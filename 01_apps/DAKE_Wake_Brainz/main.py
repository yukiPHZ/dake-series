from __future__ import annotations

import argparse
import json
import os
import re
import socket
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from flask import Flask, jsonify, render_template

try:
    from wakeonlan import send_magic_packet as wakeonlan_send_magic_packet
except Exception:  # pragma: no cover - handled at runtime for a clear UI error.
    wakeonlan_send_magic_packet = None


APP_NAME = "DAKE_Wake_Brainz"
APP_SUBTITLE = "補助脳 GATE"
DEFAULT_CONFIG = {
    "target_mac": "",
    "target_ip": "",
    "web_port": 8766,
    "broadcast_ip": "255.255.255.255",
    "wake_port": 9,
    "ping_timeout_ms": 1000,
}
MAC_RE = re.compile(r"^[0-9A-F]{2}(:[0-9A-F]{2}){5}$")


@dataclass(frozen=True)
class WakeConfig:
    target_mac: str
    target_ip: str
    web_port: int
    broadcast_ip: str
    wake_port: int
    ping_timeout_ms: int
    config_path: Path
    config_exists: bool
    config_error: str | None = None

    @property
    def wake_ready(self) -> bool:
        return bool(self.target_mac)

    @property
    def status_ready(self) -> bool:
        return bool(self.target_ip)


class ConfigError(ValueError):
    pass


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def config_path() -> Path:
    return app_dir() / "config.json"


def coerce_int(value: Any, default: int, minimum: int, maximum: int) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        return default
    if number < minimum or number > maximum:
        return default
    return number


def load_config(path: Path | None = None) -> WakeConfig:
    resolved_path = path or config_path()
    data: dict[str, Any] = dict(DEFAULT_CONFIG)
    exists = resolved_path.exists()
    error: str | None = None

    if exists:
        try:
            loaded = json.loads(resolved_path.read_text(encoding="utf-8-sig"))
            if isinstance(loaded, dict):
                data.update(loaded)
            else:
                error = "config.json must be a JSON object."
        except Exception as exc:
            error = f"config.json could not be read: {exc}"

    return WakeConfig(
        target_mac=str(data.get("target_mac") or "").strip(),
        target_ip=str(data.get("target_ip") or "").strip(),
        web_port=coerce_int(data.get("web_port"), DEFAULT_CONFIG["web_port"], 1, 65535),
        broadcast_ip=str(data.get("broadcast_ip") or DEFAULT_CONFIG["broadcast_ip"]).strip(),
        wake_port=coerce_int(data.get("wake_port"), DEFAULT_CONFIG["wake_port"], 1, 65535),
        ping_timeout_ms=coerce_int(data.get("ping_timeout_ms"), DEFAULT_CONFIG["ping_timeout_ms"], 250, 10000),
        config_path=resolved_path,
        config_exists=exists,
        config_error=error,
    )


def normalize_mac(mac_address: str) -> str:
    compact = re.sub(r"[^0-9A-Fa-f]", "", mac_address or "")
    if len(compact) != 12 or not re.fullmatch(r"[0-9A-Fa-f]{12}", compact):
        raise ConfigError("config.json の target_mac を AA:BB:CC:DD:EE:FF 形式で設定してください。")
    normalized = ":".join(compact[index : index + 2].upper() for index in range(0, 12, 2))
    if not MAC_RE.fullmatch(normalized):
        raise ConfigError("config.json の target_mac が不正です。")
    return normalized


def send_wake_packet(config: WakeConfig) -> None:
    if not config.target_mac:
        raise ConfigError("config.json の target_mac が未設定です。")
    if wakeonlan_send_magic_packet is None:
        raise RuntimeError("wakeonlan が見つかりません。pip install -r requirements.txt を実行してください。")

    mac_address = normalize_mac(config.target_mac)
    wakeonlan_send_magic_packet(
        mac_address,
        ip_address=config.broadcast_ip or DEFAULT_CONFIG["broadcast_ip"],
        port=config.wake_port,
    )


def ping_host(host: str, timeout_ms: int = 1000) -> bool:
    target = str(host or "").strip()
    if not target:
        return False

    if os.name == "nt":
        command = ["ping", "-n", "1", "-w", str(timeout_ms), target]
        kwargs: dict[str, Any] = {"creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0)}
    else:
        seconds = max(1, round(timeout_ms / 1000))
        command = ["ping", "-c", "1", "-W", str(seconds), target]
        kwargs = {}

    try:
        result = subprocess.run(
            command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
            **kwargs,
        )
        return result.returncode == 0
    except Exception:
        return False


def local_ip_hint() -> str:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            return sock.getsockname()[0]
    except Exception:
        return "LIVAZ_IP"


def status_payload(config: WakeConfig) -> dict[str, Any]:
    online = ping_host(config.target_ip, config.ping_timeout_ms) if config.status_ready else False
    return {
        "ok": True,
        "target": "3070Ti",
        "online": online,
        "status": "ONLINE" if online else "OFFLINE",
        "target_ip": config.target_ip,
        "wake_ready": config.wake_ready,
        "status_ready": config.status_ready,
        "config_exists": config.config_exists,
        "config_error": config.config_error,
        "checked_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def create_app() -> Flask:
    app = Flask(__name__)

    @app.get("/")
    def index() -> str:
        config = load_config()
        return render_template(
            "index.html",
            app_name=APP_NAME,
            app_subtitle=APP_SUBTITLE,
            target_ip=config.target_ip,
            wake_ready=config.wake_ready,
            config_exists=config.config_exists,
            config_error=config.config_error,
        )

    @app.get("/api/status")
    def api_status():
        return jsonify(status_payload(load_config()))

    @app.post("/api/wake")
    def api_wake():
        config = load_config()
        try:
            send_wake_packet(config)
        except ConfigError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400
        except Exception as exc:
            return jsonify({"ok": False, "message": f"Wake failed: {exc}"}), 500

        return jsonify(
            {
                "ok": True,
                "message": "Magic packet sent.",
                "sent_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "target": "3070Ti",
            }
        )

    return app


def run_launch_check() -> int:
    config = load_config()
    normalize_mac("AA:BB:CC:DD:EE:FF")
    payload = status_payload(config)
    if payload["target"] != "3070Ti":
        raise RuntimeError("status payload check failed")
    print("LAUNCH CHECK OK")
    print(f"config_exists={config.config_exists}")
    print(f"web_port={config.web_port}")
    return 0


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=f"{APP_NAME} {APP_SUBTITLE}")
    parser.add_argument("--host", default="0.0.0.0", help="Bind host. Use 0.0.0.0 for LAN access.")
    parser.add_argument("--port", type=int, default=None, help="Override config.json web_port.")
    parser.add_argument("--launch-check", action="store_true", help="Run a lightweight startup check and exit.")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    if args.launch_check:
        return run_launch_check()

    config = load_config()
    port = args.port or config.web_port
    app = create_app()
    print(f"{APP_NAME} {APP_SUBTITLE}")
    print(f"Local: http://localhost:{port}")
    print(f"LAN:   http://{local_ip_hint()}:{port}")
    if config.config_error:
        print(config.config_error)
    elif not config.config_exists:
        print("config.json not found. UI will stay safe; create it from config.example.json.")
    app.run(host=args.host, port=port, debug=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
