from __future__ import annotations

import argparse
import json
import os
import queue
import re
import socket
import subprocess
import sys
import tempfile
import threading
import time
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path


APP_NAME = "補助脳 起こす"
WINDOW_TITLE = "補助脳 起こす - Wake-on-LAN Trigger"
COPYRIGHT = "PEAKHEADZ / DAKE"

UI_TEXT = {
    "subtitle": "Wake-on-LAN Trigger",
    "section_target": "Target",
    "section_status": "Status",
    "section_log": "Log",
    "label_pc_name": "PC Name",
    "label_mac_address": "MAC Address",
    "label_broadcast_ip": "Broadcast IP",
    "label_port": "Port",
    "placeholder_pc_name": "BRAINZ-PC",
    "placeholder_mac": "AA:BB:CC:DD:EE:FF",
    "button_save": "Save",
    "button_wake": "Wake",
    "button_check": "Check",
    "button_working": "Working...",
    "status_offline": "OFFLINE",
    "status_starting": "STARTING...",
    "status_online": "ONLINE",
    "status_ready": "READY",
    "status_error": "ERROR",
    "log_ready": "Wake tool ready.",
    "log_saved": "Config saved.",
    "log_wake_sent": "Wake packet sent.",
    "log_waiting_response": "Waiting for response...",
    "log_online": "Brainz PC online.",
    "log_offline": "Brainz PC offline.",
    "log_checking": "Checking status...",
    "log_invalid_mac": "Invalid MAC Address.",
    "log_send_failed": "Wake packet failed: {error}",
    "log_no_pc_name": "PC Name is empty. Wake packet was sent without ping check.",
    "phrase_brainz_connected": "補助脳基地へ接続しました。",
    "dialog_error": "Error",
    "error_mac_required": "MAC Addressを入力してください。",
    "error_invalid_mac": "MAC Address形式を確認してください。例: AA:BB:CC:DD:EE:FF",
    "error_invalid_port": "Portは1から65535の数字で入力してください。",
    "error_config_save": "設定保存に失敗しました: {error}",
}

COLORS = {
    "bg": "#050912",
    "panel": "#08111C",
    "panel_soft": "#0B1624",
    "accent": "#2F8CFF",
    "accent_hover": "#1D6FD1",
    "border": "#23466F",
    "text": "#EAF2FF",
    "muted": "#9FB4CC",
    "weak": "#6F829A",
    "online": "#63E6BE",
    "offline": "#8796A8",
    "error": "#FF8A8A",
}


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return app_dir()


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


@dataclass
class WakeConfig:
    pc_name: str = ""
    mac_address: str = ""
    broadcast_ip: str = "255.255.255.255"
    port: int = 9


class ConfigStore:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or (app_dir() / "data" / "config" / "config.json")

    def load(self) -> WakeConfig:
        if not self.path.exists():
            return WakeConfig()
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except Exception:
            return WakeConfig()
        try:
            port = parse_port(data.get("port", 9))
        except ValueError:
            port = 9
        return WakeConfig(
            pc_name=str(data.get("pc_name", "") or ""),
            mac_address=str(data.get("mac_address", "") or ""),
            broadcast_ip=str(data.get("broadcast_ip", "255.255.255.255") or "255.255.255.255"),
            port=port,
        )

    def save(self, config: WakeConfig) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(asdict(config), ensure_ascii=False, indent=2), encoding="utf-8")


def parse_port(value: object) -> int:
    try:
        port = int(str(value).strip())
    except Exception as exc:
        raise ValueError(UI_TEXT["error_invalid_port"]) from exc
    if port < 1 or port > 65535:
        raise ValueError(UI_TEXT["error_invalid_port"])
    return port


def normalize_mac(mac_address: str) -> str:
    raw = str(mac_address or "").strip()
    if not raw:
        raise ValueError(UI_TEXT["error_mac_required"])
    clean = re.sub(r"[^0-9A-Fa-f]", "", raw)
    if len(clean) != 12 or not re.fullmatch(r"[0-9A-Fa-f]{12}", clean):
        raise ValueError(UI_TEXT["error_invalid_mac"])
    pairs = [clean[index : index + 2].upper() for index in range(0, 12, 2)]
    return ":".join(pairs)


def build_magic_packet(mac_address: str) -> bytes:
    normalized = normalize_mac(mac_address)
    mac_bytes = bytes.fromhex(normalized.replace(":", ""))
    return b"\xff" * 6 + mac_bytes * 16


def send_magic_packet(mac_address: str, broadcast_ip: str = "255.255.255.255", port: int = 9) -> None:
    packet = build_magic_packet(mac_address)
    with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
        sock.sendto(packet, (broadcast_ip or "255.255.255.255", int(port)))


def probe_host(host: str, timeout_ms: int = 1000) -> bool:
    target = str(host or "").strip()
    if not target:
        return False
    if os.name == "nt":
        command = ["ping", "-n", "1", "-w", str(timeout_ms), target]
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    else:
        command = ["ping", "-c", "1", "-W", str(max(1, timeout_ms // 1000)), target]
        creationflags = 0
    try:
        result = subprocess.run(
            command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
            creationflags=creationflags,
        )
        return result.returncode == 0
    except Exception:
        return False


def write_log_line(message: str) -> None:
    log_dir = app_dir() / "data" / "logs"
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / f"wake_{datetime.now().strftime('%Y%m%d')}.log"
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(f"[{now_text()}] {message}\n")
    except Exception:
        return


def run_launch_check() -> int:
    ConfigStore().load()
    packet = build_magic_packet("AA:BB:CC:DD:EE:FF")
    if len(packet) != 102:
        raise RuntimeError("magic packet length check failed")
    print("LAUNCH CHECK OK")
    return 0


def run_smoke_test() -> int:
    with tempfile.TemporaryDirectory() as tmp:
        store = ConfigStore(Path(tmp) / "config.json")
        config = WakeConfig(
            pc_name="127.0.0.1",
            mac_address="AA:BB:CC:DD:EE:FF",
            broadcast_ip="127.0.0.1",
            port=9,
        )
        store.save(config)
        loaded = store.load()
        if loaded != config:
            raise RuntimeError("config roundtrip failed")
        invalid_mac_ok = False
        try:
            normalize_mac("invalid")
        except ValueError:
            invalid_mac_ok = True
        if not invalid_mac_ok:
            raise RuntimeError("invalid MAC check failed")
        packet = build_magic_packet(config.mac_address)
        if len(packet) != 102 or not packet.startswith(b"\xff" * 6):
            raise RuntimeError("magic packet build failed")
        send_magic_packet(config.mac_address, config.broadcast_ip, config.port)
        localhost_online = probe_host("127.0.0.1", timeout_ms=500)
        if not localhost_online:
            raise RuntimeError("localhost status check failed")

    print("SMOKE OK")
    print(f"packet_bytes={len(packet)}")
    print("wake_send=ok")
    print(f"localhost_online={localhost_online}")
    print(f"invalid_mac={invalid_mac_ok}")
    return 0


def run_gui(gui_smoke_seconds: float | None = None) -> int:
    import customtkinter as ctk
    from tkinter import messagebox

    try:
        from PIL import Image
    except Exception:
        Image = None

    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    class WakeBrainzApp(ctk.CTk):
        def __init__(self) -> None:
            super().__init__()
            self.title(WINDOW_TITLE)
            self.geometry("720x720")
            self.minsize(640, 640)
            self.configure(fg_color=COLORS["bg"])
            self.events: queue.Queue[tuple[str, object]] = queue.Queue()
            self.worker: threading.Thread | None = None
            self.config_store = ConfigStore()
            self.config_data = self.config_store.load()
            self.logo_image = None

            self.pc_name_var = ctk.StringVar(value=self.config_data.pc_name)
            self.mac_var = ctk.StringVar(value=self.config_data.mac_address)
            self.broadcast_var = ctk.StringVar(value=self.config_data.broadcast_ip)
            self.port_var = ctk.StringVar(value=str(self.config_data.port))
            self.status_var = ctk.StringVar(value=UI_TEXT["status_ready"])

            self._build_ui(Image)
            self._append_log(UI_TEXT["log_ready"])
            self.after(120, self._poll_events)
            if gui_smoke_seconds is not None:
                self.after(max(200, int(gui_smoke_seconds * 1000)), self.destroy)

        def _build_ui(self, image_module: object) -> None:
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(2, weight=1)

            header = ctk.CTkFrame(self, fg_color="transparent")
            header.grid(row=0, column=0, sticky="ew", padx=24, pady=(22, 10))
            header.grid_columnconfigure(1, weight=1)

            logo_path = resource_dir() / "assets" / "peakheadz_logo.png"
            if image_module is not None and logo_path.exists():
                try:
                    image = image_module.open(logo_path)
                    self.logo_image = ctk.CTkImage(light_image=image, dark_image=image, size=(38, 38))
                    ctk.CTkLabel(header, image=self.logo_image, text="").grid(row=0, column=0, rowspan=2, padx=(0, 12))
                except Exception:
                    self.logo_image = None

            ctk.CTkLabel(
                header,
                text=APP_NAME,
                text_color=COLORS["text"],
                font=("Yu Gothic UI", 26),
                anchor="w",
            ).grid(row=0, column=1, sticky="w")
            ctk.CTkLabel(
                header,
                text=UI_TEXT["subtitle"],
                text_color=COLORS["muted"],
                font=("Segoe UI", 13),
                anchor="w",
            ).grid(row=1, column=1, sticky="w", pady=(2, 0))

            target = self._panel()
            target.grid(row=1, column=0, sticky="ew", padx=24, pady=(8, 12))
            target.grid_columnconfigure(1, weight=1)

            ctk.CTkLabel(
                target,
                text=UI_TEXT["section_target"],
                text_color=COLORS["text"],
                font=("Segoe UI", 15),
                anchor="w",
            ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=18, pady=(16, 10))

            self._form_row(target, 1, UI_TEXT["label_pc_name"], self.pc_name_var, UI_TEXT["placeholder_pc_name"])
            self._form_row(target, 2, UI_TEXT["label_mac_address"], self.mac_var, UI_TEXT["placeholder_mac"])
            self._form_row(target, 3, UI_TEXT["label_broadcast_ip"], self.broadcast_var, "255.255.255.255")
            self._form_row(target, 4, UI_TEXT["label_port"], self.port_var, "9")

            buttons = ctk.CTkFrame(target, fg_color="transparent")
            buttons.grid(row=5, column=0, columnspan=2, sticky="ew", padx=18, pady=(10, 18))
            buttons.grid_columnconfigure((0, 1, 2), weight=1)
            self.save_button = self._button(buttons, UI_TEXT["button_save"], self._save_config)
            self.save_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))
            self.check_button = self._button(buttons, UI_TEXT["button_check"], self._check_status)
            self.check_button.grid(row=0, column=1, sticky="ew", padx=8)
            self.wake_button = self._button(buttons, UI_TEXT["button_wake"], self._wake)
            self.wake_button.grid(row=0, column=2, sticky="ew", padx=(8, 0))

            status_panel = self._panel()
            status_panel.grid(row=2, column=0, sticky="nsew", padx=24, pady=(0, 22))
            status_panel.grid_columnconfigure(0, weight=1)
            status_panel.grid_rowconfigure(3, weight=1)

            ctk.CTkLabel(
                status_panel,
                text=UI_TEXT["section_status"],
                text_color=COLORS["text"],
                font=("Segoe UI", 15),
                anchor="w",
            ).grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 6))
            self.status_label = ctk.CTkLabel(
                status_panel,
                textvariable=self.status_var,
                text_color=COLORS["muted"],
                font=("Segoe UI", 22),
                anchor="w",
            )
            self.status_label.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 12))

            ctk.CTkLabel(
                status_panel,
                text=UI_TEXT["section_log"],
                text_color=COLORS["text"],
                font=("Segoe UI", 15),
                anchor="w",
            ).grid(row=2, column=0, sticky="ew", padx=18, pady=(2, 8))
            self.log_box = ctk.CTkTextbox(
                status_panel,
                fg_color=COLORS["bg"],
                border_color=COLORS["border"],
                border_width=1,
                text_color=COLORS["muted"],
                font=("Consolas", 12),
                wrap="word",
            )
            self.log_box.grid(row=3, column=0, sticky="nsew", padx=18, pady=(0, 18))

        def _panel(self) -> object:
            return ctk.CTkFrame(self, fg_color=COLORS["panel"], border_color=COLORS["border"], border_width=1, corner_radius=8)

        def _form_row(self, parent: object, row: int, label: str, variable: object, placeholder: str) -> None:
            import customtkinter as ctk

            ctk.CTkLabel(
                parent,
                text=label,
                text_color=COLORS["muted"],
                font=("Yu Gothic UI", 13),
                anchor="w",
            ).grid(row=row, column=0, sticky="w", padx=(18, 12), pady=8)
            ctk.CTkEntry(
                parent,
                textvariable=variable,
                placeholder_text=placeholder,
                fg_color=COLORS["panel_soft"],
                border_color=COLORS["border"],
                text_color=COLORS["text"],
                font=("Segoe UI", 13),
                height=34,
            ).grid(row=row, column=1, sticky="ew", padx=(0, 18), pady=8)

        def _button(self, parent: object, label: str, command: object) -> object:
            import customtkinter as ctk

            return ctk.CTkButton(
                parent,
                text=label,
                command=command,
                height=38,
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                text_color=COLORS["text"],
                font=("Yu Gothic UI", 13),
            )

        def _read_config_from_inputs(self) -> WakeConfig:
            port = parse_port(self.port_var.get())
            mac = normalize_mac(self.mac_var.get())
            return WakeConfig(
                pc_name=str(self.pc_name_var.get()).strip(),
                mac_address=mac,
                broadcast_ip=str(self.broadcast_var.get()).strip() or "255.255.255.255",
                port=port,
            )

        def _save_config(self) -> bool:
            try:
                config = self._read_config_from_inputs()
                self.config_store.save(config)
                self.config_data = config
                self.mac_var.set(config.mac_address)
                self._append_log(UI_TEXT["log_saved"])
                return True
            except Exception as exc:
                self._set_status(UI_TEXT["status_error"])
                self._append_log(str(exc))
                messagebox.showerror(UI_TEXT["dialog_error"], str(exc))
                return False

        def _wake(self) -> None:
            if self.worker and self.worker.is_alive():
                return
            try:
                config = self._read_config_from_inputs()
                self.config_store.save(config)
                self.config_data = config
                self.mac_var.set(config.mac_address)
            except Exception as exc:
                self._set_status(UI_TEXT["status_error"])
                self._append_log(UI_TEXT["log_invalid_mac"])
                messagebox.showerror(UI_TEXT["dialog_error"], str(exc))
                return
            self._set_status(UI_TEXT["status_starting"])
            self.wake_button.configure(state="disabled", text=UI_TEXT["button_working"])
            self.check_button.configure(state="disabled")
            self.worker = threading.Thread(target=self._wake_worker, args=(config,), daemon=True)
            self.worker.start()

        def _wake_worker(self, config: WakeConfig) -> None:
            try:
                send_magic_packet(config.mac_address, config.broadcast_ip, config.port)
                self.events.put(("log", UI_TEXT["log_wake_sent"]))
                if not config.pc_name:
                    self.events.put(("log", UI_TEXT["log_no_pc_name"]))
                    return
                self.events.put(("log", UI_TEXT["log_waiting_response"]))
                for _index in range(10):
                    if probe_host(config.pc_name, timeout_ms=1000):
                        self.events.put(("status", UI_TEXT["status_online"]))
                        self.events.put(("log", UI_TEXT["log_online"]))
                        self.events.put(("log", UI_TEXT["phrase_brainz_connected"]))
                        return
                    time.sleep(2)
                self.events.put(("status", UI_TEXT["status_offline"]))
                self.events.put(("log", UI_TEXT["log_offline"]))
            except Exception as exc:
                self.events.put(("status", UI_TEXT["status_error"]))
                self.events.put(("log", UI_TEXT["log_send_failed"].format(error=exc)))
            finally:
                self.events.put(("wake_done", ""))

        def _check_status(self) -> None:
            if self.worker and self.worker.is_alive():
                return
            host = str(self.pc_name_var.get()).strip()
            self._set_status(UI_TEXT["status_starting"])
            self._append_log(UI_TEXT["log_checking"])
            self.check_button.configure(state="disabled", text=UI_TEXT["button_working"])
            self.worker = threading.Thread(target=self._check_worker, args=(host,), daemon=True)
            self.worker.start()

        def _check_worker(self, host: str) -> None:
            try:
                if probe_host(host, timeout_ms=1000):
                    self.events.put(("status", UI_TEXT["status_online"]))
                    self.events.put(("log", UI_TEXT["log_online"]))
                    self.events.put(("log", UI_TEXT["phrase_brainz_connected"]))
                else:
                    self.events.put(("status", UI_TEXT["status_offline"]))
                    self.events.put(("log", UI_TEXT["log_offline"]))
            finally:
                self.events.put(("check_done", ""))

        def _poll_events(self) -> None:
            while True:
                try:
                    event, payload = self.events.get_nowait()
                except queue.Empty:
                    break
                if event == "log":
                    self._append_log(str(payload))
                elif event == "status":
                    self._set_status(str(payload))
                elif event == "wake_done":
                    self.wake_button.configure(state="normal", text=UI_TEXT["button_wake"])
                    self.check_button.configure(state="normal", text=UI_TEXT["button_check"])
                elif event == "check_done":
                    self.check_button.configure(state="normal", text=UI_TEXT["button_check"])
            self.after(120, self._poll_events)

        def _set_status(self, value: str) -> None:
            self.status_var.set(value)
            color = COLORS["muted"]
            if value == UI_TEXT["status_online"]:
                color = COLORS["online"]
            elif value == UI_TEXT["status_offline"]:
                color = COLORS["offline"]
            elif value == UI_TEXT["status_error"]:
                color = COLORS["error"]
            self.status_label.configure(text_color=color)

        def _append_log(self, message: str) -> None:
            line = f"[{now_text()}] {message}"
            self.log_box.insert("end", line + "\n")
            self.log_box.see("end")
            write_log_line(message)

    app = WakeBrainzApp()
    app.mainloop()
    return 0


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=WINDOW_TITLE)
    parser.add_argument("--launch-check", action="store_true")
    parser.add_argument("--smoke-test", action="store_true")
    parser.add_argument("--gui-smoke-seconds", type=float, default=None)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    if args.launch_check:
        return run_launch_check()
    if args.smoke_test:
        return run_smoke_test()
    return run_gui(args.gui_smoke_seconds)


if __name__ == "__main__":
    raise SystemExit(main())
