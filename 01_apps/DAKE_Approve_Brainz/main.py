# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import html
import json
import os
import socket
import sys
import tempfile
import threading
import time
import urllib.parse
import urllib.request
import webbrowser
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox


APP_NAME = "DakeApproveBrainz"
WINDOW_TITLE = "承認Brainz"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

SERVER_HOST = "0.0.0.0"
DEFAULT_PORT = 8765

UI_TEXT = {
    "display_name": "承認Brainz",
    "status_server_running": "ローカル承認サーバー稼働中",
    "status_server_error": "ローカル承認サーバーを起動できませんでした: {error}",
    "url_label": "URL",
    "local_url_label": "PC確認用",
    "lan_url_label": "スマホ用",
    "pending_status_label": "pending.json の状態",
    "pending_ready": "承認待ち: {id} / {summary}",
    "pending_missing_created": "pending.json がなかったため、サンプルを生成しました。",
    "pending_read_error": "pending.json を読み込めませんでした。JSON形式を確認してください: {error}",
    "result_status_label": "result.json の最終判断",
    "result_none": "まだ判断はありません",
    "result_ready": "{decision} / {decided_at}",
    "result_read_error": "result.json を読み込めませんでした。JSON形式を確認してください: {error}",
    "button_open_browser": "ブラウザで開く",
    "button_regenerate_sample": "サンプル承認待ちを再生成",
    "button_exit": "終了",
    "button_refresh": "更新",
    "button_approve": "承認する",
    "button_deny": "却下する",
    "message_sample_regenerated": "サンプル承認待ちを再生成しました。",
    "message_open_failed": "ブラウザを開けませんでした: {error}",
    "message_approved": "承認しました",
    "message_denied": "却下しました",
    "message_saved": "判断を保存しました。",
    "dialog_error_title": "エラー",
    "dialog_info_title": "承認Brainz",
    "web_title": "承認Brainz",
    "web_heading": "承認待ち 1件",
    "web_target": "対象",
    "web_project": "場所",
    "web_summary": "内容",
    "web_risk": "危険度",
    "web_detail": "詳細",
    "web_current_result": "最終判断",
    "web_no_pending": "承認待ちはありません",
    "web_error": "処理できませんでした: {error}",
    "sample_target": "Codex",
    "sample_project": "DAKE_series",
    "sample_summary": "README更新 / Git push確認",
    "sample_risk": "中",
    "sample_detail": "これはDAKE_Approve_Brainzの初期動作確認用サンプルです。",
    "sample_created_at": "",
    "decision_approve": "approve",
    "decision_deny": "deny",
    "decision_label_approve": "承認",
    "decision_label_deny": "却下",
    "unknown": "不明",
    "health_ok": "OK",
    "launch_check_ok": "LAUNCH CHECK OK",
    "smoke_check_ok": "SMOKE OK",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_tagline": "止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": "｜",
    "footer_copyright": COPYRIGHT,
}

COLORS = {
    "bg": "#F6F7F8",
    "panel": "#FFFFFF",
    "border": "#D7DCE2",
    "text": "#20242A",
    "muted": "#68717D",
    "button": "#26313D",
    "button_hover": "#1E2731",
    "approve": "#126C43",
    "deny": "#8A2F2B",
    "soft": "#EEF1F4",
}


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def icon_path() -> Path:
    return app_dir().parents[1] / "02_assets" / "dake_icon.ico"


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def escape(value: object) -> str:
    return html.escape(str(value if value is not None else ""), quote=True)


def sample_pending() -> dict[str, str]:
    return {
        "id": "sample-001",
        "target": UI_TEXT["sample_target"],
        "project": UI_TEXT["sample_project"],
        "summary": UI_TEXT["sample_summary"],
        "risk": UI_TEXT["sample_risk"],
        "detail": UI_TEXT["sample_detail"],
        "created_at": UI_TEXT["sample_created_at"],
    }


class ApprovalStore:
    def __init__(self, root: Path | None = None) -> None:
        self.root = root or app_dir()
        self.data_dir = self.root / "data"
        self.pending_path = self.data_dir / "pending.json"
        self.result_path = self.data_dir / "result.json"
        self._lock = threading.RLock()

    def ensure_dirs(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)

    def write_json(self, path: Path, payload: dict[str, object]) -> None:
        self.ensure_dirs()
        temp_path = path.with_suffix(path.suffix + ".tmp")
        temp_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        temp_path.replace(path)

    def ensure_pending(self) -> tuple[dict[str, object], str]:
        with self._lock:
            if not self.pending_path.exists():
                payload = sample_pending()
                self.write_json(self.pending_path, payload)
                return payload, UI_TEXT["pending_missing_created"]
            return self.read_pending_locked(), ""

    def read_pending_locked(self) -> dict[str, object]:
        data = json.loads(self.pending_path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("pending.json root must be an object")
        return data

    def read_pending(self) -> tuple[dict[str, object], str]:
        try:
            return self.ensure_pending()
        except Exception as exc:
            return sample_pending(), UI_TEXT["pending_read_error"].format(error=exc)

    def read_result(self) -> tuple[dict[str, object] | None, str]:
        with self._lock:
            if not self.result_path.exists():
                return None, ""
            try:
                data = json.loads(self.result_path.read_text(encoding="utf-8"))
                if not isinstance(data, dict):
                    raise ValueError("result.json root must be an object")
                return data, ""
            except Exception as exc:
                return None, UI_TEXT["result_read_error"].format(error=exc)

    def save_decision(self, decision: str) -> dict[str, object]:
        if decision not in {UI_TEXT["decision_approve"], UI_TEXT["decision_deny"]}:
            raise ValueError(UI_TEXT["web_error"].format(error=UI_TEXT["unknown"]))
        with self._lock:
            pending = self.ensure_pending()[0]
            payload = {
                "id": str(pending.get("id", "")),
                "decision": decision,
                "decided_at": now_text(),
            }
            self.write_json(self.result_path, payload)
            return payload

    def regenerate_sample(self, clear_result: bool = True) -> None:
        with self._lock:
            self.write_json(self.pending_path, sample_pending())
            if clear_result and self.result_path.exists():
                self.result_path.unlink()


def decision_label(decision: object) -> str:
    if decision == UI_TEXT["decision_approve"]:
        return UI_TEXT["decision_label_approve"]
    if decision == UI_TEXT["decision_deny"]:
        return UI_TEXT["decision_label_deny"]
    return str(decision or UI_TEXT["unknown"])


def footer_html() -> str:
    separator = escape(UI_TEXT["footer_separator"])
    first_line = (
        f"{escape(UI_TEXT['footer_left'])} {separator} "
        f"{escape(UI_TEXT['footer_tagline'])}"
    )
    second_line = (
        f"{escape(UI_TEXT['footer_link_1'])} {separator} "
        f"{escape(UI_TEXT['footer_link_2'])} {separator} "
        f"{escape(UI_TEXT['footer_copyright'])}"
    )
    return f"<footer><div>{first_line}</div><div>{second_line}</div></footer>"


def render_web_page(store: ApprovalStore, message: str = "") -> str:
    pending, pending_error = store.read_pending()
    result, result_error = store.read_result()

    result_text = UI_TEXT["result_none"]
    if result:
        result_text = UI_TEXT["result_ready"].format(
            decision=decision_label(result.get("decision", "")),
            decided_at=result.get("decided_at", ""),
        )

    notices = []
    if message:
        notices.append(f"<div class=\"notice\">{escape(message)}</div>")
    for error_text in (pending_error, result_error):
        if error_text:
            notices.append(f"<div class=\"error\">{escape(error_text)}</div>")

    rows = [
        (UI_TEXT["web_target"], pending.get("target", "")),
        (UI_TEXT["web_project"], pending.get("project", "")),
        (UI_TEXT["web_summary"], pending.get("summary", "")),
        (UI_TEXT["web_risk"], pending.get("risk", "")),
        (UI_TEXT["web_detail"], pending.get("detail", "")),
    ]
    row_html = "\n".join(
        f"<div class=\"row\"><dt>{escape(label)}</dt><dd>{escape(value)}</dd></div>"
        for label, value in rows
    )

    return f"""<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{escape(UI_TEXT["web_title"])}</title>
  <style>
    :root {{
      color-scheme: light;
      font-family: "BIZ UDPGothic", "Yu Gothic UI", Meiryo, sans-serif;
      background: {COLORS["bg"]};
      color: {COLORS["text"]};
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      background: {COLORS["bg"]};
      color: {COLORS["text"]};
    }}
    main {{
      width: min(720px, 100%);
      margin: 0 auto;
      padding: 18px 14px 24px;
    }}
    header {{
      padding: 10px 2px 14px;
    }}
    h1 {{
      margin: 0;
      font-size: 24px;
      font-weight: 700;
      letter-spacing: 0;
    }}
    .panel {{
      background: {COLORS["panel"]};
      border: 1px solid {COLORS["border"]};
      border-radius: 8px;
      padding: 16px;
    }}
    .notice, .error {{
      margin: 0 0 12px;
      padding: 12px;
      border-radius: 8px;
      font-weight: 700;
      line-height: 1.5;
    }}
    .notice {{
      background: #EAF6EF;
      color: {COLORS["approve"]};
      border: 1px solid #B9DEC8;
    }}
    .error {{
      background: #FAEDEC;
      color: {COLORS["deny"]};
      border: 1px solid #E4B8B4;
    }}
    dl {{
      margin: 0;
    }}
    .row {{
      padding: 12px 0;
      border-bottom: 1px solid {COLORS["soft"]};
    }}
    .row:last-child {{
      border-bottom: 0;
    }}
    dt {{
      margin: 0 0 5px;
      color: {COLORS["muted"]};
      font-size: 13px;
      font-weight: 700;
    }}
    dd {{
      margin: 0;
      font-size: 16px;
      line-height: 1.55;
      overflow-wrap: anywhere;
    }}
    .result {{
      margin-top: 14px;
      color: {COLORS["muted"]};
      font-size: 14px;
      line-height: 1.5;
    }}
    .actions {{
      display: grid;
      grid-template-columns: 1fr;
      gap: 10px;
      margin-top: 16px;
    }}
    button, a.button {{
      width: 100%;
      min-height: 48px;
      border: 0;
      border-radius: 8px;
      padding: 12px 14px;
      font: inherit;
      font-weight: 700;
      color: #FFFFFF;
      text-align: center;
      text-decoration: none;
      cursor: pointer;
    }}
    .approve {{ background: {COLORS["approve"]}; }}
    .deny {{ background: {COLORS["deny"]}; }}
    .refresh {{ background: {COLORS["button"]}; }}
    footer {{
      margin-top: 18px;
      padding: 12px 2px 0;
      color: {COLORS["muted"]};
      font-size: 12px;
      line-height: 1.7;
    }}
    @media (min-width: 520px) {{
      .actions {{
        grid-template-columns: 1fr 1fr 1fr;
      }}
    }}
  </style>
</head>
<body>
  <main>
    <header>
      <h1>{escape(UI_TEXT["web_heading"])}</h1>
    </header>
    {"".join(notices)}
    <section class="panel">
      <dl>
        {row_html}
      </dl>
      <div class="result">{escape(UI_TEXT["web_current_result"])}: {escape(result_text)}</div>
      <div class="actions">
        <form method="post" action="/approve"><button class="approve" type="submit">{escape(UI_TEXT["button_approve"])}</button></form>
        <form method="post" action="/deny"><button class="deny" type="submit">{escape(UI_TEXT["button_deny"])}</button></form>
        <form method="get" action="/"><button class="refresh" type="submit">{escape(UI_TEXT["button_refresh"])}</button></form>
      </div>
    </section>
    {footer_html()}
  </main>
</body>
</html>
"""


class ApprovalHTTPServer(ThreadingHTTPServer):
    def __init__(self, server_address: tuple[str, int], store: ApprovalStore) -> None:
        super().__init__(server_address, ApprovalRequestHandler)
        self.store = store
        self.daemon_threads = True


class ApprovalRequestHandler(BaseHTTPRequestHandler):
    server_version = f"{APP_NAME}/0.1"

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/healthz":
            self.send_text(UI_TEXT["health_ok"], content_type="text/plain; charset=utf-8")
            return
        if parsed.path not in {"/", ""}:
            self.send_error(HTTPStatus.NOT_FOUND)
            return

        params = urllib.parse.parse_qs(parsed.query)
        message = ""
        if params.get("decision", [""])[0] == UI_TEXT["decision_approve"]:
            message = UI_TEXT["message_approved"]
        elif params.get("decision", [""])[0] == UI_TEXT["decision_deny"]:
            message = UI_TEXT["message_denied"]
        self.send_html(render_web_page(self.server.store, message=message))  # type: ignore[attr-defined]

    def do_POST(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/approve":
            self.save_and_redirect(UI_TEXT["decision_approve"])
            return
        if parsed.path == "/deny":
            self.save_and_redirect(UI_TEXT["decision_deny"])
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def save_and_redirect(self, decision: str) -> None:
        try:
            self.server.store.save_decision(decision)  # type: ignore[attr-defined]
            self.send_response(HTTPStatus.SEE_OTHER)
            self.send_header("Location", f"/?decision={urllib.parse.quote(decision)}")
            self.end_headers()
        except Exception as exc:
            self.send_response(HTTPStatus.INTERNAL_SERVER_ERROR)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            message = UI_TEXT["web_error"].format(error=exc)
            self.wfile.write(render_web_page(self.server.store, message=message).encode("utf-8"))  # type: ignore[attr-defined]

    def send_html(self, body: str) -> None:
        payload = body.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def send_text(self, body: str, content_type: str) -> None:
        payload = body.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def log_message(self, format: str, *args: object) -> None:
        return


class ApprovalWebServer:
    def __init__(self, store: ApprovalStore, host: str = SERVER_HOST, port: int = DEFAULT_PORT) -> None:
        self.store = store
        self.host = host
        self.port = port
        self.httpd: ApprovalHTTPServer | None = None
        self.thread: threading.Thread | None = None
        self.error = ""

    def start(self) -> bool:
        try:
            self.store.read_pending()
            self.httpd = ApprovalHTTPServer((self.host, self.port), self.store)
            self.port = int(self.httpd.server_address[1])
        except Exception as exc:
            self.error = str(exc)
            return False

        self.thread = threading.Thread(target=self._serve, name="ApproveBrainzWeb", daemon=True)
        self.thread.start()
        return True

    def _serve(self) -> None:
        if not self.httpd:
            return
        try:
            self.httpd.serve_forever(poll_interval=0.2)
        except Exception as exc:
            self.error = str(exc)

    def stop(self) -> None:
        if not self.httpd:
            return
        self.httpd.shutdown()
        self.httpd.server_close()
        if self.thread and self.thread.is_alive():
            self.thread.join(timeout=2.0)

    def is_running(self) -> bool:
        return self.httpd is not None and not self.error


def local_ip() -> str:
    candidates: list[str] = []
    try:
        host_name = socket.gethostname()
        candidates.extend(socket.gethostbyname_ex(host_name)[2])
    except Exception:
        pass
    for candidate in candidates:
        if candidate and not candidate.startswith("127."):
            return candidate
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            ip_address = sock.getsockname()[0]
            if ip_address:
                return ip_address
    except Exception:
        pass
    return "127.0.0.1"


def web_url(host: str, port: int) -> str:
    return f"http://{host}:{port}"


def choose_font(root: tk.Tk) -> str:
    families = set(tkfont.families(root))
    for family in ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo"):
        if family in families:
            return family
    return "TkDefaultFont"


class ApproveBrainzApp(tk.Tk):
    def __init__(self, store: ApprovalStore, server: ApprovalWebServer) -> None:
        super().__init__()
        self.store = store
        self.server = server
        self.font_family = choose_font(self)
        self.local_url = web_url("127.0.0.1", self.server.port)
        self.lan_url = web_url(local_ip(), self.server.port)

        self.title(WINDOW_TITLE)
        self.geometry("560x420")
        self.minsize(480, 360)
        self.configure(bg=COLORS["bg"])
        icon = icon_path()
        if icon.exists():
            try:
                self.iconbitmap(str(icon))
            except Exception:
                pass

        self.url_var = tk.StringVar(value=f"{UI_TEXT['local_url_label']}: {self.local_url}\n{UI_TEXT['lan_url_label']}: {self.lan_url}")
        self.server_var = tk.StringVar()
        self.pending_var = tk.StringVar()
        self.result_var = tk.StringVar()
        self.message_var = tk.StringVar()

        self.build_ui()
        self.refresh_status()
        self.protocol("WM_DELETE_WINDOW", self.close)

    def font(self, size: int, weight: str = "normal") -> tuple[str, int, str]:
        return (self.font_family, size, weight)

    def build_ui(self) -> None:
        outer = tk.Frame(self, bg=COLORS["bg"], padx=18, pady=16)
        outer.pack(fill="both", expand=True)

        title = tk.Label(
            outer,
            text=UI_TEXT["display_name"],
            bg=COLORS["bg"],
            fg=COLORS["text"],
            font=self.font(22, "bold"),
            anchor="w",
        )
        title.pack(fill="x")

        server = tk.Label(
            outer,
            textvariable=self.server_var,
            bg=COLORS["bg"],
            fg=COLORS["muted"],
            font=self.font(11),
            anchor="w",
        )
        server.pack(fill="x", pady=(4, 12))

        panel = tk.Frame(outer, bg=COLORS["panel"], highlightthickness=1, highlightbackground=COLORS["border"], padx=14, pady=14)
        panel.pack(fill="both", expand=True)

        self.add_section(panel, UI_TEXT["url_label"], self.url_var)
        self.add_section(panel, UI_TEXT["pending_status_label"], self.pending_var)
        self.add_section(panel, UI_TEXT["result_status_label"], self.result_var)

        message = tk.Label(
            panel,
            textvariable=self.message_var,
            bg=COLORS["panel"],
            fg=COLORS["approve"],
            font=self.font(10, "bold"),
            anchor="w",
            justify="left",
            wraplength=480,
        )
        message.pack(fill="x", pady=(2, 8))

        buttons = tk.Frame(panel, bg=COLORS["panel"])
        buttons.pack(fill="x", pady=(4, 0))

        self.make_button(buttons, UI_TEXT["button_open_browser"], self.open_browser).pack(fill="x", pady=(0, 8))
        self.make_button(buttons, UI_TEXT["button_regenerate_sample"], self.regenerate_sample).pack(fill="x", pady=(0, 8))
        self.make_button(buttons, UI_TEXT["button_exit"], self.close, fill=COLORS["soft"], fg=COLORS["text"]).pack(fill="x")

        footer = tk.Label(
            outer,
            text=self.footer_text(),
            bg=COLORS["bg"],
            fg=COLORS["muted"],
            font=self.font(9),
            anchor="w",
            justify="left",
            wraplength=520,
        )
        footer.pack(fill="x", pady=(12, 0))

    def add_section(self, parent: tk.Widget, label_text: str, variable: tk.StringVar) -> None:
        label = tk.Label(
            parent,
            text=label_text,
            bg=COLORS["panel"],
            fg=COLORS["muted"],
            font=self.font(10, "bold"),
            anchor="w",
        )
        label.pack(fill="x")

        value = tk.Label(
            parent,
            textvariable=variable,
            bg=COLORS["panel"],
            fg=COLORS["text"],
            font=self.font(11),
            anchor="w",
            justify="left",
            wraplength=490,
        )
        value.pack(fill="x", pady=(3, 12))

    def make_button(self, parent: tk.Widget, label: str, command: object, fill: str | None = None, fg: str = "#FFFFFF") -> tk.Button:
        return tk.Button(
            parent,
            text=label,
            command=command,
            bg=fill or COLORS["button"],
            fg=fg,
            activebackground=COLORS["button_hover"],
            activeforeground="#FFFFFF",
            relief="flat",
            bd=0,
            padx=12,
            pady=10,
            cursor="hand2",
            font=self.font(11, "bold"),
        )

    def footer_text(self) -> str:
        separator = UI_TEXT["footer_separator"]
        first_line = f"{UI_TEXT['footer_left']} {separator} {UI_TEXT['footer_tagline']}"
        second_line = (
            f"{UI_TEXT['footer_link_1']} {separator} "
            f"{UI_TEXT['footer_link_2']} {separator} "
            f"{UI_TEXT['footer_copyright']}"
        )
        return f"{first_line}\n{second_line}"

    def refresh_status(self) -> None:
        if self.server.error:
            self.server_var.set(UI_TEXT["status_server_error"].format(error=self.server.error))
        else:
            self.server_var.set(UI_TEXT["status_server_running"])

        pending, pending_error = self.store.read_pending()
        if pending_error:
            self.pending_var.set(pending_error)
        else:
            self.pending_var.set(
                UI_TEXT["pending_ready"].format(
                    id=pending.get("id", ""),
                    summary=pending.get("summary", ""),
                )
            )

        result, result_error = self.store.read_result()
        if result_error:
            self.result_var.set(result_error)
        elif result:
            self.result_var.set(
                UI_TEXT["result_ready"].format(
                    decision=decision_label(result.get("decision", "")),
                    decided_at=result.get("decided_at", ""),
                )
            )
        else:
            self.result_var.set(UI_TEXT["result_none"])

        self.after(1500, self.refresh_status)

    def open_browser(self) -> None:
        try:
            webbrowser.open(self.local_url)
        except Exception as exc:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["message_open_failed"].format(error=exc))

    def regenerate_sample(self) -> None:
        try:
            self.store.regenerate_sample(clear_result=True)
            self.message_var.set(UI_TEXT["message_sample_regenerated"])
            self.refresh_status()
        except Exception as exc:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["web_error"].format(error=exc))

    def close(self) -> None:
        self.server.stop()
        self.destroy()


def run_gui(gui_smoke_seconds: float | None = None) -> int:
    store = ApprovalStore()
    server = ApprovalWebServer(store)
    server.start()
    app = ApproveBrainzApp(store, server)
    if gui_smoke_seconds is not None:
        app.after(int(gui_smoke_seconds * 1000), app.close)
    app.mainloop()
    return 0


def run_serve_seconds(seconds: float) -> int:
    store = ApprovalStore()
    server = ApprovalWebServer(store)
    if not server.start():
        raise RuntimeError(f"server start failed: {server.error}")
    try:
        print(web_url("127.0.0.1", server.port))
        time.sleep(max(0.1, seconds))
    finally:
        server.stop()
    return 0


def run_launch_check() -> int:
    store = ApprovalStore()
    pending, _ = store.read_pending()
    if not isinstance(pending, dict):
        raise RuntimeError("pending check failed")
    print(UI_TEXT["launch_check_ok"])
    return 0


def post_form(url: str) -> str:
    request = urllib.request.Request(url, data=b"", method="POST")
    with urllib.request.urlopen(request, timeout=5) as response:
        return response.read().decode("utf-8", errors="replace")


def run_smoke_test() -> int:
    with tempfile.TemporaryDirectory() as tmp:
        store = ApprovalStore(Path(tmp))
        pending, _ = store.ensure_pending()
        if pending.get("id") != "sample-001":
            raise RuntimeError("pending sample generation failed")

        server = ApprovalWebServer(store, host="127.0.0.1", port=0)
        if not server.start():
            raise RuntimeError(f"server start failed: {server.error}")
        try:
            base_url = web_url("127.0.0.1", server.port)
            with urllib.request.urlopen(f"{base_url}/", timeout=5) as response:
                html_text = response.read().decode("utf-8", errors="replace")
            if UI_TEXT["web_heading"] not in html_text:
                raise RuntimeError("web page heading check failed")

            post_form(f"{base_url}/approve")
            result, _ = store.read_result()
            if not result or result.get("decision") != UI_TEXT["decision_approve"]:
                raise RuntimeError("approve decision save failed")

            post_form(f"{base_url}/deny")
            result, _ = store.read_result()
            if not result or result.get("decision") != UI_TEXT["decision_deny"]:
                raise RuntimeError("deny decision save failed")
        finally:
            server.stop()

    print(UI_TEXT["smoke_check_ok"])
    return 0


def run_server_check() -> int:
    store = ApprovalStore()
    server = ApprovalWebServer(store, host="127.0.0.1", port=DEFAULT_PORT)
    if not server.start():
        raise RuntimeError(f"server start failed: {server.error}")
    try:
        base_url = web_url("127.0.0.1", server.port)
        with urllib.request.urlopen(f"{base_url}/", timeout=5) as response:
            html_text = response.read().decode("utf-8", errors="replace")
        if UI_TEXT["web_heading"] not in html_text:
            raise RuntimeError("web page heading check failed")

        post_form(f"{base_url}/approve")
        result, _ = store.read_result()
        if not result or result.get("decision") != UI_TEXT["decision_approve"]:
            raise RuntimeError("approve decision save failed")

        post_form(f"{base_url}/deny")
        result, _ = store.read_result()
        if not result or result.get("decision") != UI_TEXT["decision_deny"]:
            raise RuntimeError("deny decision save failed")
    finally:
        server.stop()

    print(f"{UI_TEXT['smoke_check_ok']}: {web_url('127.0.0.1', DEFAULT_PORT)}")
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(prog=APP_NAME)
    parser.add_argument("--launch-check", action="store_true")
    parser.add_argument("--smoke-test", action="store_true")
    parser.add_argument("--server-check", action="store_true")
    parser.add_argument("--gui-smoke-seconds", type=float)
    parser.add_argument("--serve-seconds", type=float)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.launch_check:
        return run_launch_check()
    if args.smoke_test:
        return run_smoke_test()
    if args.server_check:
        return run_server_check()
    if args.serve_seconds is not None:
        return run_serve_seconds(args.serve_seconds)
    return run_gui(gui_smoke_seconds=args.gui_smoke_seconds)


if __name__ == "__main__":
    raise SystemExit(main())
