# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import io
import ipaddress
import os
import queue
import re
import secrets
import shutil
import socket
import sys
import threading
import webbrowser
from pathlib import Path
from typing import Any

import qrcode
import tkinter as tk
from flask import Flask, jsonify, render_template_string, request
from PIL import Image, ImageOps, ImageTk
from tkinter import font as tkfont
from tkinter import messagebox, ttk
from werkzeug.serving import make_server

try:
    import pillow_heif

    pillow_heif.register_heif_opener()
    HEIF_AVAILABLE = True
except Exception:
    HEIF_AVAILABLE = False


APP_NAME = "Dake画像iPhoneToPC"
WINDOW_TITLE = "iPhone画像をPCに送る"
COPYRIGHT = "© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "brand_series": "シンプルそれDAKEシリーズ",
    "brand_phrase": "止まらない、迷わない、すぐ終わる。",
    "main_title": "iPhone画像をPCに送る",
    "main_description": "同じWi-FiのiPhoneから、写真をこのPCへ保存します。",
    "helper_url_label": "補助URL",
    "helper_url_description": "QRが読めない場合だけ、iPhoneのブラウザに入力してください。",
    "batch_recommendation_pc": "大量写真は30〜50枚ずつ送ると安定します。",
    "progress_title": "受信状況",
    "status_label": "状態",
    "current_file_label": "現在処理中",
    "success_count_label": "成功件数",
    "failed_count_label": "失敗件数",
    "total_count_label": "全体件数",
    "failed_files_label": "失敗したファイル",
    "failed_files_empty": "なし",
    "button_copy_url": "URLをコピー",
    "button_open_downloads": "保存先を開く",
    "button_stop": "受信を停止",
    "button_restart_receiving": "受信を再開",
    "button_stopped": "停止済み",
    "status_starting": "起動中",
    "status_waiting": "待機中",
    "status_waiting_upload": "受信待機中です",
    "status_receiving": "受信中",
    "status_receiving_progress": "受信中：{current} / {total}",
    "status_saving": "保存中",
    "status_saving_file": "保存中：{filename}",
    "status_converting": "変換中",
    "status_convert_file": "HEIC変換中：{filename}",
    "status_saved": "保存完了",
    "status_batch_complete": "保存完了：成功 {success}件 / 失敗 {failed}件",
    "status_ready_next": "続けて送信できます",
    "status_ready_continue": "続けて送信できます",
    "status_receiving_restarted": "受信を再開しました",
    "status_stopping": "受信停止中",
    "status_stopped": "受信を停止しました",
    "status_server_stopped": "受信停止中",
    "status_receiving_stopped": "受信停止中",
    "status_error": "エラー",
    "status_url_copied": "URLをコピーしました",
    "status_opened_downloads": "保存先を開きました",
    "status_not_running": "受信サーバーを起動できませんでした",
    "current_file_empty": "iPhoneからの送信を待っています",
    "current_waiting": "iPhoneからの送信を待っています",
    "current_receiving_started": "受信を開始しました",
    "current_saving": "{filename} を保存中",
    "current_converting": "{filename} をJPGに変換中",
    "current_complete": "保存が完了しました",
    "current_ready_next": "続けて別の写真を送れます",
    "current_ready_continue": "続けて別の写真を送れます",
    "current_stopped": "受信を停止しました。もう一度受け取る場合は、受信を再開してください。",
    "current_restarted": "新しいQRを読み取ってください",
    "current_failed": "{filename} の保存に失敗しました",
    "current_no_success": "保存できた写真がありませんでした",
    "same_wifi_hint": "同じWi-FiのiPhoneからアクセスしてください。",
    "firewall_hint": "Windowsファイアウォールの許可が必要な場合があります。",
    "company_wifi_hint": "会社Wi-Fiでは端末同士の通信が止められている場合があります。",
    "downloads_hint": "写真はDownloads直下へ保存します。",
    "dialog_error_title": "エラー",
    "dialog_info_title": "お知らせ",
    "dialog_copy_failed": "URLをコピーできませんでした。",
    "dialog_open_failed": "保存先を開けませんでした。",
    "dialog_server_failed": "受信サーバーを起動できませんでした。\n同じWi-Fiに接続されているか確認してください。",
    "error_ip_not_found": "ローカルIPが取得できませんでした。同じWi-Fiに接続されているか確認してください。",
    "error_port_not_found": "受信用ポートを開けませんでした。しばらくしてから、もう一度起動してください。",
    "error_save_failed": "保存に失敗しました。",
    "error_heic_failed": "HEIC変換に失敗しました。",
    "error_heic_dependency": "HEIC変換に必要なライブラリを読み込めませんでした。",
    "error_no_file": "ファイルが届きませんでした。",
    "error_token_title": "このURLでは送信できません",
    "error_token_message": "QRコードをもう一度読み取ってください。",
    "error_stopped_title": "受信を停止しました",
    "error_stopped_message": "PC側で受信が停止されています。",
    "error_root_title": "QRコードから開いてください",
    "error_root_message": "この画面からは写真を送信できません。",
    "html_title": "iPhone画像をPCに送る",
    "html_description": "写真を選んで送信してください。",
    "html_note_browser": "送信中は、この画面を閉じないでください。",
    "html_note_photo": "送信完了まで写真を削除・編集しないでください。",
    "mobile_recommended_batch": "写真が多い場合は、30〜50枚ずつ分けて送ってください。",
    "mobile_many_files_notice": "枚数が多めです。止まる場合は30〜50枚ずつ分けてください。",
    "mobile_too_many_files_notice": "100枚以上選択されています。\niPhone側で止まる場合があるため、30〜50枚ずつ分けて送ることをおすすめします。",
    "mobile_too_many_block": "100枚以上は一度に送れません。\n30〜50枚ずつ分けて選択してください。",
    "mobile_large_picker_notice": "100枚以上を一度に選ぶと、iPhone側で止まる場合があります。",
    "mobile_count_ok": "そのまま送信可能です。",
    "html_choose_label": "写真を選択",
    "html_send_button": "送信する",
    "html_count_none": "写真が選択されていません。",
    "html_count_selected": "{count}枚を送信します。",
    "html_selected_count": "選択枚数：{count}枚",
    "html_selected_hint": "送信ボタンを押すと、このPCへ保存します。",
    "html_uploading": "送信中です。完了するまでこの画面を閉じないでください。",
    "html_uploading_title": "送信中です。",
    "html_uploading_keep_open": "完了するまでこの画面を閉じないでください。",
    "html_uploading_photo": "写真アプリで写真を削除・編集しないでください。",
    "html_done": "送信完了",
    "html_done_message": "送信完了しました",
    "html_success_label": "成功",
    "html_failed_label": "失敗",
    "html_total_label": "合計",
    "html_result_summary": "{done}：{success_label} {success}件 / {failed_label} {failed}件",
    "html_total_summary": "{total_label} {total}件",
    "html_retry_hint": "失敗した写真は、もう一度選んで送信してください。",
    "html_next_action": "別の写真を送る場合は、もう一度写真を選んでください。",
    "html_next_batch_hint": "大量写真は30〜50枚ずつ分けると安定します。",
    "html_finish_hint": "終了する場合は、このタブを閉じてください。",
    "html_send_more": "別の写真を送る",
    "html_finish": "終了する",
    "html_close_tab_message": "このタブを閉じて終了してください。共有端末では、開いたままにしないでください。",
    "html_no_file": "写真を選択してください。",
    "html_network_error": "送信できませんでした。同じWi-Fiを確認して、もう一度お試しください。",
    "html_failed_files": "失敗したファイル",
    "html_server_stopped": "受信は停止されています。PC側で「受信を再開」を押してから、もう一度QRを読み取ってください。",
    "mobile_select_title": "写真を選ぶ",
    "mobile_selected_count": "選択枚数：{count}枚",
    "mobile_uploading": "送信中です。",
    "mobile_upload_complete": "送信完了",
    "mobile_upload_failed": "送信できませんでした",
    "mobile_send_more": "別の写真を送る",
    "mobile_finish": "終了する",
    "mobile_close_tab_message": "このタブを閉じて終了してください。共有端末では、開いたままにしないでください。",
    "mobile_server_stopped": "受信は停止されています。PC側で「受信を再開」を押してから、もう一度QRを読み取ってください。",
    "mobile_restart_required": "PC側で「受信を再開」を押してから、もう一度QRを読み取ってください。",
    "mobile_close_page": "画面を閉じる",
    "mobile_reload": "再読み込み",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_left_separator": " / ",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
}

LINK_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

THEME = {
    "background": "#F6F7F9",
    "card": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "success": "#12B76A",
    "danger": "#D92D20",
    "soft_accent": "#EAF2FF",
    "progress_trough": "#EAF2FF",
}

FONT_CANDIDATES = ("BIZ UDPGothic", "Yu Gothic UI", "Meiryo")
PORT_START = 5000
PORT_SCAN_COUNT = 100
QUEUE_POLL_MS = 80
WINDOW_SIZE = "920x740"
WINDOW_MIN_SIZE = (820, 680)
INVALID_FILENAME_CHARS = r'<>:"/\\|?*'
RESERVED_WINDOWS_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    "COM1",
    "COM2",
    "COM3",
    "COM4",
    "COM5",
    "COM6",
    "COM7",
    "COM8",
    "COM9",
    "LPT1",
    "LPT2",
    "LPT3",
    "LPT4",
    "LPT5",
    "LPT6",
    "LPT7",
    "LPT8",
    "LPT9",
}

UPLOAD_TEMPLATE = """
<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ text.html_title }}</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #F6F7F9;
      --card: #FFFFFF;
      --text: #1E2430;
      --muted: #667085;
      --border: #E6EAF0;
      --accent: #2F6FED;
      --accent-hover: #2458BF;
      --success: #12B76A;
      --danger: #D92D20;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      background: var(--bg);
      color: var(--text);
      font-family: "BIZ UDPGothic", "Yu Gothic UI", Meiryo, sans-serif;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 18px;
    }
    main {
      width: min(100%, 520px);
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 24px;
    }
    h1 {
      margin: 0 0 10px;
      font-size: 24px;
      line-height: 1.35;
      letter-spacing: 0;
    }
    p {
      margin: 0;
      color: var(--muted);
      font-size: 15px;
      line-height: 1.7;
    }
    .notes {
      margin: 18px 0;
      padding: 14px 16px;
      border: 1px solid var(--border);
      border-radius: 8px;
      background: #FAFBFC;
    }
    label {
      display: block;
      margin: 0 0 8px;
      font-weight: 700;
      font-size: 15px;
    }
    input[type="file"] {
      width: 100%;
      padding: 14px;
      border: 1px solid var(--border);
      border-radius: 8px;
      background: #FFFFFF;
      font-size: 16px;
    }
    button {
      width: 100%;
      margin-top: 16px;
      border: 0;
      border-radius: 8px;
      padding: 15px 18px;
      color: #FFFFFF;
      background: var(--accent);
      font-weight: 700;
      font-size: 17px;
      font-family: inherit;
    }
    button:active { background: var(--accent-hover); }
    button:disabled {
      opacity: .62;
      background: var(--muted);
    }
    .secondary-button {
      color: var(--text);
      background: #FFFFFF;
      border: 1px solid var(--border);
    }
    .secondary-button:active { background: #EAF2FF; }
    .count {
      margin-top: 10px;
      color: var(--muted);
      font-size: 14px;
      white-space: pre-wrap;
    }
    .result {
      margin-top: 16px;
      padding: 14px 16px;
      border: 1px solid var(--border);
      border-radius: 8px;
      background: #FAFBFC;
      color: var(--text);
      font-size: 15px;
      line-height: 1.7;
      white-space: pre-wrap;
    }
    .warn {
      margin-top: 10px;
      color: var(--danger);
      font-size: 14px;
      line-height: 1.6;
      white-space: pre-wrap;
    }
    .success { color: var(--success); font-weight: 700; }
    .danger { color: var(--danger); font-weight: 700; }
    .actions {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
      margin-top: 16px;
    }
    .actions button {
      margin-top: 0;
      font-size: 15px;
    }
    .hidden { display: none; }
  </style>
</head>
<body>
  <main>
    <h1>{{ text.html_title }}</h1>
    <p>{{ text.html_description }}</p>
    <div class="notes">
      <p>{{ text.html_note_browser }}</p>
      <p>{{ text.html_note_photo }}</p>
      <p>{{ text.mobile_recommended_batch }}</p>
      <p>{{ text.mobile_large_picker_notice }}</p>
    </div>
    <form id="upload-form">
      <label for="photos">{{ text.html_choose_label }}</label>
      <input id="photos" name="photos" type="file" accept="image/*" multiple>
      <div id="count" class="count">{{ text.html_count_none }}</div>
      <div id="selection-warning" class="warn" hidden></div>
      <button id="send-button" type="submit">{{ text.html_send_button }}</button>
    </form>
    <div id="result" class="result" hidden></div>
    <div id="next-actions" class="actions hidden">
      <button id="send-more-button" class="secondary-button" type="button">{{ text.mobile_send_more }}</button>
      <button id="finish-button" class="secondary-button" type="button">{{ text.mobile_finish }}</button>
    </div>
  </main>
  <script>
    const TEXT = {{ text|tojson }};
    const form = document.getElementById("upload-form");
    const input = document.getElementById("photos");
    const count = document.getElementById("count");
    const selectionWarning = document.getElementById("selection-warning");
    const button = document.getElementById("send-button");
    const result = document.getElementById("result");
    const nextActions = document.getElementById("next-actions");
    const sendMoreButton = document.getElementById("send-more-button");
    const finishButton = document.getElementById("finish-button");

    function showResult(message) {
      result.hidden = false;
      result.textContent = message;
    }

    function hideNextActions() {
      nextActions.classList.add("hidden");
    }

    function showNextActions() {
      nextActions.classList.remove("hidden");
    }

    input.addEventListener("change", () => {
      const selected = input.files.length;
      hideNextActions();
      result.hidden = true;
      button.disabled = false;
      selectionWarning.hidden = true;
      selectionWarning.textContent = "";

      if (selected === 0) {
        count.textContent = TEXT.html_count_none;
        return;
      }

      count.textContent = TEXT.html_selected_count.replace("{count}", selected) + "\\n" + TEXT.html_selected_hint;
      if (selected <= 50) {
        selectionWarning.hidden = false;
        selectionWarning.textContent = TEXT.mobile_count_ok;
      } else if (selected < 100) {
        selectionWarning.hidden = false;
        selectionWarning.textContent = TEXT.mobile_many_files_notice;
      } else {
        selectionWarning.hidden = false;
        selectionWarning.textContent = TEXT.mobile_too_many_files_notice + "\\n" + TEXT.mobile_too_many_block;
        button.disabled = true;
      }
    });

    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      hideNextActions();
      if (input.files.length === 0) {
        showResult(TEXT.html_no_file);
        return;
      }

      const data = new FormData();
      for (const file of input.files) {
        data.append("photos", file, file.name);
      }

      button.disabled = true;
      input.disabled = true;
      showResult(TEXT.html_uploading_title + "\\n" + TEXT.html_uploading_keep_open + "\\n" + TEXT.html_uploading_photo);

      try {
        const response = await fetch(window.location.href, {
          method: "POST",
          body: data
        });
        const payload = await response.json();
        if (!response.ok) {
          showResult(payload.message || TEXT.html_network_error);
          showNextActions();
          return;
        }
        let message = TEXT.html_done_message + "\\n";
        message += TEXT.html_result_summary
          .replace("{done}", TEXT.mobile_upload_complete)
          .replace("{success_label}", TEXT.html_success_label)
          .replace("{success}", payload.success)
          .replace("{failed_label}", TEXT.html_failed_label)
          .replace("{failed}", payload.failed);
        message += "\\n" + TEXT.html_total_summary
          .replace("{total_label}", TEXT.html_total_label)
          .replace("{total}", payload.total);
        if (payload.failed_files && payload.failed_files.length > 0) {
          message += `\\n\\n${TEXT.html_failed_files}\\n` + payload.failed_files.join("\\n");
          message += `\\n\\n${TEXT.html_retry_hint}`;
        }
        message += "\\n\\n" + TEXT.html_next_action + "\\n" + TEXT.html_next_batch_hint + "\\n" + TEXT.html_finish_hint;
        showResult(message);
        showNextActions();
      } catch (error) {
        showResult(TEXT.mobile_upload_failed + "\\n" + TEXT.html_network_error);
        showNextActions();
      } finally {
        button.disabled = false;
        input.disabled = false;
      }
    });

    sendMoreButton.addEventListener("click", () => {
      input.value = "";
      count.textContent = TEXT.html_count_none;
      selectionWarning.hidden = true;
      selectionWarning.textContent = "";
      button.disabled = false;
      result.hidden = true;
      hideNextActions();
      input.focus();
      input.click();
    });

    finishButton.addEventListener("click", () => {
      window.close();
      window.setTimeout(() => {
        showResult(TEXT.mobile_close_tab_message);
      }, 250);
    });
  </script>
</body>
</html>
"""

ERROR_TEMPLATE = """
<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ title }}</title>
  <style>
    body {
      margin: 0;
      min-height: 100vh;
      background: #F6F7F9;
      color: #1E2430;
      font-family: "BIZ UDPGothic", "Yu Gothic UI", Meiryo, sans-serif;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 18px;
    }
    main {
      width: min(100%, 480px);
      background: #FFFFFF;
      border: 1px solid #E6EAF0;
      border-radius: 8px;
      padding: 24px;
    }
    h1 {
      margin: 0 0 10px;
      font-size: 22px;
      letter-spacing: 0;
    }
    p {
      margin: 0;
      color: #667085;
      line-height: 1.7;
      font-size: 15px;
    }
  </style>
</head>
<body>
  <main>
    <h1>{{ title }}</h1>
    <p>{{ message }}</p>
  </main>
</body>
</html>
"""

STOPPED_TEMPLATE = """
<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ text.error_stopped_title }}</title>
  <style>
    body {
      margin: 0;
      min-height: 100vh;
      background: #F6F7F9;
      color: #1E2430;
      font-family: "BIZ UDPGothic", "Yu Gothic UI", Meiryo, sans-serif;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 18px;
    }
    main {
      width: min(100%, 480px);
      background: #FFFFFF;
      border: 1px solid #E6EAF0;
      border-radius: 8px;
      padding: 24px;
    }
    h1 {
      margin: 0 0 10px;
      font-size: 22px;
      letter-spacing: 0;
    }
    p {
      margin: 0;
      color: #667085;
      line-height: 1.7;
      font-size: 15px;
      white-space: pre-wrap;
    }
    .actions {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
      margin-top: 18px;
    }
    button {
      border: 1px solid #E6EAF0;
      border-radius: 8px;
      padding: 12px 14px;
      color: #1E2430;
      background: #FFFFFF;
      font-weight: 700;
      font-size: 15px;
      font-family: inherit;
    }
    button:active { background: #EAF2FF; }
  </style>
</head>
<body>
  <main>
    <h1>{{ text.error_stopped_title }}</h1>
    <p>{{ text.mobile_server_stopped }}</p>
    <div class="actions">
      <button type="button" onclick="window.close()">{{ text.mobile_close_page }}</button>
      <button type="button" onclick="window.location.reload()">{{ text.mobile_reload }}</button>
    </div>
  </main>
</body>
</html>
"""


def set_windows_app_id() -> None:
    if os.name != "nt":
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "Shimarisu.DAKE.ImageiPhoneToPC"
        )
    except Exception:
        pass


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_icon_candidates() -> list[Path]:
    base_dir = get_base_dir()
    source_dir = Path(__file__).resolve().parent
    return [
        source_dir / ".." / ".." / "02_assets" / "dake_icon.ico",
        base_dir / ".." / ".." / "02_assets" / "dake_icon.ico",
        base_dir / ".." / ".." / ".." / "02_assets" / "dake_icon.ico",
    ]


def choose_font_family(root: tk.Tk) -> str:
    available = set(tkfont.families(root))
    for family in FONT_CANDIDATES:
        if family in available:
            return family
    return "TkDefaultFont"


def get_downloads_dir() -> Path:
    user_profile = os.environ.get("USERPROFILE")
    if user_profile:
        downloads = Path(user_profile) / "Downloads"
    else:
        downloads = Path.home() / "Downloads"
    downloads.mkdir(parents=True, exist_ok=True)
    return downloads


def get_local_ip() -> str:
    candidates: list[str] = []
    try:
        host_name = socket.gethostname()
        for family, _, _, _, sockaddr in socket.getaddrinfo(host_name, None, socket.AF_INET):
            candidates.append(str(sockaddr[0]))
    except Exception:
        pass

    for candidate in candidates:
        if is_usable_lan_ip(candidate):
            return candidate

    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            candidate = str(sock.getsockname()[0])
            if is_usable_lan_ip(candidate):
                return candidate
    except Exception:
        pass

    raise RuntimeError(UI_TEXT["error_ip_not_found"])


def is_usable_lan_ip(value: str) -> bool:
    try:
        address = ipaddress.ip_address(value)
    except ValueError:
        return False
    return address.version == 4 and not address.is_loopback and not address.is_link_local


def find_open_port(start_port: int = PORT_START) -> int:
    for port in range(start_port, start_port + PORT_SCAN_COUNT):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind(("0.0.0.0", port))
            except OSError:
                continue
            return port
    raise RuntimeError(UI_TEXT["error_port_not_found"])


def sanitize_filename(filename: str) -> str:
    raw_name = (filename or "").replace("\\", "/").split("/")[-1].strip()
    if not raw_name:
        raw_name = "image"
    table = str.maketrans({char: "_" for char in INVALID_FILENAME_CHARS})
    cleaned = raw_name.translate(table)
    cleaned = re.sub(r"[\x00-\x1f]", "_", cleaned)
    cleaned = cleaned.strip(" .")
    if not cleaned:
        cleaned = "image"

    path = Path(cleaned)
    stem = path.stem or "image"
    suffix = path.suffix
    if stem.upper() in RESERVED_WINDOWS_NAMES:
        stem = f"{stem}_image"
    return f"{stem}{suffix}"


def unique_download_path(downloads_dir: Path, filename: str) -> Path:
    cleaned = sanitize_filename(filename)
    path = downloads_dir / cleaned
    stem = path.stem
    suffix = path.suffix
    counter = 1
    while path.exists():
        path = downloads_dir / f"{stem}_{counter}{suffix}"
        counter += 1
    return path


def jpeg_safe_image(image: Image.Image) -> Image.Image:
    fixed_image = ImageOps.exif_transpose(image)
    if fixed_image.mode == "RGB":
        return fixed_image
    if fixed_image.mode == "RGBA":
        background = Image.new("RGB", fixed_image.size, "#FFFFFF")
        background.paste(fixed_image, mask=fixed_image.getchannel("A"))
        return background
    return fixed_image.convert("RGB")


def render_error_page(title: str, message: str, status_code: int) -> tuple[str, int]:
    return render_template_string(ERROR_TEMPLATE, title=title, message=message), status_code


def make_qr_image(url: str, size: int = 238) -> Image.Image:
    qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=8, border=2)
    qr.add_data(url)
    qr.make(fit=True)
    image = qr.make_image(fill_color=THEME["text"], back_color="#FFFFFF").convert("RGB")
    return image.resize((size, size), Image.Resampling.NEAREST)


class UploadServer:
    def __init__(self, ui_queue: queue.Queue[dict[str, Any]]) -> None:
        self.ui_queue = ui_queue
        self.downloads_dir = get_downloads_dir()
        self.token = secrets.token_hex(8)
        self.host_ip = get_local_ip()
        self.port = find_open_port()
        self.url = f"http://{self.host_ip}:{self.port}/upload/{self.token}"
        self.app = Flask(__name__)
        self.server: Any | None = None
        self.thread: threading.Thread | None = None
        self.accepting = threading.Event()
        self.accepting.set()
        self.lock = threading.Lock()
        self.success_total = 0
        self.failed_total = 0
        self.total_received = 0
        self.failed_files: list[str] = []
        self._setup_routes()

    def _setup_routes(self) -> None:
        @self.app.route("/", methods=["GET", "POST"])
        def root() -> tuple[str, int]:
            return render_error_page(
                UI_TEXT["error_root_title"],
                UI_TEXT["error_root_message"],
                403,
            )

        @self.app.route("/upload", methods=["GET", "POST"])
        @self.app.route("/upload/", methods=["GET", "POST"])
        def upload_without_token() -> tuple[str, int]:
            return render_error_page(
                UI_TEXT["error_token_title"],
                UI_TEXT["error_token_message"],
                403,
            )

        @self.app.route("/upload/<incoming_token>", methods=["GET", "POST"])
        def upload(incoming_token: str) -> Any:
            if incoming_token != self.token:
                return render_error_page(
                    UI_TEXT["error_token_title"],
                    UI_TEXT["error_token_message"],
                    403,
                )

            if not self.accepting.is_set():
                if request.method == "POST":
                    return jsonify({"message": UI_TEXT["html_server_stopped"]}), 503
                return render_template_string(STOPPED_TEMPLATE, text=UI_TEXT), 503

            if request.method == "GET":
                return render_template_string(UPLOAD_TEMPLATE, text=UI_TEXT)

            return self._handle_upload()

    def start(self) -> None:
        self.server = make_server("0.0.0.0", self.port, self.app, threaded=True)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self._post(
            "server_started",
            status=UI_TEXT["status_waiting_upload"],
            current_text=UI_TEXT["current_waiting"],
        )

    def stop(self) -> None:
        self.accepting.clear()
        self._post(
            "server_stopped",
            status=UI_TEXT["status_receiving_stopped"],
            current_text=UI_TEXT["current_stopped"],
        )

    def restart(self) -> None:
        self.token = secrets.token_hex(8)
        self.url = f"http://{self.host_ip}:{self.port}/upload/{self.token}"
        self.accepting.set()
        self._post(
            "server_restarted",
            status=UI_TEXT["status_receiving_restarted"],
            current_text=UI_TEXT["current_restarted"],
            url=self.url,
        )

    def _handle_upload(self) -> Any:
        files = [item for item in request.files.getlist("photos") if item and item.filename]
        total = len(files)
        if total == 0:
            return jsonify(
                {
                    "success": 0,
                    "failed": 0,
                    "total": 0,
                    "failed_files": [],
                    "message": UI_TEXT["error_no_file"],
                }
            ), 400

        self._post(
            "receive_started",
            total=total,
            batch_success=0,
            batch_failed=0,
            status=UI_TEXT["status_receiving"],
            current_text=UI_TEXT["current_receiving_started"],
        )
        success = 0
        failed = 0
        failed_files: list[str] = []

        for index, file_storage in enumerate(files, start=1):
            filename = sanitize_filename(file_storage.filename or "image")
            self._post(
                "file_processing",
                status=UI_TEXT["status_receiving_progress"].format(current=index, total=total),
                current=index,
                total=total,
                phase=1,
                current_file=filename,
                current_text=UI_TEXT["current_saving"].format(filename=filename),
            )
            try:
                saved_path = self._save_file(file_storage, filename, index, total)
                success += 1
                with self.lock:
                    self.success_total += 1
                    self.total_received += 1
                    success_total = self.success_total
                    failed_total = self.failed_total
                    total_received = self.total_received
                self._post(
                    "file_saved",
                    status=UI_TEXT["status_saving_file"].format(filename=saved_path.name),
                    current=index,
                    total=total,
                    phase=3,
                    current_file=saved_path.name,
                    current_text=UI_TEXT["current_saving"].format(filename=saved_path.name),
                    success=success_total,
                    failed=failed_total,
                    received=total_received,
                )
            except Exception:
                failed += 1
                failed_files.append(filename)
                with self.lock:
                    self.failed_total += 1
                    self.total_received += 1
                    self.failed_files.append(filename)
                    success_total = self.success_total
                    failed_total = self.failed_total
                    total_received = self.total_received
                    failed_snapshot = list(self.failed_files[-8:])
                self._post(
                    "file_failed",
                    status=UI_TEXT["status_error"],
                    current=index,
                    total=total,
                    phase=3,
                    current_file=filename,
                    current_text=UI_TEXT["current_failed"].format(filename=filename),
                    success=success_total,
                    failed=failed_total,
                    received=total_received,
                    failed_files=failed_snapshot,
                )

        with self.lock:
            success_total = self.success_total
            failed_total = self.failed_total
            total_received = self.total_received
            failed_snapshot = list(self.failed_files[-8:])

        self._post(
            "batch_complete",
            status=UI_TEXT["status_batch_complete"].format(success=success, failed=failed),
            current_text=UI_TEXT["current_complete"] if success > 0 else UI_TEXT["current_no_success"],
            success=success_total,
            failed=failed_total,
            received=total_received,
            batch_success=success,
            batch_failed=failed,
            batch_total=total,
            failed_files=failed_snapshot,
        )

        return jsonify(
            {
                "success": success,
                "failed": failed,
                "total": total,
                "failed_files": failed_files,
            }
        )

    def _save_file(self, file_storage: Any, filename: str, index: int, total: int) -> Path:
        suffix = Path(filename).suffix.lower()
        if suffix in {".heic", ".heif"}:
            return self._save_heic_as_jpeg(file_storage, filename, index, total)

        for _attempt in range(1000):
            destination = unique_download_path(self.downloads_dir, filename)
            self._post(
                "file_saving",
                status=UI_TEXT["status_saving_file"].format(filename=destination.name),
                current_file=destination.name,
                current_text=UI_TEXT["current_saving"].format(filename=destination.name),
                current=index,
                total=total,
                phase=2,
            )
            try:
                file_storage.stream.seek(0)
                with destination.open("xb") as output:
                    shutil.copyfileobj(file_storage.stream, output, length=1024 * 1024)
                return destination
            except FileExistsError:
                continue
            except Exception:
                try:
                    destination.unlink(missing_ok=True)
                except Exception:
                    pass
                raise
        raise RuntimeError(UI_TEXT["error_save_failed"])

    def _save_heic_as_jpeg(self, file_storage: Any, filename: str, index: int, total: int) -> Path:
        if not HEIF_AVAILABLE:
            raise RuntimeError(UI_TEXT["error_heic_dependency"])

        jpeg_name = f"{Path(filename).stem}.jpg"
        self._post(
            "file_converting",
            status=UI_TEXT["status_convert_file"].format(filename=filename),
            current_file=filename,
            current_text=UI_TEXT["current_converting"].format(filename=filename),
            current=index,
            total=total,
            phase=2,
        )
        file_storage.stream.seek(0)
        try:
            image = Image.open(file_storage.stream)
            safe_image = jpeg_safe_image(image)
        except Exception as exc:
            raise RuntimeError(UI_TEXT["error_heic_failed"]) from exc

        for _attempt in range(1000):
            destination = unique_download_path(self.downloads_dir, jpeg_name)
            try:
                with destination.open("xb") as output:
                    safe_image.save(output, "JPEG", quality=95)
                return destination
            except FileExistsError:
                continue
            except Exception as exc:
                try:
                    destination.unlink(missing_ok=True)
                except Exception:
                    pass
                raise RuntimeError(UI_TEXT["error_heic_failed"]) from exc
        raise RuntimeError(UI_TEXT["error_save_failed"])

    def _post(self, event_type: str, **payload: Any) -> None:
        payload["type"] = event_type
        self.ui_queue.put(payload)


class HoverButton(tk.Button):
    def __init__(
        self,
        parent: tk.Misc,
        *,
        text_key: str,
        command: Any,
        font: tuple[str, int, str] | tuple[str, int],
        primary: bool = False,
    ) -> None:
        self.normal_bg = THEME["accent"] if primary else THEME["card"]
        self.hover_bg = THEME["accent_hover"] if primary else THEME["soft_accent"]
        self.normal_fg = "#FFFFFF" if primary else THEME["text"]
        self.disabled_bg = "#D0D5DD"
        super().__init__(
            parent,
            text=UI_TEXT[text_key],
            command=command,
            font=font,
            fg=self.normal_fg,
            bg=self.normal_bg,
            activeforeground=self.normal_fg,
            activebackground=self.hover_bg,
            relief="solid",
            bd=1,
            highlightthickness=0,
            padx=16,
            pady=8,
            cursor="hand2",
        )
        self.configure(borderwidth=0 if primary else 1)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, _event: tk.Event) -> None:
        if str(self["state"]) == "disabled":
            return
        self.configure(bg=self.hover_bg)

    def _on_leave(self, _event: tk.Event) -> None:
        if str(self["state"]) == "disabled":
            return
        self.configure(bg=self.normal_bg)

    def disable_with_text(self, text_key: str) -> None:
        self.configure(text=UI_TEXT[text_key], state="disabled", bg=self.disabled_bg, fg="#FFFFFF")

    def set_text_command(self, text_key: str, command: Any) -> None:
        self.configure(
            text=UI_TEXT[text_key],
            command=command,
            state="normal",
            bg=self.normal_bg,
            fg=self.normal_fg,
        )


class DakeImageIPhoneToPCApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=THEME["background"])
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self._apply_icon()

        self.font_family = choose_font_family(root)
        self.fonts = {
            "title": (self.font_family, 20, "bold"),
            "description": (self.font_family, 10),
            "section": (self.font_family, 12, "bold"),
            "body": (self.font_family, 10),
            "body_bold": (self.font_family, 10, "bold"),
            "small": (self.font_family, 9),
            "url": (self.font_family, 8),
            "footer": (self.font_family, 8),
            "button": (self.font_family, 10, "bold"),
        }

        self.ui_queue: queue.Queue[dict[str, Any]] = queue.Queue()
        self.server: UploadServer | None = None
        self.qr_photo: ImageTk.PhotoImage | None = None
        self.stop_thread: threading.Thread | None = None

        self.url_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value=UI_TEXT["status_starting"])
        self.current_file_var = tk.StringVar(value=UI_TEXT["current_file_empty"])
        self.success_var = tk.StringVar(value="0")
        self.failed_var = tk.StringVar(value="0")
        self.total_var = tk.StringVar(value="0")
        self.failed_files_var = tk.StringVar(value=UI_TEXT["failed_files_empty"])
        self.progress_value = tk.IntVar(value=0)
        self.progress_total = 1
        self.last_opened_batch_id: int | None = None
        self.current_batch_id = 0

        self._configure_styles()
        self._build_ui()
        self._start_server()
        self._poll_queue()

    def _apply_icon(self) -> None:
        for candidate in get_icon_candidates():
            try:
                resolved = candidate.resolve()
                if resolved.exists():
                    self.root.iconbitmap(str(resolved))
                    return
            except Exception:
                continue

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Dake.Horizontal.TProgressbar",
            troughcolor=THEME["progress_trough"],
            background=THEME["accent"],
            bordercolor=THEME["border"],
            lightcolor=THEME["accent"],
            darkcolor=THEME["accent"],
        )

    def _build_ui(self) -> None:
        content = tk.Frame(self.root, bg=THEME["background"])
        content.pack(side="top", fill="both", expand=True, padx=26, pady=(22, 0))

        self._build_header(content)
        main_area = tk.Frame(content, bg=THEME["background"])
        main_area.pack(fill="both", expand=True)
        main_area.columnconfigure(0, weight=1)
        main_area.columnconfigure(1, weight=1)
        main_area.rowconfigure(0, weight=1)

        self._build_qr_card(main_area)
        self._build_progress_card(main_area)
        self._build_footer()

    def _build_header(self, parent: tk.Misc) -> None:
        header = tk.Frame(parent, bg=THEME["background"])
        header.pack(fill="x", pady=(0, 18))
        header.columnconfigure(0, weight=0)
        header.columnconfigure(1, weight=1)
        tk.Label(
            header,
            text=UI_TEXT["main_title"],
            font=self.fonts["title"],
            fg=THEME["text"],
            bg=THEME["background"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, 18))
        tk.Label(
            header,
            text=UI_TEXT["main_description"],
            font=self.fonts["description"],
            fg=THEME["muted"],
            bg=THEME["background"],
            anchor="w",
            wraplength=560,
            justify="left",
        ).grid(row=0, column=1, sticky="w", pady=(3, 0))

    def _build_qr_card(self, parent: tk.Misc) -> None:
        card = self._card(parent)
        card.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=(0, 14))
        card.columnconfigure(0, weight=1)

        self.qr_label = tk.Label(card, bg=THEME["card"])
        self.qr_label.grid(row=0, column=0, pady=(24, 12))

        tk.Label(
            card,
            textvariable=self.status_var,
            font=self.fonts["small"],
            fg=THEME["muted"],
            bg=THEME["card"],
            anchor="center",
        ).grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 14))

        url_frame = tk.Frame(card, bg=THEME["card"])
        url_frame.grid(row=2, column=0, sticky="ew", padx=22, pady=(0, 12))
        url_frame.columnconfigure(0, weight=1)
        tk.Label(
            url_frame,
            text=UI_TEXT["helper_url_label"],
            font=self.fonts["small"],
            fg=THEME["muted"],
            bg=THEME["card"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            url_frame,
            text=UI_TEXT["helper_url_description"],
            font=self.fonts["small"],
            fg=THEME["muted"],
            bg=THEME["card"],
            anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=(3, 0))
        tk.Label(
            url_frame,
            textvariable=self.url_var,
            font=self.fonts["url"],
            fg=THEME["muted"],
            bg=THEME["card"],
            anchor="w",
        ).grid(row=2, column=0, sticky="ew", pady=(5, 0))

        hints = tk.Frame(card, bg=THEME["card"])
        hints.grid(row=3, column=0, sticky="ew", padx=22, pady=(0, 18))
        for row, key in enumerate(("downloads_hint", "batch_recommendation_pc")):
            tk.Label(
                hints,
                text=UI_TEXT[key],
                font=self.fonts["small"],
                fg=THEME["muted"],
                bg=THEME["card"],
                anchor="w",
                justify="left",
                wraplength=360,
            ).grid(row=row, column=0, sticky="w", pady=(0, 3))

        actions = tk.Frame(card, bg=THEME["card"])
        actions.grid(row=4, column=0, sticky="ew", padx=22, pady=(0, 24))
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)
        self.copy_button = HoverButton(
            actions,
            text_key="button_copy_url",
            command=self.copy_url,
            font=self.fonts["button"],
        )
        self.copy_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        HoverButton(
            actions,
            text_key="button_open_downloads",
            command=self.open_downloads,
            font=self.fonts["button"],
        ).grid(row=0, column=1, sticky="ew")

    def _build_progress_card(self, parent: tk.Misc) -> None:
        card = self._card(parent)
        card.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=(0, 14))
        card.columnconfigure(0, weight=1)

        tk.Label(
            card,
            text=UI_TEXT["progress_title"],
            font=self.fonts["section"],
            fg=THEME["text"],
            bg=THEME["card"],
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=22, pady=(22, 14))

        self.progress = ttk.Progressbar(
            card,
            mode="determinate",
            variable=self.progress_value,
            maximum=1,
            style="Dake.Horizontal.TProgressbar",
        )
        self.progress.grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 18))

        fields = tk.Frame(card, bg=THEME["card"])
        fields.grid(row=2, column=0, sticky="ew", padx=22)
        fields.columnconfigure(1, weight=1)
        self._progress_row(fields, 0, "status_label", self.status_var)
        self._progress_row(fields, 1, "current_file_label", self.current_file_var)
        self._progress_row(fields, 2, "success_count_label", self.success_var, THEME["success"])
        self._progress_row(fields, 3, "failed_count_label", self.failed_var, THEME["danger"])
        self._progress_row(fields, 4, "total_count_label", self.total_var)

        tk.Label(
            card,
            text=UI_TEXT["failed_files_label"],
            font=self.fonts["small"],
            fg=THEME["muted"],
            bg=THEME["card"],
            anchor="w",
        ).grid(row=3, column=0, sticky="ew", padx=22, pady=(20, 6))
        tk.Label(
            card,
            textvariable=self.failed_files_var,
            font=self.fonts["small"],
            fg=THEME["muted"],
            bg=THEME["card"],
            justify="left",
            anchor="nw",
            wraplength=360,
        ).grid(row=4, column=0, sticky="nsew", padx=22, pady=(0, 20))

        stop_area = tk.Frame(card, bg=THEME["card"])
        stop_area.grid(row=5, column=0, sticky="e", padx=22, pady=(0, 22))
        self.stop_button = HoverButton(
            stop_area,
            text_key="button_stop",
            command=self.stop_receiving,
            font=self.fonts["button"],
        )
        self.stop_button.pack(side="right")

    def _progress_row(
        self,
        parent: tk.Misc,
        row: int,
        label_key: str,
        variable: tk.StringVar,
        value_color: str | None = None,
    ) -> None:
        tk.Label(
            parent,
            text=UI_TEXT[label_key],
            font=self.fonts["small"],
            fg=THEME["muted"],
            bg=THEME["card"],
            anchor="w",
            width=12,
        ).grid(row=row, column=0, sticky="w", pady=7)
        tk.Label(
            parent,
            textvariable=variable,
            font=self.fonts["body_bold"],
            fg=value_color or THEME["text"],
            bg=THEME["card"],
            anchor="w",
            wraplength=260,
            justify="left",
        ).grid(row=row, column=1, sticky="ew", pady=7)

    def _build_footer(self) -> None:
        footer = tk.Frame(self.root, bg=THEME["background"])
        footer.pack(side="bottom", fill="x", padx=26, pady=(8, 14))
        left = tk.Frame(footer, bg=THEME["background"])
        left.pack(side="left", anchor="w")
        tk.Label(
            left,
            text=UI_TEXT["footer_left"] + UI_TEXT["footer_left_separator"] + UI_TEXT["brand_phrase"],
            font=self.fonts["footer"],
            fg=THEME["muted"],
            bg=THEME["background"],
            anchor="w",
        ).pack(anchor="w")

        right = tk.Frame(footer, bg=THEME["background"])
        right.pack(side="right", anchor="e")
        self._footer_link(right, "footer_link_1").pack(side="left")
        self._footer_separator(right).pack(side="left")
        self._footer_link(right, "footer_link_2").pack(side="left")
        self._footer_separator(right).pack(side="left")
        tk.Label(
            right,
            text=UI_TEXT["footer_copyright"],
            font=self.fonts["footer"],
            fg=THEME["muted"],
            bg=THEME["background"],
        ).pack(side="left")

    def _footer_link(self, parent: tk.Misc, text_key: str) -> tk.Label:
        label = tk.Label(
            parent,
            text=UI_TEXT[text_key],
            font=self.fonts["footer"],
            fg=THEME["muted"],
            bg=THEME["background"],
            cursor="hand2",
        )
        label.bind("<Button-1>", lambda _event: webbrowser.open(LINK_URLS[text_key], new=2))
        label.bind("<Enter>", lambda _event: label.configure(fg=THEME["accent"]))
        label.bind("<Leave>", lambda _event: label.configure(fg=THEME["muted"]))
        return label

    def _footer_separator(self, parent: tk.Misc) -> tk.Label:
        return tk.Label(
            parent,
            text=UI_TEXT["footer_separator"],
            font=self.fonts["footer"],
            fg=THEME["muted"],
            bg=THEME["background"],
        )

    def _card(self, parent: tk.Misc) -> tk.Frame:
        return tk.Frame(
            parent,
            bg=THEME["card"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            highlightcolor=THEME["border"],
        )

    def _start_server(self) -> None:
        try:
            self.server = UploadServer(self.ui_queue)
            self.server.start()
            self.url_var.set(self.server.url)
            self._refresh_qr(self.server.url)
        except Exception:
            self.status_var.set(UI_TEXT["status_not_running"])
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["dialog_server_failed"])

    def _refresh_qr(self, url: str) -> None:
        qr_image = make_qr_image(url)
        self.qr_photo = ImageTk.PhotoImage(qr_image)
        self.qr_label.configure(image=self.qr_photo)

    def _poll_queue(self) -> None:
        try:
            while True:
                event = self.ui_queue.get_nowait()
                self._handle_event(event)
        except queue.Empty:
            pass
        self.root.after(QUEUE_POLL_MS, self._poll_queue)

    def _handle_event(self, event: dict[str, Any]) -> None:
        status = event.get("status")
        if status:
            self.status_var.set(str(status))

        if "current_text" in event:
            self.current_file_var.set(str(event["current_text"]))
        elif "current_file" in event:
            self.current_file_var.set(str(event["current_file"]))
        if "success" in event:
            self.success_var.set(str(event["success"]))
        if "failed" in event:
            self.failed_var.set(str(event["failed"]))
        if "batch_total" in event:
            self.total_var.set(str(event["batch_total"]))
        elif "total" in event:
            self.total_var.set(str(event["total"]))
        elif "received" in event:
            self.total_var.set(str(event["received"]))
        if "failed_files" in event:
            failed_files = event.get("failed_files") or []
            self.failed_files_var.set("\n".join(map(str, failed_files)) if failed_files else UI_TEXT["failed_files_empty"])

        event_type = str(event.get("type", ""))
        if "url" in event:
            url = str(event["url"])
            self.url_var.set(url)
            self._refresh_qr(url)

        if event_type == "receive_started":
            total = max(1, int(event.get("total", 1)))
            self.current_batch_id += 1
            self.progress_total = total * 3
            self.progress.configure(maximum=self.progress_total)
            self.progress_value.set(1)
        elif event_type in {"file_processing", "file_saving", "file_converting", "file_saved", "file_failed"}:
            current = int(event.get("current", self.progress_value.get()))
            phase = int(event.get("phase", 1))
            progress = ((max(1, current) - 1) * 3) + max(1, min(3, phase))
            self.progress_value.set(min(progress, self.progress_total))
        elif event_type == "batch_complete":
            self.progress_value.set(self.progress_total)
            batch_success = int(event.get("batch_success", 0))
            if batch_success > 0:
                self._open_downloads_after_batch(self.current_batch_id)
                self.status_var.set(UI_TEXT["status_ready_next"])
                self.current_file_var.set(UI_TEXT["current_ready_next"])
            else:
                self.status_var.set(UI_TEXT["status_error"])
                self.current_file_var.set(UI_TEXT["current_no_success"])
        elif event_type == "server_stopped":
            self.progress_value.set(0)
            self.stop_thread = None
            self.stop_button.set_text_command("button_restart_receiving", self.restart_receiving)
        elif event_type == "server_restarted":
            self.progress_total = 1
            self.progress.configure(maximum=1)
            self.progress_value.set(0)
            self.failed_files_var.set(UI_TEXT["failed_files_empty"])
            self.stop_button.set_text_command("button_stop", self.stop_receiving)

    def copy_url(self) -> None:
        if not self.url_var.get():
            return
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(self.url_var.get())
            self.status_var.set(UI_TEXT["status_url_copied"])
        except Exception:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["dialog_copy_failed"])

    def open_downloads(self) -> None:
        try:
            os.startfile(get_downloads_dir())
            self.status_var.set(UI_TEXT["status_opened_downloads"])
        except Exception:
            messagebox.showerror(UI_TEXT["dialog_error_title"], UI_TEXT["dialog_open_failed"])

    def _open_downloads_after_batch(self, batch_id: int) -> None:
        if self.last_opened_batch_id == batch_id:
            return
        self.last_opened_batch_id = batch_id
        try:
            os.startfile(get_downloads_dir())
        except Exception:
            self.status_var.set(UI_TEXT["dialog_open_failed"])

    def stop_receiving(self) -> None:
        if self.server is None or self.stop_thread is not None:
            return
        self.status_var.set(UI_TEXT["status_stopping"])
        self.current_file_var.set(UI_TEXT["current_stopped"])
        self.stop_button.disable_with_text("button_restart_receiving")
        self.stop_thread = threading.Thread(target=self.server.stop, daemon=True)
        self.stop_thread.start()

    def restart_receiving(self) -> None:
        if self.server is None:
            return
        self.server.restart()

    def on_close(self) -> None:
        if self.server is not None:
            self.server.accepting.clear()
            if self.server.server is not None:
                threading.Thread(target=self.server.server.shutdown, daemon=True).start()
        self.root.destroy()


def main() -> None:
    set_windows_app_id()
    root = tk.Tk()
    DakeImageIPhoneToPCApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
