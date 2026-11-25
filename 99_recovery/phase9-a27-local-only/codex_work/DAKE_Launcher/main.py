# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import webbrowser
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

try:
    import ctypes
except Exception:
    ctypes = None


APP_NAME = "Dakeランチャー"
WINDOW_TITLE = "Dakeランチャー"
COPYRIGHT = "© 2026 しまりす不動産"

UI_TEXT = {
    "header_title": "DAKEツール",
    "header_description": "使うアプリを選んで起動します。",
    "recent_title": "最近使った",
    "recent_empty": "最近使ったアプリはまだありません",
    "app_list_title": "アプリ一覧",
    "button_launch": "起動",
    "button_choose_path": "場所を指定",
    "status_missing": "未検出",
    "release_link": "配布ページ",
    "dialog_choose_exe_title": "exeの場所を指定",
    "dialog_error_title": "エラー",
    "dialog_invalid_exe": "exeファイルを選択してください。",
    "dialog_launch_failed": "起動できませんでした。\n\n{error}",
    "filetype_exe": "実行ファイル",
    "filetype_all": "すべてのファイル",
    "footer_left": "シンプルそれDAKEシリーズ",
    "footer_tagline": "止まらない、迷わない、すぐ終わる。",
    "footer_link_1": "戸建買取査定",
    "footer_link_2": "Instagram",
    "footer_separator": " ｜ ",
    "footer_copyright": COPYRIGHT,
}


APP_FOLDER_NAME = "DAKE_Launcher"
CONFIG_FILENAME = "DAKE_Launcher_config.json"
README_FILENAME = "README.md"
RECENT_LIMIT = 2

COLORS = {
    "base_bg": "#F6F7F9",
    "card_bg": "#FFFFFF",
    "text": "#1E2430",
    "muted": "#667085",
    "border": "#E6EAF0",
    "accent": "#2F6FED",
    "accent_hover": "#2458BF",
    "selection_bg": "#EAF2FF",
    "missing_bg": "#FFF7E6",
    "missing_text": "#B54708",
    "idle_bg": "#F2F4F7",
    "quiet_muted": "#98A2B3",
    "scrollbar_thumb": "#E1E6EE",
    "scrollbar_thumb_hover": "#D3DBE6",
}

LINK_URLS = {
    "footer_link_1": "https://sakurayk.notion.site/22ea54b5298d80928443ec7b4d20143d?pvs=74",
    "footer_link_2": "https://www.instagram.com/kikuta.shimarisu_fudosan",
}

FONT_CANDIDATES = ["BIZ UDPGothic", "Yu Gothic UI", "Meiryo"]
DAKE_META_PATTERN = re.compile(
    r"##\s*DAKE_META\s*```(?:json)?\s*(\{.*?\})\s*```",
    re.IGNORECASE | re.DOTALL,
)


@dataclass(frozen=True)
class AppMeta:
    app_key: str
    display_name: str
    launcher_title: str
    launcher_description: str
    folder_name: str
    exe_name: str
    release_url: str
    status: str
    show_in_launcher: bool
    folder_path: Path

    @property
    def standard_exe_path(self) -> Path:
        return self.folder_path / "dist" / self.exe_name


def launcher_dir() -> Path:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        if exe_dir.name.lower() == "dist":
            return exe_dir.parent
        return exe_dir
    return Path(__file__).resolve().parent


def apps_root() -> Path:
    return launcher_dir().parent


def series_root() -> Path:
    return apps_root().parent


def config_path() -> Path:
    return series_root() / "04_data" / "configs" / CONFIG_FILENAME


def icon_path() -> Path:
    if getattr(sys, "frozen", False):
        meipass = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
        frozen_icon = meipass / "dake_icon.ico"
        if frozen_icon.exists():
            return frozen_icon
    return series_root() / "02_assets" / "dake_icon.ico"


def set_windows_app_id() -> None:
    if not sys.platform.startswith("win") or ctypes is None:
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("dake.launcher")
    except Exception:
        pass


def apply_window_icon(window: tk.Misc) -> None:
    try:
        icon = icon_path()
        if icon.exists():
            window.iconbitmap(str(icon))
            window.iconbitmap(default=str(icon))
    except Exception:
        pass


def choose_font_family(root: tk.Tk) -> str:
    try:
        available = set(tkfont.families(root))
    except Exception:
        return "TkDefaultFont"
    for family in FONT_CANDIDATES:
        if family in available:
            return family
    return "TkDefaultFont"


def read_text(path: Path) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp932"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def extract_dake_meta(readme_path: Path) -> dict | None:
    try:
        content = read_text(readme_path)
    except OSError:
        return None

    match = DAKE_META_PATTERN.search(content)
    if not match:
        return None

    try:
        meta = json.loads(match.group(1))
    except json.JSONDecodeError:
        return None

    if not isinstance(meta, dict):
        return None
    return meta


def bool_from_meta(value: object) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() == "true"
    return False


def load_config() -> dict:
    path = config_path()
    default = {"recent_apps": [], "custom_exe_paths": {}}
    if not path.exists():
        return default

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return default

    if not isinstance(data, dict):
        return default

    recent_apps = data.get("recent_apps", [])
    custom_exe_paths = data.get("custom_exe_paths", {})
    if not isinstance(recent_apps, list):
        recent_apps = []
    if not isinstance(custom_exe_paths, dict):
        custom_exe_paths = {}

    clean_recent = [str(item) for item in recent_apps if isinstance(item, str)][:RECENT_LIMIT]
    clean_paths = {
        str(key): str(value)
        for key, value in custom_exe_paths.items()
        if isinstance(key, str) and isinstance(value, str)
    }
    return {"recent_apps": clean_recent, "custom_exe_paths": clean_paths}


def save_config(config: dict) -> None:
    path = config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "recent_apps": config.get("recent_apps", [])[:RECENT_LIMIT],
        "custom_exe_paths": config.get("custom_exe_paths", {}),
    }
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def read_all_apps() -> tuple[list[AppMeta], int]:
    root = apps_root()
    apps: list[AppMeta] = []
    meta_count = 0
    if not root.exists():
        return apps, meta_count

    for folder in sorted(root.iterdir(), key=lambda item: item.name.lower()):
        if not folder.is_dir():
            continue
        if folder.name == APP_FOLDER_NAME:
            continue

        meta = extract_dake_meta(folder / README_FILENAME)
        if meta is None:
            continue

        meta_count += 1
        app_key = str(meta.get("app_key", "")).strip()
        exe_name = str(meta.get("exe_name", "")).strip()
        if not app_key or not exe_name:
            continue

        display_name = str(meta.get("display_name", app_key)).strip() or app_key
        launcher_title = str(meta.get("launcher_title", display_name)).strip() or display_name
        launcher_description = str(meta.get("launcher_description", "")).strip()
        folder_name = str(meta.get("folder_name", folder.name)).strip() or folder.name
        apps.append(
            AppMeta(
                app_key=app_key,
                display_name=display_name,
                launcher_title=launcher_title,
                launcher_description=launcher_description,
                folder_name=folder_name,
                exe_name=exe_name,
                release_url=str(meta.get("release_url", "")).strip(),
                status=str(meta.get("status", "")).strip(),
                show_in_launcher=bool_from_meta(meta.get("show_in_launcher")),
                folder_path=folder,
            )
        )

    return apps, meta_count


class LauncherApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.font_family = choose_font_family(root)
        self.config = load_config()
        self.apps: list[AppMeta] = []
        self.visible_apps: list[AppMeta] = []
        self.app_by_key: dict[str, AppMeta] = {}
        self.meta_count = 0
        self._scroll_bind_target: tk.Canvas | None = None

        self.root.title(WINDOW_TITLE)
        self.root.geometry("820x680")
        self.root.minsize(720, 560)
        self.root.configure(bg=COLORS["base_bg"])
        apply_window_icon(self.root)

        self.configure_styles()
        self.build_ui()
        self.reload_apps()

    def configure_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "Primary.TButton",
            font=(self.font_family, 10, "bold"),
            padding=(12, 8),
            background=COLORS["accent"],
            foreground="#FFFFFF",
            bordercolor=COLORS["accent"],
            lightcolor=COLORS["accent"],
            darkcolor=COLORS["accent"],
            relief="flat",
        )
        style.map(
            "Primary.TButton",
            background=[("active", COLORS["accent_hover"]), ("disabled", COLORS["idle_bg"])],
            foreground=[("disabled", COLORS["muted"])],
            bordercolor=[("active", COLORS["accent_hover"])],
        )

        style.configure(
            "Secondary.TButton",
            font=(self.font_family, 10),
            padding=(12, 8),
            background=COLORS["card_bg"],
            foreground=COLORS["text"],
            bordercolor=COLORS["border"],
            lightcolor=COLORS["card_bg"],
            darkcolor=COLORS["card_bg"],
            relief="solid",
        )
        style.map(
            "Secondary.TButton",
            background=[("active", COLORS["selection_bg"])],
            foreground=[("disabled", COLORS["muted"])],
            bordercolor=[("active", COLORS["accent"])],
        )

        style.configure(
            "Launcher.Vertical.TScrollbar",
            background=COLORS["scrollbar_thumb"],
            darkcolor=COLORS["scrollbar_thumb"],
            lightcolor=COLORS["scrollbar_thumb"],
            troughcolor=COLORS["base_bg"],
            bordercolor=COLORS["base_bg"],
            arrowcolor=COLORS["muted"],
            relief="flat",
            width=10,
        )
        style.map(
            "Launcher.Vertical.TScrollbar",
            background=[("active", COLORS["scrollbar_thumb_hover"])],
            darkcolor=[("active", COLORS["scrollbar_thumb_hover"])],
            lightcolor=[("active", COLORS["scrollbar_thumb_hover"])],
        )

    def build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=COLORS["base_bg"])
        outer.pack(fill="both", expand=True, padx=28, pady=(24, 18))

        self.build_header(outer)
        self.build_scroll_area(outer)
        self.build_footer(outer)

    def build_header(self, parent: tk.Frame) -> None:
        header = tk.Frame(parent, bg=COLORS["base_bg"])
        header.pack(fill="x", pady=(0, 20))

        tk.Label(
            header,
            text=UI_TEXT["header_title"],
            bg=COLORS["base_bg"],
            fg=COLORS["text"],
            font=(self.font_family, 22, "bold"),
            anchor="w",
        ).pack(anchor="w")

        tk.Label(
            header,
            text=UI_TEXT["header_description"],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 11),
            anchor="w",
        ).pack(anchor="w", pady=(6, 0))

    def build_scroll_area(self, parent: tk.Frame) -> None:
        shell = tk.Frame(parent, bg=COLORS["base_bg"])
        shell.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(shell, bg=COLORS["base_bg"], highlightthickness=0, bd=0)
        self.canvas.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(
            shell,
            orient="vertical",
            command=self.canvas.yview,
            style="Launcher.Vertical.TScrollbar",
        )
        scrollbar.pack(side="right", fill="y", padx=(10, 0))
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.content_frame = tk.Frame(self.canvas, bg=COLORS["base_bg"])
        self.content_window = self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw")

        self.content_frame.bind("<Configure>", self.update_scrollregion)
        self.canvas.bind("<Configure>", self.fit_content_width)
        self.canvas.bind("<Enter>", self.bind_mousewheel)
        self.canvas.bind("<Leave>", self.unbind_mousewheel)

    def build_footer(self, parent: tk.Frame) -> None:
        footer = tk.Frame(parent, bg=COLORS["base_bg"])
        footer.pack(fill="x", pady=(14, 0))

        line_one = tk.Frame(footer, bg=COLORS["base_bg"])
        line_one.pack(anchor="center")
        self.create_footer_text(line_one, "footer_left")
        self.create_footer_text(line_one, "footer_separator")
        self.create_footer_text(line_one, "footer_tagline")

        line_two = tk.Frame(footer, bg=COLORS["base_bg"])
        line_two.pack(anchor="center", pady=(4, 0))
        self.create_footer_link(line_two, "footer_link_1")
        self.create_footer_text(line_two, "footer_separator")
        self.create_footer_link(line_two, "footer_link_2")
        self.create_footer_text(line_two, "footer_separator")
        self.create_footer_text(line_two, "footer_copyright")

    def create_footer_link(self, parent: tk.Frame, text_key: str) -> None:
        label = tk.Label(
            parent,
            text=UI_TEXT[text_key],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            cursor="hand2",
            font=(self.font_family, 9),
        )
        label.pack(side="left")
        label.bind("<Button-1>", lambda _event, key=text_key: webbrowser.open(LINK_URLS[key]))
        label.bind("<Enter>", lambda _event, target=label: target.configure(fg=COLORS["accent"]))
        label.bind("<Leave>", lambda _event, target=label: target.configure(fg=COLORS["muted"]))

    def create_footer_text(self, parent: tk.Frame, text_key: str) -> None:
        tk.Label(
            parent,
            text=UI_TEXT[text_key],
            bg=COLORS["base_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 9),
        ).pack(side="left")

    def bind_mousewheel(self, _event) -> None:
        self._scroll_bind_target = self.canvas
        self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)

    def unbind_mousewheel(self, _event) -> None:
        if self._scroll_bind_target is self.canvas:
            self.canvas.unbind_all("<MouseWheel>")
            self._scroll_bind_target = None

    def on_mousewheel(self, event) -> None:
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def update_scrollregion(self, _event=None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def fit_content_width(self, event) -> None:
        self.canvas.itemconfigure(self.content_window, width=event.width)

    def reload_apps(self) -> None:
        self.apps, self.meta_count = read_all_apps()
        self.visible_apps = [app for app in self.apps if app.show_in_launcher]
        self.app_by_key = {app.app_key: app for app in self.visible_apps}
        self.cleanup_config()
        self.render_apps()

    def cleanup_config(self) -> None:
        changed = False
        known_keys = set(self.app_by_key)

        recent = []
        for app_key in self.config.get("recent_apps", []):
            if app_key in known_keys and app_key not in recent:
                recent.append(app_key)
        recent = recent[:RECENT_LIMIT]
        if recent != self.config.get("recent_apps", []):
            self.config["recent_apps"] = recent
            changed = True

        custom_paths = dict(self.config.get("custom_exe_paths", {}))
        for app_key, value in list(custom_paths.items()):
            app = self.app_by_key.get(app_key)
            if app is None:
                custom_paths.pop(app_key, None)
                changed = True
                continue

            path = Path(value)
            if not path.exists() or not path.is_file():
                custom_paths.pop(app_key, None)
                changed = True

        if custom_paths != self.config.get("custom_exe_paths", {}):
            self.config["custom_exe_paths"] = custom_paths
            changed = True

        if changed:
            save_config(self.config)

    def render_apps(self) -> None:
        for child in self.content_frame.winfo_children():
            child.destroy()

        self.render_recent_section()
        self.render_list_section()
        self.update_scrollregion()

    def render_recent_section(self) -> None:
        section = self.create_section(UI_TEXT["recent_title"], quiet=True)
        recent_apps = [
            self.app_by_key[app_key]
            for app_key in self.config.get("recent_apps", [])
            if app_key in self.app_by_key
        ]

        if not recent_apps:
            tk.Label(
                section,
                text=UI_TEXT["recent_empty"],
                bg=COLORS["base_bg"],
                fg=COLORS["border"],
                font=(self.font_family, 8),
                anchor="w",
            ).pack(fill="x", pady=(0, 8))
            return

        for app in recent_apps:
            self.create_app_card(section, app)

    def render_list_section(self) -> None:
        section = self.create_section(UI_TEXT["app_list_title"])
        for app in self.visible_apps:
            self.create_app_card(section, app)

    def create_section(self, title: str, quiet: bool = False) -> tk.Frame:
        section = tk.Frame(self.content_frame, bg=COLORS["base_bg"])
        section.pack(fill="x", pady=(0, 16))
        tk.Label(
            section,
            text=title,
            bg=COLORS["base_bg"],
            fg=COLORS["quiet_muted"] if quiet else COLORS["text"],
            font=(self.font_family, 11 if quiet else 13, "normal" if quiet else "bold"),
            anchor="w",
        ).pack(fill="x", pady=(0, 8 if quiet else 10))
        return section

    def create_app_card(self, parent: tk.Frame, app: AppMeta) -> None:
        exe_path = self.resolve_exe_path(app)
        is_missing = exe_path is None

        card = tk.Frame(
            parent,
            bg=COLORS["card_bg"],
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            bd=0,
        )
        card.pack(fill="x", pady=(0, 8))

        row = tk.Frame(card, bg=COLORS["card_bg"])
        row.pack(fill="x", padx=(22, 16), pady=11)
        row.grid_columnconfigure(0, weight=1)

        text_area = tk.Frame(row, bg=COLORS["card_bg"])
        text_area.grid(row=0, column=0, sticky="ew", padx=(0, 16))

        title_label = tk.Label(
            text_area,
            text=app.launcher_title,
            bg=COLORS["card_bg"],
            fg=COLORS["text"],
            font=(self.font_family, 12, "bold"),
            anchor="w",
            justify="left",
        )
        title_label.pack(fill="x")

        description_label = tk.Label(
            text_area,
            text=app.launcher_description,
            bg=COLORS["card_bg"],
            fg=COLORS["muted"],
            font=(self.font_family, 10),
            anchor="w",
            justify="left",
            wraplength=480,
        )
        description_label.pack(fill="x", pady=(5, 0))
        text_area.bind(
            "<Configure>",
            lambda event, label=description_label: label.configure(
                wraplength=max(220, event.width - 12)
            ),
        )

        action_area = tk.Frame(row, bg=COLORS["card_bg"])
        action_area.grid(row=0, column=1, sticky="e")

        if is_missing:
            tk.Label(
                action_area,
                text=UI_TEXT["status_missing"],
                bg=COLORS["missing_bg"],
                fg=COLORS["missing_text"],
                font=(self.font_family, 9, "bold"),
                padx=10,
                pady=4,
            ).pack(anchor="e", pady=(0, 8))

        if is_missing:
            button_style = "Secondary.TButton"
            button_text = UI_TEXT["button_choose_path"]
            button_command = lambda selected_app=app: self.choose_custom_exe(selected_app)
        else:
            button_style = "Primary.TButton"
            button_text = UI_TEXT["button_launch"]
            button_command = lambda selected_app=app: self.launch_app(selected_app)
        ttk.Button(
            action_area,
            text=button_text,
            style=button_style,
            command=button_command,
            width=12,
        ).pack(anchor="e")

        if is_missing and app.release_url:
            release_label = tk.Label(
                action_area,
                text=UI_TEXT["release_link"],
                bg=COLORS["card_bg"],
                fg=COLORS["muted"],
                cursor="hand2",
                font=(self.font_family, 9),
            )
            release_label.pack(anchor="e", pady=(8, 0))
            release_label.bind("<Button-1>", lambda _event, url=app.release_url: webbrowser.open(url))
            release_label.bind("<Enter>", lambda _event, target=release_label: target.configure(fg=COLORS["accent"]))
            release_label.bind("<Leave>", lambda _event, target=release_label: target.configure(fg=COLORS["muted"]))

    def resolve_exe_path(self, app: AppMeta) -> Path | None:
        custom = self.config.get("custom_exe_paths", {}).get(app.app_key)
        if custom:
            custom_path = Path(custom)
            if custom_path.exists() and custom_path.is_file():
                return custom_path

        standard_path = app.standard_exe_path
        if standard_path.exists() and standard_path.is_file():
            return standard_path
        return None

    def choose_custom_exe(self, app: AppMeta) -> None:
        selected = filedialog.askopenfilename(
            parent=self.root,
            title=UI_TEXT["dialog_choose_exe_title"],
            filetypes=[
                (UI_TEXT["filetype_exe"], "*.exe"),
                (UI_TEXT["filetype_all"], "*.*"),
            ],
        )
        if not selected:
            return

        path = Path(selected)
        if not path.exists() or path.suffix.lower() != ".exe":
            messagebox.showerror(
                UI_TEXT["dialog_error_title"],
                UI_TEXT["dialog_invalid_exe"],
                parent=self.root,
            )
            return

        custom_paths = self.config.setdefault("custom_exe_paths", {})
        try:
            if path.resolve() == app.standard_exe_path.resolve():
                custom_paths.pop(app.app_key, None)
            else:
                custom_paths[app.app_key] = str(path)
        except OSError:
            custom_paths[app.app_key] = str(path)

        save_config(self.config)
        self.reload_apps()

    def launch_app(self, app: AppMeta) -> None:
        exe_path = self.resolve_exe_path(app)
        if exe_path is None:
            self.choose_custom_exe(app)
            return

        try:
            subprocess.Popen([str(exe_path)], cwd=str(exe_path.parent))
        except OSError as exc:
            messagebox.showerror(
                UI_TEXT["dialog_error_title"],
                UI_TEXT["dialog_launch_failed"].format(error=exc),
                parent=self.root,
            )
            return

        self.record_recent(app.app_key)
        save_config(self.config)
        self.render_apps()

    def record_recent(self, app_key: str) -> None:
        recent = [key for key in self.config.get("recent_apps", []) if key != app_key]
        self.config["recent_apps"] = [app_key, *recent][:RECENT_LIMIT]


def main() -> None:
    set_windows_app_id()
    root = tk.Tk()
    LauncherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
