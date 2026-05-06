# -*- coding: utf-8 -*-
from __future__ import annotations

import datetime as dt
import json
import re
import sys
import traceback
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
from PIL import Image, ImageOps, UnidentifiedImageError


APP_NAME = "holiday-jinja 投稿DAKE"
CONFIG_NAME = "config.json"
DEFAULT_SITE_PATH = Path("C:/Users/yukiz/devlop/holiday-jinja-site")
PREVIEW_URL = "http://127.0.0.1:4173/"
SUPPORTED_IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp"}
POST_ID_RE = re.compile(r"^hj-(\d+)$")
MAX_IMAGE_EDGE = 2000
JPEG_QUALITY = 90

UI_FONT_TITLE = ("Yu Gothic UI", 21, "bold")
UI_FONT_SECTION = ("Yu Gothic UI", 19, "bold")
UI_FONT_LABEL = ("Yu Gothic UI", 12)
UI_FONT_INPUT = ("Yu Gothic UI", 14)
UI_FONT_BUTTON = ("Yu Gothic UI", 13, "bold")
UI_FONT_META = ("Segoe UI", 16, "bold")
UI_FONT_STATUS = ("Yu Gothic UI", 13)

COLORS = {
    "bg": "#0f1115",
    "panel": "#16181d",
    "panel_2": "#101216",
    "preview": "#050506",
    "text": "#e8e8e8",
    "sub": "#9ca3af",
    "accent": "#c9c9c9",
    "button": "#2b2d33",
    "button_hover": "#3a3d45",
    "entry": "#101216",
    "border": "#2a2d34",
}


@dataclass
class PostDraft:
    posts: list[dict]
    post_id: str
    image_rel_path: str
    image_path: Path
    title: str
    text: str
    location: str
    date: str


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = get_app_dir()
CONFIG_PATH = APP_DIR / CONFIG_NAME


class HolidayJinjaPostApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1180x760")
        self.minsize(1040, 680)
        self.configure(fg_color=COLORS["bg"])

        self.site_path: Path | None = None
        self.selected_image_path: Path | None = None
        self.preview_source: Image.Image | None = None
        self.preview_ctk_image: ctk.CTkImage | None = None

        self.site_path_var = ctk.StringVar(value="未設定")
        self.generated_id_var = ctk.StringVar(value="hj-001")
        self.date_var = ctk.StringVar(value=dt.date.today().isoformat())
        self.status_var = ctk.StringVar(value="写真を選んでください")

        self._setup_icon()
        self._build_ui()

        loaded_site = self._load_site_path()
        if loaded_site is not None and loaded_site.exists():
            self._apply_site_path(loaded_site, save=not CONFIG_PATH.exists())
        else:
            self.after(250, self.choose_site_folder)

    def _setup_icon(self) -> None:
        base = APP_DIR
        candidates = [
            base / "dake_icon.ico",
            base.parent.parent / "02_assets" / "dake_icon.ico",
            base.parent.parent.parent / "02_assets" / "dake_icon.ico",
        ]
        for icon_path in candidates:
            if icon_path.exists():
                try:
                    self.iconbitmap(str(icon_path))
                except Exception:
                    pass
                return

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=3)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)

        left = ctk.CTkFrame(self, fg_color=COLORS["panel"], corner_radius=8)
        left.grid(row=0, column=0, padx=(18, 9), pady=(18, 10), sticky="nsew")
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(1, weight=1)

        left_header = ctk.CTkFrame(left, fg_color="transparent")
        left_header.grid(row=0, column=0, padx=18, pady=(18, 12), sticky="ew")
        left_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            left_header,
            text="写真",
            text_color=COLORS["text"],
            font=UI_FONT_SECTION,
        ).grid(row=0, column=0, sticky="w")

        self.select_photo_button = ctk.CTkButton(
            left_header,
            text="写真を選択",
            command=self.select_photo,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            corner_radius=6,
            height=36,
            width=128,
            font=UI_FONT_BUTTON,
        )
        self.select_photo_button.grid(row=0, column=1, sticky="e")

        self.preview_frame = ctk.CTkFrame(left, fg_color=COLORS["preview"], corner_radius=8)
        self.preview_frame.grid(row=1, column=0, padx=18, pady=(0, 18), sticky="nsew")
        self.preview_frame.grid_rowconfigure(0, weight=1)
        self.preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_frame.bind("<Configure>", lambda _event: self._render_preview())

        self.preview_label = ctk.CTkLabel(
            self.preview_frame,
            text="写真を選んでください",
            image=None,
            text_color=COLORS["sub"],
            font=UI_FONT_INPUT,
        )
        self.preview_label.grid(row=0, column=0, padx=16, pady=16, sticky="nsew")

        right = ctk.CTkFrame(self, fg_color=COLORS["panel"], corner_radius=8)
        right.grid(row=0, column=1, padx=(9, 18), pady=(18, 10), sticky="nsew")
        right.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            right,
            text=APP_NAME,
            text_color=COLORS["text"],
            font=UI_FONT_TITLE,
        ).grid(row=0, column=0, padx=18, pady=(18, 6), sticky="w")

        site_frame = ctk.CTkFrame(right, fg_color=COLORS["panel_2"], corner_radius=8)
        site_frame.grid(row=1, column=0, padx=18, pady=(8, 14), sticky="ew")
        site_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            site_frame,
            text="holiday-jinja-site",
            text_color=COLORS["sub"],
            font=UI_FONT_LABEL,
        ).grid(row=0, column=0, padx=14, pady=(12, 2), sticky="w")

        ctk.CTkLabel(
            site_frame,
            textvariable=self.site_path_var,
            text_color=COLORS["text"],
            wraplength=330,
            justify="left",
            font=UI_FONT_INPUT,
        ).grid(row=1, column=0, padx=14, pady=(0, 12), sticky="w")

        ctk.CTkButton(
            site_frame,
            text="変更",
            command=self.choose_site_folder,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            corner_radius=6,
            width=78,
            height=30,
            font=UI_FONT_BUTTON,
        ).grid(row=0, column=1, rowspan=2, padx=(8, 14), pady=12, sticky="e")

        self.title_entry = self._add_entry(right, 2, "title")
        self.text_box = self._add_textbox(right, 3, "text")
        self.location_entry = self._add_entry(right, 4, "location")
        self.location_entry.insert(0, "Japan")

        meta = ctk.CTkFrame(right, fg_color=COLORS["panel_2"], corner_radius=8)
        meta.grid(row=5, column=0, padx=18, pady=(6, 14), sticky="ew")
        meta.grid_columnconfigure((0, 1), weight=1)

        self._add_meta_value(meta, 0, "生成されるid", self.generated_id_var)
        self._add_meta_value(meta, 1, "date", self.date_var)

        self.save_button = ctk.CTkButton(
            right,
            text="投稿を保存",
            command=self.save_post,
            fg_color=COLORS["accent"],
            hover_color="#dedede",
            text_color=COLORS["bg"],
            corner_radius=6,
            height=42,
            font=UI_FONT_BUTTON,
        )
        self.save_button.grid(row=6, column=0, padx=18, pady=(4, 10), sticky="ew")

        ctk.CTkButton(
            right,
            text="ローカルプレビューを開く",
            command=self.open_local_preview,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            corner_radius=6,
            height=38,
            font=UI_FONT_BUTTON,
        ).grid(row=7, column=0, padx=18, pady=(0, 18), sticky="ew")

        status_bar = ctk.CTkFrame(self, fg_color=COLORS["bg"], corner_radius=0)
        status_bar.grid(row=1, column=0, columnspan=2, padx=18, pady=(0, 14), sticky="ew")
        status_bar.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            status_bar,
            textvariable=self.status_var,
            text_color=COLORS["sub"],
            anchor="w",
            font=UI_FONT_STATUS,
        ).grid(row=0, column=0, sticky="ew")

    def _add_entry(self, parent: ctk.CTkFrame, row: int, label: str) -> ctk.CTkEntry:
        field = ctk.CTkFrame(parent, fg_color="transparent")
        field.grid(row=row, column=0, padx=18, pady=(6, 8), sticky="ew")
        field.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            field,
            text=label,
            text_color=COLORS["sub"],
            font=UI_FONT_LABEL,
        ).grid(row=0, column=0, pady=(0, 4), sticky="w")

        entry = ctk.CTkEntry(
            field,
            fg_color=COLORS["entry"],
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            corner_radius=6,
            height=36,
            font=UI_FONT_INPUT,
        )
        entry.grid(row=1, column=0, sticky="ew")
        return entry

    def _add_textbox(self, parent: ctk.CTkFrame, row: int, label: str) -> ctk.CTkTextbox:
        field = ctk.CTkFrame(parent, fg_color="transparent")
        field.grid(row=row, column=0, padx=18, pady=(6, 8), sticky="ew")
        field.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            field,
            text=label,
            text_color=COLORS["sub"],
            font=UI_FONT_LABEL,
        ).grid(row=0, column=0, pady=(0, 4), sticky="w")

        textbox = ctk.CTkTextbox(
            field,
            fg_color=COLORS["entry"],
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            corner_radius=6,
            height=120,
            wrap="word",
            font=UI_FONT_INPUT,
        )
        textbox.grid(row=1, column=0, sticky="ew")
        return textbox

    def _add_meta_value(self, parent: ctk.CTkFrame, col: int, label: str, value: ctk.StringVar) -> None:
        ctk.CTkLabel(
            parent,
            text=label,
            text_color=COLORS["sub"],
            font=UI_FONT_LABEL,
        ).grid(row=0, column=col, padx=14, pady=(12, 2), sticky="w")

        ctk.CTkLabel(
            parent,
            textvariable=value,
            text_color=COLORS["text"],
            font=UI_FONT_META,
        ).grid(row=1, column=col, padx=14, pady=(0, 12), sticky="w")

    def _load_site_path(self) -> Path | None:
        if not CONFIG_PATH.exists():
            return None

        try:
            config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            raw_path = config.get("site_path", "")
            return Path(raw_path).expanduser() if raw_path else None
        except Exception:
            self._set_status("エラー: config.json を読み込めません")
            return None

    def _save_config(self) -> None:
        if self.site_path is None:
            return
        CONFIG_PATH.write_text(
            json.dumps({"site_path": self.site_path.as_posix()}, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

    def choose_site_folder(self) -> None:
        initial_dir = str(self.site_path or DEFAULT_SITE_PATH.parent)
        folder = filedialog.askdirectory(
            title="holiday-jinja-site フォルダを選択",
            initialdir=initial_dir,
        )
        if not folder:
            if self.site_path is None:
                self._set_status("holiday-jinja-site フォルダを選択してください")
            return

        path = Path(folder)
        if not path.exists():
            self._set_status("エラー: フォルダが見つかりません")
            return

        if not (path / "posts.json").exists():
            self._set_status("エラー: posts.json が見つかりません")
            messagebox.showerror(APP_NAME, "選択したフォルダに posts.json が見つかりません。")
            return

        self._apply_site_path(path, save=True)

    def _apply_site_path(self, path: Path, save: bool) -> None:
        self.site_path = path.resolve()
        self.site_path_var.set(str(self.site_path))
        if save:
            self._save_config()
        self.refresh_next_id()

    def refresh_next_id(self) -> None:
        if self.site_path is None:
            self.generated_id_var.set("未設定")
            return

        try:
            posts = self._read_posts()
            self.generated_id_var.set(self._next_post_id(posts))
        except FileNotFoundError:
            self.generated_id_var.set("未設定")
            self._set_status("エラー: posts.json が見つかりません")
        except Exception:
            self.generated_id_var.set("未設定")
            self._set_status("エラー: posts.json を読み込めません")

    def select_photo(self) -> None:
        filename = filedialog.askopenfilename(
            title="投稿する写真を選択",
            filetypes=[
                ("Image files", "*.jpg *.jpeg *.png *.webp"),
                ("All files", "*.*"),
            ],
        )
        if filename:
            self.load_photo(Path(filename))

    def load_photo(self, path: Path) -> None:
        if path.suffix.lower() not in SUPPORTED_IMAGE_EXTS:
            self._set_status("エラー: jpg / png / webp を選んでください")
            return

        self._set_status("写真を読み込んでいます...")
        try:
            with Image.open(path) as image:
                self.preview_source = ImageOps.exif_transpose(image).copy()
            self.selected_image_path = path
            self._render_preview()
            self._set_status("写真を読み込みました")
        except (UnidentifiedImageError, OSError):
            self.selected_image_path = None
            self.preview_source = None
            self._set_status("エラー: 写真を読み込めません")

    def _render_preview(self) -> None:
        if self.preview_source is None:
            return

        width = max(self.preview_frame.winfo_width() - 32, 320)
        height = max(self.preview_frame.winfo_height() - 32, 260)
        if width <= 1 or height <= 1:
            self.after(100, self._render_preview)
            return

        image = self.preview_source.copy()
        image.thumbnail((width, height), Image.Resampling.LANCZOS)
        self.preview_ctk_image = ctk.CTkImage(
            light_image=image,
            dark_image=image,
            size=image.size,
        )
        self.preview_label.configure(image=self.preview_ctk_image, text="")

    def save_post(self) -> None:
        self._set_status("保存を準備しています...")
        saved_image_path: Path | None = None

        try:
            draft = self._prepare_post_draft()
            if not self._confirm_save(draft):
                self._set_status("保存をキャンセルしました")
                return

            draft.image_path.parent.mkdir(exist_ok=True)
            self._save_jpeg(draft.image_path)
            saved_image_path = draft.image_path
            self._set_status(f"{draft.image_rel_path} を保存しました")

            new_post = {
                "id": draft.post_id,
                "image": draft.image_rel_path,
                "title": draft.title,
                "text": draft.text,
                "location": draft.location,
                "date": draft.date,
                "tags": [],
            }
            draft.posts.append(new_post)
            self._write_posts(draft.posts)
            self._set_status("posts.json を更新しました")

            self._clear_form_after_save()
            self.refresh_next_id()
            self._show_saved_dialog(draft)
            self._set_status("投稿を保存しました")
        except Exception as exc:
            if saved_image_path is not None and saved_image_path.exists():
                try:
                    saved_image_path.unlink()
                except Exception:
                    pass
            self._set_status(f"エラー: {exc}")
            traceback.print_exc()

    def _prepare_post_draft(self) -> PostDraft:
        if self.site_path is None:
            raise RuntimeError("holiday-jinja-site フォルダを選択してください")
        if self.selected_image_path is None:
            raise RuntimeError("写真を選んでください")

        title = self.title_entry.get().strip()
        text = self.text_box.get("1.0", "end").strip()
        location = self.location_entry.get().strip()
        if not title:
            raise RuntimeError("title を入力してください")
        if not text:
            raise RuntimeError("text を入力してください")
        if not location:
            raise RuntimeError("location を入力してください")

        posts = self._read_posts()
        post_id = self._next_post_id(posts)
        image_rel_path = f"images/{post_id}.jpg"
        image_path = self.site_path / "images" / f"{post_id}.jpg"
        if image_path.exists():
            raise RuntimeError(f"画像ファイルが既に存在します: {image_path.name}")

        return PostDraft(
            posts=posts,
            post_id=post_id,
            image_rel_path=image_rel_path,
            image_path=image_path,
            title=title,
            text=text,
            location=location,
            date=self.date_var.get(),
        )

    def _confirm_save(self, draft: PostDraft) -> bool:
        result = {"save": False}
        dialog = ctk.CTkToplevel(self)
        dialog.title("保存前の確認")
        dialog.configure(fg_color=COLORS["bg"])
        dialog.transient(self)
        dialog.resizable(False, False)
        self._center_dialog(dialog, 560, 560)

        container = ctk.CTkFrame(dialog, fg_color=COLORS["panel"], corner_radius=8)
        container.pack(fill="both", expand=True, padx=18, pady=18)
        container.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            container,
            text="保存前の確認",
            text_color=COLORS["text"],
            font=UI_FONT_TITLE,
        ).grid(row=0, column=0, padx=18, pady=(18, 4), sticky="w")

        ctk.CTkLabel(
            container,
            text="この内容で holiday-jinja に投稿を置きます。",
            text_color=COLORS["sub"],
            font=UI_FONT_LABEL,
        ).grid(row=1, column=0, padx=18, pady=(0, 14), sticky="w")

        rows = [
            ("生成されるID", draft.post_id),
            ("保存画像名", draft.image_rel_path),
            ("title", draft.title),
            ("text", draft.text),
            ("location", draft.location),
            ("date", draft.date),
        ]
        for index, (label, value) in enumerate(rows, start=2):
            self._add_dialog_row(container, index, label, value)

        buttons = ctk.CTkFrame(container, fg_color="transparent")
        buttons.grid(row=8, column=0, padx=18, pady=(16, 18), sticky="ew")
        buttons.grid_columnconfigure((0, 1), weight=1)

        def cancel() -> None:
            result["save"] = False
            dialog.destroy()

        def save() -> None:
            result["save"] = True
            dialog.destroy()

        ctk.CTkButton(
            buttons,
            text="キャンセル",
            command=cancel,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            corner_radius=6,
            height=38,
            font=UI_FONT_BUTTON,
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")

        ctk.CTkButton(
            buttons,
            text="保存する",
            command=save,
            fg_color=COLORS["accent"],
            hover_color="#dedede",
            text_color=COLORS["bg"],
            corner_radius=6,
            height=38,
            font=UI_FONT_BUTTON,
        ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

        dialog.bind("<Escape>", lambda _event: cancel())
        dialog.bind("<Return>", lambda _event: save())
        dialog.wait_visibility()
        dialog.grab_set()
        dialog.focus_force()
        self.wait_window(dialog)
        return result["save"]

    def _show_saved_dialog(self, draft: PostDraft) -> None:
        dialog = ctk.CTkToplevel(self)
        dialog.title("保存しました")
        dialog.configure(fg_color=COLORS["bg"])
        dialog.transient(self)
        dialog.resizable(False, False)
        self._center_dialog(dialog, 540, 390)

        container = ctk.CTkFrame(dialog, fg_color=COLORS["panel"], corner_radius=8)
        container.pack(fill="both", expand=True, padx=18, pady=18)
        container.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            container,
            text="保存しました",
            text_color=COLORS["text"],
            font=UI_FONT_TITLE,
        ).grid(row=0, column=0, padx=18, pady=(18, 10), sticky="w")

        messages = [
            "投稿を保存しました",
            draft.image_rel_path,
            "posts.json を更新しました",
        ]
        for index, message in enumerate(messages, start=1):
            ctk.CTkLabel(
                container,
                text=message,
                text_color=COLORS["text"] if index == 1 else COLORS["sub"],
                font=UI_FONT_INPUT if index == 1 else UI_FONT_LABEL,
                anchor="w",
            ).grid(row=index, column=0, padx=18, pady=(0, 8), sticky="ew")

        actions = ctk.CTkFrame(container, fg_color="transparent")
        actions.grid(row=4, column=0, padx=18, pady=(14, 8), sticky="ew")
        actions.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(
            actions,
            text="ローカルプレビューを開く",
            command=self.open_local_preview,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            corner_radius=6,
            height=38,
            font=UI_FONT_BUTTON,
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")

        ctk.CTkButton(
            actions,
            text="imagesフォルダを開く",
            command=self.open_images_folder,
            fg_color=COLORS["button"],
            hover_color=COLORS["button_hover"],
            text_color=COLORS["text"],
            corner_radius=6,
            height=38,
            font=UI_FONT_BUTTON,
        ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

        ctk.CTkButton(
            container,
            text="閉じる",
            command=dialog.destroy,
            fg_color=COLORS["accent"],
            hover_color="#dedede",
            text_color=COLORS["bg"],
            corner_radius=6,
            height=38,
            font=UI_FONT_BUTTON,
        ).grid(row=5, column=0, padx=18, pady=(10, 18), sticky="ew")

        dialog.bind("<Escape>", lambda _event: dialog.destroy())
        dialog.wait_visibility()
        dialog.grab_set()
        dialog.focus_force()
        self.wait_window(dialog)

    def _add_dialog_row(self, parent: ctk.CTkFrame, row: int, label: str, value: str) -> None:
        frame = ctk.CTkFrame(parent, fg_color=COLORS["panel_2"], corner_radius=6)
        frame.grid(row=row, column=0, padx=18, pady=(0, 8), sticky="ew")
        frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            frame,
            text=label,
            text_color=COLORS["sub"],
            font=UI_FONT_LABEL,
            width=110,
            anchor="w",
        ).grid(row=0, column=0, padx=(12, 10), pady=10, sticky="nw")

        ctk.CTkLabel(
            frame,
            text=value,
            text_color=COLORS["text"],
            font=UI_FONT_INPUT,
            wraplength=360,
            justify="left",
            anchor="w",
        ).grid(row=0, column=1, padx=(0, 12), pady=10, sticky="ew")

    def _center_dialog(self, dialog: ctk.CTkToplevel, width: int, height: int) -> None:
        self.update_idletasks()
        x = self.winfo_x() + max((self.winfo_width() - width) // 2, 0)
        y = self.winfo_y() + max((self.winfo_height() - height) // 2, 0)
        dialog.geometry(f"{width}x{height}+{x}+{y}")

    def _clear_form_after_save(self) -> None:
        self.title_entry.delete(0, "end")
        self.text_box.delete("1.0", "end")
        self.selected_image_path = None
        self.preview_source = None
        self.preview_ctk_image = None
        self.preview_label.configure(image=None, text="写真を選んでください")

    def _save_jpeg(self, destination: Path) -> None:
        if self.selected_image_path is None:
            raise RuntimeError("写真を選んでください")

        with Image.open(self.selected_image_path) as source:
            image = ImageOps.exif_transpose(source)
            image.thumbnail((MAX_IMAGE_EDGE, MAX_IMAGE_EDGE), Image.Resampling.LANCZOS)
            if image.mode not in ("RGB", "L"):
                background = Image.new("RGB", image.size, COLORS["bg"])
                rgba = image.convert("RGBA")
                background.paste(rgba, mask=rgba.getchannel("A"))
                image = background
            else:
                image = image.convert("RGB")

            image.save(
                destination,
                format="JPEG",
                quality=JPEG_QUALITY,
                optimize=True,
                progressive=True,
            )

    def _read_posts(self) -> list[dict]:
        if self.site_path is None:
            raise RuntimeError("holiday-jinja-site フォルダを選択してください")

        posts_path = self.site_path / "posts.json"
        if not posts_path.exists():
            raise FileNotFoundError("posts.json が見つかりません")

        data = json.loads(posts_path.read_text(encoding="utf-8"))
        if not isinstance(data, list):
            raise ValueError("posts.json は配列形式ではありません")
        return data

    def _write_posts(self, posts: list[dict]) -> None:
        if self.site_path is None:
            raise RuntimeError("holiday-jinja-site フォルダを選択してください")

        posts_path = self.site_path / "posts.json"
        text = json.dumps(posts, ensure_ascii=False, indent=2) + "\n"
        json.loads(text)
        temp_path = posts_path.with_name("posts.json.tmp")
        temp_path.write_text(text, encoding="utf-8")
        json.loads(temp_path.read_text(encoding="utf-8"))
        temp_path.replace(posts_path)

    def _next_post_id(self, posts: list[dict]) -> str:
        max_number = 0
        width = 3
        for post in posts:
            if not isinstance(post, dict):
                continue
            match = POST_ID_RE.match(str(post.get("id", "")))
            if match:
                number_text = match.group(1)
                max_number = max(max_number, int(number_text))
                width = max(width, len(number_text))
        return f"hj-{max_number + 1:0{width}d}"

    def open_local_preview(self) -> None:
        webbrowser.open(PREVIEW_URL)
        self._set_status("ローカルプレビューを開きました")

    def open_images_folder(self) -> None:
        if self.site_path is None:
            self._set_status("エラー: holiday-jinja-site フォルダを選択してください")
            return

        images_dir = self.site_path / "images"
        if not images_dir.exists():
            self._set_status("エラー: images フォルダが見つかりません")
            return

        webbrowser.open(images_dir.resolve().as_uri())
        self._set_status("imagesフォルダを開きました")

    def _set_status(self, message: str) -> None:
        self.status_var.set(message)
        self.update_idletasks()


def main() -> None:
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")
    app = HolidayJinjaPostApp()
    app.mainloop()


if __name__ == "__main__":
    main()
