# -*- coding: utf-8 -*-
from __future__ import annotations

import threading
from pathlib import Path
from typing import Callable

from app.constants import APP_NAME, UI_TEXT


class TrayService:
    def __init__(
        self,
        icon_paths: list[Path],
        open_app: Callable[[], None],
        toggle_pause: Callable[[], None],
        is_paused: Callable[[], bool],
        open_today: Callable[[], None],
        open_folder: Callable[[], None],
        open_settings: Callable[[], None],
        quit_app: Callable[[], None],
    ) -> None:
        self.icon_paths = icon_paths
        self.open_app = open_app
        self.toggle_pause = toggle_pause
        self.is_paused = is_paused
        self.open_today = open_today
        self.open_folder = open_folder
        self.open_settings = open_settings
        self.quit_app = quit_app
        self._icon = None
        self._thread: threading.Thread | None = None
        self.available = False
        try:
            import pystray  # type: ignore
            from PIL import Image  # type: ignore

            self._pystray = pystray
            self._image_class = Image
            self.available = True
        except Exception:
            self._pystray = None
            self._image_class = None

    def start(self) -> None:
        if not self.available or self._thread:
            return
        self._thread = threading.Thread(target=self._run, name="dake-voice-tray", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        icon = self._icon
        if icon is not None:
            try:
                icon.stop()
            except Exception:
                pass

    def update(self) -> None:
        icon = self._icon
        if icon is not None:
            try:
                icon.update_menu()
            except Exception:
                pass

    def _run(self) -> None:
        if self._pystray is None or self._image_class is None:
            return
        pystray = self._pystray
        image = self._load_image()
        menu = pystray.Menu(
            pystray.MenuItem(UI_TEXT["tray_open"], lambda _icon, _item: self.open_app()),
            pystray.MenuItem(lambda _item: UI_TEXT["tray_resume"] if self.is_paused() else UI_TEXT["tray_pause"], lambda _icon, _item: self.toggle_pause()),
            pystray.MenuItem(UI_TEXT["tray_today"], lambda _icon, _item: self.open_today()),
            pystray.MenuItem(UI_TEXT["tray_folder"], lambda _icon, _item: self.open_folder()),
            pystray.MenuItem(UI_TEXT["tray_settings"], lambda _icon, _item: self.open_settings()),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(UI_TEXT["tray_exit"], lambda _icon, _item: self.quit_app()),
        )
        self._icon = pystray.Icon("DakeVoiceLog", image, APP_NAME, menu)
        self._icon.run()

    def _load_image(self):  # noqa: ANN202
        assert self._image_class is not None
        for path in self.icon_paths:
            try:
                resolved = path.resolve()
            except OSError:
                resolved = path
            if resolved.is_file():
                try:
                    return self._image_class.open(resolved)
                except Exception:
                    continue
        return self._image_class.new("RGB", (64, 64), (47, 111, 237))

