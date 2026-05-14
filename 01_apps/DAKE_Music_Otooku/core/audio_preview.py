# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import webbrowser
from dataclasses import dataclass
from pathlib import Path


AUDIO_PREVIEW_EXTENSIONS = {".mp3", ".wav"}
PREVIEW_SEARCH_DIRS = (
    Path("audio"),
    Path("audio") / "loop_pack",
    Path("video_bgm_pack") / "bgm" / "shorts",
    Path("video_bgm_pack") / "bgm" / "long",
    Path("video_bgm_pack") / "bgm" / "ambient",
    Path("video_bgm_pack") / "bgm" / "work",
)


@dataclass(frozen=True)
class AudioPreviewItem:
    path: Path
    label: str


@dataclass(frozen=True)
class PreviewResult:
    success: bool
    mode: str
    message: str


def find_audio_preview_items(project_root: Path) -> list[AudioPreviewItem]:
    project_root = Path(project_root)
    items: list[AudioPreviewItem] = []
    seen: set[Path] = set()
    for relative_dir in PREVIEW_SEARCH_DIRS:
        target_dir = project_root / relative_dir
        if not target_dir.exists():
            continue
        for path in sorted(target_dir.iterdir(), key=lambda item: item.name.lower()):
            if not path.is_file() or path.suffix.lower() not in AUDIO_PREVIEW_EXTENSIONS:
                continue
            resolved = path.resolve()
            if resolved in seen:
                continue
            seen.add(resolved)
            try:
                label = path.relative_to(project_root).as_posix()
            except ValueError:
                label = path.name
            items.append(AudioPreviewItem(path=path, label=label))
    return items


class AudioPreviewPlayer:
    def __init__(self) -> None:
        self._pygame = None
        self._pygame_checked = False
        self._mode = ""

    def _load_pygame(self):
        if self._pygame_checked:
            return self._pygame
        self._pygame_checked = True
        try:
            import pygame

            if not pygame.mixer.get_init():
                pygame.mixer.init()
            self._pygame = pygame
        except Exception:
            self._pygame = None
        return self._pygame

    def play(self, path: Path) -> PreviewResult:
        path = Path(path)
        if not path.exists() or path.suffix.lower() not in AUDIO_PREVIEW_EXTENSIONS:
            return PreviewResult(False, "missing", "audio file was not found")

        self.stop()

        pygame = self._load_pygame()
        if pygame is not None:
            try:
                pygame.mixer.music.load(str(path))
                pygame.mixer.music.play()
                self._mode = "pygame"
                return PreviewResult(True, "pygame", "playing with pygame")
            except Exception as exc:
                self._mode = ""
                pygame_error = str(exc)
        else:
            pygame_error = "pygame unavailable"

        if os.name == "nt" and path.suffix.lower() == ".wav":
            try:
                import winsound

                winsound.PlaySound(str(path), winsound.SND_FILENAME | winsound.SND_ASYNC)
                self._mode = "winsound"
                return PreviewResult(True, "winsound", "playing with winsound")
            except Exception as exc:
                pygame_error = f"{pygame_error}; winsound failed: {exc}"

        try:
            if os.name == "nt":
                os.startfile(str(path))
            else:
                webbrowser.open(path.resolve().as_uri())
            self._mode = "external"
            return PreviewResult(True, "external", "opened in default player")
        except Exception as exc:
            return PreviewResult(False, "failed", f"{pygame_error}; external player failed: {exc}")

    def stop(self) -> PreviewResult:
        if self._mode == "pygame" and self._pygame is not None:
            try:
                self._pygame.mixer.music.stop()
                self._mode = ""
                return PreviewResult(True, "pygame", "stopped")
            except Exception as exc:
                self._mode = ""
                return PreviewResult(False, "pygame", str(exc))

        if self._mode == "winsound":
            try:
                import winsound

                winsound.PlaySound(None, winsound.SND_PURGE)
                self._mode = ""
                return PreviewResult(True, "winsound", "stopped")
            except Exception as exc:
                self._mode = ""
                return PreviewResult(False, "winsound", str(exc))

        if self._mode == "external":
            self._mode = ""
            return PreviewResult(True, "external", "default player controls playback")

        return PreviewResult(True, "idle", "no active preview")
