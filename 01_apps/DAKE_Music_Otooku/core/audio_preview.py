# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import webbrowser
from dataclasses import dataclass
from pathlib import Path


AUDIO_PREVIEW_EXTENSIONS = {".mp3", ".wav"}
PREVIEW_PRIORITY = (
    Path("audio") / "generated_preview_A.wav",
    Path("audio") / "generated_preview_B.wav",
    Path("audio") / "generated_preview_C.wav",
    Path("audio") / "generated_preview.wav",
    Path("audio") / "generated_preview_A.mp3",
    Path("audio") / "generated_preview_B.mp3",
    Path("audio") / "generated_preview_C.mp3",
    Path("audio") / "generated.wav",
    Path("audio") / "loop_preview.wav",
    Path("audio") / "generated_preview.mp3",
)
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
    path: Path | None = None
    size_bytes: int = 0


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
    priority = {item.as_posix(): index for index, item in enumerate(PREVIEW_PRIORITY)}
    return sorted(
        items,
        key=lambda item: (
            priority.get(item.label, len(priority)),
            0 if item.path.suffix.lower() == ".wav" else 1,
            item.label.lower(),
        ),
    )


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
        path = Path(path).resolve()
        if not path.exists():
            return PreviewResult(False, "missing", "audio file was not found", path=path)
        if path.suffix.lower() not in AUDIO_PREVIEW_EXTENSIONS:
            return PreviewResult(False, "unsupported", "audio extension is not supported", path=path)
        size_bytes = path.stat().st_size
        if size_bytes <= 0:
            return PreviewResult(False, "empty", "audio file size is 0", path=path, size_bytes=size_bytes)

        self.stop()

        winsound_error = ""
        if os.name == "nt" and path.suffix.lower() == ".wav":
            try:
                import winsound

                winsound.PlaySound(str(path), winsound.SND_FILENAME | winsound.SND_ASYNC)
                self._mode = "winsound"
                return PreviewResult(True, "winsound", "playing with winsound", path=path, size_bytes=size_bytes)
            except Exception as exc:
                winsound_error = f"winsound failed: {exc}"

        if path.suffix.lower() == ".mp3":
            sibling_wav = path.with_suffix(".wav")
            generated_wav = path.parent / "generated_preview.wav"
            for wav_candidate in (sibling_wav, generated_wav):
                if wav_candidate.exists() and wav_candidate.stat().st_size > 0:
                    return self.play(wav_candidate)

        pygame = self._load_pygame()
        if pygame is not None:
            try:
                pygame.mixer.music.load(str(path))
                pygame.mixer.music.play()
                self._mode = "pygame"
                return PreviewResult(True, "pygame", "playing with pygame", path=path, size_bytes=size_bytes)
            except Exception as exc:
                self._mode = ""
                pygame_error = f"pygame failed: {exc}"
        else:
            pygame_error = "pygame unavailable"

        try:
            if os.name == "nt":
                os.startfile(str(path))
            else:
                webbrowser.open(path.resolve().as_uri())
            self._mode = "external"
            message = "opened in default player"
            return PreviewResult(True, "external", message, path=path, size_bytes=size_bytes)
        except Exception as exc:
            details = "; ".join(part for part in (winsound_error, pygame_error, f"external player failed: {exc}") if part)
            return PreviewResult(False, "failed", details, path=path, size_bytes=size_bytes)

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
