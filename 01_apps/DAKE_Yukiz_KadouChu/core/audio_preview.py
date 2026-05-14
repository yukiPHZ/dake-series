from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


@dataclass
class PreviewResult:
    success: bool
    backend: str
    message: str


class AudioPreviewPlayer:
    def __init__(self, allow_startfile: bool = True) -> None:
        self.allow_startfile = allow_startfile
        self.backend = ""

    def play(self, audio_path: Path) -> PreviewResult:
        path = audio_path.resolve()
        if not path.exists() or not path.is_file():
            return PreviewResult(False, "unavailable", "Audio file was not found.")

        self.stop()
        errors: list[str] = []

        try:
            import pygame  # type: ignore[import-not-found]

            if not pygame.mixer.get_init():
                pygame.mixer.init()
            pygame.mixer.music.stop()
            pygame.mixer.music.load(str(path))
            pygame.mixer.music.play()
            self.backend = "pygame"
            return PreviewResult(True, self.backend, "Preview started.")
        except Exception as exc:
            errors.append(f"pygame: {exc}")

        if path.suffix.lower() == ".wav":
            try:
                import winsound

                winsound.PlaySound(str(path), winsound.SND_FILENAME | winsound.SND_ASYNC)
                self.backend = "winsound"
                return PreviewResult(True, self.backend, "Preview started.")
            except Exception as exc:
                errors.append(f"winsound: {exc}")

        if self.allow_startfile:
            try:
                os.startfile(str(path))  # type: ignore[attr-defined]
                self.backend = "startfile"
                return PreviewResult(True, self.backend, "Preview opened with the default app.")
            except Exception as exc:
                errors.append(f"startfile: {exc}")

        return PreviewResult(False, "unavailable", "; ".join(errors) or "Preview unavailable.")

    def stop(self) -> PreviewResult:
        backend = self.backend
        try:
            if backend == "pygame":
                import pygame  # type: ignore[import-not-found]

                if pygame.mixer.get_init():
                    pygame.mixer.music.stop()
            elif backend == "winsound":
                import winsound

                winsound.PlaySound(None, winsound.SND_PURGE)
            self.backend = ""
            if backend == "startfile":
                return PreviewResult(True, backend, "External preview was opened by the default app.")
            return PreviewResult(True, backend or "none", "Preview stopped.")
        except Exception as exc:
            self.backend = ""
            return PreviewResult(False, backend or "unknown", str(exc))
