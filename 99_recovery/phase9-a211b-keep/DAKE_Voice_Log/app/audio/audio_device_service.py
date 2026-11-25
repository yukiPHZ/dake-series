# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class AudioDevice:
    id: str
    name: str
    is_default: bool = False


class AudioDeviceService:
    def __init__(self) -> None:
        self._sd = None
        try:
            import sounddevice as sd  # type: ignore

            self._sd = sd
        except Exception:
            self._sd = None

    def is_available(self) -> bool:
        return self._sd is not None

    def input_devices(self) -> list[AudioDevice]:
        if self._sd is None:
            return []
        try:
            devices = self._sd.query_devices()
            defaults = self._sd.default.device
            default_input = defaults[0] if isinstance(defaults, (list, tuple)) else defaults
        except Exception:
            return []
        result: list[AudioDevice] = []
        for index, device in enumerate(devices):
            try:
                if int(device.get("max_input_channels", 0)) <= 0:
                    continue
                result.append(
                    AudioDevice(
                        id=str(index),
                        name=str(device.get("name") or f"Input {index}"),
                        is_default=index == default_input,
                    )
                )
            except Exception:
                continue
        return result

    def default_input_name(self) -> str:
        devices = self.input_devices()
        for device in devices:
            if device.is_default:
                return device.name
        return devices[0].name if devices else ""

