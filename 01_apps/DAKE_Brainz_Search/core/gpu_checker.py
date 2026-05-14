from __future__ import annotations

import shutil
import subprocess
from dataclasses import dataclass


@dataclass(frozen=True)
class GpuStatus:
    nvidia_smi_available: bool
    gpu_detected: bool
    gpu_name: str
    driver_version: str
    memory_total: str
    message: str

    @property
    def cuda_available(self) -> bool:
        return self.gpu_detected


def check_gpu(timeout: float = 2.5) -> GpuStatus:
    command = shutil.which("nvidia-smi") or "nvidia-smi"
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        completed = subprocess.run(
            [
                command,
                "--query-gpu=name,driver_version,memory.total",
                "--format=csv,noheader",
            ],
            capture_output=True,
            text=True,
            timeout=timeout,
            creationflags=creationflags,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return GpuStatus(
            nvidia_smi_available=False,
            gpu_detected=False,
            gpu_name="",
            driver_version="",
            memory_total="",
            message=str(exc),
        )

    if completed.returncode != 0:
        return GpuStatus(
            nvidia_smi_available=True,
            gpu_detected=False,
            gpu_name="",
            driver_version="",
            memory_total="",
            message=(completed.stderr or completed.stdout).strip(),
        )

    line = (completed.stdout or "").strip().splitlines()[0] if completed.stdout.strip() else ""
    parts = [part.strip() for part in line.split(",")]
    gpu_name = parts[0] if len(parts) > 0 else ""
    driver_version = parts[1] if len(parts) > 1 else ""
    memory_total = parts[2] if len(parts) > 2 else ""
    return GpuStatus(
        nvidia_smi_available=True,
        gpu_detected=bool(gpu_name),
        gpu_name=gpu_name,
        driver_version=driver_version,
        memory_total=memory_total,
        message="ready" if gpu_name else "not detected",
    )
