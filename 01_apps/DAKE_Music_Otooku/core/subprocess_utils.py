# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import subprocess
from typing import Any


def run_hidden(command: list[str], **kwargs: Any) -> subprocess.CompletedProcess:
    """Run external tools without opening a Windows console window."""
    if os.name == "nt":
        kwargs.setdefault("creationflags", subprocess.CREATE_NO_WINDOW)
    return subprocess.run(command, **kwargs)
