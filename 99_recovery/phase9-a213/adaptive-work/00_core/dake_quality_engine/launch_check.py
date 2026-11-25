from __future__ import annotations

import sys
from collections.abc import Callable, Iterable

try:
    from .logging import write_debug_log
    from .reliability import format_user_error
except ImportError:  # pragma: no cover - direct script fallback.
    from logging import write_debug_log  # type: ignore
    from reliability import format_user_error  # type: ignore


def run_launch_check(
    *,
    checks: Iterable[Callable[[], object]] | None = None,
    create_window: Callable[[], object] | None = None,
    log_dir: str | None = None,
) -> int:
    """Run a short launch check and return process-friendly exit code."""
    try:
        for check in checks or ():
            check()

        window = create_window() if create_window is not None else None
        if window is not None:
            try:
                update = getattr(window, "update", None)
                destroy = getattr(window, "destroy", None)
                if callable(update):
                    update()
                if callable(destroy):
                    destroy()
            except Exception:
                raise
        return 0
    except Exception as exc:
        write_debug_log("launch check failed", log_dir=log_dir, exc=exc)
        print(format_user_error(exc), file=sys.stderr)
        return 1
