from __future__ import annotations

import sys
from dataclasses import dataclass
from typing import Callable, Generic, TypeVar

try:
    from .logging import write_debug_log
except ImportError:  # pragma: no cover - direct script fallback.
    from logging import write_debug_log  # type: ignore

T = TypeVar("T")


@dataclass(frozen=True)
class SafeRunResult(Generic[T]):
    ok: bool
    value: T | None = None
    error: BaseException | None = None
    user_message: str = ""


def format_user_error(exc: BaseException) -> str:
    if isinstance(exc, PermissionError):
        return "保存または読み込みの権限がありません。場所を変えてもう一度お試しください。"
    if isinstance(exc, FileNotFoundError):
        return "対象ファイルが見つかりません。移動や削除がないか確認してください。"
    if isinstance(exc, ValueError):
        return "入力内容を確認してください。"
    return "処理を完了できませんでした。もう一度お試しください。"


def show_user_error(
    message: str,
    *,
    title: str = "DAKE",
    presenter: Callable[[str, str], None] | None = None,
    use_messagebox: bool = True,
) -> None:
    if presenter is not None:
        presenter(title, message)
        return

    if use_messagebox:
        try:
            from tkinter import messagebox

            messagebox.showerror(title, message)
            return
        except Exception:
            pass

    print(message, file=sys.stderr)


def safe_run(
    func: Callable[..., T],
    *args: object,
    title: str = "DAKE",
    log_dir: str | None = None,
    show_error: bool = False,
    presenter: Callable[[str, str], None] | None = None,
    **kwargs: object,
) -> SafeRunResult[T]:
    try:
        return SafeRunResult(ok=True, value=func(*args, **kwargs))
    except (PermissionError, FileNotFoundError, ValueError, Exception) as exc:
        message = format_user_error(exc)
        write_debug_log("safe_run failed", log_dir=log_dir, exc=exc, context={"title": title})
        if show_error:
            show_user_error(message, title=title, presenter=presenter)
        return SafeRunResult(ok=False, error=exc, user_message=message)
