from __future__ import annotations

from collections import deque
from dataclasses import dataclass

from core.app_config import now_iso


@dataclass(frozen=True)
class NotificationItem:
    title: str
    message: str
    created_at: str


class NotificationQueue:
    def __init__(self, history_limit: int = 3) -> None:
        self.history_limit = max(1, history_limit)
        self._pending: deque[NotificationItem] = deque()
        self.history: list[NotificationItem] = []

    def push(self, title: str, message: str) -> NotificationItem:
        item = NotificationItem(title=title, message=message, created_at=now_iso())
        self._pending.append(item)
        return item

    def pop(self) -> NotificationItem | None:
        if not self._pending:
            return None
        return self._pending.popleft()

    def remember(self, item: NotificationItem) -> None:
        self.history.insert(0, item)
        del self.history[self.history_limit :]

    @property
    def pending_count(self) -> int:
        return len(self._pending)
