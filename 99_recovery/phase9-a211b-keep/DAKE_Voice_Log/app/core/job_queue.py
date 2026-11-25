# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta


@dataclass(frozen=True)
class RetryPolicy:
    max_retries: int = 3
    base_delay_sec: int = 60

    def can_retry(self, retries: int) -> bool:
        return retries < self.max_retries

    def next_retry_at(self, retries: int, now: datetime | None = None) -> datetime:
        current = now or datetime.now()
        delay = self.base_delay_sec * max(1, 2 ** max(0, retries - 1))
        return current + timedelta(seconds=delay)

