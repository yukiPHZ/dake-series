# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import random
import sys
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path
import tkinter as tk


APP_NAME = "DAKE_Game_Diver_Catch"
WINDOW_TITLE = "Dake潜って獲る"
COPYRIGHT = "Copyright 2026 DAKE / Yukihiko Kikuta"

UI_TEXT = {
    "display_name": "Dake潜って獲る",
    "score_label": "SCORE",
    "best_label": "BEST",
    "life_label": "LIFE",
    "air_label": "AIR",
    "time_label": "TIME",
    "hold_label": "HOLD",
    "mode_label": "MODE",
    "mode_normal": "NORMAL",
    "mode_timeless": "TIMELESS",
    "hold_small_fish": "小魚",
    "hold_big_fish": "大魚",
    "hold_squid": "イカ",
    "hold_treasure": "宝箱",
    "ready_hint": "ENTER START",
    "playing_hint": "欲張りすぎず、船へ戻る",
    "paused_title": "PAUSED",
    "paused_hint": "ENTER RESUME",
    "game_over_title": "GAME OVER",
    "game_over_hint": "ENTER / R RESTART",
    "time_up_title": "TIME UP",
    "status_ready": "ENTERで開始",
    "status_diving": "潜水開始",
    "status_need_water": "海中で銛を撃つ",
    "status_harpoon": "銛を突いた",
    "status_harpoon_busy": "銛を戻す",
    "status_hold_full": "持ちきれない",
    "status_return_ship": "船へ戻れ",
    "status_missed": "銛は外れた",
    "status_caught": "{item}を獲った",
    "status_more_hold": "もう少し持てる",
    "status_heavy": "重い",
    "status_big_catch": "大漁だ",
    "status_big_score": "大漁だ +{points}",
    "status_scored": "持ち帰り成功 +{points}",
    "status_hit": "サメに接触",
    "status_hit_lost": "サメに接触、獲物を失った",
    "status_air_low": "息が苦しい",
    "status_air_out": "酸素切れ",
    "status_time_up": "TIME UP",
    "status_kraken_warning": "クラーケン接近",
    "status_kraken_close": "逃げろ",
    "status_kraken_hit": "海の王者に接触",
    "status_mode_changed": "Tでモード切替",
    "status_game_over": "GAME OVER",
    "status_paused": "一時停止",
    "status_resumed": "再開",
    "save_failed": "BEST保存に失敗",
    "launch_check_ok": "LAUNCH CHECK OK",
}

CONFIG_NAME = "diver_catch_config.json"
WINDOW_WIDTH = 500
WINDOW_HEIGHT = 650
GRID_COLUMNS = 7
GRID_ROWS = 8
GRID_LEFT = 40
GRID_TOP = 138
CELL_WIDTH = 60
CELL_HEIGHT = 45
SHIP_Y = 68
SEA_TOP = GRID_TOP - 12
SEA_BOTTOM = GRID_TOP + GRID_ROWS * CELL_HEIGHT
FOOTER_TOP = 538
FOOTER_BOTTOM = 640
MAX_LIFE = 3
HOLD_CAPACITY = 3
HARPOON_DISPLAY_SECONDS = 0.16
SHARK_BASE_SECONDS = 0.94
SHARK_MIN_SECONDS = 0.42
AIR_MAX = 100.0
NORMAL_TIME_LIMIT = 60.0
TREASURE_MIN_ROW = GRID_ROWS - 2
SINK_BASE_SECONDS = 2.4
SINK_TREASURE_SECONDS = 1.55
KRAKEN_START_SECONDS = 18.0
KRAKEN_DURATION_SECONDS = 13.0
KRAKEN_MOVE_SECONDS = 1.15
KRAKEN_NORMAL_CHANCE = 0.62
KRAKEN_TIMELESS_CHANCE = 0.72
KRAKEN_MAX_SPAWNS_NORMAL = 1
TIMELESS_MAX_SHARKS = 4

PREY_POINTS = {
    "small_fish": 10,
    "big_fish": 30,
    "squid": 50,
    "treasure": 100,
}

PREY_WEIGHT = {
    "small_fish": 1,
    "big_fish": 2,
    "squid": 2,
    "treasure": 3,
}

PREY_MOVE_SECONDS = {
    "small_fish": 1.15,
    "big_fish": 1.7,
    "squid": 1.45,
}

COLORS = {
    "lcd_bg": "#c4d6a3",
    "lcd_line": "#a9bd86",
    "lcd_shadow": "#91a873",
    "lcd_detail": "#9fb680",
    "ink": "#1f2c1c",
    "ink_soft": "#31422b",
    "panel": "#b4c98f",
    "water": "#b8cc95",
    "highlight": "#e2edbf",
}

FONTS = {
    "title": ("Yu Gothic UI", 22, "bold"),
    "lcd_large": ("Consolas", 18, "bold"),
    "lcd": ("Consolas", 12, "bold"),
    "label": ("Yu Gothic UI", 10, "bold"),
    "small": ("Yu Gothic UI", 8),
}


@dataclass
class Prey:
    kind: str
    col: int
    row: int
    direction: int = 1
    timer: float = 0.0


@dataclass
class Shark:
    col: int
    row: int
    direction: int
    timer: float = 0.0


@dataclass
class Harpoon:
    col: int
    row: int
    direction: int
    timer: float = HARPOON_DISPLAY_SECONDS


@dataclass
class Kraken:
    col: int
    row: int
    timer: float = KRAKEN_DURATION_SECONDS
    move_timer: float = KRAKEN_MOVE_SECONDS


@dataclass
class BackgroundItem:
    kind: str
    col: int
    row: int
    variant: int = 0


@dataclass
class GameModel:
    best_score: int = 0
    random: random.Random = field(default_factory=random.Random)
    state: str = "idle"
    mode: str = "normal"
    score: int = 0
    life: int = MAX_LIFE
    air: float = AIR_MAX
    time_remaining: float = NORMAL_TIME_LIMIT
    hold: list[str] = field(default_factory=list)
    diver_col: int = GRID_COLUMNS // 2
    diver_row: int = -1
    facing: int = 1
    elapsed: float = 0.0
    prey: list[Prey] = field(default_factory=list)
    sharks: list[Shark] = field(default_factory=list)
    harpoon: Harpoon | None = None
    kraken: Kraken | None = None
    background_items: list[BackgroundItem] = field(default_factory=list)
    last_shark_rows: list[int] = field(default_factory=list)
    move_cooldown: float = 0.0
    inactivity_timer: float = 0.0
    kraken_spawn_timer: float = KRAKEN_START_SECONDS
    kraken_spawn_count: int = 0
    score_popup_value: int = 0
    score_popup_timer: float = 0.0
    message_key: str = "status_ready"
    message_args: dict[str, object] = field(default_factory=dict)
    best_dirty: bool = False

    def start_game(self) -> None:
        self.state = "playing"
        self.score = 0
        self.life = MAX_LIFE
        self.air = AIR_MAX
        self.time_remaining = NORMAL_TIME_LIMIT
        self.hold.clear()
        self.diver_col = GRID_COLUMNS // 2
        self.diver_row = -1
        self.facing = 1
        self.elapsed = 0.0
        self.prey.clear()
        self.sharks.clear()
        self.harpoon = None
        self.kraken = None
        self.last_shark_rows.clear()
        self.move_cooldown = 0.0
        self.inactivity_timer = 0.0
        self.kraken_spawn_timer = KRAKEN_START_SECONDS + self.random.uniform(4.0, 13.0)
        self.kraken_spawn_count = 0
        self.score_popup_value = 0
        self.score_popup_timer = 0.0
        self.message_key = "status_diving"
        self.message_args = {}
        self._generate_background()
        self._ensure_sharks()
        self._ensure_prey(min_count=2)

    def toggle_mode(self) -> None:
        if self.state not in {"idle", "game_over"}:
            return
        self.mode = "timeless" if self.mode == "normal" else "normal"
        self.time_remaining = NORMAL_TIME_LIMIT
        self.message_key = "status_mode_changed"
        self.message_args = {}

    def resume(self) -> None:
        if self.state == "paused":
            self.state = "playing"
            self.message_key = "status_resumed"
            self.message_args = {}

    def toggle_pause(self) -> None:
        if self.state == "playing":
            self.state = "paused"
            self.message_key = "status_paused"
            self.message_args = {}
        elif self.state == "paused":
            self.resume()

    def hold_used(self) -> int:
        return sum(PREY_WEIGHT[item] for item in self.hold)

    def hold_space_left(self) -> int:
        return HOLD_CAPACITY - self.hold_used()

    def can_hold(self, kind: str) -> bool:
        return PREY_WEIGHT[kind] <= self.hold_space_left()

    def move_penalty_level(self) -> int:
        used = self.hold_used()
        if "treasure" in self.hold:
            return 2
        if used >= HOLD_CAPACITY:
            return 1
        if any(item in {"big_fish", "squid"} for item in self.hold):
            return 1
        return 0

    def move_delay(self) -> float:
        level = self.move_penalty_level()
        if level == 2:
            return 0.22 if self.mode == "timeless" else 0.18
        if level == 1:
            return 0.13 if self.mode == "timeless" else 0.10
        return 0.0

    def move(self, delta_col: int, delta_row: int) -> bool:
        if self.state != "playing":
            return False

        if delta_col:
            self.facing = 1 if delta_col > 0 else -1

        if self.move_cooldown > 0:
            if self.move_penalty_level():
                self.message_key = "status_heavy"
                self.message_args = {}
            return False

        moved = False
        if self.diver_row < 0:
            old_col = self.diver_col
            self.diver_col = clamp(self.diver_col + delta_col, 0, GRID_COLUMNS - 1)
            moved = self.diver_col != old_col
            if delta_row > 0:
                self.diver_row = 0
                self.message_key = "status_diving"
                self.message_args = {}
                moved = True
        else:
            old_col = self.diver_col
            old_row = self.diver_row
            self.diver_col = clamp(self.diver_col + delta_col, 0, GRID_COLUMNS - 1)
            next_row = self.diver_row + delta_row
            if next_row < 0:
                self.diver_row = -1
                self._deliver_hold()
                moved = True
            else:
                self.diver_row = clamp(next_row, 0, GRID_ROWS - 1)
                moved = (self.diver_col, self.diver_row) != (old_col, old_row)
            self._check_shark_contact()

        if moved:
            self.inactivity_timer = 0.0
            self.move_cooldown = self.move_delay()
        return moved

    def fire_harpoon(self) -> None:
        if self.state != "playing":
            return
        if self.diver_row < 0:
            self.message_key = "status_need_water"
            self.message_args = {}
            return
        if self.harpoon is not None:
            self.message_key = "status_harpoon_busy"
            self.message_args = {}
            return

        start_col = self.diver_col + self.facing
        if not is_cell_in_grid(start_col, self.diver_row):
            self.message_key = "status_missed"
            self.message_args = {}
            return

        self.harpoon = Harpoon(col=start_col, row=self.diver_row, direction=self.facing)
        self.inactivity_timer = 0.0
        self.message_key = "status_harpoon"
        self.message_args = {}
        if not self._catch_prey_at(start_col, self.diver_row) and self.hold_space_left() <= 0:
            self.message_key = "status_return_ship"
            self.message_args = {}

    def update(self, dt: float) -> None:
        if self.state != "playing":
            return
        self.elapsed += dt
        self.move_cooldown = max(0.0, self.move_cooldown - dt)
        self.score_popup_timer = max(0.0, self.score_popup_timer - dt)
        self._update_time(dt)
        if self.state != "playing":
            return
        self._update_air(dt)
        if self.state != "playing":
            return
        self._update_inactivity_sink(dt)
        self._ensure_sharks()
        self._update_sharks(dt)
        self._update_prey(dt)
        self._update_harpoon(dt)
        self._update_kraken(dt)
        self._ensure_prey(min_count=1)
        self._check_shark_contact()
        self._check_kraken_contact()

    def shark_interval(self) -> float:
        speed_up = self.elapsed * 0.0022 + self.score * 0.00032
        if self.mode == "timeless":
            speed_up += min(0.22, self.elapsed * 0.0018)
        return max(SHARK_MIN_SECONDS, SHARK_BASE_SECONDS - speed_up)

    def _update_time(self, dt: float) -> None:
        if self.mode != "normal":
            return
        self.time_remaining = max(0.0, self.time_remaining - dt)
        if self.time_remaining <= 0:
            self._finish_game("status_time_up")

    def _update_air(self, dt: float) -> None:
        if self.diver_row < 0:
            self.air = min(AIR_MAX, self.air + dt * 82.0)
            return

        depth = self.diver_row / max(1, GRID_ROWS - 1)
        rate = 5.8 + depth * 9.2 + max(0, self.diver_row - 4) * 1.8
        if self.hold_used() >= HOLD_CAPACITY:
            rate += 1.5
        self.air = max(0.0, self.air - rate * dt)
        if self.air <= 22 and self.message_key not in {"status_kraken_warning", "status_kraken_close"}:
            self.message_key = "status_air_low"
            self.message_args = {}
        if self.air <= 0:
            self._lose_life("status_air_out", clear_kraken=False)

    def _update_inactivity_sink(self, dt: float) -> None:
        if self.diver_row < 0 or not self.hold or self.move_penalty_level() <= 0:
            self.inactivity_timer = 0.0
            return
        self.inactivity_timer += dt
        threshold = SINK_TREASURE_SECONDS if "treasure" in self.hold else SINK_BASE_SECONDS
        if self.mode == "timeless":
            threshold = max(1.25, threshold - 0.25)
        if self.inactivity_timer < threshold:
            return
        self.inactivity_timer = 0.0
        if self.diver_row < GRID_ROWS - 1:
            self.diver_row += 1
            self.message_key = "status_heavy"
            self.message_args = {}
            self._check_shark_contact()
            self._check_kraken_contact()

    def _deliver_hold(self) -> None:
        if not self.hold:
            self.message_key = "status_diving"
            self.message_args = {}
            return

        points = sum(PREY_POINTS[item] for item in self.hold)
        full_load = self.hold_used() >= HOLD_CAPACITY or len(self.hold) >= 2
        self.score += points
        if self.score > self.best_score:
            self.best_score = self.score
            self.best_dirty = True
        self.hold.clear()
        self.air = min(AIR_MAX, self.air + 35.0)
        self.inactivity_timer = 0.0
        self.score_popup_value = points
        self.score_popup_timer = 0.85
        self.message_key = "status_big_score" if full_load else "status_scored"
        self.message_args = {"points": points}
        self._ensure_prey(min_count=2)

    def _check_shark_contact(self) -> None:
        if self.state != "playing" or self.diver_row < 0:
            return
        for shark in self.sharks:
            if not is_cell_in_grid(shark.col, shark.row):
                continue
            if shark.col == self.diver_col and shark.row == self.diver_row:
                self._lose_life("status_hit_lost" if self.hold else "status_hit", clear_kraken=False)
                return

    def _check_kraken_contact(self) -> None:
        if self.state != "playing" or self.diver_row < 0 or self.kraken is None:
            return
        if self.kraken.col == self.diver_col and self.kraken.row == self.diver_row:
            self._lose_life("status_kraken_hit", clear_kraken=True)

    def _lose_life(self, message_key: str, clear_kraken: bool) -> None:
        self.life -= 1
        self.hold.clear()
        self.harpoon = None
        if clear_kraken:
            self.kraken = None
            self._schedule_next_kraken()
        self.move_cooldown = 0.0
        self.inactivity_timer = 0.0
        self.air = AIR_MAX
        self.diver_col = GRID_COLUMNS // 2
        self.diver_row = -1
        if self.life <= 0:
            self.life = 0
            self._finish_game("status_game_over")
        else:
            self.message_key = message_key
            self.message_args = {}

    def _finish_game(self, message_key: str) -> None:
        self.state = "game_over"
        self.harpoon = None
        self.kraken = None
        if self.score > self.best_score:
            self.best_score = self.score
            self.best_dirty = True
        self.message_key = message_key
        self.message_args = {}

    def _update_sharks(self, dt: float) -> None:
        interval = self.shark_interval()
        for shark in self.sharks:
            shark.timer -= dt
            while shark.timer <= 0:
                shark.col += shark.direction
                shark.timer += interval
                if shark.direction > 0 and shark.col > GRID_COLUMNS:
                    self._reset_shark(shark)
                    break
                if shark.direction < 0 and shark.col < -1:
                    self._reset_shark(shark)
                    break

    def _update_prey(self, dt: float) -> None:
        for index, item in enumerate(self.prey):
            if item.kind == "treasure":
                continue
            item.timer -= dt
            speed_factor = 1.0 + (min(0.22, self.elapsed * 0.0015) if self.mode == "timeless" else 0.0)
            interval = PREY_MOVE_SECONDS[item.kind] / speed_factor
            while item.timer <= 0:
                item.timer += interval + self.random.uniform(-0.12, 0.18)
                occupied = {
                    (other.col, other.row)
                    for other_index, other in enumerate(self.prey)
                    if other_index != index
                }
                moved = self._try_move_prey(item, item.direction, 0, occupied)
                if not moved:
                    item.direction *= -1
                    self._try_move_prey(item, item.direction, 0, occupied)

                if item.kind == "squid" and self.random.random() < 0.35:
                    vertical = self.random.choice([-1, 1])
                    self._try_move_prey(item, 0, vertical, occupied, min_row=2)

    def _update_harpoon(self, dt: float) -> None:
        if self.harpoon is None:
            return
        self.harpoon.timer -= dt
        if self.harpoon.timer <= 0:
            self.harpoon = None

    def _update_kraken(self, dt: float) -> None:
        if self.kraken is None:
            self.kraken_spawn_timer -= dt
            if self.kraken_spawn_timer <= 0 and self._can_spawn_kraken():
                self._spawn_kraken()
            return

        self.kraken.timer -= dt
        self.kraken.move_timer -= dt
        if self.kraken.timer <= 0:
            self.kraken = None
            self._schedule_next_kraken()
            return

        if self.kraken.move_timer <= 0:
            self.kraken.move_timer += max(0.74, KRAKEN_MOVE_SECONDS - self.difficulty_level() * 0.04)
            target_col = self.diver_col if self.diver_row >= 0 else GRID_COLUMNS // 2
            target_row = self.diver_row if self.diver_row >= 0 else 0
            if self.kraken.col != target_col:
                self.kraken.col += 1 if target_col > self.kraken.col else -1
            elif self.kraken.row != target_row:
                self.kraken.row += 1 if target_row > self.kraken.row else -1
            if self.random.random() < 0.22:
                self.message_key = "status_kraken_close"
                self.message_args = {}
            self._check_kraken_contact()

    def _try_move_prey(
        self,
        item: Prey,
        delta_col: int,
        delta_row: int,
        occupied: set[tuple[int, int]],
        min_row: int = 0,
    ) -> bool:
        next_col = item.col + delta_col
        next_row = item.row + delta_row
        if not (0 <= next_col < GRID_COLUMNS and min_row <= next_row < GRID_ROWS):
            return False
        if (next_col, next_row) in occupied:
            return False
        item.col = next_col
        item.row = next_row
        return True

    def _catch_prey_at(self, col: int, row: int) -> bool:
        for index, item in enumerate(self.prey):
            if item.col != col or item.row != row:
                continue

            self.harpoon = None
            if not self.can_hold(item.kind):
                self.message_key = "status_hold_full"
                self.message_args = {}
                return False

            self.hold.append(item.kind)
            del self.prey[index]
            item_key = f"hold_{item.kind}"
            used = self.hold_used()
            if used >= HOLD_CAPACITY:
                self.message_key = "status_return_ship" if item.kind == "treasure" else "status_big_catch"
                self.message_args = {}
            elif self.move_penalty_level():
                self.message_key = "status_heavy"
                self.message_args = {}
            else:
                self.message_key = "status_caught"
                self.message_args = {"item_key": item_key}
            self._ensure_prey(min_count=1)
            return True
        return False

    def _ensure_sharks(self) -> None:
        target_count = self.target_shark_count()
        while len(self.sharks) < target_count:
            self.sharks.append(self._new_shark())

    def target_shark_count(self) -> int:
        if self.mode == "timeless":
            return min(TIMELESS_MAX_SHARKS, 1 + int(self.elapsed // 42.0) + int(self.score >= 260))
        return 2 if self.elapsed >= 42.0 or self.score >= 180 else 1

    def difficulty_level(self) -> int:
        return max(0, int(self.elapsed // 45.0))

    def _new_shark(self) -> Shark:
        direction = self.random.choice([-1, 1])
        col = -1 if direction > 0 else GRID_COLUMNS
        row = self._pick_shark_row()
        timer = self.random.uniform(0.12, max(0.28, self.shark_interval()))
        return Shark(col=col, row=row, direction=direction, timer=timer)

    def _reset_shark(self, shark: Shark) -> None:
        new_shark = self._new_shark()
        shark.col = new_shark.col
        shark.row = new_shark.row
        shark.direction = new_shark.direction
        shark.timer = new_shark.timer

    def _pick_shark_row(self) -> int:
        bands = [
            list(range(1, 3)),
            list(range(3, 6)),
            list(range(6, GRID_ROWS)),
        ]
        band = self.random.choice(bands)
        recent = set(self.last_shark_rows[-3:])
        choices = [row for row in band if row not in recent]
        if self.diver_row >= 0:
            choices = [row for row in choices if row != self.diver_row] or choices
        if not choices:
            all_rows = [row for row in range(1, GRID_ROWS) if row not in recent]
            choices = all_rows or list(range(1, GRID_ROWS))
        row = self.random.choice(choices)
        self.last_shark_rows.append(row)
        self.last_shark_rows = self.last_shark_rows[-5:]
        return row

    def _ensure_prey(self, min_count: int = 1) -> None:
        target_count = self.random.randint(max(1, min_count), 3)
        while len(self.prey) < target_count:
            kind = self._random_prey_kind()
            cell = self._random_empty_cell(kind)
            if cell is None:
                return
            self.prey.append(
                Prey(
                    kind=kind,
                    col=cell[0],
                    row=cell[1],
                    direction=self.random.choice([-1, 1]),
                    timer=self.random.uniform(0.35, PREY_MOVE_SECONDS.get(kind, 1.0)),
                )
            )

    def _random_empty_cell(self, kind: str) -> tuple[int, int] | None:
        occupied = {(item.col, item.row) for item in self.prey}
        occupied.update((shark.col, shark.row) for shark in self.sharks if is_cell_in_grid(shark.col, shark.row))
        if self.diver_row >= 0:
            occupied.add((self.diver_col, self.diver_row))

        min_row = 2 if kind == "squid" else 0
        if kind == "treasure":
            min_row = TREASURE_MIN_ROW
        cells = [
            (col, row)
            for row in range(min_row, GRID_ROWS)
            for col in range(GRID_COLUMNS)
            if (col, row) not in occupied
        ]
        if not cells:
            return None
        return self.random.choice(cells)

    def _can_spawn_kraken(self) -> bool:
        if self.elapsed < KRAKEN_START_SECONDS:
            self._schedule_next_kraken()
            return False
        if self.mode == "normal" and self.kraken_spawn_count >= KRAKEN_MAX_SPAWNS_NORMAL:
            return False
        chance = KRAKEN_TIMELESS_CHANCE if self.mode == "timeless" else KRAKEN_NORMAL_CHANCE
        chance += min(0.09, self.difficulty_level() * 0.012)
        if self.random.random() > chance:
            self._schedule_next_kraken()
            return False
        return True

    def _spawn_kraken(self) -> None:
        side = self.random.choice([-1, 1])
        col = 0 if side > 0 else GRID_COLUMNS - 1
        if self.diver_row >= 0:
            row = clamp(self.diver_row + self.random.choice([-1, 0, 1]), 1, GRID_ROWS - 1)
        else:
            row = self.random.randrange(3, GRID_ROWS)
        self.kraken = Kraken(col=col, row=row)
        self.kraken_spawn_count += 1
        self.message_key = "status_kraken_warning"
        self.message_args = {}

    def _schedule_next_kraken(self) -> None:
        if self.mode == "timeless":
            base = max(13.0, 28.0 - self.difficulty_level() * 2.0)
            self.kraken_spawn_timer = base + self.random.uniform(6.0, 18.0)
        else:
            self.kraken_spawn_timer = 9999.0

    def _random_prey_kind(self) -> str:
        roll = self.random.random()
        if roll < 0.48:
            return "small_fish"
        if roll < 0.74:
            return "big_fish"
        if roll < 0.90:
            return "squid"
        return "treasure"

    def _generate_background(self) -> None:
        self.background_items.clear()
        bottom_rows = [GRID_ROWS - 2, GRID_ROWS - 1]
        edge_cols = [0, 1, GRID_COLUMNS - 2, GRID_COLUMNS - 1]
        for col in edge_cols:
            kind = self.random.choice(["rock", "coral", "seaweed"])
            row = self.random.choice(bottom_rows)
            self.background_items.append(BackgroundItem(kind=kind, col=col, row=row, variant=self.random.randint(0, 2)))

        for _index in range(4):
            self.background_items.append(
                BackgroundItem(
                    kind="bubble",
                    col=self.random.randrange(GRID_COLUMNS),
                    row=self.random.randrange(1, GRID_ROWS - 1),
                    variant=self.random.randint(0, 2),
                )
            )


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = get_app_dir()
CONFIG_PATH = APP_DIR / CONFIG_NAME


def clamp(value: int, lower: int, upper: int) -> int:
    return max(lower, min(upper, value))


def is_cell_in_grid(col: int, row: int) -> bool:
    return 0 <= col < GRID_COLUMNS and 0 <= row < GRID_ROWS


def key_token_from_keysym(keysym: str) -> str:
    token = keysym.lower()
    aliases = {
        "return": "return",
        "space": "space",
        "up": "up",
        "down": "down",
        "left": "left",
        "right": "right",
    }
    return aliases.get(token, token)


def load_best_score(path: Path = CONFIG_PATH) -> int:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return max(0, int(data.get("best_score", 0)))
    except Exception:
        return 0


def save_best_score(best_score: int, path: Path = CONFIG_PATH) -> None:
    text = json.dumps({"best_score": int(best_score)}, ensure_ascii=False, indent=2) + "\n"
    path.write_text(text, encoding="utf-8")


def emit_stream_line(text: str, error: bool = False) -> None:
    line = text + "\n"
    handle_id = -12 if error else -11
    if os.name == "nt" and getattr(sys, "frozen", False):
        if write_windows_stream(line, handle_id):
            return

    stream = sys.stderr if error else sys.stdout
    if stream is not None:
        try:
            stream.write(line)
            stream.flush()
            return
        except Exception:
            pass

    if os.name == "nt":
        write_windows_stream(line, handle_id)


def write_windows_stream(text: str, handle_id: int) -> bool:
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        kernel32.GetStdHandle.restype = ctypes.c_void_p
        handle = kernel32.GetStdHandle(ctypes.c_int(handle_id))
        invalid_handle = ctypes.c_void_p(-1).value
        if handle in (None, 0, invalid_handle):
            kernel32.AttachConsole(ctypes.c_ulong(0xFFFFFFFF))
            handle = kernel32.GetStdHandle(ctypes.c_int(handle_id))
        if handle in (None, 0, invalid_handle):
            return False

        data = text.encode("utf-8")
        buffer = ctypes.create_string_buffer(data)
        written = ctypes.c_ulong(0)
        kernel32.WriteFile.argtypes = [
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_ulong,
            ctypes.POINTER(ctypes.c_ulong),
            ctypes.c_void_p,
        ]
        kernel32.WriteFile.restype = ctypes.c_int
        return bool(kernel32.WriteFile(handle, buffer, len(data), ctypes.byref(written), None))
    except Exception:
        return False


class DiverCatchApp(tk.Tk):
    def __init__(self, config_path: Path = CONFIG_PATH) -> None:
        super().__init__()
        self.config_path = config_path
        self.model = GameModel(best_score=load_best_score(config_path))
        self.title(WINDOW_TITLE)
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.resizable(False, False)
        self.configure(bg=COLORS["lcd_bg"])
        self._setup_icon()

        self.canvas = tk.Canvas(
            self,
            width=WINDOW_WIDTH,
            height=WINDOW_HEIGHT,
            bg=COLORS["lcd_bg"],
            highlightthickness=0,
            bd=0,
        )
        self.canvas.pack(fill="both", expand=True)
        self.last_time = time.perf_counter()
        self.pressed_keys: set[str] = set()
        self._bind_controls()
        self._draw()
        self.after(33, self._tick)

    def _setup_icon(self) -> None:
        candidates = [
            APP_DIR / "dake_icon.ico",
            APP_DIR.parent.parent / "02_assets" / "dake_icon.ico",
            APP_DIR.parent.parent.parent / "02_assets" / "dake_icon.ico",
        ]
        for icon_path in candidates:
            if icon_path.exists():
                try:
                    self.iconbitmap(str(icon_path))
                except Exception:
                    pass
                return

    def _bind_controls(self) -> None:
        self.bind("<KeyPress>", self._handle_key_press)
        self.bind("<KeyRelease>", self._handle_key_release)
        self.bind("<FocusOut>", lambda _event: self.pressed_keys.clear())
        self.focus_force()

    def _handle_key_press(self, event: tk.Event) -> None:
        token = key_token_from_keysym(str(event.keysym))
        if token in self.pressed_keys:
            return
        self.pressed_keys.add(token)

        if token == "return":
            self._handle_enter()
        elif token == "r":
            self._restart()
        elif token == "t":
            self._toggle_mode()
        elif token == "p":
            self._toggle_pause()
        elif token == "space":
            self._fire()
        elif token in {"up", "w"}:
            self._move(0, -1)
        elif token in {"down", "s"}:
            self._move(0, 1)
        elif token in {"left", "a"}:
            self._move(-1, 0)
        elif token in {"right", "d"}:
            self._move(1, 0)

    def _handle_key_release(self, event: tk.Event) -> None:
        self.pressed_keys.discard(key_token_from_keysym(str(event.keysym)))

    def _handle_enter(self) -> None:
        if self.model.state in {"idle", "game_over"}:
            self.model.start_game()
        elif self.model.state == "paused":
            self.model.resume()
        self._draw()

    def _restart(self) -> None:
        self.model.start_game()
        self._save_best_if_needed()
        self._draw()

    def _toggle_pause(self) -> None:
        self.model.toggle_pause()
        self._draw()

    def _toggle_mode(self) -> None:
        self.model.toggle_mode()
        self._draw()

    def _move(self, delta_col: int, delta_row: int) -> None:
        self.model.move(delta_col, delta_row)
        self._save_best_if_needed()
        self._draw()

    def _fire(self) -> None:
        self.model.fire_harpoon()
        self._draw()

    def _tick(self) -> None:
        now = time.perf_counter()
        dt = min(0.12, now - self.last_time)
        self.last_time = now
        self.model.update(dt)
        self._save_best_if_needed()
        self._draw()
        self.after(33, self._tick)

    def _save_best_if_needed(self) -> None:
        if not self.model.best_dirty:
            return
        try:
            save_best_score(self.model.best_score, self.config_path)
            self.model.best_dirty = False
        except Exception:
            self.model.message_key = "save_failed"
            self.model.message_args = {}

    def _draw(self) -> None:
        self.canvas.delete("all")
        self._draw_lcd_background()
        self._draw_status_row()
        self._draw_ship_area()
        self._draw_grid()
        self._draw_prey()
        self._draw_harpoon()
        self._draw_diver()
        self._draw_sharks()
        self._draw_kraken()
        self._draw_score_popup()
        self._draw_footer()
        if self.model.state == "idle":
            self._draw_center_overlay(UI_TEXT["display_name"], UI_TEXT["ready_hint"])
        elif self.model.state == "paused":
            self._draw_center_overlay(UI_TEXT["paused_title"], UI_TEXT["paused_hint"])
        elif self.model.state == "game_over":
            title = UI_TEXT["time_up_title"] if self.model.message_key == "status_time_up" else UI_TEXT["game_over_title"]
            self._draw_center_overlay(title, UI_TEXT["game_over_hint"])

    def _draw_lcd_background(self) -> None:
        self.canvas.create_rectangle(0, 0, WINDOW_WIDTH, WINDOW_HEIGHT, fill=COLORS["lcd_bg"], outline="")
        for y in range(0, WINDOW_HEIGHT, 12):
            self.canvas.create_line(0, y, WINDOW_WIDTH, y, fill=COLORS["lcd_line"], width=1)

    def _draw_status_row(self) -> None:
        self.canvas.create_rectangle(10, 8, WINDOW_WIDTH - 10, 58, fill=COLORS["panel"], outline=COLORS["ink"], width=2)
        stats = [
            (UI_TEXT["score_label"], str(self.model.score)),
            (UI_TEXT["best_label"], str(self.model.best_score)),
            (UI_TEXT["life_label"], str(self.model.life)),
            (UI_TEXT["air_label"], str(int(round(self.model.air)))),
            (UI_TEXT["time_label"], self._time_text()),
            (UI_TEXT["hold_label"], self._hold_text()),
        ]
        column_width = (WINDOW_WIDTH - 20) / len(stats)
        for index, (label, value) in enumerate(stats):
            x = 20 + column_width * index
            self.canvas.create_text(x, 21, text=label, fill=COLORS["ink"], font=FONTS["lcd"], anchor="w")
            self.canvas.create_text(x, 44, text=value, fill=COLORS["ink"], font=FONTS["lcd_large"], anchor="w")
            if index:
                self.canvas.create_line(x - 7, 17, x - 7, 51, fill=COLORS["lcd_shadow"], width=1)

    def _draw_ship_area(self) -> None:
        sea_width = GRID_COLUMNS * CELL_WIDTH
        self.canvas.create_rectangle(GRID_LEFT, SHIP_Y, GRID_LEFT + sea_width, SEA_TOP, fill=COLORS["water"], outline=COLORS["ink"], width=2)
        self.canvas.create_line(GRID_LEFT, SEA_TOP, GRID_LEFT + sea_width, SEA_TOP, fill=COLORS["ink"], width=3)
        for col in range(GRID_COLUMNS):
            x = GRID_LEFT + col * CELL_WIDTH + CELL_WIDTH / 2
            self.canvas.create_arc(x - 16, SEA_TOP - 10, x + 16, SEA_TOP + 12, start=0, extent=180, outline=COLORS["lcd_shadow"], width=1)

        boat_x = self._cell_center(self.model.diver_col, 0)[0] if self.model.diver_row < 0 else WINDOW_WIDTH / 2
        boat_y = SHIP_Y + 28
        self.canvas.create_polygon(
            boat_x - 58,
            boat_y + 10,
            boat_x + 58,
            boat_y + 10,
            boat_x + 42,
            boat_y + 28,
            boat_x - 42,
            boat_y + 28,
            fill=COLORS["ink"],
            outline=COLORS["ink"],
        )
        self.canvas.create_rectangle(boat_x - 32, boat_y - 9, boat_x + 24, boat_y + 10, fill=COLORS["highlight"], outline=COLORS["ink"], width=2)
        self.canvas.create_line(boat_x + 38, boat_y + 10, boat_x + 38, SEA_TOP, fill=COLORS["ink"], width=2)

    def _draw_grid(self) -> None:
        x1 = GRID_LEFT
        y1 = GRID_TOP
        x2 = GRID_LEFT + GRID_COLUMNS * CELL_WIDTH
        y2 = GRID_TOP + GRID_ROWS * CELL_HEIGHT
        self.canvas.create_rectangle(x1, y1, x2, y2, fill=COLORS["water"], outline=COLORS["ink"], width=2)

        for row in range(GRID_ROWS):
            y = GRID_TOP + row * CELL_HEIGHT
            self.canvas.create_rectangle(x1 + 2, y + 2, x2 - 2, y + CELL_HEIGHT - 2, fill="", outline=COLORS["lcd_detail"] if row >= 4 else COLORS["lcd_line"], width=1)
        for row in range(2, GRID_ROWS):
            y = GRID_TOP + row * CELL_HEIGHT + CELL_HEIGHT * 0.55
            self.canvas.create_line(x1 + 8, y, x2 - 8, y, fill=COLORS["lcd_detail"], width=1)

        self._draw_background_items()
        self.canvas.create_line(x1, y2 - 4, x2, y2 - 4, fill=COLORS["lcd_shadow"], width=2)

        for col in range(1, GRID_COLUMNS):
            x = GRID_LEFT + col * CELL_WIDTH
            self.canvas.create_line(x, y1, x, y2, fill=COLORS["lcd_shadow"], width=1)
        for row in range(1, GRID_ROWS):
            y = GRID_TOP + row * CELL_HEIGHT
            self.canvas.create_line(x1, y, x2, y, fill=COLORS["lcd_shadow"], width=1)

    def _draw_background_items(self) -> None:
        for item in self.model.background_items:
            x, y = self._cell_center(item.col, item.row)
            if item.kind == "rock":
                self.canvas.create_arc(x - 20, y + 6, x + 20, y + 32, start=0, extent=180, outline=COLORS["lcd_shadow"], width=2)
                self.canvas.create_arc(x - 10, y + 11, x + 26, y + 31, start=0, extent=180, outline=COLORS["lcd_detail"], width=1)
            elif item.kind == "coral":
                base_y = y + 18
                self.canvas.create_line(x, base_y, x, base_y - 22, fill=COLORS["lcd_shadow"], width=2)
                self.canvas.create_line(x, base_y - 11, x - 12, base_y - 20, fill=COLORS["lcd_shadow"], width=2)
                self.canvas.create_line(x, base_y - 14, x + 12, base_y - 25, fill=COLORS["lcd_shadow"], width=2)
            elif item.kind == "seaweed":
                for offset in (-9, 0, 9):
                    self.canvas.create_line(x + offset, y + 22, x + offset + 5, y - 7, fill=COLORS["lcd_shadow"], width=2)
            elif item.kind == "bubble":
                size = 3 + item.variant
                self.canvas.create_oval(x - size, y - size, x + size, y + size, outline=COLORS["lcd_detail"], width=1)

    def _draw_prey(self) -> None:
        for item in self.model.prey:
            x, y = self._cell_center(item.col, item.row)
            if item.kind == "small_fish":
                self._draw_fish(x, y, 18, 8, item.direction)
            elif item.kind == "big_fish":
                self._draw_fish(x, y, 28, 12, item.direction)
            elif item.kind == "squid":
                self._draw_squid(x, y)
            else:
                self._draw_treasure(x, y)

    def _draw_sharks(self) -> None:
        for shark in self.model.sharks:
            x, y = self._cell_center(shark.col, shark.row)
            direction = 1 if shark.direction >= 0 else -1
            self.canvas.create_polygon(
                x - 29 * direction,
                y,
                x - 9 * direction,
                y - 15,
                x + 22 * direction,
                y - 8,
                x + 31 * direction,
                y,
                x + 22 * direction,
                y + 8,
                x - 9 * direction,
                y + 15,
                fill=COLORS["ink"],
                outline=COLORS["ink"],
            )
            self.canvas.create_polygon(
                x - 3 * direction,
                y - 13,
                x + 8 * direction,
                y - 28,
                x + 13 * direction,
                y - 12,
                fill=COLORS["ink"],
                outline=COLORS["ink"],
            )
            self.canvas.create_oval(x + 15 * direction - 2, y - 5, x + 15 * direction + 2, y - 1, fill=COLORS["highlight"], outline="")

    def _draw_kraken(self) -> None:
        kraken = self.model.kraken
        if kraken is None:
            return
        x, y = self._cell_center(kraken.col, kraken.row)
        self.canvas.create_oval(x - 22, y - 23, x + 22, y + 14, fill=COLORS["ink"], outline=COLORS["ink"])
        self.canvas.create_arc(x - 26, y - 31, x + 26, y + 21, start=20, extent=140, outline=COLORS["highlight"], width=2)
        for offset in (-18, -9, 0, 9, 18):
            self.canvas.create_line(x + offset, y + 10, x + offset - 8, y + 31, fill=COLORS["ink"], width=3)
        self.canvas.create_oval(x - 9, y - 7, x - 4, y - 2, fill=COLORS["highlight"], outline="")
        self.canvas.create_oval(x + 4, y - 7, x + 9, y - 2, fill=COLORS["highlight"], outline="")

    def _draw_diver(self) -> None:
        if self.model.diver_row < 0:
            x = self._cell_center(self.model.diver_col, 0)[0]
            y = SHIP_Y + 22
        else:
            x, y = self._cell_center(self.model.diver_col, self.model.diver_row)
        direction = 1 if self.model.facing >= 0 else -1
        harpoon_pose = self.model.harpoon is not None

        tank_x = x - 13 * direction
        self.canvas.create_rectangle(tank_x - 5, y - 16, tank_x + 5, y + 8, fill=COLORS["ink"], outline=COLORS["ink"])
        self.canvas.create_line(tank_x, y - 18, tank_x, y - 22, fill=COLORS["ink"], width=2)

        self.canvas.create_polygon(
            x - 10 * direction,
            y - 5,
            x + 12 * direction,
            y - 10,
            x + 18 * direction,
            y + 7,
            x - 5 * direction,
            y + 12,
            fill=COLORS["ink"],
            outline=COLORS["ink"],
        )
        head_x1, head_x2 = sorted((x + 8 * direction, x + 24 * direction))
        mask_x1, mask_x2 = sorted((x + 13 * direction, x + 25 * direction))
        self.canvas.create_oval(head_x1, y - 23, head_x2, y - 7, fill=COLORS["highlight"], outline=COLORS["ink"], width=2)
        self.canvas.create_rectangle(mask_x1, y - 18, mask_x2, y - 11, fill=COLORS["ink"], outline=COLORS["ink"])
        self.canvas.create_line(x + 18 * direction, y - 22, x + 18 * direction, y - 31, fill=COLORS["ink"], width=2)
        self.canvas.create_line(x + 18 * direction, y - 31, x + 27 * direction, y - 31, fill=COLORS["ink"], width=2)
        self.canvas.create_line(x + 11 * direction, y - 9, tank_x, y - 4, fill=COLORS["ink"], width=2)

        arm_y = y + (0 if harpoon_pose else 4)
        hand_x = x + (31 if harpoon_pose else 23) * direction
        self.canvas.create_line(x + 11 * direction, y - 2, hand_x, arm_y, fill=COLORS["ink"], width=3)
        self.canvas.create_line(x + 7 * direction, y + 5, x + 20 * direction, y + 10, fill=COLORS["ink"], width=3)
        self.canvas.create_line(hand_x - 7 * direction, arm_y - 5, hand_x + 15 * direction, arm_y - 7, fill=COLORS["ink"], width=2)

        self._draw_fin(x - 6 * direction, y + 12, -1, direction)
        self._draw_fin(x - 15 * direction, y + 15, 1, direction)

        if self.model.hold:
            bag_x1, bag_x2 = sorted((x - 23 * direction, x - 10 * direction))
            self.canvas.create_rectangle(bag_x1, y - 2, bag_x2, y + 11, fill=COLORS["highlight"], outline=COLORS["ink"], width=2)

    def _draw_fin(self, x: float, y: float, lift: int, direction: int) -> None:
        self.canvas.create_line(x, y, x - 13 * direction, y + 12 + lift * 3, fill=COLORS["ink"], width=3)
        self.canvas.create_polygon(
            x - 14 * direction,
            y + 11 + lift * 3,
            x - 29 * direction,
            y + 7 + lift * 3,
            x - 21 * direction,
            y + 18 + lift * 3,
            fill=COLORS["ink"],
            outline=COLORS["ink"],
        )

    def _draw_harpoon(self) -> None:
        harpoon = self.model.harpoon
        if harpoon is None:
            return
        x, y = self._cell_center(harpoon.col, harpoon.row)
        direction = 1 if harpoon.direction >= 0 else -1
        self.canvas.create_line(x - 20 * direction, y, x + 20 * direction, y, fill=COLORS["ink"], width=3)
        self.canvas.create_polygon(
            x + 23 * direction,
            y,
            x + 12 * direction,
            y - 6,
            x + 12 * direction,
            y + 6,
            fill=COLORS["ink"],
            outline=COLORS["ink"],
        )

    def _draw_fish(self, x: float, y: float, width: float, height: float, direction: int) -> None:
        direction = 1 if direction >= 0 else -1
        self.canvas.create_oval(x - width / 2, y - height, x + width / 2, y + height, fill=COLORS["ink"], outline=COLORS["ink"])
        tail_x = x - width / 2 * direction
        self.canvas.create_polygon(
            tail_x,
            y,
            tail_x - 10 * direction,
            y - 8,
            tail_x - 10 * direction,
            y + 8,
            fill=COLORS["ink"],
            outline=COLORS["ink"],
        )
        eye_x = x + width / 4 * direction
        self.canvas.create_oval(eye_x - 2, y - 4, eye_x + 2, y, fill=COLORS["highlight"], outline="")

    def _draw_squid(self, x: float, y: float) -> None:
        self.canvas.create_oval(x - 12, y - 17, x + 12, y + 8, fill=COLORS["ink"], outline=COLORS["ink"])
        self.canvas.create_polygon(x - 12, y - 7, x, y - 24, x + 12, y - 7, fill=COLORS["ink"], outline=COLORS["ink"])
        for offset in (-9, -3, 3, 9):
            self.canvas.create_line(x + offset, y + 7, x + offset - 4, y + 19, fill=COLORS["ink"], width=2)
        self.canvas.create_oval(x - 5, y - 5, x - 2, y - 2, fill=COLORS["highlight"], outline="")
        self.canvas.create_oval(x + 2, y - 5, x + 5, y - 2, fill=COLORS["highlight"], outline="")

    def _draw_treasure(self, x: float, y: float) -> None:
        self.canvas.create_rectangle(x - 18, y - 10, x + 18, y + 13, fill=COLORS["ink"], outline=COLORS["ink"])
        self.canvas.create_arc(x - 18, y - 20, x + 18, y + 8, start=0, extent=180, fill=COLORS["ink"], outline=COLORS["ink"])
        self.canvas.create_rectangle(x - 4, y - 2, x + 4, y + 9, fill=COLORS["highlight"], outline=COLORS["highlight"])

    def _draw_score_popup(self) -> None:
        if self.model.score_popup_timer <= 0 or self.model.score_popup_value <= 0:
            return
        y = SHIP_Y + 16 - (0.85 - self.model.score_popup_timer) * 18
        self.canvas.create_text(WINDOW_WIDTH / 2, y, text=f"+{self.model.score_popup_value}", fill=COLORS["ink"], font=FONTS["lcd_large"])

    def _draw_footer(self) -> None:
        self.canvas.create_rectangle(10, FOOTER_TOP, WINDOW_WIDTH - 10, FOOTER_BOTTOM, fill=COLORS["panel"], outline=COLORS["ink"], width=2)
        self.canvas.create_text(22, FOOTER_TOP + 24, text=self._status_text(), fill=COLORS["ink"], font=FONTS["label"], anchor="w")
        self.canvas.create_text(22, FOOTER_TOP + 52, text=self._mode_text(), fill=COLORS["ink"], font=FONTS["lcd"], anchor="w")
        self.canvas.create_text(22, FOOTER_TOP + 78, text=UI_TEXT["playing_hint"], fill=COLORS["ink_soft"], font=FONTS["small"], anchor="w")
        self.canvas.create_text(WINDOW_WIDTH - 22, FOOTER_TOP + 78, text=COPYRIGHT, fill=COLORS["ink_soft"], font=FONTS["small"], anchor="e")

    def _draw_center_overlay(self, title: str, hint: str) -> None:
        self.canvas.create_rectangle(64, 260, WINDOW_WIDTH - 64, 354, fill=COLORS["lcd_bg"], outline=COLORS["ink"], width=2)
        self.canvas.create_text(WINDOW_WIDTH / 2, 294, text=title, fill=COLORS["ink"], font=FONTS["title"])
        self.canvas.create_text(WINDOW_WIDTH / 2, 330, text=hint, fill=COLORS["ink"], font=FONTS["lcd"])

    def _cell_center(self, col: int, row: int) -> tuple[float, float]:
        return (
            GRID_LEFT + col * CELL_WIDTH + CELL_WIDTH / 2,
            GRID_TOP + row * CELL_HEIGHT + CELL_HEIGHT / 2,
        )

    def _hold_text(self) -> str:
        return f"{self.model.hold_used()}/{HOLD_CAPACITY}"

    def _time_text(self) -> str:
        if self.model.mode == "timeless":
            return "--"
        return str(max(0, int(self.model.time_remaining + 0.99)))

    def _mode_text(self) -> str:
        mode_text = UI_TEXT["mode_timeless"] if self.model.mode == "timeless" else UI_TEXT["mode_normal"]
        return f"{UI_TEXT['mode_label']} {mode_text}  T:MODE"

    def _status_text(self) -> str:
        args = dict(self.model.message_args)
        item_key = args.pop("item_key", None)
        if item_key is not None:
            args["item"] = UI_TEXT[str(item_key)]
        return UI_TEXT[self.model.message_key].format(**args)


def run_launch_check() -> None:
    with tempfile.TemporaryDirectory() as temp_dir:
        config_path = Path(temp_dir) / CONFIG_NAME
        save_best_score(50, config_path)
        assert load_best_score(config_path) == 50

        model = GameModel(best_score=load_best_score(config_path), random=random.Random(7))
        assert model.state == "idle"
        assert model.mode == "normal"
        assert model.air == AIR_MAX
        assert model.time_remaining == NORMAL_TIME_LIMIT
        assert model.kraken is None
        assert model.inactivity_timer == 0.0
        assert model.kraken_spawn_count == 0
        assert GRID_COLUMNS == 7
        assert GRID_ROWS == 8
        assert HOLD_CAPACITY == 3
        assert HARPOON_DISPLAY_SECONDS <= 0.2
        assert key_token_from_keysym("space") == "space"
        assert key_token_from_keysym("W") == "w"

        model.toggle_mode()
        assert model.mode == "timeless"
        assert model.target_shark_count() == 1
        model.toggle_mode()
        assert model.mode == "normal"

        model.start_game()
        assert model.state == "playing"
        assert model.life == MAX_LIFE
        assert model.air == AIR_MAX
        assert 59.0 < model.time_remaining <= NORMAL_TIME_LIMIT
        assert model.hold_used() == 0
        assert len(model.background_items) >= 4
        assert 1 <= len(model.sharks) <= 2
        assert 1 <= len(model.prey) <= 3

        for _index in range(20):
            cell = model._random_empty_cell("treasure")
            assert cell is None or cell[1] >= TREASURE_MIN_ROW

        model.prey = [Prey("small_fish", 5, 0)]
        model.sharks = []
        model.diver_col = 3
        model.diver_row = 0
        model.facing = 1
        model.fire_harpoon()
        assert model.hold == []
        assert model.harpoon is not None
        assert model.harpoon.col == 4
        model._update_harpoon(1.0)
        assert model.harpoon is None

        model.prey = [Prey("small_fish", 4, 0)]
        model.fire_harpoon()
        assert model.hold == ["small_fish"]
        assert model.hold_used() == 1

        model._update_harpoon(1.0)
        model.prey = [Prey("small_fish", 4, 0)]
        model.fire_harpoon()
        model._update_harpoon(1.0)
        model.prey = [Prey("small_fish", 4, 0)]
        model.fire_harpoon()
        assert model.hold_used() == 3
        assert model.move_penalty_level() == 1

        model.prey = [Prey("treasure", 4, 0)]
        model.fire_harpoon()
        assert model.hold == ["small_fish", "small_fish", "small_fish"]
        assert model.message_key in {"status_return_ship", "status_hold_full"}

        model.move_cooldown = 0.0
        model.diver_row = 0
        model.move(0, -1)
        assert model.score == 30
        assert model.hold_used() == 0
        assert model.score_popup_value == 30

        model.diver_row = 0
        model.air = AIR_MAX
        model._update_air(1.0)
        shallow_air = model.air
        model.air = AIR_MAX
        model.diver_row = GRID_ROWS - 1
        model._update_air(1.0)
        deep_air = model.air
        assert deep_air < shallow_air

        model.air = 0.1
        model.life = 2
        model.hold = ["small_fish"]
        model.diver_row = GRID_ROWS - 1
        model._update_air(1.0)
        assert model.life == 1
        assert model.diver_row == -1
        assert model.hold_used() == 0
        assert model.air == AIR_MAX

        model.hold = ["treasure"]
        treasure_delay = model.move_delay()
        model.hold = ["big_fish"]
        big_fish_delay = model.move_delay()
        assert treasure_delay > big_fish_delay > 0

        model.hold = ["treasure"]
        model.diver_row = 2
        model.diver_col = 3
        model.inactivity_timer = SINK_TREASURE_SECONDS
        model._update_inactivity_sink(0.01)
        assert model.diver_row == 3

        model.hold.clear()
        model.prey = [Prey("squid", 4, 3, direction=-1, timer=0.0)]
        model.diver_col = 3
        model.diver_row = 3
        model.facing = 1
        model.fire_harpoon()
        assert model.hold == ["squid"]
        assert model.hold_used() == 2
        assert PREY_POINTS["squid"] == 50

        moving_prey = Prey("small_fish", 1, 2, direction=1, timer=0.0)
        model.prey = [moving_prey]
        model._update_prey(1.4)
        assert moving_prey.col != 1

        model.last_shark_rows = [2, 2, 2]
        row = model._pick_shark_row()
        assert 1 <= row < GRID_ROWS

        model.kraken = Kraken(col=0, row=6, timer=KRAKEN_DURATION_SECONDS, move_timer=0.0)
        model.diver_col = 3
        model.diver_row = 6
        model._update_kraken(0.01)
        assert model.kraken is not None
        assert model.kraken.col == 1

        model.kraken = None
        model.kraken_spawn_timer = 0.0
        model.elapsed = KRAKEN_START_SECONDS + 1.0
        model.random = random.Random(1)
        model._update_kraken(0.01)
        assert model.kraken is not None or model.kraken_spawn_timer > 0

        model.mode = "normal"
        model.state = "playing"
        model.time_remaining = 0.01
        model.update(0.05)
        assert model.state == "game_over"
        assert model.message_key == "status_time_up"

        model = GameModel(best_score=load_best_score(config_path), random=random.Random(9))
        model.toggle_mode()
        assert model.mode == "timeless"
        model.start_game()
        model.elapsed = 130.0
        assert model.target_shark_count() >= 3
        model.update(0.05)
        assert model.state == "playing"

        model.best_dirty = False
        model.state = "playing"
        model.life = 1
        model.hold = ["squid"]
        model.diver_col = 1
        model.diver_row = 2
        model.sharks = [Shark(col=1, row=2, direction=1, timer=1.0)]
        model.update(0.01)
        assert model.state == "game_over"
        assert model.life == 0
        assert model.hold_used() == 0

        save_best_score(max(model.best_score, 120), config_path)
        assert load_best_score(config_path) >= 120

    emit_stream_line(UI_TEXT["launch_check_ok"])


def main() -> None:
    if "--launch-check" in sys.argv:
        try:
            run_launch_check()
        except Exception as exc:
            emit_stream_line(f"LAUNCH CHECK FAILED: {exc}", error=True)
            sys.exit(1)
        return

    app = DiverCatchApp()
    app.mainloop()


if __name__ == "__main__":
    main()
