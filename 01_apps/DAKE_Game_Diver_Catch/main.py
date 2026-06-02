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
    "hold_label": "HOLD",
    "hold_empty": "--",
    "hold_small_fish": "小魚",
    "hold_big_fish": "大魚",
    "hold_treasure": "宝箱",
    "ready_hint": "ENTER START",
    "playing_hint": "銛で獲って船へ戻る",
    "paused_title": "PAUSED",
    "paused_hint": "ENTER RESUME",
    "game_over_title": "GAME OVER",
    "game_over_hint": "ENTER / R RESTART",
    "status_ready": "ENTERで開始",
    "status_diving": "潜水開始",
    "status_need_water": "海中で銛を撃つ",
    "status_harpoon": "銛発射",
    "status_harpoon_busy": "銛を回収中",
    "status_hold_full": "獲物を船へ戻す",
    "status_missed": "銛は外れた",
    "status_caught": "{item}を確保",
    "status_scored": "{points}点加算",
    "status_hit": "サメに接触",
    "status_hit_lost": "サメに接触、獲物を失った",
    "status_game_over": "GAME OVER",
    "status_paused": "一時停止",
    "status_resumed": "再開",
    "save_failed": "BEST保存に失敗",
    "launch_check_ok": "LAUNCH CHECK OK",
}

CONFIG_NAME = "diver_catch_config.json"
WINDOW_WIDTH = 420
WINDOW_HEIGHT = 540
GRID_COLUMNS = 5
GRID_ROWS = 6
GRID_LEFT = 40
GRID_TOP = 138
CELL_WIDTH = 68
CELL_HEIGHT = 50
SHIP_Y = 70
SEA_TOP = GRID_TOP - 10
SEA_BOTTOM = GRID_TOP + GRID_ROWS * CELL_HEIGHT
MAX_LIFE = 3
HARPOON_STEP_SECONDS = 0.085
SHARK_BASE_SECONDS = 0.82
SHARK_MIN_SECONDS = 0.28

PREY_POINTS = {
    "small_fish": 10,
    "big_fish": 30,
    "treasure": 100,
}

COLORS = {
    "lcd_bg": "#c4d6a3",
    "lcd_line": "#a9bd86",
    "lcd_shadow": "#91a873",
    "ink": "#1f2c1c",
    "ink_soft": "#31422b",
    "panel": "#b4c98f",
    "water": "#b8cc95",
    "button": "#9db77a",
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
    timer: float = HARPOON_STEP_SECONDS
    steps_left: int = GRID_COLUMNS


@dataclass
class GameModel:
    best_score: int = 0
    random: random.Random = field(default_factory=random.Random)
    state: str = "idle"
    score: int = 0
    life: int = MAX_LIFE
    hold: str | None = None
    diver_col: int = GRID_COLUMNS // 2
    diver_row: int = -1
    facing: int = 1
    elapsed: float = 0.0
    prey: list[Prey] = field(default_factory=list)
    sharks: list[Shark] = field(default_factory=list)
    harpoon: Harpoon | None = None
    message_key: str = "status_ready"
    message_args: dict[str, object] = field(default_factory=dict)
    best_dirty: bool = False

    def start_game(self) -> None:
        self.state = "playing"
        self.score = 0
        self.life = MAX_LIFE
        self.hold = None
        self.diver_col = GRID_COLUMNS // 2
        self.diver_row = -1
        self.facing = 1
        self.elapsed = 0.0
        self.prey.clear()
        self.sharks.clear()
        self.harpoon = None
        self.message_key = "status_diving"
        self.message_args = {}
        self._ensure_sharks()
        self._ensure_prey(min_count=2)

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

    def move(self, delta_col: int, delta_row: int) -> None:
        if self.state != "playing":
            return

        if delta_col:
            self.facing = 1 if delta_col > 0 else -1

        if self.diver_row < 0:
            self.diver_col = clamp(self.diver_col + delta_col, 0, GRID_COLUMNS - 1)
            if delta_row > 0:
                self.diver_row = 0
                self.message_key = "status_diving"
                self.message_args = {}
            return

        self.diver_col = clamp(self.diver_col + delta_col, 0, GRID_COLUMNS - 1)
        next_row = self.diver_row + delta_row
        if next_row < 0:
            self.diver_row = -1
            self._deliver_hold()
            return
        self.diver_row = clamp(next_row, 0, GRID_ROWS - 1)
        self._check_shark_contact()

    def fire_harpoon(self) -> None:
        if self.state != "playing":
            return
        if self.diver_row < 0:
            self.message_key = "status_need_water"
            self.message_args = {}
            return
        if self.hold is not None:
            self.message_key = "status_hold_full"
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
        self.message_key = "status_harpoon"
        self.message_args = {}
        self._catch_prey_at(start_col, self.diver_row)

    def update(self, dt: float) -> None:
        if self.state != "playing":
            return
        self.elapsed += dt
        self._ensure_sharks()
        self._update_sharks(dt)
        self._update_harpoon(dt)
        self._ensure_prey(min_count=1)
        self._check_shark_contact()

    def shark_interval(self) -> float:
        speed_up = self.elapsed * 0.006 + self.score * 0.0013
        return max(SHARK_MIN_SECONDS, SHARK_BASE_SECONDS - speed_up)

    def _deliver_hold(self) -> None:
        if self.hold is None:
            self.message_key = "status_diving"
            self.message_args = {}
            return

        points = PREY_POINTS[self.hold]
        self.score += points
        if self.score > self.best_score:
            self.best_score = self.score
            self.best_dirty = True
        self.hold = None
        self.message_key = "status_scored"
        self.message_args = {"points": points}
        self._ensure_prey(min_count=2)

    def _check_shark_contact(self) -> None:
        if self.state != "playing" or self.diver_row < 0:
            return
        for shark in self.sharks:
            if shark.col == self.diver_col and shark.row == self.diver_row:
                had_hold = self.hold is not None
                self.life -= 1
                self.hold = None
                self.harpoon = None
                self.diver_col = GRID_COLUMNS // 2
                self.diver_row = -1
                if self.life <= 0:
                    self.life = 0
                    self.state = "game_over"
                    if self.score > self.best_score:
                        self.best_score = self.score
                        self.best_dirty = True
                    self.message_key = "status_game_over"
                    self.message_args = {}
                else:
                    self.message_key = "status_hit_lost" if had_hold else "status_hit"
                    self.message_args = {}
                return

    def _update_sharks(self, dt: float) -> None:
        interval = self.shark_interval()
        for shark in self.sharks:
            shark.timer -= dt
            while shark.timer <= 0:
                next_col = shark.col + shark.direction
                if not 0 <= next_col < GRID_COLUMNS:
                    shark.direction *= -1
                    next_col = shark.col + shark.direction
                shark.col = clamp(next_col, 0, GRID_COLUMNS - 1)
                shark.timer += interval

    def _update_harpoon(self, dt: float) -> None:
        if self.harpoon is None:
            return
        self.harpoon.timer -= dt
        while self.harpoon is not None and self.harpoon.timer <= 0:
            self.harpoon.col += self.harpoon.direction
            self.harpoon.steps_left -= 1
            self.harpoon.timer += HARPOON_STEP_SECONDS
            if not is_cell_in_grid(self.harpoon.col, self.harpoon.row) or self.harpoon.steps_left < 0:
                self.harpoon = None
                self.message_key = "status_missed"
                self.message_args = {}
                return
            self._catch_prey_at(self.harpoon.col, self.harpoon.row)

    def _catch_prey_at(self, col: int, row: int) -> bool:
        if self.hold is not None:
            return False
        for index, item in enumerate(self.prey):
            if item.col == col and item.row == row:
                self.hold = item.kind
                del self.prey[index]
                self.harpoon = None
                self.message_key = "status_caught"
                self.message_args = {"item_key": f"hold_{item.kind}"}
                self._ensure_prey(min_count=1)
                return True
        return False

    def _ensure_sharks(self) -> None:
        target_count = 2 if self.elapsed >= 30.0 or self.score >= 150 else 1
        while len(self.sharks) < target_count:
            used_rows = {shark.row for shark in self.sharks}
            row_choices = [row for row in range(1, GRID_ROWS) if row not in used_rows]
            row = self.random.choice(row_choices or list(range(1, GRID_ROWS)))
            direction = self.random.choice([-1, 1])
            col = 0 if direction > 0 else GRID_COLUMNS - 1
            timer = self.random.uniform(0.15, self.shark_interval())
            self.sharks.append(Shark(col=col, row=row, direction=direction, timer=timer))

    def _ensure_prey(self, min_count: int = 1) -> None:
        target_count = self.random.randint(max(1, min_count), 3)
        while len(self.prey) < target_count:
            cell = self._random_empty_cell()
            if cell is None:
                return
            self.prey.append(Prey(kind=self._random_prey_kind(), col=cell[0], row=cell[1]))

    def _random_empty_cell(self) -> tuple[int, int] | None:
        occupied = {(item.col, item.row) for item in self.prey}
        occupied.update((shark.col, shark.row) for shark in self.sharks)
        if self.diver_row >= 0:
            occupied.add((self.diver_col, self.diver_row))

        cells = [
            (col, row)
            for row in range(GRID_ROWS)
            for col in range(GRID_COLUMNS)
            if (col, row) not in occupied
        ]
        if not cells:
            return None
        return self.random.choice(cells)

    def _random_prey_kind(self) -> str:
        roll = self.random.random()
        if roll < 0.58:
            return "small_fish"
        if roll < 0.88:
            return "big_fish"
        return "treasure"


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
        self.bind("<Return>", lambda _event: self._handle_enter())
        self.bind("<r>", lambda _event: self._restart())
        self.bind("<R>", lambda _event: self._restart())
        self.bind("<space>", lambda _event: self._fire())
        self.bind("<p>", lambda _event: self._toggle_pause())
        self.bind("<P>", lambda _event: self._toggle_pause())
        self.bind("<Up>", lambda _event: self._move(0, -1))
        self.bind("<Down>", lambda _event: self._move(0, 1))
        self.bind("<Left>", lambda _event: self._move(-1, 0))
        self.bind("<Right>", lambda _event: self._move(1, 0))
        self.bind("<w>", lambda _event: self._move(0, -1))
        self.bind("<W>", lambda _event: self._move(0, -1))
        self.bind("<s>", lambda _event: self._move(0, 1))
        self.bind("<S>", lambda _event: self._move(0, 1))
        self.bind("<a>", lambda _event: self._move(-1, 0))
        self.bind("<A>", lambda _event: self._move(-1, 0))
        self.bind("<d>", lambda _event: self._move(1, 0))
        self.bind("<D>", lambda _event: self._move(1, 0))
        self.focus_force()

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
        self._draw_footer()
        if self.model.state == "idle":
            self._draw_center_overlay(UI_TEXT["display_name"], UI_TEXT["ready_hint"])
        elif self.model.state == "paused":
            self._draw_center_overlay(UI_TEXT["paused_title"], UI_TEXT["paused_hint"])
        elif self.model.state == "game_over":
            self._draw_center_overlay(UI_TEXT["game_over_title"], UI_TEXT["game_over_hint"])

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
        self.canvas.create_rectangle(GRID_LEFT, SHIP_Y, GRID_LEFT + GRID_COLUMNS * CELL_WIDTH, SEA_TOP, fill=COLORS["water"], outline=COLORS["ink"], width=2)
        self.canvas.create_line(GRID_LEFT, SEA_TOP, GRID_LEFT + GRID_COLUMNS * CELL_WIDTH, SEA_TOP, fill=COLORS["ink"], width=3)
        for col in range(GRID_COLUMNS):
            x = GRID_LEFT + col * CELL_WIDTH + CELL_WIDTH / 2
            self.canvas.create_arc(x - 16, SEA_TOP - 10, x + 16, SEA_TOP + 12, start=0, extent=180, outline=COLORS["lcd_shadow"], width=1)

        boat_x = self._cell_center(self.model.diver_col, 0)[0] if self.model.diver_row < 0 else WINDOW_WIDTH / 2
        boat_y = SHIP_Y + 30
        self.canvas.create_polygon(
            boat_x - 54,
            boat_y + 10,
            boat_x + 54,
            boat_y + 10,
            boat_x + 38,
            boat_y + 28,
            boat_x - 38,
            boat_y + 28,
            fill=COLORS["ink"],
            outline=COLORS["ink"],
        )
        self.canvas.create_rectangle(boat_x - 30, boat_y - 8, boat_x + 22, boat_y + 10, fill=COLORS["highlight"], outline=COLORS["ink"], width=2)
        self.canvas.create_line(boat_x + 34, boat_y + 10, boat_x + 34, SEA_TOP, fill=COLORS["ink"], width=2)

    def _draw_grid(self) -> None:
        x1 = GRID_LEFT
        y1 = GRID_TOP
        x2 = GRID_LEFT + GRID_COLUMNS * CELL_WIDTH
        y2 = GRID_TOP + GRID_ROWS * CELL_HEIGHT
        self.canvas.create_rectangle(x1, y1, x2, y2, fill=COLORS["water"], outline=COLORS["ink"], width=2)
        for col in range(1, GRID_COLUMNS):
            x = GRID_LEFT + col * CELL_WIDTH
            self.canvas.create_line(x, y1, x, y2, fill=COLORS["lcd_shadow"], width=1)
        for row in range(1, GRID_ROWS):
            y = GRID_TOP + row * CELL_HEIGHT
            self.canvas.create_line(x1, y, x2, y, fill=COLORS["lcd_shadow"], width=1)

    def _draw_prey(self) -> None:
        for item in self.model.prey:
            x, y = self._cell_center(item.col, item.row)
            if item.kind == "small_fish":
                self._draw_fish(x, y, 18, 9)
            elif item.kind == "big_fish":
                self._draw_fish(x, y, 28, 13)
            else:
                self._draw_treasure(x, y)

    def _draw_sharks(self) -> None:
        for shark in self.model.sharks:
            x, y = self._cell_center(shark.col, shark.row)
            direction = 1 if shark.direction >= 0 else -1
            self.canvas.create_polygon(
                x - 28 * direction,
                y,
                x - 8 * direction,
                y - 15,
                x + 22 * direction,
                y - 8,
                x + 30 * direction,
                y,
                x + 22 * direction,
                y + 8,
                x - 8 * direction,
                y + 15,
                fill=COLORS["ink"],
                outline=COLORS["ink"],
            )
            self.canvas.create_polygon(
                x - 2 * direction,
                y - 14,
                x + 8 * direction,
                y - 28,
                x + 12 * direction,
                y - 12,
                fill=COLORS["ink"],
                outline=COLORS["ink"],
            )
            self.canvas.create_oval(x + 15 * direction - 2, y - 5, x + 15 * direction + 2, y - 1, fill=COLORS["highlight"], outline="")

    def _draw_diver(self) -> None:
        if self.model.diver_row < 0:
            x = self._cell_center(self.model.diver_col, 0)[0]
            y = SHIP_Y + 23
        else:
            x, y = self._cell_center(self.model.diver_col, self.model.diver_row)
        direction = 1 if self.model.facing >= 0 else -1

        self.canvas.create_oval(x - 8, y - 19, x + 8, y - 3, fill=COLORS["highlight"], outline=COLORS["ink"], width=2)
        self.canvas.create_rectangle(x - 8, y - 3, x + 8, y + 17, fill=COLORS["ink"], outline=COLORS["ink"])
        self.canvas.create_line(x + 8 * direction, y + 2, x + 23 * direction, y + 7, fill=COLORS["ink"], width=3)
        self.canvas.create_line(x - 6, y + 17, x - 18, y + 26, fill=COLORS["ink"], width=3)
        self.canvas.create_line(x + 6, y + 17, x + 18, y + 26, fill=COLORS["ink"], width=3)
        self.canvas.create_line(x + 16 * direction, y - 10, x + 28 * direction, y - 12, fill=COLORS["ink"], width=2)
        if self.model.hold is not None:
            self.canvas.create_rectangle(x - 19, y - 3, x - 7, y + 9, fill=COLORS["highlight"], outline=COLORS["ink"], width=2)

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

    def _draw_fish(self, x: float, y: float, width: float, height: float) -> None:
        self.canvas.create_oval(x - width / 2, y - height, x + width / 2, y + height, fill=COLORS["ink"], outline=COLORS["ink"])
        self.canvas.create_polygon(
            x - width / 2,
            y,
            x - width / 2 - 10,
            y - 8,
            x - width / 2 - 10,
            y + 8,
            fill=COLORS["ink"],
            outline=COLORS["ink"],
        )
        self.canvas.create_oval(x + width / 4, y - 4, x + width / 4 + 3, y - 1, fill=COLORS["highlight"], outline="")

    def _draw_treasure(self, x: float, y: float) -> None:
        self.canvas.create_rectangle(x - 18, y - 10, x + 18, y + 13, fill=COLORS["ink"], outline=COLORS["ink"])
        self.canvas.create_arc(x - 18, y - 20, x + 18, y + 8, start=0, extent=180, fill=COLORS["ink"], outline=COLORS["ink"])
        self.canvas.create_rectangle(x - 4, y - 2, x + 4, y + 9, fill=COLORS["highlight"], outline=COLORS["highlight"])

    def _draw_footer(self) -> None:
        self.canvas.create_rectangle(10, 460, WINDOW_WIDTH - 10, 530, fill=COLORS["panel"], outline=COLORS["ink"], width=2)
        self.canvas.create_text(22, 482, text=self._status_text(), fill=COLORS["ink"], font=FONTS["label"], anchor="w")
        self.canvas.create_text(22, 510, text=UI_TEXT["playing_hint"], fill=COLORS["ink_soft"], font=FONTS["small"], anchor="w")
        self.canvas.create_text(WINDOW_WIDTH - 22, 510, text=COPYRIGHT, fill=COLORS["ink_soft"], font=FONTS["small"], anchor="e")

    def _draw_center_overlay(self, title: str, hint: str) -> None:
        self.canvas.create_rectangle(48, 224, WINDOW_WIDTH - 48, 318, fill=COLORS["lcd_bg"], outline=COLORS["ink"], width=2)
        self.canvas.create_text(WINDOW_WIDTH / 2, 258, text=title, fill=COLORS["ink"], font=FONTS["title"])
        self.canvas.create_text(WINDOW_WIDTH / 2, 294, text=hint, fill=COLORS["ink"], font=FONTS["lcd"])

    def _cell_center(self, col: int, row: int) -> tuple[float, float]:
        return (
            GRID_LEFT + col * CELL_WIDTH + CELL_WIDTH / 2,
            GRID_TOP + row * CELL_HEIGHT + CELL_HEIGHT / 2,
        )

    def _hold_text(self) -> str:
        if self.model.hold is None:
            return UI_TEXT["hold_empty"]
        return UI_TEXT[f"hold_{self.model.hold}"]

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
        model.start_game()
        assert model.state == "playing"
        assert model.life == MAX_LIFE
        assert 1 <= len(model.sharks) <= 2
        assert 1 <= len(model.prey) <= 3

        model.prey = [Prey("small_fish", 3, 0)]
        model.sharks = []
        model.diver_col = 2
        model.diver_row = 0
        model.facing = 1
        model.fire_harpoon()
        assert model.hold == "small_fish"
        model.move(0, -1)
        assert model.score == 10
        assert model.hold is None

        model.hold = "treasure"
        model.diver_row = 0
        model.move(0, -1)
        assert model.score == 110
        assert model.best_score == 110
        assert model.best_dirty
        save_best_score(model.best_score, config_path)
        assert load_best_score(config_path) == 110

        model.best_dirty = False
        model.state = "playing"
        model.life = 1
        model.hold = "big_fish"
        model.diver_col = 1
        model.diver_row = 2
        model.sharks = [Shark(col=1, row=2, direction=1, timer=1.0)]
        model.update(0.01)
        assert model.state == "game_over"
        assert model.life == 0
        assert model.hold is None

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
