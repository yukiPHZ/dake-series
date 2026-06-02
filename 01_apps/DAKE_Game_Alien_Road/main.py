# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import random
import sys
import tempfile
import time
import urllib.parse
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from types import SimpleNamespace
import tkinter as tk


APP_NAME = "DAKE Alien Road"
WINDOW_TITLE = "エイリアンロード"
COPYRIGHT = "© 2026 しまりす不動産 / Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "ready_title": "エイリアンロード",
    "ready_hint": "SPACE / Enter で開始",
    "running_status": "走行中",
    "game_over": "GAME OVER",
    "game_over_hint": "結果を残して、もう一度走れます",
    "score_label": "SCORE",
    "stage_label": "STAGE",
    "speed_label": "SPEED",
    "best_label": "BEST",
    "button_left": "← 左",
    "button_right": "右 →",
    "button_copy_post": "投稿文をコピー",
    "button_open_x": "X投稿画面を開く",
    "button_restart": "もう一度遊ぶ",
    "copied_status": "投稿文をコピーしました",
    "x_open_status": "X投稿画面を開きました",
    "post_score": "スコア",
    "post_tagline": "左か右か、それDAKE。",
    "post_hashtag": "#DAKEAlienRoad",
    "launch_check_ok": "launch-check ok",
}

CONFIG_NAME = "alien_road_config.json"
WINDOW_WIDTH = 480
WINDOW_HEIGHT = 720
LANE_COUNT = 5
HORIZON_Y = 126
ROAD_BOTTOM_Y = 598
PLAYER_Y = 552
BASE_SPEED = 0.42
SPEED_GROWTH = 0.006
STAGE_SECONDS = 18.0

COLORS = {
    "lcd_bg": "#9fb77a",
    "lcd_shadow": "#839963",
    "ink": "#26391f",
    "ink_light": "#35502c",
    "road": "#8dab67",
    "road_line": "#51653a",
    "panel": "#7f9460",
    "button": "#718558",
    "button_active": "#64794e",
    "highlight": "#d4e4a8",
}

FONTS = {
    "title": ("Yu Gothic UI", 24, "bold"),
    "lcd_large": ("Consolas", 20, "bold"),
    "lcd": ("Consolas", 13, "bold"),
    "label": ("Yu Gothic UI", 11, "bold"),
    "button": ("Yu Gothic UI", 12, "bold"),
}


@dataclass
class Obstacle:
    lane: float
    z: float
    kind: str
    move_dir: int = 0


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = get_app_dir()
CONFIG_PATH = APP_DIR / CONFIG_NAME


def load_best_score(path: Path = CONFIG_PATH) -> int:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return max(0, int(data.get("best_score", 0)))
    except Exception:
        return 0


def save_best_score(best_score: int, path: Path = CONFIG_PATH) -> None:
    text = json.dumps({"best_score": int(best_score)}, ensure_ascii=False, indent=2) + "\n"
    path.write_text(text, encoding="utf-8")


def clamp_lane(lane: int) -> int:
    return max(0, min(LANE_COUNT - 1, lane))


def rects_overlap(left: tuple[float, float, float, float], right: tuple[float, float, float, float]) -> bool:
    return left[0] < right[2] and left[2] > right[0] and left[1] < right[3] and left[3] > right[1]


def make_post_text(score: int, stage: int, speed: int) -> str:
    return "\n".join(
        [
            APP_NAME,
            f"{UI_TEXT['post_score']}: {score}",
            f"{UI_TEXT['stage_label']}: {stage}",
            f"{UI_TEXT['speed_label']}: {speed}",
            UI_TEXT["post_tagline"],
            UI_TEXT["post_hashtag"],
        ]
    )


def make_x_post_url(post_text: str) -> str:
    return "https://twitter.com/intent/tweet?text=" + urllib.parse.quote(post_text)


class AlienRoadApp(tk.Tk):
    def __init__(self, config_path: Path = CONFIG_PATH, visible: bool = True) -> None:
        super().__init__()
        self.config_path = config_path
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

        self.random = random.Random()
        self.click_zones: list[tuple[tuple[int, int, int, int], str]] = []
        self.best_score = load_best_score(self.config_path)
        self.status_message = UI_TEXT["ready_hint"]
        self.state = "ready"
        self.player_lane = LANE_COUNT // 2
        self.score = 0
        self.stage = 1
        self.elapsed = 0.0
        self.current_speed = BASE_SPEED
        self.spawn_timer = 0.8
        self.road_phase = 0.0
        self.obstacles: list[Obstacle] = []
        self.last_time = time.perf_counter()

        if not visible:
            self.withdraw()

        self._bind_controls()
        self._draw()
        self.after(16, self._tick)

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
        self.bind("<Left>", lambda _event: self.move_player(-1))
        self.bind("<Right>", lambda _event: self.move_player(1))
        self.bind("<a>", lambda _event: self.move_player(-1))
        self.bind("<A>", lambda _event: self.move_player(-1))
        self.bind("<d>", lambda _event: self.move_player(1))
        self.bind("<D>", lambda _event: self.move_player(1))
        self.bind("<space>", lambda _event: self.start_or_restart())
        self.bind("<Return>", lambda _event: self.start_or_restart())
        self.canvas.bind("<Button-1>", self._handle_click)
        self.focus_force()

    def start_or_restart(self) -> None:
        if self.state in {"ready", "game_over"}:
            self.start_game()

    def start_game(self) -> None:
        self.state = "running"
        self.player_lane = LANE_COUNT // 2
        self.score = 0
        self.stage = 1
        self.elapsed = 0.0
        self.current_speed = BASE_SPEED
        self.spawn_timer = 0.5
        self.road_phase = 0.0
        self.obstacles.clear()
        self.status_message = UI_TEXT["running_status"]
        self.last_time = time.perf_counter()
        self._draw()

    def move_player(self, direction: int) -> None:
        if self.state not in {"ready", "running"}:
            return
        self.player_lane = clamp_lane(self.player_lane + direction)
        self._draw()

    def _handle_click(self, event: tk.Event) -> None:
        for rect, action in reversed(self.click_zones):
            x1, y1, x2, y2 = rect
            if x1 <= event.x <= x2 and y1 <= event.y <= y2:
                self._run_click_action(action)
                return

    def _run_click_action(self, action: str) -> None:
        if action == "left":
            self.move_player(-1)
        elif action == "right":
            self.move_player(1)
        elif action == "copy_post":
            self.copy_post_text()
        elif action == "open_x":
            self.open_x_post()
        elif action == "restart":
            self.start_game()

    def _tick(self) -> None:
        now = time.perf_counter()
        dt = min(0.05, now - self.last_time)
        self.last_time = now

        if self.state == "running":
            self._update_game(dt)
            self._draw()

        self.after(16, self._tick)

    def _update_game(self, dt: float) -> None:
        self.elapsed += dt
        self.stage = 1 + int(self.elapsed // STAGE_SECONDS)
        self.current_speed = BASE_SPEED + (self.elapsed * SPEED_GROWTH) + ((self.stage - 1) * 0.025)
        self.score += int((52 + self.current_speed * 110) * dt)
        self.road_phase = (self.road_phase + self.current_speed * dt * 1.8) % 1.0

        for obstacle in self.obstacles:
            obstacle.z += self.current_speed * dt
            if obstacle.kind == "alien" and obstacle.z > 0.12:
                obstacle.lane += obstacle.move_dir * dt * (0.55 + self.stage * 0.04)
                if obstacle.lane <= 0:
                    obstacle.lane = 0
                    obstacle.move_dir = 1
                elif obstacle.lane >= LANE_COUNT - 1:
                    obstacle.lane = LANE_COUNT - 1
                    obstacle.move_dir = -1

        self.obstacles = [obstacle for obstacle in self.obstacles if obstacle.z <= 1.18]
        self.spawn_timer -= dt
        if self.spawn_timer <= 0:
            self._spawn_obstacles()

        if self._check_collision():
            self._game_over()

    def _spawn_obstacles(self) -> None:
        lanes = list(range(LANE_COUNT))
        self.random.shuffle(lanes)
        count = 2 if self.stage >= 4 and self.random.random() < 0.26 else 1
        for lane in lanes[:count]:
            alien_chance = min(0.28 + self.stage * 0.035, 0.55)
            kind = "alien" if self.random.random() < alien_chance else "block"
            move_dir = self.random.choice([-1, 1]) if kind == "alien" else 0
            self.obstacles.append(Obstacle(lane=float(lane), z=-0.02, kind=kind, move_dir=move_dir))

        interval = max(0.45, 1.05 - self.stage * 0.045 - self.elapsed * 0.002)
        self.spawn_timer = interval + self.random.uniform(-0.12, 0.18)

    def _check_collision(self) -> bool:
        player_bounds = self._player_bounds()
        for obstacle in self.obstacles:
            if obstacle.z < 0.68:
                continue
            if rects_overlap(player_bounds, self._obstacle_bounds(obstacle)):
                return True
        return False

    def _game_over(self) -> None:
        self.state = "game_over"
        self.status_message = UI_TEXT["game_over"]
        if self.score > self.best_score:
            self.best_score = self.score
            save_best_score(self.best_score, self.config_path)

    def copy_post_text(self) -> str:
        post_text = make_post_text(self.score, self.stage, self.display_speed())
        self.clipboard_clear()
        self.clipboard_append(post_text)
        self.status_message = UI_TEXT["copied_status"]
        self._draw()
        return post_text

    def open_x_post(self) -> str:
        url = make_x_post_url(make_post_text(self.score, self.stage, self.display_speed()))
        webbrowser.open(url)
        self.status_message = UI_TEXT["x_open_status"]
        self._draw()
        return url

    def display_speed(self) -> int:
        return max(1, int(round(self.current_speed * 10)))

    def _lane_center(self, lane: float, y: float) -> float:
        left, right = self._road_edges(y)
        lane_width = (right - left) / LANE_COUNT
        return left + lane_width * (lane + 0.5)

    def _road_edges(self, y: float) -> tuple[float, float]:
        progress = max(0.0, min(1.0, (y - HORIZON_Y) / (ROAD_BOTTOM_Y - HORIZON_Y)))
        half_width = 40 + (218 - 40) * (progress**1.35)
        return WINDOW_WIDTH / 2 - half_width, WINDOW_WIDTH / 2 + half_width

    def _y_from_z(self, z: float) -> float:
        progress = max(0.0, min(1.0, z))
        return HORIZON_Y + (ROAD_BOTTOM_Y - HORIZON_Y) * (progress**1.24)

    def _player_bounds(self) -> tuple[float, float, float, float]:
        x = self._lane_center(float(self.player_lane), PLAYER_Y)
        return x - 26, PLAYER_Y - 31, x + 26, PLAYER_Y + 24

    def _obstacle_bounds(self, obstacle: Obstacle) -> tuple[float, float, float, float]:
        y = self._y_from_z(obstacle.z)
        x = self._lane_center(obstacle.lane, y)
        progress = max(0.18, min(1.0, obstacle.z))
        width = 18 + 48 * progress
        height = 18 + 54 * progress
        return x - width / 2, y - height / 2, x + width / 2, y + height / 2

    def _draw(self) -> None:
        self.click_zones.clear()
        self.canvas.delete("all")
        self._draw_lcd_background()
        self._draw_header()
        self._draw_road()
        self._draw_obstacles()
        self._draw_player()
        self._draw_controls()
        self._draw_footer()
        if self.state == "ready":
            self._draw_ready_overlay()
        elif self.state == "game_over":
            self._draw_game_over_overlay()

    def _draw_lcd_background(self) -> None:
        self.canvas.create_rectangle(0, 0, WINDOW_WIDTH, WINDOW_HEIGHT, fill=COLORS["lcd_bg"], outline="")
        for y in range(0, WINDOW_HEIGHT, 12):
            self.canvas.create_line(0, y, WINDOW_WIDTH, y, fill=COLORS["lcd_shadow"], width=1)

    def _draw_header(self) -> None:
        self.canvas.create_rectangle(12, 10, WINDOW_WIDTH - 12, 82, fill=COLORS["panel"], outline=COLORS["ink"], width=2)
        stats = [
            (UI_TEXT["score_label"], str(self.score)),
            (UI_TEXT["stage_label"], str(self.stage)),
            (UI_TEXT["speed_label"], str(self.display_speed())),
            (UI_TEXT["best_label"], str(self.best_score)),
        ]
        column_width = (WINDOW_WIDTH - 40) / len(stats)
        for index, (label, value) in enumerate(stats):
            x = 22 + column_width * index
            self.canvas.create_text(x, 24, text=label, fill=COLORS["ink"], font=FONTS["lcd"], anchor="w")
            self.canvas.create_text(x, 55, text=value, fill=COLORS["ink"], font=FONTS["lcd_large"], anchor="w")
            if index:
                self.canvas.create_line(x - 8, 20, x - 8, 70, fill=COLORS["road_line"], width=1)

    def _draw_road(self) -> None:
        top_left, top_right = self._road_edges(HORIZON_Y)
        bottom_left, bottom_right = self._road_edges(ROAD_BOTTOM_Y)
        self.canvas.create_polygon(
            top_left,
            HORIZON_Y,
            top_right,
            HORIZON_Y,
            bottom_right,
            ROAD_BOTTOM_Y,
            bottom_left,
            ROAD_BOTTOM_Y,
            fill=COLORS["road"],
            outline=COLORS["ink"],
            width=2,
        )

        for lane_index in range(1, LANE_COUNT):
            top_x = top_left + (top_right - top_left) * lane_index / LANE_COUNT
            bottom_x = bottom_left + (bottom_right - bottom_left) * lane_index / LANE_COUNT
            self.canvas.create_line(top_x, HORIZON_Y, bottom_x, ROAD_BOTTOM_Y, fill=COLORS["road_line"], width=2)

        for band in range(13):
            progress = ((band + self.road_phase) / 13) ** 1.45
            y = HORIZON_Y + (ROAD_BOTTOM_Y - HORIZON_Y) * progress
            left, right = self._road_edges(y)
            self.canvas.create_line(left, y, right, y, fill=COLORS["road_line"], width=1)

        self.canvas.create_line(top_left, HORIZON_Y, bottom_left, ROAD_BOTTOM_Y, fill=COLORS["ink"], width=3)
        self.canvas.create_line(top_right, HORIZON_Y, bottom_right, ROAD_BOTTOM_Y, fill=COLORS["ink"], width=3)

    def _draw_obstacles(self) -> None:
        for obstacle in sorted(self.obstacles, key=lambda item: item.z):
            y = self._y_from_z(obstacle.z)
            x = self._lane_center(obstacle.lane, y)
            progress = max(0.18, min(1.0, obstacle.z))
            width = 18 + 48 * progress
            height = 18 + 54 * progress
            if obstacle.kind == "alien":
                self._draw_alien(x, y, width, height)
            else:
                self._draw_block(x, y, width, height)

    def _draw_block(self, x: float, y: float, width: float, height: float) -> None:
        self.canvas.create_rectangle(
            x - width / 2,
            y - height / 2,
            x + width / 2,
            y + height / 2,
            fill=COLORS["ink_light"],
            outline=COLORS["ink"],
            width=2,
        )
        self.canvas.create_line(x - width / 2, y, x + width / 2, y, fill=COLORS["lcd_bg"], width=2)
        self.canvas.create_line(x, y - height / 2, x, y + height / 2, fill=COLORS["lcd_bg"], width=2)

    def _draw_alien(self, x: float, y: float, width: float, height: float) -> None:
        self.canvas.create_polygon(
            x,
            y - height / 2,
            x + width / 2,
            y,
            x + width * 0.22,
            y + height / 2,
            x - width * 0.22,
            y + height / 2,
            x - width / 2,
            y,
            fill=COLORS["ink"],
            outline=COLORS["ink"],
        )
        eye_size = max(2, width * 0.08)
        self.canvas.create_oval(x - width * 0.18, y - eye_size, x - width * 0.18 + eye_size, y + eye_size, fill=COLORS["highlight"], outline="")
        self.canvas.create_oval(x + width * 0.18 - eye_size, y - eye_size, x + width * 0.18, y + eye_size, fill=COLORS["highlight"], outline="")

    def _draw_player(self) -> None:
        x = self._lane_center(float(self.player_lane), PLAYER_Y)
        y = PLAYER_Y
        self.canvas.create_polygon(
            x,
            y - 36,
            x + 31,
            y + 16,
            x + 12,
            y + 12,
            x + 5,
            y + 28,
            x - 5,
            y + 28,
            x - 12,
            y + 12,
            x - 31,
            y + 16,
            fill=COLORS["ink"],
            outline=COLORS["lcd_bg"],
            width=2,
        )
        self.canvas.create_rectangle(x - 7, y - 16, x + 7, y + 8, fill=COLORS["highlight"], outline=COLORS["ink"])

    def _draw_controls(self) -> None:
        self._draw_button((24, 624, 208, 680), UI_TEXT["button_left"], "left")
        self._draw_button((272, 624, 456, 680), UI_TEXT["button_right"], "right")

    def _draw_footer(self) -> None:
        self.canvas.create_text(20, 700, text=self.status_message, fill=COLORS["ink"], font=FONTS["label"], anchor="w")
        self.canvas.create_text(WINDOW_WIDTH - 20, 700, text=COPYRIGHT, fill=COLORS["ink"], font=("Yu Gothic UI", 8), anchor="e")

    def _draw_ready_overlay(self) -> None:
        self.canvas.create_rectangle(46, 226, WINDOW_WIDTH - 46, 348, fill=COLORS["lcd_bg"], outline=COLORS["ink"], width=2)
        self.canvas.create_text(WINDOW_WIDTH / 2, 266, text=UI_TEXT["ready_title"], fill=COLORS["ink"], font=FONTS["title"])
        self.canvas.create_text(WINDOW_WIDTH / 2, 310, text=UI_TEXT["ready_hint"], fill=COLORS["ink"], font=FONTS["label"])

    def _draw_game_over_overlay(self) -> None:
        self.canvas.create_rectangle(34, 202, WINDOW_WIDTH - 34, 498, fill=COLORS["lcd_bg"], outline=COLORS["ink"], width=2)
        self.canvas.create_text(WINDOW_WIDTH / 2, 238, text=UI_TEXT["game_over"], fill=COLORS["ink"], font=FONTS["lcd_large"])
        self.canvas.create_text(WINDOW_WIDTH / 2, 270, text=UI_TEXT["game_over_hint"], fill=COLORS["ink"], font=FONTS["label"])
        self.canvas.create_text(
            WINDOW_WIDTH / 2,
            306,
            text=f"{UI_TEXT['score_label']} {self.score}  {UI_TEXT['stage_label']} {self.stage}  {UI_TEXT['speed_label']} {self.display_speed()}",
            fill=COLORS["ink"],
            font=FONTS["lcd"],
        )
        self._draw_button((70, 338, 410, 376), UI_TEXT["button_copy_post"], "copy_post")
        self._draw_button((70, 388, 410, 426), UI_TEXT["button_open_x"], "open_x")
        self._draw_button((70, 438, 410, 476), UI_TEXT["button_restart"], "restart")

    def _draw_button(self, rect: tuple[int, int, int, int], label: str, action: str) -> None:
        x1, y1, x2, y2 = rect
        self.canvas.create_rectangle(x1, y1, x2, y2, fill=COLORS["button"], outline=COLORS["ink"], width=2)
        self.canvas.create_line(x1 + 4, y2 - 5, x2 - 4, y2 - 5, fill=COLORS["button_active"], width=2)
        self.canvas.create_text((x1 + x2) / 2, (y1 + y2) / 2, text=label, fill=COLORS["ink"], font=FONTS["button"])
        self.click_zones.append((rect, action))


def run_launch_check() -> None:
    with tempfile.TemporaryDirectory() as temp_dir:
        config_path = Path(temp_dir) / CONFIG_NAME
        save_best_score(12, config_path)
        assert load_best_score(config_path) == 12

        app = AlienRoadApp(config_path=config_path, visible=False)
        app.update()
        app.start_or_restart()
        assert app.state == "running"

        app.player_lane = 2
        app._handle_click(SimpleNamespace(x=60, y=650))
        assert app.player_lane == 1
        app._handle_click(SimpleNamespace(x=420, y=650))
        assert app.player_lane == 2
        app.move_player(-1)
        assert app.player_lane == 1

        app.player_lane = 2
        app.obstacles = [Obstacle(lane=2.0, z=0.92, kind="block")]
        assert app._check_collision()
        app._game_over()
        assert app.state == "game_over"
        assert app.best_score >= app.score

        post_text = app.copy_post_text()
        assert APP_NAME in post_text
        assert UI_TEXT["post_tagline"] in post_text
        assert UI_TEXT["post_hashtag"] in post_text
        x_url = make_x_post_url(post_text)
        assert x_url.startswith("https://twitter.com/intent/tweet?text=")
        assert "%23DAKEAlienRoad" in x_url
        app.destroy()

    print(UI_TEXT["launch_check_ok"])


def main() -> None:
    if "--launch-check" in sys.argv:
        run_launch_check()
        return

    app = AlienRoadApp()
    app.mainloop()


if __name__ == "__main__":
    main()
