# -*- coding: utf-8 -*-
"""Capture DAKE formal release evidence.

This tool adds the "出荷確認動画 / 動作証跡" part of the DAKE formal shipping
line. The video is not YouTube content production. It is local release evidence
created while confirming the app can launch and its main operation was checked.

Windows app recording is handled with ffmpeg desktop capture. pywinauto is used
only as an optional focus helper when it is installed. Playwright is not assumed
to control Tkinter executables; web checks stay separate from exe checks.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any

from release_source_policy import app_dirs as source_app_dirs, app_url_for, find_app as source_find_app, read_app_source, site_slug


ROOT = Path(__file__).resolve().parents[1]
APPS_DIR = ROOT / "01_apps"
DEFAULT_SITE_ROOT = Path(os.environ.get("DAKEAPP_SITE_ROOT", r"C:\Users\yukiz\devlop\dakeapp-site"))
JST = timezone(timedelta(hours=9))
RELEASE_DIR_NAME = "release_artifacts"
DEMO_MP4 = "demo.mp4"
DEMO_WEBM = "demo.webm"
SHIPPING_STATUS = "available"


@dataclass
class StepState:
    ok: bool = False
    detail: str = ""


@dataclass
class CaptureOutcome:
    app_dir: Path
    meta: dict[str, Any]
    app_url: str
    release_dir: Path
    checks: dict[str, StepState] = field(default_factory=dict)
    video_path: Path | None = None
    video_webm_path: Path | None = None
    stage: str = "failed"
    reason: str = ""
    release_record: dict[str, Any] = field(default_factory=dict)

    @property
    def app_key(self) -> str:
        return str(self.meta.get("app_key") or self.app_dir.name)

    @property
    def display_name(self) -> str:
        return str(
            self.meta.get("display_name")
            or self.meta.get("site_title")
            or self.meta.get("launcher_title")
            or self.app_dir.name
        )


def now_iso() -> str:
    return datetime.now(JST).isoformat(timespec="seconds")


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text.rstrip() + "\n", encoding="utf-8", newline="\n")


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n")


def extract_json_section(text: str, heading: str) -> dict[str, Any] | None:
    match = re.search(rf"(?s)^\s*##\s+{re.escape(heading)}\s*```json\s*(.*?)\s*```", text, re.MULTILINE)
    if not match:
        return None
    loaded = json.loads(match.group(1))
    return loaded if isinstance(loaded, dict) else None


def read_meta(app_dir: Path) -> tuple[dict[str, Any], str]:
    source = read_app_source(app_dir, ROOT)
    return source.meta, source.error


def app_dirs(only_available: bool) -> list[Path]:
    dirs: list[Path] = []
    for path in source_app_dirs(APPS_DIR):
        if only_available and read_app_source(path, ROOT).status != SHIPPING_STATUS:
            continue
        dirs.append(path)
    return dirs


def find_app(identifier: str) -> Path:
    return source_find_app(APPS_DIR, identifier, ROOT)



def rel(path: Path, base: Path) -> str:
    try:
        return path.relative_to(base).as_posix()
    except ValueError:
        return str(path)


def find_booth_product(app_dir: Path) -> Path | None:
    for candidate in (app_dir / "booth_ready" / "booth_product.txt", app_dir / "booth_product.txt"):
        if candidate.exists():
            return candidate
    return None


def booth_url_state(app_dir: Path) -> tuple[bool, str]:
    product = find_booth_product(app_dir)
    if product is None:
        return False, "booth_product.txt missing"
    text = read_text(product)
    match = re.search(r"(?ms)^# URL\s*\n(.*?)(?=\n# |\Z)", text)
    if not match:
        return False, "# URL section missing"
    url = match.group(1).strip()
    if not url:
        return False, "BOOTH URL empty"
    return True, url


def find_exe(app_dir: Path, meta: dict[str, Any]) -> Path | None:
    exe_name = str(meta.get("exe_name") or "").strip()
    if exe_name:
        candidate = app_dir / "dist" / exe_name
        if candidate.exists():
            return candidate
    dist = app_dir / "dist"
    exes = sorted(dist.glob("*.exe")) if dist.exists() else []
    return exes[0] if exes else None


def has_launch_check(app_dir: Path) -> bool:
    main_py = app_dir / "main.py"
    if not main_py.exists():
        return False
    text = read_text(main_py)
    return "--launch-check" in text or "launch_check" in text or "run_launch_check" in text


def terminate_process_tree(process: subprocess.Popen[Any]) -> None:
    if process.poll() is not None:
        return
    if os.name == "nt":
        subprocess.run(
            ["taskkill", "/PID", str(process.pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
    else:
        process.terminate()


def run_launch_check(app_dir: Path, exe_path: Path, timeout_seconds: float) -> StepState:
    if has_launch_check(app_dir):
        args = [str(exe_path), "--launch-check"]
    else:
        return StepState(False, "--launch-check not implemented; use --manual-operation-ok after GUI smoke")
    try:
        completed = subprocess.run(
            args,
            cwd=str(app_dir),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=timeout_seconds,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return StepState(False, "--launch-check timed out")
    except Exception as exc:
        return StepState(False, str(exc))
    output = (completed.stderr or completed.stdout or "").strip().replace("\r\n", " ").replace("\n", " ")
    return StepState(completed.returncode == 0, output or f"exit={completed.returncode}")


def run_operation_command(command: str, app_dir: Path, timeout_seconds: float) -> StepState:
    if not command:
        return StepState(False, "operation command not provided")
    try:
        completed = subprocess.run(
            command,
            cwd=str(app_dir),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            shell=True,
            timeout=timeout_seconds,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return StepState(False, "operation command timed out")
    except Exception as exc:
        return StepState(False, str(exc))
    output = (completed.stderr or completed.stdout or "").strip().replace("\r\n", " ").replace("\n", " ")
    return StepState(completed.returncode == 0, output or f"exit={completed.returncode}")


def focus_with_pywinauto(process_id: int) -> str:
    try:
        from pywinauto import Application  # type: ignore
    except Exception:
        return "pywinauto not installed"
    try:
        app = Application(backend="uia").connect(process=process_id, timeout=6)
        window = app.top_window()
        window.set_focus()
        return "focused with pywinauto"
    except Exception as exc:
        return f"pywinauto focus failed: {exc}"


def start_gui_app(app_dir: Path, exe_path: Path) -> tuple[subprocess.Popen[Any] | None, StepState]:
    try:
        process = subprocess.Popen(
            [str(exe_path)],
            cwd=str(app_dir),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as exc:
        return None, StepState(False, str(exc))
    time.sleep(1.5)
    if process.poll() is not None:
        return process, StepState(False, f"process exited early: {process.returncode}")
    detail = focus_with_pywinauto(process.pid)
    return process, StepState(True, detail)


def copy_input_video(source: Path, release_dir: Path) -> Path:
    suffix = source.suffix.lower()
    if suffix not in {".mp4", ".webm"}:
        raise ValueError("input video must be .mp4 or .webm")
    if suffix == ".webm":
        target = release_dir / DEMO_WEBM
    else:
        target = release_dir / DEMO_MP4
    release_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, target)
    return target


def record_screen(target: Path, seconds: float, fps: int) -> StepState:
    ffmpeg = shutil.which("ffmpeg")
    if not ffmpeg:
        return StepState(False, "ffmpeg not found in PATH")
    target.parent.mkdir(parents=True, exist_ok=True)
    args = [ffmpeg, "-y", "-f", "gdigrab", "-framerate", str(fps), "-i", "desktop", "-t", str(seconds)]
    if target.suffix.lower() == ".webm":
        args.extend(["-c:v", "libvpx-vp9", "-b:v", "1M"])
    else:
        args.extend(["-pix_fmt", "yuv420p"])
    args.append(str(target))
    try:
        completed = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=False)
    except Exception as exc:
        return StepState(False, str(exc))
    if completed.returncode != 0:
        tail = (completed.stderr or completed.stdout or "").splitlines()[-4:]
        return StepState(False, "ffmpeg failed: " + " ".join(tail))
    return StepState(target.exists() and target.stat().st_size > 0, rel(target, target.parent.parent))


def existing_video(release_dir: Path) -> Path | None:
    for name in (DEMO_MP4, DEMO_WEBM):
        candidate = release_dir / name
        if candidate.exists() and candidate.stat().st_size > 0:
            return candidate
    return None


def verify_cloudflare(url: str, timeout_seconds: float) -> StepState:
    try:
        request = urllib.request.Request(url, headers={"User-Agent": "DAKE-release-capture/1.0"})
        with urllib.request.urlopen(request, timeout=timeout_seconds) as response:
            status = getattr(response, "status", 0)
            ok = 200 <= int(status) < 400
            return StepState(ok, f"HTTP {status}")
    except urllib.error.HTTPError as exc:
        return StepState(False, f"HTTP {exc.code}")
    except Exception as exc:
        return StepState(False, str(exc))


def file_state(path: Path, label: str) -> StepState:
    return StepState(path.exists(), rel(path, path.parent.parent) if path.exists() else f"{label} missing")



def build_release_record(outcome: CaptureOutcome, args: argparse.Namespace) -> dict[str, Any]:
    app_dir = outcome.app_dir
    meta = outcome.meta
    source = read_app_source(app_dir, ROOT)
    screenshot_path = str(meta.get("screenshot_path") or "assets/screenshot.webp")
    thumbnail_path = "assets/booth_thumbnail.jpg"
    video_rel = rel(outcome.video_path, app_dir) if outcome.video_path else ""
    mp4_rel = video_rel if outcome.video_path and outcome.video_path.suffix.lower() == ".mp4" else ""
    webm_rel = rel(outcome.video_webm_path, app_dir) if outcome.video_webm_path else ""
    checks = {
        name: {"ok": state.ok, "detail": state.detail}
        for name, state in outcome.checks.items()
    }
    record = {
        "app_key": str(meta.get("app_key") or app_dir.name),
        "display_name": outcome.display_name,
        "release_url": str(meta.get("release_url") or ""),
        "app_url": outcome.app_url,
        "screenshot_path": screenshot_path,
        "booth_thumbnail_path": thumbnail_path,
        "video_local_path": video_rel,
        "video_mp4_path": mp4_rel,
        "video_webm_path": webm_rel,
        "demo_video_url": str(meta.get("demo_video_url") or ""),
        "video_cloudinary_url": "",
        "video_r2_url": "",
        "social_release_path": str(meta.get("social_release_path") or f"{RELEASE_DIR_NAME}/social_release.json"),
        "created_at": now_iso(),
        "stage": outcome.stage,
        "reason": outcome.reason,
        "evidence_name": "出荷確認動画",
        "evidence_kind": "動作証跡",
        "youtube_upload": "not_implemented",
        "checks": checks,
        "source_policy": {
            "source_kind": source.source_kind,
            "source_path": source.source_label,
            "original_missing": source.original_missing,
            "meta_derivative_mismatch": source.derivative_mismatches,
        },
        "tool": {
            "name": "tools/release_capture.py",
            "version": 1,
            "recording": {
                "source": "input_video" if args.input_video else ("ffmpeg_gdigrab" if args.record_screen else "existing"),
                "seconds": args.seconds,
                "fps": args.fps,
            },
        },
    }
    return record


def write_release_log(outcome: CaptureOutcome) -> None:
    lines = [
        f"# DAKE 出荷ログ: {outcome.display_name}",
        "",
        f"- created_at: {now_iso()}",
        f"- stage: {outcome.stage}",
        f"- reason: {outcome.reason or '-'}",
        f"- app_url: {outcome.app_url}",
        f"- release_url: {outcome.meta.get('release_url') or ''}",
        "",
        "## 工程",
        "",
    ]
    for name, state in outcome.checks.items():
        mark = "OK" if state.ok else "NG"
        lines.append(f"- {mark}: {name} - {state.detail}")
    write_text(outcome.release_dir / "release_log.md", "\n".join(lines))

def capture_app(app_dir: Path, args: argparse.Namespace) -> CaptureOutcome:
    meta, error = read_meta(app_dir)
    release_dir = app_dir / RELEASE_DIR_NAME
    app_url = app_url_for(app_dir, meta, args.site_root)
    outcome = CaptureOutcome(app_dir=app_dir, meta=meta, app_url=app_url, release_dir=release_dir)
    release_dir.mkdir(parents=True, exist_ok=True)
    if error:
        outcome.checks["ORIGINAL / DAKE_META"] = StepState(False, error)
    else:
        outcome.checks["ORIGINAL / DAKE_META"] = StepState(True, "parsed from ORIGINAL.md or README fallback")

    screenshot = app_dir / str(meta.get("screenshot_path") or "assets/screenshot.webp")
    thumbnail = app_dir / "assets" / "booth_thumbnail.jpg"
    product = find_booth_product(app_dir)
    exe_path = find_exe(app_dir, meta)
    booth_ok, booth_detail = booth_url_state(app_dir)
    detail_slug = site_slug(app_dir, meta, args.site_root)
    detail_page = args.site_root / "public" / "apps" / detail_slug / "index.html"

    outcome.checks["アプリ起動確認"] = StepState(False, "dist/*.exe missing")
    outcome.checks["主要操作の確認"] = StepState(False, "manual operation confirmation required")
    outcome.checks["screenshot.webp 保存"] = file_state(screenshot, "screenshot.webp")
    outcome.checks["booth_thumbnail.jpg 生成"] = file_state(thumbnail, "booth_thumbnail.jpg")
    outcome.checks["booth_product.txt 生成"] = StepState(product is not None, rel(product, app_dir) if product else "booth_product.txt missing")
    outcome.checks["build / dist 生成"] = StepState(exe_path is not None, rel(exe_path, app_dir) if exe_path else "dist/*.exe missing")
    outcome.checks["GitHub Release"] = StepState(bool(meta.get("release_url")), str(meta.get("release_url") or "release_url missing"))
    outcome.checks["BOOTH ready"] = StepState((app_dir / "booth_ready").exists(), "booth_ready/" if (app_dir / "booth_ready").exists() else "booth_ready missing")
    outcome.checks["BOOTH URL確認"] = StepState(booth_ok, booth_detail)
    outcome.checks["dakeapp.com詳細ページ更新"] = StepState(detail_page.exists(), rel(detail_page, args.site_root) if detail_page.exists() else f"missing slug={detail_slug}")

    if exe_path is not None and args.launch_check:
        outcome.checks["アプリ起動確認"] = run_launch_check(app_dir, exe_path, args.timeout)

    if args.operation_command:
        outcome.checks["主要操作の確認"] = run_operation_command(args.operation_command, app_dir, args.timeout)
    elif args.manual_operation_ok:
        detail = args.operation_note or "manual operation confirmed during release capture"
        outcome.checks["主要操作の確認"] = StepState(True, detail)

    app_process: subprocess.Popen[Any] | None = None
    if args.input_video:
        if args.input_video.exists():
            try:
                copied = copy_input_video(args.input_video, release_dir)
                outcome.video_path = copied
                if copied.suffix.lower() == ".webm":
                    outcome.video_webm_path = copied
                outcome.checks["操作動画保存"] = StepState(True, rel(copied, app_dir))
            except Exception as exc:
                outcome.checks["操作動画保存"] = StepState(False, str(exc))
        else:
            outcome.checks["操作動画保存"] = StepState(False, f"input video missing: {args.input_video}")
    elif args.allow_existing_video and existing_video(release_dir):
        existing = existing_video(release_dir)
        outcome.video_path = existing
        if existing and existing.suffix.lower() == ".webm":
            outcome.video_webm_path = existing
        outcome.checks["操作動画保存"] = StepState(True, rel(existing, app_dir))
    elif args.record_screen:
        if exe_path is not None and args.launch_app:
            app_process, launch_state = start_gui_app(app_dir, exe_path)
            if not outcome.checks["アプリ起動確認"].ok:
                outcome.checks["アプリ起動確認"] = launch_state
        target_name = DEMO_WEBM if args.video_format == "webm" else DEMO_MP4
        target = release_dir / target_name
        outcome.checks["操作動画保存"] = record_screen(target, args.seconds, args.fps)
        if target.exists() and target.stat().st_size > 0:
            outcome.video_path = target
            if target.suffix.lower() == ".webm":
                outcome.video_webm_path = target
    else:
        outcome.checks["操作動画保存"] = StepState(False, "use --record-screen, --input-video, or --allow-existing-video")

    if app_process is not None:
        terminate_process_tree(app_process)

    if args.verify_cloudflare:
        outcome.checks["Cloudflare本番反映確認"] = verify_cloudflare(app_url, args.cloudflare_timeout)
    else:
        detail = "not requested"
        outcome.checks["Cloudflare本番反映確認"] = StepState(not args.require_cloudflare, detail)

    failed = [f"{name}: {state.detail}" for name, state in outcome.checks.items() if not state.ok]
    outcome.stage = "complete" if not failed else "failed"
    outcome.reason = "; ".join(failed)
    outcome.release_record = build_release_record(outcome, args)
    write_json(release_dir / "release.json", outcome.release_record)
    write_release_log(outcome)
    return outcome


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Capture DAKE release evidence video and release.json.")
    parser.add_argument("--app", action="append", help="app folder, app_key, or display_name")
    parser.add_argument("--all", action="store_true", help="process all DAKE apps")
    parser.add_argument("--only-available", action="store_true", help="with --all, process only status: available")
    parser.add_argument("--site-root", type=Path, default=DEFAULT_SITE_ROOT)
    parser.add_argument("--input-video", type=Path, help="copy an existing demo.mp4/demo.webm into release_artifacts")
    parser.add_argument("--record-screen", action="store_true", help="record desktop with ffmpeg gdigrab")
    parser.add_argument("--allow-existing-video", action="store_true", help="accept existing release_artifacts/demo.mp4 or demo.webm")
    parser.add_argument("--video-format", choices=["mp4", "webm"], default="mp4")
    parser.add_argument("--seconds", type=float, default=8.0)
    parser.add_argument("--fps", type=int, default=12)
    parser.add_argument("--launch-app", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--launch-check", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--manual-operation-ok", action="store_true", help="mark main operation as manually confirmed")
    parser.add_argument("--operation-note", default="")
    parser.add_argument("--operation-command", default="", help="command that verifies the main operation")
    parser.add_argument("--timeout", type=float, default=15.0)
    parser.add_argument("--verify-cloudflare", action="store_true")
    parser.add_argument("--require-cloudflare", action="store_true")
    parser.add_argument("--cloudflare-timeout", type=float, default=10.0)
    parser.add_argument("--fail-on-incomplete", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if not args.all and not args.app:
        raise SystemExit("use --app APP or --all")
    targets = app_dirs(args.only_available) if args.all else [find_app(item) for item in args.app or []]
    outcomes = [capture_app(app_dir, args) for app_dir in targets]
    for outcome in outcomes:
        print(f"{outcome.app_dir.name}: {outcome.stage} {outcome.reason}")
    if args.fail_on_incomplete and any(outcome.stage != "complete" for outcome in outcomes):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
