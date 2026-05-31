"""Check launchability of DAKE app executables.

This audit is intentionally limited to launch checks. It does not operate app
features, transform files, send mail, open BOOTH automation, or use real data.
Only ``status: available`` apps are checked when ``--only-available`` is used;
non-shipping statuses are listed separately.
"""

from __future__ import annotations

import argparse
import csv
import json
import subprocess
import time
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
APPS_DIR = ROOT / "01_apps"
REPORT_DIR = ROOT / "tools" / "reports"
SHIPPING_STATUS = "available"
SKIPPED_STATUSES = {"draft", "frozen", "experimental", "private", "internal"}


@dataclass
class LaunchResult:
    app: str
    display_name: str = ""
    status: str = "unknown"
    exe_path: str = ""
    exe_exists: bool = False
    launch_check_supported: bool = False
    result: str = "ERROR"
    exit_code: int | str = ""
    elapsed_seconds: float = 0.0
    note: str = ""


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def parse_dake_meta(readme_text: str) -> tuple[dict[str, Any] | None, str]:
    if "## DAKE_META" not in readme_text or "```json" not in readme_text:
        return None, "missing DAKE_META block"
    try:
        block = (
            readme_text.split("## DAKE_META", 1)[1]
            .split("```json", 1)[1]
            .split("```", 1)[0]
        )
        return json.loads(block), ""
    except IndexError:
        return None, "broken DAKE_META fence"
    except json.JSONDecodeError as exc:
        return None, f"invalid DAKE_META JSON: {exc}"


def app_dirs() -> list[Path]:
    return sorted(path for path in APPS_DIR.iterdir() if path.is_dir() and path.name.startswith("DAKE_"))


def metadata_for_app(app_dir: Path) -> tuple[dict[str, Any], str]:
    readme = app_dir / "README.md"
    if not readme.exists():
        return {}, "README.md missing"
    meta, error = parse_dake_meta(read_text(readme))
    return meta or {}, error


def find_exe(app_dir: Path, meta: dict[str, Any]) -> Path | None:
    exe_name = str(meta.get("exe_name") or "").strip()
    if exe_name:
        candidate = app_dir / "dist" / exe_name
        if candidate.exists():
            return candidate
    exes = sorted((app_dir / "dist").glob("*.exe")) if (app_dir / "dist").exists() else []
    return exes[0] if exes else None


def has_launch_check(app_dir: Path) -> bool:
    main_py = app_dir / "main.py"
    if not main_py.exists():
        return False
    text = read_text(main_py)
    return "--launch-check" in text or "launch_check" in text or "run_launch_check" in text


def terminate_process_tree(process: subprocess.Popen) -> None:
    if process.poll() is not None:
        return
    try:
        subprocess.run(
            ["taskkill", "/PID", str(process.pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=5,
            check=False,
        )
    except Exception:
        try:
            process.kill()
        except Exception:
            pass


def run_launch_check(exe_path: Path, timeout_seconds: float) -> tuple[str, int | str, float, str]:
    start = time.monotonic()
    try:
        process = subprocess.Popen(
            [str(exe_path), "--launch-check"],
            cwd=str(exe_path.parent.parent),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
    except Exception as exc:
        return "ERROR", "", time.monotonic() - start, str(exc)

    try:
        stdout, stderr = process.communicate(timeout=timeout_seconds)
    except subprocess.TimeoutExpired:
        terminate_process_tree(process)
        return "TIMEOUT", "timeout", time.monotonic() - start, "--launch-check did not finish"

    elapsed = time.monotonic() - start
    note = (stderr or "").strip().replace("\r\n", " ").replace("\n", " ")
    if not note:
        note = (stdout or "").strip().replace("\r\n", " ").replace("\n", " ")
    if process.returncode == 0:
        return "OK", process.returncode, elapsed, note
    return "ERROR", process.returncode, elapsed, note or "--launch-check returned non-zero"


def run_gui_smoke(exe_path: Path, gui_seconds: float) -> tuple[str, int | str, float, str]:
    start = time.monotonic()
    try:
        process = subprocess.Popen(
            [str(exe_path)],
            cwd=str(exe_path.parent.parent),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as exc:
        return "ERROR", "", time.monotonic() - start, str(exc)

    try:
        exit_code = process.wait(timeout=gui_seconds)
        elapsed = time.monotonic() - start
        if exit_code == 0:
            return "OK_GUI", exit_code, elapsed, "process exited normally during GUI smoke"
        return "ERROR", exit_code, elapsed, "process exited during GUI smoke"
    except subprocess.TimeoutExpired:
        elapsed = time.monotonic() - start
        terminate_process_tree(process)
        return "OK_GUI", "terminated", elapsed, "GUI process started and was stopped after smoke timeout"


def check_app(app_dir: Path, only_available: bool, timeout_seconds: float, gui_seconds: float) -> LaunchResult:
    meta, meta_error = metadata_for_app(app_dir)
    status = str(meta.get("status") or "unknown")
    display_name = str(meta.get("display_name") or meta.get("site_title") or app_dir.name)
    result = LaunchResult(app=app_dir.name, display_name=display_name, status=status)

    if only_available and status != SHIPPING_STATUS:
        result.result = "SKIPPED_STATUS"
        result.note = f"status={status}"
        return result

    if meta_error:
        result.result = "ERROR"
        result.note = meta_error
        return result

    exe_path = find_exe(app_dir, meta)
    if exe_path is None:
        result.result = "NO_EXE"
        result.note = "dist/*.exe not found"
        return result

    result.exe_path = str(exe_path.relative_to(ROOT))
    result.exe_exists = True
    result.launch_check_supported = has_launch_check(app_dir)

    if result.launch_check_supported:
        result.result, result.exit_code, result.elapsed_seconds, result.note = run_launch_check(exe_path, timeout_seconds)
    else:
        result.result, result.exit_code, result.elapsed_seconds, result.note = run_gui_smoke(exe_path, gui_seconds)
    return result


def write_csv(results: list[LaunchResult], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fields = [
        "app folder",
        "display_name",
        "status",
        "exe path",
        "exe exists",
        "launch_check supported",
        "result",
        "exit code",
        "elapsed seconds",
        "note",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fields)
        writer.writeheader()
        for item in results:
            writer.writerow(
                {
                    "app folder": item.app,
                    "display_name": item.display_name,
                    "status": item.status,
                    "exe path": item.exe_path,
                    "exe exists": item.exe_exists,
                    "launch_check supported": item.launch_check_supported,
                    "result": item.result,
                    "exit code": item.exit_code,
                    "elapsed seconds": f"{item.elapsed_seconds:.2f}",
                    "note": item.note,
                }
            )


def write_markdown(results: list[LaunchResult], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    checked = [item for item in results if item.result != "SKIPPED_STATUS"]
    counts = Counter(item.result for item in results)
    launch_check_missing = sum(1 for item in checked if item.exe_exists and not item.launch_check_supported)
    problem_results = [item for item in checked if item.result in {"TIMEOUT", "ERROR", "NO_EXE"}]

    lines = [
        "# DAKE exe Launch Check",
        "",
        "## Summary",
        "",
        f"- total apps: {len(results)}",
        f"- checked apps: {len(checked)}",
        f"- OK: {counts.get('OK', 0)}",
        f"- OK_GUI: {counts.get('OK_GUI', 0)}",
        f"- TIMEOUT: {counts.get('TIMEOUT', 0)}",
        f"- ERROR: {counts.get('ERROR', 0)}",
        f"- NO_EXE: {counts.get('NO_EXE', 0)}",
        f"- SKIPPED_STATUS: {counts.get('SKIPPED_STATUS', 0)}",
        f"- launch-check unsupported among checked exe apps: {launch_check_missing}",
        "",
        "## Problems",
        "",
    ]
    if problem_results:
        lines.extend([
            "| app | result | exit code | elapsed | note |",
            "| --- | --- | --- | --- | --- |",
        ])
        for item in problem_results:
            lines.append(
                f"| {item.app} | {item.result} | {item.exit_code} | {item.elapsed_seconds:.2f} | {item.note} |"
            )
    else:
        lines.append("- none")

    lines.extend([
        "",
        "## Checked",
        "",
        "| app | display_name | status | exe | launch-check | result | exit code | elapsed | note |",
        "| --- | --- | --- | --- | --- | --- | --- | --- | --- |",
    ])
    for item in checked:
        lines.append(
            f"| {item.app} | {item.display_name} | {item.status} | {item.exe_path} | "
            f"{item.launch_check_supported} | {item.result} | {item.exit_code} | "
            f"{item.elapsed_seconds:.2f} | {item.note} |"
        )

    skipped = [item for item in results if item.result == "SKIPPED_STATUS"]
    lines.extend([
        "",
        "## Skipped By Status",
        "",
        "| app | display_name | status | note |",
        "| --- | --- | --- | --- |",
    ])
    for item in skipped:
        lines.append(f"| {item.app} | {item.display_name} | {item.status} | {item.note} |")

    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Check DAKE dist exe launchability.")
    parser.add_argument("--only-available", action="store_true", help="check only status: available apps")
    parser.add_argument("--report-dir", default=str(REPORT_DIR), help="directory for reports")
    parser.add_argument("--timeout", type=float, default=5.0, help="launch-check timeout seconds")
    parser.add_argument("--gui-seconds", type=float, default=4.0, help="GUI smoke duration before stopping")
    args = parser.parse_args()

    results: list[LaunchResult] = []
    for app_dir in app_dirs():
        result = check_app(app_dir, args.only_available, args.timeout, args.gui_seconds)
        results.append(result)
        print(f"{result.app}: {result.result}", flush=True)
    report_dir = Path(args.report_dir)
    md_path = report_dir / "exe_launch_check.md"
    csv_path = report_dir / "exe_launch_check.csv"
    write_markdown(results, md_path)
    write_csv(results, csv_path)

    checked = [item for item in results if item.result != "SKIPPED_STATUS"]
    counts = Counter(item.result for item in results)
    launch_check_missing = sum(1 for item in checked if item.exe_exists and not item.launch_check_supported)

    print("DAKE exe Launch Check")
    print(f"total apps: {len(results)}")
    print(f"checked apps: {len(checked)}")
    for key in ("OK", "OK_GUI", "TIMEOUT", "ERROR", "NO_EXE", "SKIPPED_STATUS"):
        print(f"{key}: {counts.get(key, 0)}")
    print(f"launch-check unsupported among checked exe apps: {launch_check_missing}")
    print(f"report: {md_path}")
    print(f"csv: {csv_path}")

    problems = [item for item in checked if item.result in {"TIMEOUT", "ERROR", "NO_EXE"}]
    if problems:
        print("problems:")
        for item in problems:
            print(f"- {item.app}: {item.result} {item.note}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())