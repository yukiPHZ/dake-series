from __future__ import annotations

import json
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from core.cli_checker import run_command
from core.ollama_client import generate_ollama_text

LogCallback = Callable[[str], None]

SEQUENCE_TEXT = {
    "arranging": "補助脳：素材を並べています。",
    "flow": "補助脳：静かな流れを整えています。",
    "exported": "補助脳：横編集版を書き出しました。",
    "ffmpeg_missing": "補助脳：FFmpegが必要です。System Checkを確認してください。",
    "empty": "補助脳：Sequenceに動画がありません。",
    "normalize_failed": "補助脳：素材の整形に失敗しました。",
    "fallback": "補助脳：NVENCで試しましたが、CPUへ切り替えました。",
    "template_recommendation": "補助脳：先頭から終わりまで、静かな流れで並んでいます。",
}


def sequence_path(package_dir: Path) -> Path:
    return package_dir / "selected" / "sequence.json"


def selected_dir_for(package_dir: Path) -> Path:
    return package_dir / "selected"


def read_sequence(package_dir: Path) -> list[dict[str, Any]]:
    path = sequence_path(package_dir)
    if not path.exists() or not path.is_file():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8", errors="replace"))
    except Exception:
        return []
    if not isinstance(data, list):
        return []
    sequence: list[dict[str, Any]] = []
    for item in data:
        if not isinstance(item, dict):
            continue
        source = str(item.get("path") or "").strip()
        if not source:
            continue
        sequence.append(
            {
                "path": source,
                "duration": _float_value(item.get("duration")),
                "audio_present": bool(item.get("audio_present", True)),
            }
        )
    return sequence


def write_sequence(package_dir: Path, sequence: list[dict[str, Any]]) -> Path:
    selected_dir = selected_dir_for(package_dir)
    selected_dir.mkdir(parents=True, exist_ok=True)
    clean_items: list[dict[str, Any]] = []
    for item in sequence:
        source = str(item.get("path") or "").strip()
        if not source:
            continue
        clean_items.append(
            {
                "path": source,
                "duration": _float_value(item.get("duration")),
                "audio_present": bool(item.get("audio_present", True)),
            }
        )
    path = sequence_path(package_dir)
    path.write_text(json.dumps(clean_items, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def sequence_total_duration(sequence: list[dict[str, Any]]) -> float:
    return sum(_float_value(item.get("duration")) for item in sequence)


def build_sequence_recommendation(sequence: list[dict[str, Any]], ollama_ready: bool) -> dict[str, Any]:
    names = [Path(str(item.get("path") or "")).name for item in sequence if item.get("path")]
    fallback = _fallback_recommendation(names)
    if not ollama_ready or not names:
        return {"text": fallback, "used_ollama": False, "ollama_model": ""}

    prompt = (
        "You are the local assistant brain for Dakeユキズ稼働中.\n"
        "The user is arranging a quiet horizontal video sequence, not using a timeline editor.\n"
        "Return one short Japanese line only, starting with 補助脳：.\n"
        "Tone: calm, practical, quiet production flow. No hype.\n\n"
        "Sequence:\n"
        + "\n".join(f"{index + 1}. {name}" for index, name in enumerate(names[:12]))
    )
    result = generate_ollama_text(prompt, timeout=30)
    if result.get("ok"):
        return {
            "text": _one_line_recommendation(str(result.get("text") or fallback)),
            "used_ollama": True,
            "ollama_model": str(result.get("model") or ""),
        }
    return {"text": fallback, "used_ollama": False, "ollama_model": str(result.get("model") or "")}


def generate_horizontal_edit(
    package_dir: Path,
    sequence: list[dict[str, Any]] | None,
    ffmpeg_path: str | None,
    nvenc_online: bool,
    ollama_ready: bool,
    log: LogCallback | None = None,
) -> dict[str, Any]:
    selected_dir = selected_dir_for(package_dir)
    selected_dir.mkdir(parents=True, exist_ok=True)
    output_path = selected_dir / "horizontal_edit.mp4"
    sequence_items = sequence if sequence is not None else read_sequence(package_dir)
    write_sequence(package_dir, sequence_items)
    total_duration = sequence_total_duration(sequence_items)

    def emit(message: str) -> None:
        if log:
            log(message)

    emit(SEQUENCE_TEXT["arranging"])
    recommendation = build_sequence_recommendation(sequence_items, ollama_ready)
    if recommendation.get("text"):
        emit(str(recommendation["text"]))
    emit(SEQUENCE_TEXT["flow"])

    source_files = [Path(str(item.get("path") or "")) for item in sequence_items if item.get("path")]
    valid_sources = [path for path in source_files if path.exists() and path.is_file()]
    if not sequence_items or not valid_sources:
        _write_horizontal_log(selected_dir, sequence_items, "unavailable", False, False, output_path, SEQUENCE_TEXT["empty"])
        return _result(package_dir, selected_dir, output_path, "FAILED", "unavailable", False, False, recommendation, SEQUENCE_TEXT["empty"])
    if len(valid_sources) != len(source_files):
        missing = [str(path) for path in source_files if path not in valid_sources]
        error = "missing source files: " + ", ".join(missing[:6])
        _write_horizontal_log(selected_dir, sequence_items, "unavailable", False, False, output_path, error)
        return _result(package_dir, selected_dir, output_path, "FAILED", "unavailable", False, False, recommendation, error)
    if not ffmpeg_path:
        emit(SEQUENCE_TEXT["ffmpeg_missing"])
        _write_horizontal_log(selected_dir, sequence_items, "unavailable", False, False, output_path, "ffmpeg missing")
        return _result(package_dir, selected_dir, output_path, "FAILED", "unavailable", False, False, recommendation, "FFmpeg is required.")

    if output_path.exists():
        output_path.unlink()

    errors: list[str] = []
    try:
        with tempfile.TemporaryDirectory(prefix="_sequence_", dir=str(selected_dir)) as temp_dir:
            temp_root = Path(temp_dir)
            part_paths = _normalize_parts(ffmpeg_path, sequence_items, temp_root)
            concat_list = _write_concat_list(temp_root, part_paths)
            fallback = False

            if nvenc_online:
                completed = run_command(_concat_args(ffmpeg_path, concat_list, output_path, "h264_nvenc"), timeout=_timeout(total_duration))
                if completed.returncode == 0 and output_path.exists():
                    emit(SEQUENCE_TEXT["exported"])
                    _write_horizontal_log(selected_dir, sequence_items, "h264_nvenc", True, False, output_path, "")
                    return _result(package_dir, selected_dir, output_path, "COMPLETED", "h264_nvenc", True, False, recommendation, SEQUENCE_TEXT["exported"])
                errors.append((completed.stderr or completed.stdout or "h264_nvenc failed").strip()[-1200:])
                fallback = True
                emit(SEQUENCE_TEXT["fallback"])

            completed = run_command(_concat_args(ffmpeg_path, concat_list, output_path, "libx264"), timeout=_timeout(total_duration))
            if completed.returncode == 0 and output_path.exists():
                emit(SEQUENCE_TEXT["exported"])
                _write_horizontal_log(selected_dir, sequence_items, "libx264", False, fallback, output_path, "")
                return _result(package_dir, selected_dir, output_path, "COMPLETED", "libx264", False, fallback, recommendation, SEQUENCE_TEXT["exported"])

            errors.append((completed.stderr or completed.stdout or "libx264 failed").strip()[-1200:])
    except Exception as exc:
        errors.append(str(exc))

    error_text = "\n".join(error for error in errors if error) or "horizontal edit failed"
    _write_horizontal_log(selected_dir, sequence_items, "unavailable", False, any(errors), output_path, error_text)
    return _result(package_dir, selected_dir, output_path, "FAILED", "unavailable", False, any(errors), recommendation, error_text)


def _normalize_parts(ffmpeg_path: str, sequence: list[dict[str, Any]], temp_root: Path) -> list[Path]:
    parts: list[Path] = []
    for index, item in enumerate(sequence, start=1):
        source = Path(str(item.get("path") or ""))
        output = temp_root / f"part_{index:03d}.mp4"
        completed = run_command(_normalize_args(ffmpeg_path, source, output, bool(item.get("audio_present", True))), timeout=_timeout(_float_value(item.get("duration"))))
        if completed.returncode != 0 or not output.exists():
            message = (completed.stderr or completed.stdout or SEQUENCE_TEXT["normalize_failed"]).strip()[-1200:]
            raise RuntimeError(message)
        parts.append(output)
    return parts


def _normalize_args(ffmpeg_path: str, source: Path, output: Path, audio_present: bool) -> list[str]:
    video_filter = "scale=1920:1080:force_original_aspect_ratio=decrease,pad=1920:1080:(ow-iw)/2:(oh-ih)/2,setsar=1,format=yuv420p"
    args = [ffmpeg_path, "-y", "-i", str(source)]
    if not audio_present:
        args.extend(["-f", "lavfi", "-i", "anullsrc=channel_layout=stereo:sample_rate=48000"])
    args.extend(["-map", "0:v:0"])
    if audio_present:
        args.extend(["-map", "0:a:0?"])
    else:
        args.extend(["-map", "1:a:0"])
    args.extend(
        [
            "-vf",
            video_filter,
            "-r",
            "30",
            "-c:v",
            "libx264",
            "-preset",
            "veryfast",
            "-crf",
            "20",
            "-c:a",
            "aac",
            "-b:a",
            "192k",
            "-ar",
            "48000",
            "-ac",
            "2",
            "-shortest",
            str(output),
        ]
    )
    return args


def _concat_args(ffmpeg_path: str, concat_list: Path, output_path: Path, encoder: str) -> list[str]:
    video_codec_args = (
        ["-c:v", "h264_nvenc", "-preset", "p4", "-cq", "23"]
        if encoder == "h264_nvenc"
        else ["-c:v", "libx264", "-preset", "veryfast", "-crf", "22"]
    )
    return [
        ffmpeg_path,
        "-y",
        "-f",
        "concat",
        "-safe",
        "0",
        "-i",
        str(concat_list),
        *video_codec_args,
        "-c:a",
        "aac",
        "-b:a",
        "192k",
        "-movflags",
        "+faststart",
        str(output_path),
    ]


def _write_concat_list(temp_root: Path, part_paths: list[Path]) -> Path:
    path = temp_root / "concat_list.txt"
    lines = [f"file '{_concat_escape(part)}'" for part in part_paths]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def _write_horizontal_log(
    selected_dir: Path,
    sequence: list[dict[str, Any]],
    encoder: str,
    nvenc_used: bool,
    fallback: bool,
    output_path: Path,
    error: str,
) -> Path:
    lines = [
        f"executed_at: {datetime.now().isoformat(timespec='seconds')}",
        f"video_count: {len(sequence)}",
        f"total_duration: {sequence_total_duration(sequence):.3f}",
        f"encoder: {encoder}",
        f"nvenc_used: {str(nvenc_used).lower()}",
        f"fallback: {str(fallback).lower()}",
        "source_files:",
    ]
    lines.extend(f"- {item.get('path', '')}" for item in sequence)
    lines.extend([f"output_path: {output_path}", f"error: {error}"])
    path = selected_dir / "horizontal_edit_log.txt"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def _result(
    package_dir: Path,
    selected_dir: Path,
    output_path: Path,
    status: str,
    encoder: str,
    nvenc_used: bool,
    fallback: bool,
    recommendation: dict[str, Any],
    message: str,
) -> dict[str, Any]:
    return {
        "status": status,
        "package_dir": str(package_dir),
        "selected_dir": str(selected_dir),
        "output_path": str(output_path),
        "sequence_path": str(sequence_path(package_dir)),
        "log_path": str(selected_dir / "horizontal_edit_log.txt"),
        "encoder": encoder,
        "nvenc_used": nvenc_used,
        "fallback": fallback,
        "recommendation": recommendation.get("text", ""),
        "recommendation_used_ollama": bool(recommendation.get("used_ollama")),
        "recommendation_ollama_model": str(recommendation.get("ollama_model") or ""),
        "message": message,
    }


def _fallback_recommendation(names: list[str]) -> str:
    if len(names) >= 4:
        return f"補助脳：最後に {names[-1]} を置くと余熱が残ります。"
    if len(names) >= 3:
        return f"補助脳：{names[0]} → {names[1]} → {names[2]} の流れが自然です。"
    if len(names) >= 2:
        return f"補助脳：{names[0]} から {names[1]} へ静かにつながります。"
    if len(names) == 1:
        return f"補助脳：{names[0]} を軸に、短く出せる形です。"
    return SEQUENCE_TEXT["template_recommendation"]


def _one_line_recommendation(text: str) -> str:
    line = next((raw.strip() for raw in text.splitlines() if raw.strip()), "")
    line = line.strip().strip('"').strip("'")
    if not line:
        return SEQUENCE_TEXT["template_recommendation"]
    if not line.startswith("補助脳："):
        line = f"補助脳：{line}"
    return line[:180]


def _concat_escape(path: Path) -> str:
    return str(path.resolve()).replace("\\", "/").replace("'", "'\\''")


def _float_value(value: object) -> float:
    try:
        return max(0.0, float(value))  # type: ignore[arg-type]
    except Exception:
        return 0.0


def _timeout(duration: float) -> int:
    return max(120, min(3600, int(duration * 8) + 120))
