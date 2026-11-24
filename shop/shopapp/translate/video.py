"""听视频里的话，翻译后压成字幕，原来的声音保留。"""

from __future__ import annotations

import subprocess
import tempfile
import uuid
import wave
from dataclasses import dataclass
from pathlib import Path

from shop.shopapp.translate.engine import translate_text
from shop.shopapp.translate.languages import known_language, language_name

MAX_BYTES = 80 * 1024 * 1024
MAX_SECONDS = 180
RESULT_DIR = Path("/tmp/shop-translate")
_FONT_DIR = "/usr/share/fonts/truetype/wqy"
_SUFFIXES = {".mp4", ".mov", ".webm", ".mkv", ".m4v"}
_MODEL = None


@dataclass(frozen=True)
class Segment:
    start: float
    end: float
    text: str


def translate_upload(payload: bytes, filename: str, source: str, target: str) -> dict:
    if not payload:
        raise ValueError("请上传视频")
    if len(payload) > MAX_BYTES:
        raise ValueError("视频请小于 80MB")
    suffix = Path(filename or "video.mp4").suffix.lower()
    if suffix not in _SUFFIXES:
        raise ValueError("请上传 mp4、mov、webm 或 mkv 视频")
    work = Path(tempfile.mkdtemp(prefix="shop-video-"))
    src = work / f"source{suffix}"
    src.write_bytes(payload)
    if _duration(src) > MAX_SECONDS:
        raise ValueError("视频请短于 3 分钟")
    wav = work / "audio.wav"
    _extract_audio(src, wav)
    language = None if source == "auto" else source
    segments, detected = transcribe(str(wav), language)
    if not segments:
        raise ValueError("这段视频里没有听出说话内容")
    used = source if source != "auto" else known_language(detected)
    lines = [translate_text(segment.text, used, target)[0] for segment in segments]
    video_id = uuid.uuid4().hex
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    out = RESULT_DIR / f"{video_id}.mp4"
    srt = work / "sub.srt"
    srt.write_text(_srt(segments, lines), encoding="utf-8")
    _burn(src, srt, out)
    _trim_results()
    return {
        "source": used,
        "source_name": language_name(used),
        "target": target,
        "target_name": language_name(target),
        "transcript": "\n".join(segment.text for segment in segments),
        "translation": "\n".join(lines),
        "segments": [
            {
                "start": round(segment.start, 2),
                "end": round(segment.end, 2),
                "text": segment.text,
                "translation": line,
            }
            for segment, line in zip(segments, lines)
        ],
        "video_id": video_id,
    }


def result_path(video_id: str) -> Path | None:
    path = RESULT_DIR / f"{video_id}.mp4"
    if path.is_file():
        return path
    return None


def transcribe(path: str, language: str | None) -> tuple[list[Segment], str]:
    best: list[Segment] = []
    detected_best = ""
    for _ in range(3):
        rows, detected = _transcribe(path, language, vad_filter=True)
        if "\ufffd" in _join(rows) or _repetitive(rows):
            rows = []
        if _usable(rows, path):
            return rows, detected
        if _length(rows) > _length(best):
            best, detected_best = rows, detected
    return best, detected_best


def _join(rows: list[Segment]) -> str:
    return "\n".join(row.text for row in rows)


def _length(rows: list[Segment]) -> int:
    return sum(len(row.text) for row in rows)


def _usable(rows: list[Segment], path: str) -> bool:
    if not rows or _repetitive(rows):
        return False
    seconds = len(_read_wav(path)) / 16000
    if seconds < 2:
        return True
    return _length(rows) >= seconds * 4


def _repetitive(rows: list[Segment]) -> bool:
    return any(_repetitive_text(row.text) for row in rows)


def _repetitive_text(text: str) -> bool:
    compact = "".join(text.split())
    if len(compact) < 12:
        return False
    for size in range(1, 5):
        piece = compact[:size]
        if piece and compact.count(piece) >= 8:
            return True
    return False


def _transcribe(path: str, language: str | None, vad_filter: bool) -> tuple[list[Segment], str]:
    kwargs = {}
    if vad_filter:
        kwargs["vad_parameters"] = {"threshold": 0.35, "speech_pad_ms": 600, "min_silence_duration_ms": 200}
    segments, info = _model().transcribe(
        _read_wav(path),
        language=language,
        vad_filter=vad_filter,
        beam_size=5,
        temperature=0.0,
        condition_on_previous_text=False,
        **kwargs,
    )
    rows = []
    for item in segments:
        text = (item.text or "").strip()
        if text and not _repetitive_text(text):
            rows.append(Segment(float(item.start), float(item.end), text))
    detected = getattr(info, "language", None) or language or ""
    return rows, str(detected)


def _read_wav(path: str):
    import numpy as np

    with wave.open(path, "rb") as handle:
        frames = handle.readframes(handle.getnframes())
        width = handle.getsampwidth()
    if width != 2:
        raise ValueError("这段视频没有声音")
    return np.frombuffer(frames, dtype=np.int16).astype("float32") / 32768.0


def _model():
    global _MODEL
    if _MODEL is None:
        from faster_whisper import WhisperModel

        _MODEL = WhisperModel("small", device="cpu", compute_type="int8")
    return _MODEL


def _duration(path: Path) -> float:
    result = subprocess.run(
        [
            "ffprobe",
            "-v",
            "error",
            "-show_entries",
            "format=duration",
            "-of",
            "default=noprint_wrappers=1:nokey=1",
            str(path),
        ],
        capture_output=True,
        text=True,
    )
    try:
        return float(result.stdout.strip())
    except ValueError as exc:
        raise ValueError("这个视频没有读出来") from exc


def _extract_audio(src: Path, wav: Path) -> None:
    result = subprocess.run(
        ["ffmpeg", "-y", "-i", str(src), "-vn", "-ac", "1", "-ar", "16000", str(wav)],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0 or not wav.is_file() or wav.stat().st_size == 0:
        raise ValueError("这段视频没有声音")


def _burn(src: Path, srt: Path, out: Path) -> None:
    style = "FontName=WenQuanYi Micro Hei,FontSize=18,PrimaryColour=&H00FFFFFF,OutlineColour=&H00000000,BorderStyle=1,Outline=2,MarginV=28"
    subtitle = _filter_path(srt)
    result = subprocess.run(
        [
            "ffmpeg",
            "-y",
            "-i",
            str(src),
            "-vf",
            f"subtitles={subtitle}:fontsdir={_FONT_DIR}:force_style='{style}'",
            "-c:v",
            "libx264",
            "-preset",
            "veryfast",
            "-crf",
            "23",
            "-c:a",
            "aac",
            "-movflags",
            "+faststart",
            str(out),
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0 or not out.is_file():
        raise ValueError("译文字幕没有压进视频")


def _srt(segments: list[Segment], lines: list[str]) -> str:
    blocks = []
    for index, (segment, line) in enumerate(zip(segments, lines), start=1):
        start = segment.start
        end = segment.end if segment.end > start + 0.3 else start + 0.8
        blocks.append(f"{index}\n{_stamp(start)} --> {_stamp(end)}\n{line.strip()}\n")
    return "\n".join(blocks)


def _stamp(seconds: float) -> str:
    millis = max(0, int(round(seconds * 1000)))
    hours, millis = divmod(millis, 3_600_000)
    minutes, millis = divmod(millis, 60_000)
    secs, millis = divmod(millis, 1000)
    return f"{hours:02d}:{minutes:02d}:{secs:02d},{millis:03d}"


def _filter_path(path: Path) -> str:
    text = str(path).replace("\\", "\\\\").replace(":", "\\:").replace("'", "\\'")
    return f"'{text}'"


def _trim_results() -> None:
    files = sorted(RESULT_DIR.glob("*.mp4"), key=lambda item: item.stat().st_mtime, reverse=True)
    for old in files[20:]:
        old.unlink(missing_ok=True)
