"""把文字交给翻译接口。长文按行再按句切开，避免一次请求过长。"""

from __future__ import annotations

import json
import subprocess

import httpx

_URL = "https://translate.googleapis.com/translate_a/single"
_MEMORY = "https://api.mymemory.translated.net/get"
_LIMIT = 400
_MEMORY_CODES = {
    "zh": "zh-CN",
    "en": "en",
    "ru": "ru",
    "es": "es",
    "fr": "fr",
    "pt": "pt",
    "ar": "ar",
    "de": "de",
    "ja": "ja",
    "ko": "ko",
    "th": "th",
    "vi": "vi",
    "id": "id",
    "ms": "ms",
    "tr": "tr",
    "it": "it",
    "pl": "pl",
    "nl": "nl",
    "uk": "uk",
}


def translate_text(text: str, source: str, target: str) -> tuple[str, str]:
    detected = source
    lines = text.split("\n")
    translated_lines: list[str] = []
    for line in lines:
        if line.strip() == "":
            translated_lines.append("")
            continue
        pieces: list[str] = []
        for chunk in _chunks(line):
            piece, found = _google(chunk, "auto" if source == "auto" else source, target)
            pieces.append(piece)
            if source == "auto":
                detected = _normalize(found)
                source = detected
        translated_lines.append("".join(pieces))
    if detected == "auto":
        detected = target
    return "\n".join(translated_lines), detected


def read_google_payload(payload) -> tuple[str, str]:
    chunks = payload[0] if isinstance(payload, list) and payload else []
    text = "".join(item[0] for item in chunks if item and item[0])
    detected = ""
    if isinstance(payload, list) and len(payload) > 2 and payload[2]:
        detected = _normalize(str(payload[2]))
    return text, detected


def _google(text: str, source: str, target: str) -> tuple[str, str]:
    try:
        return _google_curl(text, source, target)
    except ValueError:
        if source == "auto":
            raise
        return _mymemory(text, source, target)


def _google_curl(text: str, source: str, target: str) -> tuple[str, str]:
    try:
        result = subprocess.run(
            [
                "curl",
                "-sS",
                "-m",
                "20",
                "-G",
                _URL,
                "--data-urlencode",
                "client=gtx",
                "--data-urlencode",
                f"sl={source}",
                "--data-urlencode",
                f"tl={target}",
                "--data-urlencode",
                "dt=t",
                "--data-urlencode",
                f"q={text}",
            ],
            capture_output=True,
            text=True,
            check=False,
        )
        payload = json.loads(result.stdout)
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError("翻译服务暂时没有连上") from exc
    translated, detected = read_google_payload(payload)
    if result.returncode != 0 or not translated:
        raise ValueError("翻译服务暂时没有连上")
    return translated, detected or source


def _mymemory(text: str, source: str, target: str) -> tuple[str, str]:
    pair = f"{_MEMORY_CODES.get(source, source)}|{_MEMORY_CODES.get(target, target)}"
    try:
        response = httpx.get(_MEMORY, params={"q": text, "langpair": pair}, timeout=20)
        response.raise_for_status()
        translated = response.json()["responseData"]["translatedText"]
    except (httpx.HTTPError, ValueError, KeyError, TypeError) as exc:
        raise ValueError("翻译服务暂时没有连上") from exc
    if not translated or str(translated).startswith("MYMEMORY WARNING"):
        raise ValueError("翻译服务暂时没有连上")
    return str(translated), source


def _normalize(code: str) -> str:
    primary = (code or "").replace("_", "-").split("-")[0].lower()
    if primary in {"zh", "cmn"}:
        return "zh"
    return primary or "auto"


def _chunks(line: str) -> list[str]:
    if len(line) <= _LIMIT:
        return [line]
    pieces: list[str] = []
    buf = ""
    for char in line:
        buf += char
        if len(buf) >= _LIMIT and char in "。！？!?.;；，,":
            pieces.append(buf)
            buf = ""
    if buf:
        pieces.append(buf)
    return pieces
