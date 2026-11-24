"""翻译可选的语言。输入可以自动识别，输出必须指定。"""

from __future__ import annotations

LANGUAGES: tuple[tuple[str, str], ...] = (
    ("zh", "中文"),
    ("en", "英语"),
    ("ru", "俄语"),
    ("es", "西班牙语"),
    ("fr", "法语"),
    ("pt", "葡萄牙语"),
    ("ar", "阿拉伯语"),
    ("de", "德语"),
    ("ja", "日语"),
    ("ko", "韩语"),
    ("th", "泰语"),
    ("vi", "越南语"),
    ("id", "印尼语"),
    ("ms", "马来语"),
    ("tr", "土耳其语"),
    ("it", "意大利语"),
    ("pl", "波兰语"),
    ("nl", "荷兰语"),
    ("uk", "乌克兰语"),
)

_NAMES = dict(LANGUAGES)


def language_name(code: str) -> str:
    if code == "auto":
        return "自动识别"
    return _NAMES.get(code, code)


def source_options() -> list[dict]:
    return [{"id": "auto", "name": "自动识别"}] + target_options()


def target_options() -> list[dict]:
    return [{"id": code, "name": name} for code, name in LANGUAGES]


def require_source(code: str) -> str:
    if code == "auto" or code in _NAMES:
        return code
    raise ValueError("没有这个输入语言")


def require_target(code: str) -> str:
    if code in _NAMES:
        return code
    raise ValueError("没有这个输出语言")


def known_language(code: str) -> str:
    primary = (code or "").replace("_", "-").split("-")[0].lower()
    if primary in _NAMES:
        return primary
    return "auto"
