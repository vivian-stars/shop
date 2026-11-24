"""已解析的运费表。默认表放在内存里，上传的新表按编号保存。"""

from __future__ import annotations

from pathlib import Path
from uuid import uuid4

from shop.shopapp.models import RateBook
from shop.shopapp.platforms import get_platform
from shop.shopapp.platforms.aliexpress.sample_book import build_rate_book, build_templates

DATA = Path(__file__).resolve().parents[1] / "data"


class RateStore:
    def __init__(self) -> None:
        self.books: dict[str, tuple[str, RateBook]] = {}
        self.load_default()

    def load_default(self) -> str:
        path = DATA / "default_rates.xlsx"
        if not (DATA / "template_qty.xlsx").exists() or not (DATA / "template_weight.xlsx").exists():
            build_templates()
        if not path.exists():
            build_rate_book()
        payload = path.read_bytes()
        book = get_platform("aliexpress").read_book(payload, path.name)
        self.books["default"] = ("aliexpress", book)
        return "default"

    def add(self, platform_id: str, payload: bytes, filename: str) -> str:
        book = get_platform(platform_id).read_book(payload, filename)
        rate_id = uuid4().hex[:12]
        self.books[rate_id] = (platform_id, book)
        return rate_id

    def get(self, rate_id: str) -> tuple[str, RateBook]:
        try:
            return self.books[rate_id or "default"]
        except KeyError as exc:
            raise KeyError("找不到这份运费表，请重新上传") from exc
