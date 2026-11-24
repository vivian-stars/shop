"""平台插件接口。运费计算和导出都从这里进。"""

from __future__ import annotations

from typing import Protocol

from shop.shopapp.models import Costs, ExportFile, Product, ProductQuote, RateBook


class Platform(Protocol):
    id: str
    name: str

    def read_book(self, payload: bytes, source_name: str) -> RateBook: ...

    def quote(
        self,
        book: RateBook,
        products: list[Product],
        channels: list[str],
        cargo: str,
        costs: Costs,
    ) -> list[ProductQuote]: ...

    def export(
        self,
        book: RateBook,
        products: list[Product],
        channels: list[str],
        cargo: str,
        costs: Costs,
        mode: str,
        qty_start: int | None,
        qty_end: int | None,
        qty_step: int,
    ) -> list[ExportFile]: ...
