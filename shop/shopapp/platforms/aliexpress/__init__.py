"""速卖通平台。"""

from __future__ import annotations

from shop.shopapp.models import Costs, ExportFile, Product, ProductQuote, RateBook
from shop.shopapp.platforms.aliexpress.calc import quote_products
from shop.shopapp.platforms.aliexpress.export import export_files
from shop.shopapp.platforms.aliexpress.parse import read_rate_book


class AliExpressPlatform:
    id = "aliexpress"
    name = "速卖通"

    def read_book(self, payload: bytes, source_name: str) -> RateBook:
        return read_rate_book(payload, source_name)

    def quote(
        self,
        book: RateBook,
        products: list[Product],
        channels: list[str],
        cargo: str,
        costs: Costs,
    ) -> list[ProductQuote]:
        missing = [name for name in channels if name not in book.channel_names()]
        if missing:
            raise KeyError("没有这些渠道：" + "、".join(missing))
        return quote_products(book, products, channels, cargo, costs)

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
    ) -> list[ExportFile]:
        missing = [name for name in channels if name not in book.channel_names()]
        if missing:
            raise KeyError("没有这些渠道：" + "、".join(missing))
        return export_files(
            book,
            products,
            channels,
            cargo,
            costs,
            mode,
            qty_start,
            qty_end,
            qty_step,
        )
