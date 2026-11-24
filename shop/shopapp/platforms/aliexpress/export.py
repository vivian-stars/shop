"""按速卖通模板写出按数量或按重量的运费表。"""

from __future__ import annotations

import io
import json
import re
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook

from shop.shopapp.models import Costs, DestinationQuote, ExportFile, Product, ProductQuote
from shop.shopapp.money import money
from shop.shopapp.platforms.aliexpress.calc import quote_products

DATA = Path(__file__).resolve().parents[3] / "data"
SELLER_RULES = {"卖家承担", "不发货"}
ALIASES = {
    "俄罗斯联邦": "俄罗斯",
    "南韩": "韩国",
    "阿拉伯联合酋长国": "阿联酋",
    "中国香港特别行政区": "香港",
    "香港特别行政区": "香港",
    "台湾省": "台湾",
}


def export_files(
    book,
    products: list[Product],
    channels: list[str],
    cargo: str,
    costs: Costs,
    mode: str,
    qty_start: int | None,
    qty_end: int | None,
    qty_step: int,
) -> list[ExportFile]:
    if mode == "quantity" and (qty_start is None or qty_end is None):
        raise ValueError("按数量计费需要填写起始采购数和截止采购数")
    if mode == "quantity" and qty_start > qty_end:
        raise ValueError("起始采购数不能大于截止采购数")
    quotes = quote_products(book, products, channels, cargo, costs)
    codes = json.loads((DATA / "line_codes.json").read_text(encoding="utf-8"))
    template = DATA / ("template_qty.xlsx" if mode == "quantity" else "template_weight.xlsx")
    files: list[ExportFile] = []
    for product_quote in quotes:
        for channel in channels:
            content, warnings = _fill(
                template,
                channel,
                codes.get(channel, ""),
                product_quote,
                mode,
                qty_start,
                qty_end,
                qty_step,
            )
            files.append(
                ExportFile(
                    filename=_filename(channel, product_quote.product, mode),
                    content=content,
                    warnings=warnings,
                )
            )
    return files


def _fill(template, channel, line_code, product_quote: ProductQuote, mode, qty_start, qty_end, qty_step):
    workbook = load_workbook(template)
    route = workbook["线路编码"]
    warnings: list[str] = []
    if channel != "无忧物流-标准":
        route["A1"] = channel
        route["C1"] = f"线路代号: {line_code}" if line_code else "线路代号: "
        if "service_name" in workbook.sheetnames:
            workbook["service_name"]["A1"] = line_code
            workbook["service_name"]["B1"] = channel
        if not line_code:
            warnings.append(f"{channel} 缺少线路代号，导入平台前需要补上")
    by_country = {_norm(row.country): row for row in product_quote.rows if row.channel == channel}
    for cells in route.iter_rows(min_row=3, max_col=19):
        name = cells[0].value
        if not name:
            continue
        rule = str(cells[4].value or "")
        if rule in SELLER_RULES:
            continue
        if mode == "quantity" and rule != "按数量计费":
            continue
        if mode == "weight" and rule != "按重量计费":
            continue
        quote = by_country.get(_norm(str(name)))
        if quote is None or not quote.shippable or quote.currency != "CNY":
            if quote is not None and quote.currency == "USD":
                warnings.append(f"{channel} / {name} 是美元报价，没有写入人民币模板")
            continue
        if mode == "quantity":
            cells[8].value = qty_start
            cells[9].value = qty_end
            cells[10].value = float(quote.final_freight)
            cells[11].value = qty_step
            cells[12].value = float(_extra(quote, qty_step))
        else:
            _write_weight(cells, quote)
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue(), _unique(warnings)


def _write_weight(cells, quote: DestinationQuote) -> None:
    if quote.mode == "first":
        first_kg = (quote.first_g or Decimal("500")) / Decimal("1000")
        step_kg = (quote.additional_g or Decimal("500")) / Decimal("1000")
        limit_kg = (quote.limit_g or Decimal("30000")) / Decimal("1000")
        cells[8].value = float(first_kg)
        cells[9].value = float(money(quote.first_fee or 0))
        cells[10].value = float(limit_kg)
        cells[11].value = float(step_kg)
        cells[12].value = float(money(quote.additional_fee or 0))
        return
    step_g = quote.step_g or Decimal("1")
    if step_g < 1:
        step_g = Decimal("1")
    step_kg = step_g / Decimal("1000")
    end_kg = (quote.band_end_g or quote.limit_g or Decimal("2000")) / Decimal("1000")
    cells[8].value = float(step_kg)
    cells[9].value = float(money(quote.registration or 0))
    cells[10].value = float(end_kg)
    cells[11].value = float(step_kg)
    per_step = (quote.rate_per_kg or Decimal("0")) * step_kg
    cells[12].value = float(money(per_step))


def _extra(quote: DestinationQuote, qty_step: int) -> Decimal:
    if quote.extra_freight is None:
        return Decimal("0")
    return money(quote.extra_freight * Decimal(qty_step))


def _filename(channel: str, product: Product, mode: str) -> str:
    label = "按数量计费" if mode == "quantity" else "按重量计费"
    weight = int(product.weight_g) if product.weight_g == product.weight_g.to_integral_value() else product.weight_g
    sku = re.sub(r"[^\w\u4e00-\u9fff.-]+", "_", product.sku).strip("_") or "SKU"
    channel_name = re.sub(r"[\\/:*?\"<>|]+", "_", channel)
    return f"{channel_name}_{sku}_{weight}g_{label}.xlsx"


def _norm(name: str) -> str:
    text = str(name).strip().replace(" ", "").replace("\u00a0", "")
    return ALIASES.get(text, text)


def _unique(items: list[str]) -> list[str]:
    seen = []
    for item in items:
        if item not in seen:
            seen.append(item)
    return seen
