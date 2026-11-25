"""按包邮基准表计算。实际运费之上加佣金，再和基准比出区域调价。"""

from __future__ import annotations

import json
from dataclasses import dataclass
from decimal import Decimal, ROUND_FLOOR
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from shop.shopapp.models import ProductQuote
from shop.shopapp.money import D, money
from shop.shopapp.store import DATA

_PRICING = json.loads((DATA / "regional_pricing.json").read_text(encoding="utf-8"))
REGIONAL_NAME: dict[str, str] = _PRICING["regional_name"]
TEMPLATE_NAME: dict[str, str] = _PRICING["template_name"]
DISCOUNT_COUNTRIES: list[str] = _PRICING["discount_countries"]
DISCOUNTS = (Decimal("0.20"), Decimal("0.25"), Decimal("0.30"), Decimal("0.35"))
MONEY_FORMAT = '"￥"#,##0.00;"￥"\\-#,##0.00'


@dataclass(frozen=True)
class BenchmarkChoice:
    amount: Decimal
    country: str
    freight: Decimal | None


@dataclass(frozen=True)
class BenchmarkRow:
    country: str
    channel: str
    template_name: str
    regional_name: str
    rate_per_kg: Decimal | None
    registration: Decimal | None
    freight: Decimal | None
    commission: Decimal | None
    total: Decimal | None
    note: str
    included: Decimal | None
    difference: Decimal | None
    regional_price: Decimal | None
    regional_set: str
    regional_amount: Decimal | None
    template_fee: Decimal | None
    extra_each: Decimal | None
    extra_diff: Decimal | None
    extra_calc: Decimal | None
    shippable: bool


def infer_benchmark(rows, commission_rate: Decimal) -> BenchmarkChoice:
    """包邮基准取买家包邮国家实际运费降序后的正中间，再去掉小数。

    买家包邮包括成本不超过基准的国家，以及区域调价国家。基准和这组国家互相决定，
    取能对上的那个整数。45 个国家时正中间是第 23 个；0.3 千克普货、佣金 20% 时是比利时。
    """
    prepared = [_prepared_freight(item, commission_rate) for item in rows if item.currency != "USD"]
    prepared = [item for item in prepared if item is not None]
    if not prepared:
        return BenchmarkChoice(Decimal("0"), "", None)
    candidates = {Decimal("0")}
    for freight, total, _regional, _name in prepared:
        candidates.add(_floor_yuan(freight))
        candidates.add(_floor_yuan(total))
    fixed = []
    for amount in sorted(candidates):
        choice = _middle_choice(prepared, amount)
        if choice is not None and choice.amount == amount:
            fixed.append(choice)
    if len(fixed) == 1:
        return fixed[0]
    reached = _walk_benchmark(prepared, Decimal("0"))
    for choice in fixed:
        if choice.amount == reached.amount:
            return choice
    if fixed:
        return fixed[0]
    return reached


def benchmark_rows(quote: ProductQuote, benchmark: Decimal, commission_rate: Decimal) -> list[BenchmarkRow]:
    rows = [_row(item, benchmark, commission_rate) for item in quote.rows if item.currency != "USD"]
    rows.sort(key=lambda item: (item.total is None, item.total or Decimal("0"), item.template_name))
    return rows


def discount_rows(rows: list[BenchmarkRow]) -> list[dict]:
    by_regional = {item.regional_name: item for item in rows if item.regional_name}
    built = []
    for name in DISCOUNT_COUNTRIES:
        source = by_regional.get(name)
        price = source.regional_price if source is not None else None
        built.append({"country": name, "price": price, "bands": _discount_bands(price)})
    return built


def workbook_bytes(quotes: list[ProductQuote], commission_rate: Decimal) -> bytes:
    book = Workbook()
    default = book.active
    used: set[str] = set()
    written = False
    single = len(quotes) == 1 and len({row.channel for row in quotes[0].rows}) <= 1
    for quote in quotes:
        channels: list[str] = []
        for row in quote.rows:
            if row.channel not in channels:
                channels.append(row.channel)
        for channel in channels:
            subset = ProductQuote(
                quote.product,
                [row for row in quote.rows if row.channel == channel],
                quote.warnings,
            )
            rate = subset.product.costs.commission_rate() if subset.product.costs is not None else commission_rate
            choice = infer_benchmark(subset.rows, rate)
            rows = benchmark_rows(subset, choice.amount, rate)
            title = channel if single else f"{quote.product.sku}-{channel}"
            sheet = book.create_sheet(_unique_title(used, title))
            _write_main(sheet, rows, quote.product.weight_g)
            discount_title = "区域定价国家调整金额计算" if single else f"区域调价-{title}"
            discount = book.create_sheet(_unique_title(used, discount_title))
            _write_discount(discount, discount_rows(rows))
            written = True
    if written:
        book.remove(default)
    buffer = BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def _prepared_freight(item, commission_rate: Decimal):
    if not item.shippable or item.rate_per_kg is None or item.billable_g is None:
        return None
    freight = item.billable_g / Decimal("1000") * item.rate_per_kg + (item.registration or Decimal("0"))
    total = freight * (Decimal("1") + commission_rate)
    regional = bool(REGIONAL_NAME.get(item.country, ""))
    return freight, total, regional, TEMPLATE_NAME.get(item.country, item.country)


def _middle_choice(prepared, benchmark: Decimal) -> BenchmarkChoice | None:
    chosen = [item for item in prepared if item[2] or item[1] <= benchmark]
    if not chosen:
        return None
    chosen.sort(key=lambda item: (-item[0], item[3]))
    freight, _total, _regional, name = chosen[len(chosen) // 2]
    return BenchmarkChoice(_floor_yuan(freight), name, freight)


def _walk_benchmark(prepared, start: Decimal) -> BenchmarkChoice:
    amount = start
    seen: set[Decimal] = set()
    last = BenchmarkChoice(Decimal("0"), "", None)
    for _ in range(len(prepared) + 2):
        choice = _middle_choice(prepared, amount)
        if choice is None:
            if amount == 0 and prepared:
                amount = max(_floor_yuan(item[0]) for item in prepared)
                continue
            return last
        last = choice
        if choice.amount == amount or choice.amount in seen:
            return choice
        seen.add(amount)
        amount = choice.amount
    return last


def _floor_yuan(value: Decimal) -> Decimal:
    return value.to_integral_value(rounding=ROUND_FLOOR)


def _row(item, benchmark: Decimal, commission_rate: Decimal) -> BenchmarkRow:
    template = TEMPLATE_NAME.get(item.country, item.country)
    regional = REGIONAL_NAME.get(item.country, "")
    if not item.shippable or item.rate_per_kg is None or item.billable_g is None:
        return BenchmarkRow(
            item.country, item.channel, template, regional, item.rate_per_kg, item.registration,
            None, None, None, "不可运", None, None, None, "否", None, None, None, None, None, False,
        )
    kilograms = item.billable_g / Decimal("1000")
    registration = item.registration or Decimal("0")
    freight = kilograms * item.rate_per_kg + registration
    commission = freight * commission_rate
    total = freight + commission
    difference = total - benchmark
    if total <= benchmark:
        note = "包邮国家"
        regional_price = difference - difference * commission_rate
        fee = each = gap = calc = None
    elif regional:
        note = "额外设置包邮国家"
        regional_price = difference + difference * commission_rate
        fee = each = gap = calc = None
    else:
        note = "不包邮国家"
        regional_price = None
        fee = difference + difference * commission_rate
        each = kilograms * item.rate_per_kg
        gap = each - benchmark
        calc = gap + gap * commission_rate
    amount = regional_price
    return BenchmarkRow(
        item.country,
        item.channel,
        template,
        regional,
        item.rate_per_kg,
        registration,
        freight,
        commission,
        total,
        note,
        benchmark,
        difference,
        regional_price,
        "是" if regional else "否",
        amount,
        fee,
        each,
        gap,
        calc,
        True,
    )


def _discount_bands(price: Decimal | None) -> list[dict]:
    if price is None:
        return [{"mark": None, "amount": None} for _ in DISCOUNTS]
    if price < 0:
        absolute = abs(price)
        return [{"mark": "-", "amount": absolute} for _ in DISCOUNTS]
    return [
        {"mark": None, "amount": money(price / (Decimal("1") - rate))}
        for rate in DISCOUNTS
    ]


def _write_main(sheet, rows: list[BenchmarkRow], weight_g: Decimal) -> None:
    kilograms = _kg_text(weight_g)
    headers = [
        "运费模板国家名称（和运费报价国家名称有不同）",
        "区域调价国家名称",
        "配送服务费\n元（RMB）/KG\n*每1g计重",
        "挂号服务费\n元（RMB）/包裹",
        f"重量={kilograms}KG\n实际运费",
        "平台佣金+联盟佣金+金币佣金+一站推费用",
        "运费成本合计",
        "备注",
        "产品中的运费",
        "差异运费",
        "区域调价价格计算",
        "是否设置区域调价",
        "区域调价金额",
        "运费模板费用计算",
        "每多买一个运费",
        "续重运费差异",
        "续重运费计算",
        "续重运费设置",
    ]
    sheet.append(headers)
    for cell in sheet[1]:
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.font = Font(bold=True)
    sheet.row_dimensions[1].height = 48
    for item in rows:
        sheet.append([
            item.template_name,
            item.regional_name or None,
            _num(item.rate_per_kg),
            _num(item.registration),
            _num(item.freight),
            _num(item.commission),
            _num(item.total),
            item.note,
            _num(item.included),
            _num(item.difference),
            _num(item.regional_price),
            item.regional_set,
            _num(item.regional_amount),
            _num(item.template_fee),
            _num(item.extra_each),
            _num(item.extra_diff),
            _num(item.extra_calc),
            None,
        ])
    for row in sheet.iter_rows(min_row=2, min_col=3, max_col=17):
        for cell in row:
            if isinstance(cell.value, Decimal):
                cell.value = float(cell.value)
            if isinstance(cell.value, (int, float)):
                cell.number_format = MONEY_FORMAT
    for index, width in enumerate([28, 18, 16, 16, 14, 22, 14, 18, 14, 12, 16, 16, 14, 16, 14, 14, 14, 14], start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.auto_filter.ref = sheet.dimensions
    sheet.freeze_panes = "A2"


def _write_discount(sheet, rows: list[dict]) -> None:
    sheet.append(["国家", "原价", None, "20%折扣", None, "25%折扣", None, "30%折扣", None, "35%折扣"])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for item in rows:
        values = [item["country"], _num(item["price"])]
        for band in item["bands"]:
            values.extend([band["mark"], _num(band["amount"])])
        sheet.append(values)
    for row in sheet.iter_rows(min_row=2, min_col=2, max_col=10):
        for cell in row:
            if isinstance(cell.value, float):
                cell.number_format = "0.00"
    sheet.column_dimensions["A"].width = 24
    sheet.freeze_panes = "A2"


def _kg_text(weight_g: Decimal) -> str:
    kilograms = (D(weight_g) / Decimal("1000")).normalize()
    return format(kilograms, "f")


def _num(value: Decimal | None):
    if value is None:
        return None
    return D(value)


def _unique_title(used: set[str], name: str) -> str:
    cleaned = name.replace("/", " ").replace("\\", " ").replace("*", " ").replace("?", " ").replace("[", " ").replace("]", " ")
    cleaned = cleaned[:31] or "运费"
    candidate = cleaned
    number = 2
    while candidate in used:
        suffix = f"-{number}"
        candidate = f"{cleaned[:31 - len(suffix)]}{suffix}"
        number += 1
    used.add(candidate)
    return candidate
