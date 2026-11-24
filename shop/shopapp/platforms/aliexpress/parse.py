"""按表头读取物流费工作簿。

每个渠道是一张表。「目录」不是渠道。同一张表里可以有多段报价，
用「运达国家」或「目的国」开头，合并单元格把普货、非普货、大包和重量段对齐。
"""

from __future__ import annotations

import io
import re
from decimal import Decimal

from openpyxl import load_workbook

from shop.shopapp.models import Channel, CountryRate, PriceBand, RateBook, Rules
from shop.shopapp.money import D

UNAVAILABLE = {"暂无服务", "-", "—", "--", "/", "无", "n/a", "N/A", "暂无"}
HEADER_KEYS = ("运达国家", "收件地中文名", "目的国")


def read_rate_book(payload: bytes, source_name: str) -> RateBook:
    workbook = load_workbook(io.BytesIO(payload), data_only=True)
    channels: list[Channel] = []
    for sheet in workbook.worksheets:
        if sheet.title.strip() == "目录":
            continue
        channel = parse_sheet(sheet)
        if channel.countries:
            channels.append(channel)
    if not channels:
        raise ValueError("这份运费表里没有读到渠道报价")
    return RateBook(channels=channels, source_name=source_name)


def parse_sheet(sheet) -> Channel:
    grid, max_row, max_col = _grid(sheet)
    header_rows = _collapse_header_rows(
        [row for row in range(1, max_row + 1) if _is_header_row(grid, row, max_col)]
    )
    preamble = _note_text(grid, 1, header_rows[0] - 1 if header_rows else 0, max_col)
    channel = Channel(name=sheet.title, rules=parse_rules(preamble))
    channel.warning = channel.rules.warning
    previous = 0
    for index, header_row in enumerate(header_rows):
        next_header = header_rows[index + 1] if index + 1 < len(header_rows) else max_row + 1
        local_notes = _note_text(grid, previous + 1, header_row - 1, max_col)
        rules = channel.rules.merge(parse_rules(local_notes))
        if _is_row_band_header(grid, header_row, max_col):
            last = _parse_row_bands(grid, header_row, next_header, max_col, rules, channel)
        else:
            block_end = _header_block_end(grid, header_row, next_header, max_col)
            last = _parse_matrix(grid, header_row, next_header, max_col, rules, channel, block_end)
        trailing = parse_rules(_note_text(grid, last + 1, next_header - 1, max_col))
        if trailing.vol_divisor or trailing.country_min_g or trailing.warning:
            channel.rules = channel.rules.merge(trailing)
        previous = last
    _assign_size_mode(channel)
    channel.currencies = {
        band.currency
        for country in channel.countries.values()
        for band in country.bands
    }
    return channel


def parse_rules(text: str) -> Rules:
    rules = Rules()
    if not text:
        return rules
    compact = re.sub(r"\s+", "", text)
    divisor = re.search(r"/(\d{4})", compact)
    if divisor and ("体积" in compact or "计泡" in compact):
        rules.vol_divisor = Decimal(divisor.group(1))
        rules.vol_scope = "large" if "大包" in compact else "all"
    side = re.search(r"单边(?:长度)?(?:不超过|<=|＜=|≤)(\d+(?:\.\d+)?)", compact)
    total = re.search(r"(?:长\+宽\+高|三边(?:长度)?之和)(?:不超过|<=|＜=|≤)(\d+(?:\.\d+)?)", compact)
    small_weight = re.search(r"(?:实际重量|申报重量)?[<＜≤]=?\s*2\s*kg", compact, re.I)
    if "＜=2kg" in compact or "<=2kg" in compact or "≤2kg" in compact:
        small_weight = True  # type: ignore[assignment]
    hard_weight = re.search(r"最大支持到(\d+(?:\.\d+)?)kg", compact, re.I)
    if side:
        rules.switch_side_cm = D(side.group(1))
    if total:
        rules.switch_sum_cm = D(total.group(1))
    if small_weight:
        rules.switch_weight_g = Decimal("2000")
    if hard_weight:
        rules.hard_weight_g = D(hard_weight.group(1)) * Decimal("1000")
    for match in re.finditer(r"([\u4e00-\u9fff、]{2,40})最低计费重量(\d+(?:\.\d+)?)kg", compact, re.I):
        kilos = D(match.group(2)) * Decimal("1000")
        for name in match.group(1).split("、"):
            if name:
                rules.country_min_g[name] = kilos
    if "燃油" in compact and any(word in compact for word in ("不含", "未含", "另计", "另收")):
        rules.warning = "报价不含燃油附加费"
    return rules


def _assign_size_mode(channel: Channel) -> None:
    has_large = any(
        band.category == "large"
        for country in channel.countries.values()
        for band in country.bands
    )
    rules = channel.rules
    if has_large:
        return
    rules.hard_side_cm = rules.switch_side_cm
    rules.hard_sum_cm = rules.switch_sum_cm
    if rules.hard_weight_g is None:
        rules.hard_weight_g = rules.switch_weight_g
    rules.switch_side_cm = None
    rules.switch_sum_cm = None
    rules.switch_weight_g = None


def _parse_matrix(grid, header_row, next_header, max_col, rules: Rules, channel: Channel, block_end: int) -> int:
    fee_row = _fee_row(grid, header_row, block_end, max_col)
    columns = _column_specs(grid, header_row, fee_row, max_col, block_end)
    name_col, code_col = _identity_columns(grid, header_row, max_col)
    min_col = next((col.index for col in columns if col.kind == "min_weight"), None)
    last = block_end
    blanks = 0
    for row in range(block_end + 1, next_header):
        name = _text(grid.get((row, name_col)))
        if _is_header_text(name) or name.endswith("列表"):
            break
        prices = _row_has_price(grid, row, columns)
        if not name and not prices:
            blanks += 1
            if blanks >= 2:
                break
            continue
        blanks = 0
        last = row
        if not name or _is_header_text(name):
            continue
        country = _country(channel, name, _text(grid.get((row, code_col))) if code_col else "")
        if min_col:
            country.min_g = _weight_text(grid.get((row, min_col))) or country.min_g
        _fill_country_bands(country, grid, row, columns, rules)
    _separate_touching_bands(channel)
    return last


def _parse_row_bands(grid, header_row, next_header, max_col, rules: Rules, channel: Channel) -> int:
    headers = {col: _text(grid.get((header_row, col))) for col in range(1, max_col + 1)}
    country_col = _find_header(headers, "目的国", "运达国家")
    code_col = _find_header(headers, "code", "Code")
    start_col = _find_header(headers, "起始重量", "克重段")
    end_col = _find_header(headers, "终止重量")
    step_col = _find_header(headers, "进位", "重量阶梯")
    rate_col = _find_rate_column(headers)
    reg_col = _find_header(headers, "挂号")
    bubble_col = _find_header(headers, "计泡系数")
    bubble_flag = _find_header(headers, "是否计泡")
    currency_col = _find_header(headers, "币种")
    unit_kg = "kg" in headers.get(start_col, "").lower()
    last_name = ""
    last_code = ""
    last = header_row
    for row in range(header_row + 1, next_header):
        label = _text(grid.get((row, 1)))
        if label.startswith("附加") or label.startswith("计费重量"):
            break
        name = _text(grid.get((row, country_col))) if country_col else ""
        if name:
            last_name = name
            last_code = _text(grid.get((row, code_col))) if code_col else last_code
        elif not last_name:
            continue
        start_cell = grid.get((row, start_col)) if start_col else None
        if start_cell in (None, "") and grid.get((row, rate_col)) in (None, ""):
            continue
        band_text = _text(start_cell)
        parsed = _parse_band(band_text, unit_kg=unit_kg or "kg" in band_text.lower())
        if parsed is None and start_col and end_col:
            start_g = _grams(grid.get((row, start_col)), False)
            end_raw = grid.get((row, end_col))
            end_g = None if _text(end_raw) in {"", "-", "—"} else _grams(end_raw, False)
            if start_g is None:
                continue
            parsed = (start_g, end_g, True, True)
        if parsed is None:
            continue
        rate = _price(grid.get((row, rate_col))) if rate_col else None
        reg = _price(grid.get((row, reg_col))) if reg_col else None
        if rate is None and reg is None:
            continue
        step = Decimal("1000") if unit_kg else Decimal("1")
        if step_col:
            step = _grams(grid.get((row, step_col)), False) or _step(_text(grid.get((row, step_col)))) or step
        divisor = rules.vol_divisor
        if bubble_flag and _text(grid.get((row, bubble_flag))) not in {"", "是"}:
            divisor = None
        if bubble_col:
            coefficient = _price(grid.get((row, bubble_col)))
            if isinstance(coefficient, Decimal):
                divisor = coefficient
        local = Rules(
            vol_divisor=divisor,
            vol_scope="all" if divisor else "none",
            country_min_g=dict(rules.country_min_g),
        )
        currency = "CNY"
        if currency_col and "美元" in _text(grid.get((row, currency_col))):
            currency = "USD"
        country = _country(channel, last_name, last_code)
        start_g, end_g, start_in, end_in = parsed
        country.bands.append(
            PriceBand(
                category="any",
                start_g=start_g,
                end_g=end_g,
                start_inclusive=start_in,
                end_inclusive=end_in,
                label=band_text,
                rate_per_kg=rate if rate != "bad" else None,
                registration=Decimal("0") if reg in (None, "bad") else reg,
                step_g=step,
                currency=currency,
                unavailable=rate == "bad",
                rules=rules.merge(local),
            )
        )
        last = row
    _separate_touching_bands(channel)
    return last


def _fill_country_bands(country: CountryRate, grid, row: int, columns: list, rules: Rules) -> None:
    grouped: dict[tuple, PriceBand] = {}
    order: list[tuple] = []
    for column in columns:
        if column.kind in {"min_weight", "ignore"}:
            continue
        value = _price(grid.get((row, column.index)))
        if value is None:
            continue
        key = (column.category, column.band_key, column.pair)
        if key not in grouped:
            start_g, end_g, start_in, end_in = column.band
            grouped[key] = PriceBand(
                category=column.category,
                start_g=start_g,
                end_g=end_g,
                start_inclusive=start_in,
                end_inclusive=end_in,
                label=column.label,
                step_g=column.step_g or Decimal("1"),
                min_g=column.min_g,
                limit_g=column.limit_g,
                currency=column.currency or "CNY",
                rules=rules,
            )
            order.append(key)
        band = grouped[key]
        if value == "bad":
            band.unavailable = True
            continue
        if column.kind == "rate":
            band.rate_per_kg = value
            band.step_g = column.step_g or band.step_g
        elif column.kind == "registration":
            band.registration = value
        elif column.kind == "first":
            band.first_fee = value
            band.first_g = column.first_g
        elif column.kind == "additional":
            band.additional_fee = value
            band.additional_g = column.step_g
        elif column.kind == "absolute":
            band.absolute = value
            band.end_g = column.point_g
            band.start_g = Decimal("0")
        if column.currency:
            band.currency = column.currency
        if column.limit_g:
            band.limit_g = column.limit_g
        if column.min_g:
            band.min_g = column.min_g
    for key in order:
        band = grouped[key]
        if band.registration is None and band.rate_per_kg is not None:
            band.registration = Decimal("0")
        if band.additional_fee is None and band.first_fee is not None:
            band.additional_fee = Decimal("0")
            band.additional_g = band.additional_g or band.first_g
        country.bands.append(band)


class _Column:
    def __init__(self, index: int, kind: str):
        self.index = index
        self.kind = kind
        self.category = "any"
        self.band = (Decimal("0"), None, True, True)
        self.band_key = "all"
        self.label = ""
        self.step_g: Decimal | None = None
        self.first_g: Decimal | None = None
        self.min_g: Decimal | None = None
        self.limit_g: Decimal | None = None
        self.currency: str | None = None
        self.point_g: Decimal | None = None
        self.pair = index


def _collapse_header_rows(rows: list[int]) -> list[int]:
    """纵向合并会让表头占连续几行，只保留这一段的第一行。"""
    kept: list[int] = []
    last = None
    for row in rows:
        if last is not None and row <= last + 1:
            last = row
            continue
        kept.append(row)
        last = row
    return kept


def _header_block_end(grid, header_row: int, next_header: int, max_col: int) -> int:
    """表头「运达国家」常纵向合并，重量段在下面一两行。"""
    name_col, _ = _identity_columns(grid, header_row, max_col)
    end = header_row
    for row in range(header_row + 1, min(header_row + 5, next_header)):
        name = _text(grid.get((row, name_col)))
        blob = " ".join(_text(grid.get((row, col))) for col in range(1, max_col + 1))
        continuation = _is_header_text(name) or (
            not name and any(word in blob for word in ("配送服务费", "挂号", "首重", "续重", "g(含)", "克", "以内"))
        )
        if not continuation:
            break
        end = row
    return end


def _column_specs(grid, header_row, fee_row, max_col, scan_end: int | None = None) -> list[_Column]:
    scan_end = fee_row if scan_end is None else scan_end
    unit_kg = False
    for row in range(header_row, scan_end + 1):
        for col in range(1, max_col + 1):
            if "单位：kg" in _text(grid.get((row, col))).lower() or "单位:kg" in _text(grid.get((row, col))).lower():
                unit_kg = True
    columns: list[_Column] = []
    last_category = "any"
    pair_anchor = 0
    last_band = (Decimal("0"), None, True, True)
    last_band_key = "all"
    last_label = ""
    last_kind = ""
    for col in range(1, max_col + 1):
        texts = [_text(grid.get((row, col))) for row in range(header_row, scan_end + 1)]
        blob = " ".join(part for part in texts if part)
        if not blob:
            continue
        fee_text = ""
        kind = None
        for text in texts:
            found = _kind(text)
            if found:
                kind = found
                fee_text = text
                break
        if kind is None:
            fee_text = _text(grid.get((fee_row, col)))
            kind = _kind(fee_text or blob)
        if kind is None and "最低计费" in blob:
            kind = "min_weight"
        point = _weight_point(fee_text) or _weight_point(blob)
        if kind is None and point is not None:
            kind = "absolute"
        if kind is None:
            continue
        column = _Column(col, kind)
        for text in texts:
            category = _category(text)
            if category:
                last_category = category
                column.category = category
            limit = _limit_g(text)
            if limit:
                column.limit_g = limit
            band = _parse_band(text, unit_kg=unit_kg)
            if band:
                column.band = band
                column.band_key = text
                column.label = text
                last_band = band
                last_band_key = text
                last_label = text
            step = _step(text)
            if step:
                column.step_g = step
            first = _first_g(text)
            if first:
                column.first_g = first
            if "首重" in text and kind == "rate":
                column.min_g = _first_g(text)
            currency = _currency(text)
            if currency:
                column.currency = currency
        if column.category == "any":
            column.category = last_category
        if column.band_key == "all" and column.kind == "registration" and last_kind == "rate":
            column.band = last_band
            column.band_key = last_band_key
            column.label = last_label
        last_kind = column.kind
        if point is not None and kind == "absolute":
            column.point_g = point
            column.band = (Decimal("0"), point, True, True)
            column.band_key = f"abs-{point}"
            column.label = fee_text or blob
        if kind in {"rate", "first"}:
            pair_anchor = col
        column.pair = pair_anchor or col
        if column.kind == "first" and column.first_g is None:
            column.first_g = Decimal("500")
        if column.kind == "additional" and column.step_g is None:
            column.step_g = Decimal("500")
        if column.limit_g and column.band[1] is None and column.kind in {"first", "rate", "registration", "additional"}:
            column.band = (column.band[0], column.limit_g, True, True)
        columns.append(column)
    return columns


def _separate_touching_bands(channel: Channel) -> None:
    for country in channel.countries.values():
        grouped: dict[str, list[PriceBand]] = {}
        for band in country.bands:
            grouped.setdefault(band.category, []).append(band)
        for bands in grouped.values():
            bands.sort(key=lambda item: (item.start_g, item.end_g or Decimal("99999999")))
            for previous, current in zip(bands, bands[1:]):
                if (
                    previous.end_g is not None
                    and current.start_g == previous.end_g
                    and previous.end_inclusive
                    and current.start_inclusive
                ):
                    current.start_inclusive = False


def _country(channel: Channel, name: str, code: str) -> CountryRate:
    key = _norm(name)
    country = channel.countries.get(key)
    if country is None:
        country = CountryRate(name=name.strip(), code=code.strip())
        channel.countries[key] = country
    elif code and not country.code:
        country.code = code.strip()
    return country


def _grid(sheet):
    max_row = min(sheet.max_row or 1, 800)
    max_col = min(sheet.max_column or 1, 90)
    cells: dict[tuple[int, int], object] = {}
    for row in sheet.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            if cell.value not in (None, ""):
                cells[(cell.row, cell.column)] = cell.value
    for merged in sheet.merged_cells.ranges:
        value = cells.get((merged.min_row, merged.min_col))
        if value is None:
            continue
        for row in range(merged.min_row, min(merged.max_row, max_row) + 1):
            for col in range(merged.min_col, min(merged.max_col, max_col) + 1):
                cells.setdefault((row, col), value)
    return cells, max_row, max_col


def _is_header_row(grid, row, max_col) -> bool:
    for col in range(1, max_col + 1):
        if _is_header_text(_text(grid.get((row, col)))):
            return True
    return False


def _is_header_text(text: str) -> bool:
    return any(text.startswith(key) for key in HEADER_KEYS)


def _is_row_band_header(grid, row, max_col) -> bool:
    blob = " ".join(_text(grid.get((row, col))) for col in range(1, max_col + 1))
    return "价格起始重量" in blob or "计费克重段" in blob


def _fee_row(grid, header_row, limit, max_col) -> int:
    point_row = None
    for row in range(header_row, limit + 1):
        blob = " ".join(_text(grid.get((row, col))) for col in range(1, max_col + 1))
        if any(word in blob for word in ("配送服务费", "挂号", "首重", "续重", "重量资费", "Item")):
            return row
        points = sum(1 for col in range(1, max_col + 1) if _weight_point(_text(grid.get((row, col)))))
        if points >= 3:
            point_row = row
    return point_row or header_row


def _identity_columns(grid, header_row, max_col) -> tuple[int, int | None]:
    name_col = 1
    code_col = None
    for col in range(1, max_col + 1):
        text = _text(grid.get((header_row, col)))
        lowered = text.lower()
        if text in {"运达国家/地区", "收件地中文名", "目的国"}:
            name_col = col
        if lowered == "code" or lowered.endswith("(code)") or "（code）" in lowered:
            code_col = col
    return name_col, code_col


def _note_text(grid, start, end, max_col) -> str:
    lines = []
    for row in range(start, end + 1):
        parts = [_text(grid.get((row, col))) for col in range(1, max_col + 1)]
        line = " ".join(part for part in parts if part)
        if line and (_is_note(line)):
            lines.append(line)
    return "\n".join(lines)


def _is_note(line: str) -> bool:
    if _is_header_text(line):
        return False
    return len(line) > 18 or any(word in line for word in ("计费", "体积", "限重", "首重", "计重", "燃油", "最低计费"))


def _row_has_price(grid, row, columns) -> bool:
    return any(_price(grid.get((row, column.index))) is not None for column in columns)


def _kind(text: str) -> str | None:
    compact = text.replace(" ", "")
    if "最低计费" in compact:
        return "min_weight"
    if "续重" in compact or "续加" in compact:
        return "additional"
    if "首重" in compact and "配送" not in compact:
        return "first"
    if "挂号" in compact or "Item服务费" in compact or "item服务费" in compact.lower() or "操作处理费" in compact:
        return "registration"
    if "配送服务费" in compact or "重量资费" in compact:
        return "rate"
    return None


def _category(text: str) -> str | None:
    compact = text.replace(" ", "")
    if "大包" in compact:
        return "large"
    special = "非普货" in compact
    general = "普货" in compact.replace("非普货", "")
    if special and general:
        return "any"
    if special and "香港" in compact:
        return "special_hk"
    if special:
        return "special"
    if general:
        return "general"
    return None


def _parse_band(text: str, unit_kg: bool = False):
    compact = text.replace(" ", "").replace("（", "(").replace("）", ")")
    if "限重" in compact and not re.search(r"\d+\s*[-~～]\s*\d+", compact):
        return None
    match = re.search(r"(\d+(?:\.\d+)?)[-~～](\d+(?:\.\d+)?)(kg|g|克)?", compact, re.I)
    if match:
        start = D(match.group(1))
        end = D(match.group(2))
        unit = (match.group(3) or "").lower()
        if unit == "kg" or (unit == "" and unit_kg):
            start *= 1000
            end *= 1000
        return (start, end, True, True)
    match = re.search(r"(\d+(?:\.\d+)?)(kg|g|克)?(以内|以下)", compact, re.I)
    if match:
        end = D(match.group(1))
        unit = (match.group(2) or "").lower()
        if unit == "kg" or "kg" in compact.lower():
            end *= 1000
        return (Decimal("0"), end, True, True)
    match = re.search(r"[>＞](\d+(?:\.\d+)?)(kg|g|克)?", compact, re.I)
    if match:
        start = D(match.group(1))
        if (match.group(2) or "").lower() == "kg":
            start *= 1000
        return (start, None, False, True)
    return None


def _weight_point(text: str) -> Decimal | None:
    compact = text.strip().lower().replace(" ", "")
    match = re.fullmatch(r"(\d+(?:\.\d+)?)(kg|g|克)", compact)
    if not match:
        return None
    value = D(match.group(1))
    if match.group(2) == "kg":
        value *= 1000
    return value


def _step(text: str) -> Decimal | None:
    compact = text.replace(" ", "")
    match = re.search(r"(?:每|按)(\d+(?:\.\d+)?)(kg|KG|g|G|克)", compact)
    if match:
        step = D(match.group(1))
        if match.group(2).lower() == "kg":
            step *= 1000
        return step
    if "每1g" in compact or "每1克" in compact:
        return Decimal("1")
    return None


def _first_g(text: str) -> Decimal | None:
    compact = text.replace(" ", "")
    match = re.search(r"(\d+(?:\.\d+)?)(kg|KG|g|G|克)首重", compact)
    if not match:
        match = re.search(r"首重(\d+(?:\.\d+)?)(kg|KG|g|G|克)", compact)
    if not match:
        return None
    value = D(match.group(1))
    if match.group(2).lower() == "kg":
        value *= 1000
    return value


def _limit_g(text: str) -> Decimal | None:
    match = re.search(r"限重\s*(\d+(?:\.\d+)?)\s*(kg|KG|g|克)", text.replace(" ", ""), re.I)
    if not match:
        return None
    value = D(match.group(1))
    if match.group(2).lower() == "kg":
        value *= 1000
    return value


def _currency(text: str) -> str | None:
    if "美元" in text or "USD" in text:
        return "USD"
    if any(word in text for word in ("RMB", "CNY", "元")):
        return "CNY"
    return None


def _price(value):
    if value is None or value == "":
        return None
    if isinstance(value, str):
        text = value.strip()
        if text in UNAVAILABLE:
            return "bad"
        text = text.replace(",", "")
        try:
            return D(text)
        except Exception:
            return None
    try:
        return D(value)
    except Exception:
        return None


def _grams(value, unit_kg: bool) -> Decimal | None:
    if value in (None, ""):
        return None
    if isinstance(value, str):
        point = _weight_text(value)
        if point is not None:
            return point
        try:
            number = D(value)
        except Exception:
            return None
    else:
        number = D(value)
    if unit_kg:
        return number * Decimal("1000")
    return number


def _weight_text(value) -> Decimal | None:
    if value in (None, ""):
        return None
    text = str(value).strip().lower().replace(" ", "")
    match = re.fullmatch(r"(\d+(?:\.\d+)?)(g|克|kg)?", text)
    if not match:
        return None
    number = D(match.group(1))
    if match.group(2) == "kg":
        number *= 1000
    return number


def _text(value) -> str:
    if value is None:
        return ""
    return str(value).replace("\n", "").strip()


def _norm(name: str) -> str:
    return name.strip().replace(" ", "").replace("\u00a0", "")


def _find_rate_column(headers: dict[int, str]) -> int | None:
    for col, text in headers.items():
        if "配送服务费" in text and "进位" not in text:
            return col
    return None


def _find_header(headers: dict[int, str], *needles: str) -> int | None:
    for col, text in headers.items():
        lowered = text.lower()
        if any(needle.lower() in lowered for needle in needles):
            return col
    return None
