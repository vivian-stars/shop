"""单个产品和批量表格。"""

from __future__ import annotations

import io
import re
from decimal import Decimal

from openpyxl import Workbook, load_workbook

from shop.shopapp.models import CostOverrides, Product
from shop.shopapp.money import D

MISSING = "未输入产品参数，不能计算运费"
SPEC_HINT = "规格请按 长x宽x高 填写，例如 60x40x20，单位是厘米"
SAMPLE_HEADERS = (
    "sku编号",
    "货物类型",
    "产品规格",
    "产品重量",
    "平台佣金",
    "联盟佣金",
    "金币佣金",
    "一站推佣金",
    "广告费用",
    "损耗",
)


def parse_spec(text: str) -> tuple[Decimal, Decimal, Decimal]:
    raw = (text or "").strip().lower().replace("cm", "").replace("厘米", "")
    raw = raw.replace("×", "x").replace("*", "x").replace("X", "x")
    parts = [part for part in re.split(r"[xX,\s]+", raw) if part]
    if len(parts) != 3:
        raise ValueError(SPEC_HINT)
    dims = tuple(D(part) for part in parts)
    if any(item <= 0 for item in dims):
        raise ValueError(SPEC_HINT)
    return dims  # type: ignore[return-value]


def parse_weight_g(value) -> Decimal:
    if value is None or str(value).strip() == "":
        raise ValueError("请填写产品重量，单位是克")
    text = str(value).strip().lower().replace(" ", "")
    if text.endswith("kg") or text.endswith("千克"):
        number = text.replace("千克", "").replace("kg", "")
        return D(number) * Decimal("1000")
    text = text.replace("克", "").replace("g", "")
    weight = D(text)
    if weight <= 0:
        raise ValueError("产品重量要大于 0")
    return weight


def product_from_fields(sku: str, spec: str, weight) -> Product:
    if not str(sku or "").strip() or not str(spec or "").strip() or weight in (None, ""):
        raise ValueError(MISSING)
    length, width, height = parse_spec(spec)
    return Product(
        sku=str(sku).strip(),
        length_cm=length,
        width_cm=width,
        height_cm=height,
        weight_g=parse_weight_g(weight),
    )


def products_from_excel(payload: bytes) -> list[Product]:
    book = load_workbook(io_bytes(payload), data_only=True)
    sheet = book.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(cell or "").strip() for cell in rows[0]]
    indexes = _column_indexes(header)
    products: list[Product] = []
    for row in rows[1:]:
        if row is None or all(cell in (None, "") for cell in row):
            continue
        sku = _cell(row, indexes.get("sku"))
        spec = _cell(row, indexes.get("spec"))
        weight = _cell(row, indexes.get("weight"))
        if not str(sku or "").strip() and not str(spec or "").strip() and weight in (None, ""):
            continue
        product = product_from_fields(str(sku or ""), str(spec or ""), weight)
        product.cargo = parse_cargo(_cell(row, indexes.get("cargo")))
        product.cost_overrides = _cost_overrides(row, indexes)
        products.append(product)
    return products


def parse_cargo(value) -> str:
    text = str(value or "").strip()
    if text == "":
        return ""
    if text in {"普货", "普通", "general"}:
        return "general"
    if text in {"非普货", "特货", "带电", "special"} or "非普" in text or "带电" in text:
        return "special"
    raise ValueError("货物类型请填普货或非普货")


def sample_workbook() -> bytes:
    book = Workbook()
    sheet = book.active
    sheet.title = "产品"
    sheet.append(list(SAMPLE_HEADERS))
    for index, width in enumerate([18, 14, 18, 14, 14, 14, 14, 14, 14, 12], start=1):
        sheet.column_dimensions[chr(64 + index)].width = width
    buffer = io.BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def _cell(row, index: int | None):
    if index is None or index >= len(row):
        return ""
    return row[index]


def _cost_overrides(row, indexes: dict[str, int | None]) -> CostOverrides | None:
    keys = ("platform", "affiliate", "coin", "one_stop", "ads", "loss")
    if all(indexes.get(key) is None for key in keys):
        return None
    overrides = CostOverrides(**{key: _percent_cell(_cell(row, indexes.get(key))) for key in keys})
    return overrides if overrides.filled() else None


def _percent_cell(value) -> Decimal | None:
    if value is None or str(value).strip() == "":
        return None
    text = str(value).strip().replace("%", "").replace("％", "")
    try:
        number = D(text)
    except Exception as exc:
        raise ValueError("成本比例要填数字") from exc
    if number < 0:
        raise ValueError("成本比例要填数字")
    return number / Decimal("100")


def io_bytes(payload: bytes):
    import io

    return io.BytesIO(payload)


def _column_indexes(header: list[str]) -> dict[str, int | None]:
    def find(*names: str) -> int | None:
        for index, cell in enumerate(header):
            folded = cell.lower()
            if any(name.lower() in folded for name in names):
                return index
        return None

    sku = find("sku", "编号")
    spec = find("规格", "尺寸")
    weight = find("重量")
    if sku is None or spec is None or weight is None:
        raise ValueError("产品表需要这些列：sku编号、产品规格、产品重量")
    return {
        "sku": sku,
        "spec": spec,
        "weight": weight,
        "cargo": find("货物类型", "货品类型"),
        "platform": find("平台佣金"),
        "affiliate": find("联盟佣金"),
        "coin": find("金币佣金"),
        "one_stop": find("一站推"),
        "ads": find("广告"),
        "loss": find("损耗"),
    }
