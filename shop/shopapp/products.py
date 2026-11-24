"""单个产品和批量表格。"""

from __future__ import annotations

import re
from decimal import Decimal

from openpyxl import load_workbook

from shop.shopapp.models import Product
from shop.shopapp.money import D

MISSING = "未输入产品参数，不能计算运费"
SPEC_HINT = "规格请按 长x宽x高 填写，例如 60x40x20，单位是厘米"


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
        sku = row[indexes["sku"]] if indexes["sku"] < len(row) else ""
        spec = row[indexes["spec"]] if indexes["spec"] < len(row) else ""
        weight = row[indexes["weight"]] if indexes["weight"] < len(row) else ""
        if not str(sku or "").strip() and not str(spec or "").strip() and weight in (None, ""):
            continue
        products.append(product_from_fields(str(sku or ""), str(spec or ""), weight))
    return products


def io_bytes(payload: bytes):
    import io

    return io.BytesIO(payload)


def _column_indexes(header: list[str]) -> dict[str, int]:
    def find(*names: str) -> int | None:
        for index, cell in enumerate(header):
            if any(name in cell for name in names):
                return index
        return None

    sku = find("sku", "SKU", "编号")
    spec = find("规格", "尺寸")
    weight = find("重量")
    if sku is None or spec is None or weight is None:
        raise ValueError("产品表需要三列：sku编号、产品规格、产品重量")
    return {"sku": sku, "spec": spec, "weight": weight}
