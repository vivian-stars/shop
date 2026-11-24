"""金额进位。运费结果保留两位小数，四舍五入。"""

from __future__ import annotations

from decimal import Decimal, ROUND_CEILING, ROUND_HALF_UP

TWOPLACES = Decimal("0.01")


def D(value) -> Decimal:
    if isinstance(value, Decimal):
        return value
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        return Decimal(str(value))
    text = str(value).strip().replace(",", "")
    return Decimal(text)


def money(value: Decimal) -> Decimal:
    return D(value).quantize(TWOPLACES, rounding=ROUND_HALF_UP)


def ceil_step(weight_g: Decimal, step_g: Decimal) -> Decimal:
    """不足一个计费档时进到下一档。正好落在档上则不变。"""
    if step_g <= 1:
        return weight_g
    units = (weight_g / step_g).to_integral_value(rounding=ROUND_CEILING)
    return units * step_g
