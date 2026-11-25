"""按渠道计费方式计算产品运费和最终运费。"""

from __future__ import annotations

from decimal import Decimal, ROUND_CEILING

from shop.shopapp.models import (
    Channel,
    Costs,
    CountryRate,
    DestinationQuote,
    PriceBand,
    Product,
    ProductQuote,
    RateBook,
)
from shop.shopapp.money import ceil_step, money

CATEGORY_LABEL = {
    "general": "小包普货",
    "special": "小包非普货",
    "special_hk": "小包非普货（香港仓）",
    "large": "大包",
    "any": "标准计费",
}


def quote_products(
    book: RateBook,
    products: list[Product],
    channel_names: list[str],
    cargo: str,
    costs: Costs,
) -> list[ProductQuote]:
    chosen = [book.channel(name) for name in channel_names]
    results = []
    for product in products:
        rows: list[DestinationQuote] = []
        warnings: list[str] = []
        for channel in chosen:
            if channel.warning and channel.warning not in warnings:
                warnings.append(f"{channel.name}：{channel.warning}")
            if "USD" in channel.currencies and "CNY" not in channel.currencies:
                warnings.append(f"{channel.name} 的报价是美元，只在页面展示，不写入人民币运费模板")
            for country in channel.countries.values():
                rows.append(quote_country(channel, country, product, product.cargo or cargo, product.costs or costs))
        results.append(ProductQuote(product=product, rows=rows, warnings=warnings))
    return results


def quote_country(
    channel: Channel,
    country: CountryRate,
    product: Product,
    cargo: str,
    costs: Costs,
) -> DestinationQuote:
    base = _blank(channel, country)
    category = _category(channel, country, product, cargo)
    if category is None:
        base.detail = "尺寸或重量超出该渠道限制"
        return base
    pool = _pool(country, category, cargo)
    if not pool:
        base.detail = "这个渠道没有对应的计费档"
        return base
    billable = _prepare_weight(product, category, channel, country, pool)
    band = _pick(pool, billable)
    if band is None:
        base.detail = "重量不在该渠道的计费档里"
        return base
    billable = _round_weight(billable, band, country, channel)
    narrowed = _pick(pool, billable)
    if narrowed is not None:
        band = narrowed
        billable = _round_weight(billable, band, country, channel)
    if band.unavailable:
        base.detail = "该重量段暂无服务"
        return base
    if band.limit_g is not None and billable > band.limit_g:
        base.detail = "超过该渠道限重"
        return base
    freight, extra, mode = _price(band, billable)
    if freight is None:
        base.detail = "该重量段没有价格"
        return base
    freight = money(freight)
    final = money(freight * costs.multiplier())
    extra_money = money(extra) if extra is not None else None
    label = CATEGORY_LABEL.get(band.category, CATEGORY_LABEL.get(category, "计费"))
    band_text = band.label or "按重量"
    return DestinationQuote(
        country=country.name,
        code=country.code,
        channel=channel.name,
        currency=band.currency,
        shippable=True,
        detail=f"{label} {band_text}，计费重量 { _grams(billable) }g",
        billable_g=billable,
        rate_per_kg=band.rate_per_kg,
        registration=band.registration,
        first_fee=band.first_fee,
        first_g=band.first_g,
        additional_fee=band.additional_fee,
        additional_g=band.additional_g,
        limit_g=band.limit_g,
        band_end_g=band.end_g,
        step_g=band.step_g if mode == "per_kg" else band.additional_g,
        freight=freight,
        final_freight=final,
        extra_freight=extra_money,
        mode=mode,
    )


def _blank(channel: Channel, country: CountryRate) -> DestinationQuote:
    currency = "USD" if channel.currencies == {"USD"} else "CNY"
    return DestinationQuote(
        country=country.name,
        code=country.code,
        channel=channel.name,
        currency=currency,
        shippable=False,
        detail="",
    )


def _category(channel: Channel, country: CountryRate, product: Product, cargo: str) -> str | None:
    rules = channel.rules
    if rules.hard_side_cm and product.side_cm > rules.hard_side_cm:
        return None
    if rules.hard_sum_cm and product.sum_cm > rules.hard_sum_cm:
        return None
    if rules.hard_weight_g and product.weight_g > rules.hard_weight_g:
        return None
    oversized = False
    if rules.switch_weight_g and product.weight_g > rules.switch_weight_g:
        oversized = True
    if rules.switch_side_cm and product.side_cm > rules.switch_side_cm:
        oversized = True
    if rules.switch_sum_cm and product.sum_cm > rules.switch_sum_cm:
        oversized = True
    has_large = any(band.category == "large" for band in country.bands)
    if oversized:
        return "large" if has_large else None
    if cargo == "special" and any(band.category == "special" for band in country.bands):
        return "special"
    if any(band.category == "general" for band in country.bands):
        return "general"
    return "any"


def _pool(country: CountryRate, category: str, cargo: str) -> list[PriceBand]:
    if category == "large":
        return [band for band in country.bands if band.category == "large"]
    if category == "special":
        return [band for band in country.bands if band.category == "special"]
    if category == "general":
        found = [band for band in country.bands if band.category == "general"]
        return found or [band for band in country.bands if band.category == "any"]
    if cargo == "special":
        special = [band for band in country.bands if band.category == "special"]
        if special:
            return special
    return [band for band in country.bands if band.category in {"any", "general"}] or list(country.bands)


def _prepare_weight(
    product: Product,
    category: str,
    channel: Channel,
    country: CountryRate,
    pool: list[PriceBand],
) -> Decimal:
    billable = product.weight_g
    divisor, scope = _volume_rule(channel, pool, category)
    if not isinstance(divisor, Decimal):
        divisor = None
    if divisor and (scope == "all" or (scope == "large" and category == "large")):
        billable = max(billable, _volumetric_g(product, divisor))
    minimum = country.min_g or channel.rules.country_min_g.get(country.name)
    if minimum:
        billable = max(billable, minimum)
    return billable


def _volume_rule(channel: Channel, pool: list[PriceBand], category: str) -> tuple[Decimal | None, str]:
    divisors = {band.rules.vol_divisor for band in pool if band.rules.vol_divisor}
    if len(divisors) == 1:
        scope = "all" if any(band.rules.vol_scope == "all" for band in pool) else channel.rules.vol_scope
        if category == "large" and channel.rules.vol_scope == "large":
            scope = "large"
        return divisors.pop(), scope or "all"
    if channel.rules.vol_divisor:
        return channel.rules.vol_divisor, channel.rules.vol_scope
    return None, "none"


def _round_weight(weight: Decimal, band: PriceBand, country: CountryRate, channel: Channel) -> Decimal:
    minimum = country.min_g or channel.rules.country_min_g.get(country.name) or band.min_g
    if band.min_g:
        minimum = max(minimum or Decimal("0"), band.min_g)
    if minimum:
        weight = max(weight, minimum)
    if band.first_fee is not None and band.first_g:
        weight = max(weight, band.first_g)
        if band.additional_g:
            weight = ceil_step(weight, band.additional_g)
        return weight
    if band.rate_per_kg is not None and band.step_g and band.step_g > 1:
        weight = ceil_step(weight, band.step_g)
    return weight


def _pick(bands: list[PriceBand], weight: Decimal) -> PriceBand | None:
    matches = [band for band in bands if band.contains(weight)]
    if not matches:
        return None
    matches.sort(key=lambda band: (band.end_g is None, band.end_g or Decimal("0")))
    return matches[0]


def _price(band: PriceBand, weight: Decimal) -> tuple[Decimal | None, Decimal | None, str]:
    if band.absolute is not None:
        return band.absolute, None, "absolute"
    if band.first_fee is not None and band.first_g:
        steps = Decimal("0")
        if weight > band.first_g and band.additional_g:
            steps = ((weight - band.first_g) / band.additional_g).to_integral_value(rounding=ROUND_CEILING)
        extra = steps * (band.additional_fee or Decimal("0"))
        return band.first_fee + extra, extra, "first"
    if band.rate_per_kg is None:
        return None, None, ""
    extra = weight / Decimal("1000") * band.rate_per_kg
    registration = band.registration or Decimal("0")
    return extra + registration, extra, "per_kg"


def _volumetric_g(product: Product, divisor: Decimal) -> Decimal:
    return product.length_cm * product.width_cm * product.height_cm / divisor * Decimal("1000")


def _grams(weight: Decimal) -> str:
    if weight == weight.to_integral_value():
        return str(int(weight))
    return format(weight.normalize(), "f")
