"""各平台共用的报价数据结构。"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal


@dataclass
class Costs:
    platform: Decimal = Decimal("0.08")
    affiliate: Decimal = Decimal("0.10")
    coin: Decimal = Decimal("0.05")
    one_stop: Decimal = Decimal("0.02")
    ads: Decimal = Decimal("0")
    loss: Decimal = Decimal("0.10")

    def multiplier(self) -> Decimal:
        return (
            Decimal("1")
            + self.platform
            + self.affiliate
            + self.coin
            + self.one_stop
            + self.ads
            + self.loss
        )


@dataclass
class Product:
    sku: str
    length_cm: Decimal
    width_cm: Decimal
    height_cm: Decimal
    weight_g: Decimal

    @property
    def sum_cm(self) -> Decimal:
        return self.length_cm + self.width_cm + self.height_cm

    @property
    def side_cm(self) -> Decimal:
        return max(self.length_cm, self.width_cm, self.height_cm)


@dataclass
class Rules:
    vol_divisor: Decimal | None = None
    vol_scope: str = "none"
    switch_sum_cm: Decimal | None = None
    switch_side_cm: Decimal | None = None
    switch_weight_g: Decimal | None = None
    hard_sum_cm: Decimal | None = None
    hard_side_cm: Decimal | None = None
    hard_weight_g: Decimal | None = None
    country_min_g: dict[str, Decimal] = field(default_factory=dict)
    warning: str = ""

    def merge(self, local: Rules) -> Rules:
        def pick(a, b):
            return b if b not in (None, "none", "", {}) else a

        mins = dict(self.country_min_g)
        mins.update(local.country_min_g)
        return Rules(
            vol_divisor=pick(self.vol_divisor, local.vol_divisor),
            vol_scope=local.vol_scope if local.vol_scope != "none" else self.vol_scope,
            switch_sum_cm=pick(self.switch_sum_cm, local.switch_sum_cm),
            switch_side_cm=pick(self.switch_side_cm, local.switch_side_cm),
            switch_weight_g=pick(self.switch_weight_g, local.switch_weight_g),
            hard_sum_cm=pick(self.hard_sum_cm, local.hard_sum_cm),
            hard_side_cm=pick(self.hard_side_cm, local.hard_side_cm),
            hard_weight_g=pick(self.hard_weight_g, local.hard_weight_g),
            country_min_g=mins,
            warning=local.warning or self.warning,
        )


@dataclass
class PriceBand:
    category: str
    start_g: Decimal
    end_g: Decimal | None
    start_inclusive: bool
    end_inclusive: bool
    label: str
    rate_per_kg: Decimal | None = None
    registration: Decimal | None = None
    first_fee: Decimal | None = None
    first_g: Decimal | None = None
    additional_fee: Decimal | None = None
    additional_g: Decimal | None = None
    step_g: Decimal = Decimal("1")
    min_g: Decimal | None = None
    limit_g: Decimal | None = None
    absolute: Decimal | None = None
    currency: str = "CNY"
    unavailable: bool = False
    rules: Rules = field(default_factory=Rules)

    def contains(self, weight_g: Decimal) -> bool:
        if weight_g < self.start_g:
            return False
        if weight_g == self.start_g and not self.start_inclusive:
            return False
        if self.end_g is None:
            return True
        if weight_g > self.end_g:
            return False
        if weight_g == self.end_g and not self.end_inclusive:
            return False
        return True


@dataclass
class CountryRate:
    name: str
    english: str = ""
    code: str = ""
    min_g: Decimal | None = None
    bands: list[PriceBand] = field(default_factory=list)


@dataclass
class Channel:
    name: str
    countries: dict[str, CountryRate] = field(default_factory=dict)
    rules: Rules = field(default_factory=Rules)
    currencies: set[str] = field(default_factory=set)
    warning: str = ""

    @property
    def country_count(self) -> int:
        return len(self.countries)


@dataclass
class RateBook:
    channels: list[Channel]
    source_name: str

    def channel(self, name: str) -> Channel:
        for item in self.channels:
            if item.name == name:
                return item
        raise KeyError(name)

    def channel_names(self) -> list[str]:
        return [item.name for item in self.channels]


@dataclass
class DestinationQuote:
    country: str
    code: str
    channel: str
    currency: str
    shippable: bool
    detail: str
    billable_g: Decimal | None = None
    rate_per_kg: Decimal | None = None
    registration: Decimal | None = None
    first_fee: Decimal | None = None
    first_g: Decimal | None = None
    additional_fee: Decimal | None = None
    additional_g: Decimal | None = None
    limit_g: Decimal | None = None
    band_end_g: Decimal | None = None
    step_g: Decimal | None = None
    freight: Decimal | None = None
    final_freight: Decimal | None = None
    extra_freight: Decimal | None = None
    mode: str = ""


@dataclass
class ProductQuote:
    product: Product
    rows: list[DestinationQuote]
    warnings: list[str] = field(default_factory=list)


@dataclass
class ExportFile:
    filename: str
    content: bytes
    warnings: list[str] = field(default_factory=list)
