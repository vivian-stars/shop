"""用内置物流费表核对计费和导出。"""

from __future__ import annotations

from decimal import Decimal
from io import BytesIO
from urllib.parse import unquote

from fastapi.testclient import TestClient
from openpyxl import load_workbook

from shop.shopapp.api import STORE, app
from shop.shopapp.models import Costs, Product
from shop.shopapp.money import money
from shop.shopapp.platforms.aliexpress.calc import quote_country

client = TestClient(app)


def _book():
    return STORE.get("default")[1]


def _channel(name="无忧物流-标准"):
    return _book().channel(name)


def _product(weight="200", spec="20x10x5", sku="SKU-1"):
    length, width, height = spec.split("x")
    return Product(sku, Decimal(length), Decimal(width), Decimal(height), Decimal(weight))


def _quote(country, product, cargo="general", costs=None, channel="无忧物流-标准"):
    found = _channel(channel).countries[country]
    return quote_country(_channel(channel), found, product, cargo, costs or Costs())


def test_australia_small_packet_uses_150_to_300_band():
    row = _quote("澳大利亚", _product())
    assert row.shippable
    assert row.billable_g == Decimal("200")
    assert row.freight == Decimal("37.32")
    assert row.final_freight == Decimal("50.38")
    assert "150~300g" in row.detail


def test_australia_large_packet_uses_volumetric_weight():
    row = _quote("澳大利亚", _product(spec="50x30x20"))
    assert row.shippable
    assert row.mode == "first"
    assert row.billable_g == Decimal("4000")
    assert row.freight == Decimal("427.86")
    assert row.final_freight == money(Decimal("427.86") * Decimal("1.35"))


def test_weight_band_boundary_and_special_cargo():
    at_150 = _quote("澳大利亚", _product(weight="150"))
    at_151 = _quote("澳大利亚", _product(weight="151"))
    special = _quote("澳大利亚", _product(), cargo="special")
    assert at_150.rate_per_kg == Decimal("102.59")
    assert at_151.rate_per_kg == Decimal("94.61")
    assert special.rate_per_kg == Decimal("85.08")
    assert special.freight == Decimal("46.11")


def test_oversize_without_large_rate_and_zero_cost_keeps_freight():
    blocked = _quote("玻利维亚", _product(spec="50x30x20"))
    plain = _quote("澳大利亚", _product(), costs=Costs(0, 0, 0, 0, 0, 0))
    korea = _quote("韩国", _product(spec="50x30x20"), costs=Costs(0, 0, 0, 0, 0, 0))
    assert blocked.shippable is False
    assert plain.final_freight == plain.freight == Decimal("37.32")
    assert korea.shippable
    assert korea.freight == Decimal("62.20")


def test_uploaded_book_includes_special_and_commercial_channels():
    names = _book().channel_names()
    for name in (
        "无忧物流-特货标准",
        "无忧物流-特货标快",
        "特货专线－简易",
        "燕文航空挂号小包",
        "4PX新邮挂号小包",
        "中邮e邮宝",
        "无忧物流-优先",
    ):
        assert name in names
    assert "目录" not in names
    standard = _quote("俄罗斯", _product(), channel="无忧物流-特货标准")
    express = _quote("俄罗斯", _product(), channel="无忧物流-特货标快")
    assert standard.freight == Decimal("36.30")
    assert standard.detail.startswith("标准计费 150~300g")
    assert express.freight == Decimal("40.00")
    bahrain = _quote("巴林", _product(), channel="无忧物流-特货标准")
    assert bahrain.freight == Decimal("34.60")


def test_minimum_weight_sea_and_oversized_rounding():
    light = quote_country(
        _channel("超级经济Global"),
        _channel("超级经济Global").countries["安道尔"],
        _product(weight="5"),
        "general",
        Costs(0, 0, 0, 0, 0, 0),
    )
    sea = quote_country(
        _channel("无忧物流-海运大件专线"),
        _channel("无忧物流-海运大件专线").countries["澳大利亚"],
        _product(weight="200", spec="10x10x10"),
        "general",
        Costs(0, 0, 0, 0, 0, 0),
    )
    huge = quote_country(
        _channel("无忧物流-超大件专线-优选仓"),
        _channel("无忧物流-超大件专线-优选仓").countries["韩国"],
        _product(weight="200", spec="10x10x10"),
        "general",
        Costs(0, 0, 0, 0, 0, 0),
    )
    assert light.billable_g == Decimal("10")
    assert light.freight == Decimal("5.86")
    assert sea.freight == Decimal("149.00")
    assert huge.billable_g == Decimal("1000")
    assert huge.freight == Decimal("36.00")


def test_usd_channel_is_marked_and_not_catalog():
    row = quote_country(
        _channel("超级经济"),
        _channel("超级经济").countries["俄罗斯"],
        _product(),
        "general",
        Costs(0, 0, 0, 0, 0, 0),
    )
    names = _book().channel_names()
    assert row.currency == "USD"
    assert row.freight == Decimal("4.91")
    assert "目录" not in names


def test_missing_product_message_and_quantity_export():
    missing = client.post("/api/quote", data={"channels": "无忧物流-标准"})
    assert missing.status_code == 400
    assert missing.json()["error"] == "未输入产品参数，不能计算运费"

    response = client.post(
        "/api/export",
        data={
            "channels": "无忧物流-标准",
            "sku": "SKU-1",
            "spec": "20x10x5",
            "weight": "200",
            "mode": "quantity",
            "qty_start": "1",
            "qty_end": "5",
            "qty_step": "1",
            "platform_fee": "8",
            "affiliate_fee": "10",
            "coin_fee": "5",
            "one_stop_fee": "2",
            "ads_fee": "0",
            "loss_fee": "10",
        },
    )
    assert response.status_code == 200
    assert "200g_按数量计费" in unquote(response.headers["content-disposition"])
    sheet = load_workbook(BytesIO(response.content))["线路编码"]
    rows = {sheet.cell(row, 1).value: row for row in range(3, sheet.max_row + 1)}
    brazil = rows["巴西"]
    argentina = rows["阿根廷"]
    australia = rows["澳大利亚"]
    assert sheet.cell(brazil, 5).value == "卖家承担"
    assert sheet.cell(brazil, 11).value is None
    assert sheet.cell(argentina, 5).value == "按数量计费"
    assert sheet.cell(argentina, 11).value == 77.36
    assert sheet.cell(australia, 9).value == 1
    assert sheet.cell(australia, 10).value == 5
    assert sheet.cell(australia, 11).value == 50.38
    assert sheet.cell(australia, 12).value == 1
    assert sheet.cell(australia, 13).value == 18.92
    assert sheet["A1"].value == "无忧物流-标准"
    assert sheet["C1"].value == "线路代号: CAINIAO_STANDARD"
    assert sheet["F1"].value == "自定义运费"
    assert sheet["H1"].value == "平台标准配置"


def test_weight_export_maps_first_and_additional_columns():
    response = client.post(
        "/api/export",
        data={
            "channels": "无忧物流-标准",
            "sku": "SKU-1",
            "spec": "20x10x5",
            "weight": "200g",
            "mode": "weight",
            "platform_fee": "0",
            "affiliate_fee": "0",
            "coin_fee": "0",
            "one_stop_fee": "0",
            "ads_fee": "0",
            "loss_fee": "0",
        },
    )
    assert response.status_code == 200
    sheet = load_workbook(BytesIO(response.content))["线路编码"]
    australia = next(row for row in range(3, sheet.max_row + 1) if sheet.cell(row, 1).value == "澳大利亚")
    assert sheet.cell(australia, 5).value == "按重量计费"
    assert sheet.cell(australia, 9).value == 0.001
    assert sheet.cell(australia, 10).value == 18.4
    assert sheet.cell(australia, 11).value == 0.3
    assert sheet.cell(australia, 12).value == 0.001
    assert sheet.cell(australia, 13).value == 0.09
    brazil = next(row for row in range(3, sheet.max_row + 1) if sheet.cell(row, 1).value == "巴西")
    assert sheet.cell(brazil, 9).value is None


def test_uploaded_rate_book_replaces_the_default():
    from openpyxl import Workbook

    book = Workbook()
    sheet = book.active
    sheet.title = "测试渠道"
    sheet["A1"] = "运达国家/地区"
    sheet["B1"] = "Code"
    sheet["C1"] = "配送服务费\n元（RMB）/KG\n*每1g计重"
    sheet["D1"] = "挂号服务费\n元（RMB）/包裹"
    sheet["A2"] = "澳大利亚"
    sheet["B2"] = "AU"
    sheet["C2"] = 10
    sheet["D2"] = 1
    buffer = BytesIO()
    book.save(buffer)
    uploaded = client.post(
        "/api/rates",
        data={"platform": "aliexpress"},
        files={"file": ("new.xlsx", buffer.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert uploaded.status_code == 200
    rate_id = uploaded.json()["rate_id"]
    quoted = client.post(
        "/api/quote",
        data={
            "rate_id": rate_id,
            "channels": "测试渠道",
            "sku": "A",
            "spec": "10x10x10",
            "weight": "1000",
            "platform_fee": "0",
            "affiliate_fee": "0",
            "coin_fee": "0",
            "one_stop_fee": "0",
            "ads_fee": "0",
            "loss_fee": "0",
        },
    )
    assert quoted.status_code == 200
    row = quoted.json()["products"][0]["rows"][0]
    assert row["freight"] == "11.00"
    assert row["final_freight"] == "11.00"
