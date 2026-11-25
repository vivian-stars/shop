"""包邮基准表要和 0.3KG 普货、基准 45 元的样表一致。"""

from __future__ import annotations

from decimal import Decimal
from io import BytesIO

from fastapi.testclient import TestClient
from openpyxl import load_workbook

from shop.shopapp.api import app
from shop.shopapp.models import Costs, Product
from shop.shopapp.platforms.aliexpress.benchmark import benchmark_rows
from shop.shopapp.platforms.aliexpress.calc import quote_products
from shop.shopapp.store import RateStore

client = TestClient(app)


def _rows():
    book = RateStore().get("default")[1]
    product = Product("SKU", Decimal("20"), Decimal("10"), Decimal("5"), Decimal("300"))
    quote = quote_products(book, [product], ["无忧物流-标准"], "general", Costs())[0]
    return {row.country: row for row in benchmark_rows(quote, Decimal("45"), Decimal("0.2"))}


def test_australia_matches_gina_sheet():
    row = _rows()["澳大利亚"]
    assert row.note == "额外设置包邮国家"
    assert row.regional_name == "澳大利亚"
    assert row.regional_set == "是"
    assert row.freight == Decimal("46.783")
    assert row.commission == Decimal("9.3566")
    assert row.total == Decimal("56.1396")
    assert row.difference == Decimal("11.1396")
    assert row.regional_price == Decimal("13.36752")
    assert row.template_fee is None


def test_malaysia_is_free_shipping_and_spain_is_extra():
    rows = _rows()
    malaysia = rows["马来西亚"]
    spain = rows["西班牙"]
    assert malaysia.note == "包邮国家"
    assert malaysia.freight == Decimal("14.9")
    assert malaysia.total == Decimal("17.88")
    assert malaysia.regional_price == Decimal("-21.696")
    assert spain.note == "额外设置包邮国家"
    assert spain.freight == Decimal("38.661")
    assert spain.regional_price == Decimal("1.67184")


def test_kazakhstan_is_free_without_regional_name_and_sri_lanka_uses_template_fee():
    rows = _rows()
    kazakhstan = rows["哈萨克斯坦"]
    sri_lanka = rows["斯里兰卡"]
    assert kazakhstan.note == "包邮国家"
    assert kazakhstan.regional_set == "否"
    assert kazakhstan.regional_price == Decimal("-19.2")
    assert sri_lanka.note == "不包邮国家"
    assert sri_lanka.regional_price is None
    assert sri_lanka.template_fee == Decimal("22.6224")
    assert sri_lanka.extra_each == Decimal("36.21")
    assert sri_lanka.extra_diff == Decimal("-8.79")


def test_export_workbook_uses_gina_headers_and_discount_sheet():
    response = client.post(
        "/api/export",
        data={
            "mode": "benchmark",
            "channels": "无忧物流-标准",
            "sku": "SKU-1",
            "spec": "20x10x5",
            "weight": "300",
            "benchmark": "45",
            "platform_fee": "8",
            "affiliate_fee": "10",
            "coin_fee": "0",
            "one_stop_fee": "2",
            "ads_fee": "0",
            "loss_fee": "0",
            "cargo": "general",
        },
    )
    assert response.status_code == 200, response.text
    book = load_workbook(BytesIO(response.content), data_only=False)
    assert book.sheetnames == ["无忧物流-标准", "区域定价国家调整金额计算"]
    sheet = book["无忧物流-标准"]
    assert "重量=0.3KG" in str(sheet["E1"].value)
    assert sheet["A2"].value == "马来西亚"
    assert _same(sheet["E2"].value, "14.9")
    assert _same(sheet["G2"].value, "17.88")
    discount = book["区域定价国家调整金额计算"]
    assert discount["A2"].value == "俄罗斯"
    assert _same(discount["D2"].value, "10.44")
    korea = {discount.cell(row, 1).value: discount.cell(row, 4).value for row in range(2, 30)}
    assert _same(korea["韩国"], "12.18816")


def _same(actual, expected: str) -> bool:
    return Decimal(str(actual)) == Decimal(expected)


def _fees(platform, affiliate, coin, one_stop, ads, loss):
    return {
        "platform_fee": platform,
        "affiliate_fee": affiliate,
        "coin_fee": coin,
        "one_stop_fee": one_stop,
        "ads_fee": ads,
        "loss_fee": loss,
    }


def test_commission_follows_the_six_cost_fields():
    common = {
        "channels": "无忧物流-标准",
        "sku": "SKU-1",
        "spec": "20x10x5",
        "weight": "300",
        "benchmark": "45",
        "cargo": "general",
    }
    low = client.post("/api/quote", data={**common, **_fees("8", "10", "0", "2", "0", "0")})
    high = client.post("/api/quote", data={**common, **_fees("8", "10", "5", "2", "0", "10")})
    assert low.status_code == 200, low.text
    assert high.status_code == 200, high.text
    low_row = next(row for row in low.json()["products"][0]["rows"] if row["country"] == "澳大利亚")
    high_row = next(row for row in high.json()["products"][0]["rows"] if row["country"] == "澳大利亚")
    assert low.json()["products"][0]["commission_percent"] == "20"
    assert low_row["benchmark_total"] == "56.1396"
    assert high.json()["products"][0]["commission_percent"] == "35"
    assert high_row["benchmark_freight"] == "46.783"
    assert high_row["benchmark_total"] == "63.15705"


def test_batch_file_exports_one_workbook_per_sku():
    from io import BytesIO
    from zipfile import ZipFile

    from openpyxl import Workbook

    book = Workbook()
    sheet = book.active
    sheet.append(["sku编号", "货物类型", "产品规格", "产品重量", "平台佣金", "联盟佣金", "金币佣金", "一站推佣金", "广告费用", "损耗"])
    sheet.append(["SKU-A", "普货", "20x10x5", 300, 8, 10, 0, 2, 0, 0])
    sheet.append(["SKU-B", "普货", "20x10x5", 300, 8, 10, 5, 2, 0, 10])
    buffer = BytesIO()
    book.save(buffer)
    response = client.post(
        "/api/export",
        data={
            "mode": "benchmark",
            "product_mode": "batch",
            "channels": "无忧物流-标准",
            "benchmark": "45",
            **_fees("1", "1", "1", "1", "1", "1"),
        },
        files={"products_file": ("products.xlsx", buffer.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert response.status_code == 200, response.text
    assert response.headers["content-type"].startswith("application/zip")
    archive = ZipFile(BytesIO(response.content))
    names = sorted(archive.namelist())
    assert names == ["300g 普货 包邮基准45元-SKU-A.xlsx", "300g 普货 包邮基准45元-SKU-B.xlsx"]
    totals = {}
    for name in names:
        sheet = load_workbook(BytesIO(archive.read(name)))["无忧物流-标准"]
        for row in range(2, sheet.max_row + 1):
            if sheet.cell(row, 1).value == "澳大利亚":
                totals[name] = sheet.cell(row, 7).value
    assert _same(totals[names[0]], "56.1396")
    assert _same(totals[names[1]], "63.15705")


def test_product_sample_is_a_blank_workbook():
    response = client.get("/api/products/sample")
    assert response.status_code == 200
    sheet = load_workbook(BytesIO(response.content)).active
    assert [cell.value for cell in sheet[1]] == [
        "sku编号", "货物类型", "产品规格", "产品重量",
        "平台佣金", "联盟佣金", "金币佣金", "一站推佣金", "广告费用", "损耗",
    ]
    assert sheet["A2"].value is None


def test_quantity_export_uses_defaults_when_fields_are_empty():
    response = client.post(
        "/api/export",
        data={
            "mode": "quantity",
            "channels": "无忧物流-标准",
            "sku": "SKU-1",
            "spec": "20x10x5",
            "weight": "200",
            "qty_start": "",
            "qty_end": "",
            "qty_step": "",
            **_fees("8", "10", "5", "2", "0", "10"),
        },
    )
    assert response.status_code == 200, response.text
    sheet = load_workbook(BytesIO(response.content))["线路编码"]
    australia = next(row for row in range(3, sheet.max_row + 1) if sheet.cell(row, 1).value == "澳大利亚")
    assert sheet.cell(australia, 9).value == 1
    assert sheet.cell(australia, 10).value == 5
    assert sheet.cell(australia, 12).value == 1
    assert sheet.cell(australia, 11).value == 50.38
