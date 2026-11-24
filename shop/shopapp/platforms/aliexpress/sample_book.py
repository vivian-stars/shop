"""生成内置运费表和两份导出模板。

物流费原件在这台环境里已经不在了。这里按当时读到的表头、合并单元格和报价行重建一份可计算的工作簿，
列的含义和计费说明与原表一致。上传完整的物流费表之后，计算会改用新表。
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

DATA = Path(__file__).resolve().parents[3] / "data"

NOTE = (
    "1. 小包1g起重，按克计费。大包计重规则，实际重量与体积重取大值，（体积重 ：长(cm)*宽(cm)*高(cm)/8000）；"
    "欧向25国1g起重，按g计费；其余国家0.5kg起重，每500g计费。"
    "小包普货、小包非普货及大包计费标准不同，部分国家不支持寄送大包货物。"
)
NOTE_GENERAL = "2. 小包普货计费：适用于普货（申报且实际品类不含有电池/非液体化妆品），且包裹重量（申报重量且实际重量）＜=2kg，且包裹单边长度<=60cm，且包裹长+宽+高<=90cm"
NOTE_SPECIAL = "3. 小包非普货计费：适用于非普货（申报且实际品类含电池/非液体化妆品），且包裹重量（申报重量且实际重量）＜=2kg，且包裹单边长度<=60cm，且包裹长+宽+高<=90cm"
NOTE_LARGE = "4. 大包计费：包裹申报重量>2kg，或包裹实际重量>2kg，或包裹单边长度>60cm，或包裹长+宽+高>90cm"

# 小包普货三段、小包非普货一段、大包首重和每 500g 续重。空值表示这一档没有报价。
STANDARD_ROWS = [
    ("阿尔巴尼亚", "Albania", "AL", 209.66, 18.7, 209.66, 18.7, 209.66, 18.7, 209.66, 19.8, 413.82, 170.83),
    ("阿尔及利亚", "Algeria", "DZ", 193.26, 17, 193.26, 17, 193.26, 17, 193.26, 18, 404.99, 138.5),
    ("安哥拉", "Angola", "AO", 176.7, 17, 176.7, 17, 176.7, 17, 176.7, 18, 362, 141.2),
    ("安提瓜和巴布达", "Antigua and Barbuda", "AG", 205.1, 17, 205.1, 17, 205.1, 17, 205.1, 18, 376.2, 155.3),
    ("亚美尼亚", "Armenia", "AM", 175.2, 16.6, 175.2, 16.6, 175.2, 16.6, 175.2, 17.6, 371, 150.1),
    ("阿鲁巴岛", "Aruba", "AW", 205.1, 17, 205.1, 17, 205.1, 17, 205.1, 18, 376.2, 155.3),
    ("澳大利亚", "Australia", "AU", 102.59, 18.4, 94.61, 18.4, 89.52, 18.4, 85.08, 29.09, 166.55, 37.33),
    ("巴林", "Bahrain", "BH", 194.8, 17.6, 187.29, 17.6, 180.33, 17.6, 205.3, 18.6, 122.79, 20.22),
    ("白俄罗斯", "Belarus", "BY", 136, 21, 136, 21, 136, 21, 165, 25, 168.5, 83.3),
    ("布基纳法索", "Burkina Faso", "BF", 205.1, 17, 205.1, 17, 205.1, 17, 205.1, 18, 376.2, 155.3),
    ("加拿大", "Canada", "CA", 122.44, 27.6, 122.44, 27.6, 122.44, 27.6, 110.59, 32.3, 171.3, 42.7),
    ("日本", "Japan", "JP", 64.45, 21.2, 64.45, 21.2, 64.45, 21.2, 64.45, 22.2, 108.13, 25.7),
    ("韩国", "Korea", "KR", 29.68, 15.9, 14.29, 22.4, 31.4, 4.4, None, None, None, None),
    ("法国", "France", "FR", 91.74, 17.02, 91.74, 17.02, 70.01, 22.02, 94.15, 17.02, 52.67, 70.71),
    ("德国", "Germany", "DE", 70.44, 18.85, 70.44, 18.85, 70.44, 18.85, 88.65, 21.85, 61.54, 54.93),
    ("意大利", "Italy", "IT", 63.8, 23.4, 63.8, 23.4, 63.8, 23.4, 91.68, 23.4, 55.3, 66.98),
    ("西班牙", "Spain", "ES", 64.67, 19.26, 64.67, 19.26, 64.67, 19.26, 85.41, 19.26, 62.34, 82.22),
    ("英国", "United Kingdom", "GB", 67.76, 15.9, 67.76, 15.9, 67.76, 15.9, 83.3, 17, 59.24, 64.04),
]

# 原模板里已经标成卖家承担的国家，承诺天数按原表。
SELLER_ROWS = [
    ("巴西", 75),
    ("马来西亚", 60),
    ("哈萨克斯坦", 60),
    ("韩国", 60),
    ("泰国", 60),
    ("新加坡", 60),
    ("印度尼西亚", 60),
    ("波兰", 60),
    ("英国", 60),
    ("柬埔寨", 60),
    ("西班牙", 60),
    ("德国", 60),
    ("日本", 60),
    ("瑞典", 60),
    ("卢森堡", 60),
    ("意大利", 60),
    ("捷克", 60),
    ("俄罗斯", 60),
    ("法国", 60),
    ("奥地利", 60),
    ("丹麦", 60),
    ("比利时", 60),
    ("斯洛伐克", 60),
]
QUANTITY_ROWS = [
    ("阿根廷", 120),
    ("博茨瓦纳", 120),
    ("加蓬", 120),
    ("玻利维亚", 120),
    ("莫桑比克", 90),
    ("巴拉圭", 90),
    ("厄瓜多尔", 90),
    ("乌拉圭", 90),
    ("斯里兰卡", 60),
    ("挪威", 60),
    ("巴基斯坦", 60),
    ("摩洛哥", 60),
    ("瑞士", 60),
    ("马耳他", 60),
    ("文莱", 60),
]


def build_rate_book(path: Path | None = None) -> Path:
    path = path or DATA / "default_rates.xlsx"
    book = Workbook()
    catalog = book.active
    catalog.title = "目录"
    catalog["A1"] = "线路名称"
    catalog["B1"] = "本表按已读取的物流费报价整理。上传新的物流费表后，以新表为准。"
    _standard_sheet(book.create_sheet("无忧物流-标准"))
    _global_sheet(book.create_sheet("超级经济Global"))
    _economy_usd_sheet(book.create_sheet("超级经济"))
    _weihai_sheet(book.create_sheet("威海优选仓-韩国经济专线"))
    _sea_sheet(book.create_sheet("无忧物流-海运大件专线"))
    _oversized_sheet(book.create_sheet("无忧物流-超大件专线-优选仓"))
    _third_party_sheet(book.create_sheet("无忧物流-标准-三方清关"))
    path.parent.mkdir(parents=True, exist_ok=True)
    book.save(path)
    return path


def build_templates(qty_path: Path | None = None, weight_path: Path | None = None) -> tuple[Path, Path]:
    qty_path = qty_path or DATA / "template_qty.xlsx"
    weight_path = weight_path or DATA / "template_weight.xlsx"
    _save_template(qty_path, "quantity")
    _save_template(weight_path, "weight")
    return qty_path, weight_path


def _standard_sheet(sheet) -> None:
    sheet["A1"] = "无忧物流-标准介绍"
    sheet["A7"] = "计费方式："
    sheet["A8"] = NOTE
    sheet["A9"] = NOTE_GENERAL
    sheet["A10"] = NOTE_SPECIAL
    sheet["A11"] = NOTE_LARGE
    _write_standard_header(sheet, 13)
    row = 16
    for item in STANDARD_ROWS:
        _write_standard_country(sheet, row, item)
        row += 1
    row += 2
    _write_russia(sheet, row)
    row += 6
    _write_ukraine(sheet, row)
    row += 6
    _write_wide(sheet, row)


def _write_standard_header(sheet, row: int) -> None:
    sheet.cell(row, 1, "运达国家/地区")
    sheet.cell(row, 2, "运达国家/地区（英文）")
    sheet.cell(row, 3, "Code")
    sheet.cell(row, 4, "小包普货计费")
    sheet.cell(row, 10, "小包非普货计费")
    sheet.cell(row, 12, "大包计费 限重30kg")
    sheet.merge_cells(start_row=row, start_column=4, end_row=row, end_column=9)
    sheet.merge_cells(start_row=row, start_column=10, end_row=row, end_column=11)
    sheet.merge_cells(start_row=row, start_column=12, end_row=row, end_column=13)
    band = row + 1
    bands = [(4, "0~150g(含)"), (6, "150~300g(含)"), (8, "300~2000g(含)"), (10, "0~2000g(含)")]
    for col, text in bands:
        sheet.cell(band, col, text)
        sheet.merge_cells(start_row=band, start_column=col, end_row=band, end_column=col + 1)
    fee = row + 2
    for col in (4, 6, 8, 10):
        sheet.cell(fee, col, "配送服务费\n元（RMB）/KG\n*每1g计重")
        sheet.cell(fee, col + 1, "挂号服务费\n元（RMB）/包裹")
    sheet.cell(fee, 12, "0.5KG首重\n元（RMB）/包裹")
    sheet.cell(fee, 13, "续重(每500G)\n元（RMB）/500G")


def _write_standard_country(sheet, row: int, item) -> None:
    for col, value in enumerate(item, start=1):
        if value is not None:
            sheet.cell(row, col, value)


def _write_russia(sheet, row: int) -> None:
    sheet.cell(row, 1, "运达国家/地区")
    sheet.cell(row, 2, "运达国家/地区（英文）")
    sheet.cell(row, 3, "Code")
    sheet.cell(row, 4, "小包普货计费")
    sheet.cell(row, 10, "小包非普货计费")
    sheet.cell(row, 16, "大包计费 限重30kg")
    sheet.merge_cells(start_row=row, start_column=4, end_row=row, end_column=9)
    sheet.merge_cells(start_row=row, start_column=10, end_row=row, end_column=15)
    sheet.merge_cells(start_row=row, start_column=16, end_row=row, end_column=17)
    labels = [(4, "0~150g(含)"), (6, "150~300g(含)"), (8, "300~2000g(含)"), (10, "0~150g(含)"), (12, "150~300g(含)"), (14, "300~2000g(含)")]
    for col, text in labels:
        sheet.cell(row + 1, col, text)
        sheet.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 1)
    for col in (4, 6, 8, 10, 12, 14):
        sheet.cell(row + 2, col, "配送服务费\n元（RMB）/KG\n*每1g计重")
        sheet.cell(row + 2, col + 1, "挂号服务费\n元（RMB）/包裹")
    sheet.cell(row + 2, 16, "0.5KG首重\n元（RMB）/包裹")
    sheet.cell(row + 2, 17, "续重(每500G)\n元（RMB）/500G")
    values = ["俄罗斯", "Russian Federation", "RU", 81, 17, 81, 19, 83, 20, 86.5, 17, 86.5, 19, 90, 20, 97, 74]
    for col, value in enumerate(values, start=1):
        sheet.cell(row + 3, col, value)


def _write_ukraine(sheet, row: int) -> None:
    sheet.cell(row, 1, "运达国家/地区")
    sheet.cell(row, 2, "运达国家/地区（英文）")
    sheet.cell(row, 3, "Code")
    sheet.cell(row, 4, "小包普货计费")
    sheet.cell(row, 6, "小包非普货计费（其他仓）")
    sheet.cell(row, 10, "小包非普货计费（香港仓）")
    sheet.cell(row, 12, "大包计费 限重30kg")
    sheet.merge_cells(start_row=row, start_column=4, end_row=row, end_column=5)
    sheet.merge_cells(start_row=row, start_column=6, end_row=row, end_column=9)
    sheet.merge_cells(start_row=row, start_column=10, end_row=row, end_column=11)
    sheet.merge_cells(start_row=row, start_column=12, end_row=row, end_column=13)
    sheet.cell(row + 1, 4, "0~2000g(含)")
    sheet.cell(row + 1, 6, "0~200g(含)")
    sheet.cell(row + 1, 8, "200~2000g(含)")
    sheet.cell(row + 1, 10, "0~2000g(含)")
    for col in (4, 6, 8, 10):
        sheet.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 1)
        sheet.cell(row + 2, col, "配送服务费\n元（RMB）/KG\n*每1g计重")
        sheet.cell(row + 2, col + 1, "挂号服务费\n元（RMB）/包裹")
    sheet.cell(row + 2, 12, "0.5KG首重\n元（RMB）/包裹")
    sheet.cell(row + 2, 13, "续重(每500G)\n元（RMB）/500G")
    values = ["乌克兰", "Ukraine", "UA", 118.98, 7.4, 0, 27.2, 132.33, 7.4, 152.97, 24.4, 65.4, 60.71]
    for col, value in enumerate(values, start=1):
        sheet.cell(row + 3, col, value)


def _write_wide(sheet, row: int) -> None:
    sheet.cell(row, 1, "运达国家/地区")
    sheet.cell(row, 2, "运达国家/地区（英文）")
    sheet.cell(row, 3, "Code")
    sheet.cell(row, 4, "小包普货计费")
    sheet.cell(row, 12, "小包非普货计费")
    sheet.cell(row, 20, "大包计费 限重30kg")
    sheet.merge_cells(start_row=row, start_column=4, end_row=row, end_column=11)
    sheet.merge_cells(start_row=row, start_column=12, end_row=row, end_column=19)
    sheet.merge_cells(start_row=row, start_column=20, end_row=row, end_column=23)
    bands = [
        (4, "0~200g(含)"),
        (6, "201~400g(含)"),
        (8, "401~700g(含)"),
        (10, "701~2000g(含)"),
        (12, "0~200g(含)"),
        (14, "201~400g(含)"),
        (16, "401~700g(含)"),
        (18, "701~2000g(含)"),
        (20, "0-2100g(含)"),
        (22, "2100-30000g"),
    ]
    for col, text in bands:
        sheet.cell(row + 1, col, text)
        sheet.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 1)
    for col in range(4, 20, 2):
        sheet.cell(row + 2, col, "配送服务费\n元（RMB）/KG\n*每1g计重")
        sheet.cell(row + 2, col + 1, "挂号服务费\n元（RMB）/包裹")
    sheet.cell(row + 2, 20, "0.5KG首重\n元（RMB）/包裹")
    sheet.cell(row + 2, 21, "续重(每500G)\n元（RMB）/500G")
    sheet.cell(row + 2, 22, "0.5KG首重\n元（RMB）/包裹")
    sheet.cell(row + 2, 23, "续重(每500G)\n元（RMB）/500G")
    us = ["美国", "United States", "US", 72.45, 30.45, 72.45, 31.25, 72.7, 32.95, 71.13, 38.15, 77.04, 31.65, 77.04, 32.55, 77.28, 34.45, 75.71, 40.85, 160.64, 27.11, 214.6, 48.3]
    br = ["巴西", "Brazil", "BR", 68.26, 29.01, 81.92, 32.01, 81.92, 34.01, 89.36, 35.01, 90.6, 29.01, 90.6, 32.01, 90.6, 34.01, 90.6, 35.01, 112.2, 87.65, 124.11, 96.44]
    for col, value in enumerate(us, start=1):
        sheet.cell(row + 3, col, value)
    for col, value in enumerate(br, start=1):
        sheet.cell(row + 4, col, value)


def _global_sheet(sheet) -> None:
    sheet["A8"] = "运达国家/地区"
    sheet["B8"] = "运达国家/地区（英文）"
    sheet["C8"] = "Code"
    sheet["D8"] = "包裹重量为1-2000克"
    sheet["F8"] = "最低计费重量"
    sheet["D9"] = "配送服务费（根据包裹重量按克计费）\nRMB/KG"
    sheet["E9"] = "Item 服务费\nRMB/包裹"
    rows = [
        ("安道尔", "Andorra", "AD", 186.34, 4, "10g"),
        ("安哥拉", "Angola", "AO", 183.34, 0.65, "10g"),
        ("亚美尼亚", "Armenia", "AM", 145.38, 0.65, "10g"),
        ("孟加拉国", "Bangladesh", "BD", 115.41, 0.65, "10g"),
        ("保加利亚", "Bulgaria", "BG", 78.88, 1.36, "10g"),
        ("芬兰", "Finland", "FI", 104.56, 7.94, "10g"),
    ]
    for offset, item in enumerate(rows):
        for col, value in enumerate(item, start=1):
            sheet.cell(10 + offset, col, value)


def _economy_usd_sheet(sheet) -> None:
    sheet["A4"] = "配送服务费（根据包裹重量按克计费）"
    sheet["C5"] = "美元（USD）/KG"
    sheet["A9"] = "运达国家/地区"
    sheet["B9"] = "运达国家/地区（英文）"
    sheet["C9"] = "Code"
    sheet["D9"] = "配送服务费\n美元/KG\n*每1g计重，限重2KG"
    sheet["E9"] = "挂号服务费\n美元/包裹"
    sheet["D10"] = "包裹重量为1-2000克"
    rows = [
        ("俄罗斯", "Russian Federation", "RU", 23.47, 0.22),
        ("乌克兰", "Ukraine", "UA", 23.48, 0.21),
        ("白俄罗斯", "Belarus", "BY", 19, 0.15),
        ("哈萨克斯坦", "Kazakhstan", "KZ", 14.14, 0.17),
        ("土耳其", "Turkey", "TR", 17.62, 0.15),
        ("芬兰", "Finland", "FI", 5.5, 3),
    ]
    for offset, item in enumerate(rows):
        for col, value in enumerate(item, start=1):
            sheet.cell(11 + offset, col, value)


def _weihai_sheet(sheet) -> None:
    sheet["A6"] = "运达国家/地区"
    sheet["B6"] = "运达国家/地区（英文）"
    sheet["C6"] = "Code"
    sheet["D6"] = "0-2000g（含）"
    sheet["F6"] = "最低计费重量"
    sheet["D7"] = "配送服务费\n元（RMB）/KG\n*每1g计重"
    sheet["E7"] = "挂号服务费\n元（RMB）/包裹"
    sheet["A8"] = "韩国"
    sheet["B8"] = "Korea"
    sheet["C8"] = "KR"
    sheet["D8"] = 49.47
    sheet["E8"] = 5.81
    sheet["F8"] = "10g"


def _sea_sheet(sheet) -> None:
    sheet["A5"] = "1. 计重规则，实际重量与体积重取大值，澳大利亚、美国最低计费重量5kg，不足5kg按5kg计费，墨西哥最低计费重量2kg，不足2kg按2kg计费；"
    headers = [
        "运达国家/地区",
        "运达国家/地区（英文）",
        "运达国家/地区（code）",
        "大小包",
        "是否计泡",
        "计泡系数",
        "货品属性",
        "计费模式",
        "始发区域",
        "计费币种",
        "价格起始重量(g)",
        "价格终止重量（含）(g)",
        "配送服务费重量进位制(g)",
        "配送服务费 /KG",
        "挂号服务费/包裹",
    ]
    for col, text in enumerate(headers, start=1):
        sheet.cell(9, col, text)
    rows = [
        ("澳大利亚", "Australia", "AU", "/", "是", 6000, "/", "配送服务费+挂号服务费", "/", "CNY", 5000, None, 1000, 22, 39),
        ("美国", "United States", "US", "/", "是", 8000, "/", "配送服务费+挂号服务费", "/", "CNY", 5000, 7000, 1000, 29, 24),
        ("美国", "United States", "US", "/", "是", 8000, "/", "配送服务费+挂号服务费", "/", "CNY", 7000, 10000, 1000, 27, 26),
        ("美国", "United States", "US", "/", "是", 8000, "/", "配送服务费+挂号服务费", "/", "CNY", 10000, 15000, 1000, 24, 26),
        ("墨西哥", "Mexico", "MX", "/", "是", 8000, "/", "配送服务费+挂号服务费", "/", "CNY", 2000, 5000, 1000, 36, 31),
        ("墨西哥", "Mexico", "MX", "/", "是", 8000, "/", "配送服务费+挂号服务费", "/", "CNY", 5000, 30000, 1000, 42, 32),
    ]
    for offset, item in enumerate(rows):
        for col, value in enumerate(item, start=1):
            if value is not None:
                sheet.cell(10 + offset, col, value)


def _oversized_sheet(sheet) -> None:
    headers = {
        2: "起始地省",
        3: "起始地市",
        5: "目的国",
        6: "计费克重段（单位：kg）",
        7: "计费重量阶梯（单位：1kg）",
        8: "计费/支付币种",
        9: "配送服务费\n元（RMB）/KG\n*每1kg计重",
        10: "挂号服务费\n元（RMB）/包裹",
    }
    for col, text in headers.items():
        sheet.cell(7, col, text)
    bands = [
        (0.001, 60, 6, 30),
        (60.001, 70, 6, 60),
        (70.001, 80, 6, 64),
        (80.001, 90, 6, 75),
        (90.001, 100, 6, 84),
        (100.001, 150, 19, 120),
        (150.001, 200, 24, 136),
    ]
    for offset, (start, end, rate, reg) in enumerate(bands):
        row = 8 + offset
        if offset == 0:
            sheet.cell(row, 2, "山东")
            sheet.cell(row, 3, "威海仓发")
            sheet.cell(row, 5, "韩国")
        sheet.cell(row, 6, f"{start}-{end}")
        sheet.cell(row, 7, "1kg")
        sheet.cell(row, 8, "RMB")
        sheet.cell(row, 9, rate)
        sheet.cell(row, 10, reg)
    sheet["B16"] = "计费重量：以体积重和实际重量的较大者为计费重，体积重计算方式为：长（CM）*宽（CM）*高（CM）/6000,计算后的单位为KG，不足1kg 的按1kg 计费"


def _third_party_sheet(sheet) -> None:
    sheet["A8"] = "计费方式："
    sheet["A9"] = NOTE
    sheet["A10"] = NOTE_GENERAL
    sheet["A11"] = NOTE_SPECIAL
    sheet["A12"] = NOTE_LARGE
    _write_standard_header(sheet, 13)
    _write_standard_country(
        sheet,
        16,
        ("法国", "France", "FR", 91.3, 17.02, 91.28, 17.02, 69.58, 22.02, 93.65, 17.02, 52.28, 70.71),
    )


def _save_template(path: Path, mode: str) -> None:
    book = Workbook()
    help_sheet = book.active
    help_sheet.title = "帮助页"
    help_sheet["A1"] = "AliExpress 运费模板使用说明"
    route = book.create_sheet("线路编码")
    countries = book.create_sheet("countries")
    modes = book.create_sheet("easy_freight_mode")
    promise = book.create_sheet("promised_delivery_time_solution")
    service = book.create_sheet("service_name")
    freight = book.create_sheet("freight_mode")
    route["A1"] = "无忧物流-标准"
    route["C1"] = "线路代号: CAINIAO_STANDARD"
    route["E1"] = "线路计费方式"
    route["F1"] = "自定义运费"
    route["G1"] = "承诺运达时间类型"
    route["H1"] = "平台标准配置"
    if mode == "quantity":
        route["I1"] = "数量运费配置"
        route["A2"] = "目的地"
        route["C2"] = "承诺运达时间（天）"
        route["E2"] = "运费计费规则(标准运费/卖家承担/按数量计费/按重量计费/不发货)"
        route["G2"] = "标准运费-减免百分比"
        route["I2"] = "起始采购数"
        route["J2"] = "截止采购数"
        route["K2"] = "首批采购运费 (¥)"
        route["L2"] = "每增加定额产品采购量"
        route["M2"] = "增加运费 (¥)"
        rule = "按数量计费"
    else:
        route["I1"] = "重量运费配置"
        route.merge_cells("I1:S1")
        route["A2"] = "目的地"
        route["C2"] = "承诺运达时间（天）"
        route["E2"] = "运费计费规则(标准运费/卖家承担/按重量计费/按重量计费/不发货)"
        route["G2"] = "标准运费-减免百分比"
        route["I2"] = "首重 (kg)"
        route["J2"] = "首重运费 (¥)"
        route["K2"] = "续重范围1 (kg)"
        route["L2"] = "续重范围1-每增加 (kg)"
        route["N2"] = "续重范围2 (kg)"
        route["O2"] = "续重范围2-每增加 (kg)"
        route["P2"] = "续重范围2-续加运费 (¥)"
        route["Q2"] = "续重范围3 (kg)"
        route["R2"] = "续重范围3-每增加 (kg)"
        route["S2"] = "续重范围3-续加运费 (¥)"
        rule = "按重量计费"
    row = 3
    seen = {name for name, _day in SELLER_ROWS}
    for name, day in SELLER_ROWS:
        route.cell(row, 1, name)
        route.cell(row, 3, day)
        route.cell(row, 5, "卖家承担")
        row += 1
    for name, day in QUANTITY_ROWS:
        route.cell(row, 1, name)
        route.cell(row, 3, day)
        route.cell(row, 5, rule)
        seen.add(name)
        row += 1
    for item in STANDARD_ROWS:
        if item[0] in seen:
            continue
        route.cell(row, 1, item[0])
        route.cell(row, 5, rule)
        row += 1
    for name in ("美国", "乌克兰"):
        if name not in seen:
            route.cell(row, 1, name)
            route.cell(row, 5, rule)
            row += 1
    service["A1"] = "CAINIAO_STANDARD"
    service["B1"] = "无忧物流-标准"
    for index, (code, label) in enumerate(
        [("allStandard", "标准运费"), ("custom", "自定义运费"), ("allFreeShipping", "卖家承担")],
        start=1,
    ):
        modes.cell(index, 1, code)
        modes.cell(index, 2, label)
    promise["A1"] = "standard"
    promise["B1"] = "平台标准配置"
    for index, (code, label) in enumerate(
        [("standard", "标准运费"), ("freeShipping", "卖家承担"), ("notShipping", "不发货"), ("weight", "按重量计费"), ("order_num", "按数量计费")],
        start=1,
    ):
        freight.cell(index, 1, code)
        freight.cell(index, 2, label)
    countries["A1"] = "AU"
    countries["B1"] = "澳大利亚"
    path.parent.mkdir(parents=True, exist_ok=True)
    book.save(path)


def build_all() -> None:
    build_rate_book()
    build_templates()


if __name__ == "__main__":
    build_all()
