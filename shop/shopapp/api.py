"""电商工具接口。运费计算是速卖通下的一项功能。"""

from __future__ import annotations

import io
import re
import zipfile
from decimal import Decimal
from urllib.parse import quote as urlquote

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

from shop.shopapp.catalog import catalog_payload, platform_by_id, tool_by_id
from shop.shopapp.models import Costs
from shop.shopapp.money import money
from shop.shopapp.platforms import get_platform
from shop.shopapp.products import MISSING, product_from_fields, products_from_excel
from shop.shopapp.store import DATA, RateStore
from shop.shopapp.translate.engine import translate_text
from shop.shopapp.translate.languages import require_source, require_target, source_options, target_options
from shop.shopapp.translate.video import result_path, translate_upload

app = FastAPI(title="电商工具")
STORE = RateStore()
STATIC = DATA.parent / "static"


@app.get("/health")
def health():
    return {"ok": True}


@app.get("/api/catalog")
def catalog():
    return catalog_payload()


@app.get("/api/translate/languages")
def translate_languages():
    return {"source": source_options(), "target": target_options()}


@app.post("/api/translate/text")
def translate_text_route(
    source: str = Form("auto"),
    target: str = Form("en"),
    text: str = Form(""),
):
    try:
        source_code = require_source(source)
        target_code = require_target(target)
        if not text.strip():
            raise ValueError("请输入要翻译的文字")
        translation, detected = translate_text(text, source_code, target_code)
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from exc
    used = detected if source_code == "auto" else source_code
    return {"source": used, "target": target_code, "translation": translation}


@app.post("/api/translate/video")
async def translate_video_route(
    source: str = Form("auto"),
    target: str = Form("en"),
    file: UploadFile = File(...),
):
    try:
        source_code = require_source(source)
        target_code = require_target(target)
        payload = await file.read()
        return translate_upload(payload, file.filename or "video.mp4", source_code, target_code)
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from exc


@app.get("/api/translate/video/{video_id}")
def translated_video(video_id: str):
    if re.fullmatch(r"[0-9a-f]{32}", video_id) is None:
        raise HTTPException(404, "没有这个视频")
    path = result_path(video_id)
    if path is None:
        raise HTTPException(404, "没有这个视频")
    return FileResponse(path, media_type="video/mp4", filename="译文字幕视频.mp4")


@app.get("/api/meta")
def meta():
    platform_id, book = STORE.get("default")
    return {
        "platforms": [{"id": item["id"], "name": item["name"]} for item in catalog_payload()["platforms"]],
        "rate_id": "default",
        "source_name": book.source_name,
        "channels": _channel_payload(platform_id, book),
        "costs": {"platform": 8, "affiliate": 10, "coin": 5, "one_stop": 2, "ads": 0, "loss": 10},
    }


@app.post("/api/rates")
async def upload_rates(
    platform: str = Form("aliexpress"),
    file: UploadFile = File(...),
):
    _require_freight(platform)
    payload = await file.read()
    if not payload:
        raise HTTPException(400, "请上传运费表")
    try:
        rate_id = STORE.add(platform, payload, file.filename or "rates.xlsx")
    except KeyError as exc:
        raise HTTPException(400, str(exc)) from exc
    except Exception as exc:
        raise HTTPException(400, f"运费表没有读成功：{exc}") from exc
    platform_id, book = STORE.get(rate_id)
    return {"rate_id": rate_id, "source_name": book.source_name, "channels": _channel_payload(platform_id, book)}


@app.post("/api/quote")
async def quote(
    platform: str = Form("aliexpress"),
    rate_id: str = Form("default"),
    channels: list[str] = Form(default=[]),
    cargo: str = Form("general"),
    sku: str = Form(""),
    spec: str = Form(""),
    weight: str = Form(""),
    products_file: UploadFile | None = File(None),
    platform_fee: str = Form(""),
    affiliate_fee: str = Form(""),
    coin_fee: str = Form(""),
    one_stop_fee: str = Form(""),
    ads_fee: str = Form(""),
    loss_fee: str = Form(""),
):
    _require_freight(platform)
    try:
        platform_id, book, products, costs = await _inputs(
            platform,
            rate_id,
            sku,
            spec,
            weight,
            products_file,
            platform_fee,
            affiliate_fee,
            coin_fee,
            one_stop_fee,
            ads_fee,
            loss_fee,
        )
        names = _channels(book, channels)
        quotes = get_platform(platform_id).quote(book, products, names, cargo, costs)
    except HTTPException:
        raise
    except KeyError as exc:
        raise HTTPException(400, str(exc)) from exc
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from exc
    return {
        "multiplier": format(costs.multiplier(), "f"),
        "products": [_product_payload(item) for item in quotes],
    }


@app.post("/api/export")
async def export_file(
    platform: str = Form("aliexpress"),
    rate_id: str = Form("default"),
    channels: list[str] = Form(default=[]),
    cargo: str = Form("general"),
    sku: str = Form(""),
    spec: str = Form(""),
    weight: str = Form(""),
    products_file: UploadFile | None = File(None),
    mode: str = Form("quantity"),
    qty_start: str = Form(""),
    qty_end: str = Form(""),
    qty_step: str = Form("1"),
    platform_fee: str = Form(""),
    affiliate_fee: str = Form(""),
    coin_fee: str = Form(""),
    one_stop_fee: str = Form(""),
    ads_fee: str = Form(""),
    loss_fee: str = Form(""),
):
    if mode not in {"quantity", "weight"}:
        raise HTTPException(400, "请选择按数量计费或按重量计费")
    _require_freight(platform)
    try:
        platform_id, book, products, costs = await _inputs(
            platform,
            rate_id,
            sku,
            spec,
            weight,
            products_file,
            platform_fee,
            affiliate_fee,
            coin_fee,
            one_stop_fee,
            ads_fee,
            loss_fee,
        )
        names = _channels(book, channels)
        start = _optional_int(qty_start)
        end = _optional_int(qty_end)
        step = _optional_int(qty_step) or 1
        files = get_platform(platform_id).export(
            book, products, names, cargo, costs, mode, start, end, step
        )
    except HTTPException:
        raise
    except KeyError as exc:
        raise HTTPException(400, str(exc)) from exc
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from exc
    warnings = [item for file in files for item in file.warnings]
    if len(files) == 1:
        return _download(files[0].filename, files[0].content, warnings)
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for file in files:
            archive.writestr(file.filename, file.content)
    weight_text = _weight_label(products)
    return _download(f"运费模板_{weight_text}.zip", buffer.getvalue(), warnings)


@app.get("/")
def index():
    return FileResponse(STATIC / "index.html")


def _channel_payload(platform_id: str, book):
    codes = {}
    code_path = DATA / "line_codes.json"
    if code_path.exists():
        import json

        codes = json.loads(code_path.read_text(encoding="utf-8"))
    return [
        {
            "name": channel.name,
            "countries": channel.country_count,
            "currencies": sorted(channel.currencies),
            "line_code": codes.get(channel.name, ""),
            "warning": channel.warning,
        }
        for channel in book.channels
    ]


async def _inputs(platform, rate_id, sku, spec, weight, products_file, platform_fee, affiliate_fee, coin_fee, one_stop_fee, ads_fee, loss_fee):
    try:
        platform_id, book = STORE.get(rate_id)
    except KeyError as exc:
        raise HTTPException(400, str(exc)) from exc
    if platform and platform != platform_id:
        raise HTTPException(400, "这份运费表不属于所选平台")
    products = []
    if products_file is not None and products_file.filename:
        payload = await products_file.read()
        if payload:
            try:
                products = products_from_excel(payload)
            except ValueError as exc:
                raise HTTPException(400, str(exc)) from exc
    if not products:
        if not sku.strip() and not spec.strip() and not weight.strip():
            raise HTTPException(400, MISSING)
        try:
            products = [product_from_fields(sku, spec, weight)]
        except ValueError as exc:
            message = str(exc)
            if not sku.strip() or not spec.strip() or not weight.strip():
                message = MISSING
            raise HTTPException(400, message) from exc
    costs = Costs(
        platform=_percent(platform_fee, "0.08"),
        affiliate=_percent(affiliate_fee, "0.10"),
        coin=_percent(coin_fee, "0.05"),
        one_stop=_percent(one_stop_fee, "0.02"),
        ads=_percent(ads_fee, "0"),
        loss=_percent(loss_fee, "0.10"),
    )
    return platform_id, book, products, costs


def _require_freight(platform: str) -> None:
    spec = platform_by_id(platform or "aliexpress")
    if spec is None:
        raise HTTPException(400, "没有这个平台")
    tool = tool_by_id(spec, "freight")
    if tool is None or not tool.ready:
        raise HTTPException(400, f"{spec.name}的运费计算还没接上")


def _channels(book, channels: list[str]) -> list[str]:
    names = [name for name in channels if name]
    if not names:
        raise HTTPException(400, "请至少选择一个运送渠道")
    return names


def _percent(raw: str, default: str) -> Decimal:
    text = (raw or "").strip().replace("%", "")
    if text == "":
        return Decimal(default)
    return Decimal(text) / Decimal("100")


def _optional_int(raw: str) -> int | None:
    text = (raw or "").strip()
    if text == "":
        return None
    try:
        return int(text)
    except ValueError as exc:
        raise HTTPException(400, "采购数要填整数") from exc


def _product_payload(item):
    product = item.product
    return {
        "sku": product.sku,
        "spec": f"{_num(product.length_cm)}x{_num(product.width_cm)}x{_num(product.height_cm)}",
        "weight_g": _num(product.weight_g),
        "warnings": item.warnings,
        "rows": [
            {
                "country": row.country,
                "code": row.code,
                "channel": row.channel,
                "currency": row.currency,
                "shippable": row.shippable,
                "detail": row.detail,
                "billable_g": None if row.billable_g is None else _num(row.billable_g),
                "rate_per_kg": _cash(row.rate_per_kg),
                "registration": _cash(row.registration),
                "first_fee": _cash(row.first_fee),
                "additional_fee": _cash(row.additional_fee),
                "freight": _cash(row.freight),
                "final_freight": _cash(row.final_freight),
                "extra_freight": _cash(row.extra_freight),
                "mode": row.mode,
            }
            for row in item.rows
        ],
    }


def _cash(value):
    if value is None:
        return None
    return format(money(value), "f")


def _num(value: Decimal) -> str:
    if value == value.to_integral_value():
        return str(int(value))
    return format(value.normalize(), "f")


def _weight_label(products) -> str:
    weights = []
    for product in products:
        text = str(int(product.weight_g)) if product.weight_g == product.weight_g.to_integral_value() else str(product.weight_g)
        if text not in weights:
            weights.append(text)
    return "_".join(f"{item}g" for item in weights[:4])


def _download(filename: str, content: bytes, warnings: list[str]) -> Response:
    media = "application/zip" if filename.endswith(".zip") else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    header = urlquote(filename)
    warning_header = urlquote(" | ".join(warnings))
    return Response(
        content,
        media_type=media,
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{header}",
            "X-Export-Warnings": warning_header,
        },
    )


@app.exception_handler(HTTPException)
async def http_error(_request, exc: HTTPException):
    if isinstance(exc.detail, str):
        return JSONResponse({"error": exc.detail}, status_code=exc.status_code)
    return JSONResponse({"error": "请求没有完成"}, status_code=exc.status_code)


app.mount("/static", StaticFiles(directory=STATIC), name="static")
