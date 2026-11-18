"""Dacha Vision Draw: рисует рамки обнаруженных объектов на фото.

Координаты рамки приходят в шкале 0-1000, как их возвращает Gemini:
box_2d = [ymin, xmin, ymax, xmax].
"""
import base64
import binascii
import hmac
import io
import multiprocessing
import os
import threading
from concurrent.futures import ProcessPoolExecutor
from typing import List, Tuple

from fastapi import FastAPI, Header, HTTPException, Request
from fastapi.responses import FileResponse
from fastapi.concurrency import run_in_threadpool
from PIL import Image, ImageDraw, ImageFont, UnidentifiedImageError
from pydantic import BaseModel, ConfigDict, Field

from xlsx_builder import XlsxBuildError, build_xlsx_from_json
import xlsx_merge
from xlsx_merge import MergeError

app = FastAPI()

DRAW_TOKEN = os.getenv("DRAW_TOKEN", "")

# Ограничение тела запроса /xlsx: 600 000 строк занимают около 52 МБ (замер).
# Запас нужен на длинные коды и на JSON с отступами.
XLSX_MAX_BODY_BYTES = 100 * 1024 * 1024
# Не больше двух сборок одновременно: каждая сборка — отдельный процесс с памятью
XLSX_PARALLEL_BUILDS = threading.Semaphore(2)
SPAWN_CONTEXT = multiprocessing.get_context("spawn")
FONT_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

# Цвет рамки по классу объекта
CLASS_COLORS = {
    "person": (230, 40, 40),   # человек — красный
    "cat": (255, 140, 0),      # кошка — оранжевый
    "dog": (30, 110, 240),     # собака — синий
}
DEFAULT_COLOR = (0, 180, 0)


class Box(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    label: str = ""
    object_class: str = Field("", alias="class")
    box_2d: List[float] = Field(..., min_length=4, max_length=4)


class DrawRequest(BaseModel):
    image: str
    boxes: List[Box] = []


class DrawResponse(BaseModel):
    image: str


class XlsxResponse(BaseModel):
    file: str
    rows: int


def check_token(authorization: str) -> None:
    if not DRAW_TOKEN:
        raise HTTPException(status_code=503, detail="DRAW_TOKEN is not configured")
    expected = f"Bearer {DRAW_TOKEN}"
    if not hmac.compare_digest(authorization or "", expected):
        raise HTTPException(status_code=401, detail="Invalid token")


def load_font(size: int) -> ImageFont.ImageFont:
    try:
        return ImageFont.truetype(FONT_PATH, size)
    except OSError:
        # Локальный запуск без шрифта DejaVu: кириллица может не отобразиться
        return ImageFont.load_default()


def to_pixels(box_2d: List[float], width: int, height: int):
    ymin, xmin, ymax, xmax = (min(max(value, 0.0), 1000.0) for value in box_2d)
    left, right = sorted((xmin / 1000 * width, xmax / 1000 * width))
    top, bottom = sorted((ymin / 1000 * height, ymax / 1000 * height))
    return round(left), round(top), round(right), round(bottom)


def draw_boxes(image: Image.Image, boxes: List[Box]) -> Image.Image:
    image = image.convert("RGB")
    width, height = image.size
    line_width = max(2, round(min(width, height) / 150))
    font = load_font(max(12, round(min(width, height) / 30)))
    painter = ImageDraw.Draw(image)

    for box in boxes:
        color = CLASS_COLORS.get(box.object_class.lower(), DEFAULT_COLOR)
        left, top, right, bottom = to_pixels(box.box_2d, width, height)
        painter.rectangle((left, top, right, bottom), outline=color, width=line_width)

        if not box.label:
            continue
        text_left, text_top, text_right, text_bottom = painter.textbbox((0, 0), box.label, font=font)
        text_width = text_right - text_left
        text_height = text_bottom - text_top
        padding = line_width
        label_height = text_height + 2 * padding
        # Подпись над рамкой; если места сверху нет — внутри рамки
        label_top = top - label_height if top >= label_height else top
        label_right = min(left + text_width + 2 * padding, width)
        painter.rectangle((left, label_top, label_right, label_top + label_height), fill=color)
        painter.text((left + padding, label_top + padding - text_top), box.label, fill=(255, 255, 255), font=font)

    return image


@app.get("/ping")
async def ping():
    return {"status": "ok", "service": "Dacha_Vision_Draw", "active": True}


@app.post("/draw", response_model=DrawResponse)
async def draw(request: DrawRequest, authorization: str = Header("")):
    check_token(authorization)

    try:
        source_bytes = base64.b64decode(request.image, validate=True)
        source_image = Image.open(io.BytesIO(source_bytes))
        source_image.load()
    except (binascii.Error, UnidentifiedImageError, OSError, ValueError):
        raise HTTPException(status_code=400, detail="Invalid image")

    result_image = draw_boxes(source_image, request.boxes)
    output = io.BytesIO()
    result_image.save(output, format="JPEG", quality=90)
    return DrawResponse(image=base64.b64encode(output.getvalue()).decode("ascii"))


def build_xlsx_in_child_process(request_body: bytes) -> Tuple[bytes, int]:
    """Разбирает запрос и собирает файл в отдельном процессе.

    Основной процесс не разбирает JSON на десятки тысяч строк: он держит только байты тела.
    Дочерний процесс завершается после сборки, и система забирает его память.
    Railway берёт плату за занятую память, поэтому основной процесс остаётся лёгким.
    """
    with XLSX_PARALLEL_BUILDS:
        with ProcessPoolExecutor(max_workers=1, mp_context=SPAWN_CONTEXT) as build_pool:
            return build_pool.submit(build_xlsx_from_json, request_body).result()


@app.post("/xlsx", response_model=XlsxResponse)
async def xlsx(request: Request, authorization: str = Header("")):
    check_token(authorization)

    request_body = await request.body()
    if len(request_body) > XLSX_MAX_BODY_BYTES:
        raise HTTPException(status_code=413, detail="Request body is too large")

    try:
        # Сборка ждёт дочерний процесс в отдельном потоке и не блокирует другие запросы
        file_bytes, rows_count = await run_in_threadpool(build_xlsx_in_child_process, request_body)
    except XlsxBuildError as build_error:
        raise HTTPException(status_code=build_error.status_code, detail=build_error.detail)

    return XlsxResponse(file=base64.b64encode(file_bytes).decode("ascii"), rows=rows_count)


# ---------- Склейка файлов частей ----------


class MergeCreateRequest(BaseModel):
    file_name: str = "result.xlsx"


class MergeCreateResponse(BaseModel):
    id: str


class MergePartRequest(BaseModel):
    number: int
    file: str


class MergeFinishRequest(BaseModel):
    parts: int
    max_level: int = 0


class MergeFinishResponse(BaseModel):
    path: str
    rows: int


def merge_error(error: MergeError) -> HTTPException:
    return HTTPException(status_code=error.status_code, detail=error.detail)


def merge_in_child_process(merge_id: str, expected_parts: int, max_level: int) -> int:
    """Склейка идёт в отдельном процессе, как сборка /xlsx: память основного процесса не растёт."""
    with XLSX_PARALLEL_BUILDS:
        with ProcessPoolExecutor(max_workers=1, mp_context=SPAWN_CONTEXT) as merge_pool:
            return merge_pool.submit(xlsx_merge.merge_parts, merge_id, expected_parts, max_level).result()


@app.post("/xlsx/merge", response_model=MergeCreateResponse)
async def merge_create(request: MergeCreateRequest, authorization: str = Header("")):
    check_token(authorization)
    return MergeCreateResponse(id=xlsx_merge.create_merge(request.file_name))


@app.post("/xlsx/merge/{merge_id}/part")
async def merge_part(merge_id: str, request: MergePartRequest, authorization: str = Header("")):
    check_token(authorization)
    try:
        part_bytes = base64.b64decode(request.file, validate=True)
    except (binascii.Error, ValueError):
        raise HTTPException(status_code=400, detail="Part file is not valid Base64")
    try:
        xlsx_merge.save_part(merge_id, request.number, part_bytes)
    except MergeError as error:
        raise merge_error(error)
    return {"number": request.number}


@app.post("/xlsx/merge/{merge_id}/finish", response_model=MergeFinishResponse)
async def merge_finish(merge_id: str, request: MergeFinishRequest, authorization: str = Header("")):
    check_token(authorization)
    try:
        xlsx_merge.read_meta(merge_id)
        rows_count = await run_in_threadpool(merge_in_child_process, merge_id, request.parts, request.max_level)
    except MergeError as error:
        raise merge_error(error)
    return MergeFinishResponse(path=f"/xlsx/merge/{merge_id}/file", rows=rows_count)


@app.get("/xlsx/merge/{merge_id}/file")
async def merge_file(merge_id: str):
    """Скачивание общего файла без токена: доступ даёт случайный номер сборки из 32 знаков."""
    try:
        meta = xlsx_merge.read_meta(merge_id)
        file_path = xlsx_merge.result_path(merge_id)
    except MergeError as error:
        raise merge_error(error)
    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=meta["file_name"],
    )
