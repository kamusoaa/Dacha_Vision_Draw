"""Dacha Vision Draw: рисует рамки обнаруженных объектов на фото.

Координаты рамки приходят в шкале 0-1000, как их возвращает Gemini:
box_2d = [ymin, xmin, ymax, xmax].
"""
import base64
import binascii
import hmac
import io
import os
from typing import List

from fastapi import FastAPI, Header, HTTPException
from PIL import Image, ImageDraw, ImageFont, UnidentifiedImageError
from pydantic import BaseModel, ConfigDict, Field

app = FastAPI()

DRAW_TOKEN = os.getenv("DRAW_TOKEN", "")
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
