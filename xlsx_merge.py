"""Склейка файлов xlsx частей одной проверки в общий файл.

«Элемент» не может держать общий файл в памяти: срабатывает предел размера объекта.
Поэтому он отправляет файлы частей по одному, а общий файл хранится здесь,
и пользователь скачивает его по ссылке.
"""

import io
import json
import os
import re
import secrets
import shutil
import time

MERGE_ROOT = os.getenv("XLSX_MERGE_DIR", "/tmp/xlsx_merge")
# Столько дней хранится сборка. При перезапуске сервиса папка /tmp очищается раньше
MERGE_KEEP_DAYS = 7
MAX_PARTS = 200
MAX_PART_BYTES = 20 * 1024 * 1024
MAX_MERGED_ROWS = 600_000
RESULT_FILE_NAME = "result.xlsx"
META_FILE_NAME = "meta.json"

_MERGE_ID_RE = re.compile(r"^[0-9a-f]{32}$")
_FILE_NAME_RE = re.compile(r"[^0-9A-Za-zА-Яа-яЁё._-]+")


class MergeError(Exception):
    def __init__(self, status_code: int, detail: str):
        # Оба аргумента в args: исключение должно пройти из дочернего процесса через pickle
        super().__init__(status_code, detail)
        self.status_code = status_code
        self.detail = detail


def merge_dir(merge_id: str) -> str:
    """Папка сборки. Номер проверяется, чтобы нельзя было выйти за пределы MERGE_ROOT."""
    if not _MERGE_ID_RE.match(merge_id or ""):
        raise MergeError(404, "Merge not found")
    return os.path.join(MERGE_ROOT, merge_id)


def remove_old_merges() -> None:
    if not os.path.isdir(MERGE_ROOT):
        return
    border_time = time.time() - MERGE_KEEP_DAYS * 24 * 60 * 60
    for merge_id in os.listdir(MERGE_ROOT):
        merge_path = os.path.join(MERGE_ROOT, merge_id)
        if os.path.isdir(merge_path) and os.path.getmtime(merge_path) < border_time:
            shutil.rmtree(merge_path, ignore_errors=True)


def create_merge(file_name: str) -> str:
    remove_old_merges()
    merge_id = secrets.token_hex(16)
    merge_path = os.path.join(MERGE_ROOT, merge_id)
    os.makedirs(merge_path)
    safe_name = _FILE_NAME_RE.sub("_", file_name or "").strip("._") or "result"
    if not safe_name.lower().endswith(".xlsx"):
        safe_name += ".xlsx"
    with open(os.path.join(merge_path, META_FILE_NAME), "w", encoding="utf-8") as meta_file:
        json.dump({"file_name": safe_name, "created": time.time()}, meta_file, ensure_ascii=False)
    return merge_id


def save_part(merge_id: str, part_number: int, part_bytes: bytes) -> None:
    merge_path = merge_dir(merge_id)
    if not os.path.isdir(merge_path):
        raise MergeError(404, "Merge not found")
    if not 1 <= part_number <= MAX_PARTS:
        raise MergeError(400, f"Part number must be from 1 to {MAX_PARTS}")
    if len(part_bytes) > MAX_PART_BYTES:
        raise MergeError(413, "Part file is too large")
    if not part_bytes.startswith(b"PK"):
        raise MergeError(400, "Part is not an xlsx file")
    with open(os.path.join(merge_path, f"part_{part_number:04d}.xlsx"), "wb") as part_file:
        part_file.write(part_bytes)


def read_meta(merge_id: str) -> dict:
    merge_path = merge_dir(merge_id)
    meta_path = os.path.join(merge_path, META_FILE_NAME)
    if not os.path.isfile(meta_path):
        raise MergeError(404, "Merge not found")
    with open(meta_path, encoding="utf-8") as meta_file:
        return json.load(meta_file)


def result_path(merge_id: str) -> str:
    merge_path = merge_dir(merge_id)
    file_path = os.path.join(merge_path, RESULT_FILE_NAME)
    if not os.path.isfile(file_path):
        raise MergeError(404, "File not found")
    return file_path


def merge_parts(merge_id: str, expected_parts: int, max_level: int) -> int:
    """Склеивает части по номеру. Шапка берётся из первой части, уровни группировки сохраняются.

    Запускается в дочернем процессе. Каждая часть открывается целиком (до 30 000 строк),
    общий файл пишется потоково. Возвращает число строк данных.
    """
    from openpyxl import Workbook, load_workbook
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.styles import PatternFill

    merge_path = merge_dir(merge_id)
    part_names = sorted(name for name in os.listdir(merge_path) if name.startswith("part_"))
    if len(part_names) != expected_parts:
        raise MergeError(409, f"Expected {expected_parts} parts, received {len(part_names)}")
    for index, part_name in enumerate(part_names, start=1):
        if part_name != f"part_{index:04d}.xlsx":
            raise MergeError(409, f"Part {index} is missing")

    first_sheet = load_workbook(os.path.join(merge_path, part_names[0]), read_only=True).active
    header_values = next(first_sheet.iter_rows(min_row=1, max_row=1, values_only=True))

    output_book = Workbook(write_only=True)
    output_sheet = output_book.create_sheet("Лист_1")
    output_sheet.sheet_properties.outlinePr.summaryBelow = False
    # Глубину группировки передаёт «Элемент»: потоковый лист пишет её до первой строки
    output_sheet.sheet_format.outlineLevelRow = max(0, min(7, max_level))

    first_book = load_workbook(os.path.join(merge_path, part_names[0]))
    first_active = first_book.active
    for column_letter, column_dimension in first_active.column_dimensions.items():
        if column_dimension.width:
            output_sheet.column_dimensions[column_letter].width = column_dimension.width
    header_fill = PatternFill("solid", fgColor="F2F2F2")
    output_sheet.append([_header_cell(output_sheet, value, header_fill, WriteOnlyCell) for value in header_values])
    first_book.close()

    output_row = 1
    for part_name in part_names:
        part_book = load_workbook(os.path.join(merge_path, part_name))
        part_sheet = part_book.active
        for row_index, row_cells in enumerate(part_sheet.iter_rows(), start=1):
            row_values = tuple(cell.value for cell in row_cells)
            if row_index == 1:
                if row_values != header_values:
                    raise MergeError(409, f"{part_name}: header differs from part 1")
                continue
            output_row += 1
            if output_row - 1 > MAX_MERGED_ROWS:
                raise MergeError(413, f"Too many rows: limit {MAX_MERGED_ROWS}")
            row_dimension = part_sheet.row_dimensions.get(row_index)
            if row_dimension is not None and row_dimension.outline_level:
                output_sheet.row_dimensions[output_row].outline_level = row_dimension.outline_level
                output_sheet.row_dimensions[output_row].hidden = bool(row_dimension.hidden)
            output_cells = []
            for cell_value in row_values:
                output_cell = WriteOnlyCell(output_sheet, value=cell_value)
                if isinstance(cell_value, str):
                    # Значение уже записано так, как его хранит Excel (_x001D_): повторно не экранируется
                    output_cell.number_format = "@"
                output_cells.append(output_cell)
            output_sheet.append(output_cells)
            if output_row in output_sheet.row_dimensions:
                del output_sheet.row_dimensions[output_row]
        part_book.close()

    output_bytes = io.BytesIO()
    output_book.save(output_bytes)
    with open(os.path.join(merge_path, RESULT_FILE_NAME), "wb") as result_file:
        result_file.write(output_bytes.getvalue())
    return output_row - 1


def _header_cell(sheet, value, fill, cell_type):
    header_cell = cell_type(sheet, value=value)
    header_cell.fill = fill
    return header_cell
