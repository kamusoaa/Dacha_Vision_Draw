"""Сборка файла Excel с группировкой строк для адреса /xlsx.

Весь модуль работает в отдельном процессе: разбор JSON, проверка запроса и сборка файла.
Процесс завершается после сборки, и система забирает его память.
openpyxl импортируется внутри функции: основной процесс сервиса его не загружает.
"""
import io
from typing import List, Literal, Optional, Tuple

from pydantic import BaseModel, Field, ValidationError

# Excel поддерживает не больше 7 уровней группировки строк
MAX_OUTLINE_LEVEL = 7
MAX_ROWS = 200_000


class XlsxBuildError(Exception):
    """Ошибка запроса. Основной процесс превращает её в HTTP-ответ."""

    def __init__(self, status_code: int, detail: str):
        super().__init__(status_code, detail)
        self.status_code = status_code
        self.detail = detail


class XlsxColumn(BaseModel):
    title: str
    width: Optional[float] = Field(None, gt=0, le=255)
    kind: Literal["text", "number"] = "text"


class XlsxRow(BaseModel):
    level: int = Field(0, ge=0, le=MAX_OUTLINE_LEVEL)
    values: List[str]


class XlsxRequest(BaseModel):
    sheet: str = Field("Лист_1", min_length=1, max_length=31)
    header_fill: str = Field("F2F2F2", pattern=r"^[0-9A-Fa-f]{6}$")
    collapsed: bool = True
    columns: List[XlsxColumn] = Field(..., min_length=1)
    rows: List[XlsxRow] = []


def to_number(text_value: str):
    """Превращает строку в число. Пустая строка даёт пустую ячейку."""
    if text_value == "":
        return None
    try:
        return int(text_value)
    except ValueError:
        pass
    try:
        return float(text_value)
    except ValueError:
        raise XlsxBuildError(400, f"Not a number: {text_value!r}")


def parse_request(request_body: bytes) -> XlsxRequest:
    """Разбирает и проверяет тело запроса."""
    try:
        xlsx_request = XlsxRequest.model_validate_json(request_body)
    except ValidationError as validation_error:
        first_error = validation_error.errors()[0]
        error_place = ".".join(str(part) for part in first_error["loc"])
        error_text = f"{error_place}: {first_error['msg']}" if error_place else first_error["msg"]
        raise XlsxBuildError(422, error_text)

    if len(xlsx_request.rows) > MAX_ROWS:
        raise XlsxBuildError(413, f"Too many rows: limit {MAX_ROWS}")

    columns_count = len(xlsx_request.columns)
    for row_number, row in enumerate(xlsx_request.rows, start=1):
        if len(row.values) != columns_count:
            raise XlsxBuildError(400, f"Row {row_number}: {len(row.values)} values, expected {columns_count}")

    return xlsx_request


def build_xlsx_from_json(request_body: bytes) -> Tuple[bytes, int]:
    """Точка входа дочернего процесса: тело запроса на входе, файл xlsx и число строк на выходе."""
    xlsx_request = parse_request(request_body)
    return build_xlsx(xlsx_request), len(xlsx_request.rows)


def build_xlsx(xlsx_request: XlsxRequest) -> bytes:
    """Собирает файл xlsx из колонок, строк и уровня каждой строки."""
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.exceptions import IllegalCharacterError

    columns = xlsx_request.columns
    workbook = Workbook()
    sheet = workbook.active
    try:
        sheet.title = xlsx_request.sheet
    except ValueError:
        raise XlsxBuildError(400, f"Invalid sheet name: {xlsx_request.sheet!r}")
    # Итог группы стоит над вложенными строками, как в образце
    sheet.sheet_properties.outlinePr.summaryBelow = False

    header_fill = PatternFill("solid", fgColor=xlsx_request.header_fill)
    sheet.append([column.title for column in columns])
    for header_cell in sheet[1]:
        header_cell.fill = header_fill

    for column_index, column in enumerate(columns, start=1):
        if column.width:
            sheet.column_dimensions[get_column_letter(column_index)].width = column.width

    number_columns = {index for index, column in enumerate(columns) if column.kind == "number"}
    max_level = 0

    for row_number, row in enumerate(xlsx_request.rows, start=2):
        row_values = [
            to_number(cell_value) if index in number_columns else cell_value
            for index, cell_value in enumerate(row.values)
        ]
        try:
            sheet.append(row_values)
        except IllegalCharacterError:
            raise XlsxBuildError(400, f"Row {row_number - 1}: illegal character in value")

        for index in range(len(columns)):
            if index not in number_columns:
                # Формат «Текст»: Excel не превратит код с нулями в число
                sheet.cell(row_number, index + 1).number_format = "@"

        if row.level > 0:
            sheet.row_dimensions[row_number].outline_level = row.level
            sheet.row_dimensions[row_number].hidden = xlsx_request.collapsed
            max_level = max(max_level, row.level)

    sheet.sheet_format.outlineLevelRow = max_level

    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()
