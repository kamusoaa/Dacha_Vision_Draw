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
# Строк в одном файле. Файл умещается на одном листе Excel (не больше 1 048 575 строк данных)
MAX_ROWS = 600_000


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
    """Собирает файл xlsx из колонок, строк и уровня каждой строки.

    Книга работает в потоковом режиме: строка сразу пишется в файл и не хранится в памяти.
    Уровень группировки строки задаётся до записи строки и удаляется после неё.
    Замер: 1 000 000 строк — около 50 МБ памяти против 723 МБ на 200 000 строк в обычном режиме.
    """
    from openpyxl import Workbook
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.styles import PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.exceptions import IllegalCharacterError

    columns = xlsx_request.columns
    workbook = Workbook(write_only=True)
    try:
        sheet = workbook.create_sheet(xlsx_request.sheet)
    except ValueError:
        raise XlsxBuildError(400, f"Invalid sheet name: {xlsx_request.sheet!r}")
    # Итог группы стоит над вложенными строками, как в образце
    sheet.sheet_properties.outlinePr.summaryBelow = False
    # Глубина группировки задаётся до первой строки: потоковый лист пишет её в начало файла
    max_level = max((row.level for row in xlsx_request.rows), default=0)
    sheet.sheet_format.outlineLevelRow = max_level

    # Ширина колонок задаётся до первой строки: потоковый лист пишет её в начало файла
    for column_index, column in enumerate(columns, start=1):
        if column.width:
            sheet.column_dimensions[get_column_letter(column_index)].width = column.width

    header_fill = PatternFill("solid", fgColor=xlsx_request.header_fill)
    header_cells = []
    for column in columns:
        header_cell = WriteOnlyCell(sheet, value=column.title)
        header_cell.fill = header_fill
        header_cells.append(header_cell)
    sheet.append(header_cells)

    number_columns = {index for index, column in enumerate(columns) if column.kind == "number"}

    for row_number, row in enumerate(xlsx_request.rows, start=2):
        row_cells = []
        for index, cell_value in enumerate(row.values):
            if index in number_columns:
                row_cells.append(to_number(cell_value))
                continue
            try:
                text_cell = WriteOnlyCell(sheet, value=cell_value)
            except IllegalCharacterError:
                raise XlsxBuildError(400, f"Row {row_number - 1}: illegal character in value")
            # Формат «Текст»: Excel не превратит код с нулями в число
            text_cell.number_format = "@"
            row_cells.append(text_cell)

        if row.level > 0:
            sheet.row_dimensions[row_number].outline_level = row.level
            sheet.row_dimensions[row_number].hidden = xlsx_request.collapsed

        sheet.append(row_cells)
        # Строка уже записана в файл: её размеры больше не нужны и не должны копиться в памяти
        if row_number in sheet.row_dimensions:
            del sheet.row_dimensions[row_number]

    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()
