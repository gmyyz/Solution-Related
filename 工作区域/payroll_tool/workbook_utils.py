import re
from copy import copy
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP


ERROR_FONT_COLOR = 'FFFF0000'
CENT = Decimal('0.01')


def _is_blank(value):
    return value is None or (isinstance(value, str) and value.strip() == '')


def _normalize_text(value):
    if value is None:
        return ''
    text = str(value).replace('\r', ' ').replace('\n', ' ').replace('\t', ' ')
    text = text.replace('\u3000', ' ').replace('\xa0', ' ')
    return ' '.join(text.split()).strip()


def _is_total_row(value):
    return _normalize_text(value) == '总计'


def _is_blank_project(value):
    text = _normalize_text(value)
    blank_texts = {
        '',
        '空白',
        '(空白)',
        '（空白）',
        '(空 白)',
        '（空 白）',
    }
    return text in blank_texts


def _to_decimal(value):
    if value is None or value == '':
        return Decimal('0')
    if isinstance(value, Decimal):
        return value
    return Decimal(str(value))


def _to_money(value):
    return _to_decimal(value).quantize(CENT, rounding=ROUND_HALF_UP)


def _to_text_money(value):
    return f'{_to_money(value):.2f}'


def _format_code(value):
    if value is None:
        return ''
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value)
    return _normalize_text(value)


def _extract_year_month(value):
    if isinstance(value, (datetime, date)):
        return int(value.year), int(value.month)

    text = _normalize_text(value)
    if not text:
        return None

    match = re.search(r'(\d{4})[-/年](\d{1,2})', text)
    if match:
        return int(match.group(1)), int(match.group(2))

    match = re.search(r'(\d{4})(\d{2})', text)
    if match:
        return int(match.group(1)), int(match.group(2))

    return None


def _copy_cell_style(source_cell, target_cell):
    target_cell._style = copy(source_cell._style)


def _copy_font_red(cell):
    font = copy(cell.font)
    font.color = ERROR_FONT_COLOR
    cell.font = font


def _mark_row_red(ws, row_idx, max_col):
    for col_idx in range(1, max_col + 1):
        _copy_font_red(ws.cell(row=row_idx, column=col_idx))


def _mark_range_red(ws, row_idx, start_col, end_col):
    for col_idx in range(start_col, end_col + 1):
        _copy_font_red(ws.cell(row=row_idx, column=col_idx))


def _find_total_row(ws):
    for row_idx in range(1, ws.max_row + 1):
        if _is_total_row(ws.cell(row=row_idx, column=1).value):
            return row_idx
    return None


def _find_payroll_header_row(ws):
    expected_headers = ('公司', '部门', '项目', '类别')
    for row_idx in range(1, min(ws.max_row, 5) + 1):
        actual_headers = tuple(_normalize_text(ws.cell(row=row_idx, column=col_idx).value) for col_idx in range(1, 5))
        if actual_headers == expected_headers:
            return row_idx
    raise ValueError('未能识别工资单表头：前 5 行中未找到“公司、部门、项目、类别”')


def _recalculate_total_row(ws, total_row_idx, data_start_row=3):
    ws.cell(row=total_row_idx, column=1).value = '总计'
    for col_idx in range(2, 21):
        ws.cell(row=total_row_idx, column=col_idx).value = None

    for col_idx in range(5, 19):
        total_value = sum(
            (_to_decimal(ws.cell(row=row_idx, column=col_idx).value) for row_idx in range(data_start_row, total_row_idx)),
            Decimal('0'),
        )
        ws.cell(row=total_row_idx, column=col_idx).value = float(
            total_value.quantize(CENT, rounding=ROUND_HALF_UP)
        )
