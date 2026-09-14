"""payroll_output: extracted payroll responsibilities; business rules preserved."""
import io
from .workbook_io import load_workbook, workbook_scope
from .workbook_utils import (
    _copy_cell_style,
    _find_payroll_header_row,
    _find_total_row,
    _format_code,
    _is_blank,
    _is_blank_project,
    _is_total_row,
    _mark_range_red,
    _mark_row_red,
    _normalize_text,
    _recalculate_total_row,
    _to_decimal,
    _to_money,
    _to_text_money,
)
from .paths import (
    _build_company_output_path,
)


def _write_bank_summary(ws, bank_results, header_ref):
    start_col = 22  # V列
    headers = ['公司', '核对事项', '工资表金额', '银行流水金额', '差额', '核对月份', '结果', '说明']
    widths = [12, 12, 14, 14, 14, 12, 10, 42]

    for idx, header in enumerate(headers):
        cell = ws.cell(row=2, column=start_col + idx)
        cell.value = header
        _copy_cell_style(header_ref, cell)
        ws.column_dimensions[cell.column_letter].width = widths[idx]

    text_ref = ws['A3']
    num_ref = ws['Q3']
    for row_offset, result in enumerate(bank_results, start=3):
        values = [
            result['company'],
            result['item_label'],
            result['payroll_amount'],
            result['bank_amount'],
            result['diff'],
            result['period_label'],
            '通过' if result['passed'] else '异常',
            result['note'],
        ]

        for col_offset, value in enumerate(values):
            cell = ws.cell(row=row_offset, column=start_col + col_offset)
            cell.value = value
            if col_offset in (2, 3, 4):
                _copy_cell_style(num_ref, cell)
            else:
                _copy_cell_style(text_ref, cell)

        if not result['passed']:
            _mark_range_red(ws, row_offset, start_col, start_col + len(headers) - 1)


def _clear_bank_summary(ws, end_row=12):
    for row_idx in range(2, end_row + 1):
        for col_idx in range(22, 30):
            ws.cell(row=row_idx, column=col_idx).value = None


@workbook_scope
def _save_company_workbooks(base_wb, base_dir, input_path, payroll_year, payroll_month, company_results, log):
    buffer = io.BytesIO()
    base_wb.save(buffer)
    workbook_bytes = buffer.getvalue()
    output_paths = {}

    for company, bank_results in company_results.items():
        company_output_path = _build_company_output_path(base_dir, company, payroll_year, payroll_month, input_path)
        company_wb = load_workbook(io.BytesIO(workbook_bytes))
        ws = company_wb[company_wb.sheetnames[0]]
        header_row_idx = _find_payroll_header_row(ws)
        data_start_row = header_row_idx + 1

        total_row_idx = _find_total_row(ws)
        if total_row_idx is None:
            raise ValueError('拆分公司文件时未找到“总计”行')

        if ws.max_row > total_row_idx:
            ws.delete_rows(total_row_idx + 1, ws.max_row - total_row_idx)

        for row_idx in range(total_row_idx - 1, data_start_row - 1, -1):
            row_company = _normalize_text(ws.cell(row=row_idx, column=1).value)
            row_values = [ws.cell(row=row_idx, column=col_idx).value for col_idx in range(1, 21)]
            if all(_is_blank(value) for value in row_values):
                ws.delete_rows(row_idx, 1)
                continue
            if row_company != company:
                ws.delete_rows(row_idx, 1)

        new_total_row_idx = _find_total_row(ws)
        if new_total_row_idx is None:
            raise ValueError(f'拆分 {company} 文件时总计行丢失')

        _recalculate_total_row(ws, new_total_row_idx, data_start_row)
        _clear_bank_summary(ws)
        header_ref = ws.cell(row=header_row_idx, column=18 if ws.max_column >= 18 else 17)
        _write_bank_summary(ws, bank_results, header_ref)

        company_wb.save(company_output_path)
        output_paths[company] = company_output_path
        log(f'{company} 文件已生成：{company_output_path}', 'ok')

    return output_paths
