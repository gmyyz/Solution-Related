"""Shared payroll parsing and validation for precheck and generation."""
from decimal import Decimal, ROUND_HALF_UP
from .workbook_io import load_workbook, workbook_scope
from .contexts import _lookup_internal_order_from_timesheet
from .rules import EMBEDDED_BONUS_NET_BY_PAYMENT
from .workbook_utils import (
    _find_payroll_header_row, _is_blank, _normalize_text, _is_total_row,
    _copy_cell_style, _to_decimal, _to_money, _mark_row_red,
)


@workbook_scope
def _prepare_payroll_preview(
    input_path,
    cost_center_map,
    timesheet_context,
    selected_companies,
    payment_period=None,
):
    wb = load_workbook(input_path)
    ws = wb[wb.sheetnames[0]]
    header_row_idx = _find_payroll_header_row(ws)
    data_start_row = header_row_idx + 1

    fill_a = fill_b = 0
    last_a = None
    last_b = None
    last_a_style_cell = None
    last_b_style_cell = None
    total_row_idx = ws.max_row + 1
    issue_rows = {}
    company_issue_counts = {}
    payroll_summary = {}
    rows_by_company = {}

    if _is_blank(ws.cell(row=data_start_row, column=1).value):
        for probe_row_idx in range(data_start_row, ws.max_row + 1):
            company_value = _normalize_text(ws.cell(row=probe_row_idx, column=1).value)
            if _is_total_row(company_value):
                break
            if company_value:
                if company_value == '耐数信息':
                    last_a = '耐数电子'
                break

    for row_idx in range(data_start_row, ws.max_row + 1):
        cell_a = ws.cell(row=row_idx, column=1)
        cell_b = ws.cell(row=row_idx, column=2)

        if _is_total_row(cell_a.value):
            total_row_idx = row_idx
            break

        if _is_blank(cell_a.value):
            if last_a is not None:
                cell_a.value = last_a
                fill_a += 1
                if last_a_style_cell is not None:
                    _copy_cell_style(last_a_style_cell, cell_a)
        else:
            last_a = cell_a.value
            last_a_style_cell = cell_a

        if _is_blank(cell_b.value):
            if last_b is not None:
                cell_b.value = last_b
                fill_b += 1
                if last_b_style_cell is not None:
                    _copy_cell_style(last_b_style_cell, cell_b)
        else:
            last_b = cell_b.value
            last_b_style_cell = cell_b

    embedded_bonus_adjustments = {}
    if payment_period:
        company_codes = {'耐数电子': '2050', '耐数信息': '2060'}
        for company in selected_companies:
            company_code = company_codes.get(company)
            bonus = EMBEDDED_BONUS_NET_BY_PAYMENT.get((company_code, *payment_period))
            if bonus is None:
                continue

            raw_total = Decimal('0')
            recalculated_total = Decimal('0')
            for row_idx in range(data_start_row, total_row_idx):
                if _normalize_text(ws.cell(row=row_idx, column=1).value) != company:
                    continue
                expected_raw = _to_decimal(ws.cell(row=row_idx, column=5).value)
                for col_idx in range(11, 17):
                    expected_raw -= _to_decimal(ws.cell(row=row_idx, column=col_idx).value)
                raw_total += _to_decimal(ws.cell(row=row_idx, column=17).value)
                recalculated_total += expected_raw

            detected_bonus = (raw_total - recalculated_total).quantize(
                Decimal('0.01'),
                rounding=ROUND_HALF_UP,
            )
            if detected_bonus != bonus:
                raise ValueError(
                    f'{company} {payment_period[0]}-{payment_period[1]:02d} 工资单内含年终奖 '
                    f'应为 {bonus}，按实发逻辑检测为 {detected_bonus}，已停止处理'
                )
            embedded_bonus_adjustments[company] = bonus

    header_ref = ws.cell(row=header_row_idx, column=18 if ws.max_column >= 18 else 17)
    data_style_col = 18 if ws.max_column >= 18 else 17
    s_header = ws.cell(row=header_row_idx, column=19)
    t_header = ws.cell(row=header_row_idx, column=20)
    s_header.value = '成本中心'
    _copy_cell_style(header_ref, s_header)
    t_header.value = '内部订单'
    _copy_cell_style(header_ref, t_header)

    if ws.column_dimensions['R'].width:
        ws.column_dimensions['S'].width = ws.column_dimensions['R'].width
    if ws.column_dimensions['Q'].width:
        ws.column_dimensions['T'].width = ws.column_dimensions['Q'].width

    for row_idx in range(data_start_row, total_row_idx):
        row_values = [ws.cell(row=row_idx, column=col).value for col in range(1, 19)]
        if all(_is_blank(value) for value in row_values):
            continue

        row_issues = []
        company = _normalize_text(ws.cell(row=row_idx, column=1).value)
        dept = _normalize_text(ws.cell(row=row_idx, column=2).value)
        project = _normalize_text(ws.cell(row=row_idx, column=3).value)
        rows_by_company.setdefault(company, []).append(row_idx)
        if company not in selected_companies:
            continue

        expected_raw = _to_decimal(ws.cell(row=row_idx, column=5).value)
        for col_idx in range(11, 17):
            expected_raw -= _to_decimal(ws.cell(row=row_idx, column=col_idx).value)
        expected_actual = expected_raw.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        if company in embedded_bonus_adjustments:
            ws.cell(row=row_idx, column=17).value = float(expected_raw)
        actual_value = _to_money(ws.cell(row=row_idx, column=17).value)

        company_summary = payroll_summary.setdefault(
            company,
            {'salary': Decimal('0'), 'fund': Decimal('0'), 'tax': Decimal('0'), 'social': Decimal('0')},
        )
        company_summary['salary'] += _to_decimal(ws.cell(row=row_idx, column=17).value)
        company_summary['fund'] += _to_decimal(ws.cell(row=row_idx, column=10).value) + _to_decimal(
            ws.cell(row=row_idx, column=14).value
        )
        company_summary['tax'] += _to_decimal(ws.cell(row=row_idx, column=15).value)
        company_summary['social'] += sum(
            (_to_decimal(ws.cell(row=row_idx, column=col_idx).value) for col_idx in (6, 7, 8, 9, 11, 12, 13)),
            Decimal('0'),
        )

        if expected_actual != actual_value:
            row_issues.append(f'实发校验失败(Q={actual_value}, 计算值={expected_actual})')

        company_cost_center_map = cost_center_map.get(company)
        if company_cost_center_map is None:
            row_issues.append(f'无法识别公司：{company or "空"}')
            cost_center_value = ''
            internal_order_value = ''
        else:
            cost_center_value = company_cost_center_map.get(dept, '')
            if not cost_center_value:
                row_issues.append(f'成本中心未匹配：{dept or "空"}')
            internal_order_value, project_for_match = _lookup_internal_order_from_timesheet(
                timesheet_context,
                company,
                dept,
                project,
            )
            if not internal_order_value:
                row_issues.append(f'内部订单未匹配：{project_for_match or "空"}')

        s_cell = ws.cell(row=row_idx, column=19)
        t_cell = ws.cell(row=row_idx, column=20)
        ref_cell = ws.cell(row=row_idx, column=data_style_col)
        _copy_cell_style(ref_cell, s_cell)
        _copy_cell_style(ref_cell, t_cell)
        s_cell.value = cost_center_value
        t_cell.value = internal_order_value

        if row_issues:
            issue_rows[row_idx] = {'company': company, 'issues': row_issues}
            company_issue_counts.setdefault(company, {'formula': 0, 'cost_center': 0, 'order': 0})
            for issue in row_issues:
                if issue.startswith('实发校验失败'):
                    company_issue_counts[company]['formula'] += 1
                elif issue.startswith('成本中心未匹配') or issue.startswith('无法识别公司'):
                    company_issue_counts[company]['cost_center'] += 1
                elif issue.startswith('内部订单未匹配'):
                    company_issue_counts[company]['order'] += 1

    for row_idx in issue_rows:
        _mark_row_red(ws, row_idx, 20)

    return {
        'fill_a': fill_a, 'fill_b': fill_b,
        'header_ref': header_ref,
        'workbook': wb,
        'worksheet': ws,
        'total_row_idx': total_row_idx,
        'issue_rows': issue_rows,
        'company_issue_counts': company_issue_counts,
        'payroll_summary': payroll_summary,
        'rows_by_company': rows_by_company,
        'embedded_bonus_adjustments': embedded_bonus_adjustments,
    }
