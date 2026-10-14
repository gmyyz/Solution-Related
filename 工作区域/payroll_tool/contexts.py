"""contexts: extracted payroll responsibilities; business rules preserved."""
import re
from decimal import Decimal, ROUND_HALF_UP, ROUND_DOWN
from .workbook_io import load_workbook, workbook_scope
from .rules import (
    BANK_REASON_CODE,
    BANK_VOUCHER_TYPE,
    BONUS_DEPARTMENT_ACCRUAL_OVERRIDES,
    BONUS_TAX_BY_PAYMENT_PERIOD,
    DEPARTMENT_DAILY_ORDER_OVERRIDES,
)
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
from .domain import (
    COMPANY_CODE_TO_NAME,
    CO_DEBIT_GL_BY_SOURCE,
    CO_SPECIAL_COST_ELEMENTS,
    _bonus_company_alias,
    _bonus_months_for_period,
    _build_internal_order_candidates,
    _extract_bonus_department,
    _match_order_from_kok3,
    _normalize_department_for_order,
    _shift_month,
)


@workbook_scope
def _match_timesheet_internal_orders(timesheet_path, save_changes=True):
    wb = load_workbook(timesheet_path)
    summary_ws = wb['工时汇总']
    kok3_ws = wb['KOK3']
    extra_ws = wb['补充逻辑'] if '补充逻辑' in wb.sheetnames else None

    extra_lookup = {}
    if extra_ws is not None:
        for row_idx in range(1, extra_ws.max_row + 1):
            project_name = _normalize_text(extra_ws.cell(row=row_idx, column=1).value)
            search_text = _normalize_text(extra_ws.cell(row=row_idx, column=2).value)
            if project_name and search_text:
                extra_lookup[project_name] = search_text

    company_code_by_name = {'耐数电子': 2050, '耐数信息': 2060}
    kok3_index = {}
    for row_idx in range(2, kok3_ws.max_row + 1):
        order = _format_code(kok3_ws.cell(row=row_idx, column=1).value)
        desc = _normalize_text(kok3_ws.cell(row=row_idx, column=9).value)
        company_code = _format_code(kok3_ws.cell(row=row_idx, column=12).value)
        if not order or not desc or not company_code:
            continue
        try:
            company_key = int(company_code)
        except ValueError:
            continue
        kok3_index.setdefault(company_key, []).append((desc, order))

    matched_rows = 0
    unmatched_rows = []
    used_extra_rows = 0

    for row_idx in range(2, summary_ws.max_row + 1):
        company = _normalize_text(summary_ws.cell(row=row_idx, column=6).value)
        project = _normalize_text(summary_ws.cell(row=row_idx, column=7).value)
        target_cell = summary_ws.cell(row=row_idx, column=13)
        target_cell.value = None

        if not company or not project:
            continue

        company_code = company_code_by_name.get(company)
        if company_code is None:
            unmatched_rows.append({'row_idx': row_idx, 'company': company, 'project': project, 'reason': 'unknown_company'})
            continue

        order, used_key = _match_order_from_kok3(kok3_index, company_code, project, extra_lookup)
        if order:
            target_cell.value = order
            matched_rows += 1
            if used_key and used_key != project:
                used_extra_rows += 1
        else:
            unmatched_rows.append({'row_idx': row_idx, 'company': company, 'project': project, 'reason': 'no_match'})

    if save_changes:
        wb.save(timesheet_path)

    project_index = {}
    allocation_index = {}
    for row_idx in range(2, summary_ws.max_row + 1):
        year_value = summary_ws.cell(row=row_idx, column=2).value
        month_value = summary_ws.cell(row=row_idx, column=3).value
        dept = _normalize_department_for_order(summary_ws.cell(row=row_idx, column=5).value)
        company = _normalize_text(summary_ws.cell(row=row_idx, column=6).value)
        project = _normalize_text(summary_ws.cell(row=row_idx, column=7).value)
        hours = _to_decimal(summary_ws.cell(row=row_idx, column=11).value)
        order = _format_code(summary_ws.cell(row=row_idx, column=13).value)
        item_type = _normalize_text(summary_ws.cell(row=row_idx, column=10).value)
        if not company or not project or not order:
            continue
        key = (company, project)
        if key not in project_index:
            project_index[key] = {'order': order, 'types': set()}
        elif project_index[key]['order'] != order:
            raise ValueError(
                f'工时汇总中「{company}」+「{project}」对应多个不同内部订单：'
                f'{project_index[key]["order"]} 与 {order}，请核对工时数据'
            )
        if item_type:
            project_index[key]['types'].add(item_type)

        month_match = re.search(r'(\d{1,2})', _normalize_text(month_value))
        month_num = int(month_match.group(1)) if month_match else None
        try:
            year_num = int(year_value) if year_value is not None else None
        except (TypeError, ValueError):
            year_num = None
        if year_num and month_num and dept and hours > Decimal('0'):
            alloc_bucket = allocation_index.setdefault((year_num, month_num, company, dept), {})
            alloc_bucket[order] = alloc_bucket.get(order, Decimal('0')) + hours

    return project_index, {
        'matched_rows': matched_rows,
        'used_extra_rows': used_extra_rows,
        'unmatched_rows': unmatched_rows,
    }


@workbook_scope
def _load_timesheet_match_context(timesheet_path):
    wb = load_workbook(timesheet_path, data_only=True)
    summary_ws = wb['工时汇总']
    kok3_ws = wb['KOK3']
    extra_ws = wb['补充逻辑'] if '补充逻辑' in wb.sheetnames else None

    extra_lookup = {}
    if extra_ws is not None:
        for row_idx in range(1, extra_ws.max_row + 1):
            project_name = _normalize_text(extra_ws.cell(row=row_idx, column=1).value)
            search_text = _normalize_text(extra_ws.cell(row=row_idx, column=2).value)
            if project_name and search_text:
                extra_lookup[project_name] = search_text

    kok3_index = {}
    for row_idx in range(2, kok3_ws.max_row + 1):
        order = _format_code(kok3_ws.cell(row=row_idx, column=1).value)
        desc = _normalize_text(kok3_ws.cell(row=row_idx, column=9).value)
        company_code = _format_code(kok3_ws.cell(row=row_idx, column=12).value)
        if not order or not desc or not company_code:
            continue
        try:
            company_key = int(company_code)
        except ValueError:
            continue
        kok3_index.setdefault(company_key, []).append((desc, order))

    company_code_by_name = {'耐数电子': 2050, '耐数信息': 2060}
    project_index = {}
    allocation_index = {}
    for row_idx in range(2, summary_ws.max_row + 1):
        year_value = summary_ws.cell(row=row_idx, column=2).value
        month_value = summary_ws.cell(row=row_idx, column=3).value
        dept = _normalize_department_for_order(summary_ws.cell(row=row_idx, column=5).value)
        company = _normalize_text(summary_ws.cell(row=row_idx, column=6).value)
        project = _normalize_text(summary_ws.cell(row=row_idx, column=7).value)
        hours = _to_decimal(summary_ws.cell(row=row_idx, column=11).value)
        order = ''
        item_type = _normalize_text(summary_ws.cell(row=row_idx, column=10).value)
        company_code = company_code_by_name.get(company)
        if company_code is not None:
            order, _ = _match_order_from_kok3(kok3_index, company_code, project, extra_lookup)
        if not company or not project or not order:
            continue
        key = (company, project)
        if key not in project_index:
            project_index[key] = {'order': order, 'types': set()}
        elif project_index[key]['order'] != order:
            raise ValueError(
                f'工时汇总中「{company}」+「{project}」对应多个不同内部订单：'
                f'{project_index[key]["order"]} 与 {order}，请核对工时数据'
            )
        if item_type:
            project_index[key]['types'].add(item_type)

        month_match = re.search(r'(\d{1,2})', _normalize_text(month_value))
        month_num = int(month_match.group(1)) if month_match else None
        try:
            year_num = int(year_value) if year_value is not None else None
        except (TypeError, ValueError):
            year_num = None
        if year_num and month_num and dept and hours > Decimal('0'):
            alloc_bucket = allocation_index.setdefault((year_num, month_num, company, dept), {})
            alloc_bucket[order] = alloc_bucket.get(order, Decimal('0')) + hours

    return {
        'project_index': project_index,
        'extra_lookup': extra_lookup,
        'kok3_index': kok3_index,
        'allocation_index': allocation_index,
    }


def _lookup_internal_order_from_timesheet(timesheet_context, company, dept, project):
    company_code_by_name = {'耐数电子': 2050, '耐数信息': 2060}
    cc = company_code_by_name.get(company)
    dept_text = _normalize_department_for_order(dept)
    candidates = _build_internal_order_candidates(dept, project)

    for candidate in candidates:
        entry = timesheet_context['project_index'].get((company, candidate))
        if entry and entry.get('order'):
            return entry['order'], candidate
        if cc is not None:
            order, used_key = _match_order_from_kok3(
                timesheet_context['kok3_index'],
                cc,
                candidate,
                timesheet_context['extra_lookup'],
            )
            if order:
                return order, used_key
    daily_project = f'{dept_text}日常工作' if dept_text else ''
    override_order = DEPARTMENT_DAILY_ORDER_OVERRIDES.get((company, dept_text), '')
    if override_order and daily_project in candidates:
        return override_order, daily_project
    return '', (candidates[0] if candidates else '')


def _resolve_a4_internal_order(timesheet_context, company, dept, project, default_order):
    dept_text = _normalize_department_for_order(dept)
    if dept_text != 'OSD运营支持部' or _is_blank_project(project):
        return default_order

    project_entry = timesheet_context['project_index'].get((company, _normalize_text(project)))
    types = project_entry.get('types', set()) if project_entry else set()
    if not project_entry or not any('自研' in t for t in types):
        return default_order

    daily_entry = timesheet_context['project_index'].get((company, 'OSD运营支持部日常工作'))
    if daily_entry and daily_entry.get('order'):
        return daily_entry['order']
    company_code_by_name = {'耐数电子': 2050, '耐数信息': 2060}
    order, _ = _match_order_from_kok3(
        timesheet_context['kok3_index'],
        company_code_by_name.get(company),
        'OSD运营支持部日常工作',
        timesheet_context['extra_lookup'],
    )
    if order:
        return order
    return default_order


@workbook_scope
def _load_bonus_context(bonus_path, cost_center_map):
    # 奖金源文件只用于读取基础数据，必须保持原始格式完全不被脚本触碰。
    wb = load_workbook(bonus_path, data_only=True, read_only=True)
    stats_ws = wb['部门统计']
    ratio_ws = wb['计提比例']

    ratio_map = {}
    for row_idx, ratio_ws_cells in enumerate(ratio_ws.iter_rows(min_row=1, max_col=max(ratio_ws.max_column, 2)), 1):
        month_text = _normalize_text(ratio_ws_cells[(1) - 1].value)
        ratio_value = ratio_ws_cells[(2) - 1].value
        month_match = re.search(r'(\d{1,2})', month_text)
        if not month_match or ratio_value in (None, ''):
            continue
        ratio_map[int(month_match.group(1))] = _to_decimal(ratio_value)

    rows_by_company = {'耐数电子': [], '耐数信息': []}
    for row_idx, stats_ws_cells in enumerate(stats_ws.iter_rows(min_row=2, max_col=max(stats_ws.max_column, 14)), 2):
        company = _bonus_company_alias(stats_ws_cells[(1) - 1].value)
        dept = _extract_bonus_department(stats_ws_cells[(2) - 1].value)
        if not company or not dept:
            continue

        monthly_amounts = {}
        for month in range(1, 13):
            monthly_amounts[month] = _to_money(stats_ws_cells[(2 + month) - 1].value)

        cost_center = cost_center_map.get(company, {}).get(dept, '')
        rows_by_company[company].append(
            {
                'row_idx': row_idx,
                'dept': dept,
                'cost_center': cost_center,
                'monthly_amounts': monthly_amounts,
            }
        )

    return {
        'rows_by_company': rows_by_company,
        'ratio_map': ratio_map,
    }


@workbook_scope
def _load_co_workorder_context(co_path):
    wb = load_workbook(co_path, data_only=True, read_only=True)
    alloc_ws = wb['CO工单分摊'] if 'CO工单分摊' in wb.sheetnames else wb[wb.sheetnames[0]]
    expense_ws = wb['待分摊费用'] if '待分摊费用' in wb.sheetnames else (wb[wb.sheetnames[1]] if len(wb.sheetnames) > 1 else None)
    if expense_ws is None:
        raise ValueError('CO工单分摊文件缺少“待分摊费用”工作表')

    allocation_rows = []
    for row_idx, cells in enumerate(alloc_ws.iter_rows(min_row=2, max_row=alloc_ws.max_row - 1, max_col=7), 2):
        order = _format_code(cells[0].value)
        total_actual = _to_money(cells[6].value)
        if total_actual == Decimal('0.00'):
            continue
        ratio_base = _to_money(cells[4].value)
        if not order:
            raise ValueError(f'CO工单分摊 第 {row_idx} 行 A列内部订单为空，无法分摊')
        if ratio_base <= Decimal('0.00'):
            raise ValueError(f'CO工单分摊 第 {row_idx} 行 E列实际成本借方必须大于 0')
        allocation_rows.append(
            {
                'row_idx': row_idx,
                'order': order,
                'ratio_base': ratio_base,
                'total_actual': total_actual,
            }
        )

    if not allocation_rows:
        raise ValueError('CO工单分摊中未找到 G列不为0 的可分摊明细行')

    total_ratio_base = sum((item['ratio_base'] for item in allocation_rows), Decimal('0.00'))
    if total_ratio_base <= Decimal('0.00'):
        raise ValueError('CO工单分摊可分摊明细行的 E列实际成本借方合计必须大于 0')

    summary_total = _to_money(alloc_ws.cell(row=alloc_ws.max_row, column=7).value)
    selected_total = sum((item['total_actual'] for item in allocation_rows), Decimal('0.00')).quantize(Decimal('0.01'))
    if summary_total != selected_total:
        raise ValueError(
            f'CO工单分摊校验失败：最后一行 G列合计 {summary_total} 与明细 G列合计 {selected_total} 不一致'
        )

    expense_rows = []
    for row_idx, expense_ws_cells in enumerate(expense_ws.iter_rows(min_row=2, max_col=max(expense_ws.max_column, 4)), 2):
        source_account = _format_code(expense_ws_cells[(1) - 1].value)
        amount = _to_money(expense_ws_cells[(4) - 1].value)
        if not source_account or amount == Decimal('0.00'):
            continue
        debit_gl = (
            CO_DEBIT_GL_BY_SOURCE['special']
            if source_account in CO_SPECIAL_COST_ELEMENTS
            else CO_DEBIT_GL_BY_SOURCE['other']
        )
        expense_rows.append(
            {
                'row_idx': row_idx,
                'source_account': source_account,
                'debit_gl': debit_gl,
                'amount': amount,
            }
        )

    if not expense_rows:
        raise ValueError('待分摊费用工作表中未找到 D列非0 的待分摊费用')

    return {
        'allocation_rows': allocation_rows,
        'total_ratio_base': total_ratio_base,
        'summary_total': summary_total,
        'expense_rows': expense_rows,
    }


@workbook_scope
def _load_shared_expense_context(expense_path, cost_center_map):
    wb = load_workbook(expense_path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    reverse_map = _build_cost_center_reverse_map(cost_center_map)
    rows_by_company = {'耐数电子': [], '耐数信息': []}

    for row_idx, ws_cells in enumerate(ws.iter_rows(min_row=2, max_col=max(ws.max_column, 59)), 2):
        company_code = _format_code(ws_cells[(5) - 1].value)
        company = COMPANY_CODE_TO_NAME.get(company_code)
        if not company:
            continue

        gl_account = _format_code(ws_cells[(10) - 1].value)
        amount = _to_money(ws_cells[(12) - 1].value)
        cost_center = _format_code(ws_cells[(19) - 1].value)
        dept_text = _normalize_department_for_order(ws_cells[(20) - 1].value)
        credit_order = _format_code(ws_cells[(21) - 1].value)
        functional_area = _format_code(ws_cells[(59) - 1].value)
        if not gl_account or amount == Decimal('0.00') or not cost_center:
            continue
        if functional_area == '1000':
            continue

        allocation_cost_center = cost_center
        if company == '耐数电子' and cost_center == '20502050':
            allocation_cost_center = '20502060'

        mapped_dept = reverse_map.get(company, {}).get(cost_center, '')
        allocation_dept = reverse_map.get(company, {}).get(allocation_cost_center, '')
        dept = allocation_dept or mapped_dept or dept_text
        dept = _normalize_department_for_order(dept)
        if not dept:
            raise ValueError(f'待分摊费用 第 {row_idx} 行无法根据成本中心 {cost_center} 识别标准部门')

        rows_by_company[company].append(
            {
                'row_idx': row_idx,
                'gl_account': gl_account,
                'amount': amount,
                'cost_center': cost_center,
                'allocation_cost_center': allocation_cost_center,
                'dept': dept,
                'credit_order': credit_order,
            }
        )

    return {
        'rows_by_company': rows_by_company,
    }


def _build_bonus_department_amounts(company, payroll_year, payroll_month, bonus_context):
    posting_year, posting_month = _shift_month(payroll_year, payroll_month, 1)
    ratio = bonus_context['ratio_map'].get(posting_month)
    if ratio is None:
        raise ValueError(f'奖金计提比例表中未找到 {posting_month} 月系数')

    source_months = _bonus_months_for_period(posting_year, posting_month)
    rows = []
    for item in bonus_context['rows_by_company'].get(company, []):
        if not item['cost_center']:
            raise ValueError(f'{company} 奖金数据中部门 {item["dept"]} 未匹配到成本中心')
        base_amount = sum((item['monthly_amounts'].get(month, Decimal('0.00')) for month in source_months), Decimal('0.00'))
        amount = (base_amount * ratio).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        amount = BONUS_DEPARTMENT_ACCRUAL_OVERRIDES.get(
            (company, posting_year, posting_month, item['dept']),
            amount,
        )
        if amount == Decimal('0.00'):
            continue
        rows.append(
            {
                'dept': item['dept'],
                'cost_center': item['cost_center'],
                'base_amount': base_amount.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP),
                'amount': amount,
            }
        )

    return {
        'posting_year': posting_year,
        'posting_month': posting_month,
        'ratio': ratio,
        'source_months': source_months,
        'rows': rows,
    }


@workbook_scope
def _load_mappings(mapping_path):
    """仅从 Mapping 读取成本中心；内部订单由工时数据.xlsx 匹配。"""
    wb = load_workbook(mapping_path, data_only=True)

    cost_center_ws = wb['成本中心']

    cost_center_map = {'耐数电子': {}, '耐数信息': {}}

    for row_idx in range(2, cost_center_ws.max_row + 1):
        dept_2050 = _normalize_text(cost_center_ws.cell(row=row_idx, column=2).value)
        code_2050 = _format_code(cost_center_ws.cell(row=row_idx, column=3).value)
        dept_2060 = _normalize_text(cost_center_ws.cell(row=row_idx, column=7).value)
        code_2060 = _format_code(cost_center_ws.cell(row=row_idx, column=8).value)

        if dept_2050 and code_2050 and dept_2050 not in cost_center_map['耐数电子']:
            cost_center_map['耐数电子'][dept_2050] = code_2050
        if dept_2060 and code_2060 and dept_2060 not in cost_center_map['耐数信息']:
            cost_center_map['耐数信息'][dept_2060] = code_2060

    return cost_center_map


def _build_cost_center_reverse_map(cost_center_map):
    reverse_map = {'耐数电子': {}, '耐数信息': {}}
    for company, dept_map in cost_center_map.items():
        for dept, cost_center in dept_map.items():
            dept_text = _normalize_department_for_order(dept)
            cc = _format_code(cost_center)
            if cc and dept_text and cc not in reverse_map[company]:
                reverse_map[company][cc] = dept_text
    return reverse_map
