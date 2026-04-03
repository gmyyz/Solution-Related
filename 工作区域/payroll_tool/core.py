import glob
import io
import os
import queue
import re
import threading
import tkinter as tk
from copy import copy
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP, ROUND_DOWN
from itertools import combinations
from tkinter import ttk, messagebox

from openpyxl import load_workbook

from .options import (
    ACTUAL_VOUCHERS,
    ACCRUAL_VOUCHERS,
    BONUS_VOUCHERS,
    CO_VOUCHERS,
    RD_VOUCHERS,
    VOUCHER_DISPLAY_ORDER,
    normalize_run_options,
    requires_bank_data,
    requires_bonus_data,
    requires_co_data,
    requires_shared_expense_data,
)


# ============================================================
# 配色常量（RIGOL 品牌风格：黄黑白）
# ============================================================

COLOR_PRIMARY = '#FFD700'
COLOR_SUCCESS = '#FFD700'
COLOR_DANGER = '#FF4444'
COLOR_WARNING = '#FFA500'
COLOR_BG = '#1A1A1A'
COLOR_CARD = '#2A2A2A'
COLOR_BORDER = '#3D3D3D'
COLOR_TEXT_MAIN = '#FFFFFF'
COLOR_TEXT_SUB = '#AAAAAA'
COLOR_HEADER_BG = '#111111'

RAW_FILE_PATTERN = '人力成本研发项目分摊* - to财务-原始.xlsx'
BANK_FILE_PATTERN = '银行流水*.xlsx'
ERROR_FONT_COLOR = 'FFFF0000'
COMPANY_NAME_TO_CODE = {'耐数电子': '2050', '耐数信息': '2060'}
COMPANY_CODE_TO_NAME = {value: key for key, value in COMPANY_NAME_TO_CODE.items()}
SOCIAL_DEBIT_ACCOUNT_COLUMNS = [
    (6, '2211020002'),
    (7, '2211030002'),
    (8, '2211040001'),
    (9, '2211060002'),
    (11, '2211020001'),
    (12, '2211030001'),
    (13, '2211060001'),
]
SOCIAL_TEXT_BY_ACCOUNT = {
    '2211020002': '支付{period}工作期间公司承担养老保险',
    '2211030002': '支付{period}工作期间公司承担失业保险',
    '2211040001': '支付{period}工作期间公司承担工伤保险',
    '2211060002': '支付{period}工作期间公司承担医疗保险',
    '2211020001': '支付{period}工作期间个人承担养老保险',
    '2211030001': '支付{period}工作期间个人承担失业保险',
    '2211060001': '支付{period}工作期间个人承担医疗保险',
}
ACCRUAL_DEBIT_SPECS = [
    (5, '6601010001', '薪酬'),
    (6, '6601030008', '养老保险'),
    (8, '6601030003', '工伤保险'),
    (7, '6601030002', '失业保险'),
    (9, '6601030005', '医疗保险'),
    (10, '6601030006', '住房公积金'),
]
ACCRUAL_CREDIT_SPECS = [
    (10, '2211070002', '公司承担住房公积金'),
    (9, '2211060002', '公司承担医疗保险'),
    (7, '2211030002', '公司承担失业保险'),
    (8, '2211040001', '公司承担工伤保险'),
    (6, '2211020002', '公司承担养老保险'),
    (14, '2211070001', '个人承担住房公积金'),
    (13, '2211060001', '个人承担医疗保险'),
    (12, '2211030001', '个人承担失业保险'),
    (11, '2211020001', '个人承担养老保险'),
]
ACCOUNTS_TO_CROSS_CHECK = [
    '2211060002',
    '2211030002',
    '2211040001',
    '2211020002',
    '2221070000',
    '2211010001',
    '2211060001',
    '2211030001',
    '2211020001',
]
CO_SPECIAL_COST_ELEMENTS = {
    '6601010001',
    '6601030002',
    '6601030003',
    '6601030005',
    '6601030006',
    '6601030008',
}
CO_DEBIT_GL_BY_SOURCE = {
    'special': '5001010011',
    'other': '5001010012',
}
CO_CREDIT_COST_CENTER = '20502020'
CO_CREDIT_ORDER = '9201856'


def _apply_style(root):
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('TFrame', background=COLOR_BG)
    style.configure(
        'Horizontal.TProgressbar',
        background=COLOR_PRIMARY,
        troughcolor='#333333',
        borderwidth=0,
        thickness=8,
    )


def _btn(parent, text, bg, command, padx=20, pady=8):
    if bg in (COLOR_PRIMARY, COLOR_SUCCESS):
        fg_color = '#000000'
        active_bg = '#E6C200'
    elif bg == COLOR_DANGER:
        fg_color = '#FFFFFF'
        active_bg = '#CC2222'
    else:
        fg_color = '#FFFFFF'
        active_bg = '#555555'
    return tk.Button(
        parent,
        text=text,
        bg=bg,
        fg=fg_color,
        font=('微软雅黑', 11, 'bold'),
        relief='flat',
        padx=padx,
        pady=pady,
        cursor='hand2',
        activebackground=active_bg,
        activeforeground=fg_color,
        command=command,
    )


def _circle_label(parent, text, size=26, bg_color=None, fg_color='#000000', parent_bg=None):
    bg_color = bg_color or COLOR_PRIMARY
    parent_bg = parent_bg or COLOR_BG
    canvas = tk.Canvas(parent, width=size, height=size, bg=parent_bg, highlightthickness=0)
    canvas.create_oval(1, 1, size - 1, size - 1, fill=bg_color, outline='')
    canvas.create_text(
        size // 2,
        size // 2,
        text=text,
        font=('微软雅黑', max(size // 3, 8), 'bold'),
        fill=fg_color,
        anchor='center',
    )
    return canvas


def _get_raw_dir(base_dir):
    return os.path.join(base_dir, '原始数据', '工资单')


def _get_mapping_path(base_dir):
    return os.path.join(base_dir, 'Mapping表.xlsx')


def _get_bank_dir(base_dir):
    return os.path.join(base_dir, '原始数据', '银行流水')


def _get_timesheet_path(base_dir):
    return os.path.join(base_dir, '原始数据', '工时数据', '工时数据.xlsx')


def _get_bonus_path(base_dir):
    return os.path.join(base_dir, '原始数据', '奖金数据', '年终奖计提2026_ - to财务.xlsx')


def _get_shared_expense_path(base_dir, payroll_year, payroll_month):
    post_year, post_month = _shift_month(payroll_year, payroll_month, 1)
    period_text = f'{post_year % 100:02d}{post_month:02d}'
    return os.path.join(base_dir, '原始数据', '待分摊费用', f'待分摊费用{period_text}.xlsx')


def _get_co_workorder_path(base_dir, payroll_year, payroll_month):
    post_year, post_month = _shift_month(payroll_year, payroll_month, 1)
    folder_name = f'{post_year % 100:02d}{post_month:02d}'
    return os.path.join(base_dir, '耐数电子', folder_name, 'CO工单分摊', 'CO工单分摊.xlsx')


def _get_voucher_template_path(base_dir):
    return os.path.join(os.path.dirname(base_dir), '模版性文件', '2060总账凭证导入-实际202602薪酬.XLS')


def _get_logo_path(base_dir):
    candidates = [
        os.path.join(base_dir, 'rigol_logo.png'),
        os.path.join(os.path.dirname(base_dir), '模版性文件', 'rigol_logo.png'),
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def _find_raw_files(raw_dir):
    return sorted(glob.glob(os.path.join(raw_dir, RAW_FILE_PATTERN)))


def _find_bank_files(bank_dir):
    return sorted(glob.glob(os.path.join(bank_dir, BANK_FILE_PATTERN)))


def _build_company_output_path(base_dir, company, payroll_year, payroll_month, input_path):
    post_year, post_month = _shift_month(payroll_year, payroll_month, 1)
    folder_name = f'{post_year % 100:02d}{post_month:02d}'
    company_dir = os.path.join(base_dir, company, folder_name)
    os.makedirs(company_dir, exist_ok=True)

    output_name = f'{payroll_month}月工资单-整理后.xlsx'
    return os.path.join(company_dir, output_name)


def _build_voucher_output_path(base_dir, company, payroll_year, payroll_month):
    company_code = COMPANY_NAME_TO_CODE[company]
    post_year, post_month = _shift_month(payroll_year, payroll_month, 1)
    folder_name = f'{post_year % 100:02d}{post_month:02d}'
    company_dir = os.path.join(base_dir, company, folder_name)
    os.makedirs(company_dir, exist_ok=True)
    return os.path.join(company_dir, f'{company_code}总账凭证导入-实际{post_year}{post_month:02d}薪酬.XLS')


def _build_accrual_voucher_output_path(base_dir, company, payroll_year, payroll_month):
    company_code = COMPANY_NAME_TO_CODE[company]
    post_year, post_month = _shift_month(payroll_year, payroll_month, 1)
    folder_name = f'{post_year % 100:02d}{post_month:02d}'
    company_dir = os.path.join(base_dir, company, folder_name)
    os.makedirs(company_dir, exist_ok=True)
    return os.path.join(company_dir, f'{company_code}总账凭证导入-计提{post_year}{post_month:02d}薪酬.XLS')


def _build_bonus_voucher_output_path(base_dir, company, payroll_year, payroll_month):
    company_code = COMPANY_NAME_TO_CODE[company]
    post_year, post_month = _shift_month(payroll_year, payroll_month, 1)
    folder_name = f'{post_year % 100:02d}{post_month:02d}'
    company_dir = os.path.join(base_dir, company, folder_name)
    os.makedirs(company_dir, exist_ok=True)
    return os.path.join(company_dir, f'{company_code}总账凭证导入-计提{_bonus_label_for_filename(post_year, post_month)}年终奖.XLS')


def _build_co_voucher_output_path(base_dir, payroll_year, payroll_month):
    post_year, post_month = _shift_month(payroll_year, payroll_month, 1)
    folder_name = f'{post_year % 100:02d}{post_month:02d}'
    co_dir = os.path.join(base_dir, '耐数电子', folder_name, 'CO工单分摊')
    os.makedirs(co_dir, exist_ok=True)
    return os.path.join(co_dir, f'2050总账凭证导入-CO工单分摊{post_year}{post_month:02d}.XLS')


def _build_rd_allocation_voucher_output_path(base_dir, company, payroll_year, payroll_month):
    company_code = COMPANY_NAME_TO_CODE[company]
    post_year, post_month = _shift_month(payroll_year, payroll_month, 1)
    folder_name = f'{post_year % 100:02d}{post_month:02d}'
    company_dir = os.path.join(base_dir, company, folder_name)
    os.makedirs(company_dir, exist_ok=True)
    return os.path.join(company_dir, f'{company_code}总账-研发费用分摊{post_year % 100:02d}{post_month:02d}.XLS')


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
    return Decimal(str(value))


def _to_money(value):
    return _to_decimal(value).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)


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


def _allocate_amount_by_weights(total_amount, weighted_items):
    total_amount = _to_money(total_amount)
    positive_items = [(key, _to_decimal(weight)) for key, weight in weighted_items if _to_decimal(weight) > Decimal('0')]
    if total_amount == Decimal('0.00') or not positive_items:
        return {}

    total_weight = sum((weight for _, weight in positive_items), Decimal('0'))
    if total_weight <= Decimal('0'):
        return {}

    allocations = {}
    remainders = []
    allocated = Decimal('0.00')
    cent = Decimal('0.01')

    for key, weight in positive_items:
        raw = total_amount * weight / total_weight
        floored = raw.quantize(cent, rounding=ROUND_DOWN)
        allocations[key] = floored
        allocated += floored
        remainders.append((raw - floored, key))

    remaining_cents = int(((total_amount - allocated) / cent).quantize(Decimal('1'), rounding=ROUND_HALF_UP))
    remainders.sort(key=lambda item: (item[0], item[1]), reverse=True)

    for idx in range(remaining_cents):
        _, key = remainders[idx % len(remainders)]
        allocations[key] += cent

    return {key: amount.quantize(cent) for key, amount in allocations.items() if amount != Decimal('0.00')}


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


def _extract_payroll_period(input_path):
    name = os.path.basename(input_path)
    match = re.search(r'(\d{6})', name)
    if not match:
        raise ValueError(f'无法从工资表文件名识别账期：{name}')
    yyyymm = match.group(1)
    return int(yyyymm[:4]), int(yyyymm[4:])


def _shift_month(year, month, offset):
    total = year * 12 + (month - 1) + offset
    return total // 12, total % 12 + 1


def _normalize_department_for_order(dept):
    text = _normalize_text(dept)
    aliases = {
        'NAISHU SOLUTION': 'OSD运营支持部',
        'LHD硬件逻辑部': 'HLD硬件逻辑部',
        '逻辑部': 'HLD硬件逻辑部',
        '运营支持部OSD': 'OSD运营支持部',
        'OSD运营支持部': 'OSD运营支持部',
    }
    return aliases.get(text, text)


def _build_internal_order_candidates(dept, project):
    if not _is_blank_project(project):
        return [_normalize_text(project)]

    dept_text = _normalize_department_for_order(dept)
    if not dept_text:
        return []

    return [f'{dept_text}日常工作']


def _lookup_internal_order(order_map, dept, project):
    for candidate in _build_internal_order_candidates(dept, project):
        value = order_map.get(candidate, '')
        if value:
            return value, candidate
    return '', (_build_internal_order_candidates(dept, project)[0] if _build_internal_order_candidates(dept, project) else '')


def _match_order_from_kok3(kok3_index, company_code, project_name, extra_lookup):
    """先对描述做精确匹配，再允许「搜索词包含于 KOK3 描述(I列)」的包含匹配。"""
    if not project_name:
        return '', project_name

    search_terms = [project_name]
    extra_term = extra_lookup.get(project_name, '')
    if extra_term and extra_term not in search_terms:
        search_terms.append(extra_term)

    rows = kok3_index.get(company_code, [])

    for search_text in search_terms:
        exact = [(desc, order) for desc, order in rows if search_text == desc]
        if exact:
            orders = {order for _, order in exact}
            if len(orders) == 1:
                return exact[0][1], search_text
            return '', search_text

    for search_text in search_terms:
        contain = [(desc, order) for desc, order in rows if search_text != desc and search_text in desc]
        if contain:
            orders = {order for _, order in contain}
            if len(orders) == 1:
                return contain[0][1], search_text
            return '', search_text

    return '', (search_terms[-1] if search_terms else project_name)


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

    return {
        'project_index': project_index,
        'extra_lookup': extra_lookup,
        'kok3_index': kok3_index,
        'allocation_index': allocation_index,
    }


def _lookup_internal_order_from_timesheet(timesheet_context, company, dept, project):
    company_code_by_name = {'耐数电子': 2050, '耐数信息': 2060}
    cc = company_code_by_name.get(company)

    for candidate in _build_internal_order_candidates(dept, project):
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
    return '', (_build_internal_order_candidates(dept, project)[0] if _build_internal_order_candidates(dept, project) else '')


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


def _load_bonus_context(bonus_path, cost_center_map):
    # 奖金源文件只用于读取基础数据，必须保持原始格式完全不被脚本触碰。
    wb = load_workbook(bonus_path, data_only=True, read_only=True)
    stats_ws = wb['部门统计']
    ratio_ws = wb['计提比例']

    ratio_map = {}
    for row_idx in range(1, ratio_ws.max_row + 1):
        month_text = _normalize_text(ratio_ws.cell(row=row_idx, column=1).value)
        ratio_value = ratio_ws.cell(row=row_idx, column=2).value
        month_match = re.search(r'(\d{1,2})', month_text)
        if not month_match or ratio_value in (None, ''):
            continue
        ratio_map[int(month_match.group(1))] = _to_decimal(ratio_value)

    rows_by_company = {'耐数电子': [], '耐数信息': []}
    for row_idx in range(2, stats_ws.max_row + 1):
        company = _bonus_company_alias(stats_ws.cell(row=row_idx, column=1).value)
        dept = _extract_bonus_department(stats_ws.cell(row=row_idx, column=2).value)
        if not company or not dept:
            continue

        monthly_amounts = {}
        for month in range(1, 13):
            monthly_amounts[month] = _to_money(stats_ws.cell(row=row_idx, column=2 + month).value)

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


def _load_co_workorder_context(co_path):
    wb = load_workbook(co_path, data_only=True, read_only=True)
    alloc_ws = wb['CO工单分摊'] if 'CO工单分摊' in wb.sheetnames else wb[wb.sheetnames[0]]
    expense_ws = wb['待分摊费用'] if '待分摊费用' in wb.sheetnames else (wb[wb.sheetnames[1]] if len(wb.sheetnames) > 1 else None)
    if expense_ws is None:
        raise ValueError('CO工单分摊文件缺少“待分摊费用”工作表')

    allocation_rows = []
    for row_idx in range(2, alloc_ws.max_row):
        order = _format_code(alloc_ws.cell(row=row_idx, column=1).value)
        total_actual = _to_money(alloc_ws.cell(row=row_idx, column=7).value)
        if total_actual == Decimal('0.00'):
            continue
        ratio_base = _to_money(alloc_ws.cell(row=row_idx, column=5).value)
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
    for row_idx in range(2, expense_ws.max_row + 1):
        source_account = _format_code(expense_ws.cell(row=row_idx, column=1).value)
        amount = _to_money(expense_ws.cell(row=row_idx, column=4).value)
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


def _load_shared_expense_context(expense_path, cost_center_map):
    wb = load_workbook(expense_path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    reverse_map = _build_cost_center_reverse_map(cost_center_map)
    rows_by_company = {'耐数电子': [], '耐数信息': []}

    for row_idx in range(2, ws.max_row + 1):
        company_code = _format_code(ws.cell(row=row_idx, column=5).value)
        company = COMPANY_CODE_TO_NAME.get(company_code)
        if not company:
            continue

        gl_account = _format_code(ws.cell(row=row_idx, column=10).value)
        amount = _to_money(ws.cell(row=row_idx, column=12).value)
        cost_center = _format_code(ws.cell(row=row_idx, column=19).value)
        dept_text = _normalize_department_for_order(ws.cell(row=row_idx, column=20).value)
        credit_order = _format_code(ws.cell(row=row_idx, column=21).value)
        if not gl_account or amount == Decimal('0.00') or not cost_center:
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


def _load_bank_records(bank_dir):
    bank_records = {}

    for path in _find_bank_files(bank_dir):
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]

        for row_idx in range(2, ws.max_row + 1):
            company_code = _format_code(ws.cell(row=row_idx, column=2).value)
            bank_account = _format_code(ws.cell(row=row_idx, column=3).value)
            trans_date = ws.cell(row=row_idx, column=4).value
            trans_period = _extract_year_month(trans_date)
            payment_name = _normalize_text(ws.cell(row=row_idx, column=13).value)
            outgoing_amt = _to_money(ws.cell(row=row_idx, column=15).value)
            usage = _normalize_text(ws.cell(row=row_idx, column=16).value)
            bank_subject = _format_code(ws.cell(row=row_idx, column=34).value)

            if not company_code or trans_period is None:
                continue

            key = (company_code, trans_period[0], trans_period[1])
            bank_records.setdefault(key, []).append(
                {
                    'file': os.path.basename(path),
                    'row_idx': row_idx,
                    'company_code': company_code,
                    'bank_account': bank_account,
                    'trans_date': trans_date,
                    'payment_name': payment_name,
                    'outgoing_amt': outgoing_amt,
                    'usage': usage,
                    'bank_subject': bank_subject,
                }
            )

    return bank_records


def _summarize_bank_records(bank_records, company_code, year, month):
    records = bank_records.get((company_code, year, month), [])
    salary_records = [record for record in records if '代发工资' in record['usage']]
    salary_amount = sum((record['outgoing_amt'] for record in salary_records), Decimal('0.00'))
    fund_amount = sum(
        (
            record['outgoing_amt']
            for record in records
            if '北京住房公积金管理中心' in record['payment_name']
        ),
        Decimal('0.00'),
    )
    treasury_amounts = sorted(
        [
            record['outgoing_amt']
            for record in records
            if '国家金库北京市分库' in record['payment_name'] and record['outgoing_amt'] > 0
        ]
    )
    tax_amount = treasury_amounts[0] if treasury_amounts else Decimal('0.00')
    social_amount = sum(treasury_amounts[1:], Decimal('0.00'))

    return {
        'salary': salary_amount,
        'salary_records': salary_records,
        'fund': fund_amount,
        'tax': tax_amount,
        'social': social_amount,
        'files': sorted({record['file'] for record in records}),
        'accounts': sorted({record['bank_account'] for record in records if record['bank_account']}),
        'record_count': len(records),
    }


def _match_salary_combination(salary_records, target_amount):
    target = _to_money(target_amount)
    if not salary_records:
        return {
            'matched': False,
            'matched_amount': Decimal('0.00'),
            'all_amount': Decimal('0.00'),
            'matched_records': [],
            'note': '未找到代发工资流水',
        }

    all_amount = sum((record['outgoing_amt'] for record in salary_records), Decimal('0.00'))
    max_size = min(5, len(salary_records))
    for size in range(1, max_size + 1):
        for combo in combinations(salary_records, size):
            combo_amount = sum((record['outgoing_amt'] for record in combo), Decimal('0.00'))
            if combo_amount == target:
                combo_desc = ' + '.join(str(record['outgoing_amt']) for record in combo)
                return {
                    'matched': True,
                    'matched_amount': combo_amount,
                    'all_amount': all_amount,
                    'matched_records': list(combo),
                    'note': f'匹配到 {size} 条代发工资组合：{combo_desc}',
                }

    return {
        'matched': False,
        'matched_amount': all_amount,
        'all_amount': all_amount,
        'matched_records': [],
        'note': f'未找到匹配组合；全部代发工资合计={all_amount}',
    }


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


def _select_treasury_records(records):
    treasury_records = sorted(
        [
            record
            for record in records
            if '国家金库北京市分库' in record['payment_name'] and record['outgoing_amt'] > 0
        ],
        key=lambda item: item['outgoing_amt'],
    )
    if not treasury_records:
        return None, []
    return treasury_records[0], treasury_records[1:]


def _get_last_bank_date(records):
    dates = [record['trans_date'] for record in records if isinstance(record['trans_date'], (datetime, date))]
    if not dates:
        raise ValueError('未找到可用的银行交易日期')
    return max(dates)


def _sap_date_value(value):
    if not isinstance(value, (datetime, date)):
        raise ValueError(f'无法转换为凭证日期：{value}')
    return int(value.strftime('%Y%m%d'))


def _period_text(year, month):
    return f'{year}年{month}月'


def _bonus_months_for_period(year, month):
    if month in (3, 6, 9, 12):
        return list(range(1, month + 1))
    return [month]


def _bonus_label_for_filename(year, month):
    if month in (3, 6, 9, 12):
        if month == 3:
            return f'{year}Q1'
        return f'Q1-Q{month // 3}'
    return f'{year}{month:02d}'


def _bonus_label_for_text(year, month):
    return _bonus_label_for_filename(year, month)


def _extract_bonus_department(value):
    text = _normalize_text(value)
    if not text:
        return ''
    return _normalize_department_for_order(text.split('/')[-1])


def _bonus_company_alias(value):
    text = _normalize_text(value)
    mapping = {
        '北京普源耐数电子有限公司': '耐数电子',
        '北京耐数信息有限公司': '耐数信息',
    }
    return mapping.get(text, '')


def _month_end_date(year, month):
    next_year, next_month = _shift_month(year, month, 1)
    first_of_next = date(next_year, next_month, 1)
    return first_of_next.fromordinal(first_of_next.toordinal() - 1)


def _make_voucher_row(
    group,
    company_code,
    posting_date,
    gl_account,
    amount,
    text,
    posting_key,
    cost_center='',
    order='',
    voucher_type='ZZ',
):
    return {
        'A': group,
        'B': company_code,
        'C': voucher_type,
        'D': _sap_date_value(posting_date),
        'E': _sap_date_value(posting_date),
        'F': 'CNY',
        'G': posting_key,
        'H': '',
        'I': gl_account,
        'J': cost_center,
        'K': order,
        'L': '',
        'M': _to_text_money(amount),
        'N': text,
        'O': '',
        'P': '',
        'Q': '',
    }


def _build_a1_rows(company, company_code, ws, total_row_idx, bank_records, payroll_year, payroll_month):
    next_year, next_month = _shift_month(payroll_year, payroll_month, 1)
    month_records = bank_records.get((company_code, next_year, next_month), [])
    _, social_records = _select_treasury_records(month_records)
    if not social_records:
        raise ValueError(f'{company} 未找到社保对应的银行流水')

    posting_date = _get_last_bank_date(social_records)
    text = f'支付{_period_text(payroll_year, payroll_month)}工作期间社保'
    rows = []
    period_text = _period_text(payroll_year, payroll_month)

    for col_idx, gl_account in SOCIAL_DEBIT_ACCOUNT_COLUMNS:
        amount = _to_money(ws.cell(total_row_idx, col_idx).value)
        if amount == Decimal('0.00'):
            continue
        debit_text = SOCIAL_TEXT_BY_ACCOUNT[gl_account].format(period=period_text)
        rows.append(_make_voucher_row('A1', company_code, posting_date, gl_account, amount, debit_text, 40))

    for record in social_records:
        credit_row = _make_voucher_row(
            'A1',
            company_code,
            posting_date,
            record['bank_subject'],
            record['outgoing_amt'],
            text,
            50,
        )
        credit_row['P'] = '204'
        rows.append(
            credit_row
        )

    return rows


def _build_a2_rows(company, company_code, bank_records, payroll_year, payroll_month):
    next_year, next_month = _shift_month(payroll_year, payroll_month, 1)
    month_records = bank_records.get((company_code, next_year, next_month), [])
    fund_records = [record for record in month_records if '北京住房公积金管理中心' in record['payment_name']]
    if not fund_records:
        raise ValueError(f'{company} 未找到公积金对应的银行流水')

    posting_date = _get_last_bank_date(fund_records)
    text = f'支付{_period_text(next_year, next_month)}工作期间公积金'
    total_amount = sum((record['outgoing_amt'] for record in fund_records), Decimal('0.00'))
    personal_amount = (total_amount / Decimal('2')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    company_amount = (total_amount - personal_amount).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    period_text = _period_text(next_year, next_month)

    rows = [
        _make_voucher_row(
            'A2',
            company_code,
            posting_date,
            '2211070001',
            personal_amount,
            f'支付{period_text}工作期间个人承担公积金',
            40,
        ),
        _make_voucher_row(
            'A2',
            company_code,
            posting_date,
            '2211070002',
            company_amount,
            f'支付{period_text}工作期间公司承担公积金',
            40,
        ),
    ]

    for record in fund_records:
        credit_row = _make_voucher_row(
            'A2',
            company_code,
            posting_date,
            record['bank_subject'],
            record['outgoing_amt'],
            text,
            50,
        )
        credit_row['P'] = '204'
        rows.append(
            credit_row
        )

    return rows


def _build_a3_rows(company, company_code, ws, bank_records, payroll_year, payroll_month):
    next_year, next_month = _shift_month(payroll_year, payroll_month, 1)
    month_records = bank_records.get((company_code, next_year, next_month), [])
    salary_records = [record for record in month_records if '代发工资' in record['usage']]
    tax_record, _ = _select_treasury_records(month_records)
    if tax_record is None:
        raise ValueError(f'{company} 未找到个税对应的银行流水')

    labor_rows = []
    salary_total = Decimal('0')
    tax_total = Decimal('0')
    total_row_idx = _find_total_row(ws)
    if total_row_idx is None:
        raise ValueError(f'{company} 工资单中未找到总计行')
    salary_match_target = _to_money(ws.cell(total_row_idx, 17).value)
    for row_idx in range(3, ws.max_row + 1):
        if _is_total_row(ws.cell(row_idx, 1).value):
            break
        item_type = _normalize_text(ws.cell(row_idx, 4).value)
        if item_type == '劳务费':
            labor_rows.append(
                {
                    'amount': _to_money(ws.cell(row_idx, 5).value),
                    'cost_center': _format_code(ws.cell(row_idx, 19).value),
                    'order': _format_code(ws.cell(row_idx, 20).value),
                }
            )
        elif item_type == '薪资':
            salary_total += _to_decimal(ws.cell(row_idx, 17).value)
            tax_total += _to_decimal(ws.cell(row_idx, 15).value)

    salary_total = salary_total.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    tax_total = tax_total.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

    salary_match = _match_salary_combination(salary_records, salary_match_target)
    if not salary_match['matched'] or not salary_match['matched_records']:
        raise ValueError(f'{company} A3 未找到匹配工资金额的银行流水组合')

    credit_records = salary_match['matched_records'] + [tax_record]
    posting_date = _get_last_bank_date(credit_records)
    text = f'支付{_period_text(payroll_year, payroll_month)}工作期间工资及劳务费'
    period_text = _period_text(payroll_year, payroll_month)

    rows = []
    for labor_row in labor_rows:
        rows.append(
            _make_voucher_row(
                'A3',
                company_code,
                posting_date,
                '6601110001',
                labor_row['amount'],
                f'支付{period_text}工作期间劳务费',
                40,
                labor_row['cost_center'],
                labor_row['order'],
            )
        )

    if tax_total != Decimal('0.00'):
        rows.append(
            _make_voucher_row(
                'A3',
                company_code,
                posting_date,
                '2221070000',
                tax_total,
                f'支付{period_text}工作期间个人所得税',
                40,
            )
        )
    if salary_total != Decimal('0.00'):
        rows.append(
            _make_voucher_row(
                'A3',
                company_code,
                posting_date,
                '2211010001',
                salary_total,
                f'支付{period_text}工作期间薪酬',
                40,
            )
        )

    for record in salary_match['matched_records']:
        credit_row = _make_voucher_row(
            'A3',
            company_code,
            posting_date,
            record['bank_subject'],
            record['outgoing_amt'],
            text,
            50,
        )
        credit_row['P'] = '204'
        rows.append(
            credit_row
        )

    tax_credit_row = _make_voucher_row(
        'A3',
        company_code,
        posting_date,
        tax_record['bank_subject'],
        tax_record['outgoing_amt'],
        f'支付{period_text}工作期间个人所得税',
        50,
    )
    tax_credit_row['P'] = '204'
    rows.append(tax_credit_row)

    return rows


def _build_a4_rows(company, company_code, ws, total_row_idx, payroll_year, payroll_month, timesheet_context):
    posting_year, posting_month = _shift_month(payroll_year, payroll_month, 1)
    posting_date = _month_end_date(posting_year, posting_month)
    period_text = _period_text(payroll_year, payroll_month)

    grouped_amounts = {}
    debit_totals = {}
    salary_total = Decimal('0')
    tax_total = Decimal('0')
    other_total = Decimal('0')

    for row_idx in range(3, ws.max_row + 1):
        if _is_total_row(ws.cell(row_idx, 1).value):
            break

        item_type = _normalize_text(ws.cell(row_idx, 4).value)
        if item_type == '劳务费':
            continue

        dept = _normalize_text(ws.cell(row_idx, 2).value)
        project = _normalize_text(ws.cell(row_idx, 3).value)
        cost_center = _format_code(ws.cell(row_idx, 19).value)
        order = _resolve_a4_internal_order(
            timesheet_context,
            company,
            dept,
            project,
            _format_code(ws.cell(row_idx, 20).value),
        )
        key = (cost_center, order)
        bucket = grouped_amounts.setdefault(key, {})

        for col_idx, account, label in ACCRUAL_DEBIT_SPECS:
            amount = _to_decimal(ws.cell(row_idx, col_idx).value)
            if amount == Decimal('0'):
                continue
            bucket[account] = bucket.get(account, Decimal('0')) + amount
            debit_totals[account] = debit_totals.get(account, Decimal('0')) + amount

        if item_type == '薪资':
            salary_total += _to_decimal(ws.cell(row_idx, 17).value)
            tax_total += _to_decimal(ws.cell(row_idx, 15).value)
            other_total += _to_decimal(ws.cell(row_idx, 16).value)

    salary_total = salary_total.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    tax_total = tax_total.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    other_total = other_total.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

    rows = []
    sorted_groups = sorted(grouped_amounts.items())
    rounded_grouped_amounts = {}
    account_last_keys = {}
    rounded_account_totals = {account: Decimal('0.00') for _, account, _ in ACCRUAL_DEBIT_SPECS}

    for key, account_map in sorted_groups:
        rounded_account_map = {}
        for _, account, _ in ACCRUAL_DEBIT_SPECS:
            raw_amount = account_map.get(account, Decimal('0'))
            rounded_amount = raw_amount.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            rounded_account_map[account] = rounded_amount
            rounded_account_totals[account] += rounded_amount
            if raw_amount != Decimal('0'):
                account_last_keys[account] = key
        rounded_grouped_amounts[key] = rounded_account_map

    for _, account, _ in ACCRUAL_DEBIT_SPECS:
        target_total = debit_totals.get(account, Decimal('0')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        diff = (target_total - rounded_account_totals.get(account, Decimal('0.00'))).quantize(Decimal('0.01'))
        if diff != Decimal('0.00') and account in account_last_keys:
            rounded_grouped_amounts[account_last_keys[account]][account] += diff

    for (cost_center, order), _ in sorted_groups:
        account_map = rounded_grouped_amounts[(cost_center, order)]
        for _, account, label in ACCRUAL_DEBIT_SPECS:
            amount = account_map.get(account, Decimal('0.00'))
            if amount == Decimal('0.00'):
                continue
            rows.append(
                _make_voucher_row(
                    'A4',
                    company_code,
                    posting_date,
                    account,
                    amount,
                    f'实际入账{period_text}工作期间{label}',
                    40,
                    cost_center,
                    order,
                )
            )

    for col_idx, account, label in ACCRUAL_CREDIT_SPECS:
        amount = _to_money(ws.cell(total_row_idx, col_idx).value)
        if amount == Decimal('0.00'):
            continue
        rows.append(
            _make_voucher_row(
                'A4',
                company_code,
                posting_date,
                account,
                amount,
                f'实际入账{period_text}工作期间{label}',
                50,
            )
        )

    if tax_total != Decimal('0.00'):
        rows.append(
            _make_voucher_row(
                'A4',
                company_code,
                posting_date,
                '2221070000',
                tax_total,
                f'实际入账{period_text}工作期间应交个税',
                50,
            )
        )
    if salary_total != Decimal('0.00'):
        rows.append(
            _make_voucher_row(
                'A4',
                company_code,
                posting_date,
                '2211010001',
                salary_total,
                f'实际入账{period_text}工作期间实发工资',
                50,
            )
        )
    if other_total != Decimal('0.00'):
        rows.append(
            _make_voucher_row(
                'A4',
                company_code,
                posting_date,
                '2211010001',
                other_total,
                f'实际入账{period_text}工作期间其他项目',
                50,
            )
        )

    return rows


def _collect_a5_base_data(ws):
    grouped_amounts = {}
    debit_totals = {}
    dept_by_cost_center = {}
    salary_total = Decimal('0')
    tax_total = Decimal('0')
    other_total = Decimal('0')

    for row_idx in range(3, ws.max_row + 1):
        if _is_total_row(ws.cell(row_idx, 1).value):
            break

        item_type = _normalize_text(ws.cell(row_idx, 4).value)
        if item_type == '劳务费':
            continue

        dept = _normalize_department_for_order(ws.cell(row_idx, 2).value)
        cost_center = _format_code(ws.cell(row_idx, 19).value)
        if cost_center and dept and cost_center not in dept_by_cost_center:
            dept_by_cost_center[cost_center] = dept

        bucket = grouped_amounts.setdefault(cost_center, {})
        for col_idx, account, _ in ACCRUAL_DEBIT_SPECS:
            amount = _to_decimal(ws.cell(row_idx, col_idx).value)
            if amount == Decimal('0'):
                continue
            bucket[account] = bucket.get(account, Decimal('0')) + amount
            debit_totals[account] = debit_totals.get(account, Decimal('0')) + amount

        if item_type == '薪资':
            salary_total += _to_decimal(ws.cell(row_idx, 17).value)
            tax_total += _to_decimal(ws.cell(row_idx, 15).value)
            other_total += _to_decimal(ws.cell(row_idx, 16).value)

    return {
        'grouped_amounts': grouped_amounts,
        'debit_totals': debit_totals,
        'dept_by_cost_center': dept_by_cost_center,
        'salary_total': salary_total.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP),
        'tax_total': tax_total.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP),
        'other_total': other_total.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP),
    }


def _build_a5_rows(company, company_code, ws, total_row_idx, payroll_year, payroll_month):
    accrual_data = _collect_a5_base_data(ws)
    posting_year, posting_month = _shift_month(payroll_year, payroll_month, 1)
    posting_date = _month_end_date(posting_year, posting_month)
    period_text = _period_text(posting_year, posting_month)

    rows = []
    sorted_groups = sorted(accrual_data['grouped_amounts'].items())
    rounded_grouped_amounts = {}
    account_last_keys = {}
    rounded_account_totals = {account: Decimal('0.00') for _, account, _ in ACCRUAL_DEBIT_SPECS}

    for key, account_map in sorted_groups:
        rounded_account_map = {}
        for _, account, _ in ACCRUAL_DEBIT_SPECS:
            raw_amount = account_map.get(account, Decimal('0'))
            rounded_amount = raw_amount.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            rounded_account_map[account] = rounded_amount
            rounded_account_totals[account] += rounded_amount
            if raw_amount != Decimal('0'):
                account_last_keys[account] = key
        rounded_grouped_amounts[key] = rounded_account_map

    for _, account, _ in ACCRUAL_DEBIT_SPECS:
        target_total = accrual_data['debit_totals'].get(account, Decimal('0')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        diff = (target_total - rounded_account_totals.get(account, Decimal('0.00'))).quantize(Decimal('0.01'))
        if diff != Decimal('0.00') and account in account_last_keys:
            rounded_grouped_amounts[account_last_keys[account]][account] += diff

    for cost_center, _ in sorted_groups:
        account_map = rounded_grouped_amounts[cost_center]
        for _, account, label in ACCRUAL_DEBIT_SPECS:
            amount = account_map.get(account, Decimal('0.00'))
            if amount == Decimal('0.00'):
                continue
            rows.append(
                _make_voucher_row(
                    'A5',
                    company_code,
                    posting_date,
                    account,
                    amount,
                    f'计提{period_text}工作期间{label}',
                    40,
                    cost_center,
                )
            )

    for col_idx, account, label in ACCRUAL_CREDIT_SPECS:
        amount = _to_money(ws.cell(total_row_idx, col_idx).value)
        if amount == Decimal('0.00'):
            continue
        rows.append(
            _make_voucher_row(
                'A5',
                company_code,
                posting_date,
                account,
                amount,
                f'计提{period_text}工作期间{label}',
                50,
            )
        )

    if accrual_data['tax_total'] != Decimal('0.00'):
        rows.append(
            _make_voucher_row(
                'A5',
                company_code,
                posting_date,
                '2221070000',
                accrual_data['tax_total'],
                f'计提{period_text}工作期间应交个税',
                50,
            )
        )
    if accrual_data['salary_total'] != Decimal('0.00'):
        rows.append(
            _make_voucher_row(
                'A5',
                company_code,
                posting_date,
                '2211010001',
                accrual_data['salary_total'],
                f'计提{period_text}工作期间实发工资',
                50,
            )
        )
    if accrual_data['other_total'] != Decimal('0.00'):
        rows.append(
            _make_voucher_row(
                'A5',
                company_code,
                posting_date,
                '2211010001',
                accrual_data['other_total'],
                f'计提{period_text}工作期间其他项目',
                50,
            )
        )

    return rows


def _build_a6_rows(company, company_code, ws, payroll_year, payroll_month, timesheet_context):
    accrual_data = _collect_a5_base_data(ws)
    posting_year, posting_month = _shift_month(payroll_year, payroll_month, 1)
    posting_date = _month_end_date(posting_year, posting_month)
    text = f'根据研发工时分摊{posting_year}年{posting_month}月研发费用'

    rows = []
    alloc_rows_by_key = {}

    for cost_center, account_map in sorted(accrual_data['grouped_amounts'].items()):
        dept = accrual_data['dept_by_cost_center'].get(cost_center, '')
        if not dept or dept == 'OSD运营支持部':
            continue

        alloc_bucket = timesheet_context['allocation_index'].get((posting_year, posting_month, company, dept))
        if not alloc_bucket:
            raise ValueError(f'{company} {dept} 在 {posting_year}年{posting_month}月 未找到 K>0 的工时项目，无法生成 A6')

        order_hours = sorted(alloc_bucket.items())
        total_hours = sum((hours for _, hours in order_hours), Decimal('0'))
        if total_hours <= Decimal('0'):
            raise ValueError(f'{company} {dept} 在 {posting_year}年{posting_month}月 工时合计为 0，无法生成 A6')

        for _, account, _ in ACCRUAL_DEBIT_SPECS:
            amount = account_map.get(account, Decimal('0')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            if amount == Decimal('0.00'):
                continue

            rows.append(
                _make_voucher_row(
                    'A6',
                    company_code,
                    posting_date,
                    account,
                    amount,
                    text,
                    50,
                    cost_center,
                )
            )

            rounded_total = Decimal('0.00')
            last_key = None
            for order, hours in order_hours:
                alloc_amount = (amount * hours / total_hours).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                key = (cost_center, order, account)
                alloc_rows_by_key[key] = alloc_rows_by_key.get(key, Decimal('0.00')) + alloc_amount
                rounded_total += alloc_amount
                last_key = key

            diff = (amount - rounded_total).quantize(Decimal('0.01'))
            if diff != Decimal('0.00') and last_key is not None:
                alloc_rows_by_key[last_key] += diff

    for (cost_center, order, account), amount in sorted(alloc_rows_by_key.items()):
        amount = amount.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        if amount == Decimal('0.00'):
            continue
        rows.append(
            _make_voucher_row(
                'A6',
                company_code,
                posting_date,
                account,
                amount,
                text,
                40,
                cost_center,
                order,
            )
        )

    return rows


def _bonus_allocation_months(year, month):
    if month in (3, 6, 9, 12):
        return list(range(1, month + 1))
    return [month]


def _get_bonus_allocation_bucket(timesheet_context, year, month, company, dept):
    bucket = {}
    for alloc_month in _bonus_allocation_months(year, month):
        month_bucket = timesheet_context['allocation_index'].get((year, alloc_month, company, dept), {})
        for order, hours in month_bucket.items():
            bucket[order] = bucket.get(order, Decimal('0')) + hours
    return bucket


def _build_a7_rows(company, company_code, payroll_year, payroll_month, bonus_context):
    bonus_data = _build_bonus_department_amounts(company, payroll_year, payroll_month, bonus_context)
    posting_date = _month_end_date(bonus_data['posting_year'], bonus_data['posting_month'])
    label = _bonus_label_for_text(bonus_data['posting_year'], bonus_data['posting_month'])
    text = f'计提{label}年终奖'

    rows = []
    total_amount = Decimal('0.00')
    for item in bonus_data['rows']:
        rows.append(
            _make_voucher_row(
                'A7',
                company_code,
                posting_date,
                '6601010003',
                item['amount'],
                text,
                40,
                item['cost_center'],
            )
        )
        total_amount += item['amount']

    if total_amount != Decimal('0.00'):
        rows.append(
            _make_voucher_row(
                'A7',
                company_code,
                posting_date,
                '2211010001',
                total_amount,
                text,
                50,
            )
        )

    return rows


def _build_a8_rows(company, company_code, payroll_year, payroll_month, bonus_context, timesheet_context):
    bonus_data = _build_bonus_department_amounts(company, payroll_year, payroll_month, bonus_context)
    posting_date = _month_end_date(bonus_data['posting_year'], bonus_data['posting_month'])
    label = _bonus_label_for_text(bonus_data['posting_year'], bonus_data['posting_month'])
    text = f'根据研发工时分摊{label}年终奖'

    rows = []
    alloc_rows_by_key = {}

    for item in bonus_data['rows']:
        dept = item['dept']
        if dept == 'OSD运营支持部':
            continue

        amount = item['amount']
        cost_center = item['cost_center']
        alloc_bucket = _get_bonus_allocation_bucket(
            timesheet_context,
            bonus_data['posting_year'],
            bonus_data['posting_month'],
            company,
            dept,
        )
        if not alloc_bucket:
            raise ValueError(f'{company} {dept} 在奖金分摊口径下未找到可用工时，无法生成 A8')

        order_hours = sorted((order, hours) for order, hours in alloc_bucket.items() if hours > Decimal('0'))
        total_hours = sum((hours for _, hours in order_hours), Decimal('0'))
        if total_hours <= Decimal('0'):
            raise ValueError(f'{company} {dept} 在奖金分摊口径下工时合计为 0，无法生成 A8')

        rows.append(
            _make_voucher_row(
                'A8',
                company_code,
                posting_date,
                '6601010003',
                amount,
                text,
                50,
                cost_center,
            )
        )

        rounded_total = Decimal('0.00')
        last_key = None
        for order, hours in order_hours:
            alloc_amount = (amount * hours / total_hours).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            key = (cost_center, order)
            alloc_rows_by_key[key] = alloc_rows_by_key.get(key, Decimal('0.00')) + alloc_amount
            rounded_total += alloc_amount
            last_key = key

        diff = (amount - rounded_total).quantize(Decimal('0.01'))
        if diff != Decimal('0.00') and last_key is not None:
            alloc_rows_by_key[last_key] += diff

    for (cost_center, order), amount in sorted(alloc_rows_by_key.items()):
        amount = amount.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        if amount == Decimal('0.00'):
            continue
        rows.append(
            _make_voucher_row(
                'A8',
                company_code,
                posting_date,
                '6601010003',
                amount,
                text,
                40,
                cost_center,
                order,
            )
        )

    return rows


def _build_a9_rows(company_code, payroll_year, payroll_month, co_context):
    posting_year, posting_month = _shift_month(payroll_year, payroll_month, 1)
    posting_date = _month_end_date(posting_year, posting_month)
    text = f'CO工单分摊{posting_year}年{posting_month}月'

    rows = []
    debit_allocations = {}
    allocation_rows = co_context['allocation_rows']
    total_ratio_base = co_context['total_ratio_base']

    for item in co_context['expense_rows']:
        credit_row = _make_voucher_row(
            'A9',
            company_code,
            posting_date,
            item['source_account'],
            item['amount'],
            text,
            50,
            CO_CREDIT_COST_CENTER,
            CO_CREDIT_ORDER,
        )
        credit_row['O'] = 'x'
        rows.append(
            credit_row
        )

        rounded_total = Decimal('0.00')
        last_key = None
        for alloc in allocation_rows:
            alloc_amount = (item['amount'] * alloc['ratio_base'] / total_ratio_base).quantize(
                Decimal('0.01'),
                rounding=ROUND_HALF_UP,
            )
            key = (item['debit_gl'], alloc['order'])
            debit_allocations[key] = debit_allocations.get(key, Decimal('0.00')) + alloc_amount
            rounded_total += alloc_amount
            last_key = key

        diff = (item['amount'] - rounded_total).quantize(Decimal('0.01'))
        if diff != Decimal('0.00') and last_key is not None:
            debit_allocations[last_key] += diff

    for (debit_gl, order), amount in sorted(debit_allocations.items()):
        amount = amount.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        if amount == Decimal('0.00'):
            continue
        rows.append(
            _make_voucher_row(
                'A9',
                company_code,
                posting_date,
                debit_gl,
                amount,
                text,
                40,
                '',
                order,
            )
        )

    return rows


def _build_a10_rows(company, company_code, payroll_year, payroll_month, expense_context, timesheet_context):
    posting_year, posting_month = _shift_month(payroll_year, payroll_month, 1)
    posting_date = _month_end_date(posting_year, posting_month)
    text = f'根据研发工时分摊{posting_year}年{posting_month}月研发费用'

    rows = []
    alloc_rows_by_key = {}
    grouped_items = {}

    for item in expense_context['rows_by_company'].get(company, []):
        dept = item['dept']
        if dept == 'OSD运营支持部' or 'OSD' in dept:
            continue

        key = (item['gl_account'], dept)
        grouped = grouped_items.get(key)
        if grouped is None:
            grouped_items[key] = {
                'gl_account': item['gl_account'],
                'dept': dept,
                'amount': item['amount'],
                'cost_center': item['cost_center'],
                'allocation_cost_center': item['allocation_cost_center'],
                'credit_order': item['credit_order'],
                'row_indices': [item['row_idx']],
            }
            continue

        if (
            grouped['cost_center'] != item['cost_center']
            or grouped['allocation_cost_center'] != item['allocation_cost_center']
            or grouped['credit_order'] != item['credit_order']
        ):
            raise ValueError(
                f'{company} A10 聚合冲突：科目 {item["gl_account"]} / 部门 {dept} 对应了多个成本中心或订单，'
                f'涉及待分摊费用行 {grouped["row_indices"] + [item["row_idx"]]}'
            )

        grouped['amount'] += item['amount']
        grouped['row_indices'].append(item['row_idx'])

    for item in grouped_items.values():
        item['amount'] = _to_money(item['amount'])
        if item['amount'] == Decimal('0.00'):
            continue
        if item['amount'] < Decimal('0.00'):
            raise ValueError(
                f'{company} A10 聚合后金额为负数：科目 {item["gl_account"]} / 部门 {item["dept"]} / 金额 {item["amount"]}'
            )

        dept = item['dept']
        alloc_bucket = dict(timesheet_context['allocation_index'].get((posting_year, posting_month, company, dept), {}))
        if dept == 'HLD硬件逻辑部':
            alloc_bucket.pop('9201856', None)

        order_hours = sorted((order, hours) for order, hours in alloc_bucket.items() if hours > Decimal('0'))
        if not order_hours and company == '耐数信息' and item['cost_center'] == '20602020':
            company_bucket = {}
            for (year_num, month_num, bucket_company, _), bucket in timesheet_context['allocation_index'].items():
                if year_num != posting_year or month_num != posting_month or bucket_company != company:
                    continue
                for order, hours in bucket.items():
                    if hours > Decimal('0'):
                        company_bucket[order] = company_bucket.get(order, Decimal('0')) + hours
            order_hours = sorted(company_bucket.items())
        if not order_hours:
            raise ValueError(f'{company} {dept} 在 {posting_year}年{posting_month}月 未找到可用于 A10 分摊的工时项目')

        total_hours = sum((hours for _, hours in order_hours), Decimal('0'))
        if total_hours <= Decimal('0'):
            raise ValueError(f'{company} {dept} 在 {posting_year}年{posting_month}月 工时合计为 0，无法生成 A10')

        rows.append(
            _make_voucher_row(
                'A10',
                company_code,
                posting_date,
                item['gl_account'],
                item['amount'],
                text,
                50,
                item['cost_center'],
                item['credit_order'],
                voucher_type='SA',
            )
        )

        allocations = _allocate_amount_by_weights(
            item['amount'],
            [((item['gl_account'], item['allocation_cost_center'], order), hours) for order, hours in order_hours],
        )
        for key, alloc_amount in allocations.items():
            alloc_rows_by_key[key] = alloc_rows_by_key.get(key, Decimal('0.00')) + alloc_amount

    for (gl_account, cost_center, order), amount in sorted(alloc_rows_by_key.items()):
        amount = amount.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        if amount == Decimal('0.00'):
            continue
        rows.append(
            _make_voucher_row(
                'A10',
                company_code,
                posting_date,
                gl_account,
                amount,
                text,
                40,
                cost_center,
                order,
                voucher_type='SA',
            )
        )

    return rows


def _clear_voucher_template_rows(ws):
    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, 18):
            ws.cell(row=row_idx, column=col_idx).value = None


def _write_voucher_rows_to_template(ws, voucher_rows):
    for row_offset, row_data in enumerate(voucher_rows, start=2):
        for col_letter, value in row_data.items():
            ws[f'{col_letter}{row_offset}'] = value


def _validate_voucher_groups(voucher_rows):
    group_totals = {}
    for row in voucher_rows:
        amount = _to_money(row['M'])
        group = row['A']
        totals = group_totals.setdefault(group, {'debit': Decimal('0.00'), 'credit': Decimal('0.00')})
        if str(row['G']) == '40':
            totals['debit'] += amount
        elif str(row['G']) == '50':
            totals['credit'] += amount

    for group, totals in group_totals.items():
        if totals['debit'].quantize(Decimal('0.01')) != totals['credit'].quantize(Decimal('0.01')):
            raise ValueError(f'{group} 借贷不平：借方 {totals["debit"]}，贷方 {totals["credit"]}')


def _validate_cross_group_accounts(voucher_rows):
    account_totals = {}
    other_project_credit = Decimal('0.00')

    for row in voucher_rows:
        account = _format_code(row['I'])
        amount = _to_money(row['M'])
        bucket = account_totals.setdefault(account, {'debit': Decimal('0.00'), 'credit': Decimal('0.00')})
        if str(row['G']) == '40':
            bucket['debit'] += amount
        elif str(row['G']) == '50':
            bucket['credit'] += amount

        if account == '2211010001' and str(row['G']) == '50' and row['N'].endswith('其他项目'):
            other_project_credit += amount

    issues = []
    for account in ACCOUNTS_TO_CROSS_CHECK:
        totals = account_totals.get(account, {'debit': Decimal('0.00'), 'credit': Decimal('0.00')})
        debit_total = totals['debit'].quantize(Decimal('0.01'))
        credit_total = totals['credit'].quantize(Decimal('0.01'))

        if account == '2211010001':
            adjusted_credit = (credit_total - other_project_credit).quantize(Decimal('0.01'))
            if debit_total != adjusted_credit:
                issues.append(
                    f'{account} 跨凭证不平：借方 {debit_total}，贷方(扣除其他项目后) {adjusted_credit}'
                )
        elif debit_total != credit_total:
            issues.append(f'{account} 跨凭证不平：借方 {debit_total}，贷方 {credit_total}')

    return issues


def _generate_voucher_files(
    base_dir,
    company_output_paths,
    bank_records,
    payroll_year,
    payroll_month,
    log,
    timesheet_context,
    bonus_context,
    run_options=None,
):
    run_options = normalize_run_options(run_options)
    template_path = _get_voucher_template_path(base_dir)
    template_bytes = open(template_path, 'rb').read()
    voucher_paths = {}
    validation_summary = {}
    co_context = None
    shared_expense_context = None
    co_workorder_path = _get_co_workorder_path(base_dir, payroll_year, payroll_month)
    shared_expense_path = _get_shared_expense_path(base_dir, payroll_year, payroll_month)

    if requires_co_data(run_options):
        if os.path.isfile(co_workorder_path):
            log('正在读取 CO工单分摊 数据…')
            co_context = _load_co_workorder_context(co_workorder_path)
            log(
                f'CO工单分摊读取完成：{len(co_context["allocation_rows"])} 个参与分摊订单，'
                f'{len(co_context["expense_rows"])} 行待分摊费用',
                'ok',
            )
        else:
            log(f'未找到 CO工单分摊 文件，已跳过 A9：{co_workorder_path}', 'warn')

    if requires_shared_expense_data(run_options):
        if os.path.isfile(shared_expense_path):
            log('正在读取 待分摊费用 数据…')
            cost_center_map = _load_mappings(_get_mapping_path(base_dir))
            shared_expense_context = _load_shared_expense_context(shared_expense_path, cost_center_map)
            log(
                f'待分摊费用读取完成：2050 {len(shared_expense_context["rows_by_company"]["耐数电子"])} 行，'
                f'2060 {len(shared_expense_context["rows_by_company"]["耐数信息"])} 行',
                'ok',
            )
        else:
            log(f'未找到 待分摊费用 文件，已跳过 A10：{shared_expense_path}', 'warn')

    for company, company_output_path in company_output_paths.items():
        if not run_options.wants_company(company):
            continue

        company_code = COMPANY_NAME_TO_CODE[company]
        company_wb = load_workbook(company_output_path, data_only=True)
        company_ws = company_wb[company_wb.sheetnames[0]]
        total_row_idx = _find_total_row(company_ws)
        if total_row_idx is None:
            raise ValueError(f'{company} 工资单中未找到总计行')

        actual_voucher_rows = []
        if run_options.wants_voucher('A1'):
            actual_voucher_rows.extend(
                _build_a1_rows(company, company_code, company_ws, total_row_idx, bank_records, payroll_year, payroll_month)
            )
        if run_options.wants_voucher('A2'):
            actual_voucher_rows.extend(_build_a2_rows(company, company_code, bank_records, payroll_year, payroll_month))
        if run_options.wants_voucher('A3'):
            actual_voucher_rows.extend(
                _build_a3_rows(company, company_code, company_ws, bank_records, payroll_year, payroll_month)
            )
        if run_options.wants_voucher('A4'):
            actual_voucher_rows.extend(
                _build_a4_rows(company, company_code, company_ws, total_row_idx, payroll_year, payroll_month, timesheet_context)
            )

        accrual_voucher_rows = []
        if run_options.wants_voucher('A5'):
            accrual_voucher_rows.extend(
                _build_a5_rows(company, company_code, company_ws, total_row_idx, payroll_year, payroll_month)
            )
        if run_options.wants_voucher('A6'):
            accrual_voucher_rows.extend(
                _build_a6_rows(company, company_code, company_ws, payroll_year, payroll_month, timesheet_context)
            )

        bonus_voucher_rows = []
        if bonus_context is not None and run_options.wants_voucher('A7'):
            bonus_voucher_rows.extend(_build_a7_rows(company, company_code, payroll_year, payroll_month, bonus_context))
        if bonus_context is not None and run_options.wants_voucher('A8'):
            bonus_voucher_rows.extend(
                _build_a8_rows(company, company_code, payroll_year, payroll_month, bonus_context, timesheet_context)
            )

        co_voucher_rows = []
        if company == '耐数电子' and co_context is not None and run_options.wants_voucher('A9'):
            co_voucher_rows.extend(_build_a9_rows(company_code, payroll_year, payroll_month, co_context))

        rd_voucher_rows = []
        if shared_expense_context is not None and run_options.wants_voucher('A10'):
            company_expense_rows = shared_expense_context['rows_by_company'].get(company, [])
            if company_expense_rows:
                rd_voucher_rows.extend(
                    _build_a10_rows(company, company_code, payroll_year, payroll_month, shared_expense_context, timesheet_context)
                )

        _validate_voucher_groups(actual_voucher_rows)
        _validate_voucher_groups(accrual_voucher_rows)
        _validate_voucher_groups(bonus_voucher_rows)
        if co_voucher_rows:
            _validate_voucher_groups(co_voucher_rows)
        if rd_voucher_rows:
            _validate_voucher_groups(rd_voucher_rows)

        cross_group_checked = all(run_options.wants_voucher(voucher) for voucher in ACTUAL_VOUCHERS)
        cross_group_issues = _validate_cross_group_accounts(actual_voucher_rows) if cross_group_checked else []

        actual_voucher_output_path = ''
        if actual_voucher_rows:
            actual_voucher_wb = load_workbook(io.BytesIO(template_bytes))
            actual_voucher_ws = actual_voucher_wb[actual_voucher_wb.sheetnames[0]]
            _clear_voucher_template_rows(actual_voucher_ws)
            _write_voucher_rows_to_template(actual_voucher_ws, actual_voucher_rows)
            actual_voucher_output_path = _build_voucher_output_path(base_dir, company, payroll_year, payroll_month)
            actual_voucher_wb.save(actual_voucher_output_path)
            voucher_paths[f'{company}（实际）'] = actual_voucher_output_path
            log(f'{company} 实际凭证文件已生成：{actual_voucher_output_path}', 'ok')

        accrual_voucher_output_path = ''
        if accrual_voucher_rows:
            accrual_voucher_wb = load_workbook(io.BytesIO(template_bytes))
            accrual_voucher_ws = accrual_voucher_wb[accrual_voucher_wb.sheetnames[0]]
            _clear_voucher_template_rows(accrual_voucher_ws)
            _write_voucher_rows_to_template(accrual_voucher_ws, accrual_voucher_rows)
            accrual_voucher_output_path = _build_accrual_voucher_output_path(base_dir, company, payroll_year, payroll_month)
            accrual_voucher_wb.save(accrual_voucher_output_path)
            voucher_paths[f'{company}（计提）'] = accrual_voucher_output_path
            log(f'{company} 计提凭证文件已生成：{accrual_voucher_output_path}', 'ok')

        bonus_voucher_output_path = ''
        if bonus_voucher_rows:
            bonus_voucher_wb = load_workbook(io.BytesIO(template_bytes))
            bonus_voucher_ws = bonus_voucher_wb[bonus_voucher_wb.sheetnames[0]]
            _clear_voucher_template_rows(bonus_voucher_ws)
            _write_voucher_rows_to_template(bonus_voucher_ws, bonus_voucher_rows)
            bonus_voucher_output_path = _build_bonus_voucher_output_path(base_dir, company, payroll_year, payroll_month)
            bonus_voucher_wb.save(bonus_voucher_output_path)
            voucher_paths[f'{company}（年终奖）'] = bonus_voucher_output_path
            log(f'{company} 年终奖凭证文件已生成：{bonus_voucher_output_path}', 'ok')

        if co_voucher_rows:
            co_voucher_wb = load_workbook(io.BytesIO(template_bytes))
            co_voucher_ws = co_voucher_wb[co_voucher_wb.sheetnames[0]]
            _clear_voucher_template_rows(co_voucher_ws)
            _write_voucher_rows_to_template(co_voucher_ws, co_voucher_rows)
            co_voucher_output_path = _build_co_voucher_output_path(base_dir, payroll_year, payroll_month)
            co_voucher_wb.save(co_voucher_output_path)
            voucher_paths[f'{company}（CO工单分摊）'] = co_voucher_output_path
            log(f'{company} CO工单分摊凭证文件已生成：{co_voucher_output_path}', 'ok')

        if rd_voucher_rows:
            rd_voucher_wb = load_workbook(io.BytesIO(template_bytes))
            rd_voucher_ws = rd_voucher_wb[rd_voucher_wb.sheetnames[0]]
            _clear_voucher_template_rows(rd_voucher_ws)
            _write_voucher_rows_to_template(rd_voucher_ws, rd_voucher_rows)
            rd_voucher_output_path = _build_rd_allocation_voucher_output_path(base_dir, company, payroll_year, payroll_month)
            rd_voucher_wb.save(rd_voucher_output_path)
            voucher_paths[f'{company}（研发费用分摊）'] = rd_voucher_output_path
            log(f'{company} 研发费用分摊凭证文件已生成：{rd_voucher_output_path}', 'ok')

        validation_summary[company] = {
            'a4_balanced': True,
            'a5_balanced': True,
            'a6_balanced': True,
            'a7_balanced': True,
            'a8_balanced': True,
            'a9_balanced': True,
            'a10_balanced': True,
            'group_balances': {voucher: True for voucher in VOUCHER_DISPLAY_ORDER if run_options.wants_voucher(voucher)},
            'cross_group_issues': cross_group_issues,
            'cross_group_checked': cross_group_checked,
        }

        if cross_group_checked and cross_group_issues:
            for issue in cross_group_issues:
                log(f'{company} 跨凭证校验异常：{issue}', 'warn')
        elif cross_group_checked:
            log(f'{company} A1-A4 科目对冲校验通过', 'ok')
        else:
            log(f'{company} 未选择完整的 A1-A4，已跳过跨凭证对冲校验', 'warn')

    return voucher_paths, validation_summary


def _clear_bank_summary(ws, end_row=12):
    for row_idx in range(2, end_row + 1):
        for col_idx in range(22, 30):
            ws.cell(row=row_idx, column=col_idx).value = None


def _find_total_row(ws):
    for row_idx in range(1, ws.max_row + 1):
        if _is_total_row(ws.cell(row=row_idx, column=1).value):
            return row_idx
    return None


def _recalculate_total_row(ws, total_row_idx):
    ws.cell(row=total_row_idx, column=1).value = '总计'
    for col_idx in range(2, 21):
        ws.cell(row=total_row_idx, column=col_idx).value = None

    for col_idx in range(5, 19):
        total_value = sum((_to_decimal(ws.cell(row=row_idx, column=col_idx).value) for row_idx in range(3, total_row_idx)), Decimal('0'))
        ws.cell(row=total_row_idx, column=col_idx).value = float(
            total_value.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        )


def _save_company_workbooks(base_wb, base_dir, input_path, payroll_year, payroll_month, company_results, log):
    buffer = io.BytesIO()
    base_wb.save(buffer)
    workbook_bytes = buffer.getvalue()
    output_paths = {}

    for company, bank_results in company_results.items():
        company_output_path = _build_company_output_path(base_dir, company, payroll_year, payroll_month, input_path)
        company_wb = load_workbook(io.BytesIO(workbook_bytes))
        ws = company_wb[company_wb.sheetnames[0]]

        total_row_idx = _find_total_row(ws)
        if total_row_idx is None:
            raise ValueError('拆分公司文件时未找到“总计”行')

        if ws.max_row > total_row_idx:
            ws.delete_rows(total_row_idx + 1, ws.max_row - total_row_idx)

        for row_idx in range(total_row_idx - 1, 2, -1):
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

        _recalculate_total_row(ws, new_total_row_idx)
        _clear_bank_summary(ws)
        header_ref = ws['R2'] if ws.max_column >= 18 else ws['Q2']
        _write_bank_summary(ws, bank_results, header_ref)

        company_wb.save(company_output_path)
        output_paths[company] = company_output_path
        log(f'{company} 文件已生成：{company_output_path}', 'ok')

    return output_paths


def fill_first_sheet_ab(input_path, base_dir, mapping_path, bank_dir, log, run_options=None):
    run_options = normalize_run_options(run_options)
    log('正在读取原始工资单…')
    wb = load_workbook(input_path)
    ws = wb[wb.sheetnames[0]]
    log(f'首个工作表：{ws.title}，共 {ws.max_row} 行')
    log('正在读取 Mapping 表…')
    cost_center_map = _load_mappings(mapping_path)
    log('Mapping 表读取完成', 'ok')
    timesheet_path = _get_timesheet_path(base_dir)
    bonus_path = _get_bonus_path(base_dir)
    if not os.path.isfile(timesheet_path):
        raise FileNotFoundError(f'找不到工时数据文件（请先放置）：{timesheet_path}')
    if requires_bonus_data(run_options) and not os.path.isfile(bonus_path):
        raise FileNotFoundError(f'找不到奖金数据文件（请先放置）：{bonus_path}')
    log('正在匹配工时数据中的内部订单…')
    _, timesheet_match_summary = _match_timesheet_internal_orders(timesheet_path, save_changes=False)
    if timesheet_match_summary['unmatched_rows']:
        raise ValueError(
            f'工时数据仍有 {len(timesheet_match_summary["unmatched_rows"])} 行项目未匹配到内部订单，无法继续后续处理'
        )
    timesheet_context = _load_timesheet_match_context(timesheet_path)
    bonus_context = _load_bonus_context(bonus_path, cost_center_map) if requires_bonus_data(run_options) else None
    log(
        f'工时内部订单匹配完成：已匹配 {timesheet_match_summary["matched_rows"]} 行，'
        f'补充逻辑命中 {timesheet_match_summary["used_extra_rows"]} 行',
        'ok',
    )
    bank_records = {}
    if requires_bank_data(run_options):
        log('正在读取银行流水…')
        bank_records = _load_bank_records(bank_dir)
        log('银行流水读取完成', 'ok')
    else:
        log('本次未选择 A1-A3，已跳过银行流水读取', 'warn')

    last_a = None
    last_b = None
    last_a_style_cell = None
    last_b_style_cell = None
    fill_a = 0
    fill_b = 0
    total_row_idx = ws.max_row + 1
    issue_rows = {}
    validation_issue_count = 0
    cost_center_issue_count = 0
    internal_order_issue_count = 0
    payroll_summary = {}
    rows_by_company = {}

    for row_idx in range(1, ws.max_row + 1):
        cell_a = ws.cell(row=row_idx, column=1)
        cell_b = ws.cell(row=row_idx, column=2)

        if _is_total_row(cell_a.value):
            log(f'在第 {row_idx} 行识别到“总计”，从该行开始停止填充')
            total_row_idx = row_idx
            break

        if _is_blank(cell_a.value):
            if last_a is not None:
                cell_a.value = last_a
                if last_a_style_cell is not None:
                    _copy_cell_style(last_a_style_cell, cell_a)
                fill_a += 1
        else:
            last_a = cell_a.value
            last_a_style_cell = cell_a

        if _is_blank(cell_b.value):
            if last_b is not None:
                cell_b.value = last_b
                if last_b_style_cell is not None:
                    _copy_cell_style(last_b_style_cell, cell_b)
                fill_b += 1
        else:
            last_b = cell_b.value
            last_b_style_cell = cell_b

    log(f'A 列填充 {fill_a} 个空白单元格', 'ok')
    log(f'B 列填充 {fill_b} 个空白单元格', 'ok')

    header_ref = ws['R2'] if ws.max_column >= 18 else ws['Q2']
    data_style_col = 18 if ws.max_column >= 18 else 17
    ws['S2'].value = '成本中心'
    _copy_cell_style(header_ref, ws['S2'])
    ws['T2'].value = '内部订单'
    _copy_cell_style(header_ref, ws['T2'])

    if ws.column_dimensions['R'].width:
        ws.column_dimensions['S'].width = ws.column_dimensions['R'].width
    if ws.column_dimensions['Q'].width:
        ws.column_dimensions['T'].width = ws.column_dimensions['Q'].width

    for row_idx in range(3, total_row_idx):
        row_values = [ws.cell(row=row_idx, column=col).value for col in range(1, 19)]
        if all(_is_blank(value) for value in row_values):
            continue

        row_issues = []
        company = _normalize_text(ws.cell(row=row_idx, column=1).value)
        dept = _normalize_text(ws.cell(row=row_idx, column=2).value)
        project = _normalize_text(ws.cell(row=row_idx, column=3).value)
        item_type = _normalize_text(ws.cell(row=row_idx, column=4).value)
        rows_by_company.setdefault(company, []).append(row_idx)

        expected_actual = _to_decimal(ws.cell(row=row_idx, column=5).value)
        for col_idx in range(11, 17):
            expected_actual -= _to_decimal(ws.cell(row=row_idx, column=col_idx).value)
        expected_actual = expected_actual.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        actual_value = _to_money(ws.cell(row=row_idx, column=17).value)
        company_summary = payroll_summary.setdefault(
            company,
            {
                'salary': Decimal('0'),
                'fund': Decimal('0'),
                'tax': Decimal('0'),
                'social': Decimal('0'),
            },
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
            validation_issue_count += 1
            row_issues.append(f'实发校验失败(Q={actual_value}, 计算值={expected_actual})')

        company_cost_center_map = cost_center_map.get(company)
        if company_cost_center_map is None:
            row_issues.append(f'无法识别公司：{company or "空"}')
            cost_center_value = ''
            internal_order_value = ''
            cost_center_issue_count += 1
            internal_order_issue_count += 1
        else:
            cost_center_value = company_cost_center_map.get(dept, '')
            if not cost_center_value:
                cost_center_issue_count += 1
                row_issues.append(f'成本中心未匹配：{dept or "空"}')

            internal_order_value, project_for_match = _lookup_internal_order_from_timesheet(
                timesheet_context,
                company,
                dept,
                project,
            )
            if not internal_order_value:
                internal_order_issue_count += 1
                row_issues.append(f'内部订单未匹配：{project_for_match or "空"}')

        s_cell = ws.cell(row=row_idx, column=19)
        t_cell = ws.cell(row=row_idx, column=20)
        ref_cell = ws.cell(row=row_idx, column=data_style_col)
        _copy_cell_style(ref_cell, s_cell)
        _copy_cell_style(ref_cell, t_cell)
        s_cell.value = cost_center_value
        t_cell.value = internal_order_value

        if row_issues:
            issue_rows[row_idx] = row_issues

    for row_idx in issue_rows:
        _mark_row_red(ws, row_idx, 20)

    if internal_order_issue_count:
        raise ValueError(
            f'工资单中有 {internal_order_issue_count} 行内部订单(T列)未匹配，已停止处理，请修正工时数据或工资单项目后再运行'
        )

    if issue_rows:
        log(f'共发现 {len(issue_rows)} 行异常，已在输出文件中标红', 'warn')
    else:
        log('所有数据校验和匹配均通过', 'ok')

    payroll_year, payroll_month = _extract_payroll_period(input_path)
    next_year, next_month = _shift_month(payroll_year, payroll_month, 1)
    bank_results = []
    company_bank_results = {}
    bank_issue_count = 0
    bank_salary_issue_count = 0
    bank_fund_issue_count = 0
    bank_tax_issue_count = 0
    bank_social_issue_count = 0
    bank_issue_companies = set()

    selected_companies = [company for company in payroll_summary if run_options.wants_company(company)]
    if requires_bank_data(run_options):
        log('开始核对银行流水…')
        for company in selected_companies:
            summary = payroll_summary[company]
            company_code = COMPANY_NAME_TO_CODE.get(company)
            if not company_code:
                continue

            salary_tax_bank = _summarize_bank_records(bank_records, company_code, next_year, next_month)
            fund_bank = _summarize_bank_records(bank_records, company_code, payroll_year, payroll_month)
            salary_match = _match_salary_combination(
                salary_tax_bank['salary_records'],
                summary['salary'].quantize(Decimal('0.01'), rounding=ROUND_HALF_UP),
            )

            checks = [
                ('salary', '实发工资', summary['salary'].quantize(Decimal('0.01'), rounding=ROUND_HALF_UP), salary_match['matched_amount'], f'{next_year}-{next_month:02d}'),
                ('fund', '公积金', summary['fund'].quantize(Decimal('0.01'), rounding=ROUND_HALF_UP), fund_bank['fund'], f'{payroll_year}-{payroll_month:02d}'),
                ('tax', '个税', summary['tax'].quantize(Decimal('0.01'), rounding=ROUND_HALF_UP), salary_tax_bank['tax'], f'{next_year}-{next_month:02d}'),
                ('social', '社保', summary['social'].quantize(Decimal('0.01'), rounding=ROUND_HALF_UP), salary_tax_bank['social'], f'{next_year}-{next_month:02d}'),
            ]

            for item_key, item_label, payroll_amount, bank_amount, period_label in checks:
                diff = (payroll_amount - bank_amount).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                passed = diff == Decimal('0.00')
                note = ''

                if item_key == 'salary':
                    passed = salary_match['matched']
                    diff = (payroll_amount - bank_amount).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                    note = salary_match['note']

                if item_key == 'fund' and company_code == '2050' and payroll_year == 2026 and payroll_month == 2:
                    passed = abs(diff) == Decimal('3650.00')
                    if passed:
                        note = '202602 特殊规则：公积金允许绝对差额 3650'

                source_files = salary_tax_bank['files'] if item_key in ('salary', 'tax', 'social') else fund_bank['files']
                if not source_files:
                    note = f'未找到 {period_label} 的匹配银行流水'
                elif not note:
                    note = '来源：' + '、'.join(source_files)
                elif item_key == 'salary':
                    note = note + '；来源：' + '、'.join(source_files)

                if not passed:
                    bank_issue_count += 1
                    bank_issue_companies.add(company)
                    if item_key == 'salary':
                        bank_salary_issue_count += 1
                    elif item_key == 'fund':
                        bank_fund_issue_count += 1
                    elif item_key == 'tax':
                        bank_tax_issue_count += 1
                    elif item_key == 'social':
                        bank_social_issue_count += 1
                    log(f'{company}{item_label}核对异常：工资表={payroll_amount}，流水={bank_amount}，差额={diff}', 'warn')
                else:
                    log(f'{company}{item_label}核对通过：{payroll_amount}', 'ok')

                bank_results.append(
                    {
                        'company': company,
                        'item_label': item_label,
                        'payroll_amount': payroll_amount,
                        'bank_amount': bank_amount,
                        'diff': diff,
                        'period_label': period_label,
                        'passed': passed,
                        'note': note,
                    }
                )
            company_bank_results[company] = [item for item in bank_results if item['company'] == company]
    else:
        for company in selected_companies:
            company_bank_results[company] = []
        log('本次未选择 A1-A3，已跳过银行核对', 'warn')

    _write_bank_summary(ws, bank_results, header_ref)

    for company in bank_issue_companies:
        for row_idx in rows_by_company.get(company, []):
            _mark_row_red(ws, row_idx, 20)

    output_paths = _save_company_workbooks(wb, base_dir, input_path, payroll_year, payroll_month, company_bank_results, log)
    voucher_paths, voucher_validation_summary = _generate_voucher_files(
        base_dir,
        output_paths,
        bank_records,
        payroll_year,
        payroll_month,
        log,
        timesheet_context,
        bonus_context,
        run_options=run_options,
    )

    return {
        'sheet_name': ws.title,
        'row_count': ws.max_row,
        'fill_a': fill_a,
        'fill_b': fill_b,
        'issue_count': len(issue_rows),
        'validation_issue_count': validation_issue_count,
        'cost_center_issue_count': cost_center_issue_count,
        'internal_order_issue_count': internal_order_issue_count,
        'bank_issue_count': bank_issue_count,
        'bank_salary_issue_count': bank_salary_issue_count,
        'bank_fund_issue_count': bank_fund_issue_count,
        'bank_tax_issue_count': bank_tax_issue_count,
        'bank_social_issue_count': bank_social_issue_count,
        'output_paths': output_paths,
        'voucher_paths': voucher_paths,
        'voucher_validation_summary': voucher_validation_summary,
    }


