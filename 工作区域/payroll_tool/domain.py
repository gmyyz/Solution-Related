"""domain: extracted payroll responsibilities; business rules preserved."""
import os
import re
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP, ROUND_DOWN
from .banking import (
    bank_scope,
    _get_treasury_records,
    _load_bank_records,
    _load_bank_records_with_stats,
    _match_bank_record_combination,
    _match_salary_combination,
    _match_treasury_combination,
    _normalize_bank_date_key,
    _summarize_bank_records,
)
from .rules import (
    skips_department,
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


RAW_FILE_PATTERNS = (
    '人力成本研发项目分摊* - to财务-原始.xlsx',
    '人力成本研发项目分摊* - to财务.xlsx',
    '耐数人力成本分摊*-to财务.xlsx',
    '耐数人力成本分摊* - to财务.xlsx',
)


RAW_FILE_PATTERN = RAW_FILE_PATTERNS[0]


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
    '6601010003',
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


EXAMPLE_IMAGE_DIRNAME = '示例截图'


def _allocate_amount_by_weights(total_amount, weighted_items):
    total_amount = _to_money(total_amount)
    if total_amount < 0:
        raise ValueError('分摊金额不能为负数')
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


def _get_bonus_tax_adjustment(company_code, payment_year, payment_month):
    return BONUS_TAX_BY_PAYMENT_PERIOD.get(
        (_format_code(company_code), payment_year, payment_month),
        Decimal('0.00'),
    )


def _select_treasury_records(records):
    treasury_records = _get_treasury_records(records)
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


def _rebalance_expense_debit_rows(rows, preferred_accounts, voucher_label):
    debit_total = Decimal('0.00')
    credit_total = Decimal('0.00')
    for row in rows:
        amount = _to_money(row.get('M'))
        if str(row.get('G')) == '40':
            debit_total += amount
        elif str(row.get('G')) == '50':
            credit_total += amount

    adjustment = (credit_total - debit_total).quantize(Decimal('0.01'))
    if adjustment == Decimal('0.00'):
        return rows

    preferred_set = set(preferred_accounts)
    for row in reversed(rows):
        if str(row.get('G')) != '40' or row.get('I') not in preferred_set:
            continue
        current_amount = _to_money(row.get('M'))
        new_amount = (current_amount + adjustment).quantize(Decimal('0.01'))
        if new_amount < Decimal('0.00'):
            continue
        row['M'] = _to_text_money(new_amount)
        return rows

    raise ValueError(f'{voucher_label} 存在 {adjustment} 尾差，但未找到可补差的费用借方行')


def _amount_to_ten_thousand_yuan(amount):
    return (amount / Decimal('10000')).quantize(Decimal('0.0001'), rounding=ROUND_HALF_UP)


def _should_skip_a5_accrual(company, posting_year, posting_month, dept):
    return skips_department(company, posting_year, posting_month, dept, 'A5')


def _should_skip_a10_allocation(company, posting_year, posting_month, dept):
    return skips_department(company, posting_year, posting_month, dept, 'A10')
