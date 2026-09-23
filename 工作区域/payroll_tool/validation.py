"""validation: extracted payroll responsibilities; business rules preserved."""
from decimal import Decimal, ROUND_HALF_UP, ROUND_DOWN
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
    ACCOUNTS_TO_CROSS_CHECK,
)
from .workbook_io import load_workbook


def verify_written_vouchers(path, expected_rows):
    """Compare independently reloaded A:Q values, not just aggregate balance."""
    wb = load_workbook(path, data_only=True)
    try:
        actual = [row for row in wb.active.iter_rows(min_row=2, max_col=17, values_only=True)
                  if any(value is not None and value != '' for value in row)]
        expected = [tuple(row.get(chr(65+i)) for i in range(17)) for row in expected_rows]
        def canonical(rows):
            return [tuple('' if value is None else str(value) for value in row) for row in rows]
        if canonical(actual) != canonical(expected):
            raise ValueError(f'导出文件与生成记录不一致：{path}')
    finally:
        wb.close()


def _validate_voucher_groups(voucher_rows):
    group_totals = {}
    for row in voucher_rows:
        amount = _to_money(row['M'])
        if str(row['G']) not in ('40', '50'):
            raise ValueError(f'未知记账码：{row["G"]}')
        if str(row['G']) == '40' and amount < 0:
            raise ValueError(f'{row["A"]} 分摊或费用借方金额为负数：{amount}')
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
    prior_bonus_payment_debit = Decimal('0.00')

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
        if (
            account == '2211010001'
            and str(row['G']) == '40'
            and row['N'] == '支付2025年9月部分年终奖'
        ):
            prior_bonus_payment_debit += amount

    issues = []
    for account in ACCOUNTS_TO_CROSS_CHECK:
        totals = account_totals.get(account, {'debit': Decimal('0.00'), 'credit': Decimal('0.00')})
        debit_total = totals['debit'].quantize(Decimal('0.01'))
        credit_total = totals['credit'].quantize(Decimal('0.01'))

        if account == '2211010001':
            adjusted_debit = (debit_total - prior_bonus_payment_debit).quantize(Decimal('0.01'))
            adjusted_credit = (credit_total - other_project_credit).quantize(Decimal('0.01'))
            if adjusted_debit != adjusted_credit:
                issues.append(
                    f'{account} 跨凭证不平：借方(扣除历史年终奖支付后) {adjusted_debit}，'
                    f'贷方(扣除其他项目后) {adjusted_credit}'
                )
        elif debit_total != credit_total:
            issues.append(f'{account} 跨凭证不平：借方 {debit_total}，贷方 {credit_total}')

    return issues
