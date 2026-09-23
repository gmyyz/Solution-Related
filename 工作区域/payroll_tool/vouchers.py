"""vouchers: extracted payroll responsibilities; business rules preserved."""
from decimal import Decimal, ROUND_HALF_UP, ROUND_DOWN
from .banking import (
    bank_scope,
    _get_treasury_records,
    _load_bank_records,
    _load_bank_records_with_stats,
    _match_bank_record_combination,
    _match_salary_excluding_bonus,
    _match_treasury_combination,
    _normalize_bank_date_key,
    _summarize_bank_records,
)
from .rules import (
    BANK_REASON_CODE,
    BANK_VOUCHER_TYPE,
    BONUS_DEPARTMENT_ACCRUAL_OVERRIDES,
    EMBEDDED_BONUS_NET_BY_PAYMENT,
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
    ACCRUAL_CREDIT_SPECS,
    ACCRUAL_DEBIT_SPECS,
    CO_CREDIT_COST_CENTER,
    CO_CREDIT_ORDER,
    SOCIAL_DEBIT_ACCOUNT_COLUMNS,
    SOCIAL_TEXT_BY_ACCOUNT,
    _allocate_amount_by_weights,
    _bonus_label_for_text,
    _get_bonus_tax_adjustment,
    _get_last_bank_date,
    _make_voucher_row,
    _month_end_date,
    _normalize_department_for_order,
    _period_text,
    _rebalance_expense_debit_rows,
    _shift_month,
    _should_skip_one_time_august_2026_a10,
    _should_skip_one_time_june_2026_accrual,
)
from .contexts import (
    _build_bonus_department_amounts,
    _lookup_internal_order_from_timesheet,
    _resolve_a4_internal_order,
)


def _build_a1_rows(company, company_code, ws, total_row_idx, bank_records, payroll_year, payroll_month):
    next_year, next_month = _shift_month(payroll_year, payroll_month, 1)
    month_records = bank_records.get((company_code, next_year, next_month), [])
    social_target = sum(
        (_to_money(ws.cell(total_row_idx, col_idx).value) for col_idx, _ in SOCIAL_DEBIT_ACCOUNT_COLUMNS),
        Decimal('0.00'),
    )
    social_match = _match_treasury_combination(month_records, social_target, '社保国库流水')
    social_records = social_match['matched_records']
    if not social_match['matched'] or not social_records:
        raise ValueError(f'{company} 未找到匹配社保金额的银行流水组合：{social_match["note"]}')

    posting_date = _get_last_bank_date(social_records)
    text = f'支付{_period_text(payroll_year, payroll_month)}工作期间社保'
    rows = []
    period_text = _period_text(payroll_year, payroll_month)

    for col_idx, gl_account in SOCIAL_DEBIT_ACCOUNT_COLUMNS:
        amount = _to_money(ws.cell(total_row_idx, col_idx).value)
        if amount == Decimal('0.00'):
            continue
        debit_text = SOCIAL_TEXT_BY_ACCOUNT[gl_account].format(period=period_text)
        rows.append(
            _make_voucher_row(
                'A1',
                company_code,
                posting_date,
                gl_account,
                amount,
                debit_text,
                40,
                voucher_type=BANK_VOUCHER_TYPE,
            )
        )

    for record in social_records:
        credit_row = _make_voucher_row(
            'A1',
            company_code,
            posting_date,
            record['bank_subject'],
            record['outgoing_amt'],
            text,
            50,
            voucher_type=BANK_VOUCHER_TYPE,
        )
        credit_row['P'] = BANK_REASON_CODE
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
            voucher_type=BANK_VOUCHER_TYPE,
        ),
        _make_voucher_row(
            'A2',
            company_code,
            posting_date,
            '2211070002',
            company_amount,
            f'支付{period_text}工作期间公司承担公积金',
            40,
            voucher_type=BANK_VOUCHER_TYPE,
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
            voucher_type=BANK_VOUCHER_TYPE,
        )
        credit_row['P'] = BANK_REASON_CODE
        rows.append(
            credit_row
        )

    return rows


def _build_a3_rows(company, company_code, ws, bank_records, payroll_year, payroll_month, timesheet_context):
    next_year, next_month = _shift_month(payroll_year, payroll_month, 1)
    month_records = bank_records.get((company_code, next_year, next_month), [])
    salary_records = [record for record in month_records if '代发工资' in record['usage']]

    labor_rows = []
    severance_rows = []
    salary_total = Decimal('0')
    tax_total = Decimal('0')
    labor_tax_total = Decimal('0')
    bank_tax_target = Decimal('0')
    total_row_idx = _find_total_row(ws)
    if total_row_idx is None:
        raise ValueError(f'{company} 工资单中未找到总计行')
    salary_match_target = _to_money(ws.cell(total_row_idx, 17).value)
    data_start_row = _find_payroll_header_row(ws) + 1
    for row_idx in range(data_start_row, ws.max_row + 1):
        if _is_total_row(ws.cell(row_idx, 1).value):
            break
        item_type = _normalize_text(ws.cell(row_idx, 4).value)
        if item_type == '劳务费':
            labor_tax_total += _to_decimal(ws.cell(row_idx, 15).value)
            bank_tax_target += _to_decimal(ws.cell(row_idx, 15).value)
            labor_rows.append(
                {
                    'amount': _to_money(ws.cell(row_idx, 5).value),
                    'cost_center': _format_code(ws.cell(row_idx, 19).value),
                    'order': _format_code(ws.cell(row_idx, 20).value),
                }
            )
        elif item_type == '离职补偿金':
            dept = _normalize_department_for_order(ws.cell(row_idx, 2).value)
            order, daily_project = _lookup_internal_order_from_timesheet(
                timesheet_context,
                company,
                dept,
                '',
            )
            if not order:
                raise ValueError(
                    f'{company} 离职补偿金第 {row_idx} 行未找到部门日常内部订单：'
                    f'部门“{dept}”，匹配项目“{daily_project}”'
                )
            severance_rows.append(
                {
                    'amount': _to_money(ws.cell(row_idx, 17).value),
                    'cost_center': _format_code(ws.cell(row_idx, 19).value),
                    'order': order,
                }
            )
        elif item_type == '薪资':
            salary_total += _to_decimal(ws.cell(row_idx, 17).value)
            tax_total += _to_decimal(ws.cell(row_idx, 15).value)
            bank_tax_target += _to_decimal(ws.cell(row_idx, 15).value)

    salary_total = salary_total.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    tax_total = tax_total.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    labor_tax_total = labor_tax_total.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    bank_tax_target = bank_tax_target.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    bonus_tax_adjustment = _get_bonus_tax_adjustment(company_code, next_year, next_month)
    bank_tax_target = (bank_tax_target + bonus_tax_adjustment).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

    embedded_bonus = EMBEDDED_BONUS_NET_BY_PAYMENT.get(
        (company_code, next_year, next_month),
        Decimal('0.00'),
    )
    salary_match = _match_salary_excluding_bonus(
        salary_records,
        salary_match_target,
        embedded_bonus,
    )
    if not salary_match['matched'] or not salary_match['matched_records']:
        raise ValueError(f'{company} A3 未找到匹配工资金额的银行流水组合')
    tax_match = _match_treasury_combination(month_records, bank_tax_target, '个税国库流水')
    tax_records = tax_match['matched_records']
    if not tax_match['matched'] or not tax_records:
        raise ValueError(f'{company} A3 未找到匹配个税金额的银行流水组合：{tax_match["note"]}')

    credit_records = salary_match['matched_records'] + tax_records
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
                voucher_type=BANK_VOUCHER_TYPE,
            )
        )

    for severance_row in severance_rows:
        rows.append(
            _make_voucher_row(
                'A3',
                company_code,
                posting_date,
                '6601990001',
                severance_row['amount'],
                '离职补偿金',
                40,
                severance_row['cost_center'],
                severance_row['order'],
                voucher_type=BANK_VOUCHER_TYPE,
            )
        )

    tax_debit_amount = (tax_match['matched_amount'] - labor_tax_total).quantize(
        Decimal('0.01'),
        rounding=ROUND_HALF_UP,
    )
    if tax_debit_amount < Decimal('0.00'):
        raise ValueError(f'{company} A3 个税银行流水小于劳务费个税，无法生成 2221070000 借方')
    if tax_debit_amount != Decimal('0.00'):
        rows.append(
            _make_voucher_row(
                'A3',
                company_code,
                posting_date,
                '2221070000',
                tax_debit_amount,
                f'支付{period_text}工作期间个人所得税',
                40,
                voucher_type=BANK_VOUCHER_TYPE,
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
                voucher_type=BANK_VOUCHER_TYPE,
            )
        )
    if embedded_bonus != Decimal('0.00'):
        rows.append(
            _make_voucher_row(
                'A3',
                company_code,
                posting_date,
                '2211010001',
                embedded_bonus,
                '支付2025年9月部分年终奖',
                40,
                voucher_type=BANK_VOUCHER_TYPE,
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
            voucher_type=BANK_VOUCHER_TYPE,
        )
        credit_row['P'] = BANK_REASON_CODE
        rows.append(
            credit_row
        )

    if embedded_bonus != Decimal('0.00'):
        gross_salary_records = salary_match.get('gross_matched_records', [])
        if not gross_salary_records:
            raise ValueError(f'{company} A3 年终奖银行贷方缺少可复制的工资流水')
        bonus_credit_row = _make_voucher_row(
            'A3',
            company_code,
            posting_date,
            gross_salary_records[0]['bank_subject'],
            embedded_bonus,
            '支付2025年9月部分年终奖',
            50,
            voucher_type=BANK_VOUCHER_TYPE,
        )
        bonus_credit_row['P'] = BANK_REASON_CODE
        rows.append(bonus_credit_row)

    for tax_record in tax_records:
        tax_credit_row = _make_voucher_row(
            'A3',
            company_code,
            posting_date,
            tax_record['bank_subject'],
            tax_record['outgoing_amt'],
            f'支付{period_text}工作期间个人所得税',
            50,
            voucher_type=BANK_VOUCHER_TYPE,
        )
        tax_credit_row['P'] = BANK_REASON_CODE
        rows.append(tax_credit_row)

    return _rebalance_expense_debit_rows(rows, ('6601110001', '6601990001'), f'{company} A3')


def _build_a4_rows(company, company_code, ws, total_row_idx, payroll_year, payroll_month, timesheet_context):
    posting_year, posting_month = _shift_month(payroll_year, payroll_month, 1)
    posting_date = _month_end_date(posting_year, posting_month)
    period_text = _period_text(payroll_year, payroll_month)

    grouped_amounts = {}
    debit_totals = {}
    salary_total = Decimal('0')
    tax_total = Decimal('0')
    other_total = Decimal('0')

    data_start_row = _find_payroll_header_row(ws) + 1
    for row_idx in range(data_start_row, ws.max_row + 1):
        if _is_total_row(ws.cell(row_idx, 1).value):
            break

        item_type = _normalize_text(ws.cell(row_idx, 4).value)
        if item_type in ('劳务费', '离职补偿金'):
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

    expense_accounts = tuple(account for _, account, _ in ACCRUAL_DEBIT_SPECS)
    return _rebalance_expense_debit_rows(rows, expense_accounts, f'{company} A4')


def _collect_a5_base_data(ws, company, posting_year, posting_month):
    grouped_amounts = {}
    debit_totals = {}
    dept_by_cost_center = {}
    salary_total = Decimal('0')
    tax_total = Decimal('0')
    other_total = Decimal('0')

    data_start_row = _find_payroll_header_row(ws) + 1
    for row_idx in range(data_start_row, ws.max_row + 1):
        if _is_total_row(ws.cell(row_idx, 1).value):
            break

        item_type = _normalize_text(ws.cell(row_idx, 4).value)
        if item_type in ('劳务费', '离职补偿金'):
            continue

        dept = _normalize_department_for_order(ws.cell(row_idx, 2).value)
        if _should_skip_one_time_june_2026_accrual(company, posting_year, posting_month, dept):
            continue
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
    posting_year, posting_month = _shift_month(payroll_year, payroll_month, 1)
    accrual_data = _collect_a5_base_data(ws, company, posting_year, posting_month)
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

    expense_accounts = tuple(account for _, account, _ in ACCRUAL_DEBIT_SPECS)
    return _rebalance_expense_debit_rows(rows, expense_accounts, f'{company} A5')


def _build_a6_rows(company, company_code, ws, payroll_year, payroll_month, timesheet_context):
    posting_year, posting_month = _shift_month(payroll_year, payroll_month, 1)
    accrual_data = _collect_a5_base_data(ws, company, posting_year, posting_month)
    posting_date = _month_end_date(posting_year, posting_month)
    text = f'根据研发工时分摊计提的{posting_year}年{posting_month}月人工费用'

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
    text = f'根据研发工时分摊计提的{label}年终奖'

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
        if _should_skip_one_time_august_2026_a10(company, posting_year, posting_month, dept):
            continue
        if _should_skip_one_time_june_2026_accrual(company, posting_year, posting_month, dept):
            continue

        key = (
            item['gl_account'],
            dept,
            item['cost_center'],
            item['allocation_cost_center'],
            item['credit_order'],
        )
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
