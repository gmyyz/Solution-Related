from .payroll_input import _prepare_payroll_preview
from .rules import DEPARTMENT_VOUCHER_SKIPS, EMBEDDED_BONUS_NET_BY_PAYMENT
"""core: extracted payroll responsibilities; business rules preserved."""
import os
from decimal import Decimal, ROUND_HALF_UP, ROUND_DOWN
from .workbook_io import load_workbook, workbook_scope
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
from .compensation_report import generate_compensation_report
from .publication import output_target
from .options import (
    ACTUAL_VOUCHERS,
    VOUCHER_DISPLAY_ORDER,
    build_batch_layout,
    format_period_label,
    normalize_run_options,
    requires_bank_data,
    requires_bonus_data,
    requires_co_data,
    requires_shared_expense_data,
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
    ACCOUNTS_TO_CROSS_CHECK,
    ACCRUAL_CREDIT_SPECS,
    ACCRUAL_DEBIT_SPECS,
    COMPANY_CODE_TO_NAME,
    COMPANY_NAME_TO_CODE,
    CO_CREDIT_COST_CENTER,
    CO_CREDIT_ORDER,
    CO_DEBIT_GL_BY_SOURCE,
    CO_SPECIAL_COST_ELEMENTS,
    EXAMPLE_IMAGE_DIRNAME,
    RAW_FILE_PATTERN,
    RAW_FILE_PATTERNS,
    SOCIAL_DEBIT_ACCOUNT_COLUMNS,
    SOCIAL_TEXT_BY_ACCOUNT,
    _allocate_amount_by_weights,
    _amount_to_ten_thousand_yuan,
    _bonus_company_alias,
    _bonus_label_for_filename,
    _bonus_label_for_text,
    _bonus_months_for_period,
    _build_internal_order_candidates,
    _extract_bonus_department,
    _extract_payroll_period,
    _get_bonus_tax_adjustment,
    _get_last_bank_date,
    _lookup_internal_order,
    _make_voucher_row,
    _match_order_from_kok3,
    _month_end_date,
    _normalize_department_for_order,
    _period_text,
    _rebalance_expense_debit_rows,
    _sap_date_value,
    _select_treasury_records,
    _shift_month,
    _should_skip_one_time_august_2026_a10,
    _should_skip_one_time_june_2026_accrual,
)
from .voucher_output import (
    _generate_voucher_files,
)
from .paths import (
    _build_accrual_voucher_output_path,
    _build_bonus_voucher_output_path,
    _build_co_voucher_output_path,
    _build_company_output_path,
    _build_key_tax_source_output_path,
    _build_rd_allocation_voucher_output_path,
    _build_voucher_output_path,
    _ensure_directory,
    _find_raw_files,
    _get_bank_dir,
    _get_bonus_path,
    _get_co_workorder_path,
    _get_logo_path,
    _get_mapping_path,
    _get_raw_dir,
    _get_shared_expense_path,
    _get_timesheet_path,
    _get_voucher_template_path,
)
from .contexts import (
    _build_bonus_department_amounts,
    _build_cost_center_reverse_map,
    _load_bonus_context,
    _load_co_workorder_context,
    _load_mappings,
    _load_shared_expense_context,
    _load_timesheet_match_context,
    _lookup_internal_order_from_timesheet,
    _match_timesheet_internal_orders,
    _resolve_a4_internal_order,
)
from .payroll_output import (
    _clear_bank_summary,
    _save_company_workbooks,
    _write_bank_summary,
)
from .tax_report import (
    _append_key_tax_source_row,
    _build_company_key_tax_source_rows,
    _describe_bank_records,
    _write_key_tax_source_report,
    _write_key_tax_source_workbook,
)
from .artifacts import (
    _archive_timesheet_snapshot,
    _copy_example_images_for_batch,
    _render_example_images,
    _render_markdown_list,
    _render_markdown_table,
    _render_voucher_detail_sections,
    _status_label,
    _voucher_detail_specs,
    _voucher_example_assets,
    _voucher_summary_rows,
    _write_run_artifacts,
    _write_utf8_file,
)
from .vouchers import (
    _bonus_allocation_months,
    _build_a10_rows,
    _build_a1_rows,
    _build_a2_rows,
    _build_a3_rows,
    _build_a4_rows,
    _build_a5_rows,
    _build_a6_rows,
    _build_a7_rows,
    _build_a8_rows,
    _build_a9_rows,
    _clear_voucher_template_rows,
    _collect_a5_base_data,
    _get_bonus_allocation_bucket,
    _write_voucher_rows_to_template,
)
from .validation import (
    _validate_cross_group_accounts,
    _validate_voucher_groups,
)


@workbook_scope
def fill_first_sheet_ab(input_path, base_dir, mapping_path, bank_dir, log, run_options=None):
    run_options = normalize_run_options(run_options)
    mapping_path = mapping_path or _get_mapping_path(base_dir)
    bank_dir = bank_dir or _get_bank_dir(base_dir, run_options.processing_year, run_options.processing_month)
    external_log = log
    log_entries = []

    def log(text, tag=''):
        log_entries.append({'tag': tag or 'info', 'text': text})
        external_log(text, tag)

    for (company, year, month, dept), vouchers in DEPARTMENT_VOUCHER_SKIPS.items():
        matched = set(run_options.vouchers).intersection(vouchers)
        if company in run_options.companies and (year, month) == run_options.processing_period and matched:
            log(f'命中限定月份例外：{company} {year}-{month:02d} {dept} 跳过 {sorted(matched)}', 'warn')
    log('正在读取原始工资单…')
    wb = load_workbook(input_path)
    ws = wb[wb.sheetnames[0]]
    header_row_idx = _find_payroll_header_row(ws)
    data_start_row = header_row_idx + 1
    log(f'首个工作表：{ws.title}，共 {ws.max_row} 行')
    log(f'工资单表头位于第 {header_row_idx} 行，数据从第 {data_start_row} 行开始', 'ok')
    log('正在读取 Mapping 表…')
    cost_center_map = _load_mappings(mapping_path)
    log('Mapping 表读取完成', 'ok')
    timesheet_path = _get_timesheet_path(base_dir, run_options.processing_year, run_options.processing_month)
    bonus_path = _get_bonus_path(base_dir, run_options.processing_year, run_options.processing_month)
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
    bank_scan_stats = {
        'source_file_count': 0,
        'scanned_row_count': 0,
        'deduped_row_count': 0,
        'kept_row_count': 0,
    }
    if requires_bank_data(run_options):
        log('正在读取银行流水…')
        bank_load_result = _load_bank_records_with_stats(bank_dir, scope=bank_scope(run_options))
        for warning in bank_load_result.get('warnings', []):
            log(warning, 'warn')
        bank_records = bank_load_result['records']
        bank_scan_stats = bank_load_result['stats']
        log(
            '银行流水读取完成：'
            f'扫描 {bank_scan_stats["source_file_count"]} 个文件，'
            f'保留 {bank_scan_stats["kept_row_count"]} 行，'
            f'去重 {bank_scan_stats["deduped_row_count"]} 行',
            'ok',
        )
    else:
        log('本次未选择 A1-A3，已跳过银行流水读取', 'warn')

    preview = _prepare_payroll_preview(
        input_path,
        cost_center_map,
        timesheet_context,
        set(run_options.companies),
        payment_period=run_options.processing_period,
    )
    wb, ws = preview['workbook'], preview['worksheet']
    header_ref = preview['header_ref']
    fill_a, fill_b = preview['fill_a'], preview['fill_b']
    issue_rows = preview['issue_rows']
    counts = list(preview['company_issue_counts'].values())
    validation_issue_count = sum(c['formula'] for c in counts)
    cost_center_issue_count = sum(c['cost_center'] for c in counts)
    internal_order_issue_count = sum(c['order'] for c in counts)
    payroll_summary, rows_by_company = preview['payroll_summary'], preview['rows_by_company']
    for company, bonus in preview['embedded_bonus_adjustments'].items():
        log(
            f'{company}实发已扣除银行代发中内含的年终奖 {bonus}；'
            '工资分摊按不含年终奖口径生成，A3单列年终奖借方及银行贷方',
            'warn',
        )
    if internal_order_issue_count:
        raise ValueError(f'工资单有 {internal_order_issue_count} 行内部订单未匹配，已停止处理')
    log(f'A 列填充 {fill_a} 个空白单元格；B 列填充 {fill_b} 个空白单元格', 'ok')
    if issue_rows:
        log(f'共发现 {len(issue_rows)} 行异常，已在输出文件中标红', 'warn')
    else:
        log('所有数据校验和匹配均通过', 'ok')

    payroll_year, payroll_month = run_options.payroll_period
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
            embedded_bonus = EMBEDDED_BONUS_NET_BY_PAYMENT.get(
                (company_code, next_year, next_month),
                Decimal('0.00'),
            )
            salary_match = _match_salary_excluding_bonus(
                salary_tax_bank['salary_records'],
                summary['salary'].quantize(Decimal('0.01'), rounding=ROUND_HALF_UP),
                embedded_bonus,
            )
            tax_match = _match_bank_record_combination(
                salary_tax_bank['treasury_records'],
                (
                    summary['tax']
                    + _get_bonus_tax_adjustment(company_code, next_year, next_month)
                ).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP),
                '个税国库流水',
            )
            social_match = _match_bank_record_combination(
                salary_tax_bank['treasury_records'],
                summary['social'].quantize(Decimal('0.01'), rounding=ROUND_HALF_UP),
                '社保国库流水',
            )

            checks = [
                ('salary', '实发工资', summary['salary'].quantize(Decimal('0.01'), rounding=ROUND_HALF_UP), salary_match['matched_amount'], f'{next_year}-{next_month:02d}'),
                ('fund', '公积金', summary['fund'].quantize(Decimal('0.01'), rounding=ROUND_HALF_UP), fund_bank['fund'], f'{payroll_year}-{payroll_month:02d}'),
                (
                    'tax',
                    '个税',
                    (
                        summary['tax']
                        + _get_bonus_tax_adjustment(company_code, next_year, next_month)
                    ).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP),
                    tax_match['matched_amount'],
                    f'{next_year}-{next_month:02d}',
                ),
                ('social', '社保', summary['social'].quantize(Decimal('0.01'), rounding=ROUND_HALF_UP), social_match['matched_amount'], f'{next_year}-{next_month:02d}'),
            ]

            for item_key, item_label, payroll_amount, bank_amount, period_label in checks:
                diff = (payroll_amount - bank_amount).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                passed = diff == Decimal('0.00')
                note = ''

                if item_key == 'salary':
                    passed = salary_match['matched']
                    diff = (payroll_amount - bank_amount).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                    note = salary_match['note']
                elif item_key == 'tax':
                    passed = tax_match['matched']
                    diff = (payroll_amount - bank_amount).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                    note = tax_match['note']
                    bonus_tax_adjustment = _get_bonus_tax_adjustment(company_code, next_year, next_month)
                    if bonus_tax_adjustment:
                        note = f'{note}；已包含上月年终奖个税 {bonus_tax_adjustment}'
                elif item_key == 'social':
                    passed = social_match['matched']
                    diff = (payroll_amount - bank_amount).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                    note = social_match['note']

                if item_key == 'fund' and company_code == '2050' and payroll_year == 2026 and payroll_month == 2:
                    passed = abs(diff) == Decimal('3650.00')
                    if passed:
                        note = '202602 特殊规则：公积金允许绝对差额 3650'

                source_files = salary_tax_bank['files'] if item_key in ('salary', 'tax', 'social') else fund_bank['files']
                if not source_files:
                    note = f'未找到 {period_label} 的匹配银行流水'
                elif not note:
                    note = '来源：' + '、'.join(source_files)
                elif item_key in ('salary', 'tax', 'social'):
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
    key_tax_source_paths = {}
    optional_errors = []
    try:
        tax_source_bank_records = bank_records if bank_records else _load_bank_records(bank_dir)
        key_tax_source_paths = _write_key_tax_source_report(
            base_dir,
            output_paths,
            tax_source_bank_records,
            payroll_year,
            payroll_month,
        )
        for path in key_tax_source_paths.values():
            log(f'重点税源采集信息已生成：{path}', 'ok')
    except Exception as exc:
        optional_errors.append(f'重点税源采集信息本次未更新：{exc}')
        log(optional_errors[-1], 'warn')
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
    compensation_report_result = {}
    try:
        compensation_report_result = generate_compensation_report(base_dir, output_path=output_target(os.path.join(base_dir, '03_运行输出', '实发薪酬表.xlsx')))
        report_path = compensation_report_result.get('path')
        row_counts = compensation_report_result.get('rows_by_company', {})
        log(
            '实发薪酬表已生成：'
            f'{report_path}（2050 {row_counts.get("2050", 0)} 行，2060 {row_counts.get("2060", 0)} 行）',
            'ok',
        )
    except Exception as exc:
        optional_errors.append(f'实发薪酬表本次未更新：{exc}')
        log(optional_errors[-1], 'warn')
    artifact_paths = _write_run_artifacts(
        base_dir,
        run_options,
        input_path,
        mapping_path,
        bank_dir,
        timesheet_path,
        bonus_path,
        _get_co_workorder_path(base_dir, payroll_year, payroll_month),
        _get_shared_expense_path(base_dir, payroll_year, payroll_month),
        output_paths,
        voucher_paths,
        voucher_validation_summary,
        log_entries,
        bank_scan_stats,
        compensation_report_result.get('path'),
    )
    if key_tax_source_paths:
        artifact_paths['key_tax_source'] = key_tax_source_paths
    if compensation_report_result:
        artifact_paths['compensation_report'] = compensation_report_result

    return {
        'optional_errors': optional_errors,
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
        'bank_scan_stats': bank_scan_stats,
        'processing_period_label': format_period_label(run_options.processing_year, run_options.processing_month),
        'payroll_period_label': format_period_label(payroll_year, payroll_month),
        'output_paths': output_paths,
        'voucher_paths': voucher_paths,
        'voucher_validation_summary': voucher_validation_summary,
        'artifact_paths': artifact_paths,
    }
