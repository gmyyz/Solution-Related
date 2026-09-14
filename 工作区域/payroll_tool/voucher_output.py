"""voucher_output: extracted payroll responsibilities; business rules preserved."""
import io
import os
from .workbook_io import load_workbook, workbook_scope
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
    COMPANY_NAME_TO_CODE,
)
from .vouchers import (
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
    _write_voucher_rows_to_template,
)
from .paths import (
    _build_accrual_voucher_output_path,
    _build_bonus_voucher_output_path,
    _build_co_voucher_output_path,
    _build_rd_allocation_voucher_output_path,
    _build_voucher_output_path,
    _get_co_workorder_path,
    _get_mapping_path,
    _get_shared_expense_path,
    _get_voucher_template_path,
)
from .contexts import (
    _load_co_workorder_context,
    _load_mappings,
    _load_shared_expense_context,
)
from .validation import (
    verify_written_vouchers,
    _validate_cross_group_accounts,
    _validate_voucher_groups,
)


@workbook_scope
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
    with open(template_path, 'rb') as template_file:
        template_bytes = template_file.read()
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
                _build_a3_rows(
                    company,
                    company_code,
                    company_ws,
                    bank_records,
                    payroll_year,
                    payroll_month,
                    timesheet_context,
                )
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
        if run_options.wants_any('A1', 'A2', 'A3', 'A4'):
            actual_voucher_wb = load_workbook(io.BytesIO(template_bytes))
            actual_voucher_ws = actual_voucher_wb[actual_voucher_wb.sheetnames[0]]
            _clear_voucher_template_rows(actual_voucher_ws)
            _write_voucher_rows_to_template(actual_voucher_ws, actual_voucher_rows)
            actual_voucher_output_path = _build_voucher_output_path(base_dir, company, payroll_year, payroll_month)
            actual_voucher_wb.save(actual_voucher_output_path)
            verify_written_vouchers(actual_voucher_output_path, actual_voucher_rows)
            voucher_paths[f'{company}（实际）'] = actual_voucher_output_path
            log(f'{company} 实际凭证文件已生成：{actual_voucher_output_path}', 'ok')

        accrual_voucher_output_path = ''
        if run_options.wants_any('A5', 'A6'):
            accrual_voucher_wb = load_workbook(io.BytesIO(template_bytes))
            accrual_voucher_ws = accrual_voucher_wb[accrual_voucher_wb.sheetnames[0]]
            _clear_voucher_template_rows(accrual_voucher_ws)
            _write_voucher_rows_to_template(accrual_voucher_ws, accrual_voucher_rows)
            accrual_voucher_output_path = _build_accrual_voucher_output_path(base_dir, company, payroll_year, payroll_month)
            accrual_voucher_wb.save(accrual_voucher_output_path)
            verify_written_vouchers(accrual_voucher_output_path, accrual_voucher_rows)
            voucher_paths[f'{company}（计提）'] = accrual_voucher_output_path
            log(f'{company} 计提凭证文件已生成：{accrual_voucher_output_path}', 'ok')

        bonus_voucher_output_path = ''
        if run_options.wants_any('A7', 'A8'):
            bonus_voucher_wb = load_workbook(io.BytesIO(template_bytes))
            bonus_voucher_ws = bonus_voucher_wb[bonus_voucher_wb.sheetnames[0]]
            _clear_voucher_template_rows(bonus_voucher_ws)
            _write_voucher_rows_to_template(bonus_voucher_ws, bonus_voucher_rows)
            bonus_voucher_output_path = _build_bonus_voucher_output_path(base_dir, company, payroll_year, payroll_month)
            bonus_voucher_wb.save(bonus_voucher_output_path)
            verify_written_vouchers(bonus_voucher_output_path, bonus_voucher_rows)
            voucher_paths[f'{company}（年终奖）'] = bonus_voucher_output_path
            log(f'{company} 年终奖凭证文件已生成：{bonus_voucher_output_path}', 'ok')

        if company == '耐数电子' and run_options.wants_voucher('A9'):
            co_voucher_wb = load_workbook(io.BytesIO(template_bytes))
            co_voucher_ws = co_voucher_wb[co_voucher_wb.sheetnames[0]]
            _clear_voucher_template_rows(co_voucher_ws)
            _write_voucher_rows_to_template(co_voucher_ws, co_voucher_rows)
            co_voucher_output_path = _build_co_voucher_output_path(base_dir, payroll_year, payroll_month)
            co_voucher_wb.save(co_voucher_output_path)
            verify_written_vouchers(co_voucher_output_path, co_voucher_rows)
            voucher_paths[f'{company}（CO工单分摊）'] = co_voucher_output_path
            log(f'{company} CO工单分摊凭证文件已生成：{co_voucher_output_path}', 'ok')

        if run_options.wants_voucher('A10'):
            rd_voucher_wb = load_workbook(io.BytesIO(template_bytes))
            rd_voucher_ws = rd_voucher_wb[rd_voucher_wb.sheetnames[0]]
            _clear_voucher_template_rows(rd_voucher_ws)
            _write_voucher_rows_to_template(rd_voucher_ws, rd_voucher_rows)
            rd_voucher_output_path = _build_rd_allocation_voucher_output_path(base_dir, company, payroll_year, payroll_month)
            rd_voucher_wb.save(rd_voucher_output_path)
            verify_written_vouchers(rd_voucher_output_path, rd_voucher_rows)
            voucher_paths[f'{company}（研发费用分摊）'] = rd_voucher_output_path
            log(f'{company} 研发费用分摊凭证文件已生成：{rd_voucher_output_path}', 'ok')

        generated = actual_voucher_rows + accrual_voucher_rows + bonus_voucher_rows + co_voucher_rows + rd_voucher_rows
        statuses = {
            voucher: ('not_selected' if not run_options.wants_voucher(voucher) else
                      'not_applicable' if voucher == 'A9' and company != '耐数电子' else
                      'passed' if any(row['A'] == voucher for row in generated) else 'no_data')
            for voucher in VOUCHER_DISPLAY_ORDER
        }
        validation_summary[company] = {
            **{f'{v.lower()}_balanced': True if status == 'passed' else None for v, status in statuses.items()},
            'group_balances': {v: True if status == 'passed' else None for v, status in statuses.items()},
            'group_status': statuses,
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

        company_wb.close()

    return voucher_paths, validation_summary
