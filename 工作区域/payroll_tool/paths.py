"""paths: extracted payroll responsibilities; business rules preserved."""
import glob
import os
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
from .domain import (
    COMPANY_NAME_TO_CODE,
    RAW_FILE_PATTERNS,
    _bonus_label_for_filename,
    _extract_payroll_period,
    _shift_month,
)


def _get_raw_dir(base_dir, processing_year=None, processing_month=None):
    if processing_year is not None and processing_month is not None:
        return build_batch_layout(base_dir, processing_year, processing_month).raw_dir
    return os.path.join(base_dir, '原始数据', '工资单')


def _get_mapping_path(base_dir):
    primary = os.path.join(base_dir, '01_基础资料', '01_映射与规则', 'Mapping表.xlsx')
    if os.path.exists(primary):
        return primary
    return os.path.join(base_dir, 'Mapping表.xlsx')


def _get_bank_dir(base_dir, processing_year=None, processing_month=None):
    if processing_year is not None and processing_month is not None:
        return build_batch_layout(base_dir, processing_year, processing_month).bank_dir
    return os.path.join(base_dir, '原始数据', '银行流水')


def _get_timesheet_path(base_dir, processing_year=None, processing_month=None):
    if processing_year is not None and processing_month is not None:
        return build_batch_layout(base_dir, processing_year, processing_month).timesheet_path
    master_path = os.path.join(base_dir, '01_基础资料', '02_工时数据', '工时数据.xlsx')
    if os.path.exists(master_path):
        return master_path
    return os.path.join(base_dir, '原始数据', '工时数据', '工时数据.xlsx')


def _get_bonus_path(base_dir, processing_year=None, processing_month=None):
    if processing_year is not None and processing_month is not None:
        return build_batch_layout(base_dir, processing_year, processing_month).bonus_path
    return os.path.join(base_dir, '原始数据', '奖金数据', '年终奖计提2026_ - to财务.xlsx')


def _get_shared_expense_path(base_dir, payroll_year, payroll_month):
    post_year, post_month = _shift_month(payroll_year, payroll_month, 1)
    return build_batch_layout(base_dir, post_year, post_month).shared_expense_path


def _get_co_workorder_path(base_dir, payroll_year, payroll_month):
    post_year, post_month = _shift_month(payroll_year, payroll_month, 1)
    return build_batch_layout(base_dir, post_year, post_month).co_workorder_path


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


def _find_raw_files(raw_dir, payroll_year=None, payroll_month=None):
    paths = []
    seen = set()
    for pattern in RAW_FILE_PATTERNS:
        for path in sorted(glob.glob(os.path.join(raw_dir, pattern))):
            normalized = os.path.normcase(path)
            if normalized in seen:
                continue
            seen.add(normalized)
            paths.append(path)
    if payroll_year is None or payroll_month is None:
        return paths
    matched = []
    for path in paths:
        try:
            file_year, file_month = _extract_payroll_period(path)
        except ValueError:
            continue
        if file_year == payroll_year and file_month == payroll_month:
            matched.append(path)
    return matched


def _build_company_output_path(base_dir, company, payroll_year, payroll_month, input_path):
    post_year, post_month = _shift_month(payroll_year, payroll_month, 1)
    layout = build_batch_layout(base_dir, post_year, post_month)
    company_dir = os.path.join(layout.company_output_root(company), '01_工资整理')
    os.makedirs(company_dir, exist_ok=True)

    output_name = f'{payroll_month}月工资单-整理后.xlsx'
    return os.path.join(company_dir, output_name)


def _build_voucher_output_path(base_dir, company, payroll_year, payroll_month):
    company_code = COMPANY_NAME_TO_CODE[company]
    post_year, post_month = _shift_month(payroll_year, payroll_month, 1)
    layout = build_batch_layout(base_dir, post_year, post_month)
    company_dir = os.path.join(layout.company_output_root(company), '02_实际凭证')
    os.makedirs(company_dir, exist_ok=True)
    return os.path.join(company_dir, f'{company_code}总账凭证导入-实际{post_year}{post_month:02d}薪酬.XLS')


def _build_accrual_voucher_output_path(base_dir, company, payroll_year, payroll_month):
    company_code = COMPANY_NAME_TO_CODE[company]
    post_year, post_month = _shift_month(payroll_year, payroll_month, 1)
    layout = build_batch_layout(base_dir, post_year, post_month)
    company_dir = os.path.join(layout.company_output_root(company), '03_计提凭证')
    os.makedirs(company_dir, exist_ok=True)
    return os.path.join(company_dir, f'{company_code}总账凭证导入-计提{post_year}{post_month:02d}薪酬.XLS')


def _build_bonus_voucher_output_path(base_dir, company, payroll_year, payroll_month):
    company_code = COMPANY_NAME_TO_CODE[company]
    post_year, post_month = _shift_month(payroll_year, payroll_month, 1)
    layout = build_batch_layout(base_dir, post_year, post_month)
    company_dir = os.path.join(layout.company_output_root(company), '04_年终奖凭证')
    os.makedirs(company_dir, exist_ok=True)
    return os.path.join(company_dir, f'{company_code}总账凭证导入-计提{_bonus_label_for_filename(post_year, post_month)}年终奖.XLS')


def _build_co_voucher_output_path(base_dir, payroll_year, payroll_month):
    post_year, post_month = _shift_month(payroll_year, payroll_month, 1)
    layout = build_batch_layout(base_dir, post_year, post_month)
    co_dir = os.path.join(layout.company_output_root('耐数电子'), '06_CO工单分摊')
    os.makedirs(co_dir, exist_ok=True)
    return os.path.join(co_dir, f'2050总账凭证导入-CO工单分摊{post_year}{post_month:02d}.XLS')


def _build_rd_allocation_voucher_output_path(base_dir, company, payroll_year, payroll_month):
    company_code = COMPANY_NAME_TO_CODE[company]
    post_year, post_month = _shift_month(payroll_year, payroll_month, 1)
    layout = build_batch_layout(base_dir, post_year, post_month)
    company_dir = os.path.join(layout.company_output_root(company), '05_研发费用分摊')
    os.makedirs(company_dir, exist_ok=True)
    return os.path.join(company_dir, f'{company_code}总账-研发费用分摊{post_year % 100:02d}{post_month:02d}.XLS')


def _build_key_tax_source_output_path(base_dir, company, processing_year, processing_month):
    company_code = COMPANY_NAME_TO_CODE[company]
    layout = build_batch_layout(base_dir, processing_year, processing_month)
    output_dir = os.path.join(layout.company_output_root(company), '07_重点税源采集信息')
    os.makedirs(output_dir, exist_ok=True)
    return os.path.join(output_dir, f'{company_code}重点税源采集信息{processing_year % 100:02d}{processing_month:02d}.xlsx')


def _ensure_directory(path):
    os.makedirs(path, exist_ok=True)
    return path
