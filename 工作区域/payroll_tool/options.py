import os
from dataclasses import dataclass
from datetime import date

COMPANY_OPTIONS = (
    ('耐数电子', '2050 耐数电子'),
    ('耐数信息', '2060 耐数信息'),
)
COMPANY_NAME_TO_CODE = {company: label.split()[0] for company, label in COMPANY_OPTIONS}

VOUCHER_DISPLAY_ORDER = ('A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9', 'A10')

VOUCHER_LABELS = {
    'A1': 'A1 社保与公积金发放',
    'A2': 'A2 公积金发放',
    'A3': 'A3 工资与个税发放',
    'A4': 'A4 实际薪酬分摊',
    'A5': 'A5 薪酬计提',
    'A6': 'A6 计提薪酬分摊',
    'A7': 'A7 年终奖计提',
    'A8': 'A8 年终奖分摊',
    'A9': 'A9 CO工单分摊',
    'A10': 'A10 研发费用分摊',
}

VOUCHER_PRESETS = {
    'actual': ('A1', 'A2', 'A3', 'A4'),
    'accrual': ('A5', 'A6'),
    'bonus': ('A7', 'A8'),
    'co': ('A9',),
    'rd': ('A10',),
    'all': VOUCHER_DISPLAY_ORDER,
}

ACTUAL_VOUCHERS = ('A1', 'A2', 'A3', 'A4')
ACCRUAL_VOUCHERS = ('A5', 'A6')
BONUS_VOUCHERS = ('A7', 'A8')
CO_VOUCHERS = ('A9',)
RD_VOUCHERS = ('A10',)

VOUCHER_FILE_GROUPS = {
    'actual': ACTUAL_VOUCHERS,
    'accrual': ACCRUAL_VOUCHERS,
    'bonus': BONUS_VOUCHERS,
    'co': CO_VOUCHERS,
    'rd': RD_VOUCHERS,
}


@dataclass(frozen=True)
class RunOptions:
    companies: tuple[str, ...]
    vouchers: tuple[str, ...]
    posting_year: int | None = None
    posting_month: int | None = None

    def wants_company(self, company: str) -> bool:
        return company in self.companies

    def wants_voucher(self, voucher: str) -> bool:
        return voucher in self.vouchers

    def wants_any(self, *voucher_ids: str) -> bool:
        return any(voucher in self.vouchers for voucher in voucher_ids)

    @property
    def processing_year(self) -> int:
        year, _ = resolve_processing_period(self.posting_year, self.posting_month)
        return year

    @property
    def processing_month(self) -> int:
        _, month = resolve_processing_period(self.posting_year, self.posting_month)
        return month

    @property
    def payroll_year(self) -> int:
        year, _ = shift_month(self.processing_year, self.processing_month, -1)
        return year

    @property
    def payroll_month(self) -> int:
        _, month = shift_month(self.processing_year, self.processing_month, -1)
        return month

    @property
    def processing_period(self) -> tuple[int, int]:
        return self.processing_year, self.processing_month

    @property
    def payroll_period(self) -> tuple[int, int]:
        return self.payroll_year, self.payroll_month

    @property
    def processing_label(self) -> str:
        return format_period_label(self.processing_year, self.processing_month)

    @property
    def payroll_label(self) -> str:
        return format_period_label(self.payroll_year, self.payroll_month)

    @property
    def processing_yymm(self) -> str:
        return format_period_yymm(self.processing_year, self.processing_month)

    @property
    def batch_name(self) -> str:
        return build_batch_name(self.processing_year, self.processing_month)


@dataclass(frozen=True)
class BatchLayout:
    base_dir: str
    processing_year: int
    processing_month: int
    payroll_year: int
    payroll_month: int
    batch_name: str
    batch_code: str
    master_data_root: str
    monthly_input_root: str
    run_output_root: str
    archive_root: str
    governance_root: str
    mapping_path: str
    raw_dir: str
    bank_dir: str
    timesheet_path: str
    bonus_path: str
    shared_expense_path: str
    co_workorder_path: str

    def company_output_root(self, company: str) -> str:
        company_code = COMPANY_NAME_TO_CODE[company]
        return os.path.join(self.run_output_root, f'{company_code}_{company}')

    def company_input_root(self, company: str) -> str:
        company_code = COMPANY_NAME_TO_CODE[company]
        return os.path.join(self.monthly_input_root, f'{company_code}_{company}')

    @property
    def input_manifest_dir(self) -> str:
        return os.path.join(self.archive_root, '01_输入清单')

    @property
    def precheck_dir(self) -> str:
        return os.path.join(self.archive_root, '02_预检报告')

    @property
    def run_log_dir(self) -> str:
        return os.path.join(self.archive_root, '03_运行日志')

    @property
    def process_note_dir(self) -> str:
        return os.path.join(self.archive_root, '04_过程说明')

    @property
    def output_manifest_dir(self) -> str:
        return os.path.join(self.archive_root, '05_输出清单')

    @property
    def batch_note_dir(self) -> str:
        return os.path.join(self.monthly_input_root, '07_批次说明')


def shift_month(year: int, month: int, offset: int) -> tuple[int, int]:
    total = year * 12 + (month - 1) + offset
    return total // 12, total % 12 + 1


def get_default_processing_period(today: date | None = None) -> tuple[int, int]:
    today = today or date.today()
    if today.day <= 5:
        return shift_month(today.year, today.month, -1)
    return today.year, today.month


def resolve_processing_period(year: int | None, month: int | None) -> tuple[int, int]:
    default_year, default_month = get_default_processing_period()
    year = year or default_year
    month = month or default_month
    if not 1 <= int(month) <= 12:
        return default_year, default_month
    return int(year), int(month)


def format_period_label(year: int, month: int) -> str:
    return f'{year}年{month}月'


def format_period_yymm(year: int, month: int) -> str:
    return f'{year % 100:02d}{month:02d}'


def build_batch_name(year: int, month: int) -> str:
    return f'{year}-{month:02d}处理批次({format_period_yymm(year, month)})'


def _prefer_existing_path(primary: str, fallback: str) -> str:
    if os.path.exists(primary):
        return primary
    return fallback


def _prefer_existing_dir(primary: str, fallback: str) -> str:
    if os.path.isdir(primary):
        return primary
    return fallback


def build_batch_layout(base_dir: str, processing_year: int, processing_month: int) -> BatchLayout:
    payroll_year, payroll_month = shift_month(processing_year, processing_month, -1)
    batch_code = format_period_yymm(processing_year, processing_month)
    batch_name = build_batch_name(processing_year, processing_month)

    legacy_raw_root = os.path.join(base_dir, '原始数据')
    legacy_mapping_path = os.path.join(base_dir, 'Mapping表.xlsx')
    legacy_raw_dir = os.path.join(legacy_raw_root, '工资单')
    legacy_bank_dir = os.path.join(legacy_raw_root, '银行流水')
    legacy_timesheet_path = os.path.join(legacy_raw_root, '工时数据', '工时数据.xlsx')
    legacy_bonus_path = os.path.join(legacy_raw_root, '奖金数据', '年终奖计提2026_ - to财务.xlsx')
    legacy_shared_expense_path = os.path.join(legacy_raw_root, '待分摊费用', f'待分摊费用{batch_code}.xlsx')
    legacy_co_path = os.path.join(base_dir, '耐数电子', batch_code, 'CO工单分摊', 'CO工单分摊.xlsx')

    master_data_root = os.path.join(base_dir, '01_基础资料')
    monthly_input_root = os.path.join(base_dir, '02_月度输入', batch_name)
    run_output_root = os.path.join(base_dir, '03_运行输出', batch_name)
    archive_root = os.path.join(base_dir, '04_归档留痕', batch_name)
    governance_root = os.path.join(base_dir, '05_流程规范')

    mapping_path = _prefer_existing_path(
        os.path.join(master_data_root, '01_映射与规则', 'Mapping表.xlsx'),
        legacy_mapping_path,
    )
    master_timesheet_path = os.path.join(master_data_root, '02_工时数据', '工时数据.xlsx')
    monthly_timesheet_path = os.path.join(monthly_input_root, '03_工时数据', '工时数据.xlsx')
    raw_dir = _prefer_existing_dir(
        os.path.join(monthly_input_root, '01_工资单'),
        legacy_raw_dir,
    )
    bank_dir = _prefer_existing_dir(
        os.path.join(monthly_input_root, '02_银行流水'),
        legacy_bank_dir,
    )
    timesheet_path = _prefer_existing_path(
        master_timesheet_path,
        _prefer_existing_path(monthly_timesheet_path, legacy_timesheet_path),
    )
    bonus_path = _prefer_existing_path(
        os.path.join(monthly_input_root, '04_奖金数据', '年终奖计提2026_ - to财务.xlsx'),
        legacy_bonus_path,
    )
    co_workorder_path = _prefer_existing_path(
        os.path.join(monthly_input_root, '05_CO工单分摊', 'CO工单分摊.xlsx'),
        legacy_co_path,
    )
    shared_expense_path = _prefer_existing_path(
        os.path.join(monthly_input_root, '06_待分摊费用', f'待分摊费用{batch_code}.xlsx'),
        legacy_shared_expense_path,
    )

    return BatchLayout(
        base_dir=base_dir,
        processing_year=processing_year,
        processing_month=processing_month,
        payroll_year=payroll_year,
        payroll_month=payroll_month,
        batch_name=batch_name,
        batch_code=batch_code,
        master_data_root=master_data_root,
        monthly_input_root=monthly_input_root,
        run_output_root=run_output_root,
        archive_root=archive_root,
        governance_root=governance_root,
        mapping_path=mapping_path,
        raw_dir=raw_dir,
        bank_dir=bank_dir,
        timesheet_path=timesheet_path,
        bonus_path=bonus_path,
        shared_expense_path=shared_expense_path,
        co_workorder_path=co_workorder_path,
    )


def build_batch_layout_from_options(base_dir: str, run_options: RunOptions | None) -> BatchLayout:
    normalized = normalize_run_options(run_options)
    return build_batch_layout(base_dir, normalized.processing_year, normalized.processing_month)


def normalize_run_options(run_options: RunOptions | None) -> RunOptions:
    default_year, default_month = get_default_processing_period()
    if run_options is None:
        return RunOptions(
            companies=tuple(company for company, _ in COMPANY_OPTIONS),
            vouchers=VOUCHER_DISPLAY_ORDER,
            posting_year=default_year,
            posting_month=default_month,
        )
    companies = tuple(company for company, _ in COMPANY_OPTIONS if company in run_options.companies)
    vouchers = tuple(voucher for voucher in VOUCHER_DISPLAY_ORDER if voucher in run_options.vouchers)
    posting_year, posting_month = resolve_processing_period(run_options.posting_year, run_options.posting_month)
    return RunOptions(companies=companies, vouchers=vouchers, posting_year=posting_year, posting_month=posting_month)


def requires_bank_data(run_options: RunOptions) -> bool:
    return run_options.wants_any('A1', 'A2', 'A3')


def requires_bonus_data(run_options: RunOptions) -> bool:
    return run_options.wants_any(*BONUS_VOUCHERS)


def requires_co_data(run_options: RunOptions) -> bool:
    return run_options.wants_any(*CO_VOUCHERS)


def requires_shared_expense_data(run_options: RunOptions) -> bool:
    return run_options.wants_any(*RD_VOUCHERS)
