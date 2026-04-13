from dataclasses import dataclass
from datetime import date

COMPANY_OPTIONS = (
    ('耐数电子', '2050 耐数电子'),
    ('耐数信息', '2060 耐数信息'),
)

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
