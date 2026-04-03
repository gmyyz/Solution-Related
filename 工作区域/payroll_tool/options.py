from dataclasses import dataclass

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

    def wants_company(self, company: str) -> bool:
        return company in self.companies

    def wants_voucher(self, voucher: str) -> bool:
        return voucher in self.vouchers

    def wants_any(self, *voucher_ids: str) -> bool:
        return any(voucher in self.vouchers for voucher in voucher_ids)


def normalize_run_options(run_options: RunOptions | None) -> RunOptions:
    if run_options is None:
        return RunOptions(
            companies=tuple(company for company, _ in COMPANY_OPTIONS),
            vouchers=VOUCHER_DISPLAY_ORDER,
        )
    companies = tuple(company for company, _ in COMPANY_OPTIONS if company in run_options.companies)
    vouchers = tuple(voucher for voucher in VOUCHER_DISPLAY_ORDER if voucher in run_options.vouchers)
    return RunOptions(companies=companies, vouchers=vouchers)


def requires_bank_data(run_options: RunOptions) -> bool:
    return run_options.wants_any('A1', 'A2', 'A3')


def requires_bonus_data(run_options: RunOptions) -> bool:
    return run_options.wants_any(*BONUS_VOUCHERS)


def requires_co_data(run_options: RunOptions) -> bool:
    return run_options.wants_any(*CO_VOUCHERS)


def requires_shared_expense_data(run_options: RunOptions) -> bool:
    return run_options.wants_any(*RD_VOUCHERS)
