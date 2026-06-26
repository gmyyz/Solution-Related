from decimal import Decimal


MAX_BANK_MATCH_COMBO_SIZE = 6
BANK_VOUCHER_TYPE = 'KZ'
BANK_REASON_CODE = '202'

BONUS_TAX_BY_PAYMENT_PERIOD = {
    ('2050', 2026, 4): Decimal('75720.04'),
    ('2060', 2026, 4): Decimal('47375.96'),
}

BONUS_NET_BY_PAYMENT = {
    ('2050', 2025, 3): Decimal('867242.06'),
    ('2060', 2025, 3): Decimal('945263.44'),
    ('2050', 2026, 3): Decimal('992732.96'),
    ('2060', 2026, 3): Decimal('349912.24'),
}

BONUS_TAX_BY_PAYMENT = {
    ('2050', 2025, 3): Decimal('83297.45'),
    ('2060', 2025, 3): Decimal('99952.56'),
    ('2050', 2026, 4): Decimal('75720.04'),
    ('2060', 2026, 4): Decimal('47375.96'),
}

BONUS_NOTES = {
    ('2050', 2025, 3): '年终奖实发按补充信息列示；流水定位：2050招行对帐单-2503.xlsx第19行',
    ('2060', 2025, 3): '年终奖实发按补充信息列示；流水定位：2060招行对帐单-2503.xlsx第17行',
    ('2050', 2026, 3): '年终奖实发按补充信息列示；银行流水两笔：937,881.44 + 54,851.52 = 992,732.96',
    ('2060', 2026, 3): '年终奖实发按补充信息列示；流水定位：银行流水-206003.xlsx第3行',
}

FUND_BANK_OVERRIDE = {
    ('2050', 2026, 2): Decimal('91162.00'),
}

SOCIAL_COMPONENT_OVERRIDE_BY_PAYMENT = {
    ('2050', 2025, 10): {
        'employee_pension': Decimal('41391.89'),
        'company_pension': Decimal('82783.77'),
        'employee_unemployment': Decimal('2587.05'),
        'company_unemployment': Decimal('2587.05'),
        'company_injury': Decimal('1034.78'),
        'employee_medical': Decimal('10446.97'),
        'company_medical': Decimal('50705.08'),
    },
    ('2050', 2025, 11): {
        'employee_pension': Decimal('40787.60'),
        'company_pension': Decimal('81575.20'),
        'employee_unemployment': Decimal('2549.31'),
        'company_unemployment': Decimal('2549.31'),
        'company_injury': Decimal('1019.68'),
        'employee_medical': Decimal('10295.90'),
        'company_medical': Decimal('49964.82'),
    },
}

TAX_BANK_AMOUNT_HINT = {
    ('2050', 2025, 3): Decimal('96394.68'),
    ('2050', 2025, 5): Decimal('19265.20'),
}
