import glob
import os
from datetime import date, datetime
from decimal import Decimal
from itertools import combinations

from .workbook_io import load_workbook, workbook_scope

from .rules import MAX_BANK_MATCH_COMBO_SIZE
from .bank_identity import TransactionIndex
from .workbook_utils import (
    _extract_year_month,
    _format_code,
    _normalize_text,
    _to_money,
)


BANK_FILE_PATTERN = '银行流水*.xlsx'


def _find_bank_files(bank_dir):
    return sorted(glob.glob(os.path.join(bank_dir, BANK_FILE_PATTERN)))


def _normalize_bank_date_key(value):
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return _normalize_text(value)


@workbook_scope
def _load_bank_records_with_stats(bank_dir, scope=None):
    bank_records = {}
    index = TransactionIndex()
    files = _find_bank_files(bank_dir)
    scanned = 0
    for path in files:
        wb = load_workbook(path, data_only=True, read_only=True)
        try:
            ws = wb.worksheets[0]
            iterator = ws.iter_rows(values_only=True)
            headers = next(iterator, ())
            id_col = next((i for i, v in enumerate(headers) if _normalize_text(v) == '交易流水号'), None)
            for row_idx, values in enumerate(iterator, 2):
                def at(col):
                    return values[col - 1] if col <= len(values) else None
                code = _format_code(at(2))
                period = _extract_year_month(at(4))
                if not code or period is None:
                    continue
                key = (code, *period)
                if scope is not None and key not in scope:
                    continue
                record = {
                    'file': os.path.abspath(path), 'sheet': ws.title, 'row_idx': row_idx,
                    'company_code': code, 'bank_account': _format_code(at(3)),
                    'trans_date': at(4), 'currency': _normalize_text(at(5)),
                    'payment_name': _normalize_text(at(13)),
                    'incoming_amt': _to_money(at(14)), 'outgoing_amt': _to_money(at(15)),
                    'usage': _normalize_text(at(16)), 'counterparty_account': _format_code(at(20)),
                    'transaction_id': _format_code(at(id_col + 1)) if id_col is not None else '',
                    'bank_subject': _format_code(at(34)),
                }
                scanned += 1
                if index.add(record):
                    bank_records.setdefault(key, []).append(record)
        finally:
            wb.close()
    return {'records': bank_records, 'warnings': index.warnings, 'stats': {
        'source_file_count': len(files), 'scanned_row_count': scanned,
        'deduped_row_count': index.deduped, 'kept_row_count': len(index.records),
        'warnings': index.warnings,
    }}


def bank_scope(run_options):
    from .options import COMPANY_NAME_TO_CODE
    periods = set()
    if run_options.wants_any('A1', 'A3'):
        periods.add(run_options.processing_period)
    if run_options.wants_voucher('A2'):
        periods.add(run_options.payroll_period)
        # Existing A2 generation uses the processing month's fund payments,
        # while the HR reconciliation also looks at the payroll month.
        periods.add(run_options.processing_period)
    return {(COMPANY_NAME_TO_CODE[c], *period) for c in run_options.companies for period in periods}


def _load_bank_records(bank_dir):
    return _load_bank_records_with_stats(bank_dir)['records']


def _get_treasury_records(records):
    return sorted(
        [
            record
            for record in records
            if '国家金库北京市分库' in record['payment_name'] and record['outgoing_amt'] > 0
        ],
        key=lambda item: (_normalize_bank_date_key(item['trans_date']), item['row_idx'], item['outgoing_amt']),
    )


def _summarize_bank_records(bank_records, company_code, year, month):
    records = bank_records.get((company_code, year, month), [])
    salary_records = [record for record in records if '代发工资' in record['usage']]
    salary_amount = sum((record['outgoing_amt'] for record in salary_records), Decimal('0.00'))
    fund_amount = sum(
        (
            record['outgoing_amt']
            for record in records
            if '北京住房公积金管理中心' in record['payment_name']
        ),
        Decimal('0.00'),
    )
    treasury_records = _get_treasury_records(records)
    treasury_amount = sum((record['outgoing_amt'] for record in treasury_records), Decimal('0.00'))

    return {
        'salary': salary_amount,
        'salary_records': salary_records,
        'fund': fund_amount,
        'tax': Decimal('0.00'),
        'social': Decimal('0.00'),
        'treasury': treasury_amount,
        'treasury_records': treasury_records,
        'files': sorted({record['file'] for record in records}),
        'accounts': sorted({record['bank_account'] for record in records if record['bank_account']}),
        'record_count': len(records),
    }


def _match_bank_record_combination(records, target_amount, label):
    target = _to_money(target_amount)
    if not records:
        return {
            'matched': False,
            'matched_amount': Decimal('0.00'),
            'all_amount': Decimal('0.00'),
            'matched_records': [],
            'note': f'未找到{label}',
        }

    all_amount = sum((record['outgoing_amt'] for record in records), Decimal('0.00'))
    max_size = min(MAX_BANK_MATCH_COMBO_SIZE, len(records))
    for size in range(1, max_size + 1):
        for combo in combinations(records, size):
            combo_amount = sum((record['outgoing_amt'] for record in combo), Decimal('0.00'))
            if combo_amount == target:
                combo_desc = ' + '.join(str(record['outgoing_amt']) for record in combo)
                return {
                    'matched': True,
                    'matched_amount': combo_amount,
                    'all_amount': all_amount,
                    'matched_records': list(combo),
                    'note': f'匹配到 {size} 条{label}组合：{combo_desc}',
                }

    return {
        'matched': False,
        'matched_amount': all_amount,
        'all_amount': all_amount,
        'matched_records': [],
        'note': f'未找到匹配组合；全部{label}合计={all_amount}',
    }


def _match_salary_combination(salary_records, target_amount):
    return _match_bank_record_combination(salary_records, target_amount, '代发工资')


def _match_treasury_combination(records, target_amount, label):
    return _match_bank_record_combination(_get_treasury_records(records), target_amount, label)
