"""Conservative transaction identity shared by payroll and historical reports."""
from pathlib import Path


class BankIdentityError(ValueError):
    pass


CORE_FIELDS = ('trans_date', 'currency', 'incoming_amt', 'outgoing_amt', 'counterparty_account')
FEATURE_FIELDS = ('company_code', 'bank_account') + CORE_FIELDS + ('payment_name', 'usage')


def _value(value):
    if hasattr(value, 'date'):
        return value.date().isoformat()
    if hasattr(value, 'isoformat'):
        return value.isoformat()
    return str(value).strip() if value is not None else ''


def _compatible(a, b, fields):
    # Missing information is not evidence that two payments are different.
    return all(not _value(a.get(f)) or not _value(b.get(f)) or
               _value(a.get(f)) == _value(b.get(f)) for f in fields)


class TransactionIndex:
    def __init__(self):
        self.records = []
        self.warnings = []
        self.deduped = 0

    def add(self, record):
        source = {'file': str(Path(record['file']).resolve()),
                  'sheet': record.get('sheet', ''), 'row': record['row_idx']}
        record['sources'] = [source]
        key = self.key(record)
        for prior in self.records:
            prior_key = self.key(prior)
            refs = f'{prior["sources"]} / {source}'
            if key and key == prior_key:
                conflicts = [f for f in CORE_FIELDS if _value(prior.get(f)) != _value(record.get(f))]
                if conflicts:
                    raise BankIdentityError(f'相同交易流水号字段冲突 {conflicts}：{refs}')
                prior['sources'].append(source)
                self.deduped += 1
                return False
            if key and prior_key:
                continue
            if not _compatible(prior, record, FEATURE_FIELDS):
                continue
            if prior['sources'][0]['file'] != source['file']:
                raise BankIdentityError(f'跨文件重复流水缺少可靠交易标识，请核对源资料：{refs}')
            self.warnings.append(f'单文件同特征流水按原次数保留：{refs}')
        self.records.append(record)
        return True

    @staticmethod
    def key(record):
        account = _value(record.get('bank_account'))
        transaction_id = _value(record.get('transaction_id'))
        if account and transaction_id:
            return (_value(record.get('company_code')), account, transaction_id)
        return None
