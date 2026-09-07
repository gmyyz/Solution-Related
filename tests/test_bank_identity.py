import sys
import tempfile
import unittest
from pathlib import Path
from decimal import Decimal
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / '工作区域'))
from payroll_tool.bank_identity import TransactionIndex, BankIdentityError
from payroll_tool.banking import _load_bank_records_with_stats
from openpyxl import Workbook


def record(file='one.xlsx', account='a', tid='', amount='100', row=2):
    return dict(file=file, row_idx=row, company_code='2050', bank_account=account,
                transaction_id=tid, trans_date='2026-03-01', currency='CNY',
                incoming_amt=Decimal('0'), outgoing_amt=Decimal(amount),
                counterparty_account='b', payment_name='name', usage='salary')


class BankIdentityTests(unittest.TestCase):
    def test_different_accounts_and_ids_are_kept(self):
        idx = TransactionIndex()
        for r in [record(account='a', tid='1'), record(account='b', tid='1'), record(tid='2')]:
            self.assertTrue(idx.add(r))
        self.assertEqual(len(idx.records), 3)

    def test_identified_duplicate_merges_sources(self):
        idx = TransactionIndex()
        idx.add(record(tid='1'))
        self.assertFalse(idx.add(record(file='two.xlsx', tid='1')))
        self.assertEqual(len(idx.records[0]['sources']), 2)

    def test_conflicting_identity_blocks(self):
        idx = TransactionIndex(); idx.add(record(tid='1'))
        with self.assertRaises(BankIdentityError):
            idx.add(record(tid='1', amount='101'))

    def test_cross_file_unknown_blocks_but_same_file_is_preserved(self):
        idx = TransactionIndex(); idx.add(record())
        self.assertTrue(idx.add(record(row=3)))
        self.assertTrue(idx.warnings)
        with self.assertRaises(BankIdentityError):
            idx.add(record(file='two.xlsx'))

    def test_missing_account_is_not_evidence_of_difference(self):
        idx = TransactionIndex(); idx.add(record())
        with self.assertRaises(BankIdentityError):
            idx.add(record(file='two.xlsx', account=''))

    def test_real_reader_closes_on_success_and_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            for name in ['银行流水1.xlsx', '银行流水2.xlsx']:
                wb=Workbook(); ws=wb.active
                ws.append(['h'] * 40)
                row=[None]*40
                for col, value in {2:'2050',3:'a',4:'2026-03-01',15:100,16:'salary'}.items():
                    row[col-1]=value
                ws.append(row); wb.save(Path(tmp)/name); wb.close()
            with self.assertRaises(BankIdentityError):
                _load_bank_records_with_stats(tmp)
            second=Path(tmp)/'银行流水2.xlsx'; second.unlink()
            self.assertEqual(_load_bank_records_with_stats(tmp)['stats']['kept_row_count'], 1)
            (Path(tmp)/'银行流水1.xlsx').unlink()


if __name__ == '__main__': unittest.main()
