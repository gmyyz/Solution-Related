from decimal import Decimal
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / '工作区域'))

from payroll_tool.banking import _match_bank_record_combination  # noqa: E402
from payroll_tool.core import _allocate_amount_by_weights  # noqa: E402
from payroll_tool.rules import MAX_BANK_MATCH_COMBO_SIZE  # noqa: E402
from payroll_tool.workbook_utils import _to_decimal, _to_money  # noqa: E402


class UtilityTests(unittest.TestCase):
    def test_money_helpers(self):
        self.assertEqual(_to_decimal(None), Decimal('0'))
        self.assertEqual(_to_decimal('12.345'), Decimal('12.345'))
        self.assertEqual(_to_money('12.345'), Decimal('12.35'))

    def test_allocate_amount_by_weights_rebalances_cents(self):
        result = _allocate_amount_by_weights(Decimal('100.00'), [('a', 1), ('b', 1), ('c', 1)])
        self.assertEqual(sum(result.values()), Decimal('100.00'))
        self.assertEqual(sorted(result.values()), [Decimal('33.33'), Decimal('33.33'), Decimal('33.34')])

    def test_bank_combination_match(self):
        records = [
            {'outgoing_amt': Decimal('10.00')},
            {'outgoing_amt': Decimal('20.00')},
            {'outgoing_amt': Decimal('30.00')},
        ]
        result = _match_bank_record_combination(records, Decimal('50.00'), '测试流水')
        self.assertTrue(result['matched'])
        self.assertEqual(result['matched_amount'], Decimal('50.00'))
        self.assertEqual(len(result['matched_records']), 2)

    def test_bank_combination_respects_configured_limit(self):
        self.assertEqual(MAX_BANK_MATCH_COMBO_SIZE, 6)


if __name__ == '__main__':
    unittest.main()
