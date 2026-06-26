import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / '工作区域'))

from payroll_tool.options import (  # noqa: E402
    RunOptions,
    build_batch_name,
    format_period_label,
    format_period_yymm,
    resolve_processing_period,
    shift_month,
)


class OptionsTests(unittest.TestCase):
    def test_shift_month_crosses_year_boundary(self):
        self.assertEqual(shift_month(2026, 1, -1), (2025, 12))
        self.assertEqual(shift_month(2026, 12, 1), (2027, 1))

    def test_run_options_derives_payroll_period(self):
        options = RunOptions(companies=('耐数电子',), vouchers=('A1',), posting_year=2026, posting_month=3)
        self.assertEqual(options.processing_period, (2026, 3))
        self.assertEqual(options.payroll_period, (2026, 2))
        self.assertEqual(options.processing_yymm, '2603')
        self.assertEqual(options.batch_name, '2026-03处理批次(2603)')

    def test_period_formatters(self):
        self.assertEqual(resolve_processing_period(2026, 3), (2026, 3))
        self.assertEqual(format_period_label(2026, 3), '2026年3月')
        self.assertEqual(format_period_yymm(2026, 3), '2603')
        self.assertEqual(build_batch_name(2026, 3), '2026-03处理批次(2603)')


if __name__ == '__main__':
    unittest.main()
