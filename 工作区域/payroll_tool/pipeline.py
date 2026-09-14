from .core import fill_first_sheet_ab
from .options import normalize_run_options
from .publication import run_transaction


def execute_payroll_run(input_path, base_dir, mapping_path, bank_dir, log, run_options=None):
    return run_transaction(input_path, base_dir, mapping_path, bank_dir, log,
                           normalize_run_options(run_options), fill_first_sheet_ab)
