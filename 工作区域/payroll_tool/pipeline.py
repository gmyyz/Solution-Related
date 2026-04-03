from .core import fill_first_sheet_ab
from .options import normalize_run_options


def execute_payroll_run(input_path, base_dir, mapping_path, bank_dir, log, run_options=None):
    return fill_first_sheet_ab(
        input_path,
        base_dir,
        mapping_path,
        bank_dir,
        log,
        run_options=normalize_run_options(run_options),
    )
