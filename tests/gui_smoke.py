"""Windows/Tk smoke test against a temporary synthetic batch, with hidden UI."""
import sys
import tempfile
import time
from pathlib import Path
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'工作区域'))
from payroll_tool import gui
from test_payroll_integration import PayrollIntegrationTests


def wait(app,predicate):
    deadline=time.monotonic()+25
    while not predicate():
        app.root.update()
        if time.monotonic()>deadline:raise TimeoutError('GUI smoke timeout')
        time.sleep(.01)


def main():
    with tempfile.TemporaryDirectory(prefix='payroll-gui-') as tmp:
        base,layout,raw,options=PayrollIntegrationTests()._create_minimal_batch(Path(tmp))
        original=gui.tk.Tk
        def hidden_root():
            root=original();root.withdraw();root.deiconify=lambda:None
            return root
        with patch.object(gui,'__file__',str(base/'payroll_tool/gui.py')),patch.object(gui.tk,'Tk',side_effect=hidden_root):
            app=gui.MainWindow()
            try:
                app._suspend_option_refresh=True
                app._processing_year_var.set(2026);app._processing_month_var.set(3)
                for company,var in app._company_vars.items():var.set(company=='耐数电子')
                for voucher,var in app._voucher_vars.items():var.set(voucher in ('A5','A6'))
                app._suspend_option_refresh=False
                app._start_precheck()
                wait(app,lambda:not app._precheck_running)
                assert app._precheck_result['can_run'],app._precheck_result
                app._start_run()
                wait(app,lambda:app._result is not None)
                assert app._result['status'].startswith('success'),app._result
                assert Path(app._result['history_path']).is_dir()
                print('GUI precheck, run, result and history: PASS')
            finally:
                app._cancel_guide_refresh();app.root.destroy()


if __name__=='__main__':main()
