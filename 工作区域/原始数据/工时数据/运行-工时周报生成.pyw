import os
import sys
import traceback
import tkinter as tk
from tkinter import messagebox, ttk


def _get_base_dir():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.dirname(os.path.dirname(current_dir))


BASE_DIR = _get_base_dir()
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from payroll_tool.timesheet_weekly_report import export_timesheet_weekly_report, list_available_months


class MonthPicker:
    def __init__(self, root, months):
        self.root = root
        self.result = None

        self.window = tk.Toplevel(root)
        self.window.title('选择工时月份')
        self.window.resizable(False, False)
        self.window.protocol('WM_DELETE_WINDOW', self._cancel)

        tk.Label(self.window, text='请选择要生成的工时月份：', padx=16, pady=12).pack(anchor='w')

        self.month_var = tk.StringVar(value=f'2026年{months[0]:02d}月')
        self.combo = ttk.Combobox(
            self.window,
            textvariable=self.month_var,
            state='readonly',
            values=[f'2026年{month:02d}月' for month in months],
            width=18,
        )
        self.combo.pack(padx=16, pady=(0, 12))
        self.combo.current(0)

        button_frame = tk.Frame(self.window)
        button_frame.pack(fill='x', padx=16, pady=(0, 16))
        tk.Button(button_frame, text='生成', width=10, command=self._confirm).pack(side='left')
        tk.Button(button_frame, text='取消', width=10, command=self._cancel).pack(side='right')

        self.window.transient(root)
        self.window.grab_set()
        self.window.update_idletasks()
        x = root.winfo_screenwidth() // 2 - self.window.winfo_width() // 2
        y = root.winfo_screenheight() // 2 - self.window.winfo_height() // 2
        self.window.geometry(f'+{x}+{y}')

    def _confirm(self):
        text = self.month_var.get()
        self.result = int(''.join(ch for ch in text if ch.isdigit())[-2:])
        self.window.destroy()

    def _cancel(self):
        self.result = None
        self.window.destroy()


def main():
    root = tk.Tk()
    root.withdraw()

    try:
        months = list_available_months(BASE_DIR, year=2026)
        if not months:
            raise ValueError('未找到 2026 年可生成的非零工时月份')

        picker = MonthPicker(root, months)
        root.wait_window(picker.window)
        if picker.result is None:
            return

        output_path = export_timesheet_weekly_report(BASE_DIR, 2026, picker.result)
    except Exception as exc:
        messagebox.showerror('工时周报生成失败', f'{exc}\n\n{traceback.format_exc()}')
        raise

    messagebox.showinfo('工时周报生成完成', f'已生成文件：\n{os.path.basename(output_path)}')


if __name__ == '__main__':
    main()
