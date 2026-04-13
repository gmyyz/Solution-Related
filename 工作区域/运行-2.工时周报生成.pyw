import os
import traceback
import tkinter as tk
from tkinter import messagebox

from payroll_tool.timesheet_weekly_report import generate_default_timesheet_weekly_reports


def main():
    root = tk.Tk()
    root.withdraw()
    base_dir = os.path.dirname(os.path.abspath(__file__))

    try:
        output_paths = generate_default_timesheet_weekly_reports(base_dir)
    except Exception as exc:
        messagebox.showerror('工时周报生成失败', f'{exc}\n\n{traceback.format_exc()}')
        raise

    output_names = '\n'.join(os.path.basename(path) for path in output_paths)
    messagebox.showinfo('工时周报生成完成', f'已生成以下文件：\n{output_names}')


if __name__ == '__main__':
    main()
