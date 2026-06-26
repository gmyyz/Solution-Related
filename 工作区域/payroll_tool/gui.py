import os
import queue
import threading
import tkinter as tk
from tkinter import messagebox, ttk

from .core import (
    BANK_FILE_PATTERN,
    COLOR_BG,
    COLOR_BORDER,
    COLOR_CARD,
    COLOR_DANGER,
    COLOR_HEADER_BG,
    COLOR_PRIMARY,
    COLOR_SUCCESS,
    COLOR_TEXT_MAIN,
    COLOR_TEXT_SUB,
    COLOR_WARNING,
    RAW_FILE_PATTERN,
    _apply_style,
    _btn,
    _circle_label,
    _find_bank_files,
    _find_raw_files,
    _get_bank_dir,
    _get_bonus_path,
    _get_co_workorder_path,
    _get_logo_path,
    _get_mapping_path,
    _get_raw_dir,
    _get_shared_expense_path,
    _get_timesheet_path,
)
from .options import (
    COMPANY_OPTIONS,
    VOUCHER_DISPLAY_ORDER,
    VOUCHER_LABELS,
    VOUCHER_PRESETS,
    RunOptions,
    build_batch_layout_from_options,
    get_default_processing_period,
    requires_bank_data,
    requires_bonus_data,
    requires_co_data,
    requires_shared_expense_data,
)
from .pipeline import execute_payroll_run
from .precheck import run_startup_precheck


class MainWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.title('工资奖金凭证处理工具')
        self.root.configure(bg=COLOR_BG)
        self.root.resizable(True, True)
        self.root.minsize(700, 520)
        self.root.protocol('WM_DELETE_WINDOW', self._on_close)
        _apply_style(self.root)

        self._phase = 'GUIDE'
        self._q = queue.Queue()
        self._base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self._raw_dir = _get_raw_dir(self._base_dir)
        self._mapping_path = _get_mapping_path(self._base_dir)
        self._bank_dir = _get_bank_dir(self._base_dir)
        self._timesheet_path = _get_timesheet_path(self._base_dir)
        self._bonus_path = _get_bonus_path(self._base_dir)
        self._result = None
        default_year, default_month = get_default_processing_period()
        self._processing_year_var = tk.IntVar(value=default_year)
        self._processing_month_var = tk.IntVar(value=default_month)
        self._company_vars = {company: tk.BooleanVar(value=True) for company, _ in COMPANY_OPTIONS}
        self._voucher_vars = {voucher: tk.BooleanVar(value=True) for voucher in VOUCHER_DISPLAY_ORDER}
        self._suspend_option_refresh = False
        self._precheck_result = None
        self._precheck_signature = None
        self._precheck_running = False
        self._guide_content = None
        self._checklist_section = None
        self._paths_section = None
        self._requirements_section = None
        self._precheck_text = None
        self._precheck_status_var = tk.StringVar(value='尚未检查当前选择')
        self._precheck_summary_frame = None
        self._precheck_company_frame = None
        self._precheck_detail_container = None
        self._precheck_detail_expanded = False
        self._precheck_toggle_btn = None
        self._processing_year_var.trace_add('write', lambda *_: self._on_selection_change())
        self._processing_month_var.trace_add('write', lambda *_: self._on_selection_change())

        self._logo_img = None
        try:
            from PIL import Image, ImageTk

            logo_path = _get_logo_path(self._base_dir)
            if logo_path:
                img = Image.open(logo_path).resize((51, 50), Image.LANCZOS)
                self._logo_img = ImageTk.PhotoImage(img)
        except Exception:
            pass

        self._build_header()
        self._body = tk.Frame(self.root, bg=COLOR_BG)
        self._body.pack(fill='both', expand=True)
        self._footer = tk.Frame(self._body, bg=COLOR_BG, pady=12)
        self._footer.pack(side='bottom', fill='x')

        self._show_guide()
        self.root.update_idletasks()
        self.root.deiconify()

    def _on_close(self):
        if self._phase == 'RUN':
            return
        self.root.destroy()

    def _build_header(self):
        self._header_frame = tk.Frame(self.root, bg=COLOR_HEADER_BG, pady=0)
        self._header_frame.pack(fill='x')
        tk.Frame(self._header_frame, bg=COLOR_PRIMARY, height=3).pack(fill='x')
        inner = tk.Frame(self._header_frame, bg=COLOR_HEADER_BG, pady=10)
        inner.pack(fill='x')
        self._header_inner = inner

        if self._logo_img:
            tk.Label(inner, image=self._logo_img, bg=COLOR_HEADER_BG).pack(side='left', padx=(14, 12))

        title_frame = tk.Frame(inner, bg=COLOR_HEADER_BG)
        title_frame.pack(side='left', fill='both', expand=True)
        self._header_title = tk.Label(
            title_frame,
            text='工资奖金凭证处理工具',
            font=('微软雅黑', 15, 'bold'),
            fg=COLOR_PRIMARY,
            bg=COLOR_HEADER_BG,
        )
        self._header_title.pack(anchor='w')
        self._header_sub = tk.Label(
            title_frame,
            text='先选择处理月份，再执行预检并生成 A1-A10 凭证',
            font=('微软雅黑', 9),
            fg=COLOR_TEXT_SUB,
            bg=COLOR_HEADER_BG,
        )
        self._header_sub.pack(anchor='w', pady=(2, 0))

        tk.Label(
            inner,
            text='Developed by Earnest Yin',
            font=('Consolas', 10, 'bold'),
            fg='#888888',
            bg=COLOR_HEADER_BG,
        ).pack(side='right', padx=(0, 16), anchor='s')

    def _set_header(self, title, sub='', bg=COLOR_HEADER_BG):
        for widget in self._header_frame.winfo_children():
            if isinstance(widget, tk.Frame) and widget is not self._header_inner:
                widget.configure(bg=bg if bg != COLOR_HEADER_BG else COLOR_PRIMARY)
                break
        self._header_title.configure(text=title, fg=COLOR_PRIMARY if bg == COLOR_HEADER_BG else bg)
        self._header_sub.configure(text=sub)

    def _clear_body(self):
        for widget in self._body.winfo_children():
            if widget is not self._footer:
                widget.destroy()

    def _clear_footer(self):
        for widget in self._footer.winfo_children():
            widget.destroy()

    def _get_run_options(self):
        default_year, default_month = get_default_processing_period()
        try:
            posting_year = int(self._processing_year_var.get())
            posting_month = int(self._processing_month_var.get())
        except tk.TclError:
            posting_year, posting_month = default_year, default_month
        companies = tuple(company for company, _ in COMPANY_OPTIONS if self._company_vars[company].get())
        vouchers = tuple(voucher for voucher in VOUCHER_DISPLAY_ORDER if self._voucher_vars[voucher].get())
        return RunOptions(
            companies=companies,
            vouchers=vouchers,
            posting_year=posting_year,
            posting_month=posting_month,
        )

    def _get_batch_layout(self, run_options):
        return build_batch_layout_from_options(self._base_dir, run_options)

    def _current_selection_signature(self):
        options = self._get_run_options()
        return options.processing_year, options.processing_month, options.companies, options.vouchers

    def _get_raw_files_for_selection(self, run_options):
        layout = self._get_batch_layout(run_options)
        return _find_raw_files(layout.raw_dir, run_options.payroll_year, run_options.payroll_month)

    def _mark_precheck_stale(self):
        self._precheck_result = None
        self._precheck_signature = None
        if self._precheck_text is not None:
            self._render_precheck_result()

    def _on_selection_change(self):
        if self._suspend_option_refresh:
            return
        self._mark_precheck_stale()
        if self._guide_content is not None and self._phase == 'GUIDE':
            self._refresh_guide_sections()

    def _set_voucher_selection(self, voucher_ids):
        selected = set(voucher_ids)
        self._suspend_option_refresh = True
        for voucher, var in self._voucher_vars.items():
            var.set(voucher in selected)
        self._suspend_option_refresh = False
        self._on_selection_change()

    def _voucher_requirement_text(self, voucher):
        mapping = {
            'A1': '工资单、Mapping、工时、银行流水',
            'A2': '工资单、Mapping、工时、银行流水',
            'A3': '工资单、Mapping、工时、银行流水',
            'A4': '工资单、Mapping、工时',
            'A5': '工资单、Mapping、工时',
            'A6': '工资单、Mapping、工时',
            'A7': '工资单、Mapping、工时、奖金数据',
            'A8': '工资单、Mapping、工时、奖金数据',
            'A9': '工资单、Mapping、工时、CO工单分摊',
            'A10': '工资单、Mapping、工时、待分摊费用',
        }
        return mapping.get(voucher, '工资单、Mapping、工时')

    def _build_section_title(self, parent, text):
        tk.Label(
            parent,
            text=text,
            font=('微软雅黑', 11, 'bold'),
            fg=COLOR_PRIMARY,
            bg=COLOR_BG,
            anchor='w',
        ).pack(fill='x')
        tk.Frame(parent, bg=COLOR_PRIMARY, height=2).pack(fill='x', pady=(2, 10))

    def _status_palette(self, status):
        mapping = {
            'ok': {'fg': '#000000', 'bg': COLOR_SUCCESS, 'label': '通过'},
            'warn': {'fg': '#000000', 'bg': COLOR_WARNING, 'label': '预警'},
            'error': {'fg': '#FFFFFF', 'bg': COLOR_DANGER, 'label': '阻断'},
            'info': {'fg': COLOR_TEXT_MAIN, 'bg': '#444444', 'label': '未检查'},
        }
        return mapping.get(status, mapping['info'])

    def _build_metric_card(self, parent, title, value, status='info'):
        palette = self._status_palette(status)
        card = tk.Frame(parent, bg=COLOR_CARD, highlightbackground=COLOR_BORDER, highlightthickness=1, padx=12, pady=10)
        card.pack(side='left', fill='both', expand=True, padx=(0, 10))
        tk.Label(card, text=title, font=('微软雅黑', 9), fg=COLOR_TEXT_SUB, bg=COLOR_CARD, anchor='w').pack(fill='x')
        tk.Label(
            card,
            text=value,
            font=('微软雅黑', 16, 'bold'),
            fg=palette['bg'] if status != 'info' else COLOR_TEXT_MAIN,
            bg=COLOR_CARD,
            anchor='w',
        ).pack(fill='x', pady=(6, 0))

    def _build_status_badge(self, parent, text, status, side='left', padx=(0, 6), pady=0):
        palette = self._status_palette(status)
        tk.Label(
            parent,
            text=f' {text} ',
            font=('微软雅黑', 8, 'bold'),
            fg=palette['fg'],
            bg=palette['bg'],
            padx=4,
            pady=2,
        ).pack(side=side, padx=padx, pady=pady)

    def _toggle_precheck_details(self):
        self._precheck_detail_expanded = not self._precheck_detail_expanded
        if self._precheck_detail_container is not None:
            if self._precheck_detail_expanded:
                self._precheck_detail_container.pack(fill='both', expand=True, pady=(10, 0))
            else:
                self._precheck_detail_container.pack_forget()
        if self._precheck_toggle_btn is not None:
            self._precheck_toggle_btn.configure(text='收起详细明细' if self._precheck_detail_expanded else '展开详细明细')

    def _build_file_status_row(self, parent, level_text, level_fg, level_bg, file_text, status_text, status_fg, status_bg, note):
        row = tk.Frame(parent, bg=COLOR_BG, pady=4)
        row.pack(fill='x')
        tk.Label(
            row,
            text=f' {level_text} ',
            font=('微软雅黑', 8, 'bold'),
            fg=level_fg,
            bg=level_bg,
            padx=4,
        ).pack(side='left', padx=(0, 8))
        tk.Label(
            row,
            text=file_text,
            font=('微软雅黑', 10, 'bold'),
            fg=COLOR_TEXT_MAIN,
            bg=COLOR_BG,
            width=40,
            anchor='w',
        ).pack(side='left')
        tk.Label(
            row,
            text=f' {status_text} ',
            font=('微软雅黑', 8, 'bold'),
            fg=status_fg,
            bg=status_bg,
            padx=4,
        ).pack(side='left', padx=(0, 10))
        tk.Label(
            row,
            text=note,
            font=('微软雅黑', 9),
            fg=COLOR_TEXT_SUB,
            bg=COLOR_BG,
            anchor='w',
        ).pack(side='left')

    def _build_selection_section(self, parent):
        card = tk.Frame(parent, bg=COLOR_CARD, highlightbackground='#3D3D3D', highlightthickness=1, padx=14, pady=12)
        card.pack(fill='x', pady=(0, 14))

        tk.Label(
            card,
            text='本次生成范围',
            font=('微软雅黑', 11, 'bold'),
            fg=COLOR_PRIMARY,
            bg=COLOR_CARD,
            anchor='w',
        ).pack(fill='x')
        tk.Frame(card, bg=COLOR_PRIMARY, height=2).pack(fill='x', pady=(2, 10))

        period_row = tk.Frame(card, bg=COLOR_CARD)
        period_row.pack(fill='x', pady=(0, 8))
        tk.Label(period_row, text='处理月份：', font=('微软雅黑', 10, 'bold'), fg=COLOR_TEXT_MAIN, bg=COLOR_CARD).pack(side='left')
        tk.Spinbox(
            period_row,
            from_=2024,
            to=2035,
            textvariable=self._processing_year_var,
            width=6,
            font=('微软雅黑', 10),
            justify='center',
            command=self._on_selection_change,
            bg=COLOR_BG,
            fg=COLOR_TEXT_MAIN,
            buttonbackground='#444444',
            insertbackground=COLOR_PRIMARY,
            relief='flat',
        ).pack(side='left', padx=(8, 6))
        tk.Label(period_row, text='年', font=('微软雅黑', 10), fg=COLOR_TEXT_MAIN, bg=COLOR_CARD).pack(side='left')
        tk.Spinbox(
            period_row,
            from_=1,
            to=12,
            format='%02.0f',
            textvariable=self._processing_month_var,
            width=4,
            font=('微软雅黑', 10),
            justify='center',
            command=self._on_selection_change,
            bg=COLOR_BG,
            fg=COLOR_TEXT_MAIN,
            buttonbackground='#444444',
            insertbackground=COLOR_PRIMARY,
            relief='flat',
        ).pack(side='left', padx=(8, 6))
        tk.Label(period_row, text='月', font=('微软雅黑', 10), fg=COLOR_TEXT_MAIN, bg=COLOR_CARD).pack(side='left')
        tk.Label(
            period_row,
            text='每月 1-5 日默认回到上个月；工资单会自动匹配上一月文件',
            font=('微软雅黑', 9),
            fg=COLOR_TEXT_SUB,
            bg=COLOR_CARD,
        ).pack(side='left', padx=(14, 0))

        company_row = tk.Frame(card, bg=COLOR_CARD)
        company_row.pack(fill='x', pady=(0, 8))
        tk.Label(company_row, text='公司：', font=('微软雅黑', 10, 'bold'), fg=COLOR_TEXT_MAIN, bg=COLOR_CARD).pack(side='left')
        for company, label in COMPANY_OPTIONS:
            tk.Checkbutton(
                company_row,
                text=label,
                variable=self._company_vars[company],
                command=self._on_selection_change,
                bg=COLOR_CARD,
                fg=COLOR_TEXT_MAIN,
                selectcolor=COLOR_BG,
                activebackground=COLOR_CARD,
                activeforeground=COLOR_TEXT_MAIN,
                font=('微软雅黑', 10),
            ).pack(side='left', padx=(8, 14))

        voucher_head = tk.Frame(card, bg=COLOR_CARD)
        voucher_head.pack(fill='x', pady=(4, 6))
        tk.Label(voucher_head, text='凭证：', font=('微软雅黑', 10, 'bold'), fg=COLOR_TEXT_MAIN, bg=COLOR_CARD).pack(
            side='left'
        )
        preset_buttons = [
            ('全选', VOUCHER_PRESETS['all']),
            ('A1-A4', VOUCHER_PRESETS['actual']),
            ('A5-A6', VOUCHER_PRESETS['accrual']),
            ('A7-A8', VOUCHER_PRESETS['bonus']),
            ('A9', VOUCHER_PRESETS['co']),
            ('A10', VOUCHER_PRESETS['rd']),
        ]
        for label, voucher_ids in preset_buttons:
            _btn(voucher_head, label, '#444444', lambda ids=voucher_ids: self._set_voucher_selection(ids), padx=10, pady=3).pack(
                side='left', padx=(0, 8)
            )

        voucher_grid = tk.Frame(card, bg=COLOR_CARD)
        voucher_grid.pack(fill='x')
        for idx, voucher in enumerate(VOUCHER_DISPLAY_ORDER):
            tk.Checkbutton(
                voucher_grid,
                text=VOUCHER_LABELS[voucher],
                variable=self._voucher_vars[voucher],
                command=self._on_selection_change,
                bg=COLOR_CARD,
                fg=COLOR_TEXT_MAIN,
                selectcolor=COLOR_BG,
                activebackground=COLOR_CARD,
                activeforeground=COLOR_TEXT_MAIN,
                font=('微软雅黑', 9),
                anchor='w',
                width=24,
            ).grid(row=idx // 2, column=idx % 2, sticky='w', padx=(0, 18), pady=3)

    def _collect_missing_inputs(self, run_options, raw_files):
        missing = []
        layout = self._get_batch_layout(run_options)
        payroll_period = run_options.payroll_period
        if not run_options.companies:
            missing.append('至少选择 1 家公司')
        if not run_options.vouchers:
            missing.append('至少选择 1 张凭证')
        if len(raw_files) != 1:
            missing.append(f'工资单目录下必须且只能有 1 份 {run_options.payroll_label} 的原始工资单')
        if not os.path.exists(layout.mapping_path):
            missing.append('缺少 Mapping表.xlsx')
        if not os.path.exists(layout.timesheet_path):
            missing.append('缺少 工时数据.xlsx')
        if requires_bonus_data(run_options) and not os.path.exists(layout.bonus_path):
            missing.append('A7/A8 需要 年终奖计提文件')
        if requires_bank_data(run_options) and not _find_bank_files(layout.bank_dir):
            missing.append('A1-A3 需要 银行流水 文件')
        if run_options.wants_voucher('A9') and '耐数电子' not in run_options.companies:
            missing.append('A9 仅适用于耐数电子，请勾选 2050')
        if run_options.wants_voucher('A9'):
            co_path = _get_co_workorder_path(self._base_dir, payroll_period[0], payroll_period[1])
            if not os.path.exists(co_path):
                missing.append('A9 需要 CO工单分摊.xlsx')
        if requires_shared_expense_data(run_options):
            shared_path = _get_shared_expense_path(self._base_dir, payroll_period[0], payroll_period[1])
            if not os.path.exists(shared_path):
                missing.append('A10 需要 待分摊费用YYMM.xlsx')
        return missing

    def _refresh_checklist_section(self, run_options, raw_files):
        for widget in self._checklist_section.winfo_children():
            widget.destroy()

        layout = self._get_batch_layout(run_options)
        payroll_period = run_options.payroll_period
        co_path = _get_co_workorder_path(self._base_dir, payroll_period[0], payroll_period[1])
        shared_path = _get_shared_expense_path(self._base_dir, payroll_period[0], payroll_period[1])
        mapping_exists = os.path.exists(layout.mapping_path)
        timesheet_exists = os.path.exists(layout.timesheet_path)
        bonus_exists = os.path.exists(layout.bonus_path)
        bank_exists = len(_find_bank_files(layout.bank_dir)) > 0
        co_exists = bool(co_path) and os.path.exists(co_path)
        shared_exists = bool(shared_path) and os.path.exists(shared_path)

        self._build_section_title(self._checklist_section, '所需资料清单')
        file_rows = [
            (
                '必须',
                COLOR_DANGER,
                '#3A1A1A',
                os.path.join(
                    os.path.relpath(layout.raw_dir, self._base_dir),
                    f'人力成本研发项目分摊{payroll_period[0]}{payroll_period[1]:02d} - to财务-原始.xlsx',
                ),
                '✓ 已找到' if len(raw_files) == 1 else ('✗ 未找到' if not raw_files else '✗ 找到多个'),
                '#000000' if len(raw_files) == 1 else '#FFFFFF',
                COLOR_SUCCESS if len(raw_files) == 1 else COLOR_DANGER,
                f'全部凭证共用，对应工资所属月份 {run_options.payroll_label}',
            ),
            (
                '必须',
                COLOR_DANGER,
                '#3A1A1A',
                os.path.relpath(layout.mapping_path, self._base_dir),
                '✓ 已找到' if mapping_exists else '✗ 未找到',
                '#000000' if mapping_exists else '#FFFFFF',
                COLOR_SUCCESS if mapping_exists else COLOR_DANGER,
                '全部凭证共用',
            ),
            (
                '必须',
                COLOR_DANGER,
                '#3A1A1A',
                os.path.relpath(layout.timesheet_path, self._base_dir),
                '✓ 已找到' if timesheet_exists else '✗ 未找到',
                '#000000' if timesheet_exists else '#FFFFFF',
                COLOR_SUCCESS if timesheet_exists else COLOR_DANGER,
                '全部凭证共用，且会检查 KOK3/内部订单',
            ),
            (
                '必须' if requires_bank_data(run_options) else '按需',
                COLOR_DANGER if requires_bank_data(run_options) else '#000000',
                '#3A1A1A' if requires_bank_data(run_options) else COLOR_WARNING,
                os.path.join(os.path.relpath(layout.bank_dir, self._base_dir), BANK_FILE_PATTERN),
                '✓ 已找到' if bank_exists else ('✗ 未找到' if requires_bank_data(run_options) else '· 未提供'),
                '#000000' if bank_exists else '#FFFFFF',
                COLOR_SUCCESS if bank_exists else (COLOR_DANGER if requires_bank_data(run_options) else '#666666'),
                f'A1-A3，启动时会全目录扫描并按处理月份 {run_options.processing_label} 筛选',
            ),
            (
                '必须' if requires_bonus_data(run_options) else '按需',
                COLOR_DANGER if requires_bonus_data(run_options) else '#000000',
                '#3A1A1A' if requires_bonus_data(run_options) else COLOR_WARNING,
                os.path.relpath(layout.bonus_path, self._base_dir),
                '✓ 已找到' if bonus_exists else ('✗ 未找到' if requires_bonus_data(run_options) else '· 未提供'),
                '#000000' if bonus_exists else '#FFFFFF',
                COLOR_SUCCESS if bonus_exists else (COLOR_DANGER if requires_bonus_data(run_options) else '#666666'),
                'A7-A8',
            ),
            (
                '必须' if run_options.wants_voucher('A9') else '按需',
                COLOR_DANGER if run_options.wants_voucher('A9') else '#000000',
                '#3A1A1A' if run_options.wants_voucher('A9') else COLOR_WARNING,
                os.path.relpath(co_path, self._base_dir),
                '✓ 已找到' if co_exists else ('✗ 未找到' if run_options.wants_voucher('A9') else '· 未提供'),
                '#000000' if co_exists else '#FFFFFF',
                COLOR_SUCCESS if co_exists else (COLOR_DANGER if run_options.wants_voucher('A9') else '#666666'),
                f'A9，对应处理月份 {run_options.processing_label}',
            ),
            (
                '必须' if requires_shared_expense_data(run_options) else '按需',
                COLOR_DANGER if requires_shared_expense_data(run_options) else '#000000',
                '#3A1A1A' if requires_shared_expense_data(run_options) else COLOR_WARNING,
                os.path.relpath(shared_path, self._base_dir),
                '✓ 已找到' if shared_exists else ('✗ 未找到' if requires_shared_expense_data(run_options) else '· 未提供'),
                '#000000' if shared_exists else '#FFFFFF',
                COLOR_SUCCESS if shared_exists else (COLOR_DANGER if requires_shared_expense_data(run_options) else '#666666'),
                f'A10，对应处理月份 {run_options.processing_label}',
            ),
        ]
        for row_data in file_rows:
            self._build_file_status_row(self._checklist_section, *row_data)

    def _refresh_requirements_section(self, run_options):
        for widget in self._requirements_section.winfo_children():
            widget.destroy()

        self._build_section_title(self._requirements_section, '按凭证资料需求')
        tk.Label(
            self._requirements_section,
            text=f'当前处理月份：{run_options.processing_label}；对应工资所属月份：{run_options.payroll_label}',
            font=('微软雅黑', 9),
            fg=COLOR_TEXT_SUB,
            bg=COLOR_BG,
            anchor='w',
        ).pack(fill='x', pady=(0, 6))
        for voucher in run_options.vouchers:
            row = tk.Frame(self._requirements_section, bg=COLOR_BG, pady=2)
            row.pack(fill='x')
            tk.Label(
                row,
                text=VOUCHER_LABELS[voucher],
                font=('微软雅黑', 9, 'bold'),
                fg=COLOR_TEXT_MAIN,
                bg=COLOR_BG,
                width=24,
                anchor='w',
            ).pack(side='left')
            tk.Label(
                row,
                text=self._voucher_requirement_text(voucher),
                font=('微软雅黑', 9),
                fg=COLOR_TEXT_SUB,
                bg=COLOR_BG,
                anchor='w',
            ).pack(side='left')

    def _refresh_paths_section(self, run_options, raw_files):
        for widget in self._paths_section.winfo_children():
            widget.destroy()

        layout = self._get_batch_layout(run_options)
        path_lines = [
            f'当前处理月份：{run_options.processing_label}',
            f'对应工资所属月份：{run_options.payroll_label}',
            f'月度输入目录：{layout.monthly_input_root}',
            f'运行输出目录：{layout.run_output_root}',
            f'留痕归档目录：{layout.archive_root}',
            f'工资单目录：{layout.raw_dir}',
            f'Mapping 表：{layout.mapping_path}',
            f'银行流水目录：{layout.bank_dir}（全目录扫描后按处理月份筛选）',
            f'工时数据：{layout.timesheet_path}',
            f'奖金数据：{layout.bonus_path}',
        ]
        if raw_files:
            show_name = os.path.basename(raw_files[0]) if len(raw_files) == 1 else '；'.join(os.path.basename(p) for p in raw_files[:2])
            path_lines.insert(4, f'当前识别工资单：{show_name}')
        payroll_period = run_options.payroll_period
        co_path = _get_co_workorder_path(self._base_dir, payroll_period[0], payroll_period[1])
        shared_path = _get_shared_expense_path(self._base_dir, payroll_period[0], payroll_period[1])
        path_lines.append(f'CO工单分摊：{co_path}')
        path_lines.append(f'待分摊费用：{shared_path}')

        for line in path_lines:
            tk.Label(
                self._paths_section,
                text=line,
                font=('微软雅黑', 9),
                fg=COLOR_TEXT_SUB,
                bg=COLOR_BG,
                anchor='w',
            ).pack(fill='x', pady=(2, 0))

    def _render_precheck_result(self):
        if self._precheck_text is None or self._precheck_summary_frame is None or self._precheck_company_frame is None:
            return

        for widget in self._precheck_summary_frame.winfo_children():
            widget.destroy()
        for widget in self._precheck_company_frame.winfo_children():
            widget.destroy()

        self._precheck_text.configure(state='normal')
        self._precheck_text.delete('1.0', 'end')
        self._precheck_text.tag_configure('ok', foreground=COLOR_SUCCESS)
        self._precheck_text.tag_configure('warn', foreground=COLOR_WARNING)
        self._precheck_text.tag_configure('err', foreground=COLOR_DANGER)
        self._precheck_text.tag_configure('info', foreground=COLOR_TEXT_SUB)

        run_options = self._get_run_options()
        signature = self._current_selection_signature()
        if self._precheck_running:
            self._precheck_status_var.set('正在检查当前选择，请稍候…')
            summary_row = tk.Frame(self._precheck_summary_frame, bg=COLOR_CARD)
            summary_row.pack(fill='x')
            self._build_metric_card(summary_row, '总体状态', '检查中', 'info')
            self._build_metric_card(summary_row, '阻断项', '-', 'info')
            self._build_metric_card(summary_row, '预警项', '-', 'info')
            self._build_metric_card(summary_row, '处理月份', run_options.processing_yymm, 'info')
            self._precheck_text.insert('end', '预检进行中，将按当前公司和凭证选择逐项校验基础数据。\n', 'info')
        elif self._precheck_result is None or self._precheck_signature != signature:
            self._precheck_status_var.set('尚未检查当前选择')
            summary_row = tk.Frame(self._precheck_summary_frame, bg=COLOR_CARD)
            summary_row.pack(fill='x')
            self._build_metric_card(summary_row, '总体状态', '未检查', 'info')
            self._build_metric_card(summary_row, '已选公司', str(len(run_options.companies)), 'info')
            self._build_metric_card(summary_row, '已选凭证', str(len(run_options.vouchers)), 'info')
            self._build_metric_card(summary_row, '操作建议', '先检查', 'info')
            self._precheck_text.insert('end', '请点击“检查当前选择”，查看当前凭证组合是否具备运行条件。\n', 'info')
        else:
            summary = self._precheck_result.get('summary', '预检完成')
            self._precheck_status_var.set(summary)
            overall_status = self._precheck_result.get('overall_status', 'info')
            blockers = self._precheck_result.get('blockers', [])
            warnings = self._precheck_result.get('warnings', [])
            summary_row = tk.Frame(self._precheck_summary_frame, bg=COLOR_CARD)
            summary_row.pack(fill='x')
            self._build_metric_card(summary_row, '总体状态', self._status_palette(overall_status)['label'], overall_status)
            self._build_metric_card(summary_row, '阻断项', str(len(blockers)), 'error' if blockers else 'ok')
            self._build_metric_card(summary_row, '预警项', str(len(warnings)), 'warn' if warnings else 'ok')
            bank_stats = self._precheck_result.get('bank_scan_stats', {})
            fourth_title = '可运行'
            fourth_value = '是' if self._precheck_result.get('can_run') else '否'
            fourth_status = 'ok' if self._precheck_result.get('can_run') else 'error'
            if requires_bank_data(run_options) and bank_stats:
                fourth_title = '银行去重'
                fourth_value = str(bank_stats.get('deduped_row_count', 0))
                fourth_status = 'warn' if bank_stats.get('deduped_row_count', 0) else 'ok'
            self._build_metric_card(summary_row, fourth_title, fourth_value, fourth_status)

            company_checks = self._precheck_result.get('company_checks', {})
            for company, company_data in company_checks.items():
                card = tk.Frame(
                    self._precheck_company_frame,
                    bg=COLOR_CARD,
                    highlightbackground=COLOR_BORDER,
                    highlightthickness=1,
                    padx=12,
                    pady=10,
                )
                card.pack(fill='x', pady=(0, 8))
                head = tk.Frame(card, bg=COLOR_CARD)
                head.pack(fill='x')
                tk.Label(
                    head,
                    text=company,
                    font=('微软雅黑', 10, 'bold'),
                    fg=COLOR_TEXT_MAIN,
                    bg=COLOR_CARD,
                    anchor='w',
                ).pack(side='left')
                self._build_status_badge(head, self._status_palette(company_data.get('status', 'info'))['label'], company_data.get('status', 'info'), side='right', padx=(6, 0))

                meta = tk.Frame(card, bg=COLOR_CARD)
                meta.pack(fill='x', pady=(6, 6))
                self._build_status_badge(meta, f'阻断 {len(company_data.get("blockers", []))}', 'error' if company_data.get('blockers') else 'ok')
                self._build_status_badge(meta, f'预警 {len(company_data.get("warnings", []))}', 'warn' if company_data.get('warnings') else 'ok')

                voucher_row = tk.Frame(card, bg=COLOR_CARD)
                voucher_row.pack(fill='x')
                for voucher in run_options.vouchers:
                    voucher_info = company_data.get('voucher_status', {}).get(voucher, {'status': 'info', 'message': '未检查'})
                    item = tk.Frame(voucher_row, bg='#222222', padx=8, pady=6)
                    item.pack(side='left', padx=(0, 8), pady=(0, 4))
                    tk.Label(
                        item,
                        text=voucher,
                        font=('Consolas', 9, 'bold'),
                        fg=COLOR_TEXT_MAIN,
                        bg='#222222',
                    ).pack(anchor='w')
                    tk.Label(
                        item,
                        text=self._status_palette(voucher_info.get('status', 'info'))['label'],
                        font=('微软雅黑', 8),
                        fg=self._status_palette(voucher_info.get('status', 'info'))['bg'],
                        bg='#222222',
                    ).pack(anchor='w')

            for item in self._precheck_result.get('display_lines', []):
                tag = item['status']
                if tag == 'error':
                    tag = 'err'
                elif tag not in ('ok', 'warn', 'err'):
                    tag = 'info'
                self._precheck_text.insert('end', item['text'] + '\n', tag)
            report_path = self._precheck_result.get('report_path')
            if report_path:
                self._precheck_text.insert('end', f'\n预检报告已写入：{report_path}\n', 'info')
        self._precheck_text.configure(state='disabled')

    def _start_precheck(self):
        if self._precheck_running:
            return
        self._precheck_running = True
        signature = self._current_selection_signature()
        self._precheck_signature = signature
        self._precheck_result = None
        self._render_precheck_result()
        self._refresh_footer()

        run_options = self._get_run_options()
        layout = self._get_batch_layout(run_options)

        def worker():
            try:
                result = run_startup_precheck(self._base_dir, layout.mapping_path, layout.bank_dir, run_options)
            except Exception as exc:
                result = {
                    'summary': str(exc),
                    'display_lines': [{'status': 'error', 'text': str(exc)}],
                    'can_run': False,
                    'overall_status': 'error',
                }
            self.root.after(0, lambda: self._finish_precheck(signature, result))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_precheck(self, signature, result):
        self._precheck_running = False
        self._precheck_signature = signature
        self._precheck_result = result
        self._render_precheck_result()
        self._refresh_footer()

    def _refresh_footer(self):
        self._clear_footer()
        run_options = self._get_run_options()
        raw_files = self._get_raw_files_for_selection(run_options)
        missing_items = self._collect_missing_inputs(run_options, raw_files)
        signature = self._current_selection_signature()
        precheck_ready = self._precheck_result is not None and self._precheck_signature == signature and not self._precheck_running
        precheck_can_run = precheck_ready and self._precheck_result.get('can_run', False)

        check_btn = _btn(self._footer, '检查当前选择', '#444444', self._start_precheck, padx=16, pady=6)
        if self._precheck_running:
            check_btn.configure(state='disabled')
        check_btn.pack(side='left', padx=(30, 10))

        start_btn = _btn(self._footer, '  文件已准备好，开始运行  ', COLOR_SUCCESS, self._start_run)
        if missing_items or not precheck_can_run or self._precheck_running:
            start_btn.configure(state='disabled')
        start_btn.pack(side='left', padx=(0, 10))

        if missing_items:
            message = '当前还不能运行：' + '；'.join(missing_items[:3])
            color = COLOR_DANGER
        elif self._precheck_running:
            message = '正在检查当前选择，检查完成后会自动更新可运行状态'
            color = COLOR_WARNING
        elif not precheck_ready:
            message = '请先点击“检查当前选择”，确认当前凭证组合具备运行条件'
            color = COLOR_WARNING
        elif not self._precheck_result.get('can_run', False):
            blockers = self._precheck_result.get('blockers', [])
            message = '当前还不能运行：' + '；'.join(blockers[:2] or ['存在阻断项'])
            color = COLOR_DANGER
        else:
            message = self._precheck_result.get('summary', '预检通过，可运行')
            color = COLOR_SUCCESS

        tk.Label(
            self._footer,
            text=message,
            font=('微软雅黑', 9, 'bold'),
            fg=color,
            bg=COLOR_BG,
        ).pack(side='left', padx=(10, 0))
        _btn(self._footer, '取消', '#444444', self.root.destroy, padx=16, pady=6).pack(side='right', padx=(10, 30))

    def _refresh_guide_sections(self):
        run_options = self._get_run_options()
        raw_files = self._get_raw_files_for_selection(run_options)
        self._refresh_checklist_section(run_options, raw_files)
        self._refresh_requirements_section(run_options)
        self._refresh_paths_section(run_options, raw_files)
        self._render_precheck_result()
        self._refresh_footer()

    def _show_guide(self):
        self._phase = 'GUIDE'
        self._clear_body()
        self._set_header('工资奖金凭证处理工具', '先选择处理月份，再按所选公司和 A1-A10 凭证执行预检与生成')

        self._guide_content = tk.Frame(self._body, bg=COLOR_BG, padx=30, pady=14)
        self._guide_content.pack(fill='both', expand=True)

        self._build_selection_section(self._guide_content)

        self._checklist_section = tk.Frame(self._guide_content, bg=COLOR_BG)
        self._checklist_section.pack(fill='x')
        self._requirements_section = tk.Frame(self._guide_content, bg=COLOR_BG)
        self._requirements_section.pack(fill='x', pady=(10, 0))
        self._paths_section = tk.Frame(self._guide_content, bg=COLOR_BG)
        self._paths_section.pack(fill='x', pady=(8, 0))

        tk.Label(self._guide_content, text='', bg=COLOR_BG, height=1).pack()
        self._build_section_title(self._guide_content, '运行步骤')

        for num, text in [
            ('①', '先确认处理月份，系统会自动定位对应批次目录并匹配上一月唯一工资单'),
            ('②', '对首个工作表 A/B 列空白单元格按上一行非空值向下填充'),
            ('③', '校验 Q 列实发金额是否等于 E 列减 K:P 列个人承担金额'),
            ('④', '根据 Mapping 表生成 S 列成本中心，并根据工时数据生成 T 列内部订单'),
            ('⑤', '银行流水会扫描整个目录，并按当前处理月份筛选后自动去重'),
            ('⑥', '若工时项目未全部匹配到内部订单，则终止后续处理'),
            ('⑦', '先检查当前选择的数据支撑情况，再生成所选公司与 A1-A10 凭证及留痕文件'),
        ]:
            row = tk.Frame(self._guide_content, bg=COLOR_BG, pady=3)
            row.pack(fill='x')
            _circle_label(row, num, size=26, parent_bg=COLOR_BG).pack(side='left', padx=(0, 10))
            tk.Label(row, text=text, font=('微软雅黑', 10), fg=COLOR_TEXT_MAIN, bg=COLOR_BG, anchor='w').pack(side='left')

        tip = tk.Frame(self._guide_content, bg='#252500', pady=8, padx=12)
        tip.pack(fill='x', pady=(14, 0))
        tk.Label(
            tip,
            text='输出结果会写入标准批次目录：月度输入、运行输出、归档留痕三层分开；工资单文件名仍沿用工资所属月份，凭证与留痕文件按处理月份和批次归档',
            font=('微软雅黑', 9),
            fg=COLOR_PRIMARY,
            bg='#252500',
            anchor='w',
        ).pack(fill='x')

        precheck_card = tk.Frame(
            self._guide_content,
            bg=COLOR_CARD,
            highlightbackground='#3D3D3D',
            highlightthickness=1,
            padx=14,
            pady=12,
        )
        precheck_card.pack(fill='both', expand=True, pady=(14, 0))
        tk.Label(
            precheck_card,
            text='预检结果',
            font=('微软雅黑', 11, 'bold'),
            fg=COLOR_PRIMARY,
            bg=COLOR_CARD,
            anchor='w',
        ).pack(fill='x')
        tk.Frame(precheck_card, bg=COLOR_PRIMARY, height=2).pack(fill='x', pady=(2, 8))
        tk.Label(
            precheck_card,
            textvariable=self._precheck_status_var,
            font=('微软雅黑', 9, 'bold'),
            fg=COLOR_TEXT_MAIN,
            bg=COLOR_CARD,
            anchor='w',
        ).pack(fill='x', pady=(0, 8))
        self._precheck_summary_frame = tk.Frame(precheck_card, bg=COLOR_CARD)
        self._precheck_summary_frame.pack(fill='x', pady=(0, 10))
        self._precheck_company_frame = tk.Frame(precheck_card, bg=COLOR_CARD)
        self._precheck_company_frame.pack(fill='x')

        toggle_row = tk.Frame(precheck_card, bg=COLOR_CARD)
        toggle_row.pack(fill='x', pady=(8, 0))
        self._precheck_toggle_btn = _btn(toggle_row, '展开详细明细', '#444444', self._toggle_precheck_details, padx=10, pady=4)
        self._precheck_toggle_btn.pack(side='left')

        self._precheck_detail_container = tk.Frame(precheck_card, bg=COLOR_CARD)
        self._precheck_text = tk.Text(
            self._precheck_detail_container,
            height=14,
            font=('Consolas', 10),
            bg=COLOR_CARD,
            fg='#E8E8E8',
            relief='flat',
            state='disabled',
            wrap='word',
            padx=8,
            pady=6,
            insertbackground=COLOR_PRIMARY,
            selectbackground='#3A3A1A',
        )
        self._precheck_text.pack(fill='both', expand=True)

        self._refresh_guide_sections()

    def _show_progress(self):
        self._phase = 'RUN'
        self._clear_body()
        self._clear_footer()
        self._set_header('正在运行，请稍候…', '')

        body = tk.Frame(self._body, bg=COLOR_BG, padx=24, pady=14)
        body.pack(fill='both', expand=True)

        self._step_var = tk.StringVar(value='准备中…')
        tk.Label(
            body,
            textvariable=self._step_var,
            font=('微软雅黑', 10, 'bold'),
            fg=COLOR_PRIMARY,
            bg=COLOR_BG,
            anchor='w',
        ).pack(fill='x')

        self._pb_var = tk.IntVar(value=0)
        self._pb = ttk.Progressbar(
            body,
            variable=self._pb_var,
            maximum=6,
            length=650,
            mode='determinate',
            style='Horizontal.TProgressbar',
        )
        self._pb.pack(fill='x', pady=(6, 12))

        tk.Label(body, text='运行日志', font=('微软雅黑', 9, 'bold'), fg=COLOR_TEXT_SUB, bg=COLOR_BG, anchor='w').pack(fill='x')
        frame = tk.Frame(body, bg=COLOR_CARD, highlightbackground=COLOR_BORDER, highlightthickness=1)
        frame.pack(fill='both', expand=True, pady=(4, 0))
        self._log_text = tk.Text(
            frame,
            height=12,
            font=('Consolas', 11),
            bg=COLOR_CARD,
            fg='#E8E8E8',
            relief='flat',
            state='disabled',
            wrap='word',
            padx=10,
            pady=8,
            insertbackground=COLOR_PRIMARY,
            selectbackground='#3A3A1A',
        )
        self._log_text.tag_configure('ok', foreground=COLOR_SUCCESS)
        self._log_text.tag_configure('warn', foreground=COLOR_WARNING)
        self._log_text.tag_configure('err', foreground=COLOR_DANGER)
        vsb = tk.Scrollbar(frame, command=self._log_text.yview, bg=COLOR_BORDER)
        self._log_text.configure(yscrollcommand=vsb.set)
        self._log_text.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

    def _append_log(self, text, tag=''):
        self._log_text.configure(state='normal')
        self._log_text.insert('end', text + '\n', tag)
        self._log_text.see('end')
        self._log_text.configure(state='disabled')

    def _show_result(self, success, summary):
        self._phase = 'DONE'
        color = COLOR_SUCCESS if success else COLOR_DANGER
        title = '处理完成' if success else '处理失败'
        self._set_header(title, summary, bg=color)
        self._clear_footer()

        if success and self._result:
            output_lines = [f'{company}：{path}' for company, path in self._result['output_paths'].items()]
            voucher_lines = [f'{company}：{path}' for company, path in self._result['voucher_paths'].items()]
            artifact_paths = self._result.get('artifact_paths', {})
            voucher_validation = self._result.get('voucher_validation_summary', {})
            voucher_validation_lines = []
            for company, validation in voucher_validation.items():
                group_balances = validation.get('group_balances', {})
                balance_text = '；'.join(
                    f'{voucher}平衡={"通过" if passed else "异常"}' for voucher, passed in group_balances.items()
                )
                cross_group_issues = validation.get('cross_group_issues', [])
                if cross_group_issues:
                    voucher_validation_lines.append(
                        f'{company}：{balance_text or "未生成凭证"}；A1-A4科目对冲异常 {len(cross_group_issues)} 项'
                    )
                    for issue in cross_group_issues:
                        voucher_validation_lines.append(f'  - {issue}')
                else:
                    cross_text = 'A1-A4科目对冲通过' if validation.get('cross_group_checked') else '未执行 A1-A4 对冲校验'
                    voucher_validation_lines.append(f'{company}：{balance_text or "未生成凭证"}；{cross_text}')

            text = (
                f"处理月份：{self._result.get('processing_period_label', '-')}\n"
                f"工资所属月份：{self._result.get('payroll_period_label', '-')}\n"
                f"首个工作表：{self._result['sheet_name']}\n"
                f"总行数：{self._result['row_count']}\n"
                f"A 列填充：{self._result['fill_a']} 个\n"
                f"B 列填充：{self._result['fill_b']} 个\n"
                f"实发校验异常：{self._result['validation_issue_count']} 行\n"
                f"成本中心未匹配：{self._result['cost_center_issue_count']} 行\n"
                f"内部订单未匹配：{self._result['internal_order_issue_count']} 行\n"
                f"银行核对异常项：{self._result['bank_issue_count']} 项\n"
                f"银行流水扫描/去重："
                f"{self._result.get('bank_scan_stats', {}).get('source_file_count', 0)} 个文件 / "
                f"{self._result.get('bank_scan_stats', {}).get('deduped_row_count', 0)} 行重复\n"
                f"工资/公积金/个税/社保异常项："
                f"{self._result['bank_salary_issue_count']}/"
                f"{self._result['bank_fund_issue_count']}/"
                f"{self._result['bank_tax_issue_count']}/"
                f"{self._result['bank_social_issue_count']}\n"
                f"行级异常总数：{self._result['issue_count']} 行\n"
                f"凭证校验：\n" + ('\n'.join(voucher_validation_lines) if voucher_validation_lines else '未生成凭证') + '\n'
                f"输出工资文件：\n" + ('\n'.join(output_lines) if output_lines else '未输出工资文件') + '\n'
                f"输出凭证文件：\n" + ('\n'.join(voucher_lines) if voucher_lines else '未输出凭证文件') + '\n'
                f"留痕文件：\n" + (
                    '\n'.join(f'{name}：{path}' for name, path in artifact_paths.items()) if artifact_paths else '未生成留痕文件'
                )
            )
            has_voucher_issue = any(
                validation.get('cross_group_issues')
                for validation in voucher_validation.values()
            )
            has_issue = self._result['issue_count'] or self._result['bank_issue_count'] or has_voucher_issue
            self._append_log(text, 'warn' if has_issue else 'ok')
            if has_voucher_issue:
                self._append_log('存在凭证科目对冲异常，请按上方明细复核。', 'warn')

        _btn(self._footer, '关闭', '#444444', self.root.destroy, padx=20, pady=6).pack(side='left', padx=(30, 10))

    def _poll(self):
        try:
            while True:
                kind, *payload = self._q.get_nowait()
                if kind == 'step':
                    current, text = payload
                    self._pb_var.set(current)
                    self._step_var.set(text)
                elif kind == 'log':
                    text, tag = payload
                    self._append_log(text, tag)
                elif kind == 'done':
                    success, summary, result = payload
                    self._result = result
                    self._show_result(success, summary)
                    return
        except queue.Empty:
            pass
        self.root.after(80, self._poll)

    def _start_run(self):
        run_options = self._get_run_options()
        raw_files = self._get_raw_files_for_selection(run_options)
        missing_items = self._collect_missing_inputs(run_options, raw_files)
        if missing_items:
            messagebox.showerror('无法运行', '；'.join(missing_items))
            return
        if self._precheck_running:
            messagebox.showerror('无法运行', '当前正在执行预检，请等待预检完成后再运行。')
            return
        if self._precheck_result is None or self._precheck_signature != self._current_selection_signature():
            messagebox.showerror('无法运行', '请先点击“检查当前选择”，确认当前公司和凭证组合具备运行条件。')
            return
        if not self._precheck_result.get('can_run', False):
            blockers = self._precheck_result.get('blockers', [])
            messagebox.showerror('无法运行', '当前选择仍存在阻断项：' + '；'.join(blockers[:3] or ['请先处理预检问题']))
            return

        input_path = raw_files[0]
        layout = self._get_batch_layout(run_options)
        payroll_year, payroll_month = run_options.payroll_period
        co_path = _get_co_workorder_path(self._base_dir, payroll_year, payroll_month)
        shared_path = _get_shared_expense_path(self._base_dir, payroll_year, payroll_month)
        self._show_progress()

        def log(text, tag=''):
            self._q.put(('log', text, tag))

        def run():
            try:
                self._q.put(('step', 1, '检查输入文件'))
                log(f'处理月份：{run_options.processing_label}')
                log(f'工资所属月份：{run_options.payroll_label}')
                log(f'批次目录：{layout.monthly_input_root}')
                log(f'输入文件：{os.path.basename(input_path)}')
                log('本次公司：' + '、'.join(run_options.companies))
                log('本次凭证：' + '、'.join(run_options.vouchers))
                log(f'输出目录：{layout.run_output_root}')
                log(f'留痕目录：{layout.archive_root}')
                log(f'Mapping 表：{os.path.basename(layout.mapping_path)}')
                log(f'工时数据：{layout.timesheet_path}')
                if requires_bonus_data(run_options):
                    log(f'奖金数据：{layout.bonus_path}')
                if run_options.wants_voucher('A9'):
                    log(f'CO工单分摊：{co_path}')
                if run_options.wants_voucher('A10'):
                    log(f'待分摊费用：{shared_path}')
                if requires_bank_data(run_options):
                    log(f'银行流水目录：{layout.bank_dir}（全目录扫描，按处理月份筛选，完全相同行自动去重）')

                self._q.put(('step', 2, '填充首个工作表 A/B 列空白'))
                self._q.put(('step', 3, '校验实发金额并匹配成本中心/内部订单'))
                self._q.put(('step', 4, '生成所选公司和凭证'))
                result = execute_payroll_run(
                    input_path,
                    self._base_dir,
                    layout.mapping_path,
                    layout.bank_dir,
                    log,
                    run_options=run_options,
                )

                self._q.put(('step', 5, '校验输出文件'))
                log('公司拆分文件写出完成', 'ok')
                self._q.put(('step', 6, '完成'))
                summary = '处理完成，所选公司与凭证已生成'
                total_issues = result['issue_count'] + result['bank_issue_count']
                voucher_issue_count = sum(
                    len(validation.get('cross_group_issues', []))
                    for validation in result.get('voucher_validation_summary', {}).values()
                )
                if total_issues:
                    summary = f"处理完成，但有 {total_issues} 项异常已标红"
                elif voucher_issue_count:
                    summary = f"处理完成，但有 {voucher_issue_count} 项凭证对冲异常"
                self._q.put(('done', True, summary, result))
            except Exception as exc:
                self._q.put(('log', str(exc), 'err'))
                self._q.put(('done', False, str(exc), None))

        threading.Thread(target=run, daemon=True).start()
        self.root.after(50, self._poll)

    def run(self):
        self.root.mainloop()


def main():
    MainWindow().run()
