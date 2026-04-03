import os
import queue
import threading
import tkinter as tk
from tkinter import messagebox, ttk

from .core import (
    BANK_FILE_PATTERN,
    COLOR_BG,
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
    _extract_payroll_period,
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
    requires_bank_data,
    requires_bonus_data,
    requires_co_data,
    requires_shared_expense_data,
)
from .pipeline import execute_payroll_run


class MainWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title('工资单首步处理工具')
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
        self._company_vars = {company: tk.BooleanVar(value=True) for company, _ in COMPANY_OPTIONS}
        self._voucher_vars = {voucher: tk.BooleanVar(value=True) for voucher in VOUCHER_DISPLAY_ORDER}

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
            text='工资单首步处理工具',
            font=('微软雅黑', 15, 'bold'),
            fg=COLOR_PRIMARY,
            bg=COLOR_HEADER_BG,
        )
        self._header_title.pack(anchor='w')
        self._header_sub = tk.Label(
            title_frame,
            text='填充 A/B 列、匹配 S/T 列，并生成 A1-A10 凭证',
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
        companies = tuple(company for company, _ in COMPANY_OPTIONS if self._company_vars[company].get())
        vouchers = tuple(voucher for voucher in VOUCHER_DISPLAY_ORDER if self._voucher_vars[voucher].get())
        return RunOptions(companies=companies, vouchers=vouchers)

    def _set_voucher_selection(self, voucher_ids):
        selected = set(voucher_ids)
        for voucher, var in self._voucher_vars.items():
            var.set(voucher in selected)
        self._show_guide()

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

        company_row = tk.Frame(card, bg=COLOR_CARD)
        company_row.pack(fill='x', pady=(0, 8))
        tk.Label(company_row, text='公司：', font=('微软雅黑', 10, 'bold'), fg=COLOR_TEXT_MAIN, bg=COLOR_CARD).pack(side='left')
        for company, label in COMPANY_OPTIONS:
            tk.Checkbutton(
                company_row,
                text=label,
                variable=self._company_vars[company],
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
                bg=COLOR_CARD,
                fg=COLOR_TEXT_MAIN,
                selectcolor=COLOR_BG,
                activebackground=COLOR_CARD,
                activeforeground=COLOR_TEXT_MAIN,
                font=('微软雅黑', 9),
                anchor='w',
                width=24,
            ).grid(row=idx // 2, column=idx % 2, sticky='w', padx=(0, 18), pady=3)

    def _collect_missing_inputs(self, run_options, raw_files, payroll_period):
        missing = []
        if not run_options.companies:
            missing.append('至少选择 1 家公司')
        if not run_options.vouchers:
            missing.append('至少选择 1 张凭证')
        if len(raw_files) != 1:
            missing.append('工资单目录下必须且只能有 1 份原始工资单')
        if not os.path.exists(self._mapping_path):
            missing.append('缺少 Mapping表.xlsx')
        if not os.path.exists(self._timesheet_path):
            missing.append('缺少 工时数据.xlsx')
        if requires_bonus_data(run_options) and not os.path.exists(self._bonus_path):
            missing.append('A7/A8 需要 年终奖计提文件')
        if requires_bank_data(run_options) and not _find_bank_files(self._bank_dir):
            missing.append('A1-A3 需要 银行流水 文件')
        if run_options.wants_voucher('A9') and '耐数电子' not in run_options.companies:
            missing.append('A9 仅适用于耐数电子，请勾选 2050')
        if run_options.wants_voucher('A9') and payroll_period:
            co_path = _get_co_workorder_path(self._base_dir, payroll_period[0], payroll_period[1])
            if not os.path.exists(co_path):
                missing.append('A9 需要 CO工单分摊.xlsx')
        if requires_shared_expense_data(run_options) and payroll_period:
            shared_path = _get_shared_expense_path(self._base_dir, payroll_period[0], payroll_period[1])
            if not os.path.exists(shared_path):
                missing.append('A10 需要 待分摊费用YYMM.xlsx')
        return missing

    def _show_guide(self):
        self._phase = 'GUIDE'
        self._clear_body()
        self._clear_footer()
        self._set_header('工资单首步处理工具', '填充 A/B 列、匹配 S/T 列，并按所选公司和 A1-A10 凭证生成输出')

        content = tk.Frame(self._body, bg=COLOR_BG, padx=30, pady=14)
        content.pack(fill='both', expand=True)

        self._build_selection_section(content)

        run_options = self._get_run_options()
        raw_files = _find_raw_files(self._raw_dir)
        payroll_period = _extract_payroll_period(raw_files[0]) if len(raw_files) == 1 else None
        co_path = _get_co_workorder_path(self._base_dir, payroll_period[0], payroll_period[1]) if payroll_period else ''
        shared_path = _get_shared_expense_path(self._base_dir, payroll_period[0], payroll_period[1]) if payroll_period else ''
        mapping_exists = os.path.exists(self._mapping_path)
        timesheet_exists = os.path.exists(self._timesheet_path)
        bonus_exists = os.path.exists(self._bonus_path)
        bank_exists = len(_find_bank_files(self._bank_dir)) > 0
        co_exists = bool(co_path) and os.path.exists(co_path)
        shared_exists = bool(shared_path) and os.path.exists(shared_path)

        tk.Label(
            content,
            text='所需文件清单',
            font=('微软雅黑', 11, 'bold'),
            fg=COLOR_PRIMARY,
            bg=COLOR_BG,
            anchor='w',
        ).pack(fill='x')
        tk.Frame(content, bg=COLOR_PRIMARY, height=2).pack(fill='x', pady=(2, 10))
        file_rows = [
            (
                '必须',
                COLOR_DANGER,
                '#3A1A1A',
                os.path.join('原始数据', '工资单', RAW_FILE_PATTERN),
                '✓ 已找到' if len(raw_files) == 1 else ('✗ 未找到' if not raw_files else '✗ 找到多个'),
                '#000000' if len(raw_files) == 1 else '#FFFFFF',
                COLOR_SUCCESS if len(raw_files) == 1 else COLOR_DANGER,
                '每月 HR 提供的原始工资单，后续会校验金额并匹配 S/T 列',
            ),
            (
                '必须',
                COLOR_DANGER,
                '#3A1A1A',
                'Mapping表.xlsx',
                '✓ 已找到' if mapping_exists else '✗ 未找到',
                '#000000' if mapping_exists else '#FFFFFF',
                COLOR_SUCCESS if mapping_exists else COLOR_DANGER,
                '部门映射表，用于生成成本中心',
            ),
            (
                '必须',
                COLOR_DANGER,
                '#3A1A1A',
                os.path.join('原始数据', '工时数据', '工时数据.xlsx'),
                '✓ 已找到' if timesheet_exists else '✗ 未找到',
                '#000000' if timesheet_exists else '#FFFFFF',
                COLOR_SUCCESS if timesheet_exists else COLOR_DANGER,
                '内部订单来源表，运行前会先确保所有项目都能成功匹配',
            ),
            (
                '必须' if requires_bonus_data(run_options) else '按需',
                COLOR_DANGER if requires_bonus_data(run_options) else '#000000',
                '#3A1A1A' if requires_bonus_data(run_options) else COLOR_WARNING,
                os.path.join('原始数据', '奖金数据', '年终奖计提2026_ - to财务.xlsx'),
                '✓ 已找到' if bonus_exists else ('✗ 未找到' if requires_bonus_data(run_options) else '· 未提供'),
                '#000000' if bonus_exists else '#FFFFFF',
                COLOR_SUCCESS if bonus_exists else (COLOR_DANGER if requires_bonus_data(run_options) else '#666666'),
                'A7/A8 年终奖凭证使用',
            ),
            (
                '必须' if requires_bank_data(run_options) else '按需',
                COLOR_DANGER if requires_bank_data(run_options) else '#000000',
                '#3A1A1A' if requires_bank_data(run_options) else COLOR_WARNING,
                os.path.join('原始数据', '银行流水', BANK_FILE_PATTERN),
                '✓ 已找到' if bank_exists else ('✗ 未找到' if requires_bank_data(run_options) else '· 未提供'),
                '#000000' if bank_exists else '#FFFFFF',
                COLOR_SUCCESS if bank_exists else (COLOR_DANGER if requires_bank_data(run_options) else '#666666'),
                'A1-A3 发放凭证及银行核对使用',
            ),
            (
                '必须' if run_options.wants_voucher('A9') else '按需',
                COLOR_DANGER if run_options.wants_voucher('A9') else '#000000',
                '#3A1A1A' if run_options.wants_voucher('A9') else COLOR_WARNING,
                '耐数电子/处理月份/CO工单分摊/CO工单分摊.xlsx',
                '✓ 已找到' if co_exists else ('✗ 未找到' if run_options.wants_voucher('A9') else '· 未提供'),
                '#000000' if co_exists else '#FFFFFF',
                COLOR_SUCCESS if co_exists else (COLOR_DANGER if run_options.wants_voucher('A9') else '#666666'),
                'A9 仅耐数电子使用',
            ),
            (
                '必须' if requires_shared_expense_data(run_options) else '按需',
                COLOR_DANGER if requires_shared_expense_data(run_options) else '#000000',
                '#3A1A1A' if requires_shared_expense_data(run_options) else COLOR_WARNING,
                os.path.join('原始数据', '待分摊费用', '待分摊费用YYMM.xlsx'),
                '✓ 已找到' if shared_exists else ('✗ 未找到' if requires_shared_expense_data(run_options) else '· 未提供'),
                '#000000' if shared_exists else '#FFFFFF',
                COLOR_SUCCESS if shared_exists else (COLOR_DANGER if requires_shared_expense_data(run_options) else '#666666'),
                'A10 研发费用分摊使用',
            ),
        ]
        for row_data in file_rows:
            self._build_file_status_row(content, *row_data)

        tk.Label(
            content,
            text=f'工资单目录：{self._raw_dir}',
            font=('微软雅黑', 9),
            fg=COLOR_TEXT_SUB,
            bg=COLOR_BG,
            anchor='w',
        ).pack(fill='x', pady=(6, 0))

        if raw_files:
            show_name = os.path.basename(raw_files[0]) if len(raw_files) == 1 else '；'.join(os.path.basename(p) for p in raw_files[:2])
            tk.Label(
                content,
                text=f'当前识别文件：{show_name}',
                font=('微软雅黑', 9),
                fg=COLOR_TEXT_SUB,
                bg=COLOR_BG,
                anchor='w',
            ).pack(fill='x', pady=(2, 0))

        tk.Label(
            content,
            text=f'Mapping 表：{self._mapping_path}',
            font=('微软雅黑', 9),
            fg=COLOR_TEXT_SUB,
            bg=COLOR_BG,
            anchor='w',
        ).pack(fill='x', pady=(2, 0))
        tk.Label(
            content,
            text=f'银行流水目录：{self._bank_dir}',
            font=('微软雅黑', 9),
            fg=COLOR_TEXT_SUB,
            bg=COLOR_BG,
            anchor='w',
        ).pack(fill='x', pady=(2, 0))
        tk.Label(
            content,
            text=f'工时数据：{self._timesheet_path}',
            font=('微软雅黑', 9),
            fg=COLOR_TEXT_SUB,
            bg=COLOR_BG,
            anchor='w',
        ).pack(fill='x', pady=(2, 0))
        tk.Label(
            content,
            text=f'奖金数据：{self._bonus_path}',
            font=('微软雅黑', 9),
            fg=COLOR_TEXT_SUB,
            bg=COLOR_BG,
            anchor='w',
        ).pack(fill='x', pady=(2, 0))
        if co_path:
            tk.Label(
                content,
                text=f'CO工单分摊：{co_path}',
                font=('微软雅黑', 9),
                fg=COLOR_TEXT_SUB,
                bg=COLOR_BG,
                anchor='w',
            ).pack(fill='x', pady=(2, 0))

        tk.Label(content, text='', bg=COLOR_BG, height=1).pack()
        tk.Label(
            content,
            text='运行步骤',
            font=('微软雅黑', 11, 'bold'),
            fg=COLOR_PRIMARY,
            bg=COLOR_BG,
            anchor='w',
        ).pack(fill='x')
        tk.Frame(content, bg=COLOR_PRIMARY, height=2).pack(fill='x', pady=(2, 10))

        for num, text in [
            ('①', '读取 原始数据/工资单 下唯一一份原始工资单'),
            ('②', '对首个工作表 A/B 列空白单元格按上一行非空值向下填充'),
            ('③', '校验 Q 列实发金额是否等于 E 列减 K:P 列个人承担金额'),
            ('④', '根据 Mapping 表生成 S 列成本中心，并根据工时数据生成 T 列内部订单'),
            ('⑤', '若工时项目未全部匹配到内部订单，则终止后续处理'),
            ('⑥', '读取银行流水、工时、奖金和可选的 CO工单分摊数据，生成 A1-A9 凭证，如有异常则标红输出'),
        ]:
            row = tk.Frame(content, bg=COLOR_BG, pady=3)
            row.pack(fill='x')
            _circle_label(row, num, size=26, parent_bg=COLOR_BG).pack(side='left', padx=(0, 10))
            tk.Label(row, text=text, font=('微软雅黑', 10), fg=COLOR_TEXT_MAIN, bg=COLOR_BG, anchor='w').pack(side='left')

        tip = tk.Frame(content, bg='#252500', pady=8, padx=12)
        tip.pack(fill='x', pady=(14, 0))
        tk.Label(
            tip,
            text='输出文件会按公司拆分，保存到对应公司目录下的处理月份文件夹，并统一命名为“3月工资单-整理后.xlsx”这类格式；同名文件会直接替换',
            font=('微软雅黑', 9),
            fg=COLOR_PRIMARY,
            bg='#252500',
            anchor='w',
        ).pack(fill='x')

        missing_items = self._collect_missing_inputs(run_options, raw_files, payroll_period)
        can_run = not missing_items
        if can_run:
            _btn(self._footer, '  文件已准备好，开始运行  ', COLOR_SUCCESS, self._start_run).pack(
                side='left', padx=(30, 10)
            )
        else:
            tk.Label(
                self._footer,
                text='当前还不能运行：' + '；'.join(missing_items[:3]),
                font=('微软雅黑', 9, 'bold'),
                fg=COLOR_DANGER,
                bg=COLOR_BG,
            ).pack(side='left', padx=(20, 0))
        _btn(self._footer, '取消', '#444444', self.root.destroy, padx=16, pady=6).pack(side='left', padx=(10, 0))

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
                f"首个工作表：{self._result['sheet_name']}\n"
                f"总行数：{self._result['row_count']}\n"
                f"A 列填充：{self._result['fill_a']} 个\n"
                f"B 列填充：{self._result['fill_b']} 个\n"
                f"实发校验异常：{self._result['validation_issue_count']} 行\n"
                f"成本中心未匹配：{self._result['cost_center_issue_count']} 行\n"
                f"内部订单未匹配：{self._result['internal_order_issue_count']} 行\n"
                f"银行核对异常项：{self._result['bank_issue_count']} 项\n"
                f"工资/公积金/个税/社保异常项："
                f"{self._result['bank_salary_issue_count']}/"
                f"{self._result['bank_fund_issue_count']}/"
                f"{self._result['bank_tax_issue_count']}/"
                f"{self._result['bank_social_issue_count']}\n"
                f"行级异常总数：{self._result['issue_count']} 行\n"
                f"凭证校验：\n" + ('\n'.join(voucher_validation_lines) if voucher_validation_lines else '未生成凭证') + '\n'
                f"输出工资文件：\n" + ('\n'.join(output_lines) if output_lines else '未输出工资文件') + '\n'
                f"输出凭证文件：\n" + ('\n'.join(voucher_lines) if voucher_lines else '未输出凭证文件')
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
        raw_files = _find_raw_files(self._raw_dir)
        payroll_period = _extract_payroll_period(raw_files[0]) if len(raw_files) == 1 else None
        missing_items = self._collect_missing_inputs(run_options, raw_files, payroll_period)
        if missing_items:
            messagebox.showerror('无法运行', '；'.join(missing_items))
            return

        input_path = raw_files[0]
        payroll_year, payroll_month = _extract_payroll_period(input_path)
        co_path = _get_co_workorder_path(self._base_dir, payroll_year, payroll_month)
        shared_path = _get_shared_expense_path(self._base_dir, payroll_year, payroll_month)
        self._show_progress()

        def log(text, tag=''):
            self._q.put(('log', text, tag))

        def run():
            try:
                self._q.put(('step', 1, '检查输入文件'))
                log(f'输入文件：{os.path.basename(input_path)}')
                log('本次公司：' + '、'.join(run_options.companies))
                log('本次凭证：' + '、'.join(run_options.vouchers))
                log('输出位置：按公司拆分到对应公司/月份目录，同名文件直接替换')
                log(f'Mapping 表：{os.path.basename(self._mapping_path)}')
                log(f'工时数据：{self._timesheet_path}')
                if requires_bonus_data(run_options):
                    log(f'奖金数据：{self._bonus_path}')
                if run_options.wants_voucher('A9'):
                    log(f'CO工单分摊：{co_path}')
                if run_options.wants_voucher('A10'):
                    log(f'待分摊费用：{shared_path}')
                if requires_bank_data(run_options):
                    log(f'银行流水目录：{self._bank_dir}')

                self._q.put(('step', 2, '填充首个工作表 A/B 列空白'))
                self._q.put(('step', 3, '校验实发金额并匹配成本中心/内部订单'))
                self._q.put(('step', 4, '生成所选公司和凭证'))
                result = execute_payroll_run(
                    input_path,
                    self._base_dir,
                    self._mapping_path,
                    self._bank_dir,
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
