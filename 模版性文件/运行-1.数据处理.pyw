import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from dateutil.relativedelta import relativedelta
import threading
import queue
import traceback as _tb

# ============================================================
# 常量定义
# ============================================================

ACCOUNTING_SUBJECT_MAPPING = {
    66030101: '财务费用-利息收入',
    66030102: '财务费用-利息支出',
}
FUNCTION_AREA_MAPPING = {1000: '销售费用', 2000: '管理费用', 3000: '研发费用', 4000: '制造费用'}
DOMESTIC_COMPANIES = {'1000', '2000', '3000', '6000', '7000'}
FX_LETTER_MAP = {
    'US10': 'C', 'EU10': 'D', 'JP10': 'E',
    'HK10': 'F', 'SG10': 'N', 'KR10': 'X', 'MY10': 'AA',
}

# ============================================================
# 配色常量（RIGOL 品牌风格：黄黑白）
# ============================================================

COLOR_PRIMARY   = '#FFD700'   # RIGOL 黄
COLOR_SUCCESS   = '#FFD700'   # 成功/通过
COLOR_DANGER    = '#FF4444'   # 错误/危险
COLOR_WARNING   = '#FFA500'   # 警告
COLOR_BG        = '#1A1A1A'   # 主背景深黑
COLOR_CARD      = '#2A2A2A'   # 卡片/容器深灰
COLOR_BORDER    = '#3D3D3D'   # 边框
COLOR_TEXT_MAIN = '#FFFFFF'   # 主文字白色
COLOR_TEXT_SUB  = '#AAAAAA'   # 副文字浅灰
COLOR_HEADER_BG = '#111111'   # 顶部标题栏近黑

# ============================================================
# 工具函数
# ============================================================

def excel_col_to_int(col_str: str) -> int:
    res = 0
    for ch in col_str.upper():
        res = res * 26 + (ord(ch) - ord('A') + 1)
    return res - 1


def _apply_style(root):
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('TFrame', background=COLOR_BG)
    style.configure('Treeview',
        background=COLOR_CARD, fieldbackground=COLOR_CARD,
        foreground=COLOR_TEXT_MAIN, rowheight=28, font=('微软雅黑', 10), borderwidth=0)
    style.configure('Treeview.Heading',
        background='#333333', foreground=COLOR_PRIMARY,
        font=('微软雅黑', 10, 'bold'), relief='flat')
    style.map('Treeview', background=[('selected', '#3A3A1A')])
    style.configure('TNotebook', background=COLOR_BG, borderwidth=0)
    style.configure('TNotebook.Tab',
        background='#2A2A2A', foreground=COLOR_TEXT_SUB,
        font=('微软雅黑', 10), padding=(12, 6))
    style.map('TNotebook.Tab',
        background=[('selected', '#333300')],
        foreground=[('selected', COLOR_PRIMARY)])
    style.configure('Horizontal.TProgressbar',
        background=COLOR_PRIMARY, troughcolor='#333333', borderwidth=0, thickness=8)


def _btn(parent, text, bg, command, padx=20, pady=8):
    # 成功/主操作用黄底黑字，危险/取消用相应深色底白字
    if bg == COLOR_PRIMARY or bg == COLOR_SUCCESS:
        fg_color = '#000000'
        active_bg = '#E6C200'
    elif bg == COLOR_DANGER:
        fg_color = '#FFFFFF'
        active_bg = '#CC2222'
    else:
        fg_color = '#FFFFFF'
        active_bg = '#555555'
    return tk.Button(parent, text=text, bg=bg, fg=fg_color,
                     font=('微软雅黑', 11, 'bold'), relief='flat',
                     padx=padx, pady=pady, cursor='hand2',
                     activebackground=active_bg,
                     activeforeground=fg_color, command=command)


def _check_file(base_dir: str, pattern: str) -> bool:
    """检查文件/文件夹是否存在（支持 glob 通配符）"""
    import glob as _glob
    if '*' in pattern:
        return bool(_glob.glob(os.path.join(base_dir, pattern)))
    return os.path.exists(os.path.join(base_dir, pattern))


def _circle_label(parent, text: str, size: int = 26,
                  bg_color: str = None, fg_color: str = '#000000',
                  parent_bg: str = None) -> tk.Canvas:
    """用 Canvas 绘制真圆形徽标，返回 Canvas 供 pack/grid 使用"""
    if bg_color is None:
        bg_color = COLOR_PRIMARY
    if parent_bg is None:
        parent_bg = COLOR_BG
    c = tk.Canvas(parent, width=size, height=size,
                  bg=parent_bg, highlightthickness=0)
    pad = 1
    c.create_oval(pad, pad, size - pad, size - pad,
                  fill=bg_color, outline='')
    c.create_text(size // 2, size // 2, text=text,
                  font=('微软雅黑', size // 3, 'bold'),
                  fill=fg_color, anchor='center')
    return c


# ============================================================
# 数据读取与合并
# ============================================================

def load_raw_data(base_dir: str, log) -> pd.DataFrame:
    file_path    = os.path.join(base_dir, '原始表.xlsx')
    file_path_my = os.path.join(base_dir, '原始表-马来.xlsx')
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到原始表：{file_path}")
    log("正在读取原始表…")
    df = pd.read_excel(file_path)
    log(f"原始表读取完成，共 {len(df)} 行", 'ok')
    if os.path.exists(file_path_my):
        log("正在读取原始表-马来…")
        df_my = pd.read_excel(file_path_my)
        df = pd.concat([df, df_my], ignore_index=True)
        log(f"合并完成，共 {len(df)} 行", 'ok')
    else:
        log("未找到原始表-马来.xlsx，跳过")
    return df


def load_fx_lookup(base_dir: str, df: pd.DataFrame, date_col: str, log) -> tuple[dict, list]:
    fx_list_path = os.path.join(base_dir, 'FX rate list.xlsx')
    current_year = str(datetime.now().year)
    log("正在读取 FX rate list…")
    fx_xl = pd.ExcelFile(fx_list_path)
    fx_data_raw = None
    for sheet in fx_xl.sheet_names:
        tmp = pd.read_excel(fx_xl, sheet_name=sheet, header=None)
        if tmp.astype(str).apply(lambda x: x.str.contains(current_year)).any().any():
            fx_data_raw = tmp
            break
    if fx_data_raw is None:
        fx_data_raw = pd.read_excel(fx_xl, sheet_name=0, header=None)

    start_row = 0
    for i in range(len(fx_data_raw)):
        try:
            if 2000 < int(float(fx_data_raw.iloc[i, 0])) < 2100:
                start_row = i
                break
        except Exception:
            continue

    fx_clean = fx_data_raw.iloc[start_row:].copy()
    fx_clean.iloc[:, 0] = pd.to_numeric(fx_clean.iloc[:, 0], errors='coerce')
    fx_clean.iloc[:, 1] = pd.to_numeric(fx_clean.iloc[:, 1], errors='coerce')
    fx_clean = fx_clean.dropna(subset=[fx_clean.columns[0], fx_clean.columns[1]])

    unique_periods = df[date_col].dt.to_period('M').unique()
    fx_summary, fx_lookup_table = [], {}
    for period in sorted(unique_periods):
        target_date = period.to_timestamp() + relativedelta(months=1)
        t_year, t_month = target_date.year, target_date.month
        match = fx_clean[
            (fx_clean.iloc[:, 0].astype(int) == t_year) &
            (fx_clean.iloc[:, 1].astype(int) == t_month)
        ]
        status = "正常"
        if match.empty:
            match = fx_clean.sort_values(
                [fx_clean.columns[0], fx_clean.columns[1]], ascending=False).head(1)
            status = "取最新"
        row_res = {
            '原始账期': str(period),
            '汇率选取年月': f"{int(float(match.iloc[0,0]))}年{int(float(match.iloc[0,1]))}月",
            '状态': status,
        }
        for co, letter in FX_LETTER_MAP.items():
            idx = excel_col_to_int(letter)
            rate = match.iloc[0, idx] if idx < fx_data_raw.shape[1] else 1.0
            row_res[co] = f"{rate:.4f}" if isinstance(rate, (int, float)) else str(rate)
            fx_lookup_table[(period.year, period.month, co)] = rate
        fx_summary.append(row_res)
    log(f"汇率匹配完成，共 {len(fx_summary)} 个账期", 'ok')
    return fx_lookup_table, fx_summary


def clean_and_enrich(df: pd.DataFrame, date_col: str, fx_lookup_table: dict, log) -> pd.DataFrame:
    log("正在清洗数据，填充费用类型/汇率/CNY…")
    def try_num(s): return pd.to_numeric(s, errors='coerce').fillna(s)
    df['成本中心'] = try_num(df['成本中心'])
    df['会计科目'] = try_num(df['会计科目'])
    mask_kj = df['公司代码'].isin(['KR10', 'JP10'])
    df.loc[mask_kj, '本位币金额'] = df.loc[mask_kj, '本位币金额'] * 100
    df = df[~df['订单号'].astype(str).str.startswith('93')]
    v_col = '凭 证编 号' if '凭 证编 号' in df.columns else '凭证编号'
    df = df[~(
        (df['公司代码'] == 'MY10') &
        (df[v_col].astype(str).str.startswith('9')) &
        (df['会计科目'].isna())
    )]
    if '费用类型' not in df.columns:
        df.insert(df.columns.get_loc('公司代码') + 1, '费用类型', 'NA')
    is_6603 = df['会计科目'].astype(str).str.startswith('6603')
    df.loc[is_6603, '费用类型'] = '财务费用'
    function_area = pd.to_numeric(df['功能范围'], errors='coerce')
    mask_2050_2060 = df['公司代码'].astype(str).isin(['2050', '2060'])
    order_col = '订单号' if '订单号' in df.columns else None
    cond_zz    = mask_2050_2060 & (function_area == 5000)
    cond_ht    = mask_2050_2060 & (function_area == 6000)
    cond_ht_91 = mask_2050_2060 & (order_col is not None) & df[order_col].astype(str).str.startswith('91')
    df.loc[cond_zz    & ~is_6603, '费用类型'] = '制造费用'
    df.loc[cond_ht    & ~is_6603, '费用类型'] = '合同履约成本'
    df.loc[cond_ht_91 & ~is_6603, '费用类型'] = '合同履约成本'
    mask_na = df['费用类型'] == 'NA'
    df.loc[mask_na, '费用类型'] = function_area[mask_na].map(FUNCTION_AREA_MAPPING).fillna('NA')
    if '集团货币CNY' not in df.columns:
        df.insert(df.columns.get_loc('本位币金额') + 1, '集团货币CNY', 0.0)
    if '汇率' not in df.columns:
        df['汇率'] = 1.0
    co_series = df['公司代码'].astype(str)
    dt_series = df[date_col]
    domestic_mask = co_series.isin(DOMESTIC_COMPANIES)
    df.loc[domestic_mask, '汇率'] = 1.0
    df.loc[domestic_mask, '集团货币CNY'] = df.loc[domestic_mask, '本位币金额']
    foreign_mask = ~domestic_mask
    if foreign_mask.any():
        keys = list(zip(dt_series[foreign_mask].dt.year,
                        dt_series[foreign_mask].dt.month,
                        co_series[foreign_mask]))
        rates = pd.Series([fx_lookup_table.get(k, 1.0) for k in keys],
                          index=df.index[foreign_mask], dtype=float)
        df.loc[foreign_mask, '汇率'] = rates
        df.loc[foreign_mask, '集团货币CNY'] = df.loc[foreign_mask, '本位币金额'] * rates
    mask_subj = df['会计科目'].notna() & (df['会计科目文本'].isna() | (df['会计科目文本'] == ''))
    df.loc[mask_subj, '会计科目文本'] = (
        pd.to_numeric(df['会计科目'], errors='coerce').map(ACCOUNTING_SUBJECT_MAPPING))
    for col in df.columns:
        if '日期' in str(col) or pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = pd.to_datetime(df[col], errors='coerce')
    return df


def write_output(df: pd.DataFrame, output_path: str, log) -> None:
    log("正在写出 费用表.xlsx…")
    unmatched = (df[df['费用类型'] == 'NA'][['功能范围', '成本中心', '成本中心名称']].drop_duplicates())
    with pd.ExcelWriter(output_path, engine='openpyxl', datetime_format='yyyy/mm/dd') as writer:
        df.to_excel(writer, sheet_name='原始数据', index=False)
        unmatched.to_excel(writer, sheet_name='待匹配费用类型', index=False)
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len, column_letter = 0, col[0].column_letter
                for cell in col:
                    try:
                        val = str(cell.value) if cell.value is not None else ''
                        w = sum(2 if ord(c) > 256 else 1 for c in val)
                        if w > max_len: max_len = w
                    except Exception:
                        pass
                ws.column_dimensions[column_letter].width = min(max_len + 4, 60)
    log(f"文件已保存：{output_path}", 'ok')


# ============================================================
# 主窗口（含使用说明 + 进度 + 内联确认，全程只有这一个窗口）
# ============================================================

class MainWindow:
    """
    整个脚本只有这一个窗口，分三个阶段：
      1. GUIDE      — 使用说明，用户点击"开始运行"
      2. RUN        — 进度条 + 日志，子线程在跑
      3. FX_CONFIRM — 汇率确认表格（内联，替换日志区）
    """

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("费用表生成工具")
        self.root.configure(bg=COLOR_BG)
        self.root.resizable(True, True)
        self.root.minsize(680, 520)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        _apply_style(self.root)
        self.root.update_idletasks()

        self._phase = 'GUIDE'
        self._q: queue.Queue = queue.Queue()
        self._fx_event = threading.Event()
        self._fx_result = [False]
        self._df_ref = [None]
        self._fx_summary_ref = [None]
        self._fx_lookup_ref = [None]
        self._base_dir = os.path.dirname(os.path.abspath(__file__))

        # 加载 RIGOL logo
        self._logo_img = None
        try:
            from PIL import Image, ImageTk
            logo_path = os.path.join(self._base_dir, 'rigol_logo.png')
            if os.path.exists(logo_path):
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

    # ── 关闭保护 ──────────────────────────────────────────────
    def _on_close(self):
        if self._phase == 'RUN':
            return
        self.root.destroy()

    # ── 顶部标题栏 ────────────────────────────────────────────
    def _build_header(self):
        self._header_frame = tk.Frame(self.root, bg=COLOR_HEADER_BG, pady=0)
        self._header_frame.pack(fill='x')
        # 黄色顶部细线
        tk.Frame(self._header_frame, bg=COLOR_PRIMARY, height=3).pack(fill='x')
        inner = tk.Frame(self._header_frame, bg=COLOR_HEADER_BG, pady=10)
        inner.pack(fill='x')
        self._header_inner = inner
        # logo 放左侧
        if self._logo_img:
            lbl = tk.Label(inner, image=self._logo_img, bg=COLOR_HEADER_BG)
            lbl.pack(side='left', padx=(14, 12))
        # 标题文字左侧
        title_frame = tk.Frame(inner, bg=COLOR_HEADER_BG)
        title_frame.pack(side='left', fill='both', expand=True)
        self._header_title = tk.Label(title_frame, text="费用表生成工具",
            font=('微软雅黑', 15, 'bold'), fg=COLOR_PRIMARY, bg=COLOR_HEADER_BG)
        self._header_title.pack(anchor='w')
        self._header_sub = tk.Label(title_frame,
            text="运行前请确认以下文件已准备好，并放置在脚本同目录下",
            font=('微软雅黑', 9), fg=COLOR_TEXT_SUB, bg=COLOR_HEADER_BG)
        self._header_sub.pack(anchor='w', pady=(2, 0))
        # 署名放右侧
        tk.Label(inner, text="Developed by Earnest Yin",
                 font=('Consolas', 10, 'bold'), fg='#888888', bg=COLOR_HEADER_BG
                 ).pack(side='right', padx=(0, 16), anchor='s')

    def _set_header(self, title: str, sub: str = '', bg: str = COLOR_HEADER_BG):
        # bg参数用于成功/失败时改变顶部细线颜色，标题栏背景始终保持深黑
        self._header_frame.configure(bg=COLOR_HEADER_BG)
        self._header_inner.configure(bg=COLOR_HEADER_BG)
        # 更新顶部细线颜色
        for w in self._header_frame.winfo_children():
            if isinstance(w, tk.Frame) and w is not self._header_inner:
                w.configure(bg=bg if bg != COLOR_HEADER_BG else COLOR_PRIMARY)
                break
        self._header_title.configure(text=title, bg=COLOR_HEADER_BG,
            fg=COLOR_PRIMARY if bg == COLOR_HEADER_BG else bg)
        self._header_sub.configure(text=sub, bg=COLOR_HEADER_BG)

    def _clear_body(self):
        for w in self._body.winfo_children():
            if w is not self._footer:
                w.destroy()

    def _clear_footer(self):
        for w in self._footer.winfo_children():
            w.destroy()

    # ── 阶段1：使用说明（含文件完备性检测） ──────────────────
    def _show_guide(self):
        self._phase = 'GUIDE'
        self._clear_body()
        self._clear_footer()
        self._set_header("费用表生成工具", "运行前请确认以下文件已准备好，并放置在脚本同目录下")

        content = tk.Frame(self._body, bg=COLOR_BG, padx=30, pady=14)
        content.pack(fill='both', expand=True)

        tk.Label(content, text="所需文件清单",
            font=('微软雅黑', 11, 'bold'), fg=COLOR_PRIMARY, bg=COLOR_BG, anchor='w').pack(fill='x')
        tk.Frame(content, bg=COLOR_PRIMARY, height=2).pack(fill='x', pady=(2, 10))

        # 文件列表：(标签, 文件名/通配符, 描述, 是否必须)
        file_items = [
            ("必须", "原始表.xlsx",       "从 SAP 导出的费用明细原始数据",         True),
            ("可选", "原始表-马来.xlsx",   "马来西亚公司单独导出的数据（无则跳过）", False),
            ("必须", "FX rate list.xlsx", "包含各外币月度汇率的汇率表",             True),
        ]
        self._must_ok = []
        for tag, fname, desc, required in file_items:
            exists = _check_file(self._base_dir, fname)
            row = tk.Frame(content, bg=COLOR_BG, pady=4)
            row.pack(fill='x')
            tc = COLOR_DANGER if tag == '必须' else COLOR_WARNING
            tb = '#3A1A1A'    if tag == '必须' else '#2A1A00'
            tk.Label(row, text=f" {tag} ", font=('微软雅黑', 8, 'bold'),
                     fg=tc, bg=tb, padx=4).pack(side='left', padx=(0, 8))
            tk.Label(row, text=fname, font=('微软雅黑', 10, 'bold'),
                     fg=COLOR_TEXT_MAIN, bg=COLOR_BG, width=22, anchor='w').pack(side='left')
            # 状态徽标
            if exists:
                status_text, status_fg, status_bg = "✓ 已找到", '#000000', COLOR_SUCCESS
            elif required:
                status_text, status_fg, status_bg = "✗ 未找到", '#FFFFFF', COLOR_DANGER
            else:
                status_text, status_fg, status_bg = "— 未找到", COLOR_TEXT_SUB, '#333333'
            tk.Label(row, text=f" {status_text} ", font=('微软雅黑', 8, 'bold'),
                     fg=status_fg, bg=status_bg, padx=4).pack(side='left', padx=(0, 10))
            tk.Label(row, text=desc, font=('微软雅黑', 9),
                     fg=COLOR_TEXT_SUB, bg=COLOR_BG, anchor='w').pack(side='left')
            if required:
                self._must_ok.append(exists)

        tk.Label(content, text="", bg=COLOR_BG, height=1).pack()
        tk.Label(content, text="运行步骤",
            font=('微软雅黑', 11, 'bold'), fg=COLOR_PRIMARY, bg=COLOR_BG, anchor='w').pack(fill='x')
        tk.Frame(content, bg=COLOR_PRIMARY, height=2).pack(fill='x', pady=(2, 10))

        for num, text in [
            ("①", "读取原始表，自动合并马来数据"),
            ("②", "从 FX rate list 匹配汇率 → 在本窗口确认"),
            ("③", "清洗数据，自动填充费用类型、汇率及集团货币CNY"),
            ("④", "生成 费用表.xlsx（含「待匹配费用类型」诊断页）"),
        ]:
            row = tk.Frame(content, bg=COLOR_BG, pady=3)
            row.pack(fill='x')
            _circle_label(row, num, size=26, parent_bg=COLOR_BG).pack(side='left', padx=(0, 10))
            tk.Label(row, text=text, font=('微软雅黑', 10),
                     fg=COLOR_TEXT_MAIN, bg=COLOR_BG, anchor='w').pack(side='left')

        tip = tk.Frame(content, bg='#252500', pady=8, padx=12)
        tip.pack(fill='x', pady=(14, 0))
        tk.Label(tip, text="完成后请运行【运行-2.数据汇总.pyw】将数据导入费用明细表",
                 font=('微软雅黑', 9), fg=COLOR_PRIMARY, bg='#252500', anchor='w').pack(fill='x')

        all_must_ready = all(self._must_ok)
        if all_must_ready:
            _btn(self._footer, "  文件已准备好，开始运行  ", COLOR_SUCCESS, self._start_run
                 ).pack(side='left', padx=(30, 10))
        else:
            tk.Label(self._footer,
                     text="⚠  请先将标记为「✗ 未找到」的必须文件放入脚本同目录，再重新打开程序",
                     font=('微软雅黑', 9, 'bold'), fg=COLOR_DANGER, bg=COLOR_BG
                     ).pack(side='left', padx=(20, 0))
        _btn(self._footer, "取消", '#444444', self.root.destroy, padx=16, pady=6
             ).pack(side='left', padx=(10, 0))

    # ── 阶段2：进度界面 ───────────────────────────────────────
    def _show_progress(self):
        self._phase = 'RUN'
        self._clear_body()
        self._clear_footer()
        self._set_header("正在运行，请稍候…", "")

        body = tk.Frame(self._body, bg=COLOR_BG, padx=24, pady=14)
        body.pack(fill='both', expand=True)

        self._step_var = tk.StringVar(value="准备中…")
        tk.Label(body, textvariable=self._step_var,
                 font=('微软雅黑', 10, 'bold'), fg=COLOR_PRIMARY,
                 bg=COLOR_BG, anchor='w').pack(fill='x')

        self._pb_var = tk.IntVar(value=0)
        self._pb = ttk.Progressbar(body, variable=self._pb_var,
                                   maximum=4, length=650, mode='determinate',
                                   style='Horizontal.TProgressbar')
        self._pb.pack(fill='x', pady=(6, 12))

        tk.Label(body, text="运行日志",
                 font=('微软雅黑', 9, 'bold'), fg=COLOR_TEXT_SUB,
                 bg=COLOR_BG, anchor='w').pack(fill='x')
        lf = tk.Frame(body, bg=COLOR_CARD,
                      highlightbackground=COLOR_BORDER, highlightthickness=1)
        lf.pack(fill='both', expand=True, pady=(4, 0))
        self._log_text = tk.Text(lf, height=12, font=('Consolas', 11),
                                 bg=COLOR_CARD, fg='#E8E8E8',
                                 relief='flat', state='disabled',
                                 wrap='word', padx=10, pady=8,
                                 insertbackground=COLOR_PRIMARY,
                                 selectbackground='#3A3A1A')
        vsb = tk.Scrollbar(lf, command=self._log_text.yview, bg=COLOR_BORDER)
        self._log_text.configure(yscrollcommand=vsb.set)
        self._log_text.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

        self._log_text.tag_configure('ok',   foreground=COLOR_SUCCESS)
        self._log_text.tag_configure('warn', foreground=COLOR_WARNING)
        self._log_text.tag_configure('err',  foreground=COLOR_DANGER)

        self._progress_body = body

    def _append_log(self, msg: str, tag: str = 'info'):
        self._log_text.configure(state='normal')
        prefix = {'ok': '✓ ', 'warn': '⚠ ', 'err': '✗ '}.get(tag, '  ')
        self._log_text.insert('end', f"{prefix}{msg}\n", tag if tag != 'info' else '')
        self._log_text.see('end')
        self._log_text.configure(state='disabled')

    def _set_step(self, step: int, label: str):
        self._pb_var.set(step)
        self._step_var.set(f"步骤 {step}/4  {label}")

    # ── 阶段3：汇率确认（替换日志区内容） ───────────────────
    def _show_fx_confirm(self, fx_summary: list):
        self._phase = 'FX_CONFIRM'
        self._set_header("请确认匹配到的汇率", "系统已根据记账日期自动匹配对应月份汇率，确认后继续处理")
        self._step_var.set("步骤 2/4  汇率确认")

        if hasattr(self, '_log_text'):
            self._log_text.master.pack_forget()

        disp = tk.Frame(self._progress_body, bg=COLOR_BG)
        disp.pack(fill='both', expand=True, pady=(4, 0))
        self._fx_disp_frame = disp

        cols = ['原始账期', '汇率选取年月', 'HK10', 'US10', 'EU10', 'JP10', 'SG10', 'KR10', 'MY10', '状态']
        tree = ttk.Treeview(disp, columns=cols, show='headings', height=8)
        col_w = {'原始账期': 90, '汇率选取年月': 110, '状态': 70}
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=col_w.get(c, 78), anchor='center')
        tree.tag_configure('latest', background='#2A2000', foreground=COLOR_WARNING)
        for item in fx_summary:
            tag = 'latest' if item.get('状态') == '取最新' else ''
            tree.insert('', 'end', values=[item.get(c, '') for c in cols], tags=(tag,))
        vsb = ttk.Scrollbar(disp, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

        tk.Label(self._progress_body,
                 text="  ⚠ 橙色行表示未找到精确月份，已自动取最新汇率",
                 font=('微软雅黑', 9), fg=COLOR_WARNING, bg=COLOR_BG, anchor='w'
                 ).pack(fill='x', pady=(4, 0))

        self._clear_footer()
        _btn(self._footer, "  确认汇率，继续处理  ", COLOR_SUCCESS, self._fx_ok
             ).pack(side='left', padx=(30, 10))
        _btn(self._footer, "取消", '#444444', self._fx_cancel, padx=16, pady=6
             ).pack(side='left')

    def _fx_ok(self):
        self._fx_result[0] = True
        self._fx_event.set()
        if hasattr(self, '_fx_disp_frame'):
            self._fx_disp_frame.destroy()
        for w in self._progress_body.winfo_children():
            w.pack_forget()
        for w in self._progress_body.winfo_children():
            w.pack(fill='x')
        self._log_text.master.pack(fill='both', expand=True, pady=(4, 0))
        self._clear_footer()
        self._phase = 'RUN'
        self._set_header("正在运行，请稍候…", "")

    def _fx_cancel(self):
        self._fx_result[0] = False
        self._fx_event.set()

    # ── 队列轮询（主线程安全刷新UI）──────────────────────────
    def _poll(self):
        try:
            while True:
                msg = self._q.get_nowait()
                kind = msg[0]
                if kind == 'log':
                    self._append_log(msg[1], msg[2])
                elif kind == 'step':
                    self._set_step(msg[1], msg[2])
                elif kind == 'fx_confirm':
                    self._show_fx_confirm(msg[1])
                elif kind == 'done':
                    self._on_done(msg[1], msg[2])
                    return
        except queue.Empty:
            pass
        self.root.after(50, self._poll)

    def _on_done(self, success: bool, summary: str):
        self._phase = 'DONE'
        bg = COLOR_SUCCESS if success else COLOR_DANGER
        self._set_header("✓  完成！" if success else "✗  运行失败", summary, bg)
        self._pb_var.set(4)
        self._clear_footer()
        _btn(self._footer, "关闭", bg, self.root.destroy).pack()

    # ── 启动子线程 ────────────────────────────────────────────
    def _start_run(self):
        base_dir = self._base_dir
        output_path = os.path.join(base_dir, '费用表.xlsx')
        self._show_progress()

        def log(msg, tag='info'):
            self._q.put(('log', msg, tag))

        def run():
            try:
                self._q.put(('step', 1, "读取原始数据"))
                df = load_raw_data(base_dir, log)

                date_col = '记账日期'
                if date_col not in df.columns:
                    found = [c for c in df.columns if '记账日期' in str(c).replace(' ', '')]
                    if found: date_col = found[0]
                    else: raise ValueError("未找到'记账日期'列")
                df[date_col] = pd.to_datetime(df[date_col])

                self._q.put(('step', 2, "匹配汇率"))
                fx_lookup_table, fx_summary = load_fx_lookup(base_dir, df, date_col, log)

                self._fx_event.clear()
                self._q.put(('fx_confirm', fx_summary))
                self._fx_event.wait()

                if not self._fx_result[0]:
                    self._q.put(('done', False, "已取消"))
                    return

                self._q.put(('step', 3, "清洗与填充数据"))
                df = clean_and_enrich(df, date_col, fx_lookup_table, log)
                na_count = (df['费用类型'] == 'NA').sum()
                log(f"处理完成，{len(df)} 行，未匹配费用类型 {na_count} 行", 'ok')

                self._q.put(('step', 4, "写出费用表.xlsx"))
                write_output(df, output_path, log)

                self._q.put(('done', True,
                             f"原始数据 {len(df)} 行 | 未匹配费用类型 {na_count} 行"))

            except FileNotFoundError as e:
                log(str(e), 'err')
                self._q.put(('done', False, str(e)))
            except ValueError as e:
                log(str(e), 'err')
                self._q.put(('done', False, str(e)))
            except Exception as e:
                log(f"{e}", 'err')
                self._q.put(('done', False, str(e)))

        threading.Thread(target=run, daemon=True).start()
        self.root.after(50, self._poll)

    def run(self):
        self.root.mainloop()


# ============================================================
# 入口
# ============================================================

if __name__ == "__main__":
    MainWindow().run()
