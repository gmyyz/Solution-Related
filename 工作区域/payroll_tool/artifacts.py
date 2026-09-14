"""artifacts: extracted payroll responsibilities; business rules preserved."""
import os
import shutil
from datetime import date, datetime
from .options import (
    ACTUAL_VOUCHERS,
    VOUCHER_DISPLAY_ORDER,
    build_batch_layout,
    format_period_label,
    normalize_run_options,
    requires_bank_data,
    requires_bonus_data,
    requires_co_data,
    requires_shared_expense_data,
)
from .paths import (
    _ensure_directory,
)
from .domain import (
    EXAMPLE_IMAGE_DIRNAME,
)


def _write_utf8_file(path, content):
    _ensure_directory(os.path.dirname(path))
    with open(path, 'w', encoding='utf-8') as fh:
        fh.write(content)
    return path


def _render_markdown_table(headers, rows):
    lines = ['| ' + ' | '.join(headers) + ' |', '| ' + ' | '.join(['---'] * len(headers)) + ' |']
    for row in rows:
        lines.append('| ' + ' | '.join(str(item) for item in row) + ' |')
    return '\n'.join(lines)


def _render_markdown_list(items, empty_text='无'):
    if not items:
        return f'- {empty_text}'
    return '\n'.join(f'- {item}' for item in items)


def _status_label(status):
    mapping = {
        'ok': '通过',
        'warn': '预警',
        'warning': '预警',
        'error': '阻断',
        'err': '阻断',
        'info': '信息',
    }
    return mapping.get(status, status or '信息')


def _voucher_detail_specs():
    return {
        'A1': {
            'title': 'A1 社保扣款凭证',
            'purpose': '社保发放',
            'debit': '借方取工资表社保相关应付科目汇总。',
            'credit': '贷方取处理月份银行流水中的社保国库流水，按实际流水逐行展开。',
            'sources': [
                '工资整理文件总计行 `F/G/H/I/K/L/M` 七列，分别对应公司养老、公司失业、公司工伤、公司医疗、个人养老、个人失业、个人医疗。',
                '银行流水首表中 `B列公司代码`、`D列交易日期`、`M列付款名称`、`O列付款金额`、`AH列银行科目`。',
            ],
            'timesheet': [],
            'special': ['凭证类型固定 `KZ`；银行行项目原因代码固定写 `202`。'],
            'debit_detail': [
                '借方直接取工资整理文件总计行 `F/G/H/I/K/L/M` 七列金额。',
                '科目映射固定为：`F->2211020002`、`G->2211030002`、`H->2211040001`、`I->2211060002`、`K->2211020001`、`L->2211030001`、`M->2211060001`。',
            ],
            'credit_detail': [
                '先在次月银行流水里筛出 `M列付款名称=国家金库北京市分库` 且 `O列>0` 的国库流水。',
                '以工资表总计行 `F+G+H+I+K+L+M` 为目标，遍历最多 6 笔国库流水组合，匹配成功的流水作为 A1 贷方。',
                '贷方会计科目取银行流水 `AH列`，金额取 `O列`，凭证日期取 `D列` 最大日期。',
            ],
            'checks': ['A1 借贷合计 = 工资表总计行 `F+G+H+I+K+L+M` = 银行社保流水合计。'],
        },
        'A2': {
            'title': 'A2 公积金扣款凭证',
            'purpose': '公积金发放',
            'debit': '借方取公积金应付科目，按公司/个人口径汇总。',
            'credit': '贷方取银行流水中公积金付款行的银行科目与金额。',
            'sources': [
                '工资整理文件总计行 `J列公司公积金` 与 `N列个人公积金` 只用于校验。',
                '银行流水首表中 `M列付款名称=北京住房公积金管理中心` 的付款行，金额取 `O列`，银行科目取 `AH列`。',
            ],
            'timesheet': [],
            'special': ['`2050 / 202602` 特殊月份允许差额 `3650`。', '凭证类型固定 `KZ`；银行行项目原因代码固定写 `202`。'],
            'debit_detail': [
                '脚本不是直接取工资表 `J/N` 入账，而是先汇总银行公积金流水 `O列` 金额。',
                '再将银行总额按 1:1 拆成两行：个人公积金 `2211070001` 与公司公积金 `2211070002`。',
            ],
            'credit_detail': [
                '贷方逐行展开公积金银行流水。',
                '每条贷方的会计科目取银行流水 `AH列`，金额取 `O列`，凭证日期取该批流水 `D列` 最大日期。',
            ],
            'checks': ['工资表总计行 `J+N` 与银行公积金流水 `O列合计` 做核对；若属于特殊月份则按例外规则提示。'],
        },
        'A3': {
            'title': 'A3 工资扣款凭证',
            'purpose': '工资、劳务费与个税发放',
            'debit': '借方取工资表中的劳务费、个税、应付工资，劳务费按行展开。',
            'credit': '贷方取代发工资组合流水和个税国库流水。',
            'sources': [
                '工资整理文件逐行的 `D列项目类型`、`E列金额`、`O列个税`、`Q列实发工资`、`S列成本中心`、`T列内部订单`。',
                '次月银行流水首表中 `P列用途含代发工资` 的工资行，以及 `M列付款名称=国家金库北京市分库` 的国库行。',
            ],
            'timesheet': [],
            'special': ['工资、个税、社保银行侧均支持最多 6 笔流水组合匹配。', '劳务费借方逐行带出成本中心和内部订单。'],
            'debit_detail': [
                '若 `D列=劳务费`，则逐行取 `E列` 税前金额生成借方 `6601110001`，并带出该行 `S列成本中心` 与 `T列内部订单`。',
                '借方 `2221070000` 取匹配到的个税银行流水扣除劳务费个税后的金额；银行个税包含薪资、劳务费和已确认的上月年终奖个税。',
                '若 `D列=薪资`，则把各行 `Q列` 汇总生成借方 `2211010001`。',
            ],
            'credit_detail': [
                '先用工资整理文件总计行 `Q列` 作为匹配目标，在银行流水中寻找最多 6 笔 `P列用途含代发工资` 的组合。',
                '工资贷方的会计科目取这些银行行的 `AH列`，金额取各自 `O列`。',
                '个税贷方用工资整理文件全部 `O列` 个税加上已确认的上月年终奖个税作为目标，在国库流水中寻找最多 6 笔组合；匹配到的流水逐行展开。',
            ],
            'checks': ['非劳务 `Q列汇总 -> 2211010001`。', '个税银行流水组合 -> `2221070000`。', '银行侧必须匹配“工资组合 + 个税国库行”。'],
        },
        'A4': {
            'title': 'A4 上月实际工资入账',
            'purpose': '上月实际工资入账',
            'debit': '借方按成本中心+内部订单归集工资、社保、公积金等费用。',
            'credit': '贷方挂应付社保、公积金、个税、实发工资及其他项目。',
            'sources': [
                '工资整理文件逐行的 `B列部门`、`C列项目`、`D列项目类型`、`E/F/G/H/I/J` 费用列、`O/P/Q` 往来列、`S/T` 核算维度列。',
                'A4 的内部订单默认取 `T列`，但 OSD 特例会结合工时匹配结果重算。',
            ],
            'timesheet': [
                '内部订单默认取工资整理文件 `T列`。',
                '若 `B列=OSD运营支持部` 且 `C列项目` 在工时数据中被识别为“自研”，则内部订单改取 `OSD运营支持部日常工作` 对应订单。',
            ],
            'special': ['排除劳务费。', '`OSD运营支持部` 且项目类型为“自研”时，内部订单按 `OSD运营支持部日常工作` 处理。'],
            'debit_detail': [
                '跳过 `D列=劳务费` 的行。',
                '借方逐行读取 `E/F/G/H/I/J` 六列金额，并按 `S列成本中心 + T列/重算后内部订单` 聚合。',
                '科目映射固定为：`E->6601010001`、`F->6601030008`、`G->6601030002`、`H->6601030003`、`I->6601030005`、`J->6601030006`。',
            ],
            'credit_detail': [
                '贷方的社保、公积金应付科目直接取工资整理文件总计行 `F/G/H/I/J/K/L/M/N`。',
                '个税贷方取所有 `D列=薪资` 行的 `O列` 汇总，实发工资贷方取这些行 `Q列` 汇总，其他项目贷方取这些行 `P列` 汇总。',
                '其中 `F/G/H/I/J/N/M/L/K` 分别映射到 `2211020002/2211030002/2211040001/2211060002/2211070002/2211070001/2211060001/2211030001/2211020001`。',
            ],
            'checks': ['A4 自身借贷必须平。', '若同时选择 A1-A4，会做重点往来科目对冲。'],
        },
        'A5': {
            'title': 'A5 当月工资计提',
            'purpose': '当月工资计提',
            'debit': '借方与 A4 同口径，但只汇总到成本中心层级。',
            'credit': '贷方按应付工资、社保、公积金、个税等科目汇总。',
            'sources': ['与 A4 使用同一份工资整理文件，仍读取 `D/E/F/G/H/I/J/O/P/Q/S`，但借方只保留 `S列成本中心` 维度。'],
            'timesheet': [],
            'special': ['A5 是 A6 的分摊前底稿。'],
            'debit_detail': [
                '跳过 `D列=劳务费` 的行。',
                '借方仍读取 `E/F/G/H/I/J` 六列费用，但仅按 `S列成本中心` 聚合，不再区分 `T列内部订单`。',
                '科目映射与 A4 相同：`E->6601010001`、`F->6601030008`、`G->6601030002`、`H->6601030003`、`I->6601030005`、`J->6601030006`。',
            ],
            'credit_detail': [
                '贷方逻辑与 A4 一致：总计行 `F/G/H/I/J/K/L/M/N` 进应付社保公积金，各薪资行 `O/P/Q` 分别汇总到个税、其他项目、实发工资。',
            ],
            'checks': ['A5 总额应与 A4 总额一致。'],
        },
        'A6': {
            'title': 'A6 研发工时分摊',
            'purpose': '研发工时分摊',
            'debit': '借方按研发工时比例把费用分摊到内部订单。',
            'credit': '贷方先冲回 A5 原计提的非 OSD 费用。',
            'sources': [
                'A5 分摊前底稿来自工资整理文件中非劳务行的 `E/F/G/H/I/J` 六列，并按 `S列成本中心` 聚合。',
                '工时数据 `工时汇总` 页读取 `B列年份`、`C列月份`、`E列部门`、`F列公司`、`K列工时`、`M列内部订单`。',
            ],
            'timesheet': [
                '只看处理月份当月工时，例如 2603 只看 2026年3月工时。',
                '只统计 `K > 0` 的有效工时行。',
                '按 `E列部门 + F列公司 + M列内部订单` 汇总工时，再在部门内计算分摊比例。',
                '`OSD运营支持部` 不参与 A6 分摊。',
            ],
            'special': [
                '科目不变、成本中心不变，贷方只是冲回 A5。',
                '统一文本：`根据研发工时分摊计提的YYYY年M月人工费用`。',
                '分摊结果统一保留两位小数，尾差回补到最后一个内部订单。',
            ],
            'checks': ['A6 冲回金额 = A6 分摊金额。', '若某部门无可用工时，则该部门不能正常分摊。'],
            'debit_detail': [
                '借方金额来源于 A5 已聚合好的 `E/F/G/H/I/J` 六列费用。',
                '每个成本中心先反查所属部门，再把同部门当月工时按内部订单比例分配；借方成本中心沿用原 `S列`，内部订单改写为工时表 `M列`。',
            ],
            'credit_detail': [
                '贷方直接冲回 A5 的原借方金额，科目和成本中心不变，内部订单留空。',
            ],
        },
        'A7': {
            'title': 'A7 年终奖计提',
            'purpose': '年终奖计提',
            'debit': '借方按成本中心计提年终奖费用。',
            'credit': '贷方汇总到应付工资科目。',
            'sources': [
                '奖金文件 `部门统计` 页中 `A列公司`、`B列部门`、`C:N列1-12月金额`。',
                '奖金文件 `计提比例` 页中 `A列月份`、`B列计提系数`。',
            ],
            'timesheet': [],
            'special': ['季度末取 `1-当月累计奖金 × 当月系数`。', '非季度末取 `当月奖金 × 当月系数`。', '系数不追溯历史月份。'],
            'checks': ['每家公司贷方通常汇总为 1 行。'],
            'debit_detail': [
                '先按处理月份判断取数范围：季度末取 `C:N` 中 1-当月累计，非季度末只取对应月份列。',
                '每个部门金额再乘 `计提比例` 页 `B列` 的当月系数，借方科目固定为 `6601010003`，成本中心来自 Mapping 反查结果。',
            ],
            'credit_detail': [
                '贷方按公司汇总为 1 行，科目固定 `2211010001`。',
            ],
        },
        'A8': {
            'title': 'A8 年终奖分摊',
            'purpose': '年终奖分摊',
            'debit': '借方按奖金口径下的工时把年终奖分配到内部订单。',
            'credit': '贷方先冲回 A7 中需分摊的奖金费用。',
            'sources': [
                'A7 已算出的部门奖金金额与成本中心。',
                '工时数据 `工时汇总` 页的 `B/C/E/F/K/M`，季度末按累计月数合并。',
            ],
            'timesheet': [
                '1 月只看 1 月工时。',
                '2 月只看 2 月工时。',
                '3/6/9/12 月看 `1-当月` 的累计工时。',
                '`OSD` 不参与分摊。',
            ],
            'special': [
                'A8 只重分摊需要分摊的奖金部分，不一定等于整张 A7。',
                '统一文本：`根据研发工时分摊计提的期间年终奖`，例如 `根据研发工时分摊计提的2026Q1年终奖`。',
            ],
            'checks': ['A8 分摊前总额 = A8 分摊后总额。'],
            'debit_detail': [
                '借方科目固定 `6601010003`，金额取 A7 每个部门的奖金金额。',
                '借方成本中心沿用该部门成本中心，内部订单按奖金口径工时比例写入工时表 `M列` 对应订单。',
            ],
            'credit_detail': [
                '贷方先按部门冲回 A7 中需要分摊的金额，科目同样固定 `6601010003`，只带成本中心不带内部订单。',
            ],
        },
        'A9': {
            'title': 'A9 CO工单分摊',
            'purpose': 'CO工单分摊',
            'debit': '借方按 CO 工单比例分摊到内部订单，科目按规则映射。',
            'credit': '贷方取待分摊费用，固定成本中心和固定内部订单冲销。',
            'sources': [
                '`CO工单分摊.xlsx` 中 `CO工单清单` 工作表的订单比例底稿。',
                '同文件 `待分摊费用` 工作表中的待分摊费用行。',
            ],
            'timesheet': [],
            'special': [
                '仅适用于 `2050`。',
                '`G列 != 0` 的明细行参与分摊，最后一行加总仅作校验。',
                '贷方固定成本中心 `20502020`、固定内部订单 `9201856`。',
                '部分成本要素映射到 `5001010011`，其余映射到 `5001010012`。',
            ],
            'checks': ['`CO工单分摊` 明细 G 列合计必须等于最后一行加总。', 'A9 借贷合计必须一致。'],
            'debit_detail': [
                '分摊比例底稿来自 `CO工单分摊` 页：`A列内部订单`、`E列实际成本借方`、`G列总的实际成本`。',
                '仅 `G列 != 0` 的明细参与分摊，比例分母取这些明细 `E列` 合计。',
                '借方内部订单取 `A列`，借方科目按 `待分摊费用` 页 `A列成本要素` 映射到 `5001010011/5001010012`。',
            ],
            'credit_detail': [
                '贷方逐行读取 `待分摊费用` 页 `A列成本要素` 与 `D列待分摊金额`。',
                '贷方成本中心固定 `20502020`，内部订单固定 `9201856`，并在 SAP 模板 `O列反记账` 写 `x`。',
            ],
        },
        'A10': {
            'title': 'A10 研发费用分摊',
            'purpose': '研发费用分摊',
            'debit': '借方按费控科目+部门聚合后的金额，再按研发工时分摊到项目订单。',
            'credit': '贷方先按原费用科目、原成本中心、原内部订单冲回待分摊费用。',
            'sources': [
                '`待分摊费用YYMM.xlsx` 中 E/J/L/S/T/U 等关键字段。',
                '工时数据中的部门工时与内部订单。',
                'Mapping 表中的成本中心反查部门规则。',
            ],
            'timesheet': [
                '分摊月份只看处理月份当月工时。',
                '绝不跨公司分摊。',
                '默认按部门内工时分摊。',
                '`OSD` 不参与分摊。',
                '部门先用成本中心编码反查，再去工时数据中取比例，而不是直接信任文本名称。',
            ],
            'special': [
                'A10 先按 `费控科目 + 部门` 聚合，再把同组正负金额抵减后分摊。',
                '`待分摊费用` 中 `BG列功能范围 = 1000` 的费用不参与分摊。',
                '`LHD硬件逻辑部` 与 `HLD硬件逻辑部` 视为同一部门。',
                '`HLD硬件逻辑部` 分摊时，`9201856` 不进入分母，也不作为分子。',
                '`2050 / 20502050` 的借方分摊按 `20502060` 工时口径处理，但贷方仍维持原成本中心。',
                '`2060 / 20602020` 若本部门无可用工时，则改按整个 `耐数信息` 公司工时分摊。',
                '统一文本：`根据研发工时分摊YYYY年M月研发费用`。',
            ],
            'checks': ['A10 借方合计 = 贷方合计。', '每组聚合后的贷方金额应等于借方分摊合计。'],
            'debit_detail': [
                '源数据先读 `E列公司代码`、`J列会计科目`、`L列金额`、`S列成本中心`、`T列部门文本`、`U列原内部订单`、`BG列功能范围`。',
                '脚本先用 `S列` 通过 Mapping 反查标准部门，再按 `J列会计科目 + 标准部门` 聚合 `L列` 金额；同组正负先抵减，净额为 0 的组直接跳过。',
                '借方科目仍取原 `J列`，借方成本中心默认取原 `S列`，若 `2050/20502050` 特例则改用 `20502060` 作为分摊成本中心；内部订单按工时表 `M列` 比例写入。',
            ],
            'credit_detail': [
                '贷方按聚合后的组逐组冲回，科目取原 `J列`，金额取聚合后的净额，成本中心取原 `S列`，内部订单取原 `U列`。',
                '凭证类型固定 `SA`，并保持“一个聚合组一条贷方”。',
            ],
        },
    }


def _voucher_summary_rows(run_options):
    descriptions = _voucher_detail_specs()
    return [
        (
            voucher,
            descriptions[voucher]['purpose'],
            descriptions[voucher]['debit'],
            descriptions[voucher]['credit'],
        )
        for voucher in run_options.vouchers
        if voucher in descriptions
    ]


def _voucher_example_assets():
    return {
        'A1': [
            ('a1-payroll-social.png', '工资整理文件：总计行中 F:I 与 K:M 七列社保金额。'),
            ('a1-bank-social.png', '银行流水：同批两条国库流水中，较大金额对应社保。'),
        ],
        'A2': [
            ('a2-payroll-fund.png', '工资整理文件：总计行中的公司/个人公积金。'),
            ('a2-bank-fund.png', '银行流水：北京住房公积金管理中心对应付款行。'),
        ],
        'A3': [
            ('a3-payroll-salary.png', '工资整理文件：劳务费逐行、总计行 O/Q 汇总、以及 S/T 映射结果。'),
            ('a3-bank-salary-tax.png', '银行流水：代发工资组合与个税国库行。'),
        ],
        'A4': [
            ('a4-payroll-source.png', '工资整理文件：按成本中心和内部订单聚合后的来源区域。'),
            ('a4-voucher.png', 'A4 凭证：借方费用与贷方应付项目。'),
        ],
        'A5': [
            ('a5-voucher.png', 'A5 凭证：借方只到成本中心层级。'),
        ],
        'A6': [
            ('a6-timesheet.png', 'A6 工时示例：处理月份当月工时。'),
            ('a6-detail.png', 'A5-A6 核对表：A5 金额到 A6 分摊金额的展开。'),
            ('a6-voucher.png', 'A6 凭证：先冲回，再按工时分配到内部订单。'),
        ],
        'A7': [
            ('a7-bonus-source.png', 'A7 奖金来源：季度累计奖金基础金额。'),
            ('a7-bonus-ratio.png', 'A7 奖金系数：3 月对应奖金系数。'),
            ('a7-voucher.png', 'A7 凭证：借方按成本中心，贷方公司汇总。'),
        ],
        'A8': [
            ('a8-timesheet-q1.png', 'A8 工时示例：季度末看 1-3 月累计工时。'),
            ('a8-voucher.png', 'A8 凭证：冲回 A7 后重新分配到内部订单。'),
        ],
        'A9': [
            ('a9-co-source.png', 'A9 来源示例：CO工单清单工作表中的订单与比例底稿。'),
            ('a9-co-expense.png', 'A9 来源示例：待分摊费用工作表中的成本要素与金额。'),
        ],
        'A10': [
            ('a10-shared-expense-source.png', 'A10 来源示例：待分摊费用中的公司、科目、金额、成本中心和原订单。'),
        ],
    }


def _copy_example_images_for_batch(base_dir, archive_root, batch_code, vouchers):
    if batch_code != '2603':
        return
    asset_map = _voucher_example_assets()
    target_dir = os.path.join(archive_root, EXAMPLE_IMAGE_DIRNAME)
    os.makedirs(target_dir, exist_ok=True)
    for voucher in vouchers:
        for file_name, _caption in asset_map.get(voucher, []):
            source_path = os.path.join(base_dir, 'doc_assets', file_name)
            target_path = os.path.join(target_dir, file_name)
            if os.path.exists(source_path):
                shutil.copyfile(source_path, target_path)


def _archive_timesheet_snapshot(layout, timesheet_path):
    if not timesheet_path or not os.path.isfile(timesheet_path):
        return ''
    snapshot_dir = os.path.join(layout.archive_root, '06_基础资料快照')
    _ensure_directory(snapshot_dir)
    snapshot_path = os.path.join(snapshot_dir, f'{layout.batch_code}-工时数据留档.xlsx')
    if os.path.abspath(timesheet_path) != os.path.abspath(snapshot_path):
        shutil.copyfile(timesheet_path, snapshot_path)
    return snapshot_path


def _render_example_images(base_dir, batch_code, voucher):
    if batch_code != '2603':
        return ''
    asset_map = _voucher_example_assets()
    examples = asset_map.get(voucher, [])
    if not examples:
        return ''
    lines = ['#### 2026年3月样例截图', '']
    for file_name, caption in examples:
        asset_path = os.path.join(base_dir, 'doc_assets', file_name)
        if os.path.exists(asset_path):
            lines.append(f'![{voucher} 示例](./{EXAMPLE_IMAGE_DIRNAME}/{file_name})')
            lines.append(f'_{caption}_')
            lines.append('')
    return '\n'.join(lines).strip()


def _render_voucher_detail_sections(run_options, base_dir, batch_code):
    specs = _voucher_detail_specs()
    blocks = []
    for voucher in run_options.vouchers:
        spec = specs.get(voucher)
        if not spec:
            continue
        blocks.extend(
            [
                f'### {spec["title"]}',
                '',
                f'- 用途：{spec["purpose"]}',
                '',
                '#### 取数来源',
                '',
                _render_markdown_list(spec.get('sources', [])),
                '',
                '#### 借方生成逻辑',
                '',
                _render_markdown_list(spec.get('debit_detail', [spec['debit']])),
                '',
                '#### 贷方生成逻辑',
                '',
                _render_markdown_list(spec.get('credit_detail', [spec['credit']])),
                '',
            ]
        )
        if spec.get('timesheet'):
            blocks.extend(
                [
                    '#### 工时选取逻辑',
                    '',
                    _render_markdown_list(spec['timesheet']),
                    '',
                ]
            )
        if spec.get('special'):
            blocks.extend(
                [
                    '#### 特殊规则',
                    '',
                    _render_markdown_list(spec['special']),
                    '',
                ]
            )
        if spec.get('checks'):
            blocks.extend(
                [
                    '#### 关键勾稽与校验',
                    '',
                    _render_markdown_list(spec['checks']),
                    '',
                ]
            )
        example_block = _render_example_images(base_dir, batch_code, voucher)
        if example_block:
            blocks.extend([example_block, ''])
    return '\n'.join(blocks).strip()
    return [
        (
            voucher,
            descriptions[voucher]['purpose'],
            descriptions[voucher]['debit'],
            descriptions[voucher]['credit'],
        )
        for voucher in run_options.vouchers
        if voucher in descriptions
    ]


def _write_run_artifacts(
    base_dir,
    run_options,
    input_path,
    mapping_path,
    bank_dir,
    timesheet_path,
    bonus_path,
    co_path,
    shared_path,
    output_paths,
    voucher_paths,
    voucher_validation_summary,
    log_entries,
    bank_scan_stats,
    compensation_report_path=None,
):
    layout = build_batch_layout(base_dir, run_options.processing_year, run_options.processing_month)
    input_manifest_path = os.path.join(layout.archive_root, f'{layout.batch_code}-输入清单.md')
    run_log_path = os.path.join(layout.archive_root, f'{layout.batch_code}-运行日志.md')
    process_note_path = os.path.join(layout.archive_root, f'{layout.batch_code}-过程说明.md')
    output_manifest_path = os.path.join(layout.archive_root, f'{layout.batch_code}-输出清单.md')
    warning_logs = [entry['text'] for entry in log_entries if entry.get('tag') == 'warn']
    error_logs = [entry['text'] for entry in log_entries if entry.get('tag') in ('error', 'err')]
    voucher_logic_rows = _voucher_summary_rows(run_options)
    _copy_example_images_for_batch(base_dir, layout.archive_root, layout.batch_code, run_options.vouchers)
    timesheet_snapshot_path = _archive_timesheet_snapshot(layout, timesheet_path)
    voucher_detail_markdown = _render_voucher_detail_sections(run_options, base_dir, layout.batch_code)

    input_rows = [
        ('原始工资单', '必须', input_path, '本次唯一工资源文件'),
        ('Mapping表', '必须', mapping_path, '基础主数据'),
        ('工时数据', '必须', timesheet_path, '基础资料来源；已按本批次单独留档'),
        ('银行流水目录', '按需', bank_dir if requires_bank_data(run_options) else '未使用', 'A1-A3 时启用'),
        ('奖金数据', '按需', bonus_path if requires_bonus_data(run_options) else '未使用', 'A7-A8 时启用'),
        ('CO工单分摊', '按需', co_path if requires_co_data(run_options) else '未使用', 'A9 时启用'),
        ('待分摊费用', '按需', shared_path if requires_shared_expense_data(run_options) else '未使用', 'A10 时启用'),
    ]
    snapshot_rows = [('工时数据留档', timesheet_snapshot_path)] if timesheet_snapshot_path else []
    input_content = '\n'.join(
        [
            f'# {layout.batch_name} 输入清单',
            '',
            '## 一页摘要',
            '',
            f'- 批次编码：`{layout.batch_code}`',
            f'- 处理月份：`{run_options.processing_label}`',
            f'- 工资所属月份：`{run_options.payroll_label}`',
            f'- 公司范围：`{"、".join(run_options.companies) or "未选"}`',
            f'- 凭证范围：`{"、".join(run_options.vouchers) or "未选"}`',
            f'- 已记录输入项：`{len(input_rows)}` 项',
            '',
            '## 批次定位',
            '',
            _render_markdown_table(
                ['项目', '值'],
                [
                    ('月度输入目录', layout.monthly_input_root),
                    ('运行输出目录', layout.run_output_root),
                    ('归档留痕目录', layout.archive_root),
                ],
            ),
            '',
            '## 实际使用文件',
            '',
            _render_markdown_table(['资料', '要求', '路径/值', '用途'], input_rows),
            '',
            '## 基础资料版本留档',
            '',
            _render_markdown_table(['资料', '留档路径'], snapshot_rows or [('无', '无')]),
            '',
            '## 银行流水扫描统计',
            '',
            _render_markdown_table(
                ['指标', '值'],
                [
                    ('扫描文件数', bank_scan_stats.get('source_file_count', 0)),
                    ('扫描记录数', bank_scan_stats.get('scanned_row_count', 0)),
                    ('保留记录数', bank_scan_stats.get('kept_row_count', 0)),
                    ('去重记录数', bank_scan_stats.get('deduped_row_count', 0)),
                ],
            ),
            '',
            '## 使用说明',
            '',
            '- 本文档用于回答“本次到底用了什么输入资料”。',
            '- 若后续重跑本批次，应优先核对本清单中的路径与月份口径。',
            '',
        ]
    )

    log_rows = [(idx, _status_label(entry.get('tag')), entry['text']) for idx, entry in enumerate(log_entries, start=1)]
    run_log_content = '\n'.join(
        [
            f'# {layout.batch_name} 运行日志',
            '',
            '## 执行摘要',
            '',
            f'- 运行时间：`{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}`',
            f'- 处理月份：`{run_options.processing_label}`',
            f'- 公司范围：`{"、".join(run_options.companies)}`',
            f'- 凭证范围：`{"、".join(run_options.vouchers)}`',
            f'- 记录日志：`{len(log_entries)}` 条',
            f'- 预警日志：`{len(warning_logs)}` 条',
            f'- 错误日志：`{len(error_logs)}` 条',
            '',
            '## 重点关注',
            '',
            _render_markdown_list(warning_logs, empty_text='本次运行未出现预警日志'),
            '',
            '## 运行时间线',
            '',
            _render_markdown_table(['序号', '级别', '内容'], log_rows),
            '',
        ]
    )

    process_note_content = '\n'.join(
        [
            f'# {layout.batch_name} 过程说明',
            '',
            '## 文档用途',
            '',
            '- 本文档用于解释本批次凭证是按什么口径生成的。',
            '- 重点服务于复核、交接和后续审计沟通。',
            '',
            '## 本次处理口径',
            '',
            _render_markdown_table(
                ['项目', '值'],
                [
                    ('处理月份', run_options.processing_label),
                    ('工资所属月份', run_options.payroll_label),
                    ('公司范围', '、'.join(run_options.companies)),
                    ('凭证范围', '、'.join(run_options.vouchers)),
                ],
            ),
            '',
            '## 处理主链路',
            '',
            '1. 读取工资单、Mapping、工时和按需资料。',
            '2. 对工资单执行 A/B 列填充、Q 列校验、S/T 列映射。',
            '3. 如选择 A1-A3，则扫描银行流水目录、筛选处理月份并自动去重。',
            '4. 按公司拆分整理后工资单，再生成所选凭证。',
            '5. 同步输出输入清单、预检报告、运行日志、过程说明和输出清单。',
            '',
            '## 凭证生成逻辑摘要',
            '',
            _render_markdown_table(['凭证', '用途', '借方生成逻辑', '贷方生成逻辑'], voucher_logic_rows),
            '',
            '## 凭证明细逻辑',
            '',
            voucher_detail_markdown,
            '',
            '## 关键规则',
            '',
            '- 银行流水按目录全量扫描后再按处理月份筛选，并对按可靠交易流水号去重，无标识跨文件重复需核对。',
            '- A10 先按费控科目 + 部门聚合，再进行工时分摊。',
            '- 工时数据作为可滚动更新的基础资料维护；每次正式运行都会把当次使用版本复制到本批次留痕目录。',
            '- 比例分摊保留两位小数：A10 按最大余数补分；其他凭证保留现行逐笔舍入和尾差规则。',
            '',
            '## 建议复核点',
            '',
            _render_markdown_list(
                warning_logs if warning_logs else ['优先核对工资整理文件、银行核对摘要和跨凭证对冲结果。'],
                empty_text='优先核对工资整理文件、银行核对摘要和跨凭证对冲结果。',
            ),
            '',
        ]
    )

    payroll_output_rows = [(company, path) for company, path in output_paths.items()]
    voucher_output_rows = [(category, path) for category, path in voucher_paths.items()]
    supplemental_output_rows = (
        [('实发薪酬表', compensation_report_path)]
        if compensation_report_path
        else []
    )
    validation_rows = []
    for company, validation in voucher_validation_summary.items():
        issue_count = len(validation.get('cross_group_issues', []))
        validation_rows.append((company, '通过' if issue_count == 0 else f'异常 {issue_count} 项'))
    output_manifest_content = '\n'.join(
        [
            f'# {layout.batch_name} 输出清单',
            '',
            '## 输出摘要',
            '',
            f'- 输出根目录：`{layout.run_output_root}`',
            f'- 留痕根目录：`{layout.archive_root}`',
            f'- 工资整理文件：`{len(payroll_output_rows)}` 个',
            f'- 凭证文件：`{len(voucher_output_rows)}` 个',
            f'- 补充报表：`{len(supplemental_output_rows)}` 个',
            f'- 基础资料留档：`{len(snapshot_rows)}` 个',
            f'- 公司级校验结果：`{len(validation_rows)}` 项',
            '',
            '## 工资整理文件',
            '',
            _render_markdown_table(['公司', '路径'], payroll_output_rows or [('无', '无')]),
            '',
            '## 凭证文件',
            '',
            _render_markdown_table(['对象', '路径'], voucher_output_rows or [('无', '无')]),
            '',
            '## 补充报表',
            '',
            _render_markdown_table(['对象', '路径'], supplemental_output_rows or [('无', '无')]),
            '',
            '## 基础资料留档',
            '',
            _render_markdown_table(['资料', '路径'], snapshot_rows or [('无', '无')]),
            '',
            '## 校验结果',
            '',
            _render_markdown_table(['公司', '结果'], validation_rows or [('无', '无')]),
            '',
            '## 说明',
            '',
            '- 本文档用于回答“本次到底产出了什么文件”。',
            '- 若后续重跑，同名文件会被最新结果覆盖，应结合运行日志一起查看。',
            '',
        ]
    )

    return {
        'input_manifest': _write_utf8_file(input_manifest_path, input_content),
        'run_log': _write_utf8_file(run_log_path, run_log_content),
        'process_note': _write_utf8_file(process_note_path, process_note_content),
        'output_manifest': _write_utf8_file(output_manifest_path, output_manifest_content),
        'timesheet_snapshot': timesheet_snapshot_path,
    }
