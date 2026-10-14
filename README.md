# 工资奖金自动化处理平台

本地 Python/Tk 桌面程序，将工资、工时、银行及奖金资料整理为两家公司 SAP A1–A10 凭证。

## 环境与启动

已验证环境：Windows、CPython 3.13.3、Microsoft Excel、openpyxl 3.1.5。安装 Python 时保留 Tcl/Tk 支持。SAP 凭证发布时会调用本机 Excel，将工作簿转换为真正的 Excel 97–2003 XLS 格式；未安装 Excel 时会停止发布并明确报错。

在仓库根目录执行：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\pythonw.exe 工作区域/运行-1.工资奖金自动化处理平台.pyw
```

双击入口时须确保关联到安装了依赖的 Python。月度 Excel 和基础资料仍在本地维护，不随 Git 备份。

## 月度运行

1. 放入当月资料，更新工时和 KOK3。
2. 选择处理月份、公司和凭证，完成预检。
3. 生成后从原来的 `03_运行输出` 目录取最新版。
4. 查阅 `04_归档留痕/<批次>/最新运行.json` 和 `运行历史/<run_id>/run.json` 获取准确状态。

重复运行会更新选中凭证组，未选分组保留。每组来源记录在凭证旁的 `.groups.json`，不要单独修改它。全组文件不是一次运行的同义词；按组来源和输入清单复核。

可靠交易流水号用于去重。没有可靠流水号的跨文件同特征记录会阻断，请整理银行源资料后重新预检，不要修改金额绕过。单文件重复无标识交易保留次数并提示。

## 发布失败与恢复

正式文件先暂存、独立重读校验，再建立旧版备份和发布日志。失败时恢复旧版；进程中断后，下次 GUI 启动或正式运行会自动恢复未完成发布。若提示“待恢复”，关闭占用的 Excel 后重新启动，不要删除历史或锁文件。操作系统会自动释放已退出进程的锁。

实发薪酬表与重点税源报表是附属产出，失败不阻止合格凭证交付。界面和运行记录会注明“本次未更新”；旧报表不是本次成功结果。附属报表旁的 `.status.json` 记录上次成功运行。

历史目录包含旧版本、成功快照和暂存数据，默认不自动清理。Git 回退代码不会恢复或删除业务结果。

## 开发验证

```powershell
python -m unittest discover -s tests
python -m compileall -q 工作区域/payroll_tool
python tests/gui_smoke.py
python tests/historical_regression.py
```

普通测试使用合成数据。GUI 冒烟在隐藏窗口和临时批次中执行。历史回归读取本地 2603、2606、2607、2608，复制到临时目录，对比基线 `19e373e` 与当前代码输出；不会覆盖生产结果，不会把金额明细加入仓库。

职责分层：`domain/rules/workbook_utils` → `contexts/payroll_input` → `vouchers/validation` → `voucher_output/payroll_output/artifacts` → `core/pipeline/gui`。`publication` 单独负责运行锁、暂存、合并、发布及恢复。

进一步阅读：项目会话与历史整理、现状与历史差异说明，以及工作区域中的技术交接说明。
