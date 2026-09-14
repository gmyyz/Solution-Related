"""Journaled publication of a payroll run. Source paths are never redirected."""
import json
import os
import shutil
import hashlib
import re
import uuid
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from copy import copy
from .workbook_io import load_workbook

_targets = ContextVar('payroll_output_targets', default=None)


def digest(path):
    with open(path, 'rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def write_json(path, value):
    path = Path(path); path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + '.tmp')
    with open(temporary, 'w', encoding='utf-8') as stream:
        json.dump(value, stream, ensure_ascii=False, indent=2, default=str)
        stream.flush(); os.fsync(stream.fileno())
    os.replace(temporary, path)


@dataclass(frozen=True)
class OutputTargets:
    base: Path
    stage: Path

    def path(self, path):
        relative = Path(path).resolve().relative_to(self.base)
        if relative.parts[0] not in ('03_运行输出', '04_归档留痕'):
            raise ValueError(f'不是输出路径：{path}')
        return str(self.stage / relative)


def output_target(path):
    targets = _targets.get()
    return targets.path(path) if targets else str(path)


@contextmanager
def workspace_lock(base):
    path = Path(base) / '.payroll-run.lock'
    path.parent.mkdir(parents=True, exist_ok=True)
    stream = open(path, 'a+b')
    try:
        stream.seek(0)
        if os.name == 'nt':
            import msvcrt
            if path.stat().st_size == 0:
                stream.write(b'0'); stream.flush()
            stream.seek(0)
            try: msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
            except OSError as exc: raise RuntimeError('已有运行正在发布，请等待完成') from exc
        else:
            import fcntl
            try: fcntl.flock(stream, fcntl.LOCK_EX | fcntl.LOCK_NB)
            except OSError as exc: raise RuntimeError('已有运行正在发布，请等待完成') from exc
        yield
    finally:
        stream.close()  # OS releases the lock even after a process crash.


def _restore(journal, path):
    try:
        for entry in reversed(journal['entries']):
            target = Path(entry['target'])
            if entry['existed']:
                backup = Path(entry['backup'])
                if not backup.is_file() or digest(backup) != entry['old_hash']:
                    raise RuntimeError(f'恢复备份缺失或已改变：{backup}')
                if target.is_file() and digest(target) == entry['old_hash']:
                    continue
                temp = target.with_name(target.name + '.restore-' + uuid.uuid4().hex)
                shutil.copy2(backup, temp); os.replace(temp, target)
            elif target.exists():
                target.unlink()
        journal['status'] = 'rolled_back'
        write_json(path, journal)
    except Exception as exc:
        journal['status'] = 'needs_recovery'; journal['error'] = str(exc)
        write_json(path, journal)
        raise RuntimeError(f'待恢复，不能继续发布：{path}；{exc}') from exc


def recover_pending(base):
    for path in sorted((Path(base) / '04_归档留痕').glob('*/运行历史/*/publish-*.json')):
        with open(path, encoding='utf-8') as stream: journal = json.load(stream)
        if journal['status'] in ('publishing', 'needs_recovery'):
            # Only restore paths within this workspace, even if a journal was edited.
            for entry in journal['entries']:
                relative = Path(entry['target']).resolve().relative_to(Path(base).resolve())
                if relative.parts[0] not in ('03_运行输出', '04_归档留痕'):
                    raise RuntimeError('待恢复日志包含非输出路径，停止恢复')
                Path(entry['backup']).resolve().relative_to(path.parent.resolve())
            _restore(journal, path)
    for path in sorted((Path(base) / '04_归档留痕').glob('*/运行历史/*/run.json')):
        with open(path,encoding='utf-8') as stream: state=json.load(stream)
        if state['status'] not in ('generating','required_committed','needs_recovery'):
            continue
        required=path.parent/'publish-required.json'
        if required.exists():
            with open(required,encoding='utf-8') as stream: journal=json.load(stream)
            if journal['status']=='committed':
                state.update(status='success_with_warnings',
                             recovery_note='必需结果已交付；进程在完成全部附属状态记录前中断，请查看发布日志')
            else:
                state.update(status='failed',recovery_note='未完成发布已恢复旧版')
        else:
            state.update(status='failed',recovery_note='生成阶段中断，没有发布正式结果')
        write_json(path,state)


def publish_files(base, history, stage, files, label):
    entries = []
    for source in files:
        relative = source.relative_to(stage)
        target = base / relative
        backup = history / ('previous-' + label) / relative
        exists = target.exists()
        if exists:
            backup.parent.mkdir(parents=True, exist_ok=True); shutil.copy2(target, backup)
        entries.append(dict(source=str(source), target=str(target), backup=str(backup),
                            existed=exists, old_hash=digest(target) if exists else None))
    journal = {'status': 'publishing', 'entries': entries}
    journal_path = history / f'publish-{label}.json'
    write_json(journal_path, journal)
    try:
        for entry in entries:
            source, target = Path(entry['source']), Path(entry['target'])
            target.parent.mkdir(parents=True, exist_ok=True)
            snapshot = history / ('success-' + label) / source.relative_to(stage)
            snapshot.parent.mkdir(parents=True, exist_ok=True); shutil.copy2(source, snapshot)
            temporary = target.with_name(target.name + '.publish-' + uuid.uuid4().hex)
            shutil.copy2(source, temporary)
            try: os.replace(temporary, target)
            finally:
                if temporary.exists(): temporary.unlink()
        journal['status'] = 'committed'; write_json(journal_path, journal)
    except Exception:
        _restore(journal, journal_path)
        raise


def _voucher_rows(ws):
    rows = []
    for cells in ws.iter_rows(min_row=2, max_col=17):
        values = {chr(65+i): c.value for i, c in enumerate(cells)}
        if not any(v is not None and v != '' for v in values.values()): continue
        rows.append(values)
    return rows


def check_voucher_rows(rows, options, company_code, allowed):
    from .validation import _validate_voucher_groups
    from .workbook_utils import _to_money
    for row in rows:
        if row['A'] not in allowed or str(row['B']) != company_code:
            raise ValueError('旧凭证公司或分组无法识别，不能合并')
        if str(row['G']) not in ('40', '50') or not row['I']:
            raise ValueError('凭证记账码或科目不完整')
        if not isinstance(row['M'], str) or not re.fullmatch(r'-?\d+\.\d{2}', row['M']):
            raise ValueError('SAP金额必须为两位小数文本')
        amount = _to_money(row['M'])
        if str(row['G']) == '40' and amount < 0: raise ValueError('凭证借方不能为负数')
        value = row['D']
        if isinstance(value, datetime): period = (value.year, value.month)
        else:
            text = str(value).replace('-', '').replace('/', '')
            if len(text) != 8 or not text.isdigit(): raise ValueError(f'凭证日期无法识别：{value}')
            period = (int(text[:4]), int(text[4:6]))
        accepted = {options.processing_period}
        if row['A'] == 'A2': accepted.add(options.payroll_period)
        if period not in accepted: raise ValueError('旧凭证月份不符，不能合并')
        if row['D'] != row['E']: raise ValueError('凭证日期与过账日期不一致')
        if row['A'] in ('A1','A2','A3'):
            if row['C'] != 'KZ': raise ValueError('银行凭证类型必须为 KZ')
            if str(row['G']) == '50' and str(row['P']) != '202':
                raise ValueError('银行贷方原因代码必须为 202')
        if row['A'] == 'A9' and str(row['G']) == '50' and row['O'] != 'x':
            raise ValueError('A9 贷方反记账标记必须为 x')
        if row['A'] == 'A10' and row['C'] != 'SA': raise ValueError('A10 类型必须为 SA')
    _validate_voucher_groups(rows)


def merge_voucher(source, target, options, run_id):
    from .validation import _validate_cross_group_accounts
    from .options import VOUCHER_FILE_GROUPS
    new = load_workbook(source)
    old = None
    try:
        ws = new.worksheets[0]; fresh = _voucher_rows(ws)
        # Directory identity is stable even when a selected voucher has no data.
        family = {'02_实际凭证':'actual','03_计提凭证':'accrual','04_年终奖凭证':'bonus',
                  '06_CO工单分摊':'co','05_研发费用分摊':'rd'}[source.parent.name]
        allowed = set(VOUCHER_FILE_GROUPS[family]); selected = allowed.intersection(options.vouchers)
        company = next(p[:4] for p in source.parts if p.startswith(('2050_', '2060_')))
        check_voucher_rows(fresh, options, company, selected)
        origins = {}; retained = []
        if target.exists():
            old = load_workbook(target); old_ws = old.worksheets[0]
            previous = _voucher_rows(old_ws)
            check_voucher_rows(previous, options, company, allowed)
            sidecar = target.with_suffix(target.suffix + '.groups.json')
            if sidecar.exists():
                with open(sidecar, encoding='utf-8') as stream: previous_meta = json.load(stream)
            else: previous_meta = {}
            trusted = previous_meta.get('file_hash') == digest(target)
            for cells in old_ws.iter_rows(min_row=2, max_col=17):
                group = cells[0].value
                if group in allowed - selected:
                    retained.append(cells)
                    origins[group] = previous_meta.get('groups', {}).get(group) if trusted else {
                        'source': '历史文件导入基线', 'file_hash': digest(target)}
        offset = ws.max_row + 1
        if not fresh: offset = 2
        for cells in retained:
            for col, cell in enumerate(cells, 1):
                dest = ws.cell(offset, col, cell.value)
                dest._style = copy(cell._style)
                if cell.number_format: dest.number_format = cell.number_format
                if cell.comment: dest.comment = copy(cell.comment)
            if old is not None:
                dimension = copy(old.worksheets[0].row_dimensions[cells[0].row])
                dimension.index = offset
                ws.row_dimensions[offset] = dimension
            offset += 1
        combined = _voucher_rows(ws)
        check_voucher_rows(combined, options, company, allowed)
        if retained and {'A1','A2','A3','A4'} <= {r['A'] for r in combined}:
            issues = _validate_cross_group_accounts(combined)
            if issues: raise ValueError('合并后跨凭证对冲异常：' + '；'.join(issues))
        for group in selected:
            origins[group] = {'run_id': run_id, 'status': 'passed' if any(r['A']==group for r in fresh) else 'no_data'}
        new.save(source)
        write_json(source.with_suffix(source.suffix + '.groups.json'),
                   {'file_hash': digest(source), 'groups': origins})
    finally:
        new.close()
        if old is not None: old.close()


def _translate(value, stage, base):
    if isinstance(value, dict): return {k:_translate(v,stage,base) for k,v in value.items()}
    if isinstance(value, list): return [_translate(v,stage,base) for v in value]
    if isinstance(value, str): return value.replace(str(stage),str(base))
    return value


def _file_paths(value):
    if isinstance(value, dict):
        return {p for child in value.values() for p in _file_paths(child)}
    if isinstance(value, (list, tuple)):
        return {p for child in value for p in _file_paths(child)}
    if isinstance(value, str):
        return {value}
    return set()


def _without_paths(value, failed):
    if isinstance(value, dict):
        return {k:_without_paths(v,failed) for k,v in value.items()
                if not isinstance(v,str) or v not in failed}
    return value


def run_transaction(input_path, base_dir, mapping_path, bank_dir, log, options, runner):
    from .options import build_batch_layout_from_options, requires_bank_data, requires_bonus_data, requires_co_data, requires_shared_expense_data
    from .banking import _find_bank_files
    base = Path(base_dir).resolve()
    with workspace_lock(base):
        recover_pending(base)
        layout = build_batch_layout_from_options(str(base), options)
        run_id = datetime.now().strftime('%Y%m%d-%H%M%S') + '-' + uuid.uuid4().hex[:8]
        history = Path(layout.archive_root) / '运行历史' / run_id
        stage = history / 'staging'; stage.mkdir(parents=True)
        inputs = [Path(input_path), Path(mapping_path or layout.mapping_path), Path(layout.timesheet_path),
                  base.parent / '模版性文件' / '2060总账凭证导入-实际202602薪酬.XLS']
        if requires_bank_data(options): inputs += [Path(p) for p in _find_bank_files(bank_dir or layout.bank_dir)]
        if requires_bonus_data(options): inputs.append(Path(layout.bonus_path))
        if requires_co_data(options): inputs.append(Path(layout.co_workorder_path))
        if requires_shared_expense_data(options): inputs.append(Path(layout.shared_expense_path))
        missing = [str(p) for p in inputs if not p.is_file()]
        if missing: raise FileNotFoundError('缺少本次必需输入：' + '；'.join(missing))
        hashes = {str(p.resolve()):digest(p) for p in inputs}
        bank_files_before = set(_find_bank_files(bank_dir or layout.bank_dir)) if requires_bank_data(options) else set()
        state = {'run_id':run_id,'status':'generating','inputs':hashes,'companies':options.companies,'vouchers':options.vouchers}
        write_json(history/'run.json',state)
        shutil.copy2(layout.timesheet_path, history/'工时输入快照.xlsx')
        token = _targets.set(OutputTargets(base,stage))
        try:
            result = runner(input_path,str(base),mapping_path,bank_dir,log,run_options=options)
        except Exception as exc:
            state.update(status='failed',error=str(exc)); write_json(history/'run.json',state)
            raise
        finally: _targets.reset(token)
        try:
            for key, expected in hashes.items():
                if not Path(key).is_file() or digest(key) != expected:
                    raise ValueError(f'运行期间输入已变化，请重新预检：{key}')
            if requires_bank_data(options) and set(_find_bank_files(bank_dir or layout.bank_dir)) != bank_files_before:
                raise ValueError('运行期间银行流水文件清单已变化，请重新预检')
            for source in stage.rglob('*.XLS'):
                merge_voucher(source, base/source.relative_to(stage), options, run_id)
            # Generated Markdown must link to published files, never staging paths.
            for source in stage.rglob('*.md'):
                with open(source,encoding='utf-8') as stream: content=stream.read()
                with open(source,'w',encoding='utf-8') as stream: stream.write(content.replace(str(stage),str(base)))
            all_optional = [p for p in stage.rglob('*') if p.is_file() and
                        ('07_重点税源采集信息' in p.parts or p.name=='实发薪酬表.xlsx')]
            accepted = _file_paths(result)
            optional = [p for p in all_optional if str(p) in accepted]
            required = [p for p in stage.rglob('*') if p.is_file() and p not in all_optional]
            publish_files(base,history,stage,required,'required')
            state.update(status='required_committed', result=_translate(result,stage,base))
            write_json(history/'run.json',state)
        except Exception as exc:
            state.update(status='needs_recovery' if '待恢复' in str(exc) else 'failed',error=str(exc))
            write_json(history/'run.json',state); raise
        optional_errors = list(result.get('optional_errors', []))
        optional_status = {}
        failed_paths = set()
        recovery_required = False
        for i, source in enumerate(optional):
            target = base / source.relative_to(stage)
            sidecar = source.with_suffix(source.suffix + '.status.json')
            write_json(sidecar, {'run_id':run_id, 'file_hash':digest(source)})
            try:
                publish_files(base,history,stage,[source,sidecar],f'optional-{i}')
                optional_status[str(target)] = {'status':'passed','run_id':run_id}
            except Exception as exc:
                failed_paths.add(str(source))
                optional_errors.append(f'{source.name} 本次未更新：{exc}')
                optional_status[str(target)] = {'status':'not_updated','error':str(exc)}
                old_status=target.with_suffix(target.suffix+'.status.json')
                if old_status.is_file():
                    with open(old_status,encoding='utf-8') as stream:
                        optional_status[str(target)]['last_success']=json.load(stream)
                if '待恢复' in str(exc):
                    recovery_required = True
                    break
        for source in all_optional:
            if str(base/source.relative_to(stage)) not in optional_status:
                failed_paths.add(str(source))
                optional_status[str(base/source.relative_to(stage))] = {'status':'not_updated'}
        from .options import COMPANY_NAME_TO_CODE
        expected_optional = [base/'03_运行输出/实发薪酬表.xlsx'] + [
            Path(layout.company_output_root(company))/'07_重点税源采集信息'/
            f'{COMPANY_NAME_TO_CODE[company]}重点税源采集信息{layout.batch_code}.xlsx'
            for company in options.companies]
        for target in expected_optional:
            if str(target) not in optional_status:
                optional_status[str(target)] = {'status':'not_updated'}
            if optional_status[str(target)]['status'] == 'not_updated':
                prior = target.with_suffix(target.suffix+'.status.json')
                if prior.is_file():
                    with open(prior,encoding='utf-8') as stream:
                        optional_status[str(target)]['last_success'] = json.load(stream)
                elif target.is_file():
                    optional_status[str(target)]['last_success'] = {
                        'source':'历史文件导入基线','file_hash':digest(target)}
        result = _without_paths(result, failed_paths)
        result = _translate(result,stage,base)
        result.update(run_id=run_id,history_path=str(history),optional_errors=optional_errors,
                      optional_status=optional_status,
                      status='needs_recovery' if recovery_required else
                      'success_with_warnings' if optional_errors else 'success')
        # Rewrite the human-facing output list after optional publication, so a
        # retained old report is never described as a successful new output.
        manifest = result.get('artifact_paths', {}).get('output_manifest')
        if manifest and not recovery_required:
            staged_manifest = stage / Path(manifest).relative_to(base)
            lines = [f'# 本次发布结果 {run_id}', '', f'- 状态：{result["status"]}',
                     f'- 历史记录：{history}', '', '## 本次凭证与工资整理', '']
            for field in ('output_paths', 'voucher_paths'):
                lines += [f'- {label}：{path}' for label,path in result.get(field,{}).items()]
            lines += ['', '## 附属报表', '']
            lines += [f'- {path}：{detail["status"]}' for path,detail in optional_status.items()]
            lines += [f'- 本次未更新：{error}' for error in optional_errors]
            lines += ['', '凭证分组来源及无数据状态见各文件旁的 .groups.json。']
            with open(staged_manifest,'w',encoding='utf-8') as stream:stream.write('\n'.join(lines)+'\n')
            try:
                publish_files(base,history,stage,[staged_manifest],'manifest')
            except Exception as exc:
                optional_errors.append(f'最终输出清单未更新，请以本次 run.json 为准：{exc}')
                result['status']='needs_recovery' if '待恢复' in str(exc) else 'success_with_warnings'
        state.update(status=result['status'],optional_errors=optional_errors,result=result)
        write_json(history/'run.json',state)
        write_json(Path(layout.archive_root)/'最新运行.json',state)
        log(f'运行编号：{run_id}；历史记录：{history}', 'ok')
        return result
