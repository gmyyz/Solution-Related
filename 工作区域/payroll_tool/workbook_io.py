"""Owned workbook resources; file handles are released even if parsing fails."""
from contextvars import ContextVar
from functools import wraps
from io import BytesIO
import os
from os import PathLike
from pathlib import Path
import shutil
import tempfile
from openpyxl import load_workbook as _load

_owned = ContextVar('payroll_owned_workbooks', default=None)
_OLE_SIGNATURE = bytes.fromhex('D0CF11E0A1B11AE1')
_OOXML_SIGNATURE = b'PK\x03\x04'


def _file_signature(path):
    with open(path, 'rb') as stream:
        return stream.read(8)


def is_strict_xls(path):
    """Return whether a workbook is a real Excel 97-2003 OLE/BIFF file."""
    return _file_signature(path) == _OLE_SIGNATURE


def _convert_with_excel(source, destination, file_format):
    """Use desktop Excel for a real format conversion, not an extension rename."""
    if os.name != 'nt':
        raise RuntimeError('生成严格 XLS 需要在 Windows 上安装 Microsoft Excel')
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError('生成严格 XLS 需要安装 pywin32，请重新执行依赖安装') from exc

    source = Path(source).resolve()
    destination = Path(destination).resolve()
    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(
            str(source),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
        )
        workbook.CheckCompatibility = False
        workbook.SaveAs(str(destination), FileFormat=file_format)
    except Exception as exc:
        raise RuntimeError(f'Excel 格式转换失败：{source.name}') from exc
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def convert_to_strict_xls(path):
    """Atomically replace an OOXML workbook with a real Excel 97-2003 XLS."""
    path = Path(path)
    if is_strict_xls(path):
        return
    if _file_signature(path)[:4] != _OOXML_SIGNATURE:
        raise ValueError(f'无法识别待转换的 Excel 文件格式：{path}')

    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='.xls-convert-', dir=path.parent) as temp_dir:
        temp_dir = Path(temp_dir)
        source = temp_dir / 'source.xlsx'
        converted = temp_dir / 'converted.xls'
        shutil.copy2(path, source)
        _convert_with_excel(source, converted, 56)  # xlExcel8
        if not converted.is_file() or not is_strict_xls(converted):
            raise RuntimeError(f'Excel 未生成严格 XLS 文件：{path.name}')
        os.replace(converted, path)


def load_workbook(filename, *args, **kwargs):
    # Detach the source handle before parsing. This also supports the historical
    # OOXML templates carrying an .XLS extension.
    if isinstance(filename, (str, bytes, PathLike)):
        path = Path(filename)
        if is_strict_xls(path):
            with tempfile.TemporaryDirectory(prefix='.xls-read-') as temp_dir:
                converted = Path(temp_dir) / 'workbook.xlsx'
                _convert_with_excel(path, converted, 51)  # xlOpenXMLWorkbook
                with open(converted, 'rb') as stream:
                    filename = BytesIO(stream.read())
        else:
            with open(path, 'rb') as stream:
                filename = BytesIO(stream.read())
    workbook = _load(filename, *args, **kwargs)
    resources = _owned.get()
    if resources is not None:
        resources.append(workbook)
    return workbook


def workbook_scope(function):
    @wraps(function)
    def wrapped(*args, **kwargs):
        resources = []
        token = _owned.set(resources)
        try:
            return function(*args, **kwargs)
        finally:
            try:
                for workbook in reversed(resources):
                    workbook.close()
            finally:
                _owned.reset(token)
    return wrapped
