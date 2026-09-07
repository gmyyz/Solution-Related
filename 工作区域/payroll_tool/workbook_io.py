"""Owned workbook resources; file handles are released even if parsing fails."""
from contextvars import ContextVar
from functools import wraps
from io import BytesIO
from os import PathLike
from openpyxl import load_workbook as _load

_owned = ContextVar('payroll_owned_workbooks', default=None)


def load_workbook(filename, *args, **kwargs):
    # Detach the source handle before parsing. This also supports the historical
    # OOXML templates carrying an .XLS extension.
    if isinstance(filename, (str, bytes, PathLike)):
        with open(filename, 'rb') as stream:
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
