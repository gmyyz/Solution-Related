"""Payroll business package. GUI is loaded only when explicitly requested."""

def __getattr__(name):
    if name in {'MainWindow', 'main'}:
        from . import gui
        return getattr(gui, name)
    raise AttributeError(name)
