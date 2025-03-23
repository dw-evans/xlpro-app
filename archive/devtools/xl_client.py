import win32com.client

from win32typelibs import excel as xl

xlapp:xl._Application = win32com.client.Dispatch("Excel.Application")

rng = xlapp.Selection


def test_func(a:int, b:float, c, caller) -> float:
    pass


pass