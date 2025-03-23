import win32com.client 

import sys
from win32typelibs import excel as xl

from pathlib import Path

wd = Path(__file__).parent.parent

xlapp:xl._Application
try: 
    xlapp = win32com.client.GetActiveObject("Excel.Application")
    try:
        xlapp.Workbooks("workbook.xlsm").Save()
        xlapp.Workbooks("workbook.xlsm").Close()
    except:
        pass
    xlapp.Quit()
    xlapp = None
except:
    pass

import time
time.sleep(0.2)

xlapp = win32com.client.Dispatch("Excel.Application")
xlapp.Visible = True
try:
    xlapp.Workbooks("PERSONAL.xlsb").Close(True)
except:
    pass
wb = xlapp.Workbooks.Open(str((wd / "template/workbook.xlsm").resolve()))
wb.Activate()

xlapp.VBE.MainWindow.Visible = True

pass




