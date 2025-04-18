import xlpro.run_server
import sys
import threading
from win32com.client import Dispatch
import win32com.client
from win32typelibs import excel as xl

def main():

    xlpro.run_server.main()

    # t1 = threading.Thread(target=xlpro.run_server.main)
    # t1.start()
    # t1.join()

    # def thread2_func():
    #     from xlpro.server import xlproServer
    #     xlproapp:xlproServer = Dispatch('{122BB48A-57EF-4775-A28C-3F71ED0D02A7}')
    #     xlapp:xl._Application = Dispatch("Excel.Application")

    #     xlproapp.register_and_configure_wb_workspace(xlapp.ActiveWorkbook)




if __name__ == "__main__":
    main()
