import pythoncom
import win32com.client
import numpy as np
# from win32typelibs import excel as xl


def add_numbers(a, b):
    return a + b + 100

def make_array(n:int, m:int):
    return np.random.random((n, m))

# from scripts import myfunc

class xlproServer:
    _public_methods_ = [
        'executeFunction',
        'initxlpro',
        'getpid',
        'testImports',
        'printTest',
    ]
    _reg_progid_ = 'xlproServer.Application'
    # _reg_clsid_ = pythoncom.CreateGuid()
    _reg_clsid_ = '{5C9DF910-6E2E-423D-9615-A28F274F35F8}'

    def executeFunction(self, func_name, *args):
        try:
            # Dynamically call the Python function
            func = globals().get(func_name)
            if func:
                return func(*args)
            else:
                return f"Function {func_name} not found."
        except Exception as e:
            return str(e)
        
    def initxlpro(self, caller:"xl._Workbook"):
        return
        xlapp:xl._Application = win32com.client.Dispatch("Excel.Application")
        py_caller = xlapp.Workbooks(caller.Name)

    def getpid(self):
        import os
        return os.getpid()
    

    def testImports(self):
        from win32typelibs import excel
        from utils import myfunc

        return "success"
        pass
    
    def printTest(self):
        return "printTest returned"

    

if __name__ == '__main__':
    # Register the COM server
    import win32com.server.register
    win32com.server.register.UseCommandLine(xlproServer)