import pythoncom
import win32com.client
from win32com.client import Dispatch
import numpy as np
# Type that can be used for isinstance checks.
from win32com.client import _PyIDispatchType
import typing
import logging
from pathlib import Path
import os
import utils
import importlib

wd = Path(__file__).parent
# Configure logging to write to a file
logging.basicConfig(
    filename= wd / 'my_log_file.log',   # The file where logs will be saved
    level=logging.DEBUG,          # The log level (DEBUG, INFO, WARNING, etc.)
    format='%(asctime)s - %(levelname)s - %(message)s',  # The format of log messages
    datefmt='%Y-%m-%d %H:%M:%S'    # The format of the date in log messages
)

def load_functions_from_file(module_name, file_path):
    """Dynamically import all functions from a Python file."""
    # Load the module dynamically
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    
    # Get all functions in the module
    functions = {
        name: getattr(module, name)
        for name in dir(module)
        if callable(getattr(module, name))
    }
    return functions

class xlproServer:
    _public_methods_ = [
        'execute_function',
        'executeFunction',
        'getpid',
        "register_functions_in_vba",
        "register_functions_in_self",
    ]
    _reg_progid_ = 'xlproServer.Application'
    # _reg_clsid_ = pythoncom.CreateGuid()
    _reg_clsid_ = '{5C9DF910-6E2E-423D-9615-A28F274F35F8}'

    def __init__(self, *args, **kwargs):
        """On initialization, we want to import all the appropriate modules
        """
        self._func_register:dict[str, typing.Callable] = {}
        pass

    def register_functions_in_self(self):
        self._func_register = load_functions_from_file(
            "xlpro_register", 
            wd / "xlpro_register.py",
        )
        return

    def register_functions_in_vba(self, thisworkbook):
        wb = Dispatch(thisworkbook)
        self.register_functions_in_self()
        funcs = [v for k,v in self._func_register.items()]
        utils.init_xlpro_vb_dynamic_component(wb, funcs)
        return

    def execute_function(self, func_name, *args) -> str:
        try:
            func = self._func_register.get(func_name)
            if func:
                return func(*args)
            else:
                return f"Function {func_name} not found."
        except Exception as e:
            return str(e)
        
    def executeFunction(self, *args, **kwargs):
        return self.execute_function(*args, **kwargs)

    def getpid(self):
        return os.getpid()
    
    def com_obj_test(self, val:int, range:object):
        if isinstance(range, _PyIDispatchType):
            obj_new = Dispatch(range)
        return "com_obj_test complete"


from functools import wraps
    
def caller(func):
    @wraps(func)
    def wrapper(func):

        pass
    return wrapper

if __name__ == '__main__':
    # Register the COM server
    import win32com.server.register
    win32com.server.register.UseCommandLine(xlproServer)