from __future__ import annotations
import typing
import logging
from pathlib import Path
import os

import threading

import pythoncom
import win32com.client
from win32com.client import _PyIDispatchType # Type that can be used for isinstance checks.
import win32com.server.util
import win32com.server.policy

import datetime
import numpy as np

import utils

wd = Path(__file__).parent

logger = logging.getLogger(__name__)

class xlproServerAsync:
    _public_methods_ = [
        'getpid',
        "p",
        "add_data",
        "execute_function",
        "register_functions_in_vba",
        "register_functions_in_self",
        "run_async_worker",
    ]
    _reg_progid_ = 'xlproServerAsync.Application'
    _reg_clsid_ = '{122BB48A-57EF-4775-A28C-3F71ED0D02A7}'

    def __init__(self):
        self._data = []
        self._func_register:dict[str, typing.Callable] = {}
        return
    
    def getpid(self):
        return os.getpid()
    
    def p(self):
        print(f"{datetime.datetime.now().strftime('%d/%m/%Y, %H:%M:%S')} - {os.getpid()} - {threading.get_ident()}")

    def add_data(self, data):
        self._data.append(data)
        return f"self._data is now {self._data}"

    def __init__(self, *args, **kwargs):
        """On initialization, we want to import all the appropriate modules
        """
        pass

    def register_functions_in_self(self):
        self._func_register = utils.load_functions_from_file(
            "xlpro_register", 
            wd / "xlpro_register.py",
        )
        return

    def register_functions_in_vba(self, thisworkbook):
        wb = win32com.client.Dispatch(thisworkbook)
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

    def run_async_worker(self):
        import threading

        def func():
            import time
            logger.debug("waiting for 5 seconds...")
            time.sleep(5.0)
            logger.debug("finished waiting")
            return

        new_thread = threading.Thread(target=func, daemon=True)
        new_thread

        # Creates a thread that returns waiting initially, then once ready, issues a callback 
        # to excel to fetch the results.
        # 
        # 
        # 

        ...







if __name__ == '__main__':
    # import win32com.server.register
    # win32com.server.register.UseCommandLine(xlproServerAsync)