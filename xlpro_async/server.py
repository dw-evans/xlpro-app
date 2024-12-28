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
        # "execute_function",
        # "register_functions_in_vba",
        # "register_functions_in_self",
        # "run_async_worker",
    ]
    _reg_progid_ = 'xlproServerAsync.Application'
    _reg_clsid_ = '{122BB48A-57EF-4775-A28C-3F71ED0D02A7}'

    _instance = None  # Singleton instance

    def __new__(cls, *args, **kwargs):
        if cls._instance is None:
            cls._instance = super().__new__(cls, *args, **kwargs)
            cls._instance._data = []  # Initialize state
            # cls._instance._func_register:dict[str, typing.Callable] = {}
            # cls._instance._func_register = {}
        return cls._instance

    def __init__(self):
        return
    
    def getpid(self):
        return os.getpid()
    
    def p(self):
        print(f"{datetime.datetime.now().strftime('%d/%m/%Y, %H:%M:%S')} - {os.getpid()} - {threading.get_ident()}")

    def add_data(self, data):
        self._data.append(data)
        logger.debug(f"add_data caller pid: {os.getpid()}")
        logger.debug(f"add_data data: {data}")
        logger.debug(f"add_data self memid: {hex(id(self))}")
        logger.debug(f"self._data = {self._data}")
        return f"self._data is now {self._data}"
    


    # def register_functions_in_self(self):
    #     self._func_register = utils.load_functions_from_file(
    #         "xlpro_register", 
    #         wd / "xlpro_register.py",
    #     )
    #     return

    # def register_functions_in_vba(self, thisworkbook):
    #     wb = win32com.client.Dispatch(thisworkbook)
    #     self.register_functions_in_self()
    #     funcs = [v for k,v in self._func_register.items()]
    #     utils.init_xlpro_vb_dynamic_component(wb, funcs)
    #     return

    # def execute_function(self, func_name, *args) -> str:
    #     try:
    #         func = self._func_register.get(func_name)
    #         if func:
    #             return func(*args)
    #         else:
    #             return f"Function {func_name} not found."
    #     except Exception as e:
    #         return str(e)

    # def run_async_worker(self):
    #     import threading

    #     def func():
    #         import time
    #         logger.debug("waiting for 5 seconds...")
    #         time.sleep(5.0)
    #         logger.debug("finished waiting")
    #         return

    #     new_thread = threading.Thread(target=func, daemon=True)
    #     new_thread




"""
Callback structure

write async results to a register
when a function completes, it has data which signals which cells (callers) need to be recomputed

# when a function is called, the output is given a unique id in the cached results.
# the cache is filled with the temporary placeholder results until it is complete
# once the uid has been flagged complete, it looks up
# the uid can cache a set of arguments func(a, b, c) so we can look up any existing ones in the past (potentially)
# This caching of results would require a cache clear option to the formulas, maybe with an optional defaulted to no
# But at the same time we might want to prevent any big calculations from being executed unknowingly 

# Is there an opportunity to have some helper methods that can see which cells are calling the xlpro functions

# Oh and by the way we need to nest everything down a level so each workbook has its own memory space if we are doing
# the persistent memory option

# How can we issue the callback
# if a function knows its caller, it can issue a Range.Recalculate callback pretty easily.

# All of this relies on recalculate doing a cache lookup first. and if it misses, deferring the calculation to a background
# python process

# maps the function uuid to the live result
results_register = {
    # "uuid": <T>result,
    12321453:
}

# look up the function to the uuid
function_register = {
    hash(func): uuid
}

caller_register = {
    workbook.sheet.a1: uuid
}

# an async filler function can be called in the meantime for each life function
# filler function
lambda repr, t: f"{repr} has been executing for {t} seconds"

# as these are seen as complete
completion_register = {
    uuid: True/False
}

# when a function is marked complete the called must be ordered to recalculate and points to
# a cached output of the function 

"""





if __name__ == '__main__':
    ...
    # import win32com.server.register
    # win32com.server.register.UseCommandLine(xlproServerAsync)