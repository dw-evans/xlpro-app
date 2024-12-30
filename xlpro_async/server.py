from __future__ import annotations
import typing
import logging
from pathlib import Path
import os

import threading
import queue

import pythoncom
import win32com.client
from win32com.client import _PyIDispatchType # Type that can be used for isinstance checks.
import win32com.server.util
import win32com.server.policy

import datetime
import numpy as np
import pywintypes

import utils

from win32typelibs import excel as xl

wd = Path(__file__).parent


logger = logging.getLogger(__name__)

class uuid:
    count = 0
    def __new__(cls):
        cls.count += 1
        return cls.count


class xlproServerAsync:
    _public_methods_ = [
        'getpid',
        "add_data",
        "register_functions_in_vba",
        "register_functions_in_self",
        "execute_function",
        "execute_function_async",
    ]
    _reg_progid_ = 'xlproServerAsync.Application'
    _reg_clsid_ = '{122BB48A-57EF-4775-A28C-3F71ED0D02A7}'

    _instance = None  # Singleton instance
    _instance_initialized = False

    def __new__(cls, *args, **kwargs):
        if cls._instance is None:
            cls._instance = super().__new__(cls, *args, **kwargs)
            # logger.log(f"__new__ creating new instance")
            # logger.log(f"__new__ creating new instance. PID: {os.getpid()}, MEMID: {hex(id(cls._instance))}")
        return cls._instance

    def __init__(self):
        if not xlproServerAsync._instance_initialized:
            logger.info(f"__init__ initializing new instance")
            self._data = []
            xlproServerAsync._instance_initialized = True

            # XXX something errors in the init function and closes the server
            # Debugger is not cut out to catch these errors.

            self._func_register = {}
            # self.register_functions_in_self()

            # store the results uuid: result
            self._func_hash_results_map = {}
            self._func_hash_results_map_lock = threading.Lock()
            self._func_hash_result_iscomplete_map = {}

            # maps the uid to the caller and function hash
            self._func_hash_to_caller_map = {}

            # map the uuid to the working thread
            # self._func_hash_to_working_thread_map = {}

            self._result_queue = queue.Queue() # stores the results as they come in
            self._recalculate_queue = queue.Queue() # stores the cells that need to be recalculated.

            # results manager handles processing the queue of results as they come in
            # and signalling to the client manager to execute commands.
            self._results_manager_thread:ResultsManager = ResultsManager(server=self)
            self._results_manager_thread.start()

            # self._wb:xl._Workbook = None
            # self._client_manager:ClientManager = None
            self._client_manager = ClientManager(server=self)
            self._client_manager.start()


        else:
            logger.info("__init__ called however the singleton already exists and has been initialized.")
        return
    
    def register_workbook(self, thiswb):
        """Register the workbook client of the xlpro server."""
        self._wb = win32com.client.Dispatch(thiswb)
        # XXX - I am concerned about accessing this from a different thread
        # self._client_manager = ClientManager(self)

    def getpid(self):
        return os.getpid()
    
    def add_data(self, data):
        self._data.append(data)
        logger.debug(f"add_data caller pid: {os.getpid()}")
        logger.debug(f"add_data data: {data}")
        logger.debug(f"add_data self memid: {hex(id(self))}")
        logger.debug(f"self._data = {self._data}")
        return f"self._data is now {self._data}"
    
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

    def execute_function_async(self, caller, func_name, *args):
    # def execute_function_async(self, func_name, *args) -> str:
        try:
            # fetch the function that is being called.
            func = self._func_register.get(func_name)

            if not func:
                return f"Function {func_name} not found."
            
            uid = utils.hash_function_call(func, *args)

            # if the hash of the function has been marked complete, return that
            # avoid computation
            if uid in self._func_hash_result_iscomplete_map.keys():
                return self._func_hash_results_map[uid]

            # Get the caller and convert it to a range to store.
            # Range object needed to handle dynamic addresses.
            map_to_caller = True
            if map_to_caller:
                caller_dispatch = win32com.client.Dispatch(caller)
                # marshal the thread for threaded use
                caller_marshal = pythoncom.CoMarshalInterThreadInterfaceInStream(
                    pythoncom.IID_IDispatch,
                    # caller_dispatch._oleobj_
                    caller_dispatch,
                )
                self._func_hash_to_caller_map[uid] = caller_marshal
                pass

            # create the daemon thread to compute the result
            t = self._create_and_register_async_worker(uid, func, args, kwargs={})

            self._func_hash_results_map[uid] = f"Promise<{uid}>"
            self._func_hash_result_iscomplete_map[uid] = False

            t.start()

            # return the result (will be incomplete)
            return self._func_hash_results_map[uid]

        except Exception as e:
            return str(e)
        
    def _generate_uid(self):
        return uuid()
        
    def _create_and_register_async_worker(self, uid, func, args, kwargs):
        def func_wrapper():
            ret = func(*args, **kwargs)
            # add to queue and wake manager.
            self._result_queue.put((uid, ret))
            self._func_hash_result_iscomplete_map[uid] = True
            self._results_manager_thread.wake()

        t = threading.Thread(target=func_wrapper, daemon=True)

        # self._func_hash_to_working_thread_map[uid] = t

        return t

        
import time
 
class ResultsManager:
    def __init__(self, server:xlproServerAsync):
        self._stop_event = threading.Event()
        self._wake_event = threading.Event()
        self._thread = threading.Thread(target=self._watch, daemon=True)
        self._server = server

    def start(self):
        self._thread.start()
        self._thread
        logger.info(f"ResultsManager thread started: tid: {threading.get_native_id()}")

    def stop(self):
        self._stop_event.set()
        self._wake_event.set()  # Wake up if sleeping
        self._thread.join()

    def wake(self):
        self._wake_event.set()

    def _watch(self):
        """When awakened, retrieves the result queue and 
        sends to the server cache."""
        while not self._stop_event.is_set():
            logger.info("ResultsManager thread is waiting for events or timeout...")
            self._wake_event.wait(timeout=5)
            if self._wake_event.is_set():
                self._wake_event.clear()
                logger.info("ResultsManager thread woke up for an event!")
               
                # Process the whole queue once woken up
                # XXX - could replace while trye with while not stop event.
                while True:
                    try:
                        uid, val = self._server._result_queue.get()
                    except queue.Empty:
                        break
                    pass
                    # add the result to the server cache and signal update to client manager
                    with self._server._func_hash_results_map_lock:
                        self._server._func_hash_results_map[uid] = val
                        self._server._recalculate_queue.put((
                            uid, 
                            self._server._func_hash_to_caller_map[uid]
                        ))
                        self._server._client_manager.wake()
                        pass
                        # self._server._client_manager._recalculate_cell(self._server._func_hash_to_caller_map[uid])

                    time.sleep(0.01) # fairness sleep

            if self._stop_event.is_set():
                break  # Exit if stop event is set


class ClientManager:
    """Hooks to the excel client so we can trigger events"""

    def __init__(self, server):
        self._server = server
        pass

    def __init__(self, server:xlproServerAsync):
        self._stop_event = threading.Event()
        self._wake_event = threading.Event()
        self._thread = threading.Thread(target=self._watch, daemon=True)
        self._server = server

    def start(self):
        self._thread.start()
        self._thread
        logger.info(f"ClientManager thread started: tid: {threading.get_native_id()}")

    def stop(self):
        self._stop_event.set()
        self._wake_event.set()  # Wake up if sleeping
        self._thread.join()

    def wake(self):
        self._wake_event.set()

    def _recalculate_cell(self, cell:xl.Range):
        cell.Calculate()

    def comarshal_range(self, range_marshal) -> xl.Range:
        range_pyidispatch = pythoncom.CoGetInterfaceAndReleaseStream(
            range_marshal, pythoncom.IID_IDispatch,
        )
        range_dispatch = win32com.client.Dispatch(range_pyidispatch)
        return range_dispatch

    def _process_queue(self):
        """Process the queue at the current point in time. Any failed attempts get
        added back into the queue.
        Probably need to 
        """
        # XXX - Warning that qsize() is not thread safe. Shouldn't be an issue.
        queue_length = self._server._recalculate_queue.qsize()
        for _ in range(queue_length):
            try:
                uid, caller_marshal = self._server._recalculate_queue.get()
            except queue.Empty:
                return
            pass

            # if we fail to dispatch the range, it was probably deleted.
            # therefore we don't need to replace it in the queue
            replace_in_queue = True
            try:
                caller_dispatch = self.comarshal_range(caller_marshal)
            except pywintypes.com_error as e:
                if VBErrorConverter(e) == VBError.xlObjectRequired:
                    logger.error(f"ClientManager _process_queue() COM Exception: '{e}', uid: '{uid}'")
                    replace_in_queue = False
                else:
                    logger.error(f"ClientManager _process_queue() COM Exception: '{e}', uid: '{uid}'")


            # Updates will fail if excel application has a dialogue open for example.
            # Give it back to the queue to handle later.
            try:
                caller_dispatch.__getattr__("Formula2")
                caller_dispatch.Formula2 = caller_dispatch.Formula2
                # XXX - todo - consider removing the uid after this call.
            except (AttributeError, pywintypes.com_error) as e:
                # At this point 
                # XXX - Marshalling the caller back to the pool (not that we plan on using
                # this object elsewhere)
                logger.info(f"Could not recalculate caller_dispatch for uid: '{uid}'")
                caller_marshalled_back = pythoncom.CoMarshalInterThreadInterfaceInStream(
                    pythoncom.IID_IDispatch, 
                    caller_dispatch,
                )
                if replace_in_queue:
                    self._server._recalculate_queue.put((uid, caller_marshalled_back))

            time.sleep(0.01) # fairness sleep


    def _watch(self):
        """When awakened, retrieves the result queue and sends to the server cache.
        """
        pythoncom.CoInitialize()

        while not self._stop_event.is_set():
            # wait for ten seconds for an event, otherwise do a queue process to check for 
            # outstanding tasks.
            logger.info("ClientManager thread is waiting for events or timeout...")
            self._wake_event.wait(timeout=10)
            if self._wake_event.is_set():
                self._wake_event.clear()
                logger.info("ClientManager thread woke up for an event!")
            self._process_queue()

            # XXX - todo - could check for the exit event during the loop also.
            if self._stop_event.is_set():
                break

        pythoncom.CoUninitialize()


"""
-2147418111: call was rejected by callee
-2146827864: object required

"""
class VBError:
    xlCallRejectedByCallee = 1  # Call Rejected by Callee (Custom errors)
    xlObjectRequired = 424  # Object required (when an object is not available or specified)

    # hte following are generated by chatgpt

    xlNotImplemented = -2147467263  # Method or property not implemented
    xlInvalidMethod = 438  # Object doesn't support this property or method
    xlInvalidProcedureCall = 5  # Invalid procedure call or argument
    xlInvalidArgument = 9999  # Invalid argument or parameter provided

    xlOutOfMemory = -2147024882  # Insufficient memory to complete the operation
    xlInvalidReference = 91  # Object variable or With block variable not set
    xlTypeMismatch = 13  # Type mismatch error (incorrect data type)
    xlAutomationError = -2147352567  # Generic Automation error, usually related to COM failures

    xlFileNotFound = 53  # File not found (usually happens when trying to open a file)
    xlPermissionDenied = 70  # Permission denied (file access issue)
    xlFileAlreadyOpen = 101  # The file is already open in Excel
    xlFileFormatNotSupported = 1004  # The file format is not supported by Excel

    xlObjectNotFound = 3265  # Object not found in the collection (can happen when working with ranges, cells, etc.)
    xlRangeNotFound = 1004  # Range cannot be found (invalid range or sheet)
    xlApplicationDefinedError = 32755  # Application-defined or object-defined error

    xlInvalidCellReference = 1004  # Invalid cell reference (usually happens when the cell cannot be accessed)
    xlFormulaParseError = 1004  # Formula parsing error (when a formula is incorrectly written)
    xlCannotOpenWorkbook = 1004  # Cannot open the workbook (file may not exist or is not accessible)

    xlInvalidRange = 1004  # Invalid range used
    xlSheetNotFound = 9  # Sheet not found (referring to a non-existent sheet)
    xlRangeNotValid = 1004  # Range is invalid (referring to an invalid range)
    xlActiveCellRequired = 1004  # The active cell cannot be referenced (e.g., no selection)
    
    xlInvalidObject = 13  # Invalid object type (often occurs with wrong object types passed to methods)
    xlInvalidState = 3201  # Invalid operation state (e.g., when an operation cannot be performed at the moment)

    xlUserDefinedError = 9999  # Custom user-defined error (error raised explicitly by VBA code)



class VBErrorConverter:
    def __new__(cls, e):
        if isinstance(e, pythoncom.com_error):
            e_hresult = e.hresult
        elif isinstance(e, int):
            e_hresult = e
        else:
            raise TypeError(f"Type {type(e)} is not allowed")
        return VBErrorConverter.convert_hresult_to_vba_codes(e_hresult)[1]
        
    def convert_hresult_to_32b_signed(hresult):
        """Convert a signed 32-bit HRESULT to its unsigned hexadecimal representation."""
        # Ensure the number is treated as a 32-bit unsigned integer
        hex_value = hresult & 0xFFFFFFFF
        return hex_value
    @staticmethod
    def convert_hresult_to_vba_codes(hresult):
        """Returns the Facility Code, Error Code of an hresult"""
        unsigned_hresult = VBErrorConverter.convert_hresult_to_32b_signed(hresult)
        return ((unsigned_hresult >> 16) & 0xFFFF, unsigned_hresult & 0xFFFF)


from functools import wraps
def exception_return_wrapper(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs) 
        except Exception as e:
            return e
        
    return wrapper


if __name__ == '__main__':
    ...

    from win32com.client import Dispatch


    xlapp:xl._Application = Dispatch("Excel.Application")
    xlapp.Visible = True
    rng:xl.Range = xlapp.Selection
    pass
    e1 = exception_return_wrapper(lambda: rng.Address)()
    e2 = exception_return_wrapper(lambda: rng.Address)()
    e3 = exception_return_wrapper(lambda: rng.Address)()

    pass



    pass