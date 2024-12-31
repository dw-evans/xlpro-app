from __future__ import annotations
import typing
import logging
from pathlib import Path
import os

import threading
import queue
import multiprocessing
import time

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

import matplotlib.figure


wd = Path(__file__).parent


logger = logging.getLogger(__name__)

class uuid:
    count = 0
    def __new__(cls):
        cls.count += 1
        return cls.count

def load_functions_from_register() -> dict:
    """Returns a dict of functions found in the xlpro registry module.
    name: function. Runs all the imports too :)
    """
    return utils.load_functions_from_file(
            "xlpro_register", 
            wd / "xlpro_register.py",
        )

def dispatch_args_preprocessor(func, args):
    """Dispatches the  b 
    """
    f_name, args_and_types, ret_type, _ = utils.get_function_signature(func)
    arg_names = [v0 for v0, v1 in args_and_types]
    new_args = args
    if "caller" in arg_names:
        idx = arg_names.index("caller")
        caller = args[idx]
        caller_stream = utils.comarshal_release_and_get_stream(caller)
        new_args[idx] = caller_stream
        # Caller type could be many things, likely just a Range.
        pass
    if "thiswb" in arg_names:
        idx = arg_names.index("thiswb")
        thiswb = args[idx]
        thiswb_stream = utils.comarshal_release_and_get_stream(thiswb)
        new_args[idx] = thiswb_stream
        pass
    return args


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

            self._func_hash_result_display_map = {} # the result to be displayed
            self._func_hash_result_display_map_lock = threading.Lock()

            self._func_hash_results_map = {} # the actual results
            self._func_hash_results_map_lock = threading.Lock()
            self._func_hash_result_iscomplete_map = {}

            self._func_hash_result_type = {}
            self._func_hash_result_type_lock = threading.Lock()

            # function types: standard, figure, ...image...? Anything else...
            # could just tag the results hash map, or create a dict that keeps the func hash and output type
            # self._func_hash_figure_map = {} # dict that matches func call hash to the output figure.

            # maps the uid to the caller and function hash
            self._func_hash_to_caller_map = {}
            self._func_hash_to_caller_map_lock = threading.Lock() # XXX - todo - not used currently

            # maps the live processes and threads in the background.
            # self._func_hash_to_thread_map = {} # XXX - not used
            # self._func_hash_to_process_map = {} # XXX - not used

            self._threaded_result_queue = queue.Queue() # stores the results as they come in
            self._recalculate_queue = queue.Queue() # stores the cells that need to be recalculated.

            self._multiprocessing_figure_result_queue = multiprocessing.Queue() # stores the multiprocessing results

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
        self._func_register = load_functions_from_register()
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

    def execute_function_async(self, result_type:ResultType, caller, func_name, *args):
        try:
            # fetch the function that is being called.
            func = self._func_register.get(func_name)

            if not func:
                return f"Function {func_name} not found."
            
            # 
            uid = utils.hash_function_call(func, args, kwargs={})
            # uid = utils.create_random_hash()

            # if the hash of the function has been marked complete, return that
            # avoid computation
            if uid in self._func_hash_result_iscomplete_map.keys():
                # return self._func_hash_results_map[uid]
                return self._func_hash_result_display_map[uid]

            # check the result_type is valid.
            if not result_type in ResultType.as_list():
                raise Exception(f"ResultType identifier is not valid, v={result_type}, valids={ResultType.as_list()}")
            self._func_hash_result_type[uid] = result_type

            # Get the caller and convert it to a range to store.
            # Range object needed to handle dynamic addresses.
            map_to_caller = True
            if map_to_caller:
                caller_dispatch = win32com.client.Dispatch(caller)
                # marshal the thread for threaded use
                caller_marshal = pythoncom.CoMarshalInterThreadInterfaceInStream(
                    pythoncom.IID_IDispatch,
                    caller_dispatch,
                )
                self._func_hash_to_caller_map[uid] = caller_marshal
                pass

            # self._func_hash_results_map[uid] = f"Promise<{uid}>"
            self._func_hash_result_display_map[uid] = f"Promise<{uid}>"
            self._func_hash_result_iscomplete_map[uid] = False

            if result_type == ResultType.default:
                # create the daemon thread to compute the result
                t = self._create_and_register_async_worker(uid, func, args, kwargs={})
                t.start()
            elif result_type == ResultType.figure:
                fig_generating_thread = FigureGeneratingThread(
                    uid=uid,
                    func_name=func_name,
                    args=args,
                    kwargs={},
                    return_value_queue=self._threaded_result_queue, 
                    return_event=self._results_manager_thread._wake_event,
                )
                fig_generating_thread.start()

            # return the result (will be incomplete)
            return self._func_hash_result_display_map[uid]

        except Exception as e:
            return str(e)
        
    def _generate_uid(self):
        return uuid()
        
    def _create_and_register_async_worker(self, uid, func, args, kwargs):
        
        def func_wrapper():

            # XXX - todo - possible candidate here to wrap this in a try-except block
            # so we can re-attempt failed functions.
            # ATM there is a likely attribute error
            # attempts = 0
            # max_attempts = 20
            # delay = 5.0
            # while attempts < 10:
            #     try:
            #         ret = func(*args, **kwargs)
            #     except Exception as e:
            #         logger.warning(f"Async worker thread failed to get a result. Retrying in {delay} seconds")
            #         ret = f"Failed Promise{attempts}/{max_attempts}<{uid}>"
            #         attempts += 1
            #         # add to queue and wake manager regardless. We are working on it in the background I think.
            #         # self._threaded_result_queue.put((uid, ret))
            #         # self._func_hash_result_iscomplete_map[uid] = True
            #         # self._results_manager_thread.wake()
            #         time.sleep(delay)

            ret = func(*args, **kwargs)

            # add to queue and wake manager.
            self._threaded_result_queue.put((uid, ret))
            self._func_hash_result_iscomplete_map[uid] = True
            self._results_manager_thread.wake()

        t = threading.Thread(target=func_wrapper, daemon=True)
        return t
    
    def _create_and_register_figure_generating_process(self, uid, func, args, kwargs):
        def func_wrapper():
            fig = func(*args, **kwargs)
            
            fig:matplotlib.figure.Figure
            exec('matplotlib.use("Agg")')
            fp = wd / ".xlpro" / "tmp" / f"{uid}.png"
            fp.mkdir(parents=True, exist_ok=True)
                
            fig.savefig(wd / ".xlpro" / "tmp" / f"{uid}",)
            fig.savefig(fp, dpi=600)
        
            # add to queue and wake manager.
            self._multiprocessing_figure_result_queue.put((uid, fp.resolve()))
            self._results_manager_thread.wake()

        global _MULTIPROCESS_FUNCTION
        _MULTIPROCESS_FUNCTION = func_wrapper

        p = multiprocessing.Process(target=_MULTIPROCESS_FUNCTION, daemon=True)
        return p

class FigureGeneratingThread:
    """Class which maintains a thread which waits for a process to finish."""
    def __init__(self, uid, func_name, args, kwargs, return_value_queue, return_event):
        self._stop_event = threading.Event()
        self._wake_event = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)

        self._return_event = return_event
        self._return_queue = return_value_queue        

        self._process:multiprocessing.Process = None
        self._multiprocess_queue = multiprocessing.Queue()

        self._uid = uid
        self._func_name = func_name
        self._args = args
        self._kwargs = kwargs


    def start(self):
        # start the watcher
        self._thread.start()

    def stop(self):
        self._stop_event.set()
        self._wake_event.set()  # Wake up if sleeping
        self._thread.join()

    def wake(self):
        self._wake_event.set()

    def _create_process(self) -> multiprocessing.Process:

        p = multiprocessing.Process(
            target=figure_process_func, 
            daemon=True,
            kwargs={
                "uid": self._uid,
                "func_name": self._func_name,
                "args": self._args,
                "kwargs": self._kwargs,
                "queue": self._multiprocess_queue,
            }
        )
        return p
    
    def _run(self): 
        # logger.info(f"Starting FigureGeneratingThread process within thread, {threading.get_native_id()}")
        p = self._create_process()

        p.start()
        p.join() # once joined, we can check the queue

        # get the uid and value (figure path in this case)
        uid, val = self._multiprocess_queue.get()

        # send the uid/value combo to the return queue
        # the return queue should always be the value return queue 
        self._return_queue.put((uid, val))
        self._return_event.set()
        pass
        

def figure_process_func(uid, func_name, args, kwargs, queue):
    func = setup_scope_and_get_function(func_name)
    fp = wd / ".xlpro" / "tmp" / f"{uid}.png"
    fp.parent.mkdir(parents=True, exist_ok=True)
    fig:matplotlib.figure.Figure = func(*args, **kwargs)
    size_inches = np.array(fig.get_size_inches())
    fig.savefig(fp, dpi=600)
    queue.put((uid, (fp, size_inches)))
    pass

def setup_scope_and_get_function(func_name):
    import sys, os
    func_map = load_functions_from_register()
    func = func_map[func_name]
    return func


class ResultType:
    default = 0
    figure = 1
    @classmethod
    def as_list(cls):
        return [value for key, value in vars(cls).items() if isinstance(value, int)]
        

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

    def _set_results_value(self, uid, val):
        with self._server._func_hash_results_map_lock:
            self._server._func_hash_results_map[uid] = val

    def _process_queue_element(self):
        """Processes either the multiprocessing or threaded results queues"""
        # getting the value tells us the result is complete
        uid, val = self._server._threaded_result_queue.get()
        self._server._func_hash_result_iscomplete_map[uid] = True

        self._set_results_value(uid, val)
        
        self._server._recalculate_queue.put(uid)

        # wake the client manager to update the client
        self._server._client_manager.wake()

    
    def _watch(self):
        """When awakened, retrieves the result queue and sends to the server cache.
        ResultsManager can reliably be woken up so no need for background checking.
        Unlike the ClientManager
        """
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
                        self._process_queue_element()
                    except queue.Empty:
                        break
                    time.sleep(0.01) # fairness sleep

            if self._stop_event.is_set():
                break 


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

    def _comarshal_com_object(self, com_obj_marshal) -> xl.Range:
        """Convert a comarshalled range to a range dispatch (or any object)"""
        com_obj_pyidispatch = pythoncom.CoGetInterfaceAndReleaseStream(
            com_obj_marshal, pythoncom.IID_IDispatch,
        )
        com_obj_dispatch = win32com.client.Dispatch(com_obj_pyidispatch)
        return com_obj_dispatch
    
    def _set_result_display(self, uid, val) -> None:
        with self._server._func_hash_result_display_map_lock:
            self._server._func_hash_result_display_map[uid] = val

    def _get_result_display(self, uid):
        with self._server._func_hash_results_map_lock:
            return self._server._func_hash_results_map[uid]
        
    def _get_value(self, uid):
        with self._server._func_hash_results_map_lock:
            return self._server._func_hash_results_map[uid]

    def _get_caller(self, uid):
        with self._server._func_hash_to_caller_map_lock:
            return self._comarshal_com_object(self._server._func_hash_to_caller_map[uid])
        
    def _get_unmarshalled_com_object(com_object):
        return pythoncom.CoMarshalInterThreadInterfaceInStream(
            pythoncom.IID_IDispatch, 
            com_object,
        )
    def _unmarshal_com_object(self, com_object):
        return pythoncom.CoMarshalInterThreadInterfaceInStream(
            pythoncom.IID_IDispatch, 
            com_object,
        )
        
    def _update_client_default(self, uid) -> None:
        """Update the data for the default case (row-major arrays, strings, values)"""
        caller_dispatch = self._get_caller(uid)
        val = self._get_value(uid)
        self._set_result_display(uid, val)

        # update by resetting the formula
        caller_dispatch.Formula2 = caller_dispatch.Formula2

        self._unmarshal_com_object(caller_dispatch)



    def _update_client_figure(self, uid) -> None:
        """Update the data for the figure case - add a figure image to the spreadsheet"""

        caller_dispatch = self._get_caller(uid)
        val = self._get_value(uid)
        fp, size_inches = val

        size_pt = size_inches * 72

        self._set_result_display(uid, f"Figure @'{str(fp)}'")

        caller_adjacent = caller_dispatch.Cells(2,1)
        xpos, ypos = caller_adjacent.Left, caller_adjacent.Top
        width, height = size_pt

        ws = caller_adjacent.Parent
        ws.Shapes.AddPicture(str(fp.resolve()), False, True, xpos, ypos, width, height)

        caller_dispatch.Formula2 = caller_dispatch.Formula2

        self._unmarshal_com_object(caller_dispatch)


    def _process_queue(self):
        """Process the queue at the current point in time. Any failed attempts get
        added back into the queue if appropriate.
        """
        # XXX - Warning that qsize() is not thread safe. Shouldn't be an issue.
        queue_length = self._server._recalculate_queue.qsize()
        for _ in range(queue_length):
            try:
                uid = self._server._recalculate_queue.get()
            except queue.Empty:
                return
            pass
            # fetch the result type so we know how to handle it
            with self._server._func_hash_result_type_lock:
                result_type = self._server._func_hash_result_type[uid]
            # if we fail to dispatch the range, it was probably deleted.
            # therefore we don't need to replace it in the queue
            replace_in_queue = True

            # Decide whether to recycle
            try:
                if result_type == ResultType.default:
                    self._update_client_default(uid)
                elif result_type == ResultType.figure:
                    self._update_client_figure(uid)
                else:
                    raise Exception("Result type invalid")
                replace_in_queue = False
                # XXX - todo - consider removing the uid after this call
            except pywintypes.com_error as e:
                if VBErrorConverter(e) == VBError.xlObjectRequired:
                    # Range has been deleted
                    replace_in_queue = False
                    logger.warning("Object Required. Cell has been deleted. Removing from queue.")
                    # XXX - todo - if this is dropped from the queue the uid definitely needs to be marked for delete.
                elif VBErrorConverter(e) == VBError.xlCallRejectedByCallee:
                    # Call rejected - user might be in a dialogue
                    logger.debug("Call rejected, recycling in queue")

                # XXX - Marshalling the caller back to the pool in case
                logger.info(f"Could not recalculate caller_dispatch for uid: '{uid}'")

            except AttributeError as e:
                logger.warning("AttributeError during cell update")

            # if we failed to update, recycle the queue as necessary
            if replace_in_queue:
                self._server._recalculate_queue.put(uid)


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
            return -1 # -1 flags an invalid call, any eq comparisons will fail
        # return the second half of the OLE Automation Error code.
        return VBErrorConverter.convert_hresult_to_vba_codes(e_hresult)[1]
        
    @staticmethod
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

    pass