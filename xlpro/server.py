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

import numpy as np
import pywintypes

import utils

from win32typelibs import excel as xl

import matplotlib.figure

import config

import sys

cfg = config.load()
wd = Path(__file__).parent
logger = logging.getLogger(__name__)

def configure_workspace_xlpro_files(wd:Path, cfg:config.Configuration):
    xlpro_dir_path = Path() / wd / cfg.xlpro_directory
    xlpro_dir_path.mkdir(exist_ok=True)

    funcs_path = xlpro_dir_path / f"{cfg.xlpro_functions_stem}.py"
    subroutines_path = xlpro_dir_path / f"{cfg.xlpro_subroutines_stem}.py"

    p = funcs_path
    if not p.exists():
        with open(p, "w") as f:
            f.write(f"# > {p.resolve()}\n")
            f.write(f"# xlpro will automatically detect functions in this file as Excel UDFs.\n\n")

    p = subroutines_path
    if not p.exists():
        with open(p, "w") as f:
            f.write(f"# > {p.resolve()}\n")
            f.write(f"# xlpro will automatically detect functions in this file as Excel subroutines.\n\n")

    
class ServerClosedException(Exception):
    def __init__(self, *args):
        super().__init__(*args)

class xlproServer:
    _public_methods_ = [
        "getpid",

        "register_and_configure_wb_workspace",
        "register_functions_in_vba",
        "register_functions_in_workspace",

        "execute_function_async",

        "shutdown_workspace",
        "shutdown",

        "get_vba_sync_text",

        "__dev_shutdown",
    ]
    # _reg_progid_ = config.progid
    _reg_clsid_ = cfg.clsid

    _instance = None  # Singleton instance
    _instance_initialized = False

    # use this flag to tell the server to close it next time
    # it is checked.
    _is_pending_close = False

    def __new__(cls, *args, **kwargs):
        if cls._instance is None:
            cls._instance = super().__new__(cls, *args, **kwargs)
            # logger.log(f"__new__ creating new instance")
            # logger.log(f"__new__ creating new instance. PID: {os.getpid()}, MEMID: {hex(id(cls._instance))}")
        return cls._instance

    def __init__(self):
        if not xlproServer._instance_initialized:
            pythoncom.CoInitialize()
            xlproServer._instance_initialized = True
            self._workspace_map:dict[str, xlproWorkspace] = {} # uid (path) to workspace
        else:
            logger.debug("__init__ called however the singleton already exists and has been initialized.")
        return

    def getpid(self):
        return os.getpid()
    
    @classmethod
    def signal_shutdown(cls):
        cls._is_pending_close = True
    
    def force_shutdown(self):
        xlproServer.signal_shutdown()
    
    def register_and_configure_wb_workspace(self, wb_dispatch):
        uid = self._get_workspace_uid_from_wb(wb_dispatch)

        if not uid in self._workspace_map.keys():
            logger.info(f"Creating workspace '{uid}'...")
            workspace = xlproWorkspace(self, uid)
            workspace_wd = Path(uid).parent.resolve()

            workspace.set_xlpro_working_dir(workspace_wd)
            workspace.reset()

            self._workspace_map[uid] = workspace
            logger.info(f"Creation of workspace complete for '{uid}'.")

            pass
        else:
            logger.info(f"Workspace already exists, resetting workspace '{uid}'.")
            workspace = self._get_workspace_from_wb(wb_dispatch)
            workspace.reset()
            logger.info(f"Workspace reset complete for '{uid}'.")
            
            # logger.info(f"Workspace '{uid}' already exists. Shutting down and re-initializing workspace.")
            # self.shutdown_workspace_from_dispatch(wb_dispatch)
            # logger.info(f"Re-initializing workspace...")
            # self.register_and_configure_wb_workspace(wb_dispatch)
            # logger.info(f"Re-initialization complete")

    def _get_workspace_from_wb(self, wb_dispatch):
        uid = self._get_workspace_uid_from_wb(wb_dispatch)
        if not uid in self._workspace_map.keys():
            logger.info("Workbook has not been registered, initializing...")
            self.register_and_configure_wb_workspace(wb_dispatch)
        utils.comarshal_release_and_get_stream(wb_dispatch)
        return self._workspace_map[uid]
    
    def _get_workspace_uid_from_wb(self, wb_dispatch):
        wb:xl._Workbook = win32com.client.Dispatch(wb_dispatch)
        wb_path = str(Path(wb.FullName))
        utils.comarshal_release_and_get_stream(wb)
        return wb_path
        # return utils.hash_str(wb_path)

    def execute_function_async(self, wb_dispatch, caller, func_name, *args):
        workspace = self._get_workspace_from_wb(wb_dispatch)
        return workspace.execute_function_async(caller=caller, func_name=func_name, args=args)

    def register_functions_in_workspace(self, wb_dispatch):
        workspace = self._get_workspace_from_wb(wb_dispatch)
        # workspace.register_functions_in_self(utils.comarshal_release_and_get_stream(wb_dispatch))
        workspace._register_functions_in_self()

    def register_functions_in_vba(self, wb_dispatch):
        raise NotImplementedError("Obsoleted to remove combase.dll issue")

    def shutdown_workspace_from_dispatch(self, wb_dispatch):
        # XXX - todo - check this actually does anything meaninfgul
        workspace = self._get_workspace_from_wb(wb_dispatch)
        uid = self._get_workspace_uid_from_wb(wb_dispatch)
        logger.info(f"Shutting down workspace uid:'{uid}'")
        workspace:xlproWorkspace
        workspace._shutdown()
        del workspace
        del self._workspace_map[uid]
        n_live_workspaces = len(list(self._workspace_map.values())) 
        logger.info(f"There are currently {n_live_workspaces}")

    def shutdown(self):
        """Shuts down the server"""
        n_live_workspaces = len(list(self._workspace_map.values())) 
        if n_live_workspaces == 0:
            logger.info("No workspaces alive, shutting down the server...")
            xlproServer.signal_shutdown()
        logger.error(f"Unable to shutdown, {n_live_workspaces} are active. Please close these first.")
        pass

    def get_vba_sync_text(self, wb_dispatch):
        """Gets the vba code module contents to register the udfs"""
        workspace = self._get_workspace_from_wb(wb_dispatch)
        return workspace._get_vba_sync_text()

class xlproWorkspace:
    def __init__(self, server:xlproServer, wb_uid):
        self._server = server
        self._wb_uid = wb_uid
        self._wd = None # working directory

        self._func_hash_result_display_map = {} # the result to be displayed
        self._func_hash_result_display_map_lock = threading.Lock()

        self._func_hash_results_map = {} # the actual results
        self._func_hash_results_map_lock = threading.Lock()

        self._func_hash_result_iscomplete_map = {}
        self._func_hash_result_iscomplete_map_lock = threading.Lock()

        self._func_hash_result_type = {}
        self._func_hash_result_type_lock = threading.Lock()

        # maps the uid to the caller and function hash
        self._func_hash_to_caller_map = {}
        self._func_hash_to_caller_map_lock = threading.Lock() # XXX - todo - not used currently

        self._threaded_result_queue = queue.Queue() # stores the results as they come in
        self._recalculate_queue = queue.Queue() # stores the cells that need to be recalculated.

        self._multiprocessing_figure_result_queue = multiprocessing.Queue() # stores the multiprocessing results

        # results manager handles processing the queue of results as they come in
        # and signalling to the client manager to execute commands.
        self._results_manager_thread:ResultsManager = ResultsManager(server=self)
        self._results_manager_thread.start()

        self._client_manager_thread = ClientManager(server=self)
        self._client_manager_thread.start()

        self._func_hash_subthread_map = {}
        # self._func_hash_fig_generating_thread_map = {}

        self._temp_module_name:str = None
        self._valid_function_names:list[str] = []

    def set_xlpro_working_dir(self, wd:Path):
        logger.info(f"Setting working directory for workspace to '{str(wd)}'")
        self._wd = wd / cfg.xlpro_directory
        self._wd.mkdir(parents=True, exist_ok=True)

    def _configure_xlpro_files(self):
        configure_workspace_xlpro_files(self._wd.parent, cfg)

    def _register_functions_in_self(self):
        logger.info(f"Re-initializing workspace functions...")
        self._temp_module_name = f"{cfg.xlpro_functions_stem}_{utils.hash_str(self._wb_uid)}"
        utils.import_module(self._temp_module_name, self._wd / f"{cfg.xlpro_functions_stem}.py")
        self._valid_function_names = utils.get_function_names_from_module(self._temp_module_name)
        logger.info(f"Reinitialization complete.")

    def _deregister_functions_in_self(self):
        logger.info(f"Uninitializing workspace functions...")
        self._valid_function_names = []
        if self._temp_module_name is not None:
            del sys.modules[self._temp_module_name]
        logger.info(f"Uninitialization complete.")

    def reset(self):
        self._configure_xlpro_files()
        self._deregister_functions_in_self()
        self._register_functions_in_self()
        logger.info("Clearing cached results")
        self._func_hash_results_map = {}



    def register_functions_in_vba(self, wb_stream):
        raise NotImplementedError("Obsoleted to remove combase.dll issue")

    def _get_function_by_name(self, func_name):
        return getattr(sys.modules[self._temp_module_name], func_name)


    def execute_function_async(self, caller, func_name, args):
        try:
            func = self._get_function_by_name(func_name)
            if not func:
                return f"Function {func_name} not found."

            uid = utils.hash_function_call(func, utils.get_args_minus_reserved(func, args), kwargs={})

            # if the hash of the function has been marked complete, return that
            # avoid computation
            with self._func_hash_result_display_map_lock:
                if uid in self._func_hash_result_iscomplete_map.keys():
                    # return self._func_hash_results_map[uid]
                    return self._func_hash_result_display_map[uid]

            # release the com args for use in another thread. convert them to streams
            args = utils.com_args_release_to_stream_reserved(func, args)

            result_type = utils.get_func_result_type(func)
            self._func_hash_result_type[uid] = result_type

            # XXX - todo - check if the caller is a range.
            caller_stream = utils.comarshal_release_and_get_stream(caller)
            self._func_hash_to_caller_map[uid] = caller_stream

            with self._func_hash_result_display_map_lock:
                self._func_hash_result_display_map[uid] = f"Promise<{uid}>"
            with self._func_hash_result_display_map_lock:
                self._func_hash_result_iscomplete_map[uid] = False

            if result_type == ResultType.default:
                t = self._create_and_register_async_worker(uid, func, args, kwargs={})
                self._func_hash_subthread_map[uid] = t
                t.start()
            elif result_type == ResultType.figure:
                fig_generating_thread = FigureGeneratingThread(
                    wd=self._wd,
                    uid=uid,
                    func_name=func_name,
                    args=args,
                    kwargs={},
                    return_value_queue=self._threaded_result_queue, 
                    return_event=self._results_manager_thread._wake_event,
                )
                # self._func_hash_fig_generating_thread_map[uid] = fig_generating_thread
                fig_generating_thread.start()

            # return the (incomplete result)
            with self._func_hash_result_display_map_lock:
                return self._func_hash_result_display_map[uid]

        # Return the python exception string as a fallback
        except Exception as e:
            return str(e)
        
    def _create_and_register_async_worker(self, uid, func, args, kwargs):
        if kwargs:
            raise Exception("kwargs should not be here!")
        def func_wrapper():

            f = utils.type_converter_wrapper(
                utils.com_init_dispatch_release_wrapper(func)
            )
            ret = f(*args, **kwargs)

            self._threaded_result_queue.put((uid, ret))
            with self._func_hash_result_iscomplete_map_lock:
                self._func_hash_result_iscomplete_map[uid] = True

            self._results_manager_thread.wake()

        t = threading.Thread(target=func_wrapper, daemon=True)
        return t
    
    def _shutdown(self):
        for uid, t in self._func_hash_subthread_map.items():
            t:threading.Thread
            if t.is_alive():
                t.join()
        self._results_manager_thread.stop()
        self._client_manager_thread.stop()
        # for uid, p in self._func_hash_fig_generating_thread_map.items():
        #     p:FigureGeneratingThread
        #     p.stop()

    def _get_vba_sync_text(self) -> str:
        funcs = [getattr(sys.modules[self._temp_module_name], f) for f in self._valid_function_names]
        s = utils.get_xlpro_vb_dynamic_component_contents(funcs)
        del funcs
        return s
    

# wrappers must be applied before we pickle the function I believe...
@utils.type_converter_wrapper
@utils.com_init_dispatch_release_wrapper
def figure_process_func(wd:Path, uid, func_name, args, kwargs, queue):

    module_name = f"{cfg.xlpro_functions_stem}_{uid}"
    utils.import_module(f"{cfg.xlpro_functions_stem}_{uid}", wd / f"{cfg.xlpro_functions_stem}.py")
    func = getattr(sys.modules[module_name], func_name)

    fp = wd / "tmp" / f"{uid}.svg"
    fp.parent.mkdir(parents=True, exist_ok=True)

    fig:matplotlib.figure.Figure = func(*args, **kwargs)
    size_inches = np.array(fig.get_size_inches())
    fig.savefig(fp, dpi=600)

    queue.put((uid, (fp, size_inches)))
    pass

def setup_scope_and_get_function(module_name, module_path, func_name):
    utils.import_module(module_name, wd / f"{module_name}.py")
    func_map = utils.get_functions_from_module(module_name)
    func = func_map[func_name]
    return func

class FigureGeneratingThread:
    """Class which maintains a thread which waits for a process to finish."""
    def __init__(self, wd, uid, func_name, args, kwargs, return_value_queue, return_event):
        self._wd = wd
        if not isinstance(self._wd, Path):
            raise TypeError 
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

    def stop(self, timeout_s=10):
        self._stop_event.set()
        self._wake_event.set()  # Wake up if sleeping
        self._thread.join(timeout=timeout_s)
        # self._process.terminate()

    def wake(self):
        self._wake_event.set()

    def _create_process(self) -> multiprocessing.Process:

        p = multiprocessing.Process(
            # target=utils.type_converter_wrapper(
            #     utils.com_init_dispatch_release_wrapper(
            #         figure_process_func,
            #     )
            target=figure_process_func,
            daemon=True,
            kwargs={
                "wd": self._wd,
                "uid": self._uid,
                "func_name": self._func_name,
                "args": self._args,
                "kwargs": self._kwargs,
                "queue": self._multiprocess_queue,
            }
        )
        return p
    
    def _run(self): 
        p = self._create_process()
        p.start()
        p.join() # once joined, we can check the queue
        # get the uid and value (figure path in this case)
        uid, val = self._multiprocess_queue.get()
        # send the uid/value combo to the return queue
        # the return queue should always be the server value return queue 
        self._return_queue.put((uid, val))
        self._return_event.set()
        
class ResultType:
    default = 0
    figure = 1
    @classmethod
    def as_list(cls):
        return [value for key, value in vars(cls).items() if isinstance(value, int)]
        
class ResultsManager:
    def __init__(self, server:xlproWorkspace):
        self._stop_event = threading.Event()
        self._wake_event = threading.Event()
        self._thread = threading.Thread(target=self._rmgr_watch, daemon=True)
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
        uid, val = self._server._threaded_result_queue.get(timeout=0.01)
        with self._server._func_hash_result_iscomplete_map_lock:
            self._server._func_hash_result_iscomplete_map[uid] = True

        self._set_results_value(uid, val)
        
        self._server._recalculate_queue.put(uid)

        # wake the client manager to update the client
        self._server._client_manager_thread.wake()

    def _rmgr_watch(self):
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

    def __init__(self, server:xlproWorkspace):
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
            return utils.comarshal_dispatch_stream(self._server._func_hash_to_caller_map[uid])
        
    def _get_unmarshalled_com_object(self, com_object):
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

        utils.comarshal_release_and_get_stream(caller_dispatch)

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
        # ws.Shapes.AddPicture(str(fp.resolve()), False, True, xpos, ypos, width, height)
        ws.Shapes.AddPicture(str(fp.resolve()), False, True, xpos, ypos, width, height)

        caller_dispatch.Formula2 = caller_dispatch.Formula2

        utils.comarshal_release_and_get_stream(caller_dispatch)

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