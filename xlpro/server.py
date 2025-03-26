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
from win32typelibs import excel as xl
import matplotlib.figure
import sys

from xlpro import errors
from xlpro import config

cfg = config.load()
wd = Path(__file__).parent
logger = logging.getLogger(__name__)

from xlpro import _utils
from xlpro._wrappers import ModuleFunctionMapsWrapper
from xlpro._enums import FunctionTypes
from xlpro import _wrappers


from win32com.client import Dispatch

def initialize_and_get_workspace_xlpro_dir(workbook_path:Path) -> Path:
    d = workbook_path.parent / f"{workbook_path.name}.xlpro"
    d.mkdir(exist_ok=True)

    funcs_path = d / f"functions.py"
    subroutines_path = d / f"subroutines.py"

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

    return d

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
            # logger.debug("__init__ called however the singleton already exists and has been initialized.")
            pass
        return

    def getpid(self):
        return os.getpid()
    
    @classmethod
    def signal_shutdown(cls):
        cls._is_pending_close = True
    
    def force_shutdown(self):
        xlproServer.signal_shutdown()
    
    def register_and_configure_wb_workspace(self, wb_dispatch):
        # marshalling ok afaik - excel vba interface
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
        _utils.comarshal_release_and_get_stream(wb_dispatch) # marshalling ok afaik
        return self._workspace_map[uid]
    
    def _get_workspace_uid_from_wb(self, wb_dispatch):
        wb:xl._Workbook = win32com.client.Dispatch(wb_dispatch)
        wb_path = str(Path(wb.FullName))
        _utils.comarshal_release_and_get_stream(wb) # marshalling ok afaik
        return wb_path
        # return utils.hash_str(wb_path)

    def execute_function_async(self, wb_dispatch, caller, func_name, *args):
        # marshalling ok afaik - excel vba interface
        workspace = self._get_workspace_from_wb(wb_dispatch)
        return workspace.execute_function_async(caller=caller, fname=func_name, args=args)

    def register_functions_in_workspace(self, wb_dispatch):
        # marshalling ok afaik - excel vba interface
        workspace = self._get_workspace_from_wb(wb_dispatch)
        workspace._register_functions_in_self()

    def register_functions_in_vba(self, wb_dispatch):
        raise NotImplementedError("Obsoleted to remove combase.dll issue")

    def shutdown_workspace_from_dispatch(self, wb_dispatch):
        # XXX - todo - check this actually does anything meaninfgul
        # marshalling ok afaik - excel vba interface
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
        # marshalling ok afaik - excel vba interface
        workspace = self._get_workspace_from_wb(wb_dispatch)
        return workspace._get_vba_sync_text()

class xlproWorkspace:
    def __init__(self, server:xlproServer, wb_uid):
        self._server = server
        self._wb_uid = wb_uid
        self._wb_path = Path(wb_uid)
        self._wd = None # working directory

        self._uid_result_display_map = {} # the result to be displayed
        self._uid_result_display_map_lock = threading.Lock()

        self._uid_results_map = {} # the actual results
        self._uid_results_map_lock = threading.Lock()

        self._uid_result_iscomplete_map = {}
        self._uid_result_iscomplete_map_lock = threading.Lock()

        self._uid_result_type_map = {}
        self._uid_result_type_map_lock = threading.Lock()

        # maps the uid to the caller and function hash
        self._uid_to_caller_map = {}
        self._uid_to_caller_map_lock = threading.Lock() # XXX - todo - not used currently

        self._uid_pending_function_map = {} # uid: func
        self._uid_pending_function_map_lock = threading.Lock()

        self._caller_address_uid_map = {} # uid: address
        self._caller_address_uid_map_lock = threading.Lock()

        self._uid_args_cache = {}

        self._pending_function_queue = queue.Queue()
        self._result_queue = queue.Queue() # stores the results as they come in
        self._client_recalculate_queue = queue.Queue() # stores the cells that need to be recalculated.
        self._multip_fig_result_queue = multiprocessing.Queue() # stores the multiprocessing results

        # results manager handles processing the queue of results as they come in
        # and signalling to the client manager to execute commands.
        self._results_manager_thread:ResultsManager = ResultsManager(server=self)
        self._results_manager_thread.start()

        self._client_manager_thread = ClientManager(server=self)
        self._client_manager_thread.start()
        
        self._worker_manager = WorkerManager(server=self)
        self._worker_manager.start()

        self._func_hash_subthread_map = {}
        # self._func_hash_fig_generating_thread_map = {}

        self._temp_module_name:str = None

        self._module_function_maps_wrapper:ModuleFunctionMapsWrapper = None

    def set_xlpro_working_dir(self, wd:Path):
        logger.info(f"Setting working directory for workspace to '{str(wd)}'")
        self._wd = initialize_and_get_workspace_xlpro_dir(self._wb_path)

    def _configure_xlpro_files(self):
        # configure_workspace_xlpro_files(self._wd.parent, cfg)
        initialize_and_get_workspace_xlpro_dir(self._wb_path)

    def _register_functions_in_self(self):
        logger.info(f"Re-initializing workspace functions...")
        self._temp_module_name = f"{cfg.xlpro_functions_stem}_{_utils.hash_str(self._wb_uid)}"
        _wrappers.import_module_with_registration(self._temp_module_name, self._wd / f"{cfg.xlpro_functions_stem}.py")
        # self._valid_function_names = utils.get_function_names_from_module(self._temp_module_name)
        self._update_module_func_map_wrapper()

        logger.info(f"Reinitialization complete.")

    def _get_active_registered_functon_names(self):
        return [k for k, isactive in self._module_function_maps_wrapper.fname_isactive_register.items() if isactive]
    
    def _get_active_registered_functions(self):
        keys = self._get_active_registered_functon_names()
        return [self._module_function_maps_wrapper.fname_func_register[key] for key in keys]
    
    def _deregister_functions_in_self(self):
        logger.info(f"Uninitializing workspace functions...")
        self._valid_function_names = []
        if self._temp_module_name is not None:
            del sys.modules[self._temp_module_name]
        logger.info(f"Uninitialization complete.")

    def _update_module_func_map_wrapper(self):
        self._module_function_maps_wrapper = ModuleFunctionMapsWrapper(self._temp_module_name)

    def reset(self):
        self._configure_xlpro_files()
        self._deregister_functions_in_self()
        self._register_functions_in_self()
        logger.info("Clearing cached results")
        self._uid_results_map = {}
        pass

    def _get_function_by_name(self, fname):
        return self._module_function_maps_wrapper.fname_func_register[fname]

    def _clear_uid(self, uid):
        """Clear a uid from memory"""
        # XXX - todo - check if this is a valid method to purge an item from the queue.
        try:
            with self._uid_result_display_map_lock:
                del self._uid_result_display_map[uid]
        except:
            pass

        try:
            with self._uid_results_map_lock:
                del self._uid_results_map[uid]
        except:
            pass

        try:
            with self._uid_result_iscomplete_map_lock:
                del self._uid_result_iscomplete_map[uid]
        except:
            pass

        try:
            with self._uid_result_type_map_lock:
                del self._uid_result_type_map[uid]
        except:
            pass

        try:
            with self._uid_to_caller_map_lock:
                del self._uid_to_caller_map[uid]
        except:
            pass

        try:
            with self._uid_pending_function_map_lock:
                del self._uid_pending_function_map[uid]
        except:
            pass
        pass

        try:
            del self._uid_args_cache[uid]
        except:
            pass
        pass

    @staticmethod
    def _hash_excel_function_call(*args:typing.Iterable[str]):
        return _utils.hash_str(", ".join([str(x) for x in args]))

    def execute_function_async(self, caller, fname, args):
        try:
            func = self._get_function_by_name(fname)

            if not func:
                return f"Function {fname} not found."

            uid = xlproWorkspace._hash_excel_function_call(fname, *args)
            # uid = xlproWorkspace._hash_excel_function_call(caller.Address, fname, *args)

            logger.debug(f"Calling function '{fname}', uid: '{uid}', args: '{args}'")

            self._uid_args_cache[uid] = args

            # return the cached result if it exists
            with self._uid_result_display_map_lock:
                if uid in self._uid_result_iscomplete_map.keys():
                    return self._uid_result_display_map[uid]
                
            # Clear any lingering calculations coming from this caller if the result isnt cached
            # downstream functions should pick up on these being deleted
            caller_dispatch = Dispatch(caller)
            caller_addr = caller_dispatch.Address

            with self._caller_address_uid_map_lock:
                # the hash will be constant for a function/args/caller combination so this is valid
                if caller_addr in self._caller_address_uid_map.keys():
                    self._clear_uid(self._caller_address_uid_map[caller_addr])
                self._caller_address_uid_map[caller_addr] = uid

            # this was originally added to deal with a recalculation order issue where None
            # would get passed in the args mid calculation cycle
            if isinstance(args[0][0], list|tuple):
                if None in args[0][0]:
                    logger.error(f"Args are fkd '{args}'")
                    x, y = args[0][0]
                    pass
                # if args[0][0][1] is not None:
                #     x, y = args[0][0]
                #     pass

            # release the com args for use in another thread. convert them to streams
            # args = utils.com_args_release_to_stream_reserved(func, args)

            result_type = self._module_function_maps_wrapper.fname_type_register[fname]
            
            self._uid_result_type_map[uid] = result_type

            caller_stream = _utils.comarshal_release_and_get_stream(caller_dispatch) # this marshal is the OG
            self._uid_to_caller_map[uid] = caller_stream

            # configure default state for result and iscomplete status
            with self._uid_result_display_map_lock:
                self._uid_result_display_map[uid] = f"Promise<{uid}>"
            with self._uid_result_display_map_lock:
                self._uid_result_iscomplete_map[uid] = False

            # handle different function types
            if result_type == FunctionTypes.default:
                f = self._create_worker_func(uid, func, args, kwargs={})
                self._uid_pending_function_map[uid] = f
                self._pending_function_queue.put(uid)
                self._worker_manager.wake()

            elif result_type == FunctionTypes.figure:
                fig_generating_thread = FigureGeneratingThread(
                    wd=self._wd,
                    uid=uid,
                    func_name=fname,
                    args=args,
                    kwargs={},
                    return_value_queue=self._result_queue, 
                    return_event=self._results_manager_thread._wake_event,
                )
                # self._func_hash_fig_generating_thread_map[uid] = fig_generating_thread
                fig_generating_thread.start()

            # return the (incomplete result)
            with self._uid_result_display_map_lock:
                return self._uid_result_display_map[uid]

        # Return the python exception string as a fallback
        except Exception as e:
            return str(e)
        
    def _create_worker_func(self, uid, func, args, kwargs):
        if kwargs:
            raise Exception("kwargs should not be here!")
        
        def worker():
            f = _wrappers.generate_wrapped_function(self._temp_module_name, func.__name__)
            try:
                ret = f(*args, **kwargs)
                self._result_queue.put((uid, ret))
                logger.info(f"Completed function '{uid}' successfully")
            except Exception as e:
                self._result_queue.put((uid, e))
                logger.info(f"Completed function '{uid}' unsuccessfully with error {e}")
            logger.debug("Waking results manager from worker thread...")
            self._results_manager_thread.wake()
        
        return worker
            
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
        funcs = self._get_active_registered_functions()
        return _utils.get_xlpro_vb_dynamic_component_contents(funcs)
    
    def get_caller_stream(self, uid):
        with self._uid_to_caller_map_lock:
            return self._uid_to_caller_map[uid]
    
    def set_caller_stream(self, uid, val):
        with self._uid_to_caller_map_lock:
            self._uid_to_caller_map[uid] = val

    def _get_uid_debug_info(self, uid):
        a0 = self._uid_result_display_map.get(uid, None)
        a1 = self._uid_results_map.get(uid, None)
        a2 = self._uid_result_iscomplete_map.get(uid, None)
        a3 = self._uid_result_type_map.get(uid, None)
        a4 = self._uid_pending_function_map.get(uid, None)
        pythoncom.CoInitialize()
        try:
            rng_stream = self.get_caller_stream(uid)
            rng_dispatch:xl.Range = _utils.comarshal_dispatch_stream(rng_stream)
            a5 = rng_dispatch.Address
            a6 = rng_dispatch.Formula
            a7 = rng_dispatch.Value
        except Exception as e:
            a5 = e
            a6 = "ERROR"
            a7 = "ERROR"
        try:
            self.set_caller_stream(uid, _utils.comarshal_release_and_get_stream(rng_dispatch))
        except:
            pass
        a8 = self._uid_args_cache.get(uid, None)
        ret = {
            "uid": uid,
            "result_display": a0,
            "results": a1,
            "result_iscomplete": a2,
            "result_type": a3,
            "pending_function": a4,
            "Address": a5,
            "Formula": a6,
            "Value": a7,
            "Args": a8
        }
        pass
        pythoncom.CoUninitialize()
        return ret
        
# wrappers must be applied before we pickle the function I believe...
# @utils.type_converter_wrapper
# @utils.com_init_dispatch_release_wrapper
def figure_process_func(wd:Path, uid, func_name, args, kwargs, queue):

    module_name = f"{cfg.xlpro_functions_stem}_{uid}"
    _utils.import_module(f"{cfg.xlpro_functions_stem}_{uid}", wd / f"{cfg.xlpro_functions_stem}.py")
    raise NotImplementedError
    func = getattr(sys.modules[module_name], func_name)

    fp = wd / "tmp" / f"{uid}.svg"
    fp.parent.mkdir(parents=True, exist_ok=True)

    fig:matplotlib.figure.Figure = func(*args, **kwargs)
    size_inches = np.array(fig.get_size_inches())
    fig.savefig(fp, dpi=600)

    queue.put((uid, (fp, size_inches)))
    pass

def setup_scope_and_get_function(module_name, module_path, func_name):
    _utils.import_module(module_name, wd / f"{module_name}.py")
    func_map = _utils.get_udf_valid_functions_from_module(module_name)
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


class WorkerManager:
    """Manages function execution for a workspace. Sends results to the results manager"""
    def __init__(self, server:xlproWorkspace):
        self._server = server
        self._stop_event = threading.Event()
        self._wake_event = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)
        # self._threadpool:list[threading.Thread] = []
        self._threadpool_dict:dict[str, threading.Thread] = {}

    @property
    def MAX_THREADS(self):
        return cfg.max_worker_threads

    def start(self):
        self._thread.start()
        self._thread
        logger.info(f"WorkerManager thread started: tid: {threading.get_native_id()}")

    def stop(self):
        self._stop_event.set()
        self._wake_event.set()  # Wake up if sleeping
        self._thread.join()

    def wake(self):
        self._wake_event.set()

    def _process_function_queue(self):
        if len(self._threadpool_dict.keys()) < self.MAX_THREADS:
            uid = self._server._pending_function_queue.get(timeout=0.01)
            try:
                with self._server._uid_pending_function_map_lock:
                    func = self._server._uid_pending_function_map[uid]
            except KeyError:
                logger.warning(f"Pending function queue uid not available, ignoring calculation request for uid '{uid}'")
                return
            
            if uid in self._threadpool_dict:
                logger.debug(f"Rejected to start worker for uid: '{uid}', already running")
                return

            t = threading.Thread(target=func, daemon=True)
            self._threadpool_dict[uid] = t
            # XXX - todo - limit the number of attempts for a given function in some way
            # XXX - todo - support sending terminate command to lingering worker threads
            t.start()
            return

        logger.info(f"Reached maximum worker thread cap - MAX_THREADS: {self.MAX_THREADS}")

    def _clear_completed_threads(self):
        self._threadpool_dict = {uid: t for uid, t in self._threadpool_dict.items() if t.is_alive()}

    def _run(self):
        while not self._stop_event.is_set():
            logger.debug("WorkerManager thread is waiting for events or timeout...")
            self._wake_event.wait(timeout=5)
            if self._wake_event.is_set():
                self._wake_event.clear()
                logger.info("WorkerManager thread woke up for an event!")
            while True:
                try:
                    self._process_function_queue()
                except queue.Empty:
                    break
                self._clear_completed_threads()
                time.sleep(0.1) # fairness sleep

            if self._stop_event.is_set():
                break 


class ResultsManager:
    """Manages the results queue and signals update requests to the client manager"""
    def __init__(self, server:xlproWorkspace):
        self._stop_event = threading.Event()
        self._wake_event = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)
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
        with self._server._uid_results_map_lock:
            self._server._uid_results_map[uid] = val


    def _recalculate_precedents(self, uid):
        pythoncom.CoInitialize()
        caller_dispatch = _utils.comarshal_dispatch_stream(self._server.get_caller_stream(uid))
        precedents_stream = _utils.get_precedents_chain(caller_dispatch)
        
        precedents_recalculate = []
        active_formulas = self._server._get_active_registered_functon_names()

        u = [x.AddressLocal for x in precedents_stream]

        for ps_disp in precedents_stream:
            try:
                ps_disp:xl.Range
                formula = ps_disp.Formula2
                ps_disp.Formula2 = ps_disp.Formula2
                if _utils.formula_is_for_xlpro(formula, active_formulas):
                    precedents_recalculate.append(ps_disp)
                else:
                    _utils.comarshal_release_and_get_stream(ps_disp) # marshal release only afaik - obj created in this thread
            except Exception as e:
                logger.warning(f"Error during recalculate: {e}")
                continue
        v = [x.AddressLocal for x in precedents_recalculate]
        logger.debug(f"Precedents, '{len(v)}' {v[:min(12, len(v))]}")
        for ps_disp in precedents_recalculate:
            try:
                ps_disp.Formula2 = ps_disp.Formula2
            except Exception as e:
                logger.warning(f"Error during recalculate2: {e}")
                continue

            _utils.comarshal_release_and_get_stream(ps_disp) # marshal release only afaik - XXX - todo - check
        self._server.set_caller_stream(uid, _utils.comarshal_release_and_get_stream(caller_dispatch))
        pythoncom.CoUninitialize()

    def _process_queue_element(self):
        """Processes either the multiprocessing or threaded results queues
        """
        # getting the value tells us the result is complete
        uid, val = self._server._result_queue.get(timeout=0.01)
        ret = val

        # the uid may have been removed due to irrelevance from the workspace
        with self._server._uid_pending_function_map_lock:
            if not uid in self._server._uid_pending_function_map.keys():
                logger.debug(f"Uid '{uid}' not found as pending function within results manager. Ignoring this")
                return

        # if the result is an exception - we need to handle it
        if isinstance(val, Exception):
            # if the arguments weren't ready - ignore the process request
            if isinstance(val, errors.ArugmentNotReadyException):
                logger.debug(f"Arguments not ready for uid '{uid}', recycling function...")

                logger.debug(f"Arguments not ready for uid '{uid}', Attempting to recalculate precedents...")
                self._server._get_uid_debug_info(uid)
                # self._recalculate_precedents(uid)

                logger.debug(f"ResultsManager is waking the worker manager to replace '{uid}'")
                self._server._pending_function_queue.put(uid)
                self._server._worker_manager.wake()
                return
                
            elif isinstance(val, pythoncom.com_error):
                if VBErrorConverter(val) == VBError.xlCallRejectedByCallee:
                    self._server._pending_function_queue.put(uid)
                    logger.info(f"Call rejected by callee for '{uid}', recycling function...")
                    return
            
            logger.warning(f"Returned value is a generic exception: {uid}, {val}")
            ret = str(val) # convert exception to string for it to show in excel.


        # if it is a valid return, write the result to the cache
        self._set_results_value(uid, ret)
        # signal that it is complete
        with self._server._uid_result_iscomplete_map_lock:
            self._server._uid_result_iscomplete_map[uid] = True
        # remove from the pending function map once successfully completed.
        del self._server._uid_pending_function_map[uid]

        # signal to the client manager to update the client
        self._server._client_recalculate_queue.put(uid)
        logger.debug(f"ResultsManager is waking the ClientManager after successful calculation of '{uid}'")
        self._server._client_manager_thread.wake()

    def _run(self):
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
                time.sleep(0.1) # fairness sleep

            if self._stop_event.is_set():
                break 


class ClientManager:
    """Hooks to the excel client so we can trigger events"""
    def __init__(self, server:xlproWorkspace):
        self._stop_event = threading.Event()
        self._wake_event = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)
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
    
    def _set_result_display(self, uid, val) -> None:
        # XXX - todo - move logic to server
        with self._server._uid_result_display_map_lock:
            self._server._uid_result_display_map[uid] = val

    def _get_result_display(self, uid):
        # XXX - todo - move logic to server
        with self._server._uid_results_map_lock:
            return self._server._uid_results_map[uid]
        
    def _get_value(self, uid):
        # XXX - todo - move logic to server
        with self._server._uid_results_map_lock:
            return self._server._uid_results_map[uid]
        
    def _update_client_default_result(self, uid) -> None:
        """Update the data for the default case (row-major arrays, strings, values)"""
        try:
            caller_dispatch = _utils.comarshal_dispatch_stream(self._server.get_caller_stream(uid))
            val = self._get_value(uid)
            if isinstance(val, Exception):
                logger.warning(f"Value is an exception: '{e}', '{uid}'")
            self._set_result_display(uid, val)

            # update by resetting the formula
            caller_dispatch.Formula2 = caller_dispatch.Formula2
            self._server.set_caller_stream(uid, _utils.comarshal_release_and_get_stream(caller_dispatch))
        except KeyError as e:
            # XXX - todo - there is a risk of a keyerror here for some reason
            logger.error(f"Error during client update: '{uid}', {e}")
        
        except Exception as e:
            try:
                self._server.set_caller_stream(uid, _utils.comarshal_release_and_get_stream(caller_dispatch))
            except:
                pass
            raise e

    def _update_client_figure_result(self, uid) -> None:
        """Update the data for the figure case - add a figure image to the spreadsheet"""
        caller_dispatch = self._server.get_caller_stream(uid)
        try:
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
            self._server.set_caller_stream(_utils.comarshal_release_and_get_stream(caller_dispatch))

        except Exception as e:
            self._server.set_caller_stream(_utils.comarshal_release_and_get_stream(caller_dispatch))
            raise e

    def _process_queue(self):
        """Process the queue at the current point in time. Any failed attempts get
        added back into the queue if appropriate.
        """
        # XXX - Warning that qsize() is not thread safe
        # this prevents inplace recycling and potentially infinite loop
        queue_length = self._server._client_recalculate_queue.qsize()
        logger.debug(f"Client manager queue length estimate is {queue_length}")
        for _ in range(queue_length):
            try:
                uid = self._server._client_recalculate_queue.get()
                logger.debug(f"Client manager fetched uid '{uid}'")
            except queue.Empty:
                logger.debug("Client manager queue is empty")
                return
            
            with self._server._uid_result_iscomplete_map_lock:
                if not uid in self._server._uid_result_iscomplete_map.keys():
                    logger.debug(f"Uid '{uid}' not found as pending function within client manager update. Ignoring this iteration")
                    return

            pass
            # fetch the result type so we know how to handle it
            with self._server._uid_result_type_map_lock:
                result_type = self._server._uid_result_type_map[uid]
            # if we fail to dispatch the range, it was probably deleted.
            # therefore we don't need to replace it in the queue
            replace_in_queue = True

            # Decide whether to recycle
            try:
                if result_type == FunctionTypes.default:
                    self._update_client_default_result(uid)
                elif result_type == FunctionTypes.figure:
                    self._update_client_figure_result(uid)
                else:
                    raise Exception("Result type invalid")
                # if successful we don't need to replace#
                logger.debug(f"Client manager successfully processed uid '{uid}'. Not replacing")
                replace_in_queue = False
                # XXX - todo - consider removing the uid after this call
            except pywintypes.com_error as e:
                if VBErrorConverter(e) == VBError.xlObjectRequired:
                    # Range has been deleted
                    replace_in_queue = False
                    logger.error("VB Error - Object Required. Cell has been deleted. Removing from queue.")
                    # XXX - todo - if this is dropped from the queue the uid definitely needs to be marked for delete.
                elif VBErrorConverter(e) == VBError.xlCallRejectedByCallee:
                    # Call rejected - excel might be in a dialogue 
                    logger.debug("VB Error - Call rejected, recycling in queue")

                logger.info(f"Could not recalculate caller_dispatch for uid: '{uid}'")

            except AttributeError as e:
                logger.warning("AttributeError during cell update, Application may be in dialogue")
            except Exception as e:
                logger.critical(f"Error during client queue processing!")
                raise e
            finally:
                # XXX - Marshalling the caller back to the pool in case
                logger.debug("Releasing caller dispatch during ClientManager._process_queue()")

            # if we failed to update, recycle the queue as necessary

            if replace_in_queue:
                if "jsonify" in self._server._get_uid_debug_info(uid)["Formula"]:
                    pass
                self._server._client_recalculate_queue.put(uid)

            time.sleep(0.1) # fairness sleep


    def _run(self):
        """When awakened, retrieves the result queue and sends to the server cache.
        """
        pythoncom.CoInitialize()

        while not self._stop_event.is_set():
            # wait for ten seconds for an event, otherwise do a queue process to check for 
            # outstanding tasks.
            self._wake_event.wait(timeout=0.5)
            if self._wake_event.is_set():
                logger.debug("ClientManager thread woke up for an event!")
                self._wake_event.clear()
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
            return -1 # -1 flags an invalid call, -1 will make any eq comparisons fail
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