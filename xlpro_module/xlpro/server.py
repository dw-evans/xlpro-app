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
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from win32typelibs import excel as xl
import matplotlib.figure
import sys

from xlpro import errors
from xlpro import config

cfg = config.load()
wd = Path(__file__).parent
logger = logging.getLogger(__name__)

from xlpro import _utils
from xlpro._wrappers import ModuleFunctionMapsWrapper, ModuleSubMapsWrapper
from xlpro._enums import FunctionTypes
from xlpro import _wrappers
import regex as re
from copy import deepcopy

from xlpro._types import xlproImage

from win32com.client.dynamic import Dispatch

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

    return 

def get_workspace_xlpro_dir(workbook_path:Path) -> Path:
    return workbook_path.parent / f"{workbook_path.name}.xlpro"


SUB_CALLER_FLAG_STRING = "SUB_CALLER_FLAG_STRING"


class xlproServer:
    _public_methods_ = [
        "getpid",

        "register_and_configure_wb_workspace",
        "register_functions_in_vba",
        "register_functions_in_workspace",

        "execute_function_async",
        "execute_sub_async",

        "shutdown_workspace",
        "shutdown",

        "get_vba_sync_text",
        "get_vba_sync_text_subs",

        "__dev_shutdown",
    ]
    # _reg_progid_ = config.progid
    _reg_clsid_ = "undefined"

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
            self._workspace_map_lock = threading.Lock()
            self._workspace_map:dict[str, xlproWorkspace] = {} # uid (path) to workspace
            self._workspace_uid_to_workbook_path_lock = threading.Lock()
            self._workspace_uid_to_workbook_path:dict[str, str] = {} # uid to workspace path (done to eliminate reference duplication...)
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
    
    def refresh_workspace(self, wb_dispatch):
        wb_path = self._get_workspace_pathuid_from_wb(wb_dispatch)
        if not wb_path in self._workspace_map.keys():
            raise Exception("wb is not registered")
        
        workspace = self._get_workspace_from_wb(wb_dispatch)
        workspace.reset()
        pass
    
    def register_and_configure_wb_workspace(self, wb_dispatch):
        # marshalling ok afaik - excel vba interface
        wb_path = self._get_workspace_pathuid_from_wb(wb_dispatch)

        if not wb_path in self._workspace_map.keys():
            logger.info(f"Creating workspace '{wb_path}'...")
            uid = _utils.hash_str(wb_path)
            workspace = xlproWorkspace(self, wb_uid=wb_path, uid=uid)
            workspace_wd = Path(wb_path).parent.resolve()

            workspace.set_xlpro_working_dir(workspace_wd)
            workspace.reset()

            self._workspace_map[wb_path] = workspace
            self._workspace_uid_to_workbook_path[uid] = wb_path
            logger.info(f"Creation of workspace complete for '{wb_path}'.")

            pass
        else:
            logger.info(f"Workspace already exists, resetting workspace '{wb_path}'.")
            workspace = self._get_workspace_from_wb(wb_dispatch)
            workspace.reset()
            logger.info(f"Workspace reset complete for '{wb_path}'.")
            
            # logger.info(f"Workspace '{uid}' already exists. Shutting down and re-initializing workspace.")
            # self.shutdown_workspace_from_dispatch(wb_dispatch)
            # logger.info(f"Re-initializing workspace...")
            # self.register_and_configure_wb_workspace(wb_dispatch)
            # logger.info(f"Re-initialization complete")

    def _get_workspace_from_wb(self, wb_dispatch):
        uid = self._get_workspace_pathuid_from_wb(wb_dispatch)
        if not uid in self._workspace_map.keys():
            logger.info("Workbook has not been registered, initializing...")
            self.register_and_configure_wb_workspace(wb_dispatch)
        _utils.comarshal_release_and_get_stream(wb_dispatch) # marshalling ok afaik
        return self._workspace_map[uid]
    
    def _get_workspace_pathuid_from_wb(self, wb_dispatch):
        wb:"xl._Workbook" = win32com.client.Dispatch(wb_dispatch)
        wb_path = str(Path(wb.FullName))
        _utils.comarshal_release_and_get_stream(wb) # marshalling ok afaik
        return wb_path
        # return utils.hash_str(wb_path)

    def execute_function_async(self, wb_dispatch, caller, func_name, *args):
        # marshalling ok afaik - excel vba interface
        workspace = self._get_workspace_from_wb(wb_dispatch)
        return workspace.execute_function_async(caller=caller, fname=func_name, args=args)

    def execute_sub_async(self, wb_dispatch, func_name):
        # marshalling ok afaik - excel vba interface
        workspace = self._get_workspace_from_wb(wb_dispatch)
        return workspace.execute_sub_async(fname=func_name)


    def register_functions_in_workspace(self, wb_dispatch):
        # marshalling ok afaik - excel vba interface
        workspace = self._get_workspace_from_wb(wb_dispatch)
        # workspace.register_functions_in_self()
        workspace.register_subs_in_self()
        pass

    def register_functions_in_vba(self, wb_dispatch):
        raise NotImplementedError("Obsoleted to remove combase.dll issue")

    def shutdown_workspace_from_dispatch(self, wb_dispatch):
        # XXX - todo - check this actually does anything meaninfgul
        # marshalling ok afaik - excel vba interface
        workspace = self._get_workspace_from_wb(wb_dispatch)
        uid = self._get_workspace_pathuid_from_wb(wb_dispatch)
        logger.info(f"Shutting down workspace uid:'{uid}'")
        workspace:xlproWorkspace
        workspace.shutdown()
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
        return workspace.get_vba_sync_text()
    def get_vba_sync_text_subs(self, wb_dispatch):
        workspace = self._get_workspace_from_wb(wb_dispatch)
        return workspace.get_vba_sync_text_subs()

    
    def get_workspace_from_uid_thread_safe(self, uid) -> xlproWorkspace:
        with self._workspace_uid_to_workbook_path_lock:
            with self._workspace_map_lock:
                return self._workspace_map[self._workspace_uid_to_workbook_path[uid]]


class xlproWorkspace:
    def __init__(self, server:xlproServer, wb_uid, uid):
        self._server = server
        self._wb_uid = wb_uid
        self._wb_path = Path(wb_uid)
        self._uid = uid
        self._wd = None # working directory

        self._uid_result_display_map_lock = threading.Lock()
        self._uid_result_display_map:dict=None

        self._uid_results_map_lock = threading.Lock()
        self._uid_results_map:dict=None

        self._uid_result_iscomplete_map_lock = threading.Lock()
        self._uid_result_iscomplete_map:dict=None

        self._uid_result_type_map_lock = threading.Lock()
        self._uid_result_type_map:dict=None

        self._uid_to_caller_map_lock = threading.Lock() # XXX - todo - not used currently
        self._uid_to_caller_map:dict=None

        self._uid_pending_function_map_lock = threading.Lock()
        self._uid_pending_function_map:dict=None

        self._caller_address_uid_map_lock = threading.Lock()
        self._caller_address_uid_map:dict=None

        self._uid_args_cache_lock = threading.Lock() # XXX - todo - not used.
        self._uid_args_cache:dict=None

        self.reset_workspace_cache()

        # maps the uid to the caller and function hash

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

        # self._sub_worker_manager = SubWorkerManager(server=self)
        # self._sub_worker_manager.start()

        self._func_hash_subthread_map = {}
        # self._func_hash_fig_generating_thread_map = {}

        self._temp_module_name:str = None

        self._module_function_maps_wrapper:ModuleFunctionMapsWrapper = None

    def set_xlpro_working_dir(self, wd:Path):
        logger.info(f"Setting working directory for workspace to '{str(wd)}'")
        # self._wd = initialize_and_get_workspace_xlpro_dir(self._wb_path)
        self._wd = get_workspace_xlpro_dir(self._wb_path)

    # def _configure_xlpro_files(self):
    #     # configure_workspace_xlpro_files(self._wd.parent, cfg)
    #     initialize_and_get_workspace_xlpro_dir(self._wb_path)

    def register_functions_in_self(self):
        logger.info(f"Registering workspace functions...")
        self._temp_module_name = f"{cfg.xlpro_functions_stem}_{self._uid}"
        _wrappers.import_module_with_registration(self._temp_module_name, self._wd / f"{cfg.xlpro_functions_stem}.py")
        self.update_module_func_map_wrapper()
        logger.info(f"Registration complete.")

    def register_subs_in_self(self):
        logger.info(f"Registering workspace subroutines...")
        self._sub_module_name = f"{cfg.xlpro_subroutines_stem}_{self._uid}"
        _wrappers.import_module_subs_with_registration(self._sub_module_name, self._wd / f"{cfg.xlpro_subroutines_stem}.py")
        self.update_module_sub_map_wrapper()
        logger.info(f"Registration complete.")

    def get_active_registered_functon_names(self):
        return [k for k, isactive in self._module_function_maps_wrapper.fname_isactive_register.items() if isactive]
    def get_active_registered_sub_names(self):
        return [k for k, isactive in self._module_sub_maps_wrapper.subname_isactive_register.items() if isactive]
    
    def get_active_registered_functions(self):
        keys = self.get_active_registered_functon_names()
        return [self._module_function_maps_wrapper.fname_func_register[key] for key in keys]
    def get_active_registered_subs(self):
        keys = self.get_active_registered_sub_names()
        return [self._module_sub_maps_wrapper.subname_func_register[key] for key in keys]
    
    def deregister_functions_in_self(self):
        logger.info(f"Deregistering workspace functions...")
        # self._valid_function_names = []
        if self._temp_module_name is not None:
            try:
                del sys.modules[self._temp_module_name]
            except KeyError:
                pass
        logger.info(f"Deregistration complete.")
    def deregister_subs_in_self(self):
        logger.info(f"Deregistering workspace subs...")
        # self._valid_function_names = []
        if self._temp_module_name is not None:
            try:
                del sys.modules[self._temp_module_name]
            except KeyError:
                pass
        logger.info(f"Deregistration complete.")

    def update_module_func_map_wrapper(self):
        self._module_function_maps_wrapper = ModuleFunctionMapsWrapper(self._temp_module_name)
    def update_module_sub_map_wrapper(self):
        self._module_sub_maps_wrapper = ModuleSubMapsWrapper(self._sub_module_name)


    def reset(self):
        # self._configure_xlpro_files()
        logger.info("Resetting workspace")
        self.deregister_functions_in_self()
        self.register_functions_in_self()

        self.deregister_subs_in_self()
        self.register_subs_in_self()

        self.reset_workspace_cache()

    def _get_function_by_name(self, fname):
        return self._module_function_maps_wrapper.fname_func_register[fname]
    def _get_sub_by_name(self, fname):
        return self._module_sub_maps_wrapper.subname_func_register[fname]

    def reset_workspace_cache(self):
        logger.debug("Initializing hashmaps")
        self._uid_result_display_map = {} # the result to be displayed
        self._uid_results_map = {} # the actual results
        self._uid_result_iscomplete_map = {}
        self._uid_result_type_map = {}
        self._uid_to_caller_map = {}
        self._uid_pending_function_map = {} # uid: func
        self._caller_address_uid_map = {} # uid: address
        self._uid_args_cache = {}


    def clear_uid(self, uid):
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
    def hash_excel_function_call(*args:typing.Iterable[str]):
        return _utils.hash_str(", ".join([str(x) for x in args]))

    def execute_sub_async(self, fname):
        try:
            func = self._get_sub_by_name(fname)
            if not func:
                raise Exception(f"Function {fname} not found.")

            import uuid
            uid = SUB_CALLER_FLAG_STRING + str(uuid.uuid4())

            logger.debug(f"Calling subroutine '{fname}', uid: '{uid}'")

            # xxx - todo - hack to signal the clientmanager to skip!
            caller_stream = SUB_CALLER_FLAG_STRING    
            
            with self._uid_to_caller_map_lock:
                self._uid_to_caller_map[uid] = caller_stream

            # configure default state for result and iscomplete status
            # with self._uid_result_display_map_lock:
                # self._uid_result_display_map[uid] = f"Promise<{uid}>"
            with self._uid_result_iscomplete_map_lock:
                self._uid_result_iscomplete_map[uid] = False

            f = self.create_worker_func_sub(uid, func, args=tuple(), kwargs={})
            with self._uid_pending_function_map_lock:
                self._uid_pending_function_map[uid] = f
            self._pending_function_queue.put(uid)
            self._worker_manager.wake()
            return
        
        # Return the python exception string as a fallback
        except Exception as e:
            logger.critical("error during execute_sub_async please rectify!")
            raise e
            # sys.exit()
            # return 
            # return repr(errors.xlproUnhandledException(repr(e)))
        
    def execute_function_async(self, caller, fname, args):
        try:
            func = self._get_function_by_name(fname)

            if not func:
                raise Exception(f"Function {fname} not found.")
            
            args = _utils.com_args_release_to_stream_reserved(func, args)      

            args_less_reserved = _utils.get_args_minus_reserved(func, args)

            uid = Dispatch(caller).Address + xlproWorkspace.hash_excel_function_call(fname, *args_less_reserved)

            logger.debug(f"Calling function '{fname}', uid: '{uid}', args: '{args}'")

            self._uid_args_cache[uid] = args

            # return the cached result if it exists
            with self._uid_result_display_map_lock:
                with self._uid_result_iscomplete_map_lock:
                    if self._uid_result_iscomplete_map.get(uid, False):
                        ret = self._uid_result_display_map[uid]
                        logger.debug(f"result marked complete, fetched cached result {ret}")
                        return ret
                
            # Clear any lingering calculations coming from this caller if the result isnt cached
            # downstream functions should pick up on these being deleted
            caller_dispatch = Dispatch(caller)
            caller_addr = caller_dispatch.Address

            with self._caller_address_uid_map_lock:
                # the hash will be constant for a function/args/caller combination so this is valid
                if caller_addr in self._caller_address_uid_map.keys():
                    self.clear_uid(self._caller_address_uid_map[caller_addr])
                self._caller_address_uid_map[caller_addr] = uid

            # release the com args for use in another thread. convert them to streams
            # args = utils.com_args_release_to_stream_reserved(func, args)

            result_type = self._module_function_maps_wrapper.fname_type_register[fname]
            
            with self._uid_result_type_map_lock:
                self._uid_result_type_map[uid] = result_type

            caller_stream = _utils.comarshal_release_and_get_stream(caller_dispatch) # this marshal is the OG
            with self._uid_to_caller_map_lock:
                self._uid_to_caller_map[uid] = caller_stream

            # configure default state for result and iscomplete status
            with self._uid_result_display_map_lock:
                self._uid_result_display_map[uid] = f"Promise<{uid}>"
            with self._uid_result_iscomplete_map_lock:
                self._uid_result_iscomplete_map[uid] = False

            # handle different function types
            # if result_type in [FunctionTypes.array_or_value, FunctionTypes.py_object]:
            if result_type in [FunctionTypes.array_or_value, FunctionTypes.py_object]:
                f = self.create_worker_func(uid, func, args, kwargs={})
                with self._uid_pending_function_map_lock:
                    self._uid_pending_function_map[uid] = f
                self._pending_function_queue.put(uid)
                self._worker_manager.wake()

            # return the (incomplete result)
            with self._uid_result_display_map_lock:
                return self._uid_result_display_map[uid]

        # Return the python exception string as a fallback
        except Exception as e:
            return repr(errors.xlproUnhandledException(repr(e)))
        
    def create_worker_func(self, uid, func, args, kwargs):
        if kwargs:
            raise Exception("kwargs should not be here!")
        pass
        # check for py object request
        # args0 = deepcopy(args)
        args = list(args)
        for i, arg in enumerate(args):
            if isinstance(arg, str):
                if m:=re.match(r"PyObj<(.*)>", arg):
                    with self._uid_results_map_lock:
                        temp_uid = m.group(1)
                        args[i] = self._uid_results_map[temp_uid]
        def worker():
            f = _wrappers.generate_wrapped_function(self._temp_module_name, func.__name__)
            try:
                ret = f(*args, **kwargs)
                self._result_queue.put((uid, ret))
                logger.info(f"Completed function '{uid}' successfully")
            except Exception as e:
                if isinstance(e, errors.ArugmentNotReadyException):
                    pass
                elif isinstance(e, errors.ExcelArugmentIsNoneException):
                    pass
                elif isinstance(e, errors.xlproArgumentExceptionError):
                    pass
                elif isinstance(e, errors.ExcelNotAccessibleError):
                    pass
                elif isinstance(e, errors.xlproUnhandledException):
                    pass
                elif isinstance(e, AttributeError):
                    if re.match(r"^<unknown>\..*$", str(e)):
                        logger.debug(f"error looks like a COM access error, modifying it from e={repr(e)}")
                        e = errors.xlproLikelyCOMAccessError(str(e))
                self._result_queue.put((uid, e))
                logger.info(f"Completed function '{uid}' unsuccessfully with error {e}")
            logger.debug("Waking results manager from worker thread...")
            self._results_manager_thread.wake()
        
        return worker
    
    def create_worker_func_sub(self, uid, func, args, kwargs):
        def worker():
            try:
                ret = func(*args, **kwargs)
                self._result_queue.put((uid, ret))
                logger.info(f"Completed function '{uid}' successfully")
            except Exception as e:
                self._result_queue.put((uid, e))
                logger.info(f"Completed function '{uid}' unsuccessfully with error {e}")
            logger.debug("Waking results manager from worker thread...")
            self._results_manager_thread.wake()
        
        return worker
            
    def shutdown(self):
        for uid, t in self._func_hash_subthread_map.items():
            t:threading.Thread
            if t.is_alive():
                t.join()
        self._results_manager_thread.stop()
        self._client_manager_thread.stop()
        # for uid, p in self._func_hash_fig_generating_thread_map.items():
        #     p:FigureGeneratingThread
        #     p.stop()

    def get_vba_sync_text(self) -> str:
        funcs = self.get_active_registered_functions()
        return _utils.get_xlpro_vb_dynamic_component_contents(funcs)
    def get_vba_sync_text_subs(self) -> str:
        subs = self.get_active_registered_subs()
        return _utils.get_xlpro_vb_dynamic_component_contents_subs(subs)
    
    def get_caller_stream(self, uid):
        with self._uid_to_caller_map_lock:
            return self._uid_to_caller_map[uid]
    
    def set_caller_stream(self, uid, val):
        with self._uid_to_caller_map_lock:
            self._uid_to_caller_map[uid] = val

    def get_uid_debug_info(self, uid):
        with self._uid_result_display_map_lock:
            a0 = self._uid_result_display_map.get(uid, None)
        with self._uid_results_map_lock:
            a1 = self._uid_results_map.get(uid, None)
        with self._uid_result_iscomplete_map_lock:
            a2 = self._uid_result_iscomplete_map.get(uid, None)
        with self._uid_result_type_map_lock:
            a3 = self._uid_result_type_map.get(uid, None)
        with self._uid_pending_function_map_lock:
            a4 = self._uid_pending_function_map.get(uid, None)
        pythoncom.CoInitialize()
        try:
            rng_stream = self.get_caller_stream(uid)
            rng_dispatch:"xl.Range" = _utils.comarshal_dispatch_stream(rng_stream)
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
    
    def get_uid_of_val_thread_safe(self, val) -> str:
        with self._uid_results_map_lock:
            for k, v in self._uid_results_map.items():
                if v == val:
                    return k
        raise KeyError("value is not present within the uid_results_map")
        
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
        self._threadpool_dict_lock = threading.Lock()

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
        with self._threadpool_dict_lock:
            n_items = len(self._threadpool_dict.keys())
            
        if n_items < self.MAX_THREADS:
            uid = self._server._pending_function_queue.get(timeout=0.01)
            try:
                with self._server._uid_pending_function_map_lock:
                    func = self._server._uid_pending_function_map[uid]
            except KeyError:
                logger.warning(f"Pending function queue uid not available, ignoring calculation request for uid '{uid}'")
                return
            
            with self._threadpool_dict_lock:
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
            # logger.debug("WorkerManager thread is waiting for events or timeout...")
            self._wake_event.wait(timeout=0.1)
            if self._wake_event.is_set():
                self._wake_event.clear()
                logger.info("WorkerManager thread woke up for an event!")
            while True:
                try:
                    self._process_function_queue()
                except queue.Empty:
                    break
                self._clear_completed_threads()
                time.sleep(0.01) # fairness sleep

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
        active_formulas = self._server.get_active_registered_functon_names()

        u = [x.AddressLocal for x in precedents_stream]

        for ps_disp in precedents_stream:
            try:
                ps_disp:"xl.Range"
                formula = ps_disp.Formula2
                ps_disp.Formula2 = ps_disp.Formula2
                if "pd_function_create" in formula:
                    pass
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

    def _emergency_recalculate(self, uid) -> None:
        """force a recalculate"""
        raise NotImplementedError("please god dont use this looks dodgy")
        pythoncom.CoInitialize()
        try:
            caller_dispatch = _utils.comarshal_dispatch_stream(self._server.get_caller_stream(uid))
            caller_dispatch.Formula2 = caller_dispatch.Formula2
            self._server.set_caller_stream(uid, _utils.comarshal_release_and_get_stream(caller_dispatch))
        except KeyError as e:
            logger.error(f"Error during forced client update: '{uid}', {e}")
        pythoncom.CoUninitialize()

    def signal_worker_manager_to_recalculate(self, uid):
        logger.debug(f"ResultsManager is waking the worker manager to recycle '{uid}'")
        self._server._pending_function_queue.put(uid)
        self._server._worker_manager.wake()

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
            # if isinstance(val, errors.ArugmentNotReadyException):
            if isinstance(val, (errors.ArugmentNotReadyException, errors.ExcelArugmentIsNoneException)):
                logger.debug(f"Arguments not ready for uid '{uid}', recycling function...")
                self._server.get_uid_debug_info(uid)
                self.signal_worker_manager_to_recalculate(uid)
                return
            
            elif isinstance(val, errors.xlproArgumentExceptionError):
                # if an argument is an exception, do not recycle as a pending function
                # when the root failed one re-executes, it will recalculate the dependents
                pass
          
            elif isinstance(val, errors.ExcelNotAccessibleError):
                # recycle if excel is not accessible
                logger.debug(f"ExcelNotAccessible during '{uid}', recycling function...")
                self.signal_worker_manager_to_recalculate(uid)
                return

            elif isinstance(val, pythoncom.com_error):
                # recycle if excel is not accessible
                if VBErrorConverter(val) == VBError.xlCallRejectedByCallee:
                    logger.info(f"Call rejected by callee for '{uid}', recycling function...")
                    self.signal_worker_manager_to_recalculate(uid)
                    return
                elif VBErrorConverter(val) == 285: # The marshaled interface data packet (OBJREF) has an invalid or unknown format.
                    logger.info(f"OBJREF invalid '{uid}', recycling function...")
                    self.signal_worker_manager_to_recalculate(uid)
                    return
                elif VBErrorConverter(val) == 30: # A disk error occurred during a read operation.
                    logger.info(f"disk read failed '{uid}', recycling function...")
                    self.signal_worker_manager_to_recalculate(uid)
                    return
                else:
                    logger.info(f"Other COM error for '{uid}', recycling function...")
                    logger.info(f"'{repr(val)}'")
                    self.signal_worker_manager_to_recalculate(uid)
                    return


            elif isinstance(val, errors.xlproLikelyCOMAccessError):
                logger.warning(f"Likely COM access error for {uid}, recycling.")
                self.signal_worker_manager_to_recalculate(uid)
                return
            
            elif isinstance(val, UnboundLocalError):
                logger.warning(f"UnboundLocalError {uid}, recycling... (TODO fix this)")
                self.signal_worker_manager_to_recalculate(uid)
                pass

            logger.warning(f"Returned value is a generic exception: {uid}, {val}")
            # ret = repr(val) # convert exception to string for it to show in excel.

        # if it is a valid return, write the result to the cache
        self._set_results_value(uid, ret)
        # signal that it is complete
        with self._server._uid_result_iscomplete_map_lock:
            logger.debug(f"calculation marked complete {uid}")
            try:
                if "promise" in ret.lower():
                    pass
            except:
                pass
            self._server._uid_result_iscomplete_map[uid] = True
        # remove from the pending function map once successfully completed.
        with self._server._uid_pending_function_map_lock:
            logger.debug(f"calculation removed as pending function {uid}")
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
            # logger.info("ResultsManager thread is waiting for events or timeout...")
            self._wake_event.wait(timeout=0.1)
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
        
    def _update_client_pyobject_result(self, uid) -> None:
        """Update the data for the py_object case case (row-major arrays, strings, values)"""
        try:
            caller_dispatch = _utils.comarshal_dispatch_stream(self._server.get_caller_stream(uid))
            val = self._get_value(uid)
            if isinstance(val, Exception):
                logger.warning(f"Value is an exception: '{val}', '{uid}'")
                self._set_result_display(uid, repr(val))
            else:
                self._set_result_display(uid, f"PyObj<{uid}>")

            # update by resetting the formula
            caller_dispatch.Formula2 = caller_dispatch.Formula2
            # self._server.set_caller_stream(uid, _utils.comarshal_release_and_get_stream(caller_dispatch))
        except KeyError as e:
            # XXX - todo - there is a risk of a keyerror here for some reason
            logger.error(f"Error during client update: '{uid}', {e}")
        finally:
            try:
                self._server.set_caller_stream(uid, _utils.comarshal_release_and_get_stream(caller_dispatch))
            except:
                pass
        
    def _update_client_default_result(self, uid) -> None:
        """Update the data for the default case (row-major arrays, strings, values)"""
        try:
            caller_dispatch = _utils.comarshal_dispatch_stream(self._server.get_caller_stream(uid))
            val = self._get_value(uid)
            if isinstance(val, Exception):
                logger.warning(f"Value is an exception: '{val}', '{uid}'")
                self._set_result_display(uid, repr(val))
            else:
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

    def _update_client_image_result(self, uid) -> None:
        """Update the data for the figure case - add a figure image to the spreadsheet"""
        caller_dispatch = _utils.comarshal_dispatch_stream(self._server.get_caller_stream(uid))
        try:
            val = self._get_value(uid)
            if not type(val) == xlproImage:
                raise TypeError
            
            val:xlproImage
            
            fp, size_pt, xl_name = val.fp, val.size_pt, val.xl_name

            if not fp.suffix.lower()[1:] in ("emf","wmf","jpg","jpeg","jff","jpe","png","bmp","dib","rle","gif","emz","wmz","tif","tiff","svg","ico","webp"):
                raise Exception(f"Excel does not support this extension {fp.suffix}")

            caller_adjacent = caller_dispatch.Cells(2,1)
            xpos, ypos = caller_adjacent.Left, caller_adjacent.Top
            width, height = size_pt

            ws = caller_adjacent.Parent
            # ws.Shapes.AddPicture(str(fp.resolve()), False, True, xpos, ypos, width, height)

            # look for an existing shape with the same name
            try:
                shape = ws.Shapes(xl_name)
                xpos, ypos = shape.Left, shape.Top
                shape.Delete()
            except pythoncom.com_error as e:
                logger.warning(f"could not find object with name: {xl_name} to delete")

            shape = ws.Shapes.AddPicture(str(fp.resolve()), False, True, xpos, ypos, width, height)
            shape.Name = xl_name

            # self._set_result_display(uid, f"Image<{fp}>")
            self._set_result_display(uid, f"Image<{xl_name}>")

            # update by resetting the formula
            caller_dispatch.Formula2 = caller_dispatch.Formula2
            self._server.set_caller_stream(uid, _utils.comarshal_release_and_get_stream(caller_dispatch))

        except Exception as e:
            self._server.set_caller_stream(uid, _utils.comarshal_release_and_get_stream(caller_dispatch))
            raise e

    def _process_queue(self):
        """Process the queue at the current point in time. Any failed attempts get
        added back into the queue if appropriate.
        """
        # XXX - Warning that qsize() is not thread safe
        # this prevents inplace recycling and potentially infinite loop
        queue_length = self._server._client_recalculate_queue.qsize()
        # logger.debug(f"Client manager queue length estimate is {queue_length}")
        for _ in range(queue_length):
            try:
                uid = self._server._client_recalculate_queue.get()
                # logger.debug(f"Client manager fetched uid '{uid}'")
            except queue.Empty:
                # logger.debug("Client manager queue is empty")
                return
            
            with self._server._uid_result_iscomplete_map_lock:
                if not uid in self._server._uid_result_iscomplete_map.keys():
                    logger.debug(f"Uid '{uid}' not found as pending function within client manager update. Ignoring this iteration")
                    return

            pass

            # xxx - todo - hack to skip checks on subroutines
            # subroutines will never need to be recycled by the client manager
            # subroutines will be recycled by the results manager since there will
            # never be any result to reach the client.
            # so we are ok to just skip the loop here, happy days.
            caller_dispatch = self._server.get_caller_stream(uid)
            if caller_dispatch is SUB_CALLER_FLAG_STRING:
                logger.debug("caller_dispatched checked as None, assuming this is a subroutine and skipping further.")
                time.sleep(0.01) # copy the fairness sleep
                continue
            
            # fetch the result type so we know how to handle it
            with self._server._uid_result_type_map_lock:
                result_type = self._server._uid_result_type_map[uid]

            # if we fail to dispatch the range, it was probably deleted.
            # therefore we don't need to replace it in the queue
            replace_in_queue = True

            # Decide whether to recycle
            try:

                value = self._get_value(uid)
                if type(value) == xlproImage:
                    self._update_client_image_result(uid)
                elif isinstance(value, matplotlib.figure.Figure):
                    self._update_client_pyobject_result(uid)

                elif result_type == FunctionTypes.array_or_value:
                    self._update_client_default_result(uid)
                elif result_type == FunctionTypes.py_object:
                    self._update_client_pyobject_result(uid)
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
                else:
                    logger.debug(f"Other COM Error occurred, {e}, recycling in queue")

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
                self._server._client_recalculate_queue.put(uid)

            time.sleep(0.01) # fairness sleep


    def _run(self):
        """When awakened, retrieves the result queue and sends to the server cache.
        """
        pythoncom.CoInitialize()

        while not self._stop_event.is_set():
            # wait for ten seconds for an event, otherwise do a queue process to check for 
            # outstanding tasks.
            self._wake_event.wait(timeout=0.1)
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
    pass