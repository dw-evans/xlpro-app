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

CFG = config.load()
wd = Path(__file__).parent
logger = logging.getLogger(__name__)

from xlpro import _utils
from xlpro._wrappers import ModuleFunctionMapsWrapper, ModuleSubMapsWrapper
from xlpro._enums import FunctionTypes
from xlpro import _wrappers
import re
from copy import deepcopy

from xlpro._types import xlproImage, xlproExpandedType, xlproCollapsedType
from xlpro._types import ndarray1d, ndarray2d, list1d, list2d

from win32com.client.dynamic import Dispatch
from functools import wraps
from xlpro._types import ExcelArrayConverter

SLEEP_DURATION = 0.01

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

def force_clear_queue(q: queue.Queue):
    with q.mutex:
        q.queue.clear()


SUB_CALLER_FLAG_STRING = "SUB_CALLER_FLAG_STRING"

class xlproServer:
    _public_methods_ = [
        # "getpid",

        "register_and_configure_wb_workspace",
        "register_functions_in_vba",
        "register_functions_in_workspace",

        "execute_function_async",
        "execute_sub_async",
        "force_refresh_area_calculation",

        # "shutdown_workspace",
        # "shutdown",

        "get_vba_sync_text",
        "get_vba_sync_text_subs",

        # "__dev_shutdown",
    ]
    # _reg_progid_ = config.progid
    _reg_clsid_ = "undefined"

    _instance = None  # Singleton instance
    _instance_initialized = False

    # use this flag to tell the server to close it next time
    # it is checked.
    _is_pending_close = False

    # @staticmethod
    # def comsafe(func):
    #     @wraps(func)
    #     def inner(*args, **kwargs):
    #         with XLAPP_LOCK_MAINTHREAD:
    #             return func(*args, **kwargs)
    #     return inner

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
    
    @_utils.traceback_log_raise
    # @_utils.comsafe
    def register_fnames_in_workbook(self, workspace:xlproWorkspace, wb_dispatch):
        """Registers fname_* in the workbook names so the user knows what names
        Are registered"""


        def _xlinteract1():
            logger.info("Replacing workbook names beginning with reserved string `fname_`, `args_`...")
            logger.info("Adding names to workbook Name Manager...")
            logger.info("Adding pyNone and pyEmpty to workbook names...")
            wb_dispatch.Names.Add(_utils.XLPRO_EMPTY_STR, f"=\"{_utils.XLPRO_EMPTY_STR}\"")
            wb_dispatch.Names.Add(_utils.XLPRO_NONE_STR, f"=\"{_utils.XLPRO_NONE_STR}\"")

            logger.info("Adding function names to workbook names...")
            count = 0
            for name in wb_dispatch.Names:
                if name.Name.startswith("fname_"):
                    wb_dispatch.Names(name.Name).Delete()
                    count += 1
            logger.debug(f"Deleted {count} names")
            
            logger.info("Adding function names...")
            count = 0
            for fname in workspace.get_active_registered_functon_map().keys():
                wb_dispatch.Names.Add(f"fname_{fname}", f"=\"{fname}\"")
                count += 1
            logger.info(f"Added {count} function names")

            f_map = workspace.get_active_registered_functon_map()

            logger.info("Adding argument arrays to workbook names...")
            for name in wb_dispatch.Names:
                if name.Name.startswith("args_"):
                    wb_dispatch.Names(name.Name).Delete()

            count = 0
            for k, v in f_map.items():
                args:list[str] = _utils.get_excel_args_of_func(v)
                if len(args) == 0:
                    s = "=\"\""
                elif len(args) == 1:
                    s = f"=\"{args[0]}\""
                else:
                    s = "={{{}}}".format(";".join([f"\"{x}\"" for x in args]))
                (a:=f"args_{k}", b:=f"{s}")
                wb_dispatch.Names.Add(a, b)
                count += 1
                pass
            logger.info(f"Added {count} argument array names")
            logger.info("Workbook name registration complete")

                

        # _xlinteract1()
        _utils.comsafe(_xlinteract1)()
        
        return

    @_utils.traceback_log_raise
    # @_utils.comsafe
    def register_and_configure_wb_workspace(self, wb_dispatch):
        # marshalling ok afaik - excel vba interface
        wb_path = self._get_workspace_pathuid_from_wb(wb_dispatch)

        if not wb_path in self._workspace_map.keys():
            logger.info(f"Creating new workspace for '{Path(wb_path).name}'...")
            uid = _utils.hash_str(wb_path)
            workspace = xlproWorkspace(self, wb_uid=wb_path, uid=uid)
            workspace_wd = Path(wb_path).parent.resolve()

            workspace.set_xlpro_working_dir(workspace_wd)
            workspace.reset()
            self.register_fnames_in_workbook(workspace=workspace, wb_dispatch=Dispatch(wb_dispatch))


            self._workspace_map[wb_path] = workspace
            self._workspace_uid_to_workbook_path[uid] = wb_path
            logger.info(f"Creation of workspace complete for '{wb_path}'")

            pass
        else:
            logger.info(f"Workspace already exists. Resetting...")
            logger.info(f"Resetting workspace '{Path(wb_path).name}'...")
            workspace = self._get_workspace_from_wb(wb_dispatch)
            workspace.reset()
            self.register_fnames_in_workbook(workspace=workspace, wb_dispatch=Dispatch(wb_dispatch))
            logger.info(f"Workspace reset complete for '{Path(wb_path).name}'")
            
            # logger.info(f"Workspace '{uid}' already exists. Shutting down and re-initializing workspace.")
            # self.shutdown_workspace_from_dispatch(wb_dispatch)
            # logger.info(f"Re-initializing workspace...")
            # self.register_and_configure_wb_workspace(wb_dispatch)
            # logger.info(f"Re-initialization complete")

    # @_utils.comsafe
    def _get_workspace_from_wb(self, wb_dispatch):
        uid = self._get_workspace_pathuid_from_wb(wb_dispatch)
        if not uid in self._workspace_map.keys():
            logger.info("Workbook has not been registered, initializing...")
            self.register_and_configure_wb_workspace(wb_dispatch)
            logger.info("Initialization complete.")
        _utils.comarshal_release_and_get_stream(wb_dispatch) # marshalling ok afaik
        return self._workspace_map[uid]
    
    # @_utils.comsafe
    def _get_workspace_pathuid_from_wb(self, wb_dispatch):
        wb:"xl._Workbook" = win32com.client.Dispatch(wb_dispatch)
        wb_path = str(Path(wb.FullName))
        _utils.comarshal_release_and_get_stream(wb) # marshalling ok afaik
        return wb_path
        # return utils.hash_str(wb_path)


    @_utils.traceback_log_raise
    # @_utils.comsafe
    def force_refresh_area_calculation(self, wb_dispatch, rng):
        workspace = self._get_workspace_from_wb(wb_dispatch)
        rng_dispatch = Dispatch(rng)
        shtname = rng_dispatch.Parent.Name
        for area in rng_dispatch.Areas:
            for cell in area.Cells:
                workspace.force_clear_addr(sheetaddr=shtname + cell.Address)
            _utils.comsafe(lambda: rng_dispatch.Application.Run("'xlpro.xlam'!AtomicFormulaRefreshNoEvents", area))()
            # rng_dispatch.Application.Run("'xlpro.xlam'!AtomicFormulaRefreshNoEvents", area)


    # @_utils.comsafe
    def execute_function_async(self, wb_dispatch, caller, func_name, *args):
        # marshalling ok afaik - excel vba interface
        workspace = self._get_workspace_from_wb(wb_dispatch)
        ret = workspace.execute_function_async(caller=caller, fname=func_name, args=args)
        # ret = workspace.execute_function_sync(caller=caller, fname=func_name, args=args)
        return ret

    # @_utils.comsafe
    def execute_sub_async(self, wb_dispatch, func_name):
        # marshalling ok afaik - excel vba interface
        workspace = self._get_workspace_from_wb(wb_dispatch)
        return workspace.execute_sub_async(fname=func_name)


    # @_utils.comsafe
    def register_functions_in_workspace(self, wb_dispatch):
        # marshalling ok afaik - excel vba interface
        workspace = self._get_workspace_from_wb(wb_dispatch)
        # workspace.register_functions_in_self()
        workspace.register_subs_in_self()
        pass

    def register_functions_in_vba(self, wb_dispatch):
        raise NotImplementedError("Obsoleted to remove combase.dll issue")

    # @_utils.comsafe
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

    @_utils.traceback_log_raise
    # @_utils.comsafe
    def get_vba_sync_text(self, wb_dispatch):
        """Gets the vba code module contents to register the udfs"""
        # marshalling ok afaik - excel vba interface
        workspace = self._get_workspace_from_wb(wb_dispatch)
        return workspace.get_vba_sync_text()
    
    @_utils.traceback_log_raise
    # @_utils.comsafe
    def get_vba_sync_text_subs(self, wb_dispatch):
        workspace = self._get_workspace_from_wb(wb_dispatch)
        return workspace.get_vba_sync_text_subs()

    # @_utils.comsafe
    def get_workspace_from_uid_thread_safe(self, uid) -> xlproWorkspace:
        with self._workspace_uid_to_workbook_path_lock:
            with self._workspace_map_lock:
                return self._workspace_map[self._workspace_uid_to_workbook_path[uid]]


CALLER_THROTTLE_TIME_NS = 200_000_000

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
        
        self._uid_subresults_map_lock = threading.Lock()
        self._uid_subresults_map:dict[str, tuple]=None
        
        self._uid_subresults_display_map_lock = threading.Lock()
        self._uid_subresults_display_map:dict[str, tuple]=None

        self._uid_result_iscomplete_map_lock = threading.Lock()
        self._uid_result_iscomplete_map:dict=None # {uid: [subres0, subres1, ..., subresN]}

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

        self._caller_addr_timer_map_lock = threading.Lock() # XXX - todo - not used.
        self._caller_addr_timer_map:dict=None

        self._caller_function_count_map_lock = threading.Lock()
        self._caller_function_count_map:dict=None

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
        self.reset_workspace_cache()

    def set_xlpro_working_dir(self, wd:Path):
        logger.info(f"Setting working directory for workspace '{self._wb_path.name}'... ")
        # self._wd = initialize_and_get_workspace_xlpro_dir(self._wb_path)
        self._wd = get_workspace_xlpro_dir(self._wb_path)
        logger.info(f"Working directory for '{self._wb_path.name}' set at '{str(wd)}'")

    # def _configure_xlpro_files(self):
    #     # configure_workspace_xlpro_files(self._wd.parent, cfg)
    #     initialize_and_get_workspace_xlpro_dir(self._wb_path)

    def register_functions_in_self(self):
        logger.info(f"Registering workspace functions...")
        self._temp_module_name = f"{CFG.xlpro_functions_stem}_{self._uid}"
        _wrappers.import_module_with_registration(self._temp_module_name, self._wd / f"{CFG.xlpro_functions_stem}.py")
        self.update_module_func_map_wrapper()
        logger.info(f"Workspace functions registered")

    def register_subs_in_self(self):
        logger.info(f"Registering workspace subroutines...")
        self._sub_module_name = f"{CFG.xlpro_subroutines_stem}_{self._uid}"
        _wrappers.import_module_subs_with_registration(self._sub_module_name, self._wd / f"{CFG.xlpro_subroutines_stem}.py")
        self.update_module_sub_map_wrapper()
        logger.info(f"Workspace subroutines registered")

    def get_active_registered_functon_map(self):
        return {k: v for k, v in self._module_function_maps_wrapper.fname_func_register.items() if self._module_function_maps_wrapper.fname_isactive_register[k]}
        # return [k for k, isactive in self._module_function_maps_wrapper.fname_isactive_register.items() if isactive]
    def get_active_registered_sub_names(self):
        return [k for k, isactive in self._module_sub_maps_wrapper.subname_isactive_register.items() if isactive]
    
    def get_active_registered_functions(self):
        keys = self.get_active_registered_functon_map()
        return [self._module_function_maps_wrapper.fname_func_register[key] for key in keys]
    def get_active_registered_subs(self):
        keys = self.get_active_registered_sub_names()
        return [self._module_sub_maps_wrapper.subname_func_register[key] for key in keys]
    
    def deregister_functions_in_self(self):
        logger.info(f"De-registering workspace functions...")
        # self._valid_function_names = []
        if self._temp_module_name is not None:
            try:
                del sys.modules[self._temp_module_name]
            except KeyError:
                pass
        logger.info(f"Workspace function de-registration complete")
    def deregister_subs_in_self(self):
        logger.info(f"De-registering workspace subroutines...")
        # self._valid_function_names = []
        if self._temp_module_name is not None:
            try:
                del sys.modules[self._temp_module_name]
            except KeyError:
                pass
        logger.info(f"Workspace subroutine de-registration complete")

    def update_module_func_map_wrapper(self):
        self._module_function_maps_wrapper = ModuleFunctionMapsWrapper(self._temp_module_name)
    def update_module_sub_map_wrapper(self):
        self._module_sub_maps_wrapper = ModuleSubMapsWrapper(self._sub_module_name)


    def reset(self):
        # self._configure_xlpro_files()
        logger.info(f"Resetting workspace '{self._wb_path.name}'...")
        self.deregister_functions_in_self()
        self.register_functions_in_self()

        self.deregister_subs_in_self()
        self.register_subs_in_self()

        self.reset_workspace_cache()
        logger.info("Workspace reset complete...")

    def _get_function_by_name(self, fname):
        try:
            return self._module_function_maps_wrapper.fname_func_register[fname]
        except:
            raise KeyError

    def _get_sub_by_name(self, fname):
        try:
            return self._module_sub_maps_wrapper.subname_func_register[fname]
        except:
            raise KeyError

    def reset_workspace_cache(self):
        logger.debug("Initializing hashmaps")
        with self._uid_result_display_map_lock:
            self._uid_result_display_map = {} # the result to be displayed
        with self._uid_results_map_lock:
            self._uid_results_map = {} # the actual results
        with self._uid_result_iscomplete_map_lock:
            self._uid_result_iscomplete_map = {}
        with self._uid_result_type_map_lock:
            self._uid_result_type_map = {}
        with self._uid_to_caller_map_lock:
            self._uid_to_caller_map = {}
        with self._uid_pending_function_map_lock:
            self._uid_pending_function_map = {} # uid: func
        with self._caller_address_uid_map_lock:
            self._caller_address_uid_map = {} # uid: address
        with self._uid_args_cache_lock:
            self._uid_args_cache = {}
        with self._uid_subresults_map_lock:
            self._uid_subresults_map = {}
        with self._uid_subresults_display_map_lock:
            self._uid_subresults_display_map = {}
        with self._caller_addr_timer_map_lock:
            self._caller_addr_timer_map = {}
        with self._caller_function_count_map_lock:
            self._caller_function_count_map = {}

        self._worker_manager._clear_all_threads()
        force_clear_queue(self._pending_function_queue)
        force_clear_queue(self._result_queue)
        force_clear_queue(self._client_recalculate_queue)

        # self._worker_manager.clear_futures()


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

        try:
            with self._uid_subresults_map_lock:
                del self._uid_subresults_map[uid]
        except:
            pass
        try:
            with self._uid_subresults_display_map_lock:
                del self._uid_subresults_display_map[uid]
        except:
            pass


    @staticmethod
    def hash_excel_function_call(*args:typing.Iterable[str]):
        return _utils.hash_str(", ".join([str(x) for x in args]))
        # import random
        # return _utils.hash_str(", ".join([str(x if x is not None else random.random()) for x in args]))
        # return _utils.hash_str(", ".join([str(x) for x in args]))

    def execute_sub_async(self, fname):
        try:
            func = self._get_sub_by_name(fname)
            if not func:
                raise Exception(f"Function {fname} not found.")

            import uuid
            uid = SUB_CALLER_FLAG_STRING + str(uuid.uuid4())

            logger.info(f"Calling subroutine '{fname}'...")
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
            logger.critical(f"Error encountered during subroutine execution '{fname}'")
            raise e
            # sys.exit()
            # return 
            # return repr(errors.xlproUnhandledException(repr(e)))
    
    def force_clear_addr(self, sheetaddr):
        with self._caller_function_count_map_lock:
            uids = self._caller_function_count_map.get(sheetaddr, [])
            for uid in uids:
                self.clear_uid(uid)


    def execute_function_async(self, caller, fname, args):
        try:
            func = self._get_function_by_name(fname)

            if fname == "create_table_if_not_exists":
                pass

            if not func:
                raise Exception(f"Function {fname} not found.")
            
            args = _utils.com_args_release_to_stream_reserved(func, args)      

            args_less_reserved = _utils.get_args_minus_reserved(func, args)

            caller_dispatch = Dispatch(caller)
            caller_addr = caller_dispatch.Address
            caller_sheetaddr = caller_dispatch.Parent.Name + caller_dispatch.Address

            # calling_time = time.time_ns()

            uid = caller_sheetaddr + xlproWorkspace.hash_excel_function_call(fname, *args_less_reserved)

            logger.info(f"Calling function '{fname}' from '{caller_sheetaddr}'...")

            logger.debug(f"Calling function '{fname}', uid: '{uid}', args: '{args}'")

            self._uid_args_cache[uid] = args

            # return the cached result if it exists
            with self._uid_result_display_map_lock:
                with self._uid_result_iscomplete_map_lock:
                    # return_cached = False
                    if self._uid_result_iscomplete_map.get(uid, False):
                        if f"{uid}_expanded" in self._uid_result_display_map:
                            ret = self._uid_result_display_map[f"{uid}_expanded"]
                        else:
                            ret = self._uid_result_display_map[uid]
                        # if isinstance(ret, str):
                        #     if ret.startswith("Promise"):
                        #         logger.warning("Promise is marked complete, ignoring cache")
                        #     else:
                        #         return_cached = True
                        # if return_cached:
                        #     logger.debug(f"result marked complete, fetched cached result {ret}")
                        #     return ret
                        logger.debug(f"result marked complete, fetched cached result {ret}")
                        return ret
                

            # if the number of called functions from a caller exceeds a threshold, pop off the left uid 
            # and clear it 
            with self._caller_function_count_map_lock:
                if not caller_sheetaddr in self._caller_function_count_map.keys():
                    self._caller_function_count_map[caller_sheetaddr]  = [uid]
                else:
                    l:list = self._caller_function_count_map[caller_sheetaddr]
                    l.append(uid)
                    if len(l) > 32:
                        spent_uid = l.pop(0)
                        self.clear_uid(spent_uid)


            # When to wipe an existing calculation...
            # Current process: 
            # send caller, fname, args to xlpro server
            # hash addr+fname+args to create a uid for that function call
            # if there are any existing calculations for the same cell, wipe them.
            # Solutions:
            # Do not wipe any calculations, will cost speed
            # Find way to identify sub calls from a cell

            # with self._caller_address_uid_map_lock:
            #     # the hash will be constant for a function/args/caller combination so this is valid
            #     if caller_addr in self._caller_address_uid_map.keys():
            #         # XXX - todo - this chain will wipe nested calculations within the same cell
            #         # Even if we check which function is being executed we would still fail if the same
            #         # nested function call occurs from the same cell.
            #         # The function hash might pay to be generated from vba using cell range addrs
            #         # Then we canheck if ...

            #         # race condition hack, sleep if the last call from the cell was too soon!
            #         # with self._caller_addr_timer_map_lock:
            #         #     if not caller_addr in self._caller_addr_timer_map:
            #         #         self._caller_addr_timer_map[caller_addr] = calling_time
            #         #     else:
            #         #         dt = self._caller_addr_timer_map[caller_addr] - calling_time
            #         #         if dt < CALLER_THROTTLE_TIME_NS:
            #         #             time.sleep(dt/1e9)

            #         self.clear_uid(self._caller_address_uid_map[caller_addr])
            #     self._caller_address_uid_map[caller_addr] = uid

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
        

    def _handle_default_result(self, _ret, _uid):
        raise NotImplementedError
        """Default result handler, translated from client manager processing loop"""
        if type(_ret) == xlproExpandedType:
            val = _ret.data
            if type(val) == list:
                val = ExcelArrayConverter(val, list2d)
            elif type(val) == np.ndarray:
                val = ExcelArrayConverter(val, ndarray2d)
        else:
            val = _ret
        if isinstance(val, Exception):
            logger.warning(f"Value is an exception: '{val}', '{_uid}'")
            self._client_manager_thread._set_result_display(_uid, retval:=repr(val))
        else:
            self._client_manager_thread._set_result_display(_uid, retval:=val)
        return retval

    # pyobj result
    def _handle_pyobject_result(self, _ret, _uid):
        raise NotImplementedError
        """Pyobject result handler, translated from client manager processing loop"""
        if isinstance(_ret, xlproCollapsedType):
            val = _ret.data
        else:
            val = _ret
        if isinstance(val, Exception):
            logger.warning(f"Value is an exception: '{val}', '{_uid}'")
            self._client_manager_thread._set_result_display(_uid, retval:=repr(val))
        else:
            self._client_manager_thread._set_result_display(_uid, retval:=f"PyObj<{_uid}>")
        return retval
        
    # iterable result
    def _handle_iterable_result(self, _ret, _uid):
        raise NotImplementedError
        """Iterable result handler, translated from client manager processing loop"""
        if isinstance(_ret, Exception):
            logger.warning(f"Value is an exception: '{_ret}', '{_uid}'")
            self._client_manager_thread._set_result_display(_uid, retval:=repr(_ret))

        # construct a list off the iterable value
        elif isinstance(_ret, typing.Iterable):
            val_modified = []
            for i, subval in enumerate(_ret):
                if isinstance(subval, (int, float, str, bool)):
                    val_modified.append(subval)
                else:
                    val_modified.append(f"PyObj<{_uid}>_{i}")

            # create a *_expanded variant of the uid result display
            self._client_manager_thread._set_result_display(
                f"{_uid}_expanded", retval:=np.array(val_modified, dtype=object)
            )
        return retval


    def execute_function_sync(self, caller, fname, args):
        raise NotImplementedError
        try:
            func = self._get_function_by_name(fname)

            if not func:
                raise Exception(f"Function {fname} not found.")
            
            args = _utils.com_args_release_to_stream_reserved(func, args)      
            args_less_reserved = _utils.get_args_minus_reserved(func, args)

            uid = Dispatch(caller).Address + xlproWorkspace.hash_excel_function_call(fname, *args_less_reserved)
            logger.debug(f"Calling function '{fname}', uid: '{uid}', args: '{args}'")

            with self._uid_args_cache_lock:
                self._uid_args_cache[uid] = args

            # return the cached result if it exists
            with self._uid_result_display_map_lock:
                with self._uid_result_iscomplete_map_lock:
                    if self._uid_result_iscomplete_map.get(uid, False):
                        if f"{uid}_expanded" in self._uid_result_display_map:
                            ret = self._uid_result_display_map[f"{uid}_expanded"]
                        else:
                            ret = self._uid_result_display_map[uid]
                        logger.debug(f"result marked complete, fetched cached result {ret}")
                        return ret
                
            # Clear any lingering calculations coming from this caller if the result isnt cached
            # downstream functions should pick up on these being deleted
            caller_dispatch = Dispatch(caller)

            # irrelevant
            # with self._caller_address_uid_map_lock:
            #     if caller_addr in self._caller_address_uid_map.keys():
            #         self.clear_uid(self._caller_address_uid_map[caller_addr])
            #     self._caller_address_uid_map[caller_addr] = uid

            # release the com args for use in another thread. convert them to streams
            # args = utils.com_args_release_to_stream_reserved(func, args)

            result_type = self._module_function_maps_wrapper.fname_type_register[fname]
            with self._uid_result_type_map_lock:
                self._uid_result_type_map[uid] = result_type

            # leave this in, but it wont be called
            caller_stream = _utils.comarshal_release_and_get_stream(caller_dispatch) # this marshal is the OG
            with self._uid_to_caller_map_lock:
                self._uid_to_caller_map[uid] = caller_stream

            f = self.create_worker_func_sync(uid, func, args, kwargs={})
            ret = f()

            do_mark_complete = False
            if isinstance(ret, Exception):
                e = ret
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
                    pass
                elif isinstance(e, pythoncom.com_error):
                    pass
                else:
                    do_mark_complete = True

            self._results_manager_thread._set_results_value(uid, ret)

            if any([type(ret) == x for x in (str, bool, int, float)]):
                ppret = self._handle_default_result(ret, uid)
                
            # an array or ret type will send the rets directly to excel via COM
            elif result_type == FunctionTypes.array_or_value:
                if type(ret) == xlproCollapsedType:
                    ppret = self._handle_pyobject_result(ret, uid)
                else:
                    ppret = self._handle_default_result(ret, uid)

            # sends a string to excel which effectively points to a stored result
            elif result_type == FunctionTypes.py_object:
                # a py_object list must be parsed before sending to excel to ensure the contents are compliant 
                if any([type(ret) == x for x in (list, tuple)]):
                    ppret = self._handle_iterable_result(ret, uid)
                elif type(ret) == xlproExpandedType:
                    ppret = self._handle_default_result(ret, uid)
                else:
                    ppret = self._handle_pyobject_result(ret, uid)
                    
            else:
                raise Exception("Result type invalid")
        
            ppret_fetched = self._client_manager_thread._get_result_display(uid)
            
            # marking the function complete is needed to handle the cached results.
            with self._uid_result_iscomplete_map_lock:
                self._uid_result_iscomplete_map[uid] = do_mark_complete

            return ppret_fetched

        # Return the python exception string as a fallback
        except Exception as e:
            return repr(errors.xlproUnhandledException(repr(e)))
        

    def create_worker_func_sync(self, uid, func, args, kwargs):
        raise NotImplementedError

        if kwargs:
            raise Exception("kwargs should not be here!")
        pass

        def modify_args_list(args:list):
            for i, arg in enumerate(args):
                if isinstance(arg, str):
                    if not arg.startswith("PyObj"):
                        continue
                    # handle the basic pyobject case
                    if m:=re.match(r"^PyObj<(.*)>$", arg):
                        with self._uid_results_map_lock:
                            temp_uid = m.group(1)
                            args[i] = self._uid_results_map[temp_uid]
                            if isinstance(args[i], xlproCollapsedType):
                                args[i] = args[i].data 

                    # handle the case for an address request for expanded values
                    elif m:=re.match(r"^PyObj<(.*)>_(\d+)$", arg):
                        with self._uid_results_map_lock:
                            temp_uid = m.group(1)
                            temp_addr = int(m.group(2))
                            # try and look it up, pass the error through to the function if we encounter one.
                            try:
                                args[i] = self._uid_results_map[temp_uid][temp_addr]
                                if isinstance(args[i], xlproCollapsedType):
                                    args[i] = args[i].data 
                            except IndexError as e:
                                args[i] = e
                            except Exception as e:
                                raise errors.xlproUnhandledException
            return args
            
        # check for py object request
        # args0 = deepcopy(args)
        import copy
        args_original = copy.deepcopy(args)
        args = list(args)
        # modify level 0 of the args
        modify_args_list(args)
        # check if level 1 needs to be modified
        try:
            for i, arg in enumerate(args):
                # a tuple nested arg may contain pyobject strings
                if isinstance(arg, (tuple, list)):
                    args[i] = modify_args_list(list(arg))
                    for i0, arg0 in enumerate(args[i]):
                        if isinstance(arg0, (list, tuple)):
                            args[i][i0] = modify_args_list(list(arg0))
            pass
        except Exception as e:
            raise e
                
        def worker():
            fname = _wrappers.get_registered_func_name(func, self._temp_module_name)
            f = _wrappers.generate_wrapped_function(self._temp_module_name, fname)
            try:
                ret = f(*args, **kwargs)
                return ret
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
                    pass
                elif isinstance(e, pythoncom.com_error):
                    pass
                elif isinstance(e, Exception):
                    pass

                return e
            
        return worker


    def create_worker_func(self, uid, func, args, kwargs):
        if kwargs:
            raise Exception("kwargs should not be here!")
        pass

        def replace_pyobjects_in_args(args:list):
            for i, arg in enumerate(args):
                if isinstance(arg, str):
                    if not arg.startswith("PyObj"):
                        continue

                    # handle the basic pyobject case
                    if m:=re.match(r"^PyObj<(.*)>$", arg):
                        with self._uid_results_map_lock:
                            temp_uid = m.group(1)
                            args[i] = self._uid_results_map[temp_uid]
                            if isinstance(args[i], (xlproCollapsedType, xlproExpandedType)):
                                args[i] = args[i].data 

                    # handle the case for an address request for expanded values
                    elif m:=re.match(r"^PyObj<(.*)>_(\d+)$", arg):
                        with self._uid_results_map_lock:
                            temp_uid = m.group(1)
                            temp_addr = int(m.group(2))
                            # try and look it up, pass the error through to the function if we encounter one.
                            try:
                                args[i] = self._uid_results_map[temp_uid][temp_addr]
                                if isinstance(args[i], (xlproCollapsedType, xlproExpandedType)):
                                    args[i] = args[i].data 
                            except IndexError as e:
                                args[i] = e
                            except Exception as e:
                                raise errors.xlproUnhandledException
                            pass
            return args
            
        # check for py object request
        # args0 = deepcopy(args)
        import copy
        # args_original = copy.deepcopy(args)
        args = list(args)
        # modify level 0 of the args
        replace_pyobjects_in_args(args)
        # check if level 1 needs to be modified
        try:
            for i, arg in enumerate(args):
                # a tuple nested arg may contain pyobject strings
                if isinstance(arg, (tuple, list)):
                    args[i] = replace_pyobjects_in_args(list(arg))
                    for i0, arg0 in enumerate(args[i]):
                        if isinstance(arg0, (list, tuple)):
                            args[i][i0] = replace_pyobjects_in_args(list(arg0))
            pass
        except Exception as e:
            raise e

        if func.__name__ == "add_legend":
            pass
                
        def worker():
            fname = _wrappers.get_registered_func_name(func, self._temp_module_name)
            if func.__name__ == "xlpro_getitem":
                pass
            if func.__name__ == "getitem":
                pass
            if func.__name__ == "add_line":
                pass
            if func.__name__ == "mpl_add_line_unique":
                pass
            if func.__name__ == "visualise_color":
                pass
            if func.__name__ == "show":
                pass
            if func.__name__ == "mpl_set_xlims":
                pass

            # SIGNIFICANT FIX RECOMMENDED
            # PyObj variables make no sense to go through the pre-validation steps
            # These functions should only be called on the arguments coming from Excel
            #   - preprocess_arguments
            #   - pre_validate_args
            # Sequence for validation for reference:
            #   - generate_wrapped_function
            #     - _pyobj_func_wrapper
            #       - pre_validate_args (validate no bad inputs exist)
            #       - preprocess_arguments (cast based on function signature)
            #       - pre_validate_args (validate no nested bad inputs exist on preprocessed, is this just to capture ptrs...?)
            #         - pre_p_an_arg (replaces pynone strings, evalulates pointer, converts argument to a target type, duplicates repeats prevalidate arg check lol)
            #     - _com_init_dispatch_release_wrapper
            # In summary, there is some cost to doing all the extra processing to validate the python arguments
            # As they should already be legitimate. But nothing that would break the calculation cycle
            # Split off the wrappers into individual steps time permitting.

            f = _wrappers.generate_wrapped_function(self._temp_module_name, fname)
            try:
                ret = f(*args, **kwargs)
                self._result_queue.put((uid, ret))
                logger.debug(f"Completed function '{uid}' successfully")
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
                        logger.debug(f"Error looks like a COM access error, modifying it from e={repr(e)}")
                        e = errors.xlproLikelyCOMAccessError(str(e))
                elif isinstance(e, pythoncom.com_error):
                    pass
                elif isinstance(e, Exception):
                    pass
                self._result_queue.put((uid, e))
                logger.warning(f"Completed function '{uid}' unsuccessfully. Error: '{e}'")
            logger.debug("Waking results manager from worker thread...")
            self._results_manager_thread.wake()
        
        return worker
    
    def create_worker_func_sub(self, uid, func, args, kwargs):
        def worker():
            try:
                ret = func(*args, **kwargs)
                self._result_queue.put((uid, ret))
                logger.debug(f"Completed subroutine '{uid}' successfully")
            except Exception as e:
                self._result_queue.put((uid, e))
                # logger.debug(f"Completed function '{uid}' unsuccessfully with error {e}")
                logger.warning(f"Completed subroutine '{uid}' unsuccessfully. Error: '{e}'")

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
        # funcs = self.get_active_registered_functions()
        return _utils.get_xlpro_vb_dynamic_component_contents(self.get_active_registered_functon_map())
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
    module_name = f"{CFG.xlpro_functions_stem}_{uid}"
    _utils.import_module(f"{CFG.xlpro_functions_stem}_{uid}", wd / f"{CFG.xlpro_functions_stem}.py")
    raise NotImplementedError
    func = getattr(sys.modules[module_name], func_name)

    fp = wd / "tmp" / f"{uid}.svg"
    fp.parent.mkdir(parents=True, exist_ok=True)

    fig:matplotlib.figure.Figure = func(*args, **kwargs)
    size_inches = np.array(fig.get_size_inches())
    fig.savefig(fp, dpi=600)

    queue.put((uid, (fp, size_inches)))
    pass


from concurrent.futures import ThreadPoolExecutor, Future

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
        return CFG.MAX_WORKERS

    def start(self):
        self._thread.start()
        self._thread
        logger.info("Worker Manager initialized")
        logger.debug(f"WorkerManager thread started: tid: {threading.get_native_id()}")

    def stop(self):
        self._stop_event.set()
        self._wake_event.set()  # Wake up if sleeping
        self._thread.join()

    def wake(self):
        self._wake_event.set()

    @staticmethod
    def sleep_wrapper(func):
        import random
        def wrapper(*args, **kwargs):
            time.sleep(random.random() * 3.0)
            return func(*args, **kwargs)
        return wrapper
    
    # def clear_futures(self):
    #     with self._futures_lock:
    #         self._futures = {}

    def _wrap_with_cleanup(self, uid, func):
        def wrapped():
            try:
                func()
            finally:
                try:
                    with self._threadpool_dict_lock:
                        thread = self._threadpool_dict.get(uid)
                        if thread is not None and not thread.is_alive():
                            del self._threadpool_dict[uid]
                            logger.debug(f"Cleaned up thread for uid: {uid}")
                except:
                    pass
        return wrapped

    def _process_function_queue(self):
        with self._threadpool_dict_lock:
            n_items = len(self._threadpool_dict.keys())
            
        if n_items < self.MAX_THREADS:
            uid = self._server._pending_function_queue.get(timeout=0.01)
            try:
                with self._server._uid_pending_function_map_lock:
                    # func = self.sleep_wrapper(self._wrap_with_cleanup(uid, self._server._uid_pending_function_map[uid]))
                    func = self._wrap_with_cleanup(uid, self._server._uid_pending_function_map[uid])
            except KeyError:
                logger.debug(f"Pending function queue uid not available, ignoring calculation request for uid '{uid}'")
                return
            
            with self._threadpool_dict_lock:
                if uid in self._threadpool_dict:
                    if not self._threadpool_dict[uid].is_alive():
                        self._threadpool_dict[uid].join()
                        del self._threadpool_dict[uid]
                    else:
                        logger.debug(f"Rejected to start worker for uid: '{uid}', already running")
                        return

                t = threading.Thread(target=func, daemon=True)
                # t = threading.Thread(target=self.sleep_wrapper(func), daemon=True)
                self._threadpool_dict[uid] = t
            # XXX - todo - limit the number of attempts for a given function in some way
            # XXX - todo - support sending terminate command to lingering worker threads
            t.start()
            return

        logger.debug(f"Reached maximum worker thread cap - MAX_THREADS: {self.MAX_THREADS}")


    def _clear_completed_threads(self):
        with self._threadpool_dict_lock:
            self._threadpool_dict = {uid: t for uid, t in self._threadpool_dict.items() if t.is_alive()}

    def _clear_all_threads(self):
        with self._threadpool_dict_lock:
            for uid, t in self._threadpool_dict.items():
                t.join()

    def _run(self):
        while not self._stop_event.is_set():
            # logger.debug("WorkerManager thread is waiting for events or timeout...")
            self._wake_event.wait(timeout=0.1)
            if self._wake_event.is_set():
                self._wake_event.clear()
                logger.debug("WorkerManager thread woke up for an event!")
            while True:
                try:
                    self._process_function_queue()
                except queue.Empty:
                    break
                self._clear_completed_threads()
                time.sleep(SLEEP_DURATION) # fairness sleep

            if self._stop_event.is_set():
                break 


class ResultsManager:
    """Manages the results queue and signals update requests to the client manager"""
    def __init__(self, server:xlproWorkspace):
        self._stop_event = threading.Event()
        self._wake_event = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._server = server
        logger.info("Results Manager initialized")

    def start(self):
        self._thread.start()
        self._thread
        logger.debug(f"ResultsManager thread started: tid: {threading.get_native_id()}")

    def stop(self):
        self._stop_event.set()
        self._wake_event.set()  # Wake up if sleeping
        self._thread.join()

    def wake(self):
        self._wake_event.set()

    def _set_results_value(self, uid, val):
        with self._server._uid_results_map_lock:
            self._server._uid_results_map[uid] = val

    def _set_subresults_values(self, uid, val):
        with self._server._uid_results_map_lock:
            for i, subval in enumerate(val):
                self._server._uid_results_map[f"{uid}_{i}"] = subval


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

            elif isinstance(val, TypeError):
                # There is a bug of some kind where .Address on a range raises a Type error. This case deals with it
                if val.args[0] == 'This object does not support enumeration':
                    logger.debug(f"Assuming OBJREF invalid '{uid}', recycling function...")
                    # self.signal_worker_manager_to_recalculate(uid)
                    self._set_results_value(uid, ret)
                    self._server._client_recalculate_queue.put(uid)
                    logger.debug(f"ResultsManager is waking the ClientManager after successful calculation of '{uid}'")
                    self._server._client_manager_thread.wake()
                    return

            elif isinstance(val, pythoncom.com_error):
                # recycle if excel is not accessible
                if VBErrorConverter(val) == VBError.xlCallRejectedByCallee:
                    logger.debug(f"Call rejected by callee for '{uid}', recycling function...")
                    self.signal_worker_manager_to_recalculate(uid)
                    return
                elif VBErrorConverter(val) == 285: # The marshaled interface data packet (OBJREF) has an invalid or unknown format.
                    logger.debug(f"OBJREF invalid '{uid}', recycling function...")
                    # self.signal_worker_manager_to_recalculate(uid)
                    self._set_results_value(uid, ret)
                    self._server._client_recalculate_queue.put(uid)
                    logger.debug(f"ResultsManager is waking the ClientManager after successful calculation of '{uid}'")
                    self._server._client_manager_thread.wake()
                    return
                elif VBErrorConverter(val) == 30: # A disk error occurred during a read operation.
                    logger.debug(f"disk read failed '{uid}', recycling function...")
                    # self.signal_worker_manager_to_recalculate(uid)

                    # with self._server._uid_pending_function_map_lock:
                    #     logger.debug(f"calculation removed as pending function {uid}")
                    #     del self._server._uid_pending_function_map[uid]

                    # tell the client manager to issue a recalculate
                    self._set_results_value(uid, ret)
                    self._server._client_recalculate_queue.put(uid)
                    logger.debug(f"ResultsManager is waking the ClientManager after successful calculation of '{uid}'")
                    self._server._client_manager_thread.wake()
                    return
                else:
                    logger.debug(f"Other COM error for '{uid}', recycling function...")
                    logger.debug(f"'{repr(val)}'")
                    self.signal_worker_manager_to_recalculate(uid)
                    return


            elif isinstance(val, errors.xlproLikelyCOMAccessError):
                logger.debug(f"Likely COM access error for {uid}, recycling.")
                self.signal_worker_manager_to_recalculate(uid)
                return
            
            elif isinstance(val, UnboundLocalError):
                logger.debug(f"UnboundLocalError {uid}, recycling... (TODO fix this)")
                self.signal_worker_manager_to_recalculate(uid)
                pass

            logger.debug(f"Returned value is a generic exception: {uid}, {val}")
            # ret = repr(val) # convert exception to string for it to show in excel.


        # if it is a valid return, write the result to the cache
        self._set_results_value(uid, ret)
        # xxx - todo - only should expand the subresults if requested instead of every time an iterable is returned.
        # if isinstance(val, typing.Iterable):
        # if isinstance(val, list):
        #     self._set_subresults_values(uid, val)

        # signal that it is complete
        with self._server._uid_result_iscomplete_map_lock:
            logger.debug(f"calculation marked complete {uid}")
            if uid.startswith("$H$25"):
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
                logger.debug("ResultsManager thread woke up for an event!")
            # Process the whole queue once woken up
            # XXX - could replace while trye with while not stop event.
            while True:
                try:
                    self._process_queue_element()
                except queue.Empty:
                    break
                time.sleep(SLEEP_DURATION) # fairness sleep

            if self._stop_event.is_set():
                break 

import pandas as pd
import datetime
import decimal

SIMPLE_DISPLAY_TYPES_DISPLAY_CONVERSION = {
    # Native Python types
    str: lambda x: x,
    bool: lambda x: x,
    int: lambda x: x,
    float: lambda x: x,
    type(None): lambda x: _utils.XLPRO_NONE_STR,

    # NumPy numeric types
    np.bool_: lambda x: bool(x),
    np.int8: lambda x: int(x),
    np.int16: lambda x: int(x),
    np.int32: lambda x: int(x),
    np.int64: lambda x: int(x),
    np.uint8: lambda x: int(x),
    np.uint16: lambda x: int(x),
    np.uint32: lambda x: int(x),
    np.uint64: lambda x: int(x),
    np.float16: lambda x: float(x),
    np.float32: lambda x: float(x),
    np.float64: lambda x: float(x),

    np.datetime64: lambda x: _utils.datetime_to_excel(x),

    # Complex numbers
    # np.complex64: lambda x: complex(x),
    # np.complex128: lambda x: complex(x),
    # np.complexfloating: lambda x: complex(x),

    # Pandas scalar types
    pd.Int8Dtype(): lambda x: int(x),
    pd.Int16Dtype(): lambda x: int(x),
    pd.Int32Dtype(): lambda x: int(x),
    pd.Int64Dtype(): lambda x: int(x),
    pd.UInt8Dtype(): lambda x: int(x),
    pd.UInt16Dtype(): lambda x: int(x),
    pd.UInt32Dtype(): lambda x: int(x),
    pd.UInt64Dtype(): lambda x: int(x),
    pd.BooleanDtype(): lambda x: bool(x),
    pd.StringDtype(): lambda x: str(x),

    # Dates / Times
    # np.datetime64: lambda x: pd.to_datetime(x),
    # datetime.datetime: lambda x: x,
    # datetime.date: lambda x: x,
    # pd.Timestamp: lambda x: x,
    # np.timedelta64: lambda x: pd.to_timedelta(x),
    # datetime.timedelta: lambda x: x,
    # pd.Timedelta: lambda x: x,

    # Decimal
    # decimal.Decimal: lambda x: float(x),
    
}

# handle 128 bit
if hasattr(np, 'float128'):
    SIMPLE_DISPLAY_TYPES_DISPLAY_CONVERSION[np.float128] = lambda x: float(x)

SIMPLE_DISPLAY_TYPES_DISPLAY_CONVERSION_KEYS = SIMPLE_DISPLAY_TYPES_DISPLAY_CONVERSION.keys()

class ClientManager:
    """Hooks to the excel client so we can trigger events"""
    def __init__(self, server:xlproWorkspace):
        self._stop_event = threading.Event()
        self._wake_event = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._server = server
        logger.info("Client Manager initialized")

    def start(self):
        self._thread.start()
        self._thread
        logger.debug(f"ClientManager thread started: tid: {threading.get_native_id()}")

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
        with self._server._uid_result_display_map_lock:
            return self._server._uid_result_display_map[uid]
        
    def _get_value(self, uid):
        with self._server._uid_results_map_lock:
            return self._server._uid_results_map[uid]
        
    def _update_client_iterable_result(self, uid) -> None:
        """Update the data for the case where the result is a list"""
        try:
            caller_dispatch = _utils.comarshal_dispatch_stream(self._server.get_caller_stream(uid))
            iterable_val = self._get_value(uid)

            if isinstance(iterable_val, Exception):
                logger.debug(f"Value is an exception: '{iterable_val}', '{uid}'")
                self._set_result_display(uid, repr(iterable_val))

            # construct a list off the iterable value
            elif isinstance(iterable_val, typing.Iterable):
                val_modified = []
                for i, subval in enumerate(iterable_val):
                    if isinstance(subval, (int, float, str, bool)):
                        val_modified.append(subval)
                    else:
                        val_modified.append(f"PyObj<{uid}>_{i}")

                # create a *_expanded variant of the uid result display
                self._set_result_display(f"{uid}_expanded", np.array(val_modified, dtype=object).reshape(-1,1))

                pass

            # update by resetting the formula
            _utils.comsafe(lambda: caller_dispatch.Application.Run("'xlpro.xlam'!AtomicFormulaRefreshNoEvents", caller_dispatch))()
        
            # caller_dispatch.Formula2 = caller_dispatch.Formula2
        except KeyError as e:
            # XXX - todo - there is a risk of a keyerror here for some reason
            logger.debug(f"Error during client update: '{uid}', {e}")
        finally:
            try:
                self._server.set_caller_stream(uid, _utils.comarshal_release_and_get_stream(caller_dispatch))
            except:
                pass
    
    def _update_client_pyobject_result(self, uid) -> None:
        """Update the data for the py_object case (row-major arrays, strings, values)"""
        if uid.startswith("$H$25"):
            pass
        try:
            caller_dispatch = _utils.comarshal_dispatch_stream(self._server.get_caller_stream(uid))
            val0 = self._get_value(uid)
            if isinstance(val0, xlproCollapsedType):
                val = val0.data
            else:
                val = val0
            if isinstance(val, Exception):
                # TODO investigate need to show these as user warnings
                logger.debug(f"Value is an exception: '{val}', '{uid}'")
                self._set_result_display(uid, repr(val))
            else:
                if type(val) in SIMPLE_DISPLAY_TYPES_DISPLAY_CONVERSION_KEYS:
                    self._set_result_display(uid, SIMPLE_DISPLAY_TYPES_DISPLAY_CONVERSION[type(val0)](val))
                else:
                    self._set_result_display(uid, f"PyObj<{uid}>")
                    # self._set_result_display(uid, val)

            # update by resetting the formula
            _utils.comsafe(lambda: caller_dispatch.Application.Run("'xlpro.xlam'!AtomicFormulaRefreshNoEvents", caller_dispatch))()
            # caller_dispatch.Application.Run("'xlpro.xlam'!AtomicFormulaRefreshNoEvents", caller_dispatch)
            # caller_dispatch.Formula2 = caller_dispatch.Formula2
            # self._server.set_caller_stream(uid, _utils.comarshal_release_and_get_stream(caller_dispatch))
        except KeyError as e:
            # XXX - todo - there is a risk of a keyerror here for some reason
            logger.debug(f"Error during client update: '{uid}', {e}")
        finally:
            try:
                self._server.set_caller_stream(uid, _utils.comarshal_release_and_get_stream(caller_dispatch))
            except:
                pass
        
    def _update_client_default_result(self, uid) -> None:
        """Update the data for the default case (row-major arrays, strings, values)"""
        try:
            caller_dispatch = _utils.comarshal_dispatch_stream(self._server.get_caller_stream(uid))
            # handle the case where the return may be an array expand request or a basic value
            val0 = self._get_value(uid)
            if type(val0) == xlproExpandedType:
                val = val0.data
                # cast xlpro expanded types to the 2d data types
                if type(val) == list:
                    val = ExcelArrayConverter(val, list2d)
                elif type(val) == np.ndarray:
                    val = ExcelArrayConverter(val, ndarray2d)
            else:
                val = val0

            if isinstance(val, Exception):
                logger.debug(f"Value is an exception: '{val}', '{uid}'")
                self._set_result_display(uid, repr(val))
            else:
                self._set_result_display(uid, val)

            # update by resetting the formula
            _utils.comsafe(lambda: caller_dispatch.Application.Run("'xlpro.xlam'!AtomicFormulaRefreshNoEvents", caller_dispatch))()
            # caller_dispatch.Application.Run("'xlpro.xlam'!AtomicFormulaRefreshNoEvents", caller_dispatch)
            # caller_dispatch.Formula2 = caller_dispatch.Formula2
            self._server.set_caller_stream(uid, _utils.comarshal_release_and_get_stream(caller_dispatch))
        except KeyError as e:
            # XXX - todo - there is a risk of a keyerror here for some reason
            logger.debug(f"Error during client update: '{uid}', {e}")
        
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

            def _xlinteract():
                caller_adjacent = caller_dispatch.Cells(2,1)
                xpos, ypos = caller_adjacent.Left, caller_adjacent.Top
                width, height = size_pt

                ws = caller_adjacent.Parent
                # ws.Shapes.AddPicture(str(fp.resolve()), False, True, xpos, ypos, width, height)

                logger.info(f"Adding image to Excel: '{xl_name}'")
                # look for an existing shape with the same name
                ws:'xl._Worksheet'
                try:
                    shape = ws.Shapes(xl_name)
                    xpos, ypos = shape.Left, shape.Top
                    shape.Delete()
                except pythoncom.com_error as e:
                    logger.debug(f"Could not find object with name: {xl_name} to delete")

                try:
                    shape = ws.Shapes.AddPicture(str(fp.resolve()), False, True, xpos, ypos, width, height)
                    time.sleep(0.10)
                    shape.Name = xl_name
                except Exception as e:
                    shape.Delete()
                    raise e
                logger.info(f"Image '{xl_name}' added successfully")
                self._set_result_display(uid, f"Image<{xl_name}>")
                caller_dispatch.Application.Run("'xlpro.xlam'!AtomicFormulaRefreshNoEvents", caller_dispatch)
                return
            
            _utils.comsafe(_xlinteract)()
            
            # self._set_result_display(uid, f"Image<{xl_name}>")
            # _utils.comsafe(lambda: caller_dispatch.Application.Run("'xlpro.xlam'!AtomicFormulaRefreshNoEvents", caller_dispatch))()

            # self._set_result_display(uid, f"Image<{fp}>")

            # update by resetting the formula
            # caller_dispatch.Application.Run("'xlpro.xlam'!AtomicFormulaRefreshNoEvents", caller_dispatch)
            # caller_dispatch.Formula2 = caller_dispatch.Formula2
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
                if uid.lower().startswith("$e$27"):
                    pass
                # logger.debug(f"Client manager fetched uid '{uid}'")
            except queue.Empty:
                # logger.debug("Client manager queue is empty")
                return
            
            # with self._server._uid_result_iscomplete_map_lock:
            #     if not uid in self._server._uid_result_iscomplete_map.keys():
            #         logger.debug(f"Uid '{uid}' not found as pending function within client manager update. Ignoring this iteration")
            #         return

            pass

            # fetch the result type so we know how to handle it
            with self._server._uid_result_type_map_lock:
                result_type = self._server._uid_result_type_map[uid]

            # if we fail to dispatch the range, it was probably deleted.
            # therefore we don't need to replace it in the queue
            replace_in_queue = True

            # Decide whether to recycle
            try:
                # xxx - todo - hack to skip checks on subroutines
                # subroutines will never need to be recycled by the client manager
                # subroutines will be recycled by the results manager since there will
                # never be any result to reach the client.
                # so we are ok to just skip the loop here, happy days.

                caller_dispatch = self._server.get_caller_stream(uid)
                if caller_dispatch is SUB_CALLER_FLAG_STRING:
                    logger.debug("caller_dispatched checked as None, assuming this is a subroutine and skipping further.")
                    time.sleep(SLEEP_DURATION) # copy the fairness sleep
                    continue
                
                value = self._get_value(uid)

                if isinstance(value, TypeError):
                    pass

                # handle images every time
                if type(value) == xlproImage:
                    self._update_client_image_result(uid)

                if any([type(value) == x for x in (str, bool, int, float)]):
                    self._update_client_default_result(uid)
                    
                # an array or value type will send the values directly to excel via COM
                elif result_type == FunctionTypes.array_or_value:
                    if type(value) == xlproCollapsedType:
                        self._update_client_pyobject_result(uid)
                    else:
                        self._update_client_default_result(uid)

                # sends a string to excel which effectively points to a stored result
                elif result_type == FunctionTypes.py_object:
                    # a py_object list must be parsed before sending to excel to ensure the contents are compliant 
                    if any([type(value) == x for x in (list, tuple)]):
                        self._update_client_iterable_result(uid)
                    elif type(value) == xlproExpandedType:
                        self._update_client_default_result(uid)
                    else:
                        self._update_client_pyobject_result(uid)

                else:
                    msg = "Internal Server Error: Result type not valid"
                    logger.critical(msg)
                    raise Exception(msg)
                
                
                # if successful we don't need to replace#
                logger.info(f"Client manager successfully processed uid '{uid}'.")
                replace_in_queue = False
                # XXX - todo - consider removing the uid after this call
            except pywintypes.com_error as e:
                if VBErrorConverter(e) == VBError.xlObjectRequired:
                    # Range has been deleted
                    replace_in_queue = False
                    logger.warning("VB Error - Object Required. Cell has been deleted. Removing from queue.")
                    # XXX - todo - if this is dropped from the queue the uid definitely needs to be marked for delete.
                elif VBErrorConverter(e) == VBError.xlCallRejectedByCallee:
                    # Call rejected - excel might be in a dialogue 
                    logger.warning("VB Error - Call rejected, recycling in queue")
                else:
                    logger.warning(f"Other COM Error occurred, {e}, recycling in queue")

                logger.info(f"Could not recalculate caller_dispatch for uid: '{uid}'")

            except AttributeError as e:
                logger.warning("AttributeError during cell update, Application may be in dialogue")
            except KeyError as e:
                logger.warning(f"Error during client queue processing 1! {e}")
            except Exception as e:
                logger.warning(f"Error during client queue processing 2! {e}")
            finally:
                # XXX - Marshalling the caller back to the pool in case
                logger.debug("Releasing caller dispatch during ClientManager._process_queue()")
                try:
                    self._server.set_caller_stream(uid, _utils.comarshal_release_and_get_stream(caller_dispatch))
                except Exception as e:
                    pass

            # if we failed to update, recycle the queue as necessary
            if not replace_in_queue:
                pass
            if replace_in_queue:
                self._server._client_recalculate_queue.put(uid)

            time.sleep(SLEEP_DURATION) # fairness sleep


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