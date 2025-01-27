import win32com.client


from pywintypes import IID
from win32com.client import Dispatch

from win32typelibs import excel as xl
from win32typelibs import vbide

from typing import Any, Callable
import typing

import inspect
import logging

from pathlib import Path
import hashlib
import uuid
from functools import wraps
import xlpro_typing
import ctypes
import os

import textwrap
import numpy as np
import matplotlib.figure
import pythoncom
import importlib
import types
import sys
import regex as re

from xlpro_types import xlproptr
import errors
import json



logger = logging.getLogger(__name__)

VB_DYNAMIC_MODULE_NAME = "xlpro_async"

def get_function_signature(func):
    # Get the type hints from the function
    type_hints = typing.get_type_hints(func)
    
    # Get the parameter information using inspect
    signature = inspect.signature(func)
    parameters = signature.parameters
    
    # Build the output list
    result = []
    for param_name, param in parameters.items():
        # Get the type hint for the parameter or default to Any
        param_type = type_hints.get(param_name, Any)
        result.append((param_name, param_type))

    default_value_map = {
        param.name: param.default
        for param in signature.parameters.values()
        if param.default is not inspect.Parameter.empty
    }
    
    # return func.__qualname__, result, type_hints.get('return', Any), default_value_map
    return func.__name__, result, type_hints.get('return', Any), default_value_map

# Converts python type to vb type
vb_type_conversion_strings = {
    int: "Cint({})",
    float: "Cdbl({})",
    bool: "Cbool({})",
    str: "{}",
    Any: "{}"
}

# Use in function definitions
vb_type_declaration_strings = {
    int: "{} As Integer",
    float: "{} As Double",
    bool: "{} As Boolean",
    str: "{} As String",
    Any: "{} As Variant",
}

vb_range_conversion_check_string = """If TypeName({arg}) = \"Range\" Then
    {arg} = {arg}.Value
EndIf"""


def infer_func_result_type_from_type_hints(func) -> int:
    # XXX - todo - link this up with the enum in the server at some point
    f_name, args_and_types, ret_type, default_value_map = get_function_signature(func)
    if ret_type == matplotlib.figure.Figure:
        return 1
    return 0

def infer_function_type(func):
    f_name, args_and_types, ret_type, default_value_map = get_function_signature(func)
    if ret_type == matplotlib.figure.Figure:
        return 1
    return 0


def function_template_with_caller(func:Callable) -> str:
    """Returns function template string to send to VBA module.
    If the reserved `caller` argument is used, pass it to the execute function call.
    """
    func_name, args_and_types, ret_type, default_value_map = get_function_signature(func)
    docstring = func.__doc__
    arg_declaration_list = []
    arg_conversion_list = []
    arg_range_conversion_check_list = []

    for a, t in args_and_types:
        if vb_type_declaration_strings.get(t, None):
            arg_declaration_list.append( vb_type_declaration_strings[t].format(a))
        else:
            arg_declaration_list.append("{} As Variant".format(a))
        if vb_type_conversion_strings.get(t, None):
            arg_conversion_list.append(vb_type_conversion_strings[t].format(a))
        else:
            arg_conversion_list.append("{}".format(a))
        if a not in ["caller", "thiswb"]:
            if t not in [float, int, bool, str]:
                arg_range_conversion_check_list.append(
                    textwrap.indent(vb_range_conversion_check_string.format(arg=a), "    ")
            )

    # XXX - todo - ensure no reserved vba arguments are parsed!   
    
    a_list = [a for a, t in args_and_types]

    # # handle the reserved caller keyword
    # arg_idxs_to_del = []
    # if "caller" in a_list:
    #     indx = a_list.index("caller")
    #     arg_idxs_to_del.append(indx)
    #     arg_conversion_list[indx] = "Application.Caller"
    
    # # handle the reserved ActiveWorkbook keyword (pseudo thisworkbook...)
    # if "thiswb" in a_list:
    #     indx = a_list.index("thiswb")
    #     arg_idxs_to_del.append(indx)
    #     arg_conversion_list[indx] = "ActiveWorkbook"

    # if arg_idxs_to_del:
    #     arg_idxs_to_del.sort(reverse=True)
    #     for idx in arg_idxs_to_del:
    #         del arg_declaration_list[idx]

    import server 

    return f"""Function {func_name}({', '.join(arg_declaration_list)}) as Variant
    Dim xlpro As Object
    Set xlpro = GetObject("new: {server.xlproServer._reg_clsid_}")
{'\n'.join(arg_range_conversion_check_list)}
    {func_name} = xlpro.{server.xlproServer.execute_function_async.__name__}(ActiveWorkbook, Application.Caller, "{func_name}", {', '.join(arg_conversion_list)})
End Function
"""

def get_or_create_codemodule(wb:xl._Workbook, c_name:str) -> vbide._CodeModule:
    proj:vbide._VBProject = wb.VBProject

    if not c_name in [x.Name for x in proj.VBComponents]:
        comp = proj.VBComponents.Add(vbide.constants.vbext_ct_StdModule)
        comp.Name = c_name
    else:
        comp = proj.VBComponents(c_name)

    codemod:vbide._CodeModule = comp.CodeModule
    return codemod


def write_to_vb_module(s:str, vb_codemod:vbide._CodeModule):
    vb_codemod.DeleteLines(1, vb_codemod.CountOfLines)
    vb_codemod.AddFromString(s)
    pass

def init_xlpro_vb_dynamic_component(wb:xl._Workbook, func_register:list[Callable]):
    """Write a list of commands to be registered in vba."""
    vb_dynamic_comdemod = get_or_create_codemodule(wb, VB_DYNAMIC_MODULE_NAME)
    s_list = []
    for f in func_register:
        if not isinstance(f, Callable):
            raise TypeError(f"Item must be a function, {type(f)}, {f}")
        s_list.append(function_template_with_caller(f))

    write_to_vb_module("\n".join(s_list), vb_dynamic_comdemod)
    # write_to_vb_module("\n".join([s_list[0]]), vb_dynamic_comdemod)
    # write_to_vb_module("Sub fn()\n msgbox \"hello\"\nend sub", vb_dynamic_comdemod)
    # vb_dynamic_comdemod = None
    # pythoncom.CoUninitialize()

def get_xlpro_vb_dynamic_component_contents(func_register:list[Callable]) -> str:
    s_list = []
    for f in func_register:
        if not isinstance(f, Callable):
            raise TypeError(f"Item must be a function, {type(f)}, {f}")
        s_list.append(function_template_with_caller(f))
    return "\n".join(s_list)



# XXX - todo - get a better understsanding of these COM names, they can't be right lol
def comarshal_release_and_get_stream(com_dispatch):
    """Releases the COM object (PyIDispatch) from this thread and returns the stream 
    (PyIStream)"""
    stream = pythoncom.CoMarshalInterThreadInterfaceInStream(
        pythoncom.IID_IDispatch,
        com_dispatch,
    )
    com_dispatch = None
    return stream

def comarshal_dispatch_stream(com_stream):
    """Dispatch a com stream (PyIStream) to a com object"""
    com_obj_pyidispatch = pythoncom.CoGetInterfaceAndReleaseStream(
        com_stream, pythoncom.IID_IDispatch,
    )
    com_obj_dispatch = win32com.client.Dispatch(com_obj_pyidispatch)
    return com_obj_dispatch


def com_args_release_to_stream_reserved(func, args):
    """Be careful which thread this runs on!
    Marshals the caller and thiswb reserved keyword arguments for use
    in another thread. Replaces the args with streams that can be used on another thread.
    """
    f_name, args_and_types, ret_type, _ = get_function_signature(func)
    arg_names = [v0 for v0, v1 in args_and_types]
    new_args = list(args) #  args come in immutable (tuples)
    if "caller" in arg_names:
        idx = arg_names.index("caller")
        caller = args[idx]
        caller_stream = comarshal_release_and_get_stream(caller)
        new_args[idx] = caller_stream
        # Caller type could be many things, likely just a Range.
        raise NotImplementedError("Support for caller and thiswb dropped until deferred calculation flow reworked")
        pass
    if "thiswb" in arg_names:
        idx = arg_names.index("thiswb")
        thiswb = args[idx]
        thiswb_stream = comarshal_release_and_get_stream(thiswb)
        new_args[idx] = thiswb_stream
        raise NotImplementedError("Support for caller and thiswb dropped until deferred calculation flow reworked")
        pass
    return new_args

def com_args_dispatch_reserved(func, args):
    """Be careful which thread this runs on! 
    Marshals the caller and thiswb reserved keyword arguments for use
    in another thread. Replaces the args with streams that can be used on another thread.
    """
    f_name, args_and_types, ret_type, _ = get_function_signature(func)
    arg_names = [v0 for v0, v1 in args_and_types]
    new_args = list(args)
    if "caller" in arg_names:
        idx = arg_names.index("caller")
        caller = args[idx]
        caller_stream = comarshal_dispatch_stream(caller)
        new_args[idx] = caller_stream
        # Caller type could be many things, likely just a Range.
        raise NotImplementedError("Support for caller and thiswb dropped until deferred calculation flow reworked")
        pass
    if "thiswb" in arg_names:
        idx = arg_names.index("thiswb")
        thiswb = args[idx]
        thiswb_stream = comarshal_dispatch_stream(thiswb)
        new_args[idx] = thiswb_stream
        raise NotImplementedError("Support for caller and thiswb dropped until deferred calculation flow reworked")
        pass
    return new_args

def get_args_minus_reserved(func, args):
    """Returns the arguments of a function but removes the reserved keywords
    as to prevent the pyidispatch strings that are generated by memory allocation
    from contaminating the string."""
    f_name, args_and_types, ret_type, _ = get_function_signature(func)
    arg_names = [v0 for v0, v1 in args_and_types]
    arg_idxs_to_del = []
    new_args = list(args)
    if "caller" in arg_names:
        idx = arg_names.index("caller")
        arg_idxs_to_del.append(idx)
        raise NotImplementedError("Support for caller and thiswb dropped until deferred calculation flow reworked")
    if "thiswb" in arg_names:
        idx = arg_names.index("thiswb")
        arg_idxs_to_del.append(idx)
        raise NotImplementedError("Support for caller and thiswb dropped until deferred calculation flow reworked")
    arg_idxs_to_del.sort(reverse=True)
    for idx in arg_idxs_to_del:
        new_args.pop(idx)
    return new_args


def import_module(module_name, file_path):
    """Dynamically import a module with a custom name"""
    # Load the module dynamically
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    sys.modules[module_name] = module

def get_udf_valid_functions_from_module(module_name):
    """Retrieves all functions from a model"""
    # Get all functions in the module
    module = sys.modules[module_name]
    functions = [
        v
        for name in dir(module)
        if isinstance((v:=getattr(module, name)), types.FunctionType)
    ]
    return functions


def get_udf_valid_function_names_from_module(module_name):
    """Returns a list of function names within a module. Returned names satisfy
    being a valid callable function from Excel
    """
    # Get all functions in the module
    module = sys.modules[module_name]
    functions = [
        name
        for name in dir(module)
        if isinstance((v:=getattr(module, name)), types.FunctionType)
    ]
    return functions


def is_function_udf_valid(func):
    return isinstance(func, types.FunctionType)

def get_udf_valid_functions(funcs:list):
    return [f for f in funcs if is_function_udf_valid(f)]


def hash_function_call(func, *args, **kwargs):
    # Create a unique string based on the function name and its arguments
    func_name = func.__name__
    # Convert arguments to a string (including both positional and keyword arguments)
    args_str = str(args)
    kwargs_str = str(kwargs)

    # Combine the function name with its arguments
    combined_string = func_name + args_str + kwargs_str

    # Generate a hash using SHA-256 (you can also use MD5 or others depending on your needs)
    return hash_str(combined_string)

def hash_str(s:str):
    return hashlib.sha256(s.encode('utf-8')).hexdigest()


def convert_xl_2d_types_args(func, args):
    _, args_and_types, _, _ = get_function_signature(func)
    ppargs = []
    for val, (a, t) in zip(args, args_and_types):
        ppargs.append(xlpro_typing.ExcelArrayConverter(val, t))
    return ppargs

def convert_xl_2d_types_kwargs(func, kwargs):
    _, args_and_types, _, _ = get_function_signature(func)
    ppkwargs = {}
    for k, v in kwargs.items():
        for a, t in args_and_types:
            if a == k:
                ppkwargs[a] = xlpro_typing.ExcelArrayConverter(v, t)
                break

    return ppkwargs


def com_init_dispatch_release_wrapper(func):
    """Wraps com object dispatch and release around a func.
    Also appropriately configures pythoncom coinitialise"""
    raise NotImplementedError
    @wraps(func)
    def wrapper(*args, **kwargs):
        pythoncom.CoInitialize()
        # dispatch the args on this thread
        # only relevant if the reserved dispatch arguments are being used.
        args_dispatched = com_args_dispatch_reserved(func, args)

        ret = func(*args_dispatched, **kwargs)

        # must release after!
        com_args_release_to_stream_reserved(func, args_dispatched)
        pythoncom.CoUninitialize()
        
        return ret
    
    return wrapper


def show_warning(title, message):
    # MessageBox parameters: hWnd, text, caption, uType
    ctypes.windll.user32.MessageBoxW(0, message, title, 0x30)  # 0x30 = MB_ICONWARNING

def get_short_path(long_path):
    # Ensure the path exists
    if not os.path.exists(long_path):
        raise FileNotFoundError(f"The path '{long_path}' does not exist.")
    
    # Allocate a buffer for the short path
    buffer = ctypes.create_unicode_buffer(260)  # Maximum path length on Windows
    ctypes.windll.kernel32.GetShortPathNameW(long_path, buffer, len(buffer))
    
    return buffer.value



def hash_cell(rng_dispatch) -> str:
    rng = Dispatch(rng_dispatch)
    # XXX - todo - check that this also works for only worksheet files.
    ws = rng.Parent
    wb = ws.Parent
    comarshal_release_and_get_stream(rng)
    # return f"{get_short_path(wb.FullName)}::{ws.Name}::{rng.Address}"
    return f"{wb.FullName}::{ws.Name}::{rng.Address}"


from xlpro_typing import list1d, list2d
def jsonify(arr):
    """Converts range to json string"""
    from xlpro_typing import ExcelArrayConverter
    arr:list2d = ExcelArrayConverter(arr, list2d)

    if len(arr[0]) != 2:
        raise Exception("Please provide a nx2 array of key:value pairs")
    ret = {}
    for row in arr:
        if not isinstance(row[0], str):
            raise TypeError("Ensure the first column values are all strings")
        ret[row[0]] = row[1]

    return json.dumps(ret, indent=2)

def is_arg_promise(arg):
    if not isinstance(arg, str):
        return False
    return re.search(r"^Promise<\w*>", arg)

def validate_args_ready(args, kwargs):
    for arg in list(args) + [v for k, v in kwargs.items()]:
        if is_arg_promise(arg):
            raise errors.ArugmentNotReadyException

# import copy
def pre_p_an_arg(cval, target_type):
    # 0. raise error if the argument is currently a promise!
    if is_arg_promise(cval):
        raise errors.ArugmentNotReadyException
    
    # 1. check if its an xlproptr. Replace val with the ptr result
    # cval = copy.copy(val)
    if xlproptr.is_ptr(cval):
        cval = xlproptr.decode(cval).evaluate()

    
    # 2. convert an argument to a target type
    ppval = xlpro_typing.ExcelArrayConverter(cval, target_type)
    return ppval

def preprocess_arguments(func, args:typing.Iterable=None, kwargs:dict=None):
    if not args is None and not isinstance(args, typing.Iterable):
        raise TypeError("args must be an iterable")
    if not kwargs is None and not isinstance(kwargs, dict):
        raise TypeError("kwargs must be a dict")

    _, args_and_types, _, _ = get_function_signature(func)

    ppargs = []
    if args is not None:
        for val, (a, t) in zip(args, args_and_types):
            ppargs.append(pre_p_an_arg(val, t))

    ppkwargs = {}
    if kwargs is not None:
        for k, val in kwargs.items():
            for a, t in args_and_types:
                if a == k:
                    ppkwargs[a] = (pre_p_an_arg(val, t))
                    break

    return ppargs, ppkwargs
    

import threading
class ThreadWithException(threading.Thread):
    """Thread wrapper class that allows exception extraction"""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.exception = None

    def run(self):
        try:
            if self._target:
                self._target(*self._args, **self._kwargs)
        except Exception as e:
            self.exception = e  # Store the exception

    def get_exception(self):
        return self.exception


def get_precedents_chain(rng_stream):
    rng_dispatch:xl.Range = comarshal_dispatch_stream(rng_stream)
    precedents = []
    for cell in rng_dispatch.Precedents:
        precedents.append(comarshal_release_and_get_stream(cell))
    comarshal_release_and_get_stream(rng_dispatch)
    return precedents


if __name__ == "__main__":

    # jsonify_func(hash_str)

    # a = xlproptr.decode("*<a::b::c>")


    # Example target function
    def faulty_function():
        raise ValueError("Something went wrong in the thread!")

    # Using the custom thread
    thread = ThreadWithException(target=faulty_function)
    thread.start()
    thread.join()

    # Check for exceptions
    if thread.get_exception():
        print(f"Exception occurred: {thread.get_exception()}")
    else:
        print("Thread completed successfully.")




    pass

