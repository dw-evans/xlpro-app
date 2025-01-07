import win32com.client


from pywintypes import IID
from win32com.client import Dispatch

from win32typelibs import excel as xl
from win32typelibs import vbide

from typing import Any, Callable
import typing

import inspect
import logging

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
    
    return func.__qualname__, result, type_hints.get('return', Any), default_value_map

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

import textwrap
vb_range_conversion_check_string = """If TypeName({arg}) = \"Range\" Then
    {arg} = {arg}.Value
EndIf"""


import numpy as np

import matplotlib.figure
def get_func_result_type(func) -> int:
    # XXX - todo - link this up with the enum in the server at some point
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
    # handle the reserved caller keyword
    arg_idxs_to_del = []
    if "caller" in a_list:
        indx = a_list.index("caller")
        arg_idxs_to_del.append(indx)
        arg_conversion_list[indx] = "Application.Caller"
    
    # handle the reserved thisworkbook keyword
    if "thiswb" in a_list:
        indx = a_list.index("thiswb")
        arg_idxs_to_del.append(indx)
        arg_conversion_list[indx] = "ThisWorkbook"

    if arg_idxs_to_del:
        arg_idxs_to_del.sort(reverse=True)
        for idx in arg_idxs_to_del:
            del arg_declaration_list[idx]

    import server 
    # Set xlpro = CreateObject("{server.xlproServerAsync._reg_progid_}")

    return f"""Function {func_name}({', '.join(arg_declaration_list)}) as Variant
    Dim xlpro As Object
    Set xlpro = GetObject("new: {server.xlproServerAsync._reg_clsid_}")
{'\n'.join(arg_range_conversion_check_list)}
    {func_name} = xlpro.{server.xlproServerAsync.execute_function_async.__name__}(ThisWorkbook, Application.Caller, "{func_name}", {', '.join(arg_conversion_list)})
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

# xlapp = win32com.client.Dispatch("Excel.Application")
# wb = xlapp.ActiveWorkbook
# get_or_create_codemodule(wb, "ThisWorkbook")


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


import pythoncom
import threading

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
        pass
    if "thiswb" in arg_names:
        idx = arg_names.index("thiswb")
        thiswb = args[idx]
        thiswb_stream = comarshal_release_and_get_stream(thiswb)
        new_args[idx] = thiswb_stream
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
        pass
    if "thiswb" in arg_names:
        idx = arg_names.index("thiswb")
        thiswb = args[idx]
        thiswb_stream = comarshal_dispatch_stream(thiswb)
        new_args[idx] = thiswb_stream
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
    if "thiswb" in arg_names:
        idx = arg_names.index("thiswb")
        arg_idxs_to_del.append(idx)
    arg_idxs_to_del.sort(reverse=True)
    for idx in arg_idxs_to_del:
        new_args.pop(idx)
    return new_args


import importlib
import types

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
        if isinstance((v:=getattr(module, name)), types.FunctionType)
    }
    return functions

import hashlib
def hash_function_call(func, *args, **kwargs):
    # Create a unique string based on the function name and its arguments
    func_name = func.__name__
    # Convert arguments to a string (including both positional and keyword arguments)
    args_str = str(args)
    kwargs_str = str(kwargs)

    # Combine the function name with its arguments
    combined = func_name + args_str + kwargs_str

    # Generate a hash using SHA-256 (you can also use MD5 or others depending on your needs)
    hash_object = hashlib.sha256(combined.encode('utf-8'))
    return hash_object.hexdigest()

import uuid
def create_random_hash():
    r = str(uuid.uuid4())
    return hashlib.sha256(r.encode('utf-8')).hexdigest()

from functools import wraps

def convert_xl_2d_types(func, args):
    _, args_and_types, _, _ = get_function_signature(func)
    ppargs = []
    for val, (a, t) in zip(args, args_and_types):
        ppargs.append(xlpro_typing.xl2DArgConvertor(val, t))
    return ppargs

import xlpro_typing
def type_converter_wrapper(func):
    """Converts the inbound data from excel into the types specified by the user 
    e.g. The user specifies a numpy array, the inbound argument is converted from a row
    major tuple to a numpy array
    """
    # @wraps(func)
    def wrapper(*args, **kwargs):
        if kwargs:
            raise NotImplementedError("kwargs not supported atm")
        ppargs = convert_xl_2d_types(func, args)
        ret = func(*ppargs, **kwargs)
        
        # convert it back to a range format
        ret2 = xlpro_typing.xl2DArgConvertor._convert_back_to_range_format(ret)
        # XXX - todo - add some logging to this. cant handle np arrays atm
        # if ret2 != ret:
        #     logger.debug("Return value was changed to suit excel's format")
        #     pass

        return ret2

    return wrapper


def com_init_dispatch_release_wrapper(func):
    """Wraps com object dispatch and release around a func.
    Also appropriately configures pythoncom coinitialise"""
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


import ctypes
def show_warning(title, message):
    # MessageBox parameters: hWnd, text, caption, uType
    ctypes.windll.user32.MessageBoxW(0, message, title, 0x30)  # 0x30 = MB_ICONWARNING



if __name__ == "__main__":
    pass



