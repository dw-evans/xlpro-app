import win32com.client



from pywintypes import IID
from win32com.client import Dispatch

from win32typelibs import excel as xl
from win32typelibs import vbide

from typing import Any, Callable
import typing

import inspect

# CLSID = IID('{0002E157-0000-0000-C000-000000000046}')
# vbide = win32com.client.Dispatch('{0002E157-0000-0000-C000-000000000046}')

def init_xl():
    global xlapp, wb
    xlapp = win32com.client.Dispatch("Excel.Application")
    wb = xlapp.Workbooks("workbook.xlsm")

# type hints cant be set to globals in the function, so they live here
# to make the language server happy.
xlapp:xl._Application
wb:xl._Workbook

VB_DYNAMIC_MODULE_NAME = "xlpro_async"
# VB_STATIC_MODULE_NAME = "xlpro_static"

def myfunc(a:float, b:int, c:str, d) -> str:
    """the docstring hehehe"""
    return f"{a}, {b}, {c}, {d}"

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

# vb_ = {
#     "Optional {} As {} = {}",
# }

import textwrap
vb_range_conversion_check_string = """If TypeName({arg}) = \"Range\" Then
    {arg} = {arg}.Value
EndIf"""


def convert_to_array_if_range(arg:Any):
    ...

def test_func(a:int, b:float, c, caller) -> float:
    pass

def wrap_function_with_caller_arg(func):
    """Modifies the function to take a "caller" argument if it doesn't exist"""
    f_name, args_and_types, ret_type, default_value_map = get_function_signature(func)
    # XXX todo

    def wrapper(*args, **kwargs):
        ...
        
    return wrapper

import numpy as np

import matplotlib.figure
def get_func_type(func) -> int:
    # XXX - todo - link this up with the enum in the server at some point
    f_name, args_and_types, ret_type, default_value_map = get_function_signature(func)
    if ret_type == matplotlib.figure.Figure:
        return 1
    return 0

def function_template_with_caller(func:Callable) -> str:
    """Returns function template string to send to VBA module.
    If the reserved `caller` argument is used, pass it to the execute function call.
    """
    func_type = get_func_type(func)

    func_name, args_and_types, ret_type, default_value_map = get_function_signature(func)
    docstring = func.__doc__
    arg_declaration_list = []
    arg_conversion_list = []
    arg_range_conversion_check_list = []

    # if an argument is an array, convert any range to its values.
    # if isinstance(t, typing.Iterable)
    #   if TypeName(a) = "Range" then
    #   t = t.Value

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

    pass

    if arg_idxs_to_del:
        arg_idxs_to_del.sort(reverse=True)
        for idx in arg_idxs_to_del:
            del arg_declaration_list[idx]

    return f"""Function {func_name}({', '.join(arg_declaration_list)}) as Variant
    Dim xlpro_async As Object
    Set xlpro_async = CreateObject("xlproServerAsync.Application")
    xlpro_async.register_functions_in_self
{'\n'.join(arg_range_conversion_check_list)}
    {func_name} = xlpro_async.execute_function_async({func_type}, Application.Caller, "{func_name}", {', '.join(arg_conversion_list)})
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

    # XXX to do these functions need to be imported by the com server...
    

from functools import wraps

# XXX - todo implement a 
def type_converter_wrapper(func):
    """Converts the inbound data from excel into the types specified by the user 
    e.g. The user specifies a numpy array, the inbound argument is converted from a row
    major tuple to a numpy array
    """
    ...

import pythoncom
import threading

# XXX - todo - get a better understsanding of these COM names, they can't be right lol
def comarshal_release_and_get_stream(com_dispatch):
    """Releases the COM object (PyIDispatch) from this thread and returns the stream 
    (PyIStream)"""
    return pythoncom.CoMarshalInterThreadInterfaceInStream(
        pythoncom.IID_IDispatch,
        com_dispatch
    )
def comarshal_dispatch_stream(com_stream):
    """Dispatch a com stream (PyIStream) to a com object"""
    com_obj_pyidispatch = pythoncom.CoGetInterfaceAndReleaseStream(
        com_stream, pythoncom.IID_IDispatch,
    )
    com_obj_dispatch = win32com.client.Dispatch(com_obj_pyidispatch)
    return com_obj_dispatch




import importlib

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
        if callable((v:=getattr(module, name))) and not type(v) == type
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

def hash_function_call(func, args, kwargs):
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



if __name__ == "__main__":


    pass



