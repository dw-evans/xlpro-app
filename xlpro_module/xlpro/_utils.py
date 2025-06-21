import win32com.client
from PIL import Image


from pywintypes import IID
from win32com.client.dynamic import Dispatch

from typing import TYPE_CHECKING
if TYPE_CHECKING:
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
# import ._types
from xlpro import _types
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

from xlpro._types import xlproptr, ExcelArrayConverter
from xlpro import errors
import json
from xlpro._types import list1d, list2d, ndarray1d, ndarray2d


logger = logging.getLogger(__name__)

# VB_DYNAMIC_MODULE_NAME = "xlpro_async"

def get_function_types_with_fallback(func:Callable):
    try:
        hints = typing.get_type_hints(func, globalns={}, localns={})
    except NameError as e:
        # Fallback: manually replace forward references with Any
        annotations = func.__annotations__
        resolved = {}
        for k, v in annotations.items():
            if isinstance(v, str):
                try:
                    # Attempt to resolve the type
                    resolved_type = eval(v, {}, {})
                except Exception:
                    resolved_type = Any
            else:
                resolved_type = v
            resolved[k] = resolved_type
        hints = resolved

    return hints

from dataclasses import dataclass
@dataclass
class FSig:
    fname:str
    args_and_types:tuple[tuple[str, type]]
    return_type:type
    default_value_map:dict[str, Any]
    positional_only_args:tuple
    keyword_only_args:tuple

def get_function_signature(func) -> FSig:
    # Get the type hints from the function
    # type_hints = typing.get_type_hints(func)
    type_hints = get_function_types_with_fallback(func)
    
    # Get the parameter information using inspect
    signature = inspect.signature(func)
    parameters = signature.parameters
    
    # Build the output list
    argname_and_types = []
    positional_only_args = []
    keyword_only_args = []
    for param_name, param in parameters.items():
        # Get the type hint for the parameter or default to Any
        param_type = type_hints.get(param_name, Any)
        argname_and_types.append((param_name, param_type))

        # Classify argument by kind
        if param.kind in (inspect.Parameter.POSITIONAL_ONLY, inspect.Parameter.POSITIONAL_OR_KEYWORD):
            positional_only_args.append(param_name)
        if param.kind in (inspect.Parameter.KEYWORD_ONLY, inspect.Parameter.POSITIONAL_OR_KEYWORD):
            keyword_only_args.append(param_name)


    default_value_map = {
        param.name: param.default
        for param in signature.parameters.values()
        if param.default is not inspect.Parameter.empty
    }
    
    return FSig(
        fname=func.__name__, 
        args_and_types=tuple(argname_and_types), 
        return_type=type_hints.get('return', Any), 
        default_value_map=default_value_map, 
        positional_only_args=tuple(positional_only_args), 
        keyword_only_args=tuple(keyword_only_args)
    )

    # return func.__qualname__, result, type_hints.get('return', Any), default_value_map
    return func.__name__, tuple(argname_and_types), type_hints.get('return', Any), default_value_map, tuple(positional_only_args), tuple(keyword_only_args)



from xlpro._enums import FunctionTypes
import pandas as pd

def infer_func_result_type_from_type_hints(func) -> FunctionTypes:
    # XXX - todo - link this up with the enum in the server at some point
    fsig = get_function_signature(func)
    ret_type = fsig.return_type

    # if ret_type == matplotlib.figure.Figure:
    #     return FunctionTypes.py_object
    # elif ret_type == pd.DataFrame:
    #     return FunctionTypes.py_object
    # elif ret_type == pd.Series:
    #     return FunctionTypes.py_object
    if isinstance(ret_type, (str, float, int, bool)):
        return FunctionTypes.array_or_value
    # elif any([ret_type == x for x in (list1d, list2d, ndarray1d, ndarray2d)]):
    #     return FunctionTypes.array_or_value
    return FunctionTypes.py_object


# Converts python type to vb type
VB_TYPE_CONVERSION_STRINGS = {
    int: "CLngLng({})",
    float: "Cdbl({})",
    bool: "Cbool({})",
    str: "{}",
    Any: "{}"
}

# Use in function definitions
VB_TYPE_DECLARATION_STRINGS = {
    int: "{} As Long",
    float: "{} As Double",
    bool: "{} As Boolean",
    str: "{} As String",
    Any: "{} As Variant",
}

# XLPRO_DEFAULT_ARGUMENT_HINT_STR = "XLPRO_DEFAULT"
XLPRO_EMPTY_STR = "pyEmpty" # used to signal default
XLPRO_NONE_STR = "pyNone" # used to signal None (this could be a function call but seems extreme)


VB_DEFAULT_VALUE_REPR_FUNCTIONS = {
    int: lambda x: "{}".format(x),
    float: lambda x: "{}".format(float(x)),
    bool: lambda x: "True" if x else "False",
    str: lambda x: "\"{}\"".format(x),
    Any: lambda x: f"\"{XLPRO_EMPTY_STR}\"",
}





# Optional {argname} As {vbtype} = "XLPRO_DEFAULT"

VB_RANGE_CONVERSION_CHECK_STRING = """If TypeName({arg}) = \"Range\" Then
    {arg}_val = {arg}.Value
End If
"""

VB_GENERIC_CONVERSION_CHECK_STRING = """If Application.Run("'xlpro.xlam'!IsNoneOrEmpty", {arg}) Then
    {arg}_val = {arg}
Else
    {arg}_val = {arg_conversion_str}
End If
"""


RESERVED_XLPRO_KW_LOOKUPS = {
    "caller": "Application.Caller",
    "thiswb": "ActiveWorkbook",
}
RESERVED_ARGS = list(RESERVED_XLPRO_KW_LOOKUPS.keys())

# from vba_reserved_names import RESERVED_VBA_NAMES

if TYPE_CHECKING:
    from win32typelibs import excel as xl


from xlpro.vba_reserved_names import RESERVED_VBA_NAMES

def function_template_with_caller(func:Callable, fname:str=None) -> str:
    """Returns function template string to send to VBA module.
    If the reserved `caller` argument is used, pass it to the execute function call.
    """
    fsig = get_function_signature(func)
    func_name = fsig.fname
    args_and_types = fsig.args_and_types
    default_value_map = fsig.default_value_map

    if fname is not None:
        func_name = fname

    docstring = func.__doc__
    arg_declaration_list = []
    argnames_passed_to_xlpro = []
    arg_range_conversion_check_list = []
    pre_arg_dim_defs = []
    pre_arg_dim_defs = []
    argnames = [a for a, t in args_and_types]

    pre_check_template = textwrap.dedent((
        """
        If Not Application.Run("'xlpro.xlam'!CheckArgReady", {arg}) Then
            {fname} = "Promise<PENDING_PREDECENTS>"
            Exit Function
        End If"""[1:]
    ))

    if func_name == "mpl_add_line_unique":
        pass

    pre_check_arg_sequence_strs = []
    # loop over each arg and type
    # create the declaration list of strings
    # create the conversion list of strings.
    for a, t in args_and_types:
        a_orig = a
        if a in RESERVED_ARGS:
            # handle reserved kwargs
            argnames_passed_to_xlpro.append(RESERVED_XLPRO_KW_LOOKUPS[a])
            continue

        # Add a trailing underscore to the vba variable names to avoid clashes with 
        # vba reserved words
        if a.lower() in RESERVED_VBA_NAMES:
            a = f"{a}_"
            while a in argnames:
                a = f"{a}_"
        # handle any argument starting with a leading underscore
        elif a.startswith("_"):
            a = a.lstrip("_")
            a = f"{a}_"
            while a in argnames:
                a = f"{a}_"

        # should add in here some code to check if any single values are "argnotreadyexceptions" or "promises"
        # so we can significantly reduce the number of com calls.
        pre_check_arg_sequence_strs.append(pre_check_template.format(fname=fname, arg=a))

        pre_arg_dim_defs.append(f"Dim {a}_val As Variant")

        if a_orig in default_value_map.keys():
            # # Optional {argname} As {vbtype} = {defaultvalue}
            # if t in VB_DEFAULT_VALUE_REPR_FUNCTIONS.keys():
            #     arg_declaration_list.append(
            #         f"Optional {VB_TYPE_DECLARATION_STRINGS[t].format(a)} = {VB_DEFAULT_VALUE_REPR_FUNCTIONS[t](default_value_map[a])}"
            #     )
            # else:
            #     arg_declaration_list.append(
            #         f"Optional {VB_TYPE_DECLARATION_STRINGS[Any].format(a)} = {VB_DEFAULT_VALUE_REPR_FUNCTIONS[Any](default_value_map[a])}"
            #     )
            arg_declaration_list.append(
                f"Optional {a} as Variant = \"{XLPRO_EMPTY_STR}\""
            )



        # else define it in the signature with its true type
        else:
            # define the function declaration values
            if VB_TYPE_DECLARATION_STRINGS.get(t, None):
                arg_declaration_list.append(VB_TYPE_DECLARATION_STRINGS[t].format(a))
            else:
                arg_declaration_list.append(VB_TYPE_DECLARATION_STRINGS[Any].format(a))


        # # define the type conversions/casting to pass to xlpro
        # if VB_TYPE_CONVERSION_STRINGS.get(t, None):
        #     # handle any args that can be converted
        #     argnames_passed_to_xlpro.append(VB_TYPE_CONVERSION_STRINGS[t].format(a))
        # else:
        #     # handle standard args
        #     argnames_passed_to_xlpro.append("{}".format(a))

        argnames_passed_to_xlpro.append("{}_val".format(a))


        # we need to be able to handle
        if t in VB_TYPE_CONVERSION_STRINGS:
            arg_range_conversion_check_list.append(
                VB_GENERIC_CONVERSION_CHECK_STRING.format(arg=a, arg_conversion_str = VB_TYPE_CONVERSION_STRINGS[t].format(a))
            )
        else:
            arg_range_conversion_check_list.append(
                VB_GENERIC_CONVERSION_CHECK_STRING.format(arg=a, arg_conversion_str = VB_TYPE_CONVERSION_STRINGS[Any].format(a))
            )

        # convert all range inputs to their .value attribute
        if a not in RESERVED_ARGS:

            if t not in [float, int, bool, str]:
                arg_range_conversion_check_list.append(
                    VB_RANGE_CONVERSION_CHECK_STRING.format(arg=a)
                )


    from xlpro import server 
    ret = f"""Function {func_name}({', '.join(arg_declaration_list)}) as Variant
    If xlpro is Nothing Or xlpro_guid <> xlpro_guid_prev Then
        InitXlpro
    End If
{textwrap.indent('\n'.join(pre_arg_dim_defs), prefix="    ")}
{textwrap.indent('\n'.join(pre_check_arg_sequence_strs), prefix="    ")}
{textwrap.indent('\n'.join(arg_range_conversion_check_list), prefix="    ")}
    {func_name} = xlpro.{server.xlproServer.execute_function_async.__name__}(ActiveWorkbook, Application.Caller, "{func_name}"{', ' if argnames_passed_to_xlpro else ''}{', '.join(argnames_passed_to_xlpro)})
End Function
"""
    return ret

def sub_template(func:Callable) -> str:
    """Returns function template string to send to VBA module.
    If the reserved `caller` argument is used, pass it to the execute function call.
    """
    fsig = get_function_signature(func)
    func_name = fsig.fname

    from xlpro import server 

    return f"""Sub {func_name}()
    Dim xlpro As Object
    Set xlpro = GetObject("new: " & xlpro_guid)
    xlpro.{server.xlproServer.execute_sub_async.__name__} ActiveWorkbook, "{func_name}"
End Sub
"""

def get_or_create_codemodule(wb:"xl._Workbook", c_name:str) -> "vbide._CodeModule":
    proj:"vbide._VBProject" = wb.VBProject

    if not c_name in [x.Name for x in proj.VBComponents]:
        vbext_ct_StdModule            =1          # from enum vbext_ComponentType
        # comp = proj.VBComponents.Add(vbide.constants.vbext_ct_StdModule)
        comp = proj.VBComponents.Add(vbext_ct_StdModule)
        comp.Name = c_name
    else:
        comp = proj.VBComponents(c_name)

    codemod:"vbide._CodeModule" = comp.CodeModule
    return codemod


def write_to_vb_module(s:str, vb_codemod:"vbide._CodeModule"):
    vb_codemod.DeleteLines(1, vb_codemod.CountOfLines)
    vb_codemod.AddFromString(s)
    pass

def init_xlpro_vb_dynamic_component(wb:"xl._Workbook", func_register:list[Callable]):
    raise NotImplementedError("Obsoleted due to memory issues when calling this from Python")
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

# def get_xlpro_vb_dynamic_component_contents(func_register:list[Callable]) -> str:
def get_xlpro_vb_dynamic_component_contents(func_register:dict[str: Callable]) -> str:
    s_list = []

    from xlpro import server 
    s_list += [f"public const xlpro_guid as string = \"{server.xlproServer._reg_clsid_}\""]
    s_list += [f"public xlpro as object"]
    s_list += [textwrap.dedent((
        f"""
        Public xlpro_guid_prev as string
        Sub InitXlpro()
            Set xlpro = GetObject("new: " & xlpro_guid)
            xlpro_guid_prev = xlpro_guid
        End Sub
        """))
    ]

    for fname, f in func_register.items():
        if not isinstance(f, Callable):
            raise TypeError(f"Item must be a function, {type(f)}, {f}")
        s_list.append(function_template_with_caller(func=f, fname=fname))
    return "\n".join(s_list)

def get_xlpro_vb_dynamic_component_contents_subs(func_register:list[Callable]) -> str:
    s_list = []

    from xlpro import server 
    # s_list += [f"public const xlpro_guid as string = \"{server.xlproServer._reg_clsid_}\""]

    for f in func_register:
        if not isinstance(f, Callable):
            raise TypeError(f"Item must be a function, {type(f)}, {f}")
        s_list.append(sub_template(f))
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
    fsig = get_function_signature(func)
    args_and_types = fsig.args_and_types

    arg_names = [v0 for v0, v1 in args_and_types]
    new_args = list(args) #  args come in immutable (tuples)
    if "caller" in arg_names:
        idx = arg_names.index("caller")
        caller = args[idx]
        caller_stream = comarshal_release_and_get_stream(caller)
        new_args[idx] = caller_stream
    if "thiswb" in arg_names:
        idx = arg_names.index("thiswb")
        thiswb = args[idx]
        thiswb_stream = comarshal_release_and_get_stream(thiswb)
        new_args[idx] = thiswb_stream
    return new_args

def com_args_dispatch_reserved(func, args):
    """Be careful which thread this runs on! 
    Marshals the caller and thiswb reserved keyword arguments for use
    in another thread. Replaces the args with streams that can be used on another thread.
    """
    try:
        fsig = get_function_signature(func)
        args_and_types = fsig.args_and_types

        arg_names = [v0 for v0, v1 in args_and_types]
        new_args = list(args)
        if "caller" in arg_names:
            idx = arg_names.index("caller")
            caller = args[idx]
            try:
                caller_stream = comarshal_dispatch_stream(caller)
                new_args[idx] = caller_stream
            except Exception as e:
                raise e
        if "thiswb" in arg_names:
            idx = arg_names.index("thiswb")
            thiswb = args[idx]
            thiswb_stream = comarshal_dispatch_stream(thiswb)
            new_args[idx] = thiswb_stream
        return new_args
    except Exception as e:
        raise e

def get_args_minus_reserved(func, args):
    """Returns the arguments of a function but removes the reserved keywords
    as to prevent the pyidispatch strings that are generated by memory allocation
    from contaminating the string."""

    fsig = get_function_signature(func)
    args_and_types = fsig.args_and_types

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

def get_excel_args_of_func(func):
    """Returns the arguments for a function name"""
    fsig = get_function_signature(func)
    args_and_types = fsig.args_and_types
    
    arg_names = [v0 for v0, v1 in args_and_types]
    arg_idxs_to_del = []
    if "caller" in arg_names:
        idx = arg_names.index("caller")
        arg_idxs_to_del.append(idx)
    if "thiswb" in arg_names:
        idx = arg_names.index("thiswb")
        arg_idxs_to_del.append(idx)

    arg_idxs_to_del.sort(reverse=True)
    for idx in arg_idxs_to_del:
        arg_names.pop(idx)
    return arg_names


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


def count_function_args(func):
    """
    Returns the number of required and total arguments of a function.

    Args:
        func (callable): The function to inspect.

    Returns:
        dict: A dictionary with counts of 'required' and 'total' arguments.
    """
    sig = inspect.signature(func)
    params = sig.parameters.values()

    required_args = [p for p in params if p.default is inspect.Parameter.empty and p.kind in (
        inspect.Parameter.POSITIONAL_ONLY,
        inspect.Parameter.POSITIONAL_OR_KEYWORD,
        inspect.Parameter.KEYWORD_ONLY
    )]
    
    total_args = [p for p in params if p.kind in (
        inspect.Parameter.POSITIONAL_ONLY,
        inspect.Parameter.POSITIONAL_OR_KEYWORD,
        inspect.Parameter.KEYWORD_ONLY
    )]

    return len(total_args)


def get_sub_valid_functions_from_module(module_name):
    """For now, only functions wiht no arguments are supported as subroutines."""
    module = sys.modules[module_name]

    functions = [
        v
        for name in dir(module)
        if isinstance((v:=getattr(module, name)), types.FunctionType)
    ]
    valid_funcs = [
        v
        for func in functions
        if count_function_args(v:=func) == 0
    ]
    
    return valid_funcs


def get_udf_valid_function_names_from_module(module_name):
    """Returns a list of function names within a module. Returned names satisfy
    being a valid callable function from Excel
    """
    raise NotImplementedError
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
    fsig = get_function_signature(func)
    args_and_types = fsig.args_and_types
    ppargs = []
    for val, (a, t) in zip(args, args_and_types):
        ppargs.append(_types.ExcelArrayConverter(val, t))
    return ppargs

def convert_xl_2d_types_kwargs(func, kwargs):
    fsig = get_function_signature(func)
    args_and_types = fsig.args_and_types
    ppkwargs = {}
    for k, v in kwargs.items():
        for a, t in args_and_types:
            if a == k:
                ppkwargs[a] = _types.ExcelArrayConverter(v, t)
                break

    return ppkwargs


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


def jsonify(arr:list2d):
    """Converts range to json string"""
    from xlpro._types import ExcelArrayConverter
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
    return bool(re.match(r"^Promise<.+>$", arg))

def is_arg_stringified_exception(arg):
    if not isinstance(arg, str):
        return False
    return bool(re.match(r"^\w*((?:error)|(?:exception))\(.*\)$", arg, flags=re.IGNORECASE))

def pre_validate_args(args:tuple|list, kwargs:dict):
    for arg in list(args) + [v for v in kwargs.values()]:
        pre_validate_arg(arg=arg)
        
def pre_validate_arg(arg):
    if arg is None:
        raise errors.ExcelArugmentIsNoneException()
    elif is_arg_promise(arg):
        raise errors.ArugmentNotReadyException()
    elif isinstance(arg, Exception):
        raise errors.xlproArgumentExceptionError()
    elif is_arg_stringified_exception(arg):
        raise errors.xlproArgumentExceptionError()


def pre_p_an_arg(cval, target_type):
    # 0. raise error if the argument is currently a promise!
    # XXX - todo - check if this is redundant, I suspect it is
    if is_arg_promise(cval):
        raise errors.ArugmentNotReadyException()
    if isinstance(cval, Exception):
        raise errors.xlproArgumentExceptionError()
    if is_arg_stringified_exception(cval):
        raise errors.xlproArgumentExceptionError()
    
    # 1. check if its an xlproptr. Replace val with the ptr result
    # cval = copy.copy(val)
    if xlproptr.is_ptr(cval):
        cval = xlproptr.decode(cval).evaluate()

    # if m:=re.match("^PyObj<(.*)>$"):
    
    # 2. convert an argument to a target type
    ppval = _types.ExcelArrayConverter(cval, target_type)
    return ppval


def preprocess_arguments(func, args:typing.Iterable=None, kwargs:dict=None):
    if not args is None and not isinstance(args, typing.Iterable):
        raise TypeError("args must be an iterable")
    if not kwargs is None and not isinstance(kwargs, dict):
        raise TypeError("kwargs must be a dict")

    fsig = get_function_signature(func)
    args_and_types = fsig.args_and_types
    default_arguments = fsig.default_value_map

    ppargs = []
    ppkwargs = {}

    args_was_none = args == None
    kwargs_was_none = kwargs == None


    if args is not None:
        for val, (a, t) in zip(args, args_and_types):
            # handle keyword only argument which appears in args.
            if a in fsig.keyword_only_args:
                if kwargs_was_none:
                    if kwargs is None:
                        kwargs = {}
                if a in kwargs:
                    raise Exception(f"argument {a} appeard in kwargs and args!")
                kwargs[a] = val
                continue

            # check if the optional argument string has been passed
            if a in default_arguments.keys():
                if val == XLPRO_EMPTY_STR:
                    val = default_arguments[a]
                elif val == XLPRO_NONE_STR:
                    val = None
            ppargs.append(pre_p_an_arg(val, t))

    if kwargs is not None:
        for k, val in kwargs.items():
            for a, t in args_and_types:
                if a == k:
                    # check if the optional argument string has been passed
                    if a in default_arguments.keys():
                        if isinstance(val, str):
                            if val == XLPRO_EMPTY_STR:
                                val = default_arguments[a]
                            elif val == XLPRO_NONE_STR:
                                val = None
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
            self.exception = e  # Store the exceptionimport

    def get_exception(self):
        return self.exception


def get_precedents_chain(rng_dispatch:"xl.Range"):
    precedents = []
    for cell in rng_dispatch.Precedents:
        precedents.append(cell)
    return precedents

def formula_is_for_xlpro(formula:str, formulas:list[str]):
    for f in formulas:
        if f in formula:
            ret = True
            return ret
            # return True
    ret = False
    return ret
    # return False


import inspect

def get_caller_globals(frame):
    """Returns the global namespace of the module that called the current function."""
    # frame = inspect.currentframe()
    try:
        caller_frame = frame.f_back  # The frame that called `register`
        return caller_frame.f_globals  # Get the caller's globals
    finally:
        del frame  # Prevent reference cycles

import copy as _copy

def cpy(val):
    """returns a shallow copy of the object"""
    return _copy.copy(val)

def deepcpy(val):
    """returns a deep copy of the object"""
    return _copy.deepcopy(val)

from xlpro._types import xlproImage, xlproExpandedType, xlproCollapsedType

def show(val):
    """converts a value to excel-ready representation"""
    if val is None:
        raise Exception("cannot show(None)")
    tval = type(val)
    val_adj = val
    tdst:type=None
    ret = None
    calc_success = False

    if tval == pd.DataFrame:
        tdst = ndarray2d
        val_adj = val.to_numpy()
        ret = ExcelArrayConverter(val=val_adj, tdst=tdst)
        calc_success = True

    elif tval == pd.Series:
        tdst = ndarray2d
        val_adj = val.to_numpy()
        ret = ExcelArrayConverter(val=val_adj, tdst=tdst)

    elif tval in [list, tuple, list1d, list2d, ndarray1d, ndarray2d]:
        # XXX - todo - fix tuple hack in excelarrayconverter class!
        if tval == tuple:
            val_adj = list(val)
            tval = list
        tdst = tval
        ret = ExcelArrayConverter(val=val_adj, tdst=tdst)
        calc_success = True
    
    elif tval == np.ndarray:
        tdst = ndarray2d
        ret = ExcelArrayConverter(val=val_adj, tdst=tdst)
        calc_success = True

    elif isinstance(val, Exception):
        raise val
    
    elif tval == str:
        ret = val
        calc_success = True
    
    if calc_success:
        return xlproExpandedType(ret)

        # # # Constants for VARIANT type
        # # VT_ARRAY = 0x2000
        # # VT_R8 = 5  # double

        # # # Wrap as a VARIANT of type SAFEARRAY of doubles
        # # safearray_variant = VARIANT(VT_ARRAY | VT_R8, [ret.tolist()])
        # # return xlproExpandedType(safearray_variant)

        # safe_arr = automation.SafeArrayCreateVector(automation.VT_R8, 0, len(arr))
        # data_ptr = cast(safe_arr.contents.pvData, POINTER(c_double))
        # ret2 = np.ctypeslib.as_array(data_ptr, shape=(len(arr),))[:] = arr
        # return xlproExpandedType(ret2)



    # XXX - WARNING - CODE MUSTERIOSLY STOPPED WORKING?
    raise TypeError(f"type {repr(tval)} is not supported")



def px_to_pt(px, dpi):
    return px * 72 / dpi

def pt_to_px(pt, dpi):
    return pt / 72 * dpi


def show_image(val, name:str, 
    # sizex:float=None, sizey:float=None, dpi:int, format:str,
    ):
    if val is None:
        raise Exception("cannot show(None)")
    tval = type(val)
    if tval == matplotlib.figure.Figure:
        # get the singleton server
        # import xlpro.server
        # server = xlpro.server.xlproServer()

        # # find the uid of the workspace from the value received here
        # if m:=re.match(r"^functions_(.*)$", val.__globals__["name"]):
        #     uid = m.group(1)
        # else:
        #     raise Exception("could not get uid from value")
        
        # # get the workspace from the workspace uid
        # workspace = server.get_workspace_from_uid_thread_safe(uid)
        
        # # save the figure as an image in a temporary location
        # val_uid = workspace.get_uid_of_val_thread_safe(val)
        tmp:matplotlib.figure.Figure = val
        # create a tmp folder. This matches where the lockfile is created...
        tmp_path = Path(sys.executable).parent.parent.parent / ".xlpro/tmp"
        if not tmp_path.parent.exists():
            raise FileNotFoundError(f"{tmp_path.parent} does not exist!")
        tmp_path.mkdir(exist_ok=True)

        fp = tmp_path / f"{uuid.uuid4()}.png"
        tmp.savefig(fp)

        # sizex = sizex if sizex is not None else 
        ret = xlproImage(
            fp, 
            np.array(tmp.get_size_inches()) * 72,
            xl_name=name
        )
        # tmp.savefig(ret.fp, format="svg", dpi=600, backend="svg")
        tmp.savefig(ret.fp, format="png", dpi=600)

        # return the xlproImage
        return ret
    
    if tval in [str, Path]:
        if tval == str:
            fp = Path(val)
        elif tval == Path:
            fp = val
        if not fp.exists():
            raise FileNotFoundError(f"File does not exist {fp}")
        with Image.open(fp) as img:
            size = np.array(img.size)
            dpi = img.info.get("dpi")
            size_pt = px_to_pt(size, dpi)

        ret = xlproImage(
            fp=fp,
            size_pt=size_pt,
            xl_name=name
        )
        return ret
    
    # XXX - WARNING - CODE MUSTERIOSLY STOPPED WORKING?
    raise TypeError(f"type {repr(tval)} is not supported")


def show_image_with_seed(val, name:str, seed):
    return show_image(val, name)

def pow(val:np.ndarray, exp):
    return val ** exp

def mul(val:np.ndarray, rhs):
    return val * rhs

def div(val:np.ndarray, rhs):
    return val / rhs

def add(val:np.ndarray, rhs):
    return val + rhs

def subtract(val:np.ndarray, rhs):
    return val - rhs


def pytype(val):
    if val is None:
        return None
    if isinstance(val, xlproCollapsedType):
        return f"xlproCollapsedType[{str(type(val.data))}]"
    ret = str(type(val))
    return ret

def pyrepr(val):
    if val is None:
        return None
    return repr(val)

def pystr(val):
    if val is None:
        return None
    return str(val)

def pyhash(vals):
    s = "".join([str(x) if x in (float, int, str) else str(id(x)) for x in vals])
    return hash(s)
        

# def vectorize(func_name:str, args_list) -> list1d:
#     ret = []
#     func = 

#     for args in args_list
#         ret.append()

def int2rgb(color:int): # -> tuple[int, int, int]:
    r = color & 0xFF
    g = (color >> 8) & 0xFF
    b = (color >> 16) & 0xFF
    return (r, g, b)

def rgb2int(color:tuple[int, int, int]):
    r, g, b = color
    return (b << 16) + (g << 8) + r


# class ExpandedIterable(list):
#     def __init__(self, val):
#         if not isinstance(val, typing.Iterable):
#             raise TypeError(f"Type: {type(val)} cannot be expanded")

# def expand(val):
#     return ExpandedIterable(val)

import operator
def pygetitem(obj, val:int):
    """Typed wrapper for getitem"""
    return operator.getitem(obj, val)
    
def pygetattr(obj, attrname:str, default:Any):
    """Typed wrapper for getattr"""
    return getattr(obj, attrname, default)
    

# class PyNone:
#     pass

#     def __repr__(self):
#         return "<pynone>"
    
#     def __str__(self):
#         return repr(self)
    
#     @staticmethod
#     def is_pynone(s:str):
#         return s == "<pynone>"


# def pynone():
#     """Returns Python None"""
#     return PyNone()

class xlRange:
    """Signal class to preserve range passing version of xl.Range"""
    pass


# import queue
# MAIN_THREAD_QUEUE = queue.Queue()

# # Wrapper to run function on the main thread
# def run_on_main_thread(func):
#     def wrapper(*args, **kwargs):
#         result_q = queue.Queue()
#         MAIN_THREAD_QUEUE.put((func, args, kwargs, result_q))
#         return result_q.get()  # block until result is available
#     return wrapper

XLAPP_LOCK = threading.Lock()

def comsafe(func):
    @wraps(func)
    def inner(*args, **kwargs):
        # return func(*args, **kwargs)
        with XLAPP_LOCK:
            return func(*args, **kwargs)
        # return run_on_main_thread(func)(*args, **kwargs)
    return inner

# def comsafe(func):
#     @wraps(func)
#     def inner(*args, **kwargs):
#         with XLAPP_LOCK:
#             return func(*args, **kwargs)
#     return inner

# def comsafe(func):
#     @wraps(func)
#     def inner(*args, **kwargs):
#         try:
#             if not XLAPP_LOCK.acquire(timeout=5):
#                 raise RuntimeError("Timeout waiting for Excel COM access")
#             return func(*args, **kwargs)
#         except Exception as e:
#             pass

#         finally:
#             XLAPP_LOCK.release()
#     return inner


@comsafe
def create_table_if_not_exists(caller, table_name:str):
    # Locate the table
    ws = caller.Parent
    wb = ws.Parent
    found = False
    for ws in wb.Worksheets:
        for tbl in ws.ListObjects:
            if tbl.Name == table_name:
                table = tbl
                sheet = ws
                found = True
                tbl_range = tbl.Range
                break
        if found:
            break

    if not found:
        tbl_range = caller.Cells(2, 1)
        newtbl = ws.ListObjects.Add(
            SourceType=1,
            Source=tbl_range,
            XlListObjectHasHeaders=1,
        )
        newtbl.Name = table_name

    return f"xlTable(\"{table_name}\" @ '{tbl_range.Parent.Name}'!{tbl_range.Address})"




@comsafe
def create_table_from_df(caller, df:pd.DataFrame, table_name:str):
    wb = caller.Parent.Parent
    # Locate the table
    found = False
    for ws in wb.Worksheets:
        for tbl in ws.ListObjects:
            if tbl.Name == table_name:
                table = tbl
                sheet = ws
                found = True
                break
        if found:
            break

    if not found:
        return create_table_if_not_exists(caller=caller, table_name=table_name)
        # raise ValueError(f"Table '{table_name}' not found.")

    # Get starting cell
    top_left = table.Range.Cells(1, 1)

    # Get shape of DataFrame
    n_rows, n_cols = df.shape
    if n_rows == 0 or n_cols == 0:
        raise ValueError("DataFrame is empty or has no columns.")
    tblRange = table.Range

    tblcontentsRange = sheet.Range(table.Range.Cells(2,1), table.Range.Cells(tblRange.Rows.Count, tblRange.Columns.Count))
    tblcontentsRange.Formula2 = ""

    # Resize table range BEFORE writing anything
    new_range = sheet.Range(top_left, top_left.Cells(n_rows+1, n_cols))  # +1 row for header

    table.Resize(new_range)

    # Write headers
    header_range = sheet.Range(top_left, top_left.Cells(1, n_cols))
    header_range.Value = [df.columns.tolist()]

    # Write data
    data_start = top_left.Cells(2, 1)
    data_end = data_start.Cells(n_rows, n_cols)
    data_range = sheet.Range(data_start, data_end)
    data_range.Value = tuple(df.itertuples(index=False, name=None))

    print(f"✅ Table '{table_name}' updated with {n_rows} rows and {n_cols} columns.")

    return f"xlTable(\"{table_name}\" @ '{header_range.Parent.Name}'!{header_range.Address})"



def replace_table_with_df(df: pd.DataFrame, table_name: str = "Table1"):

    # Locate the table
    found = False
    for ws in wb.Worksheets:
        for tbl in ws.ListObjects:
            if tbl.Name == table_name:
                table = tbl
                sheet = ws
                found = True
                break
        if found:
            break

    if not found:
        raise ValueError(f"Table '{table_name}' not found.")

    # Get starting cell
    top_left = table.Range.Cells(1, 1)

    # Get shape of DataFrame
    n_rows, n_cols = df.shape
    if n_rows == 0 or n_cols == 0:
        raise ValueError("DataFrame is empty or has no columns.")

    # Resize table range BEFORE writing anything
    new_range = sheet.Range(top_left, top_left.Cells(n_rows+1, n_cols))  # +1 row for header
    table.Resize(new_range)

    # Write headers
    header_range = sheet.Range(top_left, top_left.Cells(1, n_cols))
    header_range.Value = [df.columns.tolist()]

    # Write data
    data_start = top_left.Cells(2, 1)
    data_end = data_start.Cells(n_rows, n_cols)
    data_range = sheet.Range(data_start, data_end)
    data_range.Value = tuple(df.itertuples(index=False, name=None))

    print(f"✅ Table '{table_name}' updated with {n_rows} rows and {n_cols} columns.")




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




