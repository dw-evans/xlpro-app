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

VB_DYNAMIC_MODULE_NAME = "xlpro"
VB_STATIC_MODULE_NAME = "xlpro_static"



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
    bool: "{} As Boonlean",
    str: "{} As String",
    Any: "{} As Variant",
}

vb_ = {
    "Optional {} As {} = {}",
}


def convert_to_array_if_range(arg:Any):
    ...

def test_func(a:int, b:float, c, caller) -> float:
    pass

def function_template_with_caller(func:Callable):
    """Returns function template string to send to VBA module.
    If the reserved `caller` argument is used, pass it to the execute function call.
    """
    f_name, args_and_types, ret_type, default_value_map = get_function_signature(func)
    docstring = func.__doc__
    arg_declaration_list = []
    arg_conversion_list = []
    
    for a, t in args_and_types:
        arg_declaration_list.append(vb_type_declaration_strings[t].format(a))
        arg_conversion_list.append(vb_type_conversion_strings[t].format(a))
    
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

    # a_list_new = 
    # for a in :
    #     if a in default_value_map.keys():
    #         idx = arg_declaration_list.index(a)
            

    #     return f"""Function {f_name}({', '.join(arg_declaration_list)}) as Variant
    #     On Error GoTo ErrorHandler
    #     initialize_xlpro
    #     {f_name} = g_xlpro.execute_function("{f_name}", {', '.join(arg_conversion_list)})
    # ErrorHandler:
    #     If Err.Number <> 0 Then
    #         MsgBox Err.Description
    #     End If
    #     Resume Next
    # End Function
    # """
    return f"""Function {f_name}({', '.join(arg_declaration_list)}) as Variant
    initialize_xlpro
    {f_name} = g_xlpro.execute_function("{f_name}", {', '.join(arg_conversion_list)})
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

def dispatch_converter_wrapper(func):
    """Wraps a function so that the arguments are dispatched when they
    come to python. Saves the user needing to do this. Some of the secret
    sauce...
    """
    @wraps
    def wrapper(*args):
        f_name, args_and_types, ret_type, _ = get_function_signature(func)
        arg_names = [v0 for v0, v1 in args_and_types]
        new_args = args
        if "caller" in arg_names:
            idx = arg_names.index("caller")
            caller = args[idx]
            caller_dispatch = Dispatch(caller)
            new_args[idx] = caller_dispatch
            # Caller type could be many things, likely just a Range.
            pass
        if "thiswb" in arg_names:
            idx = arg_names.index("thiswb")
            thiswb = args[idx]
            thiswb_dispatch = Dispatch(thiswb)
            new_args[idx] = thiswb_dispatch
            pass
        return func(*new_args)
    return wrapper


if __name__ == "__main__":


    pass



