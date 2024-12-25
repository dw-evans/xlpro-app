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

VB_MODULE_NAME = "xlpro"
VB_STATIC_MODULE_NAME = "xlpro_static"

def get_or_create_xlpro_component():
    proj:vbide._VBProject = wb.VBProject

    if not VB_MODULE_NAME in [x.Name for x in proj.VBComponents]:
        comp = proj.VBComponents.Add(vbide.constants.vbext_ct_StdModule)
        comp.Name = VB_MODULE_NAME
    else:
        comp = proj.VBComponents(VB_MODULE_NAME)

    codemod:vbide._CodeModule = comp.CodeModule
    codemod.AddFromString("hello from codemod")

    pass


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
    return func.__qualname__, result, type_hints.get('return', Any)

# Converts python type to vb type
vb_type_conversion_strings = {
    int: "Cint({})",
    float: "Cdbl({})",
    bool: "Cbool({})",
    Any: "{}"
}

# Use in function definitions
vb_type_declaration_strings = {
    int: "{} As Integer",
    float: "{} As Double",
    bool: "{} As Boonlean",
    Any: "{} As Variant"
}

def convert_to_array_if_range(arg:Any):
    ...

def function_template(func:Callable):
    """Returns function template string to send to VBA module."""
    f_name, args_and_types, ret_type = get_function_signature(func)
    docstring = func.__doc__
    arg_declaration_list = []
    arg_conversion_list = []
    for a, t in args_and_types:
        arg_declaration_list.append(vb_type_declaration_strings[t].format(a))
        arg_conversion_list.append(vb_type_conversion_strings[t].format(a))
    
    return f"""Function {f_name}({', '.join(arg_declaration_list)}) as Variant
Dim xlpro As Object
Set xlpro = CreateObject("xlproServer.Application")
{f_name} = xlpro.ExecuteFunction("add_numbers", {', '.join(arg_conversion_list)})
End Function"""

if __name__ == "__main__":
    init_xl()
    get_or_create_xlpro_component()



