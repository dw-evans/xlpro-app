
import time
import matplotlib.pyplot as plt
import numpy as np

import matplotlib.figure
import utils
import json

from pathlib import Path
from win32typelibs import excel as xl
import win32com.client


def test_function(a:int, b:float, c:bool, d, e:list, caller, thiswb):
    return "done..."

def range_test_function(a:list):
    return repr(a)

def basic_function(a:float, b:float):
    return a ** b

def py_plot(xdata:list, ydata:list, xlims:list, ylims:list, xaxislabel:str, yaxislabel:str) -> matplotlib.figure.Figure:
    fig, ax = plt.subplots()
    ax.plot(xdata, ydata)
    return fig

def func_name_to_function(func_name):
    return globals()[func_name]

def get_args_of_function(func_name) -> list:
    func = func_name_to_function(func_name)
    f_name, args_and_types, ret_type, default_value_map = utils.get_function_signature(func)
    args = [a for a, t in args_and_types]
    return d2dd(args)

def d2dd(vector:list):
    return np.array(vector).reshape(-1, 1)

def pytranspose(vector:list):
    return np.array(vector).T

def decon(vector:list):
    return json.dumps(vector)

def jsonize(arr:list):
    if len(arr[0]) != 2:
        raise Exception("Please provide a nx2 array of key:value pairs")
    ret = {}
    for row in arr:
        if not isinstance(row[0], str):
            raise TypeError("Ensure the first column values are all strings")
        ret[row[0]] = row[1]

    return json.dumps(ret, indent=2)

import os

def write_to_file(s:str, p:str, caller, thiswb):
    # XXX - todo - we can't access the workbook object while the user is editing it...
    # How do we handle this? retry calculation..?

    wb:xl._Workbook = win32com.client.Dispatch(thiswb)
    wd = Path(wb.FullName).parent
    wd_before = os.getcwd()
    os.chdir(wd)
    p_obj = Path(p)
    with open(p_obj, "w") as f:
        f.write(s)
    os.chdir(wd_before)
    return f"File written to {str(p_obj)}"

# def test(a, caller, thiswb):
#     pass
#     caller_obj = win32com.client.Dispatch(caller)
#     thiswb_obj = win32com.client.Dispatch(thiswb)
#     return "test complete"


# if __name__ == "__main__":
#     fig = create_figure()
#     plt.show()