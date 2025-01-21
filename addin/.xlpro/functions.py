
import time
import matplotlib.pyplot as plt
import numpy as np

from pathlib import Path
import sys

# wd = Path(__file__).parent
# sys.path.append(r"C:\Users\Daniel Evans\projects\xlpro\xlpro")

# wd = Path(__file__).parent
# sys.path.append(wd / "../xlpro")


import matplotlib.figure
import utils
import json

from win32typelibs import excel as xl
import win32com.client

# These must be the exact same objects!...
from xlpro_types import list1d, list2d, ndarray1d, ndarray2d
from xlpro_typing import xl2DArgConvertor

import numpy as np
import typing



def type_conversion_test_simple(range1:list1d):
    return np.array(range1) + 3.4

def type_conversion_test_simple2(range1:list2d):
    return range1

def type_conversion_test_simple3(range1:list):
    g = xl2DArgConvertor(range1, list)
    return range1

# def type_conversion_test_simple(range1:list1d):
#     # a = xl2DArgConvertor(range1, list1d)
#     # b = xl2DArgConvertor(range1, list2d)
#     # c = xl2DArgConvertor(range1, list2d[int])
#     # d = xl2DArgConvertor(range1, ndarray1d)
#     # e = xl2DArgConvertor(range1, ndarray2d)
#     # f = xl2DArgConvertor(range1, ndarray2d[bool])
#     g = xl2DArgConvertor(range1, list1d)
#     return range1


# # XXX - todo - correct list parsing...
def py_plot(xdata:list1d, ydata:list1d, xlims:list, ylims:list, xaxislabel:str, yaxislabel:str) -> matplotlib.figure.Figure:
    fig, ax = plt.subplots()
    # XXX - todo - excel always spits things out in a 2d format, a preprocess could be useful
    ax.set_xlim(np.array(xlims).tolist()[0])
    ax.set_ylim(np.array(ylims).tolist()[0])
    ax.set_xlabel(xaxislabel)
    ax.set_ylabel(yaxislabel)
    ax.plot(xdata, ydata)
    return fig



def func_name_to_function(func_name):
    return globals()[func_name]

def d2dd(vector:list):
    return np.array(vector).reshape(-1, 1)

def get_args_of_function(func_name) -> list:
    func = func_name_to_function(func_name)
    f_name, args_and_types, ret_type, default_value_map = utils.get_function_signature(func)
    args = [a for a, t in args_and_types]
    return d2dd(args)

# def pytranspose(vector:list):
#     return np.array(vector).T

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

# import os

# def write_to_file(s:str, p:str, caller:xl.Range, thiswb:xl._Workbook):
#     # XXX - todo - we can't access the workbook object while the user is editing it...
#     # How do we handle this? retry calculation..?
#     # wb:xl._Workbook = win32com.client.Dispatch(thiswb)
#     wd = Path(thiswb.FullName).parent
#     wd_before = os.getcwd()
#     os.chdir(wd)
#     p_obj = Path(p)
#     with open(p_obj, "w") as f:
#         f.write(s)
#     ret = f"File written to {str(p_obj)}"
#     caller.Interior = xlrgb(0,0,255)
#     os.chdir(wd_before)
#     return ret

# def xlrgb(r:int, g:int, b:int):
#     """Returns the color as an int interpreted by excel"""
#     return b*256**2 + g*256 + r

