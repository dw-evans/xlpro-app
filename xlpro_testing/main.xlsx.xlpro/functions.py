import matplotlib.pyplot as plt
import numpy as np

import matplotlib.figure

from win32typelibs import excel as xl

import numpy as np

import xlpro
from xlpro import list1d, list2d, ndarray1d, ndarray2d
import xlpro._utils
from xlpro._types import ExcelArrayConverter

import pandas as pd

xlpro.register()(xlpro.jsonify)

@xlpro.wrap_jsonify()
def test_function_1(rng1:ndarray1d):
    return rng1 * np.random.random(rng1.size)

@xlpro.wrap_jsonify()
def test_function_2(rng1:ndarray1d, rng2:ndarray1d, rng3:ndarray1d):
    return rng1 + rng2 + rng3

@xlpro.register()
def test_function_3(rng1:ndarray1d):
    return rng1

@xlpro.ignore()
def type_conversion_test_simple(range1:list1d):
    return np.array(range1) + 3.4
def type_conversion_test_simple2(range1:list2d):

    return range1

def type_conversion_test_simple3(range1:list):
    g = ExcelArrayConverter(range1, list)
    return range1


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
# func_name = "test_function1_json"
def func_name_to_function(func_name):
    return globals()[func_name]

def d2dd(vector:list):
    return np.array(vector).reshape(-1, 1)

@xlpro.ignore()
def get_args_of_function(func_name) -> list:
    func = func_name_to_function(func_name)
    f_name, args_and_types, ret_type, default_value_map = xlpro._utils.get_function_signature(func)
    args = [a for a, t in args_and_types]
    return d2dd(args)


def pd_function_create(x) -> pd.DataFrame:
    ret = pd.DataFrame(x[1:], columns=x[0])
    return ret

def pd_function_filter(df, col_name):
    df:pd.DataFrame
    return df[col_name]
