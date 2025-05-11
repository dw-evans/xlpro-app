import matplotlib
matplotlib.use('Agg')

import matplotlib.pyplot as plt
import numpy as np
import matplotlib.figure

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from win32typelibs import excel as xl

import numpy as np

import xlpro
from xlpro import list1d, list2d, ndarray1d, ndarray2d

import pandas as pd

xlpro.register()(xlpro.jsonify)
xlpro.register()(xlpro.show)
xlpro.register()(xlpro.show_image)
xlpro.register()(xlpro.typ)
xlpro.register()(xlpro.cpy)
xlpro.register()(xlpro.deepcpy)
xlpro.register()(xlpro.int2rgb)
# xlpro.register()(xlpro.conditional_formatter_example)

def pd_function_create(x) -> pd.DataFrame:
    ret = pd.DataFrame(x[1:], columns=x[0])
    return ret

def pd_function_filter(df, col_name) -> pd.DataFrame:
    df:pd.DataFrame
    return df[col_name]

def plot_function(x:ndarray1d, y:ndarray1d) -> matplotlib.figure.Figure:
    fig, ax = plt.subplots()
    ax.plot(x, y)
    return fig

def add_line(fig, x:ndarray1d, y:ndarray1d) -> matplotlib.figure.Figure:
    fig:matplotlib.figure.Figure
    ax = fig.axes[0]

    ax.plot(x, y, "g--")

    return fig

import random
import time

def get_self_color_value(caller:'xl.Range') -> int:
    try:
        # time.sleep(random.random() * 10.0)
        print(caller.Value)
        ret = caller.Interior.Color
    except Exception as e:
        raise e
    return ret