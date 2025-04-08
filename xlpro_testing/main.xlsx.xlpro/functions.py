import matplotlib
matplotlib.use('Agg')

import matplotlib.pyplot as plt
import numpy as np
import matplotlib.figure

from win32typelibs import excel as xl

import numpy as np

import xlpro
from xlpro import list1d, list2d, ndarray1d, ndarray2d

import pandas as pd

xlpro.register()(xlpro.jsonify)
xlpro.register()(xlpro.show)
xlpro.register()(xlpro.show_image)
xlpro.register()(xlpro.typ)
xlpro.register()(xlpro.copy)
xlpro.register()(xlpro.deepcopy)

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
