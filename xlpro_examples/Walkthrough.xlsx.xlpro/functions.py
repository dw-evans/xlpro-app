import matplotlib
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import typing
from pathlib import Path

if typing.TYPE_CHECKING:
    from matplotlib.figure import Figure
    from matplotlib.axes import Axes

import xlpro
from xlpro import list1d, list2d, ndarray1d, ndarray2d

matplotlib.use("Agg")

@xlpro.ignore
def register_defaults():
    # Display methods
    xlpro.register(xlpro.show)
    xlpro.register(xlpro.show_image)
    xlpro.register(xlpro.create_table_from_df)

    # Python operators
    xlpro.register(xlpro.pypow)
    xlpro.register(xlpro.pymul)
    xlpro.register(xlpro.pydiv)
    xlpro.register(xlpro.pymod)
    xlpro.register(xlpro.pyadd)
    xlpro.register(xlpro.pysub)
    xlpro.register(xlpro.pynot)
    xlpro.register(xlpro.pyeq)
    xlpro.register(xlpro.pyne)
    xlpro.register(xlpro.pylt)
    xlpro.register(xlpro.pyle)
    xlpro.register(xlpro.pygt)
    xlpro.register(xlpro.pyge)

    # Common Python functions
    xlpro.register(xlpro.pyrepr)
    xlpro.register(xlpro.pystr)
    xlpro.register(xlpro.pylen)
    xlpro.register(xlpro.pyshape)
    xlpro.register(xlpro.pytype)

    # Object getters
    xlpro.register(xlpro.pygetattr)
    xlpro.register(xlpro.pygetitem)

    # Object copying
    xlpro.register(xlpro.cpy)
    xlpro.register(xlpro.deepcpy)

    # Hasher to create unique seeds
    xlpro.register(xlpro.pyhash)

    xlpro.register(xlpro.condense)
    xlpro.register(xlpro.pylist)
    xlpro.register(xlpro.pytuple)
register_defaults()

@xlpro.ignore
def configure():
    import matplotlib.font_manager as fm

    fm.fontManager.addfont(
        r"C:\Users\Daniel Evans\Downloads\JetBrainsMono-2.304\fonts\ttf\JetBrainsMono-Regular.ttf"
    )
    plt.rcParams.update(
        {
            "font.family": "JetBrains Mono",
            "font.size": 10,
            "axes.titlesize": "large",
            "axes.labelsize": "medium",
            "xtick.labelsize": "small",
            "ytick.labelsize": "small",
            "legend.fontsize": "small",
            "figure.titlesize": "x-large",
        }
    )
configure()



def hello(name:str):
    return f"hello {name}"

def custom_adder(a, b):
    return a + b

def custom_adder_matrix(a, b):
    a_mat = np.array(a)
    return a_mat + b

def adder_or_subtractor(a, b, do_add:bool=True):
    if do_add:
       return a + b
    return a - b

from xlpro import list1d, list2d, ndarray1d, ndarray2d

def custom_adder_matrix(matrix:ndarray2d, constant:float):
    return matrix + constant

def custom_adder_always_vector(vector:ndarray1d, constant:float):
    return vector + constant

def create_matplotlib_figure(xdata:ndarray1d, ydata:ndarray1d, linelabel:str):
    """Function which takes x and y data and generates a matplotlib figure, 
    and labels the line in the legend."""

    # Create the figure
    fig, ax = plt.subplots()

    # Plot the data
    ln, = ax.plot(xdata, ydata)

    # Add axes labels
    ax.set_xlabel("x-values")
    ax.set_ylabel("y-values")

    # Add a legend
    ax.legend([ln], [linelabel], frameon=False)

    # resize the figure to 170 by 100 mm.
    fig.set_size_inches(np.array((170, 100)) / 25.4)

    return fig # type: matplotlib.figure.Figure

def get_banana_price():
    key = "Banana"
    # Perhaps you would fetch your prices from an api call,
    # for example
    price_cents_per_kg = {
        "Apple":     99,
        "Orange":    140,
        "banana":    87,
        "Kiwi":      180,
        "Mango":     204,
        "Pineapple": 316,
    }
    return price_cents_per_kg[key]




def myfunc(a:int, b:float, c):
    return f"{a}, {b}, {c}"

def myfunc2(a:ndarray1d[np.int32], b:ndarray1d[np.datetime64], c):
    x = list(a)
    y = a.tolist()
    # return a.tolist()
    return f"{a}, {b}, {c}"