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


# NOTE, this file is the 'entry point' of xlpro to define functions in Excel.
# all functions must be 'registered' in this file *only*.

# xlpro will auto-register any module function, like below
def hello(name:str):
    return f"Hello {name}"

# use xlpro.ignore to hide a function from Excel in this namespace
@xlpro.ignore
def helper_function(a, b):
    return a + b
