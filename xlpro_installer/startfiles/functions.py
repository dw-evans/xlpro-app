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

matplotlib.use("Agg")  # Use Non-GUI backend


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

# This file is the 'entry point' of xlpro to define functions in Excel.


def hello_world(name: str):
    return f"Hello {name}!"


# use xlpro.ignore to hide a function from Excel
@xlpro.ignore
def helper_function(a, b):
    return a + b
