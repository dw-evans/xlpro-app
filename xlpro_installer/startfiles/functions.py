import matplotlib
matplotlib.use('Agg') # Non-GUI mpl backend is required for xlpro

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

import xlpro
import xlpro._utils as utils
from xlpro import list1d, list2d, ndarray1d, ndarray2d

# Display methods
xlpro.register()(xlpro.show)
xlpro.register()(xlpro.show_image)

# Python operators
xlpro.register()(xlpro.pypow)
xlpro.register()(xlpro.pymul)
xlpro.register()(xlpro.pydiv)
xlpro.register()(xlpro.pymod)
xlpro.register()(xlpro.pyadd)
xlpro.register()(xlpro.pysub)
xlpro.register()(xlpro.pynot)
xlpro.register()(xlpro.pyeq)
xlpro.register()(xlpro.pyne)
xlpro.register()(xlpro.pylt)
xlpro.register()(xlpro.pyle)
xlpro.register()(xlpro.pygt)
xlpro.register()(xlpro.pyge)

# Common Python functions
xlpro.register()(xlpro.pyrepr)
xlpro.register()(xlpro.pystr)
xlpro.register()(xlpro.pylen)
xlpro.register()(xlpro.pyshape)
xlpro.register()(xlpro.pytype)

# Object getters
xlpro.register()(xlpro.pygetattr)
xlpro.register()(xlpro.pygetitem)

# Object copying
xlpro.register()(xlpro.cpy)
xlpro.register()(xlpro.deepcpy)

# Hasher to create unique seeds
xlpro.register()(xlpro.pyhash)

xlpro.register()(xlpro.condense)


def hello(name:str):
    return f"hello {name}"

