import matplotlib
matplotlib.use('Agg')

import matplotlib.axes
import matplotlib.figure

import matplotlib.pyplot as plt
import random

plt.rcParams.update({
    "font.family": "Consolas", 
    "font.size": 10,              
    "axes.titlesize": "large",    
    "axes.labelsize": "medium",   
    "xtick.labelsize": "small",   
    "ytick.labelsize": "small",
    "legend.fontsize": "small",
    "figure.titlesize": "x-large"
})

# plt.style.use("seaborn-v0_8")  # For global Seaborn-like styling

import numpy as np
import matplotlib.figure

# from win32typelibs import excel as xl

import numpy as np

import xlpro
import xlpro._utils as utils
from xlpro import list1d, list2d, ndarray1d, ndarray2d

import pandas as pd


xlpro.register()(xlpro.show)
xlpro.register()(xlpro.show_image)

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

xlpro.register()(xlpro.pyrepr)
xlpro.register()(xlpro.pystr)
xlpro.register()(xlpro.pylen)
xlpro.register()(xlpro.pyshape)
xlpro.register()(xlpro.pytype)

xlpro.register()(xlpro.pygetattr)
xlpro.register()(xlpro.pygetitem)

xlpro.register()(xlpro.cpy)
xlpro.register()(xlpro.deepcpy)

xlpro.register()(xlpro.pyhash)

xlpro.register()(xlpro.condense)