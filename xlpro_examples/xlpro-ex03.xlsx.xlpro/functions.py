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



import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
sns.set_theme(style="dark")

def seaborn_visualise_bivariate(
    n:int, 
    mean: ndarray1d, 
    cov:ndarray2d, 
    cmap:str, 
    figsize_mm:ndarray1d=np.array([]),
):
    # Simulate data from a bivariate Gaussian
    # n = 10000
    # mean = [0, 0]
    # cov = [(2, .4), (.4, .2)]
    rng = np.random.RandomState(0)
    x, y = rng.multivariate_normal(mean, cov, n).T

    # Draw a combo histogram and scatterplot with density contours
    f, ax = plt.subplots(figsize=figsize_mm/25.4)
    sns.scatterplot(x=x, y=y, s=5, color=".15", ax=ax)
    sns.histplot(x=x, y=y, bins=50, pthresh=.1, cmap=cmap, ax=ax)
    sns.kdeplot(x=x, y=y, levels=5, color="w", linewidths=1, ax=ax)

    return f