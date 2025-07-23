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


import matplotlib.font_manager as fm

fm.fontManager.addfont(r"C:\Users\Daniel Evans\Downloads\JetBrainsMono-2.304\fonts\ttf\JetBrainsMono-Regular.ttf")

plt.rcParams.update({
    "font.family": "JetBrains Mono", 
    "font.size": 10,              
    "axes.titlesize": "large",    
    "axes.labelsize": "medium",   
    "xtick.labelsize": "small",   
    "ytick.labelsize": "small",
    "legend.fontsize": "small",
    "figure.titlesize": "x-large"
})


import matplotlib.figure
import matplotlib.axes

def pallette():
    tab10 = plt.get_cmap("tab10").colors  # Returns 10 RGB tuples
    return tab10

def mpl_create_figure(seed=None) -> matplotlib.figure.Figure:
    fig, ax = plt.subplots()
    return fig

def mpl_get_axes_of_fig(fig:matplotlib.figure.Figure):
    return list(fig.get_axes())

def mpl_add_line_unique(
    ax:matplotlib.axes.Axes, 
    x:ndarray1d, 
    y:ndarray1d, 
    tag:str, 
    fmt:str=None,
    color:list1d=None, 
    linewidth=2,
    markersize=12,
):
    if tag is None:
        raise Exception(f"InvalidTag")
    for line in ax.get_lines():
        if tag == getattr(line, "_tag", None):
            line.remove()
    ln, = ax.plot(x, y, fmt, color=color, linewidth=linewidth, markersize=markersize)
    ln._tag = tag
    return ln

@xlpro.comsafe
def visualise_color(rgb:ndarray1d[np.int32], caller:'xl.Range'):
    rgb = rgb
    caller.Interior.Color = utils._rgb2int(np.astype(rgb, np.int32))
    return str(rgb)



def mpl_set_xax_name(ax, name:str):
    ax.set_xlabel(name)
    return f"X-axis label set to '{name}'."
    return ax



def mpl_set_yax_name(ax, name:str):
    ax.set_ylabel(name)
    return f"Y-axis label set to '{name}'."
    return ax

def mpl_set_ax_title(ax:matplotlib.axes.Axes, title:str):
    ax.set_title(title)
    return f"Axes title set to '{title}'"
    return ax

def mpl_set_fig_suptitle(fig:matplotlib.figure.Figure, suptitle:str):
    fig.suptitle(suptitle)
    return f"Figure title set to '{suptitle}'"
    return fig


def mpl_set_xlims(ax:matplotlib.axes.Axes, xlims:list1d, seed=None):
    try:
        ax.set_xlim(*xlims)
    except:
        try:
            ax.set_xlim(*xlims)
        except Exception as e:
            raise e
    return f"X-axis limits set to {xlims}."


def mpl_set_ylims(ax:matplotlib.axes.Axes, ylims:list1d, seed=None):
    ax.set_ylim(*ylims)
    return f"Y-axis limits set to {ylims}."


def mpl_add_legend(ax:matplotlib.axes.Axes, handles:list1d, labels:list1d, loc:str="upper left", frameon:bool=False, ncol:int=1):
    legend = ax.legend(
        labels=labels,
        handles=handles,
        labelspacing=1.0,        # Line spacing
        fontsize=9.0,            # Text size
        frameon=False,            # No border or background
        loc=loc,
        ncol=ncol
    )
    return legend






import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
# sns.set_theme(style="light")

def seaborn_visualise_bivariate(
    n:int, 
    mean: ndarray1d, 
    cov:ndarray2d, 
    cmap:str="mako", 
    figsize_mm:ndarray1d=np.array([]),
):
    rng = np.random.RandomState(0)
    x, y = rng.multivariate_normal(mean, cov, n).T

    # Draw a combo histogram and scatterplot with density contours
    f, ax = plt.subplots(figsize=figsize_mm/25.4)
    sns.scatterplot(x=x, y=y, s=5, color=".15", ax=ax)
    sns.histplot(x=x, y=y, bins=50, pthresh=.1, cmap=cmap, ax=ax)
    sns.kdeplot(x=x, y=y, levels=5, color="w", linewidths=1, ax=ax)

    return f

from pathlib import Path
def relpath(p:Path):
    return str(p)