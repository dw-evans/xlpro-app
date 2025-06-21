import matplotlib
matplotlib.use('Agg')

import matplotlib.pyplot as plt

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
xlpro.register()(xlpro.show_image_with_seed)
xlpro.register()(xlpro.pytype)
xlpro.register()(xlpro.cpy)
xlpro.register()(xlpro.deepcpy)
xlpro.register()(xlpro.pow)
xlpro.register()(xlpro.mul)
xlpro.register()(xlpro.div)
xlpro.register()(xlpro.add)
xlpro.register()(xlpro.subtract)
xlpro.register()(xlpro.pyhash)
xlpro.register()(xlpro.pygetattr)
xlpro.register()(xlpro.pygetitem)
xlpro.register()(xlpro.pystr)
xlpro.register()(xlpro.pyrepr)
xlpro.register()(xlpro.pylen)
xlpro.register()(xlpro.condense)


def np_linspace(
    start,
    stop,
    num: int = 50,
    endpoint: bool = True,
    retstep: bool = False,
    axis:int = 0,
    dtype: None = None,
    device:str = "cpu"
):
    return np.linspace(
        start=start,
        stop=stop,
        num=num,
        endpoint=endpoint,
        retstep=retstep,
        dtype=dtype,
        axis=axis,
        device=device
    )

def pallette():
    tab10 = plt.get_cmap("tab10").colors  # Returns 10 RGB tuples
    return tab10

@xlpro.comsafe
def visualise_color(rgb:ndarray1d[np.int32], caller:'xl.Range'):
    rgb = rgb
    caller.Interior.Color = utils.rgb2int(np.astype(rgb, np.int32))
    return str(rgb)

def mpl_create_figure_with_seed(seed) -> matplotlib.figure.Figure:
    fig, ax = plt.subplots()
    # fig.tight_layout()
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

def mpl_add_legend(ax:matplotlib.axes.Axes, handles:list1d, labels:list1d, loc:str="upper left"):
    legend = ax.legend(
        labels=labels,
        handles=handles,
        labelspacing=1.0,        # Line spacing
        fontsize=9.0,            # Text size
        frameon=False,            # No border or background
        loc=loc,
    )


    return legend

def mpl_set_xax_name(ax, name:str):
    ax.set_xlabel(name)
    return ax

def mpl_set_yax_name(ax, name:str):
    ax.set_ylabel(name)
    return ax


def mpl_set_xlims(ax:matplotlib.axes.Axes, xlims:ndarray1d):
    ax.set_xlim(*xlims)

def mpl_set_ylims(ax:matplotlib.axes.Axes, ylims:ndarray1d):
    ax.set_ylim(*ylims)



import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta

def get_stock_prices_last_week(ticker: str, days:int=365) -> pd.DataFrame:
    """
    Fetches daily stock prices for the past 7 days (including weekends).
    
    Args:
        ticker (str): Stock ticker symbol (e.g., 'AAPL', 'GOOG')
    
    Returns:
        pd.DataFrame: DataFrame with Date as index and columns: Open, High, Low, Close, Volume
    """
    end_date = datetime.now()
    start_date = end_date - timedelta(days=days)
    
    # Download data
    df = yf.download(ticker, start=start_date.strftime('%Y-%m-%d'), end=end_date.strftime('%Y-%m-%d'), interval='1d')
    
    # Optionally, reset index if needed
    df.reset_index(inplace=True)
    
    return df


