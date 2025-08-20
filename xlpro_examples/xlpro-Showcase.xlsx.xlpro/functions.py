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


import seaborn as sns

def seaborn_visualise_bivariate(
    n: int,
    mean: ndarray1d,
    cov: ndarray2d,
    cmap: str = "mako",
    figsize_mm: ndarray1d = np.array([]),
):
    """Generate a bivariate visualisation with colormap and KDE
    using seaborn
    """
    rng = np.random.RandomState(0)
    x, y = rng.multivariate_normal(mean, cov, n).T

    # Draw a combo histogram and scatterplot with density contours
    f, ax = plt.subplots(figsize=figsize_mm / 25.4)
    sns.scatterplot(x=x, y=y, s=5, color=".15", ax=ax)
    sns.histplot(x=x, y=y, bins=50, pthresh=0.1, cmap=cmap, ax=ax)
    sns.kdeplot(x=x, y=y, levels=5, color="w", linewidths=1, ax=ax)

    return f


def np_linspace(start, stop, n: int):
    return np.linspace(start, stop, n)


def pallette():
    tab10 = plt.get_cmap("tab10").colors  # Returns 10 RGB tuples
    return tab10


def mpl_create_figure(seed=None) -> "Figure":
    fig, ax = plt.subplots()
    return fig


def mpl_get_axes_of_fig(fig: "Figure"):
    return list(fig.get_axes())


def mpl_add_line_unique(
    ax: "Axes",
    x: ndarray1d,
    y: ndarray1d,
    tag: str,
    fmt: str = None,
    color: list1d = None,
    linewidth=2,
    markersize=12,
):
    if tag is None:
        raise Exception(f"InvalidTag")
    for line in ax.get_lines():
        if tag == getattr(line, "_tag", None):
            line.remove()
    (ln,) = ax.plot(x, y, fmt, color=color, linewidth=linewidth, markersize=markersize)
    ln._tag = tag
    return ln


@xlpro.comsafe
def visualise_color(rgb: ndarray1d[np.int32], caller: "xlpro.xlRange"):
    rgb = rgb
    caller.Interior.Color = xlpro._rgb2int(np.astype(rgb, np.int32))
    return str(rgb)


def mpl_set_xax_name(ax, name: str):
    ax.set_xlabel(name)
    return f"X-axis label set to '{name}'."


def mpl_set_yax_name(ax: "Axes", name: str):
    ax.set_ylabel(name)
    return f"Y-axis label set to '{name}'."


def mpl_set_ax_title(ax: "Axes", title: str):
    ax.set_title(title)
    return f"Axes title set to '{title}'"


def mpl_set_fig_suptitle(fig: "Figure", suptitle: str):
    fig.suptitle(suptitle)
    return f"Figure title set to '{suptitle}'"


def mpl_set_xlims(ax: "Axes", xlims: list1d, seed=None):
    ax.set_ylim(*xlims)
    return f"X-axis limits set to {xlims}."


def mpl_set_ylims(ax: "Axes", ylims: list1d, seed=None):
    ax.set_ylim(*ylims)
    return f"Y-axis limits set to {ylims}."


def mpl_add_legend(
    ax: "Axes",
    handles: list1d,
    labels: list1d,
    loc: str = "upper left",
    frameon: bool = False,
    ncol: int = 1,
):
    legend = ax.legend(
        labels=labels,
        handles=handles,
        labelspacing=1.0,
        fontsize=9.0,
        frameon=False,
        loc=loc,
        ncol=ncol,
    )
    return legend


def glob_path(dirpath:Path, glob:str="*", absolute=False, folders_only=False):
    g = dirpath.glob(glob)
    if folders_only:
        ret =  [x for x in g if g.is_dir()]
    else:
        ret =  [x for x in g]

    if absolute:
        return [x.resolve() for x in ret]
    return [x.relative_to(xlpro.get_xlpro_wd()) for x in ret]

def pd_read_csv(fp:Path):
    return pd.read_csv(fp)


import yfinance as yf
import datetime

def get_stock_prices(tickers: list1d[str], days:int) -> pd.DataFrame:
    """
    Fetches daily stock prices for the past number of days.
    
    Args:
        ticker (str): Stock ticker symbol (e.g., 'AAPL', 'GOOG')
        days (int): Number of days prior to today's date
    
    Returns:
        pd.DataFrame: DataFrame with Date as index and columns: Open, High, Low, Close, Volume
    """
    end_date = datetime.datetime.now()
    start_date = end_date - datetime.timedelta(days=days)
    
    # Download data
    # Progress=False is required to prevent OSError - conflict with pydevd. For reference.
    df = yf.download(tickers, start=start_date, end=end_date, progress=False)
    df.reset_index(inplace=True)
    # yf returns timezone-naive data. Correcting to allow passing to Excel natively.
    df["Date"] = df["Date"].dt.tz_localize("UTC")

    return df
