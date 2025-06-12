import matplotlib
import matplotlib.lines

import xlpro._utils
matplotlib.use('Agg')

import matplotlib.axes
import matplotlib.pyplot as plt
import numpy as np
import matplotlib.figure

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from win32typelibs import excel as xl

import numpy as np

import xlpro
from xlpro import list1d, list2d, ndarray1d, ndarray2d

import pandas as pd

import xlpro._utils as utils

xlpro.register()(xlpro.jsonify)
xlpro.register()(xlpro.show)
xlpro.register()(xlpro.show_image)
xlpro.register()(xlpro.pytype)
xlpro.register()(xlpro.cpy)
xlpro.register()(xlpro.deepcpy)
xlpro.register()(xlpro.int2rgb)
xlpro.register()(xlpro.pyrepr)

# xlpro.register()(xlpro.conditional_formatter_example)


import operator
xlpro.register(fname="getitem")(xlpro._utils.xlpro_getitem)
xlpro.register(fname="pylen")(len)
pass
# xlpro.register(operator.setitem)
# xlpro.register(operator.delitem)
# xlpro.register(len)


def pd_function_create(x) -> pd.DataFrame:
    ret = pd.DataFrame(x[1:], columns=x[0])
    return ret

def pd_function_filter(df, col_name) -> pd.DataFrame:
    df:pd.DataFrame
    return df[col_name]

def create_figure() -> matplotlib.figure.Figure:
    fig, ax = plt.subplots()
    fig.tight_layout()

    return fig

def get_axes_of_figure(fig:matplotlib.figure.Figure):
    return list(fig.get_axes())

def get_line2ds_of_axes(ax:matplotlib.axes.Axes, seed):
    return list(ax.get_lines())

def add_line(ax:matplotlib.axes.Axes, x:ndarray1d, y:ndarray2d):
    ln = ax.plot(x, y)
    return ln

def add_line_unique(ax:matplotlib.axes.Axes, x:ndarray1d, y:ndarray2d, tag:str):
    if tag is None:
        raise Exception(f"InvalidTag")
    for line in ax.get_lines():
        if tag == getattr(line, "_tag", None):
            line.remove()
    ln, = ax.plot(x, y, color=None)
    ln._tag = tag
    return ln

def add_point_marker_unique(ax:matplotlib.axes.Axes, x:ndarray1d, y:ndarray2d, tag:str):
    if tag is None:
        raise Exception(f"InvalidTag")
    for line in ax.get_lines():
        if tag == getattr(line, "_tag", None):
            line.remove()
    ln, = ax.plot(x, y, marker="o", markersize=5)
    ln._tag = tag
    return ln

def hex_to_rgb_tuple(hex_str:str):
    """Convert a hex color string (#RRGGBB or #RRGGBBAA) to a tuple of integers (R, G, B) or (R, G, B, A)."""
    hex_str = hex_str.strip().lstrip('#')
    
    if len(hex_str) not in (6, 8):
        raise ValueError("Hex color must be in the format #RRGGBB or #RRGGBBAA")

    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)

    if len(hex_str) == 8:
        a = int(hex_str[6:8], 16)
        return (r, g, b, a)
    else:
        return (r, g, b)

def condense(iterable_val):
    return xlpro.xlproCollapsedType(iterable_val)

# @xlpro.ignore()
def get_all_artists(ax:matplotlib.axes.Axes):
    lines = ax.lines                          # List of Line2D objects
    collections = ax.collections              # For things like scatter plots, contour fills
    patches = ax.patches                      # Rectangles, Circles, Polygons, etc.
    texts = ax.texts                          # All Text objects including annotations
    images = ax.images                        # Any image (imshow, etc.)
    artists = ax.artists                      # Miscellaneous Artist objects
    tables = ax.tables                        # Tables
    containers = ax.containers                # Bar containers, etc.

    all_artists = (
        ax.lines +
        ax.collections +
        ax.patches +
        ax.texts +
        ax.images +
        ax.artists +
        ax.containers +
        list(ax.tables.values())  # tables is a dict
    )

    return all_artists


def add_line_point_unique(ax:matplotlib.axes.Axes, ln:matplotlib.lines.Line2D, xpos:float, tag:str):
    for t in ax.texts:
        if tag == getattr(t, "_tag", None):
            t.remove()
    
    xdata = list(ln.get_xdata())
    ydata = list(ln.get_ydata())

    ln_color = ln.get_color()

    # Interpolate or extrapolate y value at xpos
    if len(xdata) >= 2:
        # Sort xdata and ydata together by x
        sorted_points = sorted(zip(xdata, ydata))
        xs, ys = zip(*sorted_points)
        yval = np.interp(xpos, xs, ys)
    else:
        # Not enough points to interpolate
        yval = 0.0

    # Redraw the line (only necessary outside interactive mode)
    ax.draw_artist(ln)



    # Annotate the new point
    label = str((float(xpos), float(yval)))

    annot = ax.annotate(label,
                xy=(xpos, yval),
                xytext=(5, 5),
                textcoords='offset points',
                color='white',
                fontsize=8.0)
    annot._tag = tag

    pt = add_point_marker_unique(ax, xpos, yval, f"{tag}_0")

    return annot



def show_image_with_seed(val, name:str, seed):
    return xlpro.show_image(val, name)

def create_figure_with_seed(seed):
    fig, ax = plt.subplots()
    fig.tight_layout()
    # Remove figure and axes background
    fig.patch.set_alpha(0)           # Figure background
    ax.patch.set_alpha(0)            # Axes background

    # Set all text (title, labels, ticks) to white
    ax.title.set_color('white')
    ax.xaxis.label.set_color('white')
    ax.yaxis.label.set_color('white')
    ax.tick_params(colors='white')

    # Set spine (axis lines) colors to white
    for spine in ax.spines.values():
        spine.set_color('white')
    return fig

def add_legend(ax:matplotlib.axes.Axes, handles:list1d, labels:list1d):
    legend = ax.legend(
        labels=labels,
        handles=handles,
        labelspacing=1.0,        # Line spacing
        fontsize=9.0,            # Text size
        frameon=False            # No border or background
    )

    # Set font color to white
    for text in legend.get_texts():
        text.set_color('white')
        pass

    return legend

def pyhash(vals):
    s = "".join([str(x) if x in (float, int, str) else str(id(x)) for x in vals])
    return hash(s)

def get_ax0(fig:matplotlib.figure.Figure):
    ret = fig.axes[0]
    return ret

def plot_function(x:ndarray1d, y:ndarray1d) -> matplotlib.figure.Figure:
    fig, ax = plt.subplots()
    ax.plot(x, y)
    return fig

def set_fig_size_mm(fig:matplotlib.figure.Figure, size:ndarray1d[np.float64]) -> matplotlib.figure.Figure:
    fig.set_size_inches(size / 25.4)
    return fig


def tick_formatter_partial(pattern:str):
    def func(x, pos):
        return pattern.format(x)
    ret = func
    return ret

def format_figure_tick_marks(fig):
    ...

def visualise_color(rgb:ndarray1d[int], caller:'xl.Range'):
    rgb = rgb[0]
    caller.Interior.Color = utils.rgb2int(rgb)
    return str(rgb)


def expander_test():
    fig, ax = plt.subplots()
    ret = [
        list((1, 2, 3)),
        tuple((1, 2, 3)),
        "hello",
        100,
        -123.123,
        True,
        False,
        fig,
        ax,
    ]
    return ret


def randtest():
    return [
        1, 2, 3, "hello", 100.0, True, False
    ]

def np_create_random(n:int, m:int):
    return np.random.random((n, m))

def np_create_random_typed(n:int, m:int) -> ndarray2d:
    return np.random.random((n, m))


def get_address(caller:'xl.Range'):
    return caller.Address


def pow(val:np.ndarray, exp):
    return val ** exp

def mul(val:np.ndarray, mu):
    return val * mu

def div(val:np.ndarray, di):
    return val / di

def add(val:np.ndarray, ad):
    return val + ad

# import random
# import time

# def get_self_color_value(caller:'xl.Range') -> int:
#     try:
#         # time.sleep(random.random() * 10.0)
#         print(caller.Value)
#         ret = caller.Interior.Color
#     except Exception as e:
#         raise e
#     return ret