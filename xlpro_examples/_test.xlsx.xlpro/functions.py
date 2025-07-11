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


"""
test_cast_list
test_cast_list1d
test_cast_list2d
test_cast_list_int_
test_cast_list1d_int_
test_cast_list2d_int_
test_cast_np_ndarray
test_cast_ndarray1d
test_cast_ndarray2d
test_cast_np_ndarray_int_
test_cast_np_ndarray_np_int32_
test_cast_ndarray1d_np_int32_
test_cast_ndarray2d_np_int32_
"""

def test_cast_list(a:list):
    return a

def test_cast_list1d(a:list1d):
    return a

def test_cast_list2d(a:list2d):
    return a

def test_cast_list_int_(a:list[int]):
    return a

def test_cast_list1d_int_(a:list1d[int]):
    return a

def test_cast_list2d_int_(a:list2d[int]):
    return a

def test_cast_np_ndarray(a:np.ndarray):
    return a

def test_cast_ndarray1d(a:ndarray1d):
    return a

def test_cast_ndarray2d(a:ndarray2d):
    return a

def test_cast_np_ndarray_int_(a:np.ndarray[int]):
    return a

def test_cast_np_ndarray_np_int32_(a:np.ndarray[np.int32]):
    return a

def test_cast_np_ndarray_np_float64_(a:np.ndarray[np.float64]):
    return a

def test_cast_ndarray1d_np_int32_(a:ndarray1d[np.int32]):
    return a

def test_cast_ndarray2d_np_int32_(a:ndarray2d[np.int32]):
    return a

def test_cast_ndarray1d_np_datetime64_(a:ndarray1d[np.datetime64]):
    return a

def test_cast_ndarray2d_np_datetime64_(a:ndarray2d[np.datetime64]):
    return a

def test_cast_ndarray2d_np_datetime64_show_dt(a:ndarray2d[np.datetime64]):
    return a

def test_pd_datetime_show(a:ndarray2d[np.datetime64]):
    df = pd.DataFrame(a)
    return df

def test_pd_mix_datetime_show(a:ndarray1d[np.datetime64], b:ndarray1d[np.int32]):
    df = pd.DataFrame({
        "date": a,
        "value": b,
    })
    return df

def test_pd_create_series(a:ndarray1d):
    ret = pd.Series(a)
    return ret

def test_condensed_1(a:list2d):
    ret = xlpro.xlproCollapsedType(a)
    return ret