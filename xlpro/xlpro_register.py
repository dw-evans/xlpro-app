import numpy as np
from win32com.client import Dispatch

def add_numbers(a:float, b:float):
    return a + b + 100.0

def make_array(n:int, m:int):
    return np.random.random((n, m))

def funky_formula(a, b, caller, thiswb):
    return
