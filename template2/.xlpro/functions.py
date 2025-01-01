
import time
import matplotlib.pyplot as plt
import numpy as np

from pathlib import Path
import sys

wd = Path(__file__).parent
sys.path.append(wd / "../../xlpro_async")

import matplotlib.figure
import utils
import json

from win32typelibs import excel as xl
import win32com.client
from xlpro_typing import list1d, list2d, ndarray1d, ndarray2d, xl2DArgConvertor
import numpy as np
import typing



def type_conversion_test_simple(range1):
    # a = xl2DArgConvertor(range1, list1d)
    # b = xl2DArgConvertor(range1, list2d)
    # c = xl2DArgConvertor(range1, list2d[int])
    # d = xl2DArgConvertor(range1, ndarray1d)
    # e = xl2DArgConvertor(range1, ndarray2d)
    f = xl2DArgConvertor(range1, ndarray2d[bool])
    # g = xl2DArgConvertor(range1, list1d)
    return f