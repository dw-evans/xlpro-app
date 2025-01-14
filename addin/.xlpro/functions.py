
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



def type_conversion_test_simple(range1:list1d):
    g = xl2DArgConvertor(range1, list1d)
    return range1