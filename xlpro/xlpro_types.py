from typing import TypeVar
import numpy as np

T = TypeVar('T') # arbitrary

# this preserves all of the methods for type hinting and code completion for the user
# 1d options will force attempt reduction to a 1d vector and except if the input dimensionality is wrong.
list1d = list[T]
list2d = list[list1d[T]]

ndarray1d = np.ndarray[T]
ndarray2d = np.ndarray[T, T]