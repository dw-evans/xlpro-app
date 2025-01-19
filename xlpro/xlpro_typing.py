from typing import TypeVar
import numpy as np
from typing import Any
import typing

import logging
logger = logging.getLogger(__name__)

T = TypeVar('T') # arbitrary

# this preserves all of the methods for type hinting and code completion for the user
# 1d options will force attempt reduction to a 1d vector and except if the input dimensionality is wrong.
list1d = list[T]
list2d = list[list1d[T]]

ndarray1d = np.ndarray[T]
ndarray2d = np.ndarray[T, T]

class xl2DArgConvertor:
    """Excel will convert ranges to 2d tuple arrays. Use this class to convert
    the arguments before the user's code intereacts with them to ease debugging.
    The user should only require up to 2d data with float, int, str and maybe
    boolean types so we can implement this custom logic confidently.
    """
    def __new__(cls, val:Any, tdst:type):
        if not isinstance(val, tuple):
            return val

        origin = typing.get_origin(tdst)
        args = typing.get_args(tdst)
        # extract the final or intermediate dtype to use.
        dtype = cls._convert_args_to_np_dtype(args)

        if tdst in [list1d, ndarray1d]:
            intermediate = np.array(val)
            shape = intermediate.shape
            if all([x > 1 for x in shape]):
                raise TypeError("Provided value is not compatible with list1d")
            if tdst == list1d:
                return intermediate.flatten().tolist()
            elif tdst == ndarray1d:
                return intermediate.flatten()
        
        elif tdst in [list2d, ndarray2d]:
            intermediate = np.array(val)
            if tdst == list2d:
                return intermediate.tolist()
            return intermediate

        elif tdst in [list, np.ndarray]:
            intermediate = np.array(val)
            if tdst == list:
                return intermediate.tolist()
            return intermediate
        else:
            # handle things such as list[float]
            intermediate = np.array(val)
            if origin == list:
                intermediate = np.array(val, dtype=dtype)
                return intermediate.tolist()
            
            elif origin == np.ndarray:
                intermediate = np.array(val, dtype=dtype)
                return intermediate
        if not tdst in [typing.Any]:
            raise Exception(f"Type {tdst} is too complicated or not supported for conversion attempt")
    
        return val

    @classmethod
    def _convert_args_to_np_dtype(cls, args:tuple[type]):

        # XXX - todo - we may want to support strings in a fancy way...
        # For string arrays we probably want to fetch the .Text property of 
        # a range and not the .Value property (specifically concerned with 
        # 0 != "" (blank cells))

        if len(args) == 0:
            return None
        a0 = args[0]
        if not all([x == a0 for x in args]):
            raise TypeError(f"Type arguments {a0} are too complicated to convert")
        if a0 in [T, list[T]]:
            # if the type has not been given for the list1d or list2d case return no type
            return None
        if isinstance(a0, typing.GenericAlias):
            if not typing.get_origin(a0) == list:
                raise TypeError(f"Type {a0} cannot be processed, origin must be list")
            return cls._convert_args_to_np_dtype(typing.get_args(a0))
        
        elif a0 == float:
            return np.float64
        elif a0 == int:
            return np.int64
        elif a0 == bool:
            return np.bool
        elif any([np.issubdtype(a0, x) for x in (np.floating, np.integer, np.bool)]):
            # if the user specifies a specific dtype, resort to that
            # numpy should catch any conversion errors when attempting to convert
            return a0
        else:
            # Let the numpy default behaviour run
            # XXX - todo - check how this behaves
            return None 

    @classmethod
    def _convert_back_to_range_format(cls, val):
        if isinstance(val, typing.Iterable):
            # wrap any iterable that doesn't have an iterable second type in a list
            if not isinstance(val[0], typing.Iterable):
                return [val]
            # wrap a string in a list also.
            if isinstance(val, str):
                return [val]

        return val




if __name__ == "__main__":
    r1  = xl2DArgConvertor(((1, 2,),), list1d)
    print(f"r1={r1}")
    r2  = xl2DArgConvertor(((1, 2,),), list2d)
    print(f"r2={r2}")
    r5  = xl2DArgConvertor(((1, 2,),), list)
    print(f"r5={r5}")
    r6  = xl2DArgConvertor(((1, 2,),), list[float])
    print(f"r6={r6}")
    r7  = xl2DArgConvertor(((1, 2,),), list[int])
    print(f"r7={r7}")
    r8  = xl2DArgConvertor(((0, 1,),), list[bool])
    print(f"r8={r8}")
    r3  = xl2DArgConvertor(((1, 2,),), ndarray1d)
    print(f"r3={r3}, dtype={r3.dtype}")
    r4  = xl2DArgConvertor(((1, 2,),), ndarray2d)
    print(f"r4={r4}, dtype={r4.dtype}")
    r10 = xl2DArgConvertor(((0, 1,),), np.ndarray[np.float16])
    print(f"r10={r10}, dtype={r10.dtype}")
    r11 = xl2DArgConvertor(((0, 1,),), np.ndarray[np.float32])
    print(f"r11={r11}, dtype={r11.dtype}")
    r12 = xl2DArgConvertor(((0, 1,),), np.ndarray[np.float64])
    print(f"r12={r12}, dtype={r12.dtype}")
    r13 = xl2DArgConvertor(((0, 1,),), np.ndarray[np.int16])
    print(f"r13={r13}, dtype={r13.dtype}")
    r14 = xl2DArgConvertor(((0, 1,),), np.ndarray[np.int32])
    print(f"r14={r14}, dtype={r14.dtype}")
    r15 = xl2DArgConvertor(((0, 1,),), np.ndarray[np.int64])
    print(f"r15={r15}, dtype={r15.dtype}")


pass
