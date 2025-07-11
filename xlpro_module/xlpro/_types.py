import typing
from typing import TypeVar, Generic, Any
import numpy as np
from xlpro import errors
from typing import Iterable
import numpy.typing as npt
import re
from dataclasses import dataclass, field
from win32com.client import GetActiveObject, Dispatch
import pythoncom
# from win32typelibs import excel as xl
from pathlib import Path
import pandas as pd

T = TypeVar('T')

class TypeStrEnums:
    LIST1D = "list1d"
    LIST2D = "list2d"
    NDARRAY1D = "ndarray1d"
    NDARRAY2D = "ndarray2d"

list1d = typing.Annotated[list[T], TypeStrEnums.LIST1D]
list2d = typing.Annotated[list[list[T]], TypeStrEnums.LIST2D]

ndarray1d = typing.Annotated[npt.NDArray[T], TypeStrEnums.NDARRAY1D]
ndarray2d = typing.Annotated[npt.NDArray[T], TypeStrEnums.NDARRAY2D]

STR_TO_TYPE_MAP = {
    TypeStrEnums.LIST1D: list1d,
    TypeStrEnums.LIST2D: list2d,
    TypeStrEnums.NDARRAY1D: ndarray1d,
    TypeStrEnums.NDARRAY2D: ndarray2d,
}


def _get_ndarray_annotated_dtype(_t):
    _args = typing.get_args(_t)[0]
    _dtype_alias = typing.get_args(_args)[1]
    _dtype = typing.get_args(_dtype_alias)[0]
    _dtype = None if _dtype == T else _dtype
    return _dtype

def _get_list1d_annotated_dtype(_t):
    _args = typing.get_args(_t)[0]
    _dtype = typing.get_args(_args)[0]
    _dtype = None if _dtype == T else _dtype
    return _dtype

def _get_list2d_annotated_dtype(_t):
    _args = typing.get_args(_t)[0]
    _dtype_alias = typing.get_args(_args)[0]
    _dtype = typing.get_args(_dtype_alias)[0]
    _dtype = None if _dtype == T else _dtype
    return _dtype


def _get_generic_dtype(_t):
    _args = typing.get_args(_t)
    if len(_args) > 1:
        raise TypeError(f"GenericAlias type annotation is too complicated: {_t}")
    _ret = _args[0]
    return _ret


def _is_annotated_type(tp) -> bool:
    return typing.get_origin(tp) is typing.Annotated

def _is_generic_alias_type(tp) -> bool:
    return isinstance(tp, typing.GenericAlias)


@dataclass
class xlproptr:
    """Class encoding a dynamic excel range link. Sister function ptr() to be used
    within Excel to convert an excel-native pointer to a string which is recognizable
    within this class using the decode method.
    """

    wb_path:str
    ws_name:str
    rng_addr:str
    
    def __post_init__(self):
        # XXX - todo - implement validation for xlproptr
        pass

    def __repr__(self):
        return f"xlproptr({self.wb_path}::{self.ws_name}::{self.rng_addr})"
    
    def __str__(self):
        return self.__repr__()

    @classmethod
    def decode(cls, s:str):
        """Primary entry point into class"""
        mtch = re.search(r"^\*\<(.+)::(.+)::(.+)\>", s)
        return cls(*mtch.groups())
    
    @staticmethod
    def is_ptr(s:str):
        """Crude check for a valid ptr"""
        if not isinstance(s, str):
            return False
        try:
            xlproptr.decode(s)
            return True
        except Exception:
            return False
    
    # def evaluate(self, wb_stream):
    def evaluate(self):
        """Evaluates the cell pointer's value. Returns 2d array for ranges, or a value for
        a 1x1 range object.
        For reliability, the calling workbook is passed so we ensure we don't spin up a new
        excel session.
        Workflow will be to overwrite the argument if xlproptr.is_ptr(arg) -> 
        i.e. pparg = xlproptr.decode(arg)
             args[arg_idx] = pparg
             kwargs[arg_key] = pparg
        """
        try:
            # XXX - todo - must improve deferred calculation here!

            # from utils import comarshal_release_and_get_stream, comarshal_dispatch_stream
            pythoncom.CoInitialize()
            # wb:xl._Workbook = comarshal_dispatch_stream(wb_stream)
            # Dip in and out of python com to retrieve the data. Then release.
            # XXX - todo - Check if getactiveobject is actually good here.
            # xlapp = GetActiveObject("Excel.Application")
            xlapp = Dispatch("Excel.Application")
            try:
                xlapp.Workbooks
            except AttributeError:
                raise errors.ExcelNotAccessibleError
            # XXX - todo - Check if this wb path matching check is reliable for server locations for example.
            # e.g. mapped drives may convert to server addresses. I believe resolve() corrects for this...
            if not Path(self.wb_path).resolve().__str__() in [Path(wb.FullName).resolve().__str__() for wb in xlapp.Workbooks]:
                xlapp.Workbooks.Open(self.wb_path)
            ret = xlapp.Workbooks(str(Path(self.wb_path).name)).Sheets(self.ws_name).Range(self.rng_addr).Value
            xlapp = None
            # comarshal_release_and_get_stream(wb)
            pythoncom.CoUninitialize()
            return ret
        except Exception as e:
            pass
            raise e

    
    # if any array argument arrives as a string, a pre-process step should be used
    # to evaluate the range.


@dataclass
class xlproImage:
    fp:Path
    size_pt:Iterable[float]
    xl_name:str

import pywintypes

# XXX - todo - apparently this is sensitive to imports...
# type checking broke when I refactored, presumably changed the origin of some of the objects?
class ExcelArrayConverter:
    """Excel will convert ranges to 2d tuple arrays. Use this class to convert
    2d arrays into a user-specified type before the user's code receives it.
    
    The user should only require up to 2d data with float, int, str and maybe
    boolean types so we can implement this custom logic confidently.
    """

    def __new__(cls, val:Any, tdst:type):

        
        dtype = None
        # if annotated type is provided, get the base type
        if _is_annotated_type(tdst):
            args = typing.get_args(tdst)
            tstr = args[1]
            if tstr == TypeStrEnums.NDARRAY1D:
                tdstnew = ndarray1d
            elif tstr == TypeStrEnums.NDARRAY2D:
                tdstnew = ndarray2d
            elif tstr == TypeStrEnums.LIST1D:
                tdstnew = list1d
            elif tstr == TypeStrEnums.LIST2D:
                tdstnew = list2d
            else:
                raise TypeError(f"Annotated type {tdst} not supported")
        elif _is_generic_alias_type(tdst):
            origin = typing.get_origin(tdst)
            args = typing.get_args(tdst)
            if len(args) > 1:
                raise TypeError(f"GenericAlias type {tdst} is too complex to coerce")
            # dtype = args[0]
            if origin == list:
                tdstnew = list
            elif origin == np.ndarray:
                tdstnew = np.ndarray
            else:
                raise TypeError(f"GenericAlias type {tdst} not supported. Origin {origin} not supported.")
        else:
            tdstnew = tdst

        # 1 by 1 ranges are passed as a single value, cast these back to an array tuple
        # so it is cast to the right type
        if tdstnew in [list1d, list2d, ndarray1d, ndarray2d, list, np.ndarray]:
            if isinstance(val, (float, int, str, bool, pywintypes.TimeType)):
                val = ((val,),)
        # return the incoming value if not a 2d array needing conversion
        elif not isinstance(val, (tuple, list, np.ndarray)):
            return val

        pass

        # get the dtype of tdst
        if tdstnew in [ndarray1d, ndarray2d]:
            dtype = _get_ndarray_annotated_dtype(tdst)
        elif tdstnew == list1d:
            dtype = _get_list1d_annotated_dtype(tdst)
        elif tdstnew == list2d:
            dtype = _get_list2d_annotated_dtype(tdst)
            
        # handle the list/ndarray GenericAlias casees
        elif tdstnew == list:
            if _is_generic_alias_type(tdst):
                dtype = _get_generic_dtype(tdst)
            elif tdst == list:
                dtype = None
        elif tdstnew == np.ndarray:
            if _is_generic_alias_type(tdst):
                dtype = _get_generic_dtype(tdst)
            elif tdst == np.ndarray:
                dtype = None
        
        def handle_ndarray1d(val, dtype):
            intermediate = np.array(val, dtype=dtype)
            shape = intermediate.shape
            # implicitly the shape is 2d, convert it to the user chosen list or ndarray
            if (not len(shape) == 1) and (all([x > 1 for x in shape])):
                raise TypeError(f"Provided value is not compatible with {tdstnew}")
            return intermediate.flatten()

        def handle_list1d(val, dtype):
            intermediate = np.array(val, dtype=dtype)
            shape = intermediate.shape
            # implicitly the shape is 2d, convert it to the user chosen list or ndarray
            if (not len(shape) == 1) and (all([x > 1 for x in shape])):
                raise TypeError(f"Provided value is not compatible with {tdstnew}")
            return intermediate.flatten().tolist()
        
        def handle_ndarray2d(val, dtype):
            intermediate = np.array(val, dtype=dtype)
            # convert to a 2d
            shape = intermediate.shape
            if len(shape) == 1:
                ret = intermediate.reshape(-1, 1)
            elif len(shape) == 2:
                ret = intermediate
            else:
                raise TypeError(f"Cannot convert shape {shape} to ndarray2d")
            return ret
        
        def handle_ndarray(val, dtype):
            intermediate = np.array(val, dtype=dtype)
            return intermediate
        
        def handle_list(val, dtype):
            intermediate = np.array(val, dtype=dtype)
            return intermediate.tolist()

        if tdstnew == ndarray1d:
            newval = handle_ndarray1d(val=val, dtype=dtype)

        # Note: also used to send data back to Excel
        elif tdstnew == ndarray2d:
            newval = handle_ndarray2d(val=val, dtype=dtype)

        elif tdstnew == list1d:
            newval = handle_list1d(val=val, dtype=dtype)

        # Note: also used to send data back to Excel
        elif tdstnew == list2d:
            newval = handle_list(val=val, dtype=dtype)

        elif tdstnew == list:
            newval = handle_list(val=val, dtype=dtype)
            
        elif tdstnew == np.ndarray:
            newval = handle_ndarray(val=val, dtype=dtype)
    
        elif tdstnew == typing.Any:
            newval = val
        else:
            raise TypeError(f"Type {tdstnew} is too complicated or not supported for coersion.")

        return newval


    @classmethod
    def _convert_back_to_range_format(cls, val):
        if isinstance(val, pd.DataFrame):
            return val.to_numpy()

        elif isinstance(val, typing.Iterable):
            # wrap any iterable that doesn't have an iterable second type in a list
            if not isinstance(val[0], typing.Iterable):
                return [val]
            # wrap a string in a list also.
            if isinstance(val, str):
                return [val]

        return val
    

class xlproExpandedType:
    """Class to signal that an array is to be expanded, overwrites xlproCollapsedType"""
    oned_direction_rowwise = True
    def __init__(self, arraydata):
        if isinstance(arraydata, (xlproCollapsedType, xlproExpandedType)):
            arraydata = arraydata.data
        self.data = arraydata


class xlproCollapsedType:
    """Wrapper class that signals a result to forcibly be collapsed, overwrites xlproExpandedType"""
    def __init__(self, arraydata):
        if isinstance(arraydata, (xlproCollapsedType, xlproExpandedType)):
            arraydata = arraydata.data
        self.data = arraydata



if __name__ == "__main__":
    # r1  = ExcelArrayConverter(((1, 2,),), list1d)
    # print(f"r1={r1}")
    # r2  = ExcelArrayConverter(((1, 2,),), list2d)
    # print(f"r2={r2}")
    # r5  = ExcelArrayConverter(((1, 2,),), list)
    # print(f"r5={r5}")
    # r6  = ExcelArrayConverter(((1, 2,),), list[float])
    # print(f"r6={r6}")
    # r7  = ExcelArrayConverter(((1, 2,),), list[int])
    # print(f"r7={r7}")
    # r8  = ExcelArrayConverter(((0, 1,),), list[bool])
    # print(f"r8={r8}")
    # r3  = ExcelArrayConverter(((1, 2,),), ndarray1d)
    # print(f"r3={r3}, dtype={r3.dtype}")
    # r4  = ExcelArrayConverter(((1, 2,),), ndarray2d)
    # print(f"r4={r4}, dtype={r4.dtype}")
    # r10 = ExcelArrayConverter(((0, 1,),), np.ndarray[np.float16])
    # print(f"r10={r10}, dtype={r10.dtype}")
    # r11 = ExcelArrayConverter(((0, 1,),), np.ndarray[np.float32])
    # print(f"r11={r11}, dtype={r11.dtype}")
    # r12 = ExcelArrayConverter(((0, 1,),), np.ndarray[np.float64])
    # print(f"r12={r12}, dtype={r12.dtype}")
    # r13 = ExcelArrayConverter(((0, 1,),), np.ndarray[np.int16])
    # print(f"r13={r13}, dtype={r13.dtype}")
    # r14 = ExcelArrayConverter(((0, 1,),), np.ndarray[np.int32])
    # print(f"r14={r14}, dtype={r14.dtype}")
    # r15 = ExcelArrayConverter(((0, 1,),), np.ndarray[np.int64])
    # print(f"r15={r15}, dtype={r15.dtype}")
    i = 0
    results = []
    val = ((1.5, 2.5,),)

    # t = ndarray1d
    # r = ExcelArrayConverter(val, t)
    # print(f"t={t} r{i}={r}, dtype={getattr(r, 'dtype', 'N/A')}")
    # results.append(r)
    # i += 1
    
    # t = ndarray2d
    # r = ExcelArrayConverter(val, t)
    # print(f"t={t} r{i}={r}, dtype={getattr(r, 'dtype', 'N/A')}")
    # results.append(r)
    # i += 1
    
    # t = ndarray1d[np.int32]
    # r = ExcelArrayConverter(val, t)
    # print(f"t={t} r{i}={r}, dtype={getattr(r, 'dtype', 'N/A')}")
    # results.append(r)
    # i += 1
    
    # t = ndarray2d[np.int32]
    # r = ExcelArrayConverter(val, t)
    # print(f"t={t} r{i}={r}, dtype={getattr(r, 'dtype', 'N/A')}")
    # results.append(r)
    # i += 1
    
    # t = list1d
    # r = ExcelArrayConverter(val, t)
    # print(f"t={t} r{i}={r}, dtype={getattr(r, 'dtype', 'N/A')}")
    # results.append(r)
    # i += 1
    
    # t = list2d
    # r = ExcelArrayConverter(val, t)
    # print(f"t={t} r{i}={r}, dtype={getattr(r, 'dtype', 'N/A')}")
    # results.append(r)
    # i += 1
    
    # t = list1d[int]
    # r = ExcelArrayConverter(val, t)
    # print(f"t={t} r{i}={r}, dtype={getattr(r, 'dtype', 'N/A')}")
    # results.append(r)
    # i += 1
    
    # t = list2d[int]
    # r = ExcelArrayConverter(val, t)
    # print(f"t={t} r{i}={r}, dtype={getattr(r, 'dtype', 'N/A')}")
    # results.append(r)
    # i += 1
    
    t = np.ndarray[np.int32]
    r = ExcelArrayConverter(val, t)
    print(f"t={t} r{i}={r}, dtype={getattr(r, 'dtype', 'N/A')}")
    results.append(r)
    i += 1

    t = list[int]
    r = ExcelArrayConverter(val, t)
    print(f"t={t} r{i}={r}, dtype={getattr(r, 'dtype', 'N/A')}")
    results.append(r)
    i += 1

    pass


pass

