import typing
from typing import TypeVar, Generic, Any
import numpy as np
from xlpro import errors

T = TypeVar('T') # arbitrary

# this preserves all of the methods for type hinting and code completion for the user
# 1d options will force attempt reduction to a 1d vector and except if the input dimensionality is wrong.
list1d = list[T]
list2d = list[list[T]]

ndarray1d = np.ndarray[T]
ndarray2d = np.ndarray[T, T]

import regex as re
from dataclasses import dataclass, field
from win32com.client import GetActiveObject, Dispatch
import pythoncom
# from win32typelibs import excel as xl
from pathlib import Path
import pandas as pd

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

def compare_generic_aliases(t1, t2):
    if isinstance(t1, typing.GenericAlias) and isinstance(t2, typing.GenericAlias):
        checks = []
        checks.append(t1.__origin__ == t2.__origin__)
        for st1, st2 in zip(t1.__args__, t2.__args__):
            if isinstance(st1, typing.TypeVar) and isinstance(st2, typing.TypeVar):
                checks.append(True)
                continue
            if isinstance(st1, typing.GenericAlias) and isinstance(st1, typing.GenericAlias):
                checks.append(compare_generic_aliases(st1, st2))
                continue
            checks.append(st1 == st2)
        return all(checks)
    t_generic = None
    t_other = None
    if isinstance(t1, typing.GenericAlias):
        t_generic = t1
        t_other = t2
    elif isinstance(t2, typing.GenericAlias):
        t_generic = t2
        t_other = t1
    else:
        raise TypeError("Incompatible types being checked")
    return t_generic.__origin__ == t_other

from typing import Iterable
@dataclass
class xlproImage:
    fp:Path
    size_pt:Iterable[float]
    xl_name:str

# XXX - todo - apparently this is sensitive to imports...
# type checking broke when I refactored, presumably changed the origin of some of the objects?
class ExcelArrayConverter:
    """Excel will convert ranges to 2d tuple arrays. Use this class to convert
    2d arrays into a user-specified type before the user's code receives it.
    
    The user should only require up to 2d data with float, int, str and maybe
    boolean types so we can implement this custom logic confidently.
    """
    def __new__(cls, val:Any, tdst:type):

        # excel will provide a range as tuple[tuple]
        # or a 1x1 range as a value.

        if tdst == list1d[str]:
            pass

        # return the incoming value if not a 2d array needing conversion
        if not isinstance(val, (tuple, list)):
            return val

        # dissect the destination type
        origin = typing.get_origin(tdst)
        args = typing.get_args(tdst)

        # The user can specify a type if they want - e.g. np.float64, np.int32, np.bool
        # extract the final or intermediate dtype to use.
        dtype = cls._convert_alias_args_to_np_dtype(args)

        if compare_generic_aliases(tdst, list1d):
            pass
            
        # map the input value to list1d or array1d
        if any([compare_generic_aliases(tdst, x) for x in [list1d, ndarray1d]]):
            intermediate = np.array(val)
            shape = intermediate.shape
            # implicitly the shape is 2d, convert it to the user chosen list or ndarray
            if all([x > 1 for x in shape]):
                raise TypeError("Provided value is not compatible with list1d")
            if compare_generic_aliases(tdst, list1d):
                return intermediate.flatten().tolist()
            elif compare_generic_aliases(tdst, ndarray1d):
                return intermediate.flatten()
        
        # as above
        elif any([compare_generic_aliases(tdst, x) for x in [list2d, ndarray2d]]):
            intermediate = np.array(val)
            if compare_generic_aliases(tdst, list2d):
                return intermediate.tolist()
            return intermediate

        # handle generic lists and arrays
        elif tdst in [list, np.ndarray]:
            intermediate = np.array(val)
            if tdst == list:
                return intermediate.tolist()
            return intermediate
        
        # elif tdst in [tuple,]:
        #     # XXX - todo - for some reason Excel doesn't accept a tuple of tuples as a return type. 
        #     # Needs to be a numpy array. maybe because it needs to be contigious memory?
        #     return np.array(val)

        # handle things such as list[float], i.e. <origin>[<dtype>]
        else:
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
    def _convert_alias_args_to_np_dtype(cls, args:tuple[type]):
        """Converts the args of a generic alias into a single dtype for numpy conversion
        A return of None will implicitly resort to numpy's conversion handling."""

        # XXX - todo - we may want to support strings in a fancy way...
        # For string arrays we probably want to fetch the .Text property of 
        # a range and not the .Value property (specifically concerned with 
        # 0 != "" (blank cells))
        # XXX - todo - this is a job for the vba string generator!

        if len(args) == 0:
            return None
        
        a0 = args[0]

        # list[float, int, ...] using different types is not supported
        if not all([x == a0 for x in args]):
            raise TypeError(f"Type arguments {a0} are too complicated to convert")
        
        
        # if the type has not been given for the list1d or list2d case return no type
        if a0 in [T, list[T]]:
            return None
        
        # recursively handle generic alias
        if isinstance(a0, typing.GenericAlias):
            if not typing.get_origin(a0) == list:
                raise TypeError(f"Type {a0} cannot be processed, origin must be list")
            return cls._convert_alias_args_to_np_dtype(typing.get_args(a0))
        
        # convert python types to numpy types for convenience.
        elif a0 == float:
            return np.float64
        elif a0 == int:
            return np.int64
        elif a0 == bool:
            return np.bool
        elif a0 == str:
            return str
        
        # unsure what this is for.
        elif isinstance(a0, TypeVar):
            return None
        
        # if the user specifies a specific dtype, resort to that
        # numpy should catch any conversion errors when attempting to convert
        elif any([np.issubdtype(a0, x) for x in (np.floating, np.integer, np.bool)]):
            return a0
        
        
        else:
            # Let the numpy default behaviour run
            # XXX - todo - check how this behaves
            return None 


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
    """Class to signal that an array is to be expanded."""
    def __init__(self, arraydata):
        oned_direction_rowwise = False
        if oned_direction_rowwise:
            if len((npdata:=np.array(arraydata)).shape) == 1:
                self.data = npdata.reshape(-1, 1)
        else:
            self.data = arraydata



if __name__ == "__main__":
    r1  = ExcelArrayConverter(((1, 2,),), list1d)
    print(f"r1={r1}")
    r2  = ExcelArrayConverter(((1, 2,),), list2d)
    print(f"r2={r2}")
    r5  = ExcelArrayConverter(((1, 2,),), list)
    print(f"r5={r5}")
    r6  = ExcelArrayConverter(((1, 2,),), list[float])
    print(f"r6={r6}")
    r7  = ExcelArrayConverter(((1, 2,),), list[int])
    print(f"r7={r7}")
    r8  = ExcelArrayConverter(((0, 1,),), list[bool])
    print(f"r8={r8}")
    r3  = ExcelArrayConverter(((1, 2,),), ndarray1d)
    print(f"r3={r3}, dtype={r3.dtype}")
    r4  = ExcelArrayConverter(((1, 2,),), ndarray2d)
    print(f"r4={r4}, dtype={r4.dtype}")
    r10 = ExcelArrayConverter(((0, 1,),), np.ndarray[np.float16])
    print(f"r10={r10}, dtype={r10.dtype}")
    r11 = ExcelArrayConverter(((0, 1,),), np.ndarray[np.float32])
    print(f"r11={r11}, dtype={r11.dtype}")
    r12 = ExcelArrayConverter(((0, 1,),), np.ndarray[np.float64])
    print(f"r12={r12}, dtype={r12.dtype}")
    r13 = ExcelArrayConverter(((0, 1,),), np.ndarray[np.int16])
    print(f"r13={r13}, dtype={r13.dtype}")
    r14 = ExcelArrayConverter(((0, 1,),), np.ndarray[np.int32])
    print(f"r14={r14}, dtype={r14.dtype}")
    r15 = ExcelArrayConverter(((0, 1,),), np.ndarray[np.int64])
    print(f"r15={r15}, dtype={r15.dtype}")
    pass

pass

