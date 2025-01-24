from typing import TypeVar
import numpy as np

T = TypeVar('T') # arbitrary

# this preserves all of the methods for type hinting and code completion for the user
# 1d options will force attempt reduction to a 1d vector and except if the input dimensionality is wrong.
list1d = list[T]
list2d = list[list1d[T]]

ndarray1d = np.ndarray[T]
ndarray2d = np.ndarray[T, T]



import regex as re
from dataclasses import dataclass, field
from win32com.client import GetActiveObject, Dispatch
import pythoncom
from win32typelibs import excel as xl
from pathlib import Path

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
        # XXX - todo - must improve deferred calculation here!

        # from utils import comarshal_release_and_get_stream, comarshal_dispatch_stream
        pythoncom.CoInitialize()
        # wb:xl._Workbook = comarshal_dispatch_stream(wb_stream)
        # Dip in and out of python com to retrieve the data. Then release.
        # XXX - todo - Check if getactiveobject is actually good here.
        # xlapp = GetActiveObject("Excel.Application")
        xlapp = Dispatch("Excel.Application")
        # XXX - todo - Check if this wb path matching check is reliable for server locations for example.
        # e.g. mapped drives may convert to server addresses. I believe resolve() corrects for this...
        if not Path(self.wb_path).resolve().__str__() in [Path(wb.FullName).resolve().__str__() for wb in xlapp.Workbooks]:
            xlapp.Workbooks.Open(self.wb_path)
        ret = xlapp.Workbooks(str(Path(self.wb_path).name)).Sheets(self.ws_name).Range(self.rng_addr).Value
        xlapp = None
        # comarshal_release_and_get_stream(wb)
        pythoncom.CoUninitialize()
        return ret

    
    # if any array argument arrives as a string, a pre-process step should be used
    # to evaluate the range.