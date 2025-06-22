from __future__ import annotations
from xlpro import _utils
from functools import wraps
import threading
from xlpro._enums import FunctionTypes
from xlpro import _types
from xlpro import errors
import json
import typing

# XXX - todo - maybe implement threading locks in future.
# Not sure when you'd ever have multithread during registration unless you were maybe mixing libraries?

# from typing import TYPE_CHECKING
# if TYPE_CHECKING:

from xlpro._utils import FSig
from xlpro import _utils


# map of func names to functions (not used)
_module_fname_func_register:dict = {}
_module_func_fname_register:dict = {}
# map of func names to internal function types
_module_fname_type_register:dict[dict[str, int]] = {}

_module_fname_isjsonified_register:dict[dict[str, bool]] = {}

# map of func names to their active status
_module_fname_isactive_register:dict[dict[str, bool]] = {}
_module_fname_function_signature_register:dict[dict[str, FSig]] = {}


_module_subname_func_register:dict = {}
_module_func_subname_register:dict = {}
_module_subname_isactive_register:dict[dict[str, bool]] = {}


class ModuleSubMapsWrapper:
    """Wrapper for module-specific function registry maps"""
    def __init__(self, mname):
        self.mname = mname
        self.subname_func_register:dict = _module_subname_func_register[mname]
        self.func_subname_register:dict = _module_func_subname_register[mname]
        self.subname_isactive_register:dict = _module_subname_isactive_register[mname]
        return

def _remove_subs_module_from_maps(mname):
    try:
        del _module_fname_func_register[mname]
        del _module_subname_isactive_register[mname]
    except:
        pass

def _add_sub_module(mname):
    if not mname in _module_subname_func_register:
        _module_subname_func_register[mname] = {}
        _module_func_subname_register[mname] = {}
        _module_subname_isactive_register[mname] = {}
        pass
    ...

def _sub_set_active(func, mname, state:bool):
    _add_sub_module(mname)
    # fname = get_registered_func_name(func, mname)
    fname = func.__name__
    _module_subname_isactive_register[mname][fname] = state

def _register_sub(func, _mname, _isactive):
    _add_sub_module(_mname)
    # fname = get_registered_func_name(func, _mname)
    fname = func.__name__
    _module_subname_func_register[_mname][fname] = func
    _module_func_subname_register[_mname][func] = fname
    _sub_set_active(func, _mname, _isactive)


def register_sub(isactive=True):
    """Primary interface for registration"""
    mname = _utils.get_caller_globals(inspect.currentframe())["__name__"]
    def wrapper(func):
        # fname = get_registered_func_name(func, mname)
        fname = func.__name__
        if not mname in _module_subname_func_register:
            _add_sub_module(mname)
        if not fname in _module_subname_func_register[mname]:
            _register_sub(func, mname, isactive)
        return func
    return wrapper


def sub_ignore(func):
    mname = _utils.get_caller_globals(inspect.currentframe())["__name__"]
    def wrapper(func):
        # _register(func, mname, _type, False)
        return func
    return wrapper
    ...



class ModuleFunctionMapsWrapper:
    """Wrapper for module-specific function registry maps"""
    def __init__(self, mname):
        self.mname = mname
        self.fname_func_register:dict = _module_fname_func_register[mname]
        self.func_fname_register:dict = _module_func_fname_register[mname]
        self.fname_type_register:dict = _module_fname_type_register[mname]
        self.fname_type_register_isjsonified_register:dict = _module_fname_isjsonified_register[mname]
        self.fname_isactive_register:dict = _module_fname_isactive_register[mname]
        self.fname_function_signature_register:dict = _module_fname_function_signature_register[mname]
        return

def _remove_funcs_module_from_maps(mname):
    try:
        del _module_fname_func_register[mname]
        del _module_func_fname_register[mname]
        del _module_fname_type_register[mname]
        del _module_fname_isactive_register[mname]
        del _module_fname_isjsonified_register[mname]
        del _module_fname_function_signature_register[mname]
    except:
        pass


def _add_func_module(mname):
    if not mname in _module_fname_func_register:
        _module_fname_func_register[mname] = {}
        _module_func_fname_register[mname] = {}
        _module_fname_type_register[mname] = {}
        _module_fname_isactive_register[mname] = {}
        _module_fname_isjsonified_register[mname] = {}
        _module_fname_function_signature_register[mname] = {}

def _func_set_active(func, mname, state:bool): 
    if not isinstance(state, bool):
        raise TypeError
    fname = get_registered_func_name(func, mname)
    _module_fname_isactive_register[mname][fname] = state

def _func_set_jsonified(func, mname, isjsonified:bool): 
    if not isinstance(isjsonified, bool):
        raise TypeError
    fname = get_registered_func_name(func, mname)
    if isjsonified:
        pass
    _module_fname_isjsonified_register[mname][fname] = isjsonified


def _register_func(func, mname):
    fname = get_registered_func_name(func, mname)
    _module_fname_func_register[mname][fname] = func
    _module_func_fname_register[mname][func] = fname


def _validate_func_type(_type):
    from xlpro.server import FunctionTypes
    return _type in FunctionTypes.as_list()

def _register_func_type(func, mname, _type):
    _validate_func_type(_type)
    fname = get_registered_func_name(func, mname)
    _module_fname_type_register[mname][fname] = _type

def _register_fname(func, mname, fname):
    if func in _module_func_fname_register[mname]:
        raise Exception("Attempted to register function object again")
    if fname in _module_fname_func_register[mname]:
        raise Exception("Attempted to register function name again")
    _module_func_fname_register[mname][func] = fname
    _module_fname_func_register[mname][fname] = func

def get_registered_func_name(func, mname):
    return _module_func_fname_register[mname][func]

def _register(func, _mname, _type, _isactive, fname:str=None):
    fname = fname if fname is not None else func.__name__
    _register_fname(func, _mname, fname)
    _register_func(func, _mname)
    _func_set_active(func, _mname, _isactive)
    if _type is None:
        _type = _utils.infer_func_result_type_from_type_hints(func)
    _register_func_type(func, _mname, _type)

def register(_type:None|int=None, isactive=True, fname:str=None):
    """Registers the function for xlpro. User can set function type or rely on PEP-484 type hints
    per the documentation"""
    mname = _utils.get_caller_globals(inspect.currentframe())["__name__"]
    def wrapper(func):
        _add_func_module(mname)
        _register(func=func, _mname=mname, _type=_type, _isactive=isactive, fname=fname)
        return func
    return wrapper


# def expand(func):
#     mname = _utils.get_caller_globals(inspect.currentframe())["__name__"]
#     def inner(*args, **kwargs):
#         return _utils.expand(func(*args, **kwargs))
#     return inner





# XXX - todo - register 'register' and 'ignore' as class based methods to give the option of
# calling with parentheses or not.
# Actually don't know if this is a horrendous idea.

def ignore(_type:None|int=None):
    """Registers the function for xlpro with isactive=False so it does not enter as a UDF in excel.
    User can set function type or rely on PEP-484 type hints per the documentation"""
    mname = _utils.get_caller_globals(inspect.currentframe())["__name__"]
    def wrapper(func):
        _register(func, mname, _type, False)
        return func
    return wrapper


def import_module_with_registration(mname, fpath):
    _remove_funcs_module_from_maps(mname)
    
    _utils.import_module(mname, fpath)

    valid_functions = _utils.get_udf_valid_functions_from_module(mname)
    for f in valid_functions:
        _register(f, mname, None, True)


def import_module_subs_with_registration(mname, fpath):
    _remove_subs_module_from_maps(mname)

    _utils.import_module(mname, fpath)

    valid_functions = _utils.get_sub_valid_functions_from_module(mname)
    for f in valid_functions:
        _register_sub(f, mname, _isactive=True)


def _pyobj_func_wrapper(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        f = func
        _utils.pre_validate_args(args, kwargs)
        ppargs, ppkwargs = _utils.preprocess_arguments(func=func, args=args, kwargs=kwargs)

        if func.__name__ == "mpl_set_xlims":
            pass

        # XXX - todo - could add multilevel nesting to this...
        all_args = ppargs + [v for v in ppkwargs.values()]
        for a in all_args:
            if isinstance(a, typing.Iterable):
                _utils.pre_validate_args(args=a, kwargs={}) 
        
        ret = f(*ppargs, **ppkwargs)
        return ret
    
    return wrapper


def _array_or_value_func_wrapper(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        ret = _pyobj_func_wrapper(func)(*args, **kwargs)
        ret_converted = _types.ExcelArrayConverter._convert_back_to_range_format(ret)
        return ret_converted
    return wrapper


def _jsonified_pyobj_func_wrapper(func):
    @wraps(func)
    def wrapper(s):
        _utils.pre_validate_args((s,), {})
        kwargs = json.loads(s)
        return _pyobj_func_wrapper(func)(**kwargs)
    return wrapper


def _jsonified_array_or_value_func_wrapper(func):
    @wraps(func)
    def wrapper(s):
        _utils.pre_validate_args((s,), {})
        kwargs = json.loads(s)
        return _array_or_value_func_wrapper(func)(**kwargs)
    return wrapper


import inspect
def wrap_jsonify():
    """Register a fork of the function in globals() with a single str argument"""
    # globals_dict = _utils.get_caller_globals()
    globals_dict = _utils.get_caller_globals(inspect.currentframe())
    mname = globals_dict["__name__"]
    def wrapper0(func):

        @wraps(func)
        def wrapper(*args, **kwargs):
            return func(*args, **kwargs)
        
        # register()(func)
        wrapper.__name__ = f"{get_registered_func_name(func, mname)}_json"
        wrapper.__qualname__ = wrapper.__name__

        _register(wrapper, mname, None, True)
        _func_set_jsonified(wrapper, mname, True)

        if wrapper.__name__ in globals_dict:
            raise Exception("Name conflict encountered")
        
        globals_dict[wrapper.__name__] = wrapper

        return func
    return wrapper0

import pythoncom

def _com_init_dispatch_release_wrapper(func):
    """Wraps com object dispatch and release around a func.
    Also appropriately configures pythoncom coinitialise"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        pythoncom.CoInitialize()
        # dispatch the args on this thread
        # only relevant if the reserved dispatch arguments are being used.
        try:
            args_dispatched = _utils.com_args_dispatch_reserved(func, args)
        except Exception as e:
            raise e
        ret = func(*args_dispatched, **kwargs)

        # must release after!
        _utils.com_args_release_to_stream_reserved(func, args_dispatched)
        pythoncom.CoUninitialize()
        
        return ret
    return wrapper



def generate_wrapped_function(mname, fname):
    """primary interface for generating wrapped functions which pre-parse excel arguments."""
    # func = sys.modules[mname][fname]
    func = _module_fname_func_register[mname][fname]
    ftype = _module_fname_type_register[mname][fname]
    # isactive = __module_func_name_isactive_register[mname][fname]
    isjson = _module_fname_isjsonified_register[mname].get(fname, False)

    retf:callable = None

    if ftype == FunctionTypes.py_object:
        if isjson:
            retf = _jsonified_pyobj_func_wrapper(func)
        else:
            retf = _pyobj_func_wrapper(func)
    elif ftype == FunctionTypes.array_or_value:
        if isjson:
            retf = _jsonified_array_or_value_func_wrapper(func)
        else:
            retf = _array_or_value_func_wrapper(func)

    if retf:
        return _com_init_dispatch_release_wrapper(retf)
    
    raise NotImplementedError("Function type is not supported")






# def generate_wrapped_function_sub(mname, fname):
#     """primary interface for generating wrapped functions which pre-parse excel arguments."""
#     # func = sys.modules[mname][fname]
#     func = _module_fname_func_register[mname][fname]
#     return func
