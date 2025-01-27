import utils

from functools import wraps
import threading
from xlpro_enums import FunctionTypes
import xlpro_typing

# XXX - todo - maybe implement threading locks in future.
# Not sure when you'd ever have multithread during registration unless you were maybe mixing libraries?

# map of func names to functions (not used)
_module_func_name_register:dict[dict[str, str]] = {}
# map of func names to internal function types
_module_func_name_type_register:dict[dict[str, int]] = {}

_module_func_name_isjsonified_register:dict[dict[str, bool]] = {}

# map of func names to their active status
_module_func_name_isactive_register:dict[dict[str, bool]] = {}

class ModuleFunctionMapsWrapper:
    """Wrapper for module-specific function registry maps"""
    def __init__(self, mname):
        global _module_func_name_register
        global _module_func_name_type_register
        global _module_func_name_isjsonified_register
        global _module_func_name_isactive_register

        self.mname = mname
        self.func_name_register:dict = _module_func_name_register[mname]
        self.func_name_type_register:dict = _module_func_name_type_register[mname]
        self.func_name_isjsonified_register:dict = _module_func_name_isjsonified_register[mname]
        self.func_name_isactive_register:dict = _module_func_name_isactive_register[mname]
        return

def _remove_module_from_maps(mname):
    try:
        del _module_func_name_register[mname]
        del _module_func_name_type_register[mname]
        del _module_func_name_isactive_register[mname]
        del _module_func_name_isjsonified_register[mname]
    except:
        pass


def _add_module(mname):
    if not mname in _module_func_name_register:
        _module_func_name_register[mname] = {}
        _module_func_name_type_register[mname] = {}
        _module_func_name_isactive_register[mname] = {}
        _module_func_name_isjsonified_register[mname] = {}

def _func_set_active(func, state:bool): 
    if not isinstance(state, bool):
        raise TypeError
    fname = func.__name__
    mname = func.__module__
    _add_module(mname)
    _module_func_name_isactive_register[mname][fname] = state

def _func_set_jsonified(func, isjsonified:bool): 
    if not isinstance(isjsonified, bool):
        raise TypeError
    fname = func.__name__
    mname = func.__module__
    _add_module(mname)
    if isjsonified:
        pass
    _module_func_name_isjsonified_register[mname][fname] = isjsonified

def _register_func(func):
    global _module_func_name_register
    mname = func.__module__
    fname = func.__name__
    _add_module(mname)
    _module_func_name_register[mname][fname] = func
    pass

def _validate_func_type(_type):
    from server import FunctionTypes
    return _type in FunctionTypes.as_list()

def _register_func_type(func, _type):
    _validate_func_type(_type)
    mname = func.__module__
    fname = func.__name__
    _add_module(mname)
    _module_func_name_type_register[mname][fname] = _type


def _register(func, _type, _isactive):
    _register_func(func)
    _func_set_active(func, _isactive)
    if _type is None:
        _type = utils.infer_func_result_type_from_type_hints(func)
    _register_func_type(func, _type)

def register(_type:None|int=None, isactive=True):
    """Registers the function for xlpro. User can set function type or rely on PEP-484 type hints
    per the documentation"""
    def wrapper(func):
        # register the function only if its not already been registered.
        mname = func.__module__
        fname = func.__name__
        if not mname in _module_func_name_register:
            _add_module(mname)
        if not fname in _module_func_name_register[mname]:
            _register(func, _type, isactive)
        return func
    return wrapper

def ignore(_type:None|int=None):
    """Registers the function for xlpro with isactive=False so it does not enter as a UDF in excel.
    User can set function type or rely on PEP-484 type hints per the documentation"""
    def wrapper(func):
        _register(func, _type, False)
        return func
    return wrapper


def import_module_with_registration(mname, fpath):
    _remove_module_from_maps(mname)
    
    utils.import_module(mname, fpath)

    valid_functions = utils.get_udf_valid_functions_from_module(mname)
    for f in valid_functions:
        f.__module__ = mname
        register()(f)



import errors
import json


def _default_func_wrapper(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        f = func
        utils.validate_args_ready(args, kwargs)
        ppargs, ppkwargs = utils.preprocess_arguments(func=func, args=args, kwargs=kwargs)
        
        ret = f(*ppargs, **ppkwargs)
        ret_converted = xlpro_typing.ExcelArrayConverter._convert_back_to_range_format(ret)
        return ret_converted
    return wrapper

def _jsonified_func_wrapper(func):
    @wraps(func)
    def wrapper(s):
        utils.validate_args_ready((s,), {})
        kwargs = json.loads(s)
        return _default_func_wrapper(func)(**kwargs)
    return wrapper


def jsonify_func(globals_dict:dict):
    """Register a fork of the function in globals() with a single str argument"""
    def wrapper0(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            return func(*args, **kwargs)
        
        # register()(func)
        wrapper.__name__ = f"{func.__name__}_json"
        wrapper.__qualname__ = wrapper.__name__

        register()(wrapper)
        _func_set_jsonified(wrapper, True)

        if wrapper.__name__ in globals_dict:
            raise Exception("Name conflict encountered")
        globals_dict[wrapper.__name__] = wrapper

        return func
    return wrapper0


def generate_wrapped_function(mname, fname):
    # func = sys.modules[mname][fname]
    func = _module_func_name_register[mname][fname]
    ftype = _module_func_name_type_register[mname][fname]
    # isactive = __module_func_name_isactive_register[mname][fname]
    isjson = _module_func_name_isjsonified_register[mname].get(fname, False)

    if ftype == FunctionTypes.default:
        if isjson:
            return _jsonified_func_wrapper(func)
        return _default_func_wrapper(func)
    raise Exception("Not supported")