
from win32typelibs import excel as xl
import json


def serialize_range(obj:xl.Range):
    raise NotImplementedError
    ws:xl._Worksheet = obj.Parent
    wb = ws.Parent
    return f"{obj.Parent.Parent}{obj.Parent}"


