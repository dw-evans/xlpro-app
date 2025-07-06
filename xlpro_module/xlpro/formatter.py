from win32com.client.dynamic import Dispatch
import random
from xlpro import errors
from pathlib import Path
import pandas as pd
import numpy as np
from typing import Any
import pythoncom
from win32com.client import VARIANT
import re

def binned_color_func(a, b):
    def wrapper():
        val = random.randint(a, b)
        return val
    return wrapper


def alpha_to_int(s):
    result = 0
    for char in s:
        result = result * 26 + (ord(char.lower()) - ord('a'))
    return result

def int_to_base26_alpha(n):
    if n == 0:
        return 'a'
    chars = []
    while n > 0:
        chars.append(chr(ord('a') + n % 26))
        n //= 26
    return ''.join(reversed(chars))

def split_xl_address_aadd(addr:str):
    mtch = re.match(r"^([a-z]+)(\d+)$", addr)
    if not mtch:
        raise ValueError("provided string is incompatible.")
    return mtch.groups()

import ast
import operator

# Supported operators
operators = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
    ast.Pow: operator.pow,
    ast.Lt: operator.lt,
    ast.Gt: operator.gt,
    ast.LtE: operator.le,
    ast.GtE: operator.ge,
    ast.Eq: operator.eq,
    ast.NotEq: operator.ne,
    ast.USub: operator.neg,
}

def eval_expr(expr):
    def _eval(node):
        if isinstance(node, ast.Expression):
            return _eval(node.body)
        elif isinstance(node, ast.BinOp):
            return operators[type(node.op)](_eval(node.left), _eval(node.right))
        elif isinstance(node, ast.UnaryOp):
            return operators[type(node.op)](_eval(node.operand))
        elif isinstance(node, ast.Compare):
            left = _eval(node.left)
            for op, comparator in zip(node.ops, node.comparators):
                if not operators[type(op)](left, _eval(comparator)):
                    return False
                left = _eval(comparator)
            return True
        # elif isinstance(node, ast.Num):  # for Python < 3.8
        #     return node.n
        elif isinstance(node, ast.Constant):  # for Python 3.8+
            return node.value
        else:
            raise ValueError(f"Unsupported expression: {ast.dump(node)}")
    tree = ast.parse(expr, mode='eval')
    return _eval(tree)


def rgb_to_hex(rgb) -> str:
    """Convert an (R, G, B) tuple to a hex string."""
    return "#{:02x}{:02x}{:02x}".format(*rgb)

def hex_to_rgb(hex_str) -> np.ndarray:
    """Convert a hex string to an (R, G, B) tuple. Accepts optional leading '#'."""
    hex_str = hex_str.lstrip('#')
    if len(hex_str) != 6:
        raise ValueError("Hex string must be 6 characters long.")
    return np.array(tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))


def interpolate_rgb(start_rgb, end_rgb, bins):
    """Interpolate between two RGB tuples over `bins` steps (inclusive)."""
    interpolated = np.linspace(start_rgb, end_rgb, bins)
    return [tuple(map(int, color)) for color in interpolated]

def rgb_to_xl_rep(rgb:tuple[int, int, int]) -> int:
    r, g, b = rgb
    return b + (g << 8) + (r << 16)

def convert_area_address_to_individual_addrs_col_major(area_addr:str) -> list[str]:
    """$a$1:$b$10 -> a1, b1, a2, b2, a3, b3, ..., a10, b1"""
    area_addr = area_addr.replace("$", "")
    area_addr1, area_addr2 = area_addr.split(":")

    r_val1, c_val1 = get_coordinates_from_cell_address(area_addr1)
    r_val2, c_val2 = get_coordinates_from_cell_address(area_addr2)
    
    cell_coords = np.array([(x, y) for x in range(r_val1, r_val2+1, 1) for y in range(c_val1, c_val2+1, 1)])
    cell_addrs = [f"{int_to_base26_alpha(y)}{x+1}" for (x, y) in cell_coords]

    return cell_addrs

def get_coordinates_from_cell_address(rng_str:str):
    r1, c1 = split_xl_address_aadd(rng_str)
    r_val1 = int(c1) - 1
    c_val1 = alpha_to_int(r1)
    return (r_val1, c_val1)


def conditional_formatter_example(workbook_name:str, sheet_name:str, area_names_arr:tuple[str], range_vals:tuple[Any], formula_str:str, root_cell_addr:str):
    """Colors cells which satisfy a formula"""
    import pythoncom
    pythoncom.CoInitialize()

    color_low = "ffffff"
    low_lim = 0.0
    color_high = "ff0000"
    high_lim = 1000.0
    n_bins = 50
    bins = np.linspace(low_lim, high_lim, n_bins+1)
    rgb_low = hex_to_rgb(color_low)
    rgb_high = hex_to_rgb(color_high)
    rgb_bins = interpolate_rgb(rgb_low, rgb_high, n_bins)
    xl_rgb_bins = [rgb_to_xl_rep(x) for x in rgb_bins]

    # try:
    pythoncom.CoInitialize()
    xlapp = Dispatch("Excel.Application")
    try:
        xlapp.Workbooks
    except AttributeError:
        raise errors.ExcelNotAccessibleError
    
    range_name = ",".join(area_names_arr)
    rng = xlapp.Workbooks(workbook_name).Sheets(sheet_name).Range(range_name)
    
    # create the addrs using the areas.
    addrs = []
    for area_name in area_names_arr:
        addrs += convert_area_address_to_individual_addrs_col_major(area_name)

    # calculate the coordinates of each cell address and calculate the relative positions.
    addr_coords = np.zeros(shape=(len(addrs), 2), dtype=np.int64)
    for i, addr in enumerate(addrs):
        addr_coords[i] = get_coordinates_from_cell_address(addr)

    root_cell_coords = np.array(get_coordinates_from_cell_address(root_cell_addr))
    rel_addr_coords =  addr_coords - root_cell_coords 

    relrows = rel_addr_coords[:,0]
    relcols = rel_addr_coords[:,1]

    # range_names_stripped = range_name.replace("$", "").split(",")

    data = {
        "addr": addrs,
        "relrow": relrows,
        "relcol": relcols,
        "val": range_vals,
    }

    df = pd.DataFrame(data)

    # formula_str = formula_str.replace(" ", "")

    pattern = r"r(\$?-?\d+)c(\$?-?\d+)"
    formula_args = list(re.finditer(pattern, formula_str, re.IGNORECASE))
    formula_template = formula_str
    for i, x in enumerate(formula_args):
        formula_template = re.sub(pattern, f"{{arg{i}}}", formula_template, count=1)

    n_args = len(formula_args)
    for i, x in enumerate(formula_args):

        (a1, a2) = x.groups()
        a1_is_frozen = a1.startswith("$")
        a2_is_frozen = a2.startswith("$")

        if a1_is_frozen:
            a1_val = int(a1[1:])
        else:
            a1_val = int(a1)

        if a2_is_frozen:
            a2_val = int(a2[1:])
        else:
            a2_val = int(a2)

        df[f"arg{i}r"] = df["relrow"] - a1_val if not a1_is_frozen else a1_val
        df[f"arg{i}c"] = df["recol"] - a2_val if not a2_is_frozen else a2_val

        df[f"arg{i}"] = None
        for y in df.index:
            r = df.loc[y]
            df.loc[y, f"arg{i}"] = (
                df[(df["relrow"] == r[f"arg{i}r"]) & (df["relcol"] == r[f"arg{i}c"])]["val"].values[0]
            )

    df["eval_str"] = None
    for i in df.index:
        df.loc[i, "eval_str"] = formula_template.format(**{(k:=f"arg{j}"): str(df.loc[i, k]) for j in range(n_args)})

    df["eval"] = None
    for i in df.index:
        df.loc[i, "eval"] = eval_expr(df.loc[i, "eval_str"])

    df["bin"] = None
    for i in df.index:
        df.loc[i, "bin"] = np.abs(bins - df.loc[i, "eval"]).argmin()

    pass

    df[["addr", "bin"]]

    bin_ranges = {} # {bin_idx: range_addrs}
    for i in range(n_bins):
        tmp = df[df["bin"] == i]["addr"]
        if tmp.empty:
            continue
        bin_ranges[i] = ",".join(tmp.values)

    # large calls here will be inefficient so reworking this as a sub call instead.
    # try:
    #     for bin_idx, addr in bin_ranges.items():
    #         rng = ... # worksheet.range(addr) ... 
    #         rng.Interior.Color = xl_rgb_bins[bin_idx]
    # except Exception as e:
    #     raise errors.ExcelNotAccessibleError()

    range_names = []
    colors = []
    max_areas = 25
    for k, v in bin_ranges.items():
        str_split = v.split(",")
        i = 0
        while True:
            range_names.append(",".join(str_split[i * max_areas: (v1:=min(len(str_split), (i+1) * max_areas))]))
            colors.append(xl_rgb_bins[k])
            if v1 == len(str_split):
                break
            i += 1


    # xlapp.ActiveSheet.Range("d15,e15,d36,e36,d79,e79,d108,e108,d109,e109,d191,e191,d195,e195,d227,e227,d228,e228,d266,e266,d294,e294,d323,e323,d333,e333,d341,e341,d387,e387,d396,e396,d407,e407,d477,e477,d480,e480,d496,e496,d537,e537,d579,e579,d659,e659,d762,e762,d789,e789,d889")
    xlapp.Application.Run(
        "conditional_format_handler", 
        workbook_name, 
        sheet_name, 
        range_names, 
        colors, 
        "",
    )

    pass
    # xlapp.Application.Run("test", 10, 30)
    # xlapp.Application.Run("test2", (10,31))
    # xlapp.Application.Run("test2", VARIANT(pythoncom.VT_BYREF | pythoncom.VT_ARRAY, [10,30]))
    # xlapp.Application.Run("test3", ["hello", " there"])


    xlapp = None
    # comarshal_release_and_get_stream(wb)
    pythoncom.CoUninitialize()
    return "ConditionalFormatter('Success')"


    # except Exception as e:
    #     pythoncom.CoUninitialize()
    #     raise e


if __name__ == "__main__":
    xlapp = Dispatch("Excel.Application")
    rng = xlapp.Workbooks("main.xlsx").Sheets("sheet3").Range((addr:="d10:e10000"))
    vals = np.array(rng.Value).flatten()
    pass
    conditional_formatter_example(
        workbook_name="main.xlsx",
        sheet_name="sheet3",
        area_names_arr=(addr,),
        range_vals=vals,
        formula_str="r0c$0",
        root_cell_addr="e10",
    )

"""

=conditional_formatter_example(affected_range, root_cell, formula)

# Cell highlighting/colouring
formula = "r0c-1 > 10.0"
formula = "r0$c-5 == 100.0"

# Colour binning
formula = "r0$c-5" (color (low_lim, low_color, high_lim, high_color))
formula = "r0$c-5" (color (low_lim, low_color, mid_lim, mid_color, high_lim, high_color), bins_low, bins_high)


convert affected_range to single range object
get values from single range object and map to addresses

# highlighter formula
f = f(formula_str)
f = f("r0$c-10 > 10")

# gradient formula
# mono-directional
f = f(formula_str, color_low, low_lim, color_high, high_lim, bins)
f = f(formula_str, #ffffff, 0, #ff0000, 10, bins=10)
f = f(formula_str, #ffffff, "min", #ff0000, "max", bins=10)

# bi-directional
f = f(formula_str, color_low, low_lim, color_mid, mid_lim, color_high, high_lim, bins_high, bins_low)
f = f(formula_str, #0000ff, 0, #ffffff, 5, #ff0000, 10, bins_high=4, bins_low=10)
f = f(formula_str, #0000ff, "min", #ffffff, "median", #ff0000, "max", bins_high=4, bins_low=10)

# get the formula arguments from the dataframe
args = re.findall(r"((?:r\$?-?\d+)(?:c\$?-?\d+))", formula_str, re.IGNORECASE)
extract strings: argname1: r$-1c$10 ($-1) ($10)

if startswith $
    c$10 maps to (10 - rel_col) i.e. (col-10 = 0, col-0 = 10)
    use this to calculate argname1 for each element in the table
    repeat for argname2, argname3

e.g. for argname1 == $r0$c-1
pd_df[argname1] = pd_df[
    (pd_df["rel_row"] == 0) and 
    (pd_df["rel_col"] == -1)
]
e.g. for argname2 == r$0c10
pd_df[argname1] = pd_df[
    pd_df["rel_row"] == 0) and 
    pd_df["rel_col"] == rel_col + 10)
]
# warning we need to access data outside of the range!
pd_df[argname3] = pd_df[...]

f = f(argname1, argname2, argname3)
pd_df["val"] = vectorize(f)(pd_df[argname1 || argname2 || argname3].values)

pd_df[pd_df["val"] == 0]["addr"] -> torange() -> Range.Color.Fill = #ff0000
pd_df[pd_df["val"] == 1] ...
pd_df[pd_df["val"] == -1] ...
pd_df[pd_df["val"] == 2] ...
...
pd_df[pd_df["val"] == n_bins] ...

f_map: {f"{rel_row};{rel_col}": partial(f, rel_col=rel_col, rel_row=rel_row)}

data structure
pd.Dataframe -> addr(str), relrow(int), relcol(int), val(int|float|str), result(bool|int)

result = f_map[f"{rel_row};{rel_col}"](val)

filter df by result


lambda $a1 > 10.0 applies to a1:d10

$r1$c-1 -> evaluate $r1$c-1 relative to root cell. apply to all
$r1c-1 -> evaluate $r1 relative to root cell for all. calculate c-1 for each cell(column)
r1$c-1 -> evaluate $c-1 relative to root cell for all. calculate r1 for each cell(row)

optimization
cells can be 

ux
using (0,0) as the basis may not be natural to excel users but suits going negative well.



"""