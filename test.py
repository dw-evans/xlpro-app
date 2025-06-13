import pandas as pd
import win32com.client as win32

def replace_table_with_df(df: pd.DataFrame, table_name: str = "Table1"):
    # Connect to Excel
    excel = win32.GetActiveObject("Excel.Application")
    wb = excel.ActiveWorkbook

    # Locate the table
    found = False
    for ws in wb.Worksheets:
        for tbl in ws.ListObjects:
            if tbl.Name == table_name:
                table = tbl
                sheet = ws
                found = True
                break
        if found:
            break

    if not found:
        raise ValueError(f"Table '{table_name}' not found.")

    # Get starting cell
    top_left = table.Range.Cells(1, 1)

    # Get shape of DataFrame
    n_rows, n_cols = df.shape
    if n_rows == 0 or n_cols == 0:
        raise ValueError("DataFrame is empty or has no columns.")

    # Resize table range BEFORE writing anything
    new_range = sheet.Range(top_left, top_left.Cells(n_rows+1, n_cols))  # +1 row for header
    table.Resize(new_range)

    # Write headers
    header_range = sheet.Range(top_left, top_left.Cells(1, n_cols))
    header_range.Value = [df.columns.tolist()]

    # Write data
    data_start = top_left.Cells(2, 1)
    data_end = data_start.Cells(n_rows, n_cols)
    data_range = sheet.Range(data_start, data_end)
    data_range.Value = tuple(df.itertuples(index=False, name=None))

    print(f"✅ Table '{table_name}' updated with {n_rows} rows and {n_cols} columns.")



import pandas as pd

df = pd.DataFrame({
    "Name": ["Alice", "Bob", "Charlie", "Daniel", "Dervla"],
    "Age": [25, 30, 35, 13, 23],
    "Favourite Number": [1, 1, 3, 3, 4],
})

replace_table_with_df(df, table_name="Table1")

pass