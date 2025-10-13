import xlpro

# This file is the 'entry point' of xlpro to define subroutines/macros in Excel.


@xlpro.register_sub
def write_hello_from_python(activews: "xlpro.xlWorksheet"):
    """Writes a string to cell A1 of the *active* worksheet."""
    activews.Range("A1").Value2 = "Hello from Python!"


@xlpro.register_sub
def create_new_sheet(activewb: "xlpro.xlWorkbook"):
    """Creates a new worksheet"""
    activewb.Sheets.Add2()
