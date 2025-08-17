from win32com.client import Dispatch
import datetime
import pandas as pd
from pathlib import Path

import typing

if typing.TYPE_CHECKING:
    from win32typelibs import word as wd
    from win32typelibs import excel as xl

import xlpro
from xlpro import xlWorkbook

WD = Path(__file__).parent
OUTPUT_DIRECTORY: Path = WD.parent / "Invoices"
INVOICE_TEMPLATE_PATH: Path = WD.parent / "XLP_0000_INV_000_V01.dotx"
OUTPUT_DIRECTORY.mkdir(exist_ok=True)


@xlpro.register_sub
def show_folder():
    import subprocess
    subprocess.run(["explorer.exe", str(OUTPUT_DIRECTORY.resolve())])


def type_range(rng: "wd.Range", text: str):
    rng.Text = text
    return rng


def create_wd_instance():
    wdapp: "wd._Application" = Dispatch("Word.Application")
    wdapp.Visible = True
    return wdapp

@xlpro.register_sub
def generate_invoice(thiswb: "xlWorkbook"):

    sht = thiswb.ActiveSheet
    

    # Extract data from Excel
    client_code = sht.Range("Client_Code").Text
    project_number = sht.Range("Project_Number").Text
    invoice_number = sht.Range("Invoice_Number").Text
    author = sht.Range("Author").Text
    date = sht.Range("Date").Value
    revision = sht.Range("Revision").Text
    po_number = sht.Range("Purchase_Order").Text
    bill_to_name = sht.Range("Bill_Name").Text
    bill_to_company = sht.Range("Bill_Company").Text
    bill_to_email = sht.Range("Bill_Email").Text
    bill_to_addr1 = sht.Range("Bill_Address1").Text
    bill_to_addr2 = sht.Range("Bill_Address2").Text
    bill_to_addr3 = sht.Range("Bill_Address3").Text
    additional_info = sht.Range("Additional_Information").Text
    date_due = date + datetime.timedelta(30)
    invoice_table = sht.ListObjects("Invoice_Table")

    # Generate a document UID
    document_uid = f"{client_code}_{project_number}_INV_{invoice_number}_{revision}"

    # Load table data as dataframe and calculate sums
    df_invoice_table = pd.DataFrame(
        invoice_table.DataBodyRange.Value, columns=invoice_table.HeaderRowRange.Value[0]
    )
    calc_subtotal = df_invoice_table["Amount"].values.sum()
    calc_vat = calc_subtotal * 0.20
    calc_total = calc_subtotal + calc_vat

    # Instantiate Word and create the document
    wdapp = create_wd_instance()

    doc = wdapp.Documents.Add(Template=str(INVOICE_TEMPLATE_PATH), NewTemplate=False)
    export_pdf_path = OUTPUT_DIRECTORY / f"{document_uid}.pdf"

    def update_header():
        funcs_vals = [
            (doc.Bookmarks("AUTHOR").Range, author, "AUTHOR"),
            (
                doc.Bookmarks("DETAILS_BILLED_TO").Range,
                "\n".join(
                    [
                        bill_to_name,
                        bill_to_email,
                        bill_to_company,
                        bill_to_addr1,
                        bill_to_addr2,
                        bill_to_addr3,
                    ]
                ),
                "DETAILS_BILLED_TO",
            ),
            (doc.Bookmarks("INVOICE").Range, f"{invoice_number}_{revision}", "INVOICE"),
            (doc.Bookmarks("PO").Range, po_number, "PO"),
            (doc.Bookmarks("CLIENT").Range, client_code, "CLIENT"),
            (doc.Bookmarks("DATE").Range, date.strftime("%Y-%m-%d"), "DATE"),
            (
                doc.Bookmarks("DUE_DATE").Range,
                date_due.strftime("%Y-%m-%d"),
                "DUE_DATE",
            ),
            (
                doc.Bookmarks("ADDITIONAL_INFORMATION").Range,
                additional_info,
                "ADDITIONAL_INFORMATION",
            ),
            (doc.Bookmarks("DOCUMENT_UID").Range, document_uid, "DOCUMENT_UID"),
        ]
        for rng, val, bookmark in funcs_vals:
            new_rng = type_range(rng, val)
            if bookmark:
                doc.Bookmarks.Add(bookmark, new_rng)

    def update_invoice_table():
        invoice_table_wd: "wd.Table" = doc.Tables(3)
        funcs_vals = []

        for i, row in enumerate(df_invoice_table.values):
            for j, val in enumerate(row):
                if j in [3, 4]:
                    x = (
                        invoice_table_wd.Rows(2 + i).Cells(j + 1).Range,
                        f"£\t{val:.2f}",
                        None,
                    )
                elif j == 0:
                    x = (
                        invoice_table_wd.Rows(2 + i).Cells(j + 1).Range,
                        f"{int(val)}",
                        None,
                    )
                elif j == 2:
                    x = (
                        invoice_table_wd.Rows(2 + i).Cells(j + 1).Range,
                        f"{int(val)}",
                        None,
                    )
                else:
                    x = (invoice_table_wd.Rows(2 + i).Cells(j + 1).Range, val, None)
                funcs_vals.append(x)

        # No need to reinstate bookmark
        funcs_vals += [
            (doc.Bookmarks("SUBTOTAL").Range, f"£\t{calc_subtotal:.2f}", ""),
            (doc.Bookmarks("VAT").Range, f"£\t{calc_vat:.2f}", ""),
            (doc.Bookmarks("TOTAL").Range, f"£\t{calc_total:.2f}", ""),
        ]

        for rng, val, bookmark in funcs_vals:
            new_rng = type_range(rng, val)
            if bookmark:
                doc.Bookmarks.Add(bookmark, new_rng)

    update_header()
    update_invoice_table()

    doc.ExportAsFixedFormat2(
        OutputFileName=str(export_pdf_path),
        ExportFormat=17,  # wd.constants.wdExportFormatPDF,
    )

    doc.Close(SaveChanges=False)
