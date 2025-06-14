Attribute VB_Name = "modUndoEntryPoint"

Option Explicit

' Helper to safely get a workbook by name
Public Function GetWorkbook(name As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.Name = name Then
            Set GetWorkbook = wb
            Exit Function
        End If
    Next wb
    Err.Raise vbObjectError + 1000, "UndoManager", "Workbook not found: " & name
End Function

' Helper to safely get a workbook by name
Public Function GetWorksheet(ByRef wb as Workbook, name As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Sheets
        If ws.Name = name Then
            Set GetWorksheet = ws
            Exit Function
        End If
    Next ws
    Err.Raise vbObjectError + 1000, "UndoManager", "Worksheet not found in workbook: " & "Workbook:"& wb.Name & " Sheet:" & name
End Function
