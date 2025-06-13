VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UndoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private EnumPointer As Long
Private UndoStack As Collection

' Enum for types of change
' Public Enum ChangeCategory
'     ChangeFormula = 1
' End Enum

Private Sub Class_Initialize()
    Set UndoStack = New Collection
    EnumPointer = 0
End Sub

' Track a change
Public Sub TrackChange( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal targetRange As Range, _
    ByVal val0 As Variant, _
    ByVal val1 As Variant)

    ' ByVal changeType As ChangeCategory,

    Dim item As clsUndoItem
    Set item = New clsUndoItem
    item.WorkbookName = wb.Name
    item.WorksheetName = ws.Name
    item.RangeAddress = targetRange.Address
    ' item.ChangeType = changeType
    item.Val0 = val0
    item.Val1 = val1

    ' Trim stack if we undid some changes and are now making a new one
    If EnumPointer < UndoStack.Count Then
        Do While UndoStack.Count > EnumPointer
            UndoStack.Remove UndoStack.Count
        Loop
    End If

    UndoStack.Add item
    EnumPointer = UndoStack.Count

    Debug.Print "Tracked change at " & item.WorkbookName & "/" & item.WorksheetName & "/" & item.RangeAddress
End Sub

' Undo the most recent change
Public Sub RevertChange()
    If EnumPointer = 0 Then
        Debug.Print "No changes to undo."
        Exit Sub
    End If

    Dim item As clsUndoItem
    Set item = New clsUndoItem

    Set item = UndoStack(EnumPointer)

    Dim wb As Workbook, ws As Worksheet
    Set wb = GetWorkbook(item.WorkbookName)
    Set ws = wb.Worksheets(item.WorksheetName)

    ws.Range(item.RangeAddress).Value = item.Val0

    EnumPointer = EnumPointer - 1
    Debug.Print "Reverted: " & item.WorksheetName & " " & item.RangeAddress
End Sub

' Redo the next change
Public Sub ImplementChange()
    If EnumPointer >= UndoStack.Count Then
        Debug.Print "No changes to redo."
        Exit Sub
    End If

    EnumPointer = EnumPointer + 1
    Dim item As clsUndoItem
    Set item = New clsUndoItem

    Set item = UndoStack(EnumPointer)

    Dim wb As Workbook, ws As Worksheet
    Set wb = GetWorkbook(item.WorkbookName)
    Set ws = GetWorksheet(wb, item.WorksheetName)

    ws.Range(item.RangeAddress).Value = item.Val1

    Debug.Print "Redone: " & item.WorksheetName & " " & item.RangeAddress
End Sub

' Helper to safely get a workbook by name
Private Function GetWorkbook(name As String) As Workbook
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
Private Function GetWorksheet(ByRef wb as Workbook, name As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Sheets
        If ws.Name = name Then
            Set GetWorksheet = ws
            Exit Function
        End If
    Next ws
    Err.Raise vbObjectError + 1000, "UndoManager", "Worksheet not found in workbook: " & "Workbook:"& wb.Name & " Sheet:" & name
End Function
