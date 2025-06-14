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
Private DoTrackChanges As Boolean
Private UndoHistoryLength As Long

' Enum for types of change
' Public Enum ChangeCategory
'     ChangeFormula = 1
' End Enum

Private Sub Class_Initialize()
    Set UndoStack = New Collection
    DoTrackChanges = True
    EnumPointer = 1
    UndoHistoryLength = 12
End Sub

' Track a change
Public Sub TrackChange( _
    ByVal Wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal TargetRange As Range, _
    ByVal val0 As Variant, _
    ByVal val1 As Variant)

    If Not DoTrackChanges Then
        Debug.Print "Tracking Disabled"
        Exit Sub
    End If


    ' ByVal changeType As ChangeCategory,

    Dim item As clsUndoItem
    Set item = New clsUndoItem
    item.WorkbookName = Wb.name
    item.WorksheetName = ws.name
    item.RangeAddress = TargetRange.Address
    ' item.ChangeType = changeType
    item.val0 = val0
    item.val1 = val1

    ' Trim stack if we undid some changes and are now making a new branch
    If EnumPointer < UndoStack.Count Then
        Do While UndoStack.Count > EnumPointer
            UndoStack.Remove UndoStack.Count
        Loop
    End If
    
    UndoStack.Add item
    
    ' If we exceed the maximum undo stack count by adding another
    If UndoStack.Count > UndoHistoryLength Then
        UndoStack.Remove 1
        Debug.Print "Undo Stack maximum length exceeded, dropping"
    End If

    ' Set the undo stack counter to one past the end of the collection
    EnumPointer = UndoStack.Count + 1

    Debug.Print "Tracked change at " & item.WorkbookName & "/" & item.WorksheetName & "/" & item.RangeAddress
End Sub

' Undo the most recent change
Public Sub RevertChange()
    If EnumPointer = 1 Then
        Debug.Print "No changes to undo."
        Exit Sub
    End If
    
    EnumPointer = EnumPointer - 1

    Dim item As clsUndoItem
    Set item = New clsUndoItem

    Set item = UndoStack(EnumPointer)

    Dim Wb As Workbook, ws As Worksheet
    Set Wb = GetWorkbook(item.WorkbookName)
    Set ws = Wb.Worksheets(item.WorksheetName)
    Dim TargetRange As Range
    
    Set TargetRange = ws.Range(item.RangeAddress)

    Dim DoTrackChangesBefore As Boolean
    DoTrackChangesBefore = DoTrackChanges
    DoTrackChanges = False
    TargetRange.Formula2 = item.val0
    DoTrackChanges = DoTrackChangesBefore
    
    TargetRange.Select
    
    Debug.Print "Reverted: " & item.WorksheetName & " " & item.RangeAddress
End Sub

' Redo the next change
Public Sub ImplementChange()
    If EnumPointer > UndoStack.Count Then
        Debug.Print "No changes to redo."
        Exit Sub
    End If
    


    Dim item As clsUndoItem
    Set item = New clsUndoItem

    Set item = UndoStack(EnumPointer)

    Dim Wb As Workbook, ws As Worksheet
    Set Wb = GetWorkbook(item.WorkbookName)
    Set ws = GetWorksheet(Wb, item.WorksheetName)
    Dim TargetRange As Range
    
    Set TargetRange = ws.Range(item.RangeAddress)

    Dim DoTrackChangesBefore As Boolean
    DoTrackChangesBefore = DoTrackChanges
    DoTrackChanges = False
    TargetRange.Formula2 = item.val1
    DoTrackChanges = DoTrackChangesBefore
    
    TargetRange.Select
    
    EnumPointer = EnumPointer + 1

    Debug.Print "Redone: " & item.WorksheetName & " " & item.RangeAddress
End Sub

