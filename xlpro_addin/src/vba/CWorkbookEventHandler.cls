VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWorkbookEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' Class Module: CWorkbookEventHandler
Option Explicit

Public WithEvents Wb As Workbook
' Public WithEvents Ws As Worksheet

Private UndoStack As Collection

Private formulaCache as Variant
Private rngAddressCache as String

Private Sub Wb_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Call undo tracking logic here
    Debug.Print "Change in " & Wb.Name & " on " & Sh.Name & ": " & Target.Address

    ' if the sheet changes
    If rngAddressCache <> Target.Address(False, False) Then
        Debug.Print "Warning: Cached range mismatch. Undo may be invalid."
        Exit Sub
        ' Err.Raise vbObjectError + 1000, "CWorkbookEventHandler", "Range address does not match cached address!" ' & "Workbook:"& Wb.Name & " Sheet:" & Sh.name
    End If

    ' Store the cached previous value against the new value
    Call GlobalUndoManager.TrackChange(Wb, Sh, Target, formulaCache, Target.Formula2)

End Sub

Private Sub Wb_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    ' Pre-change tracking (cache old values)
    Debug.Print "Selection in " & Wb.Name & " on " & Sh.Name & ": " & Target.Address

    ' Track the state of the target range
    Set formulaCache = Target.Formula2
    rngAddressCache = Target.Address(False, False)
End Sub



