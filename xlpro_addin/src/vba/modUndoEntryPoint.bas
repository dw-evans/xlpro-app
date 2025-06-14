Attribute VB_Name = "modUndoEntryPoint"

Option Explicit

Public GlobalUndoManager As UndoManager

Sub SetupUndo()
    Set GlobalUndoManager = New UndoManager
End Sub

Sub PerformUndo()
    If Not GlobalUndoManager Is Nothing Then
        GlobalUndoManager.RevertChange
    Else
        MsgBox "Undo manager not available.", vbExclamation
    End If
End Sub

Sub PerformRedo()
    If Not GlobalUndoManager Is Nothing Then
        GlobalUndoManager.ImplementChange
    Else
        MsgBox "Redo manager not available.", vbExclamation
    End If
End Sub


Public Sub GlobalUndoManager_RevertChange()
    GlobalUndoManager.RevertChange
End Sub

Public Sub GlobalUndoManager_ImplementChange()
    GlobalUndoManager.ImplementChange
End Sub

Sub setcustomundo()
    Application.OnUndo "xlpro Custom Undo", "GlobalUndoManager_RevertChange"
End Sub


Public Sub AtomicFormulaRefreshNoEvents(WbName As String, WsName As String, Address As String)

    Dim Wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range

    Set Wb = GetWorkbook(WbName)
    Set ws = GetWorksheet(Wb, WsName)

    Set rng = ws.Range(Address)

    ' Turn OFF all Excel events (no events will fire)
    Dim appeventsbefore As Boolean
    appeventsbefore = Application.EnableEvents

    On Error GoTo Cleanup
    Application.EnableEvents = False

    rng.Formula2 = rng.Formula2

    Application.OnUndo "xlpro Custom Undo", "GlobalUndoManager_RevertChange"

Cleanup:
        Application.EnableEvents = appeventsbefore

End Sub
