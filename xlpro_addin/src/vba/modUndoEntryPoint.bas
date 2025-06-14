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

Sub ButtonPerformUndo(ByRef control As Office.IRibbonControl)
    PerformUndo
End Sub

Sub ButtonPerformRedo(ByRef control As Office.IRibbonControl)
    PerformRedo
End Sub


' Public Sub GlobalUndoManager_RevertChange()
'     GlobalUndoManager.RevertChange
' End Sub

' Public Sub GlobalUndoManager_ImplementChange()
'     GlobalUndoManager.ImplementChange
' End Sub

Sub SetCustomUndo()
    Application.OnUndo "XLPRO: UNDO UNAVAILABLE", "WarnUndo"
End Sub

Sub WarnUndo()
    MsgBox "Native undo stack is unavailable, use xlpro-undos", vbExclamation
End Sub


Public Sub AtomicFormulaRefreshNoEvents(rng As Range)

    ' Turn OFF all Excel events (no events will fire)
    Dim appeventsbefore As Boolean
    appeventsbefore = Application.EnableEvents

    On Error GoTo Cleanup
    Application.EnableEvents = False

    rng.Formula2 = rng.Formula2

    Call SetCustomUndo

Cleanup:
        Application.EnableEvents = appeventsbefore

End Sub

