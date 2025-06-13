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