Option Explicit

' Declare a module-level variable to hold the instance of the event handler class
Dim ehandler As xlproEventHandler

Private Sub Workbook_Open()
    ResetEventHandler
End Sub

Sub ResetEventHandler()
    Set ehandler = New xlproEventHandler
    Debug.Print "xlpro Event handler initialized for the add-in."
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Not Cancel Then
        Set ehandler = Nothing
        Debug.Print "xlpro Event handler cleaned up for the add-in."
    Else
        Debug.Print "Close canceled, event handler still active."
    End If
End Sub


Sub EnableEvents()
    Application.EnableEvents = True
End Sub









