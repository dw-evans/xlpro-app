Option Explicit

' Declare a module-level variable to hold the instance of the event handler class
Dim ehandler As XlproEventHandler

Sub ResetEventHandler()
    Set ehandler = New XlproEventHandler
    Debug.Print "xlpro Event handler initialized for the add-in."
    Call initialize_workbook_guid_map
End Sub

Private Sub Workbook_Open()
    ResetEventHandler
End Sub

Public Property Get EventHandler() As Object
    Set EventHandler = ehandler
End Property

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







