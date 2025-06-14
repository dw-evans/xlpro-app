VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ThisWorkbook Module of the Add-In

Option Explicit

' Declare a module-level variable to hold the instance of the event handler class
Dim ehandler As xlproEventHandler

Private Sub Workbook_Open()
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







