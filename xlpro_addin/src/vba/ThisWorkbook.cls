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

' Workbook_Open event: Called when the add-in is loaded
Private Sub Workbook_Open()
    ' Initialize the event handler class when the add-in is loaded
    Set ehandler = New xlproEventHandler
    Debug.Print "Event handler initialized for the add-in."
End Sub

' Workbook_BeforeClose event: Called when the add-in is closed
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Only clean up the event handler if the close is not canceled
    If Not Cancel Then
        Set ehandler = Nothing
        Debug.Print "Event handler cleaned up for the add-in."
    Else
        Debug.Print "Close canceled, event handler still active."
    End If
End Sub

