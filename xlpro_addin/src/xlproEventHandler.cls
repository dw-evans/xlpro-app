VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xlproEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents app As Application
Attribute app.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set app = Application ' Initialize the app variable
    
    'initialize_workbook_guid_map
    'InitializeShell
    
End Sub

Private Sub Class_Terminate()
    Set app = Nothing ' Clean up
    
    'uninitialize_workbook_guid_map
    'UnintializeShell
    
End Sub

Private Sub app_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    'uninitialize the workbook on close
    uninitialize ActiveWorkbook
End Sub

' Overwrite the application onclose behaviour
Private Sub app_Workbook_BeforeClose(Cancel As Boolean)
    ' Only clean up the event handler if the close is not canceled
    If Not Cancel Then
        'uninitialize the workbook on close
        Debug.Print "Attempting to uninitialize workbook..."
        uninitialize ActiveWorkbook
        Debug.Print "Uninitialize completed (maybe?)"
    Else
        Debug.Print "Close canceled, event handler still active."
    End If
End Sub

' TODO - xxx - prevent any overwrites of an existing workbookbeforeclose
' function for example.



