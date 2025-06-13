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

Private workbookHandlers As Collection

Private Sub Class_Initialize()
    Set app = Application ' Initialize the app variable
    Set workbookHandlers = New Collection

    ' Attach to already open workbooks (doubt this affects anything if we auto-load xlpro.xlam)
    Dim wb As Workbook
    For Each wb In app.Workbooks
        AttachToWorkbook wb
    Next wb

    Call SetupUndo
    Debug.Print "Global undo manager initialized"
    
End Sub

Private Sub Class_Terminate()
    Set app = Nothing
    Set workbookHandlers = Nothing
End Sub

Public Sub AttachToWorkbook(wb As Workbook)
    Dim wbHandler As CWorkbookEventHandler
    Set wbHandler = New CWorkbookEventHandler
    Set wbHandler.Wb = wb
    workbookHandlers.Add wbHandler, wb.Name
End Sub


Private Sub app_WorkbookOpen(ByVal Wb As Workbook)
' Attach all newly opened workbooks to the 
    AttachToWorkbook Wb
    Debug.Print "Workbook opened: " & Wb.Name

End Sub

Private Sub app_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error Resume Next
    workbookHandlers.Remove Wb.Name
    Debug.Print "Workbook closing: " & Wb.Name
End Sub




