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

Private workbookHandlers As Collection

Private Sub Class_Initialize()
    Set app = Application ' Initialize the app variable
    Set workbookHandlers = New Collection

    ' Attach to already open workbooks (doubt this affects anything if we auto-load xlpro.xlam)
    Dim Wb As Workbook
    For Each Wb In app.Workbooks
        AttachToWorkbook Wb
    Next Wb

    Call SetupUndo
    Debug.Print "Global undo manager initialized"
    
End Sub

Private Sub Class_Terminate()
    Set app = Nothing
    Set workbookHandlers = Nothing
End Sub

Public Sub AttachToWorkbook(Wb As Workbook)
    Dim wbHandler As CWorkbookEventHandler
    Set wbHandler = New CWorkbookEventHandler
    Set wbHandler.Wb = Wb
    workbookHandlers.Add wbHandler, Wb.name
End Sub


Private Sub app_WorkbookOpen(ByVal Wb As Workbook)
' Attach all newly opened workbooks to the
    AttachToWorkbook Wb
    Debug.Print "Workbook opened: " & Wb.name

End Sub

Private Sub app_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error Resume Next
    workbookHandlers.Remove Wb.name
    Debug.Print "Workbook closing: " & Wb.name
End Sub








