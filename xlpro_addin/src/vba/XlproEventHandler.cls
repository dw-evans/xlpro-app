VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XlproEventHandler"


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

    Debug.Print "Initializing XlproEventHandler"

    ' Attach to already open workbooks (doubt this affects anything if we auto-load xlpro.xlam)
    Dim Wb As Workbook
    For Each Wb In app.Workbooks
        Debug.Print "Attaching to " & Wb.Name & "..."
        AttachToWorkbook Wb
        Debug.Print "Attach successful."
    Next Wb

    Call SetupUndo
    Debug.Print "Global undo manager successfully initialized."
End Sub

Private Sub Class_Terminate()
    Set app = Nothing
    Set workbookHandlers = Nothing
End Sub

Public Sub AttachToWorkbook(Wb As Workbook)
    ' Attach the event handler to a workbook.
    Dim wbHandler As WorkbookEventHandler
    Set wbHandler = New WorkbookEventHandler
    Set wbHandler.Wb = Wb
    Debug.Print "Creating new xlpro workbook event handler for " & Wb.Name & "..."
    workbookHandlers.Add wbHandler, Wb.name
    Debug.Print "Workbook handler created."
End Sub


Private Sub app_WorkbookOpen(ByVal Wb As Workbook)
    ' Attach all newly opened workbooks to the handler
    AttachToWorkbook Wb
    Debug.Print "Workbook has been opened: " & Wb.name

End Sub

Private Sub app_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    ' Drop workbooks off the event handler if they are closed
    if Not Cancel Then
        On Error Resume Next
        workbookHandlers.Remove Wb.name
        On Error goto 0
        Debug.Print "Workbook has been closed: " & Wb.name
    End If  
End Sub


' Private Sub app_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)

'     On error goto ErrHandler
'     ' reset the event handler and embedded undo manager
'     if ehandler is Nothing Then
'         Debug.Print "Resetting the xlpro Event Handler as it didn't exist..."
'         Call ResetEventHandler
'         Debug.Print "xlpro Event Handler reset complete."
'     End If
'     On Error goto 0
'     Exit Sub

' ErrHandler:
'     Debug.Print "Error in App_SheetSelectionChange event handler."
'     Err.Clear
' End Sub







