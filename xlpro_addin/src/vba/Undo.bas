Attribute VB_Name = "Undo"

' Module-level variables
Dim WorkbookHandlers As Collection
Dim AppHandler As New CWorkbookEvents

Sub InitEventHandlers()
    Set WorkbookHandlers = New Collection
    Set AppHandler.App = Application

    Dim wb As Workbook
    For Each wb In Application.Workbooks
        AttachToWorkbook wb
    Next wb
End Sub

Sub AttachToWorkbook(wb As Workbook)
    Dim handler As New CWorkbookEvents
    Set handler.Wb = wb
    WorkbookHandlers.Add handler
End Sub
