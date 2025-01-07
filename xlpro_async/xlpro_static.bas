Sub register_workbook()
    Dim xlpro_async As Object
    Set xlpro_async = GetObject("new: {122BB48A-57EF-4775-A28C-3F71ED0D02A7}")
    xlpro_async.register_and_configure_wb_workspace ThisWorkbook
End Sub

Sub sync_vba()
    Dim xlpro_async As Object
    Set xlpro_async = GetObject("new: {122BB48A-57EF-4775-A28C-3F71ED0D02A7}")
    xlpro_async.register_functions_in_vba ThisWorkbook
End Sub

Sub initialize()
    register_workbook
    sync_vba
End Sub

Sub getpid()
    Dim xlpro_async As Object
    Set xlpro_async = GetObject("new: {122BB48A-57EF-4775-A28C-3F71ED0D02A7}")
    Debug.Print xlpro_async.getpid
End Sub

Sub RegisterFunctionDescriptions()
    ' Register a description for the AddNumbers function
    MsgBox "registering something"
    Application.MacroOptions _
        Macro:="AddNumbers", _
        Description:="Adds two numbers and returns their sum.", _
        ArgumentDescriptions:=Array("First number to add", "Second number to add")
End Sub


