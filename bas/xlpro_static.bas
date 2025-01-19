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


Sub write_vba_sync_module()

dim contents as string
dim comp_name as string

contents = get_vba_sync_text

write_text_to_module ThisWorkbook, "xlpro_async", contents

end sub


Function get_vba_sync_text() as string

    Dim xlpro_async As Object
    Set xlpro_async = GetObject("new: {122BB48A-57EF-4775-A28C-3F71ED0D02A7}")
    dim s as string
    s = xlpro_async.get_vba_sync_text

end function


Sub write_text_to_module(wb as workbook, c_name as string, contents as string)

dim proj as vbide.vbproject

set proj = wb.vbproject

Dim names() As String ' Dynamic array to store the names
Dim i As Long

' Resize the array to the number of items in the collection
ReDim names(1 To proj.vbcomponents.Count)

' Loop through the collection
For i = 1 To proj.vbcomponents.Count
    names(i) = col(i).Name ' Extract the .Name property of each element
Next i

dim comp as vbide.codemodule
if IsInArray(c_name, names) Then
    set comp = proj.vbcomponents.add(vbext_ct_StdModule)
    comp.Name = c_name
else:
    comp = proj.vbcomponents(c_name)
end if

dim codemod as vbide.codemodule

set codemod = comp.codemodule

codemod.deletelines(1, codemod.countoflines)
codemod.addfromstring(contents)

end sub


Function IsInArray(target As String, arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If element = target Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function
