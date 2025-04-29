Attribute VB_Name = "xlpro_static"
Option Explicit

Public Const XLPRO_GUID As String = "{122BB48A-57EF-4775-A28C-3F71ED0D02A7}"
Public Const XLPRO_CLI_PATH As String = "C:\Users\Daniel Evans\projects\xlpro\.venv\Scripts\xlpro-cli.exe"
Public Const VSCODE_PATH As String = "code"

'Public Const XLPRO_CLI_PATH As String = "xlpro-cli"

Public WORKBOOK_GUID_MAP As Scripting.Dictionary
Public WORKBOOK_GUID_MAP_INITIALIZED As Boolean

Public WSCRIPT_SHELL As Object
Public WSCRIPT_SHELL_INITIALIZED As Boolean


#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Public Sub initialize_workbook_guid_map()
    If Not WORKBOOK_GUID_MAP_INITIALIZED Then
        Set WORKBOOK_GUID_MAP = New Scripting.Dictionary
        WORKBOOK_GUID_MAP_INITIALIZED = True
    End If
    Debug.Print "WORKBOOK_GUID_MAP initialized"
End Sub
Public Sub uninitialize_workbook_guid_map()
    Set WORKBOOK_GUID_MAP = Nothing
    WORKBOOK_GUID_MAP_INITIALIZED = False
    Debug.Print "WORKBOOK_GUID_MAP uninitialized"
End Sub


Public Sub InitializeShell()
    ' Only initialize the shell object once
    If Not WSCRIPT_SHELL_INITIALIZED Then
        Set WSCRIPT_SHELL = CreateObject("WScript.Shell")
        WSCRIPT_SHELL_INITIALIZED = True ' Set flag to True after initialization
    End If
    Debug.Print "WSCRIPT_SHELL initialized"
End Sub
Sub UnintializeShell()
    Set WSCRIPT_SHELL = Nothing
    WSCRIPT_SHELL_INITIALIZED = False ' Set flag to True after initialization
    Debug.Print "WSCRIPT_SHELL uninitialized"
End Sub

Sub set_workbook_guid_map_pairing(wb_name As String, guid As String)
    initialize_workbook_guid_map
    WORKBOOK_GUID_MAP.Add wb_name, guid
End Sub
Sub del_workbook_guid_map_key(wb_name As String)
    initialize_workbook_guid_map
    If WORKBOOK_GUID_MAP.Exists(wb_name) Then
        WORKBOOK_GUID_MAP.Remove wb_name
    End If
End Sub
Public Function get_workbook_guid_map_value(wb_name As String) As String
    initialize_workbook_guid_map
    If WORKBOOK_GUID_MAP.Exists(wb_name) Then
        get_workbook_guid_map_value = WORKBOOK_GUID_MAP(wb_name)
    Else
        get_workbook_guid_map_value = ""
    End If
End Function

'Public Const XLPRO_GUID_NAME As String = "XLPRO_GUID"
'Sub update_workbook_guid_value(wb As Workbook)
'    ' stores the workbook guid as a name in the workbook
'    Dim guid As String
'    guid = get_workbook_guid_map_value(wb.Name)
'    ' Add the XLPRO_GUID constant to the workbook names
'    wb.names.Add Name:=XLPRO_GUID_NAME, RefersTo:="="" & guid & """
'End Sub
'
'Function get_workbook_guid_value(wb As Workbook) As String
'    ' Fetches the guid from the workbook name
'    get_workbook_guid_value = Evaluate(wb.names(XLPRO_GUID_NAME).RefersTo)
'End Function


'------------------------------------------------------------------------
' UI ribbon elements
'------------------------------------------------------------------------
Sub xlproStart(ByRef control As Office.IRibbonControl)
    Dim taskID As Double
    Dim command As String
    
    Dim wb As Workbook
    
    Set wb = ActiveWorkbook

    ' Use the Shell function to call the program
    command = """" & XLPRO_CLI_PATH & """" & " start " & """" & wb.Path & "\" & wb.Name & """"
    Debug.Print command
    taskID = shell("cmd /c " & """" & command & """", vbNormalFocus)
    
    ' clear the workbook key if it exists
    initialize_workbook_guid_map
    del_workbook_guid_map_key wb.Name


    ' Optionally, display the Task ID of the program
    Debug.Print "Program launched with Task ID: " & taskID


End Sub

Sub xlproInit(ByRef control As Office.IRibbonControl)
    Dim taskID As Double
    Dim command As String

    ' Use the Shell function to call the program
    'taskID = Shell("cmd.exe /K xlpro", vbNormalFocus)
    command = """" & XLPRO_CLI_PATH & """" & " init " & """" & ActiveWorkbook.Path & "\" & ActiveWorkbook.Name & """"
    Debug.Print command
    taskID = shell("cmd /c " & """" & command & """", vbNormalFocus)

    ' Optionally, display the Task ID of the program
    Debug.Print "Program launched with Task ID: " & taskID

End Sub

Sub xlproRegister(ByRef control As Office.IRibbonControl)
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Debug.Print ActiveWorkbook.Path
    
    On Error GoTo RegistrationErrorHandler
    register_workbook ActiveWorkbook
    On Error GoTo 0
    
    write_vba_sync_module ActiveWorkbook
    Exit Sub
    
RegistrationErrorHandler:
    MsgBox "Error: Error during registration. Please connect the debugger and retry.", _
           vbExclamation, "Warning"
    Err.Clear
    Exit Sub
End Sub


Sub xlproStartIDE(ByRef control As Office.IRibbonControl)
    
    Dim command As String
    Dim taskID As Double
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    command = """" & VSCODE_PATH & """" & " " & """" & wb.Path & "\" & ActiveWorkbook.Name & ".xlpro" & """"
    Debug.Print command
    taskID = shell(command, vbNormalFocus)
End Sub

'------------------------------------------------------------------------
'Items below here are helper subroutines for the addin.
'------------------------------------------------------------------------

Function RunCommandAndCaptureOutput(command As String) As String
    'Dim shell As Object
    Dim tempFile As String
    Dim fNum As Integer
    Dim output As String
    Dim guid As String
    Dim timestamp As String

    ' Generate unique file name using a GUID
    'guid = Replace(CreateObject("Scriptlet.TypeLib").guid, "{", "")
    'guid = Replace(guid, "}", "")
    'tempFile = Environ$("TEMP") & "\" & "out_" & guid & ".txt"

    timestamp = Format(Now, "yyyymmdd_HHMMSS") & "_" & Right(Format(Timer, "0.000"), 3)
    tempFile = Environ$("TEMP") & "\xlpro_" & timestamp & ".txt"

    ' Redirect command output to the temp file
    command = "cmd /c """ & command & " 2> """ & tempFile & """ 1> nul"""
    'command = "cmd /c """ & command & " > """ & tempFile & """ 2>&1"""
    Debug.Print command
    ' Execute the command
    'Set shell = CreateObject("WScript.Shell")
    
    'InitializeShell
    WSCRIPT_SHELL.Run command, 0, True  ' 0=hidden window, True=wait for completion

    ' Read the output from the temp file
    fNum = FreeFile
    On Error GoTo Cleanup
    'Debug.Print tempFile
    Open tempFile For Input As #fNum
    output = Input$(LOF(fNum), fNum)
    Close #fNum

Cleanup:
    On Error Resume Next
    Kill tempFile  ' Delete the temp file
    RunCommandAndCaptureOutput = output
End Function

Function get_workbook_guid(wb As Workbook) As String
    Dim command As String
    Dim shell As Object
    Dim exec As Object
    Dim output As String
    Dim line As String

    On Error GoTo ErrHandler


'    Set shell = CreateObject("WScript.Shell")
'
    command = """" & XLPRO_CLI_PATH & """" & " guid " & """" & ActiveWorkbook.Path & "\" & ActiveWorkbook.Name & """"
    ' command = """" & XLPRO_CLI_PATH & """ guid """ & ActiveWorkbook.Path & "\" & ActiveWorkbook.Name & """"

    Debug.Print command
    output = RunCommandAndCaptureOutput(command)

    get_workbook_guid = output
    Debug.Print output
    Exit Function

ErrHandler:
    get_workbook_guid = "Error: " & Err.Description
End Function


Sub register_activeworkbook()
    register_workbook ActiveWorkbook
End Sub

Sub unregister_activeworkbook()
    uninitialize ActiveWorkbook
End Sub

Private Sub register_workbook(ByRef wb As Workbook)
    On Error GoTo 0
    initialize_workbook_guid_map
    InitializeShell
    Dim guid As String
    guid = get_workbook_guid_map_value(wb.Name)
    If guid = "" Then
        guid = get_workbook_guid(wb)
        set_workbook_guid_map_pairing wb.Name, guid
    End If

    Debug.Print "XLPRO_GUID: " & guid

    Dim xlpro_async As Object
    Set xlpro_async = GetObject("new: " & guid)

    xlpro_async.register_and_configure_wb_workspace wb

End Sub


Sub uninitialize(ByRef wb As Workbook)
'Uninitialize this workbook from the com server
    Dim xlpro_async As Object
    Dim guid As String
    guid = get_workbook_guid_map_value(wb.Name)
    Set xlpro_async = GetObject("new: " & guid)
    xlpro_async.shutdown_workspace wb
End Sub

Private Sub shutdown_xlpro(ByRef wb As Workbook)
'Attempt to shutdown the xlpro server.
    Dim xlpro_async As Object
    Dim guid As String
    guid = get_workbook_guid_map_value(wb.Name)
    Set xlpro_async = GetObject("new: " & guid)
    xlpro_async.shutdown
End Sub

Private Sub getpid(ByRef wb As Workbook)
    Dim xlpro_async As Object
    Dim guid As String
    guid = get_workbook_guid_map_value(wb.Name)
    Set xlpro_async = GetObject("new: " & guid)
    Debug.Print xlpro_async.getpid
End Sub

Private Sub shutdown_workspace(ByRef wb As Workbook)
    Dim xlpro_async As Object
    Dim guid As String
    guid = get_workbook_guid_map_value(wb.Name)
    Set xlpro_async = GetObject("new: " & guid)
    xlpro_async.shutdown_workspace wb
End Sub


' Replacement synchronization functions
Sub write_vba_sync_module(ByRef wb As Workbook)
    Dim contents As String
    Dim comp_name As String
    contents = get_vba_sync_text(wb)
    write_text_to_module wb, "xlpro_async", contents
End Sub

Function get_vba_sync_text(wb As Workbook) As String
    Dim xlpro_async As Object
    Dim guid As String
    guid = get_workbook_guid_map_value(wb.Name)
    Set xlpro_async = GetObject("new: " & guid)
    get_vba_sync_text = xlpro_async.get_vba_sync_text(wb)
    
'SyncTextErrorHandler:
'    MsgBox "Error: Error during VBA Module sync. Please connect the debugger and retry.", _
'           vbExclamation, "Warning"
'    Err.Clear
'    Exit Function
End Function

Sub write_text_to_module(ByRef wb As Workbook, c_name As String, contents As String)
    Dim proj As VBIDE.VBProject
    Set proj = wb.VBProject
    Dim names() As String ' Dynamic array to store the names
    Dim i As Long
    
    ' Resize the array to the number of items in the collection
    ReDim names(1 To proj.vbcomponents.Count)
    
    ' Loop through the collection
    For i = 1 To proj.vbcomponents.Count
        names(i) = proj.vbcomponents(i).Name ' Extract the .Name property of each element
    Next i
    
    Dim comp As VBIDE.VBComponent
    If Not IsInArray(c_name, names) Then
        Set comp = proj.vbcomponents.Add(vbext_ct_StdModule)
        comp.Name = c_name
    Else:
        Set comp = proj.vbcomponents(c_name)
    End If
    
    Dim codemod As VBIDE.codemodule
    
    Set codemod = comp.codemodule
    
    codemod.DeleteLines 1, codemod.CountOfLines
    codemod.addfromstring contents
End Sub


'------------------------------------------------------------------------
'Helper functions etc
'------------------------------------------------------------------------

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

Private Sub RegisterFunctionDescriptions()
    ' Register a description for the AddNumbers function
    MsgBox "registering something"
    Application.MacroOptions _
        Macro:="AddNumbers", _
        Description:="Adds two numbers and returns their sum.", _
        ArgumentDescriptions:=Array("First number to add", "Second number to add")
End Sub


Public Function ptr(rng As Range)
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set ws = rng.Parent
    Set wb = ws.Parent
    
    ptr = "*<" & wb.FullName & "::" & ws.Name & "::" & rng.Address & ">"
End Function


Sub test(val As Long, val2 As Long)
    ActiveSheet.Range("A1").Value = val + val2
End Sub

Sub test2(val As Variant)
    Dim i0 As Long
    i0 = LBound(val)
    ActiveSheet.Range("A1").Value = val(i0) + val(i0 + 1)
End Sub

Sub test3(val As Variant)
    Dim i0 As Long
    i0 = LBound(val)
    ActiveSheet.Range("A1").Value = val(i0) & val(i0 + 1)
End Sub

Sub conditional_format_handler(workbook_name As String, sheet_name As String, range_names As Variant, colors As Variant, return_uid As String)
    'pass the areas and conditional formatting colours for each area
    Dim Workbook As Workbook
    Dim Worksheet As Worksheet
    
    Dim area As Range
    Dim rng As Range
    Dim i As Long
    
    If Not (LBound(range_names) = LBound(colors) And UBound(colors) = UBound(range_names)) Then
        Err.Raise 9999, "xlproError", "Array bounds do not match"
    End If
    
    Set Workbook = Workbooks(workbook_name)
    Set Worksheet = Workbook.Sheets(sheet_name)

    Dim cell_color As Long

    ' create the range (which could have multi-areas)
    ' loop over the areas and set the interior color to the target color.
    For i = LBound(range_names) To UBound(range_names)
        Set rng = Worksheet.Range(range_names(i))
        cell_color = colors(i)
        For Each area In rng.areas
            area.Interior.Color = cell_color
        Next area
    Next i
    
    Set rng = Nothing
    Set Workbook = Nothing
    Set Worksheet = Nothing
    Set area = Nothing
    
    'Dim xlpro As Object
    'Set xlpro = GetObject("new: " & XLPRO_GUID)
    'xlpro.set_return_value(uid, "Success('" & uid & "')"
    
End Sub



Function conditional_formatter(formula_str As String, rng As Range, root_cell As Range) As Variant

    Dim rng_name_arr As Variant
    Dim rng_vals As Variant
    
    ' ReDim rng_name_arr(1 To rng.Count)
    Dim i As Integer
    
    Dim rng_name As String
    Dim workbook_name As String
    Dim sheet_name As String
    Dim root_cell_addr As String

    
    workbook_name = rng.Parent.Parent.Name
    sheet_name = rng.Parent.Name
    area_names_arr = get_range_name_per_area(rng)
    root_cell_addr = root_cell.Address
    
    rng_vals = get_range_values(Join(area_names_arr, ","))
    
    Dim xlpro As Object
    Set xlpro = GetObject("new: " & XLPRO_GUID)
    conditional_formatter = xlpro.execute_function_async(ActiveWorkbook, Application.Caller, "conditional_formatter_example", workbook_name, sheet_name, rng_name_arr, formula_str, root_cell_addr)
End Function

Private Function get_range_name_per_area(rng As Range) As Variant()

Dim cell As Range
Dim values() As Variant
Dim i As Long
Dim area As Range

' rng.Interior.Color = 16744319
ReDim values(1 To rng.Cells.Count)

i = 1
For Each area In rng.areas
    values(i) = area.Address
    i = i + 1
Next area

get_range_name_per_area = values

End Function

Private Function get_range_name_per_cell(rng As Range) As Variant()

Dim cell As Range
Dim values() As Variant
Dim i As Long
Dim area As Range

' rng.Interior.Color = 16744319
ReDim values(1 To rng.Cells.Count)

i = 1
For Each area In rng.areas
    For Each cell In area.Cells
        values(i) = cell.Address
        i = i + 1
    Next cell
Next area

get_range_name_per_cell = values

End Function

Private Function get_range_values(rng_name As String) As Variant()

Dim rng As Range
Dim cell As Range
Dim values() As Variant
Dim i As Long
Dim area As Range

Set rng = Range(rng_name)

rng.Interior.Color = 16744319
ReDim values(1 To rng.Cells.Count)

i = 1
For Each area In rng.areas
    For Each cell In area.Cells
        values(i) = cell.Value
        i = i + 1
    Next cell
Next area

get_range_values = values

End Function



