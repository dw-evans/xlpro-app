Attribute VB_Name = "xlpro_static"

' ### BEGIN METADATA ###
' --- xlpro_static.bas ---
' Compiled with xlpro\xlpro_addin\main.py
' At 2025-06-09, 21:48:58
' ### END METADATA ###


' ### BEGIN METADATA ###
' --- xlpro_static.bas ---
' Compiled with xlpro\xlpro_addin\main.py
' At 2025-06-08, 23:34:04
' ### END METADATA ###


' ### BEGIN METADATA ###
' --- xlpro_static.bas ---
' Compiled with xlpro\xlpro_addin\main.py
' At 2025-06-08, 23:30:21
' ### END METADATA ###



Option Explicit

Public XLPRO_CLI_PATH As String
Public VSCODE_PATH As String
' Public XLPRO_ADDIN_PATH As String

Public WORKBOOK_GUID_MAP As Object
Public WORKBOOK_GUID_MAP_INITIALIZED As Boolean

Public WSCRIPT_SHELL As Object
Public WSCRIPT_SHELL_INITIALIZED As Boolean


public const MAX_ARGS_READY_CHECKED as Integer = 32

Public rePromise As Object
Public reException As Object


#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Sub LoadXlproConfigTOML()
    Dim fso As Object
    Dim file As Object
    Dim fileText As String
    Dim userProfilePath As String
    Dim configPath As String

    ' Get the USERPROFILE environment variable
    userProfilePath = Environ("USERPROFILE")

    ' Construct the full path to config.toml (adjust the relative part as needed)
    configPath = userProfilePath & "\.xlpro\config.toml"

    ' Create FileSystemObject (late binding)
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check if file exists
    If Not fso.FileExists(configPath) Then
        MsgBox "Config file not found: " & configPath, vbExclamation
        Exit Sub
    End If

    ' Open and read file
    Set file = fso.OpenTextFile(configPath, 1) ' 1 = ForReading
    fileText = file.ReadAll
    file.Close
    Debug.Print fileText

    ' Use RegExp to extract a and b
    Dim re As Object
    Dim matches As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False

    ' Find a
    re.Pattern = "(?:^|\n)\s*xlpro_cli_path\s*=\s*""([^""]+)"""
    Set matches = re.Execute(fileText)
    If matches.Count > 0 Then
        XLPRO_CLI_PATH = matches(0).SubMatches(0)
    Else:
        GoTo RegistrationErrorHandler
    End If
    ' Find b
    re.Pattern = "(?:^|\n)\s*vscode_path\s*=\s*""([^""]+)"""
    Set matches = re.Execute(fileText)
    If matches.Count > 0 Then
        VSCODE_PATH = matches(0).SubMatches(0)
    Else:
        GoTo RegistrationErrorHandler
    End If
    ' re.Pattern = "(?:^|\n)\s*xlpro_xlam_path\s*=\s*""([^""]+)"""
    ' Set matches = re.Execute(fileText)
    ' If matches.Count > 0 Then
    '     XLPRO_ADDIN_PATH = matches(0).SubMatches(0)
    ' Else:
    '     GoTo RegistrationErrorHandler
    ' End If
    Exit Sub
    
RegistrationErrorHandler:
    MsgBox "Error: Error, could not load config.toml xlpro_cli_path or vscode_path", _
           vbExclamation, "Warning"
    Err.Clear
    Exit Sub

End Sub

Public Sub initialize_workbook_guid_map()
    If Not WORKBOOK_GUID_MAP_INITIALIZED Then
        ' Set WORKBOOK_GUID_MAP = New Scripting.Dictionary
        Set WORKBOOK_GUID_MAP = CreateObject("Scripting.Dictionary")
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
    WORKBOOK_GUID_MAP.Add wb_name, guid
End Sub
Sub del_workbook_guid_map_key(wb_name As String)
    If WORKBOOK_GUID_MAP.Exists(wb_name) Then
        WORKBOOK_GUID_MAP.Remove wb_name
    End If
End Sub
Public Function get_workbook_guid_map_value(wb_name As String) As String
    If WORKBOOK_GUID_MAP.Exists(wb_name) Then
        get_workbook_guid_map_value = WORKBOOK_GUID_MAP(wb_name)
    Else
        get_workbook_guid_map_value = ""
    End If
End Function



'------------------------------------------------------------------------
' UI ribbon elements
'------------------------------------------------------------------------

' Function RunProcessAndWaitForStdErrMsg(command As String, msg As String, timeoutSeconds As Double) as Boolean
' ' Runs a shell command in a new window and waits for a signal
'     Dim shell As Object
'     Set shell = CreateObject("WScript.Shell")

'     Dim exec As Object
'     Set exec = shell.Exec(command)

'     Dim output As String
'     Dim startTime As Single
'     startTime = Timer

'     Do While Not exec.StdErr.AtEndOfStream
'         output = exec.StdErr.ReadLine
'         Debug.Print "stderr: " & output

'         If InStr(output, msg) > 0 Then
'             MsgBox "Signal string received: " & msg
'             Exit Do
'         End If

'         If Timer - startTime > timeoutSeconds Then
'             MsgBox "Timeout waiting for signal string."
'             Exit Do
'         End If
'     Loop

'     ' Optional: Wait for process to exit
'     Do While exec.Status = 0
'         DoEvents
'         If Timer - startTime > timeoutSeconds Then
'             MsgBox "Timeout waiting for process to exit."
'             Exit Do
'         End If
'     Loop

'     MsgBox "Process exited with code: " & exec.ExitCode
' End Sub


Sub xlproStart(ByRef control As Office.IRibbonControl)
    Dim taskID As Double
    Dim command As String
    
    Dim wb As Workbook
    
    Set wb = ActiveWorkbook

    ' Load the toml xlpro configuration file to ensure the paths are correct
    LoadXlproConfigTOML

    ' Use the Shell function to call the program
    command = """" & XLPRO_CLI_PATH & """" & " start " & """" & wb.Path & "\" & wb.Name & """"
    Debug.Print command
    taskID = shell("cmd /c " & """" & command & """", vbNormalFocus)

    ' ' Run the shell command and wait for the msg in the stderr to signal that it's ready to receive a start call
    ' dim fullCommand as string
    ' Debug.Print fullCommand
    ' fullCommand = "cmd /c " & """" & command & """"    
    ' dim success as boolean
    ' success = RunProcessAndWaitForStdErrMsg command:=fullCommand, msg:="XLPRO_TRIGGER_START", timeoutSeconds:=10.0
    
    ' if not success Then
    '     MsgBox "RunProcessAndWaitForStdErr success"
    ' Else
    '     MsgBox "RunProcessAndWaitForStdErr fail"
    ' End if

    ' clear the workbook guid key if it exists
    initialize_workbook_guid_map
    del_workbook_guid_map_key wb.Name

    ' Optionally, display the Task ID of the program
    ' Debug.Print "Program launched with Task ID: " & taskID

End Sub

Sub xlproInit(ByRef control As Office.IRibbonControl)
    Dim taskID As Double
    Dim command As String

    LoadXlproConfigTOML

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
    xlproRegisterWorkbook wb
End Sub

Sub xlproRegisterWorkbook(wb as Workbook)
' Register the workbook
    Debug.Print ActiveWorkbook.Path
    
    LoadXlproConfigTOML

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

    LoadXlproConfigTOML

    command = """" & VSCODE_PATH & """" & " " & """" & wb.Path & "\" & ActiveWorkbook.Name & ".xlpro" & """"
    Debug.Print command
    taskID = shell(command, vbNormalFocus)
End Sub

Sub EditConfigGlobal()
    ' Edit the global xlpro configuration file.
    Dim configPath As String
    configPath = Environ("USERPROFILE") & "\.xlpro\config.toml"
    ' Add quotes in case the path contains spaces
    Shell "cmd /c start """" """ & configPath & """", vbNormalFocus
End Sub

Sub EditConfigGlobalButton(ByRef control As Office.IRibbonControl)
    EditConfigGlobal
End Sub

Sub PushRequirementsTxtButton(ByRef control As Office.IRibbonControl)
    PushRequirementsTxt
End Sub

Sub OpenWorkingDirButton(ByRef control As Office.IRibbonControl)
    OpenWorkingDir
End Sub

'------------------------------------------------------------------------
'Items below here are helper subroutines for the addin.
'------------------------------------------------------------------------

Sub PushRequirementsTxt()
    Dim command As String
    Dim taskID As Double
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    LoadXlproConfigTOML
    command = """" & XLPRO_CLI_PATH & """" & " write-reqs " & """" & wb.Path & "\" & wb.Name & """"
    Debug.Print command
    taskID = shell("cmd /c " & """" & command & """", vbNormalFocus)

End Sub

Sub OpenWorkingDir()
    Dim command As String
    Dim taskID As Double
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    Dim folderPath As String

    folderPath = wb.Path
    command = "explorer.exe """ & folderPath & """"
    Debug.Print command
    shell command, vbNormalFocus

End Sub



' Sub AddReferenceToMyAddin()
'     Dim vbProj As VBIDE.VBProject
'     Dim refPath As String

'     LoadXlproConfigTOML

'     ' Full path to your add-in (adjust as needed)
'     refPath = XLPRO_ADDIN_PATH

'     ' Set reference to the current project
'     Set vbProj = ThisWorkbook.VBProject

'     ' Add reference if it's not already present
'     On Error Resume Next
'     vbProj.References.AddFromFile refPath
'     If Err.Number <> 0 Then
'         MsgBox "Failed to add reference: " & Err.Description, vbExclamation
'     Else
'         MsgBox "Reference to myaddin.xlam added.", vbInformation
'     End If
'     On Error GoTo 0
' End Sub


Sub TestArg()
    Dim res as Boolean
    Dim arg as Variant

    Set arg = ActiveSheet.Range("A1")

    res = Application.Run("'xlpro.xlam'!CheckArgReady", arg)

End Sub

Sub InitRegex()
    Set rePromise = CreateObject("VBScript.RegExp")
    rePromise.Global = False
    rePromise.IgnoreCase = True
    rePromise.Pattern = "^Promise<.+>$"

    Set reException = CreateObject("VBScript.RegExp")
    reException.Global = False
    reException.IgnoreCase = True
    reException.Pattern = "^\w*((error)|(exception))\(.*\)$"
End Sub

Function IsValueReady(val as Variant) as Boolean
    If rePromise Is Nothing Or reException Is Nothing Then
        InitRegex
    End If
    If rePromise.Test(val) Then
        IsValueReady = False
        Exit Function
    ' Test for Error
    ElseIf reException.Test(val) Then
        IsValueReady = False
        Exit Function
    ElseIf IsError(val) Then
        IsValueReady = False
        Exit Function
    ElseIf IsEmpty(val) Then
        IsValueReady = False
        Exit Function
    End If
    IsValueReady = True
End Function

Public Function CheckArgReady(arg as Variant) as Boolean
    dim val as Variant
    dim subval as Variant

    Dim i as Long
    Dim j as Long
    Dim dimCount1 as Long
    Dim dimCount2 as Long

    dimCount1 = 1
    dimCount2 = 1

    If TypeName(arg) = "Range" Then
        val = arg.Value
    Else
        val = arg
    End If

    ' If the value is an array, check the dimensions do not exceed
    ' our specified limit, Assume it is valid for large arrays and let COM
    ' call handle the rest. 
    If IsArray(val) Then
        dimCount1 = UBound(val, 1)
        ' Determine if 1D or 2D array
        On Error Resume Next
        dimCount2 = UBound(val, 2)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            dimCount2 = 1
        Else
            Err.Clear
            On Error GoTo 0
            dimCount2 = Ubound(val, 2)
        End If
        ' Complete check for max inputs
        if (dimCount1 * dimCount2) > MAX_ARGS_READY_CHECKED Then
            CheckArgReady = True
            Exit Function
        End If
    End If

    ' If the value is an array, loop over all the items
    ' and return false as soon as a bad value is encountered
    If IsArray(val) Then
        ' Determine if 1D or 2D array
        dimCount1 = UBound(val, 1)
        On Error Resume Next
        dimCount2 = UBound(val, 2)
        If Err.Number <> 0 Then
            ' 1D array (either row or column)
            Err.Clear
            On Error GoTo 0
            dimCount2 = 1
            For i = LBound(val, 1) To UBound(val, 1)
                subval = val(i)
                If Not IsValueReady(subval) Then
                    CheckArgReady = False
                    Exit Function
                End If
            Next i
        Else
            ' 2D array
            On Error GoTo 0
            For i = LBound(val, 1) To UBound(val, 1)
                dimCount2 = Ubound(val, 2)
                For j = LBound(val, 2) To UBound(val, 2)
                    subval = val(i, j)
                    If Not IsValueReady(subval) Then
                        CheckArgReady = False
                        Exit Function
                    End If
                Next j
            Next i
        End If
    Else
        ' Single value (0D)
        If Not IsValueReady(val) Then
            CheckArgReady = False
            Exit Function
        End If
    End If

    CheckArgReady = True
End Function


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

' Sub reload_global_config(ByRef wb As Workbook)
' 'Use a workbook com server to reload the configuration
'     Dim xlpro_async As Object
'     Dim guid As String
'     guid = get_workbook_guid_map_value(wb.Name)
'     Set xlpro_async = GetObject("new: " & guid)
'     Dim dict as Object
'     Set dict = xlpro_async.reload_and_get_config
'     XLPRO_CLI_PATH = dict.Item("xlpro_cli_path")
'     VSCODE_PATH = dict.Item("vscode_path")
' End Sub


' Replacement synchronization functions
Sub write_vba_sync_module(ByRef wb As Workbook)
    Dim contents1 As String
    contents1 = get_vba_sync_text(wb)
    write_text_to_module wb, "xlpro_async", contents1
    Dim contents2 As String
    contents2 = get_vba_sync_text_subs(wb)
    write_text_to_module wb, "xlpro_async_subs", contents2
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
Function get_vba_sync_text_subs(wb As Workbook) As String
    Dim xlpro_async As Object
    Dim guid As String
    guid = get_workbook_guid_map_value(wb.Name)
    Set xlpro_async = GetObject("new: " & guid)
    get_vba_sync_text_subs = xlpro_async.get_vba_sync_text_subs(wb)
    
End Function

Sub write_text_to_module(ByRef wb As Workbook, c_name As String, contents As String)
    ' Dim proj As VBIDE.VBProject
    Dim proj As Object

    ' Dim codemod As VBIDE.codemodule
    Dim codemod As Object

    ' Dim comp As VBIDE.VBComponent
    Dim comp As Object

    Set proj = wb.VBProject

    Dim names() As String ' Dynamic array to store the names
    Dim i As Long
    
    ' Resize the array to the number of items in the collection
    ReDim names(1 To proj.vbcomponents.Count)
    
    ' Loop through the collection
    For i = 1 To proj.vbcomponents.Count
        names(i) = proj.vbcomponents(i).Name ' Extract the .Name property of each element
    Next i
    
    If Not IsInArray(c_name, names) Then
        ' Set comp = proj.vbcomponents.Add(vbext_ct_StdModule)
        Set comp = proj.vbcomponents.Add(1)
        comp.Name = c_name
    Else:
        Set comp = proj.vbcomponents(c_name)
    End If
    
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
' Custom pointer function which is passable to xlpro as an argument.
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set ws = rng.Parent
    Set wb = ws.Parent
    
    ptr = "*<" & wb.FullName & "::" & ws.Name & "::" & rng.Address & ">"
End Function







' Sub conditional_format_handler(workbook_name As String, sheet_name As String, range_names As Variant, colors As Variant, return_uid As String)
'     'pass the areas and conditional formatting colours for each area
'     Dim Workbook As Workbook
'     Dim Worksheet As Worksheet
    
'     Dim area As Range
'     Dim rng As Range
'     Dim i As Long
    
'     If Not (LBound(range_names) = LBound(colors) And UBound(colors) = UBound(range_names)) Then
'         Err.Raise 9999, "xlproError", "Array bounds do not match"
'     End If
    
'     Set Workbook = Workbooks(workbook_name)
'     Set Worksheet = Workbook.Sheets(sheet_name)

'     Dim cell_color As Long

'     ' create the range (which could have multi-areas)
'     ' loop over the areas and set the interior color to the target color.
'     For i = LBound(range_names) To UBound(range_names)
'         Set rng = Worksheet.Range(range_names(i))
'         cell_color = colors(i)
'         For Each area In rng.areas
'             area.Interior.Color = cell_color
'         Next area
'     Next i
    
'     Set rng = Nothing
'     Set Workbook = Nothing
'     Set Worksheet = Nothing
'     Set area = Nothing
    
'     'Dim xlpro As Object
'     'Set xlpro = GetObject("new: " & XLPRO_GUID)
'     'xlpro.set_return_value(uid, "Success('" & uid & "')"
    
' End Sub

' Function conditional_formatter(formula_str As String, rng As Range, root_cell As Range) As Variant

'     Dim rng_name_arr As Variant
'     Dim rng_vals As Variant
    
'     ' ReDim rng_name_arr(1 To rng.Count)
'     Dim i As Integer
    
'     Dim rng_name As String
'     Dim workbook_name As String
'     Dim sheet_name As String
'     Dim root_cell_addr As String
'     Dim area_names_arr As Object

    
'     workbook_name = rng.Parent.Parent.Name
'     sheet_name = rng.Parent.Name
'     area_names_arr = get_range_name_per_area(rng)
'     root_cell_addr = root_cell.Address
    
'     rng_vals = get_range_values(Join(area_names_arr, ","))
    
'     Dim xlpro As Object
'     Set xlpro = GetObject("new: " & XLPRO_GUID)
'     conditional_formatter = xlpro.execute_function_async(ActiveWorkbook, Application.Caller, "conditional_formatter_example", workbook_name, sheet_name, rng_name_arr, formula_str, root_cell_addr)
' End Function

' Private Function get_range_name_per_area(rng As Range) As Variant()

' Dim cell As Range
' Dim values() As Variant
' Dim i As Long
' Dim area As Range

' ' rng.Interior.Color = 16744319
' ReDim values(1 To rng.Cells.Count)

' i = 1
' For Each area In rng.areas
'     values(i) = area.Address
'     i = i + 1
' Next area

' get_range_name_per_area = values

' End Function

' Private Function get_range_name_per_cell(rng As Range) As Variant()

' Dim cell As Range
' Dim values() As Variant
' Dim i As Long
' Dim area As Range

' ' rng.Interior.Color = 16744319
' ReDim values(1 To rng.Cells.Count)

' i = 1
' For Each area In rng.areas
'     For Each cell In area.Cells
'         values(i) = cell.Address
'         i = i + 1
'     Next cell
' Next area

' get_range_name_per_cell = values

' End Function

' Private Function get_range_values(rng_name As String) As Variant()

' Dim rng As Range
' Dim cell As Range
' Dim values() As Variant
' Dim i As Long
' Dim area As Range

' Set rng = Range(rng_name)

' rng.Interior.Color = 16744319
' ReDim values(1 To rng.Cells.Count)

' i = 1
' For Each area In rng.areas
'     For Each cell In area.Cells
'         values(i) = cell.Value
'         i = i + 1
'     Next cell
' Next area

' get_range_values = values

' End Function



