Attribute VB_Name = "xlpro_static"


Option Explicit

Public XLPRO_CLI_PATH As String
Public VSCODE_PATH As String
Public XLPRO_SERVER_PATH As String
Public UNDOSTACKDEPTH As Long

' Public XLPRO_ADDIN_PATH As String

Public WORKBOOK_GUID_MAP As Object
Public WORKBOOK_GUID_MAP_INITIALIZED As Boolean

Public WSCRIPT_SHELL As Object
Public WSCRIPT_SHELL_INITIALIZED As Boolean

Public Const MAX_ARGS_READY_CHECKED As Integer = 32

Public rePromise As Object
Public reException As Object


#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Public Sub LoadXlproConfigTOML()
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
    re.Pattern = "(?:^|\n)\s*XLPRO_CLI_PATH\s*=\s*""([^""]+)"""
    Set matches = re.Execute(fileText)
    If matches.Count > 0 Then
        XLPRO_CLI_PATH = matches(0).SubMatches(0)
    Else:
        GoTo RegistrationErrorHandler
    End If
    ' Find b
    re.Pattern = "(?:^|\n)\s*VSCODE_PATH\s*=\s*""([^""]+)"""
    Set matches = re.Execute(fileText)
    If matches.Count > 0 Then
        VSCODE_PATH = matches(0).SubMatches(0)
    Else:
        GoTo RegistrationErrorHandler
    End If
    re.Pattern = "(?:^|\n)\s*XLPRO_SERVER_PATH\s*=\s*""([^""]+)"""
    Set matches = re.Execute(fileText)
    If matches.Count > 0 Then
        XLPRO_SERVER_PATH = matches(0).SubMatches(0)
    Else:
        GoTo RegistrationErrorHandler
    End If
    re.Pattern = "(?:^|\n)\s*UNDO_STACK_DEPTH\s*=\s*(\d+)"
    Set matches = re.Execute(fileText)
    If matches.Count > 0 Then
        UNDOSTACKDEPTH = CLng(matches(0).SubMatches(0))
    Else:
        GoTo RegistrationErrorHandler
    End If
    Exit Sub
    
RegistrationErrorHandler:
    MsgBox "Error: Error, could not load config.toml XLPRO_CLI_PATH, VSCODE_PATH, XLPRO_SERVER_PATH, UNDO_STACK_DEPTH may not be defined correctly", _
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
    ' TODO add contingency for the guid map not existing without initializing the workbook
    If WORKBOOK_GUID_MAP.Exists(wb_name) Then
        get_workbook_guid_map_value = WORKBOOK_GUID_MAP(wb_name)
    Else
        get_workbook_guid_map_value = ""
    End If
End Function



'------------------------------------------------------------------------
' UI ribbon elements
'------------------------------------------------------------------------

Sub xlproStart(ByRef control As Office.IRibbonControl)
    Dim taskID As Double
    Dim command As String
    
    Dim Wb As Workbook
    
    Set Wb = ActiveWorkbook

    ' Load the toml xlpro configuration file to ensure the paths are correct
    LoadXlproConfigTOML

    ' Use the Shell function to call the program
    command = """" & XLPRO_CLI_PATH & """" & " start " & """" & Wb.Path & "\" & Wb.name & """"
    Debug.Print command
    ' taskID = shell("cmd /c " & """" & command & """", vbNormalFocus)
    taskID = shell(command, vbNormalFocus)

    ' clear the workbook guid key if it exists
    initialize_workbook_guid_map
    del_workbook_guid_map_key Wb.name

    ' Optionally, display the Task ID of the program
    ' Debug.Print "Program launched with Task ID: " & taskID

End Sub

Sub xlproInit(ByRef control As Office.IRibbonControl)
    Dim taskID As Double
    Dim command As String

    LoadXlproConfigTOML

    ' Use the Shell function to call the program
    'taskID = Shell("cmd.exe /K xlpro", vbNormalFocus)
    ' command = """" & XLPRO_CLI_PATH & """" & " init " & """" & ActiveWorkbook.Path & "\" & ActiveWorkbook.name & """"
    ' Debug.Print command
    ' taskID = shell("cmd /c " & """" & command & """", vbNormalFocus)

    command = """" & XLPRO_CLI_PATH & """" & " init " & """" & ActiveWorkbook.Path & "\" & ActiveWorkbook.name & """"
    Debug.Print command
    taskID = shell(command, vbNormalFocus)

    ' Optionally, display the Task ID of the program
    Debug.Print "Program launched with Task ID: " & taskID

End Sub

Sub xlproRegister(ByRef control As Office.IRibbonControl)
    Dim Wb As Workbook
    Set Wb = ActiveWorkbook
    xlproRegisterWorkbook Wb
End Sub

Sub xlproRegisterWorkbook(Wb As Workbook)
' Register the workbook
    Debug.Print ActiveWorkbook.Path
    
    LoadXlproConfigTOML

    ' clear the workbook guid key if it exists
    initialize_workbook_guid_map
    del_workbook_guid_map_key Wb.name

    On Error GoTo RegistrationErrorHandler
    register_workbook ActiveWorkbook
    On Error GoTo 0
    
    write_vba_sync_module ActiveWorkbook
    Exit Sub
    
RegistrationErrorHandler:
    MsgBox "Error: Error during registration. Check the logs and and/or connect the debugger with 'Uncaught Exceptions' enabled and retry.", _
           vbExclamation, "Warning"
    Err.Clear
    Exit Sub
End Sub


Sub xlproStartIDE(ByRef control As Office.IRibbonControl)
    
    Dim command As String
    Dim taskID As Double
    
    Dim Wb As Workbook
    Set Wb = ActiveWorkbook

    LoadXlproConfigTOML

    command = """" & VSCODE_PATH & """" & " " & """" & Wb.Path & "\" & ActiveWorkbook.name & ".xlpro" & """"
    Debug.Print command
    taskID = shell(command, vbNormalFocus)
End Sub

Sub EditConfigGlobal()
    ' Edit the global xlpro configuration file.
    Dim configPath As String
    configPath = Environ("USERPROFILE") & "\.xlpro\config.toml"
    ' Add quotes in case the path contains spaces
    shell "cmd /c start """" """ & configPath & """", vbNormalFocus
    ' shell """" & configPath & """", vbNormalFocus
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

Sub OpenXlproInstallDirButton(ByRef control As Office.IRibbonControl)
    OpenXlproInstallDir
End Sub

Sub ClearVenvDataButton(ByRef control As Office.IRibbonControl)
    ClearVenvData
End Sub

Sub ClearTempDataButton(ByRef control As Office.IRibbonControl)
    ClearTempData
End Sub


Sub ClearUvPythonsButton(Byref control as Office.IRibbonControl)
    ClearUvPythons
End Sub

Sub ClearUvCacheButton(Byref control as Office.IRibbonControl)
    ClearUvCache
End Sub

Sub RemoveLinkForThisWorkbookButton(Byref control as Office.IRibbonControl)
    RemoveLink ActiveWorkbook
End Sub



Private Sub force_refresh_area_calculation(ByRef Wb As Workbook, rng as Range)

    Dim guid As String
    guid = get_workbook_guid_map_value(Wb.name)
    If guid = "" Then
        guid = get_workbook_guid(Wb)
        set_workbook_guid_map_pairing Wb.name, guid
    End If

    Debug.Print "XLPRO_GUID: " & guid

    Dim xlpro_async As Object
    Set xlpro_async = GetObject("new: " & guid)

    xlpro_async.force_refresh_area_calculation Wb, rng

End Sub

Sub force_refresh_area_calculationButton(ByRef control As Office.IRibbonControl)    
    Dim wb as Workbook
    set wb = ActiveWorkbook
    force_refresh_area_calculation wb, Application.Selection
End Sub

Sub ResetUndoButton(ByRef control As Office.IRibbonControl)    
    ' Reset the xlpro event handler manually
    Call ThisWorkbook.ResetEventHandler
End Sub
'------------------------------------------------------------------------
'Items below here are helper subroutines for the addin.
'------------------------------------------------------------------------

Sub PushRequirementsTxt()
    Dim command As String
    Dim taskID As Double
    
    Dim Wb As Workbook
    Set Wb = ActiveWorkbook

    LoadXlproConfigTOML
    command = """" & XLPRO_CLI_PATH & """" & " write-reqs " & """" & Wb.Path & "\" & Wb.name & """"
    ' Debug.Print command
    taskID = shell(command, vbNormalFocus)

End Sub

Sub ClearVenvData()
    Dim command As String
    Dim taskID As Double
    LoadXlproConfigTOML
    command = """" & XLPRO_CLI_PATH & """" & " clear-venvs" 
    taskID = shell(command, vbNormalFocus)
End Sub

Sub ClearTempData()
    Dim command As String
    Dim taskID As Double
    LoadXlproConfigTOML
    command = """" & XLPRO_CLI_PATH & """" & " clear-tmp"
    taskID = shell(command, vbNormalFocus)
End Sub

Sub ClearUvPythons()
    Dim command As String
    Dim taskID As Double
    LoadXlproConfigTOML
    command = """" & XLPRO_CLI_PATH & """" & " clear-uv-pythons"
    taskID = shell(command, vbNormalFocus)
End Sub

Sub ClearUvCache()
    Dim command As String
    Dim taskID As Double
    LoadXlproConfigTOML
    command = """" & XLPRO_CLI_PATH & """" & " clear-uv-cache"
    taskID = shell(command, vbNormalFocus)
End Sub

Sub RemoveLink(Wb as Workbook)
    Dim command As String
    Dim taskID As Double
    Set Wb = ActiveWorkbook

    LoadXlproConfigTOML
    command = """" & XLPRO_CLI_PATH & """" & " clear-venv-link " & """" & Wb.Path & "\" & Wb.name & """"
    taskID = shell(command, vbNormalFocus)
End Sub



Sub OpenWorkingDir()
    Dim command As String
    Dim taskID As Double
    
    Dim Wb As Workbook
    Set Wb = ActiveWorkbook
    
    Dim folderPath As String

    folderPath = Wb.Path
    command = "explorer.exe """ & folderPath & """"
    ' Debug.Print command
    shell command, vbNormalFocus

End Sub

Sub OpenXlproInstallDir()
    Dim command As String
    Dim taskID As Double
    
    Dim userProfilePath As String
    Dim folderPath As String

    ' Get the USERPROFILE environment variable
    userProfilePath = Environ("USERPROFILE")
    folderPath = userProfilePath & "\.xlpro"

    command = "explorer.exe """ & folderPath & """"
    ' Debug.Print command
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
    Dim res As Boolean
    Dim arg As Variant

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

Function IsValueReady(val As Variant) As Boolean
    If rePromise Is Nothing Or reException Is Nothing Then
        InitRegex
    End If
    If IsError(val) Then
        IsValueReady = False
        Exit Function
    ElseIf rePromise.Test(val) Then
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

Public Function IsNoneOrEmpty(arg As Variant) As Boolean
    If TypeName(arg) <> "String" Then
        IsNoneOrEmpty = False
        Exit Function
    End If
    IsNoneOrEmpty = (arg = "pyEmpty") Or (arg = "pyNone")
End Function

Public Function CheckArgReady(arg As Variant) As Boolean
    Dim val As Variant
    Dim subval As Variant

    Dim i As Long
    Dim j As Long
    Dim dimCount1 As Long
    Dim dimCount2 As Long

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
            dimCount2 = UBound(val, 2)
        End If
        ' Complete check for max inputs
        If (dimCount1 * dimCount2) > MAX_ARGS_READY_CHECKED Then
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
                dimCount2 = UBound(val, 2)
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

Function get_workbook_guid(Wb As Workbook) As String
    Dim command As String
    Dim shell As Object
    Dim exec As Object
    Dim output As String
    Dim line As String

    On Error GoTo ErrHandler
'    Set shell = CreateObject("WScript.Shell")
    command = """" & XLPRO_CLI_PATH & """" & " guid " & """" & ActiveWorkbook.Path & "\" & ActiveWorkbook.name & """"
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

' Sub unregister_activeworkbook()
'     uninitialize ActiveWorkbook
' End Sub
Private Sub register_workbook(ByRef Wb As Workbook)
    On Error GoTo 0
    InitializeShell
    initialize_workbook_guid_map
    Dim guid As String
    guid = get_workbook_guid_map_value(Wb.name)
    If guid = "" Then
        guid = get_workbook_guid(Wb)
        set_workbook_guid_map_pairing Wb.name, guid
    End If

    Debug.Print "XLPRO_GUID: " & guid

    Dim xlpro_async As Object
    Set xlpro_async = GetObject("new: " & guid)

    xlpro_async.register_and_configure_wb_workspace Wb

End Sub


' Sub uninitialize(ByRef Wb As Workbook)
' 'Uninitialize this workbook from the com server
'     Dim xlpro_async As Object
'     Dim guid As String
'     guid = get_workbook_guid_map_value(Wb.name)
'     Set xlpro_async = GetObject("new: " & guid)
'     xlpro_async.shutdown_workspace Wb
' End Sub

' Private Sub shutdown_xlpro(ByRef Wb As Workbook)
' 'Attempt to shutdown the xlpro server.
'     Dim xlpro_async As Object
'     Dim guid As String
'     guid = get_workbook_guid_map_value(Wb.name)
'     Set xlpro_async = GetObject("new: " & guid)
'     xlpro_async.shutdown
' End Sub

' Private Sub getpid(ByRef Wb As Workbook)
'     Dim xlpro_async As Object
'     Dim guid As String
'     guid = get_workbook_guid_map_value(Wb.name)
'     Set xlpro_async = GetObject("new: " & guid)
'     Debug.Print xlpro_async.getpid
' End Sub

' Private Sub shutdown_workspace(ByRef Wb As Workbook)
'     Dim xlpro_async As Object
'     Dim guid As String
'     guid = get_workbook_guid_map_value(Wb.name)
'     Set xlpro_async = GetObject("new: " & guid)
'     xlpro_async.shutdown_workspace Wb
' End Sub

' Sub reload_global_config(ByRef wb As Workbook)
' 'Use a workbook com server to reload the configuration
'     Dim xlpro_async As Object
'     Dim guid As String
'     guid = get_workbook_guid_map_value(wb.Name)
'     Set xlpro_async = GetObject("new: " & guid)
'     Dim dict as Object
'     Set dict = xlpro_async.reload_and_get_config
'     XLPRO_CLI_PATH = dict.Item("XLPRO_CLI_PATH")
'     VSCODE_PATH = dict.Item("VSCODE_PATH")
' End Sub


' Replacement synchronization functions
Sub write_vba_sync_module(ByRef Wb As Workbook)
    Dim contents1 As String
    contents1 = get_vba_sync_text(Wb)
    write_text_to_module Wb, "xlpro_async", contents1
    Dim contents2 As String
    contents2 = get_vba_sync_text_subs(Wb)
    write_text_to_module Wb, "xlpro_async_subs", contents2
End Sub

Function get_vba_sync_text(Wb As Workbook) As String
    Dim xlpro_async As Object
    Dim guid As String
    guid = get_workbook_guid_map_value(Wb.name)
    Set xlpro_async = GetObject("new: " & guid)
    get_vba_sync_text = xlpro_async.get_vba_sync_text(Wb)
    
'SyncTextErrorHandler:
'    MsgBox "Error: Error during VBA Module sync. Please connect the debugger and retry.", _
'           vbExclamation, "Warning"
'    Err.Clear
'    Exit Function

End Function
Function get_vba_sync_text_subs(Wb As Workbook) As String
    Dim xlpro_async As Object
    Dim guid As String
    guid = get_workbook_guid_map_value(Wb.name)
    Set xlpro_async = GetObject("new: " & guid)
    get_vba_sync_text_subs = xlpro_async.get_vba_sync_text_subs(Wb)
    
End Function

Sub write_text_to_module(ByRef Wb As Workbook, c_name As String, contents As String)
    ' Dim proj As VBIDE.VBProject
    Dim proj As Object

    ' Dim codemod As VBIDE.codemodule
    Dim codemod As Object

    ' Dim comp As VBIDE.VBComponent
    Dim comp As Object

    Set proj = Wb.VBProject

    Dim names() As String ' Dynamic array to store the names
    Dim i As Long
    
    ' Resize the array to the number of items in the collection
    ReDim names(1 To proj.vbcomponents.Count)
    
    ' Loop through the collection
    For i = 1 To proj.vbcomponents.Count
        names(i) = proj.vbcomponents(i).name ' Extract the .Name property of each element
    Next i
    
    If Not IsInArray(c_name, names) Then
        ' Set comp = proj.vbcomponents.Add(vbext_ct_StdModule)
        Set comp = proj.vbcomponents.Add(1)
        comp.name = c_name
    Else:
        Set comp = proj.vbcomponents(c_name)
    End If
    
    Set codemod = comp.codemodule
    
    codemod.DeleteLines 1, codemod.CountOfLines
    codemod.addfromstring contents
End Sub


Sub raiseSubroutineException(wb_name as String, fname as string, e_msg as String)
    dim msg as string

    msg = "Error encountered in '" & wb_name & "': '" & fname & "':" & vbnewline & "Details:"  & vbnewline &  e_msg

    MsgBox msg, vbCritical, "xlpro Subroutine Error: " & wb_name

End Sub


'------------------------------------------------------------------------
'Helper functions etc
'------------------------------------------------------------------------

Function IsInArray(Target As String, arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If element = Target Then
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
    Dim Wb As Workbook
    Dim ws As Worksheet
    
    Set ws = rng.Parent
    Set Wb = ws.Parent
    
    ptr = "*<" & Wb.FullName & "::" & ws.name & "::" & rng.Address & ">"
End Function


Public Sub ShowAsyncErrorWindow(caption as string, msg As String)
    dim frm as New UserForm1
    frm.Label1.Caption = msg
    frm.Caption = caption
    frm.Show vbModeless ' Async
End Sub




' You need a reference to "Microsoft Forms 2.0 Object Library"
' (contains the StdPicture type)

' Private Declare PtrSafe Function LoadPicture Lib "stdole2.tlb" Alias "LoadPictureA" _
'     (ByVal FileName As String) As stdole.IPictureDisp

' Public Function GetImageFromFile(filePath As String) As stdole.IPictureDisp
'     Set GetImageFromFile = LoadPicture(filePath)
' End Function


' Public Function GetImage(control As IRibbonControl) As stdole.IPictureDisp
'     Set GetImage = LoadPicture("C:\Users\Daniel Evans\projects\xlpro\xlpro_addin\assets\exports\archive-24_0.ico")
' End Function

' Public Function GetImage(control As IRibbonControl, path as string) As IPictureDisp
'     ' Dim path As String
'     ' path = ThisWorkbook.Path & "\myicon.ico"
'     ' path = "C:\Users\Daniel Evans\projects\xlpro\xlpro_addin\assets\exports\archive-24_0.ico"
'     ' path = "C:\Users\Public\archive-24_0.ico"
'     ' path = "C:\Users\Public\slide_export_0000.emf"
    
'     If Dir(path) = "" Then
'         MsgBox "Icon file not found: " & path
'         Exit Function
'     End If
    
'     GetImage = LoadPicture(path)
' End Function


' Public Sub OnLoadImage(ByVal sImageName As String, ByRef Image As Variant)
'    Set Image = LoadPicture("C:\Users\Public\" & "Capture.bmp") 
' End Sub 
