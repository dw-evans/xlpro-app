VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWorkbookEventHandler"

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True



' Class Module: CWorkbookEventHandler
Option Explicit

Public WithEvents Wb As Workbook
' Public WithEvents Ws As Worksheet

Private UndoStack As Collection

Private formulaCache As Variant
Private rngCache As Range
Private rngAddressCache As String

Private copyFormulaCache As Variant
Private cutFormulaCache As Variant
Private copyRngCache As Range

Private isCopyActive As Boolean
Private hasCopyDropped As Boolean

Private cutCopyMode As Long
Private isCutting As Boolean

Private selectionPrev As Range
Private selectionPrevAddress As String
Private selectionPrevSheet As Worksheet

Private cutRangeCacheAddress As String
Private cutRangeCacheSheet As Worksheet

Private isCutPasteOperation As Boolean

Private Sub Wb_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    
    ' Call undo tracking logic here
    Debug.Print "Change in " & Wb.name & " on " & Sh.name & ": " & Target.Address

    Dim rngExpanded As Range
    Dim endTarget As Range
    Dim endFormulaCache As Variant

    Dim isCutCopy As Boolean

    Dim toDisableIsCutting As Boolean
    Debug.Print Target.Address

    On Error GoTo ErrHandler
    If rngAddressCache <> Target.Address(False, False) Then
        Debug.Print "Warning: Cached range mismatch. Undo may be invalid."
        ' Exit Sub


        If Not copyRngCache Is Nothing Then
            ' If the first address misses, check if the expanded range (per the copy/paste region)
            ' address matches the selection (signal for a paste operation)
            Set rngExpanded = Target.Resize(copyRngCache.rows.Count, copyRngCache.Columns.Count)
            If rngExpanded.Address(False, False) <> Target.Address(False, False) Then
                Debug.Print "Warning: Secondary Cached range mismatch. Undo could not be resolved"
                Exit Sub
            End If
            Set endTarget = rngExpanded
            endFormulaCache = copyFormulaCache
            isCutCopy = True
            If isCutting Then
                If Not isCutPasteOperation Then
                    Debug.Print "this is the paste of the cut-paste operation, keeping"
                    toDisableIsCutting = True
                Else
                    Debug.Print "This is the clear of the cut-paste operation, skipping this."

                    isCutPasteOperation = False
                    Exit Sub
                End If


            End If
        Else
            Exit Sub
        End If

    Else
        Set endTarget = Target
        endFormulaCache = formulaCache

    End If
    
    Dim i As Long
    Dim j As Long
    Dim cutRange As Range
    ' For a cut change, add a step to simulate clearing the contents of the original range
    ' Add a step to replace the formula contents with blanks (not ideal, good enough for now.)
    ' Better solution could support clearcontents.
    If isCutting And (cutCopyMode = xlCut) Then
        Set cutRange = cutRangeCacheSheet.Range(cutRangeCacheAddress)
        Call GlobalUndoManager.TrackChange(Wb, Sh, cutRange, cutFormulaCache, cutRange.Formula2)
    End If
    
    If toDisableIsCutting Then
        isCutting = False
    End If
    
    ' Store the cached previous value against the new value
    Call GlobalUndoManager.TrackChange(Wb, Sh, endTarget, endFormulaCache, endTarget.Formula2)


    
    Exit Sub

ErrHandler:
    Debug.Print ("Could not track the change, maybe a multi-area range")
    Err.Clear
    On Error GoTo 0
    Exit Sub
    

End Sub



Private Sub Wb_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)

    If Target.Cells.CountLarge > 1000 Then
        Debug.Print "Too many cells to track"
        Exit Sub
    End If
    
    Dim rngExpanded As Range
    
    If Application.cutCopyMode <> False Then
        Debug.Print "User is preparing to paste or cut..."
        If Not hasCopyDropped Then
            Debug.Print "Caching target as copy state"
            hasCopyDropped = True
            ' Store the copied range so we know its size
            Set copyRngCache = rngCache

            ' copyFormulaCache = copyRngCache.Formula2
            ' Store the cut parameters immediately.
            cutCopyMode = Application.cutCopyMode
            If cutCopyMode = xlCut Then
                cutFormulaCache = selectionPrev.Formula2
                isCutting = True
                isCutPasteOperation = True
                cutRangeCacheAddress = selectionPrev.Address
                Set cutRangeCacheSheet = selectionPrev.Parent
            End If
        End If
            
        If Not copyRngCache Is Nothing Then
            'If we are waiting to copy/paste, a selection change needs to cache the formula
            'of the potentially modified range
            Set rngExpanded = Target.Resize(copyRngCache.rows.Count, copyRngCache.Columns.Count)
            copyFormulaCache = rngExpanded.Formula2

            ' Cache the selection or note that paste is pending
        End If

    Else
        hasCopyDropped = False
    End If

    
    ' Pre-change tracking (cache old values)
    Debug.Print "Selection in " & Wb.name & " on " & Sh.name & ": " & Target.Address

    ' Track the state of the target range
    formulaCache = Target.Formula2
    Set rngCache = Target
    rngAddressCache = Target.Address(False, False)
    
    Set selectionPrev = Target
    selectionPrevAddress = selectionPrev.Address
    Set selectionPrevSheet = selectionPrev.Parent
    

End Sub




