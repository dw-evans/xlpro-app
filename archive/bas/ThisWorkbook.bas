Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim answer As Integer

    If Cancel Then
        End Sub
    End If
    
    Dim xlpro_async As Object
    Set xlpro_async = GetObject("new: {122BB48A-57EF-4775-A28C-3F71ED0D02A7}")
    xlpro_async.shutdown_workspace ThisWorkbook
    
    End If
End Sub
