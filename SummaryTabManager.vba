Sub ManageSummaryTab()
    Dim ws As Worksheet
    Dim tempTabName As String
    Dim summaryExists As Boolean
    Dim wsCount As Integer
    Dim i As Integer
    
    ' Set the temporary tab name
    tempTabName = "TempSummary_" & Format(Now, "hhmmss")
    
    ' Check if Summary tab exists
    summaryExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Summary" Then
            summaryExists = True
            Exit For
        End If
    Next ws
    
    ' If Summary tab exists, rename it to temporary name
    If summaryExists Then
        Worksheets("Summary").Name = tempTabName
        MsgBox "Summary tab found and renamed to: " & tempTabName
    Else
        MsgBox "No Summary tab found in the workbook"
    End If
    
    ' Create new Summary tab
    Worksheets.Add.Name = "Summary"
    MsgBox "New Summary tab created"
    
    ' Delete the temporary tab if it exists
    If summaryExists Then
        Application.DisplayAlerts = False
        Worksheets(tempTabName).Delete
        Application.DisplayAlerts = True
        MsgBox "Temporary tab deleted: " & tempTabName
    End If
    
    ' Count remaining worksheets (excluding the new Summary tab)
    wsCount = ThisWorkbook.Worksheets.Count
    
    ' Delete all remaining tabs except the new Summary tab
    Application.DisplayAlerts = False
    
    ' Loop through worksheets from last to first to avoid index issues
    For i = wsCount To 1 Step -1
        If ThisWorkbook.Worksheets(i).Name <> "Summary" Then
            ThisWorkbook.Worksheets(i).Delete
        End If
    Next i
    
    Application.DisplayAlerts = True
    
    MsgBox "All remaining tabs have been deleted. Only the new Summary tab remains."
    
End Sub