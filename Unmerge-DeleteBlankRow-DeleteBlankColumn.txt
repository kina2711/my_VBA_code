Sub CleanSheet()
    Dim ws As Worksheet
    Dim rng As Range
    Dim i As Long
    
    Set ws = ActiveSheet
    Set rng = ws.UsedRange
    
    ' Unmerge all cells
    ws.Cells.UnMerge
    
    ' Delete empty rows
    For i = rng.Rows.Count To 1 Step -1
        If Application.WorksheetFunction.CountA(rng.Rows(i)) = 0 Then
            rng.Rows(i).Delete
        End If
    Next i
    
    ' Update range after deleting rows
    Set rng = ws.UsedRange
    
    ' Delete empty columns
    For i = rng.Columns.Count To 1 Step -1
        If Application.WorksheetFunction.CountA(rng.Columns(i)) = 0 Then
            rng.Columns(i).Delete
        End If
    Next i
End Sub
