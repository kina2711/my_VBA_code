Sub ClearUnwantedColumnsAndSort()
    Dim ws As Worksheet
    Dim rng As Range
    Dim col As Range
    Dim keepCols As Variant
    Dim i As Long
    Dim lastCol As Long
    Dim colLetter As String
    
    ' Danh sách các cột muốn giữ lại
    keepCols = Array("A", "B", "D", "E", "G", "J", "L", "M", "U", "W", "X", "AB", "AF", "AM", "AP", "AX", "BL", "BQ", "BS")
    
    ' Xác định worksheet hiện tại
    Set ws = ActiveSheet
    
    ' Xác định phạm vi dữ liệu hiện có trên worksheet
    Set rng = ws.UsedRange
    
    ' Lặp qua từng cột trong phạm vi dữ liệu
    For Each col In rng.Columns
        colLetter = Split(Cells(1, col.Column).Address, "$")(1)
        
        ' Kiểm tra nếu cột không nằm trong danh sách cần giữ lại
        If IsError(Application.Match(colLetter, keepCols, 0)) Then
            ' Xóa dữ liệu trong cột đó
            col.ClearContents
        End If
    Next col
    
    ' Xác định phạm vi dữ liệu mới sau khi xóa
    Set rng = ws.UsedRange
    
    ' Sắp xếp trang tính theo cột M
    rng.Sort Key1:=ws.Range("M1"), Order1:=xlAscending, Header:=xlYes
    
    ' Xóa các cột trống hoàn toàn sau khi sắp xếp
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For i = lastCol To 1 Step -1 ' Duyệt từ phải sang trái để tránh lỗi khi xóa
        If Application.WorksheetFunction.CountA(ws.Columns(i)) = 0 Then
            ws.Columns(i).Delete
        End If
    Next i
    
    MsgBox "Hoàn thành!", vbInformation
End Sub
