Sub MarkAndDeleteRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Thiết lập worksheet hiện tại
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Thay "Sheet1" bằng tên sheet của bạn
    
    ' Tìm hàng cuối cùng trong sheet
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' Bước 1: Duyệt từ hàng 12 đến hàng cuối, đánh dấu hàng có cột B trống
    For i = 12 To lastRow
        ' Kiểm tra ô trong cột B (cột 2) có trống không
        If Trim(ws.Cells(i, 2).Value) = "" Then
            ' Ghi số thứ tự hàng vào cột A (cột 1)
            ws.Cells(i, 1).Value = i
        End If
    Next i
    
    ' Bước 2: Duyệt ngược từ dưới lên, xóa các hàng đã đánh dấu
    For i = lastRow To 12 Step -1
        ' Nếu cột A có giá trị (số thứ tự), xóa hàng đó
        If ws.Cells(i, 1).Value <> "" Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub
