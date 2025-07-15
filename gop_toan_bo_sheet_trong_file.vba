Sub GopSheet()
    Dim ws As Worksheet
    Dim wsMerged As Worksheet
    Dim lastRow As Long
    Dim destRow As Long
    Dim mergedSheetExists As Boolean

    ' Kiểm tra xem sheet kết quả "GopSheetResult" đã tồn tại chưa
    mergedSheetExists = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "GopSheetResult" Then
            mergedSheetExists = True
            Exit For
        End If
    Next ws

    ' Nếu sheet kết quả chưa tồn tại, tạo mới
    If Not mergedSheetExists Then
        Set wsMerged = Sheets.Add(After:=Sheets(Sheets.Count))
        wsMerged.Name = "GopSheetResult"
        destRow = 1
    Else
        ' Nếu đã tồn tại, lấy dòng trống tiếp theo
        Set wsMerged = Sheets("GopSheetResult")
        destRow = wsMerged.Cells(wsMerged.Rows.Count, "A").End(xlUp).Row + 1
    End If

    ' Duyệt qua từng sheet để gộp dữ liệu
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> wsMerged.Name Then
            ' Kiểm tra xem cột B có ít nhất 2 ô chứa dữ liệu không
            If Application.WorksheetFunction.CountA(ws.Columns("B")) > 1 Then
                lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

                ' Sao chép dữ liệu từ A1 đến cột Z (tối đa 26 cột) sang sheet kết quả
                ws.Range("A1:Z" & lastRow).Copy Destination:=wsMerged.Cells(destRow, 1)

                ' Cập nhật dòng đích cho lần sao chép tiếp theo
                destRow = wsMerged.Cells(wsMerged.Rows.Count, "A").End(xlUp).Row + 1
            End If
        End If
    Next ws

    ' Hiển thị thông báo khi hoàn thành
    MsgBox "Quá trình gộp sheet đã hoàn thành!", vbInformation
End Sub
