Sub CopyRowsToNewWorkbook()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim copyRow As Long
    Dim targetWorkbook As Workbook
    Dim targetSheet As Worksheet
    Dim sourceWorkbook As Workbook
    Dim targetFilePath As String

    ' Lưu lại workbook hiện tại
    Set sourceWorkbook = ThisWorkbook

    ' Chọn file workbook đã có sẵn
    targetFilePath = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx), *.xls; *.xlsx", , "Chọn Workbook đích")

    If targetFilePath = "False" Then
        MsgBox "Không chọn file workbook đích!", vbExclamation
        Exit Sub
    End If

    ' Mở workbook đích
    Set targetWorkbook = Workbooks.Open(targetFilePath)

    ' Lấy sheet đầu tiên của workbook đích (hoặc có thể chọn sheet cụ thể)
    Set targetSheet = targetWorkbook.Sheets(1)

    ' Duyệt qua từng sheet trong workbook hiện tại
    For Each ws In sourceWorkbook.Sheets
        ' Lấy số dòng cuối cùng của sheet hiện tại
        lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row
        
        ' Duyệt qua từng dòng trong cột N của sheet
        For i = 1 To lastRow
            ' Kiểm tra nếu giá trị cột N là "Asahi"
            If ws.Cells(i, "N").Value = "Asahi" Then
                ' Copy toàn bộ dòng vào workbook đích
                ws.Rows(i).Copy
                
                ' Tìm dòng trống đầu tiên trong workbook đích
                copyRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row + 1
                
                ' Dán dữ liệu vào workbook đích
                targetSheet.Rows(copyRow).PasteSpecial Paste:=xlPasteValues
            End If
        Next i
    Next ws

    ' Lưu workbook đích và đóng lại
    targetWorkbook.Save
    targetWorkbook.Close

    MsgBox "Đã hoàn thành việc sao chép dữ liệu.", vbInformation
End Sub
