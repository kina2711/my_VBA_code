Sub InDanhSachSheet()
    Dim ws As Worksheet
    Dim summarySheet As Worksheet
    Dim i As Long

    ' Tạo hoặc xóa sheet tổng hợp nếu đã tồn tại
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets("DANH SÁCH SHEET").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set summarySheet = Worksheets.Add
    summarySheet.Name = "DANH SÁCH SHEET"

    summarySheet.Cells(1, 1).Value = "Tên sheet trong workbook:"
    i = 2
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "DANH SÁCH SHEET" Then
            summarySheet.Cells(i, 1).Value = ws.Name
            i = i + 1
        End If
    Next ws

    MsgBox "✅ Đã liệt kê toàn bộ tên sheet!", vbInformation
End Sub
