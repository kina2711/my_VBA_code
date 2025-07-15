Sub Update_DoanhThu()
    Dim wb1 As Workbook, wb2 As Workbook
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long
    Dim dict As Object
    Dim key As String
    Dim filePath As String

    ' Mở File 1 (Tổng đài)
    Set wb1 = ThisWorkbook

    ' Chọn File 2 (Doanh số Asahi)
    filePath = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Chọn file Doanh số Asahi")
    If filePath = "False" Then Exit Sub
    Set wb2 = Workbooks.Open(filePath)

    ' Lấy Sheet1 của File 2
    Set ws2 = wb2.Sheets("Sheet1")

    ' Tạo Dictionary để lưu dữ liệu File 2
    Set dict = CreateObject("Scripting.Dictionary")

    ' Xác định số dòng cuối trong Sheet1 của File 2
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row

    ' Lưu dữ liệu từ File 2 vào Dictionary (Bắt đầu từ hàng 2)
    Dim i As Long
    For i = 2 To lastRow2
        ' Thêm Cột E (cột số 5) vào key
        key = ws2.Cells(i, 1).Value & "_" & ws2.Cells(i, 3).Value & "_" & ws2.Cells(i, 5).Value & "_" & ws2.Cells(i, 14).Value & "_" & ws2.Cells(i, 7).Value & ws2.Cells(i, 6).Value
        dict(key) = ws2.Cells(i, 9).Value ' Cột I - Kế toán Asahi
    Next i

    ' Đóng File 2 sau khi lấy dữ liệu
    wb2.Close False

    ' Duyệt qua từng sheet trong File 1
    For Each ws1 In wb1.Sheets
        lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row

        If lastRow1 < 6 Then GoTo NextSheet

        For i = 6 To lastRow1
            If ws1.Cells(i, 14).Value <> "Asahi" Then GoTo NextRow

            ' Tạo key đối chiếu: Mã y tế + SĐT + Cột G + Ngày + Cột E
            key = ws1.Cells(i, 1).Value & "_" & ws1.Cells(i, 3).Value & "_" & ws1.Cells(i, 5).Value & "_" & ws1.Cells(i, 12).Value & "_" & ws1.Cells(i, 7).Value & ws1.Cells(i, 6).Value

            If dict.exists(key) Then
                ws1.Cells(i, 19).Value = dict(key)
            End If

NextRow:
        Next i

NextSheet:
    Next ws1

    MsgBox "Cập nhật doanh thu hoàn tất!", vbInformation
End Sub
