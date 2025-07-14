Sub TinhTongDoanhThuTheoPhanLoai()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sheetName As String
    Dim loaiDV As String
    Dim doanhThu As Double

    Dim totalAsahi As Double, totalGoi As Double, totalTM As Double, totalUB As Double, totalOther As Double

    ' Danh sách các sheet đặc biệt
    Dim tmSheets As Variant
    Dim ubSheets As Variant

    tmSheets = Array("TRANG NGUYEN - TELE 2", "Bich Phuong - TELE 2", "PHUONG THU - TELE 2")
    ubSheets = Array("Quynh Trang - Tele 3", "Chu Hang - Tele 3", "Tuyet Nhi - Tele 3", "Nhu Hoa - Tele 3", "Le Duyen - Tele 3")

    For Each ws In ThisWorkbook.Sheets
        sheetName = ws.Name
        lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

        ' Bỏ qua sheet không có dữ liệu
        If lastRow < 6 Then GoTo NextSheet

        ' Khởi tạo tổng
        totalTM = 0
        totalUB = 0
        totalOther = 0

        ' Duyệt từng dòng từ dòng 6
        For i = 6 To lastRow
            ' Lấy loại dịch vụ (cột V)
            If Not IsError(ws.Cells(i, 22).Value) Then
                loaiDV = Trim(CStr(ws.Cells(i, 22).Value))
            Else
                loaiDV = ""
            End If

            ' Lấy doanh thu (cột S)
            If Not IsError(ws.Cells(i, 19).Value) And IsNumeric(ws.Cells(i, 19).Value) Then
                doanhThu = CDbl(ws.Cells(i, 19).Value)
            Else
                doanhThu = 0
            End If

            ' Cộng dồn doanh thu theo phân loại
             If IsNumeric(doanhThu) Then
                Select Case loaiDV
                    Case "Tham my"
                        If IsInArray(sheetName, tmSheets) Then
                            totalTM = totalTM + doanhThu
                        Else
                            totalOther = totalOther + doanhThu
                        End If
                    Case "Ngoai - UB"
                        If IsInArray(sheetName, ubSheets) Then
                            totalUB = totalUB + doanhThu
                        Else
                            totalOther = totalOther + doanhThu
                        End If
                    Case Else
                        totalOther = totalOther + doanhThu
                End Select
            End If
        Next i

        ' Ghi kết quả ra cuối sheet
        With ws
            .Cells(lastRow + 6, 14).Value = "Tong doanh thu theo phan loai"
            .Cells(lastRow + 7, 14).Value = "Phan loai"
            .Cells(lastRow + 7, 19).Value = "Doanh thu"

            Dim r As Long: r = lastRow + 8

            If IsInArray(sheetName, tmSheets) Then
                .Cells(r, 14).Value = "Tham my"
                .Cells(r, 19).Value = totalTM: r = r + 1
            ElseIf IsInArray(sheetName, ubSheets) Then
                .Cells(r, 14).Value = "Ngoai - UB"
                .Cells(r, 19).Value = totalUB: r = r + 1
            End If

            .Cells(r, 14).Value = "Dich vu khac"
            .Cells(r, 19).Value = totalOther
        End With

NextSheet:
    Next ws

    MsgBox "Done!", vbInformation

End Sub

' Hàm kiểm tra giá trị có trong mảng không
Function IsInArray(val As String, arr As Variant) As Boolean
    Dim element
    For Each element In arr
        If val = element Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function
