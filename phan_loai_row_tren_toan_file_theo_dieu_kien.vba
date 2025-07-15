Sub PhanLoaiDichVu_Final()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim valB As String, valE As String, valF As String, valN As String
    Dim result As String
    Dim sheetName As String

    ' Chuỗi tiếng Việt viết bằng ChrW (không dùng LCase)
    Dim tm As String
    tm = ChrW(84) & ChrW(104) & ChrW(7849) & ChrW(109) & " " & ChrW(109) & ChrW(7929) ' Thẩm mỹ

    Dim ng As String
    ng = ChrW(78) & ChrW(103) & ChrW(111) & ChrW(7841) & ChrW(105) ' Ngoại

    ' Tên các sheet Thẩm mỹ (không dấu)
    Dim tmSheets As Variant
    tmSheets = Array("TRANG NGUYEN - TELE 2", "Bich Phuong - TELE 2", "PHUONG THU - TELE 2")

    ' Tên các sheet Ngoại - UB (không dấu)
    Dim ubSheets As Variant
    ubSheets = Array("Quynh Trang - Tele 3", "Chu Hang - Tele 3", "Tuyet Nhi - Tele 3", "Nhu Hoa - Tele 3", "Le Duyen - Tele 3")

    ' Duyệt tổng sheet
    For Each ws In ThisWorkbook.Worksheets
        sheetName = ws.Name
        With ws
            lastRow = .Cells(.Rows.Count, 2).End(xlUp).Row

            For i = 5 To lastRow
                valB = Trim(.Cells(i, 2).Value)
                If valB <> "" Then
                    valE = Trim(.Cells(i, 5).Value)
                    valF = Trim(.Cells(i, 6).Value)
                    valN = Trim(.Cells(i, 14).Value)

                    result = "Dich vu khac" ' Mặc định

                    ' 1. Sheet Thẩm mỹ
                    If IsInArray(sheetName, tmSheets) Then
                        If valE = tm Then
                            result = "Tham my"
                        End If
                    ' 2. Sheet Ngoại - UB
                    ElseIf IsInArray(sheetName, ubSheets) Then
                        If valE = ng Or InStr(valE, "Ung ") > 0 Then
                            result = "Ngoai - UB"
                        End If
                    End If

                    ' Ghi kết quả vào cột V
                    .Cells(i, 22).Value = result
                End If
            Next i
        End With
    Next ws

    MsgBox "Done!", vbInformation

End Sub

' Hàm phụ: kiểm tra tên sheet có trong danh sách không
Function IsInArray(val As String, arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If val = element Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function
