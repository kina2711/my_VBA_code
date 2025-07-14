Sub XoaCotVTrongTatCaSheet()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Sheets
        On Error Resume Next ' Phòng tránh lỗi nếu cột không tồn tại
        ws.Columns(22).Delete
        On Error GoTo 0 ' Bật lại thông báo lỗi sau khi thử xóa
    Next ws

    MsgBox "Đã xóa cột V (cột số 22) trong tất cả các sheet!", vbInformation

End Sub
