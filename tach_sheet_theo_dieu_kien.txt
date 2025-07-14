Sub SplitSheetsByTail1()
    Dim ws As Worksheet
    Dim wbNew As Workbook
    Dim savePath As String

    ' Tạo file mới để chứa các sheet cần tách
    Set wbNew = Workbooks.Add

    ' Kiểm tra nếu có nhiều hơn 1 sheet mặc định trong file mới thì xóa bớt
    Application.DisplayAlerts = False
    Do While wbNew.Worksheets.Count > 1
        wbNew.Worksheets(1).Delete
    Loop
    Application.DisplayAlerts = True

    ' Duyệt qua các sheet trong workbook hiện tại
    Dim sheetCount As Long
    sheetCount = 0

    For Each ws In ThisWorkbook.Worksheets
        If Right(ws.Name, 1) = "1" Then
            sheetCount = sheetCount + 1
            ws.Copy After:=wbNew.Sheets(wbNew.Sheets.Count)
        End If
    Next ws

    ' Kiểm tra nếu có sheet cần tách
    If sheetCount = 0 Then
        MsgBox "Không có sheet nào có đuôi là số 1!", vbExclamation, "Thông báo"
        wbNew.Close SaveChanges:=False
        Exit Sub
    End If

    ' Lưu file mới
    savePath = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", _
                                             Title:="Chọn nơi lưu file đã tách")

    If savePath <> "False" Then
        wbNew.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
        MsgBox "Đã tách thành công các sheet có đuôi là số 1 vào file mới!", vbInformation, "Hoàn tất"
    Else
        wbNew.Close SaveChanges:=False
        MsgBox "Đã hủy quá trình lưu file.", vbExclamation, "Hủy bỏ"
    End If
End Sub
