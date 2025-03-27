Sub ExportDataToTxt_UTF8()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range, cell As Range
    Dim filePath As String
    Dim rowData As String
    Dim i As Long, j As Long
    Dim stream As Object

    ' Lấy sheet hiện tại
    Set ws = ActiveSheet

    ' Xác định dòng và cột cuối cùng có dữ liệu
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column

    ' Đường dẫn file xuất ra ở E
    filePath = "E:\ExportedData_UTF8.txt"

    ' Tạo đối tượng ADODB.Stream để ghi file UTF-8
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Chỉ để ghi text
    stream.Charset = "utf-8" ' Định dạng UTF-8
    stream.Open

    ' Lặp qua từng dòng để ghi dữ liệu
    For i = 1 To lastRow
        rowData = "|" ' Thêm dấu | ở đầu dòng
        For j = 1 To lastCol
            rowData = rowData & ws.Cells(i, j).Value & "|"
        Next j
        ' Xóa dấu "|" cuối cùng
        rowData = Left(rowData, Len(rowData) - 1)
        
        ' Ghi dòng vào stream
        stream.WriteText rowData & vbCrLf
    Next i

    ' Lưu vào file
    stream.SaveToFile filePath, 2 ' 2 = Overwrite nếu file đã tồn tại
    stream.Close

    ' Giải phóng bộ nhớ
    Set stream = Nothing

    ' Thông báo hoàn tất
    MsgBox "Xuất dữ liệu thành công ra: " & filePath, vbInformation, "Hoàn tất"

End Sub