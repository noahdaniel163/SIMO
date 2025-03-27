Sub ProcessExcelData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    
    ' Thiết lập worksheet (chọn Sheet1, bạn có thể thay bằng tên sheet cụ thể)
    Set ws = ThisWorkbook.Sheets(1) ' Hoặc Set ws = ThisWorkbook.Sheets("SheetName")
    
    ' Tìm dòng cuối cùng có dữ liệu ở cột B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Lặp qua các dòng từ dòng 5 đến dòng cuối
    For i = 5 To lastRow
        ' Xử lý cột B
        If Not IsEmpty(ws.Cells(i, "B")) Then
            cellValue = ws.Cells(i, "B").Value
            ws.Cells(i, "B").Value = ReplaceThirdChar(cellValue)
        End If
        
        ' Xử lý cột C
        If Not IsEmpty(ws.Cells(i, "C")) Then
            cellValue = ws.Cells(i, "C").Value
            ws.Cells(i, "C").Value = ReplaceThirdChar(cellValue)
        End If
        
        ' Xử lý cột E
        If Not IsEmpty(ws.Cells(i, "E")) Then
            cellValue = ws.Cells(i, "E").Value
            ws.Cells(i, "E").Value = ReplaceThirdChar(cellValue)
        End If
        
        ' Xử lý cột I
        If Not IsEmpty(ws.Cells(i, "I")) Then
            cellValue = ws.Cells(i, "I").Value
            ws.Cells(i, "I").Value = ReplaceThirdChar(cellValue)
        End If
        
        ' Xử lý cột J
        If Not IsEmpty(ws.Cells(i, "J")) Then
            cellValue = ws.Cells(i, "J").Value
            ws.Cells(i, "J").Value = ReplaceThirdChar(cellValue)
        End If
        
        ' Xử lý cột N
        If Not IsEmpty(ws.Cells(i, "N")) Then
            cellValue = ws.Cells(i, "N").Value
            ws.Cells(i, "N").Value = ReplaceThirdChar(cellValue)
        End If
    Next i
    
    MsgBox "Đã xử lý xong!", vbInformation
End Sub

Private Function ReplaceThirdChar(value As String) As String
    ' Nếu chuỗi ngắn hơn 3 ký tự, trả về nguyên bản
    If Len(value) < 3 Then
        ReplaceThirdChar = value
        Exit Function
    End If
    
    ' Kiểm tra xem chuỗi có phải là số không
    If IsNumeric(value) Then
        ' Với số: thay bằng 000 + phần còn lại từ ký tự thứ 4
        ReplaceThirdChar = "000" & Mid(value, 4)
    Else
        ' Với text: thay ký tự thứ 3 bằng X
        ReplaceThirdChar = Left(value, 2) & "X" & Mid(value, 4)
    End If
End Function