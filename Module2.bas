Attribute VB_Name = "Module2"
Sub 通知一覧作成()

    data = 履歴抽出()

    sentence = 定型文抽出(data)
    
    bool = 通知一覧書込(data, sentence)

    If bool = True Then
        MsgBox "終了しました。", vbInformation, "OK"
    Else
        MsgBox "失敗しました。確認が必要です。", vbCritical, "エラー"
    End If

End Sub
Function 履歴抽出() As Variant()
    
    Dim data() As Variant
    
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("作成履歴")
    
    With file_name
        Dim start_point As Long
        start_point = .Cells(7, 39)
        
        start_row = .Columns(1).Find(start_point).Row
    
        max_row = .Cells(Rows.Count, 2).End(xlUp).Row
        max_col = .Cells(1, Columns.Count).End(xlToLeft).Column
        
        ReDim data(1 To max_row + start_point - 1, 1 To max_col)
        
        data = .Range(.Cells(start_row, 1), .Cells(max_row, max_col))
    End With

    履歴抽出 = data()

End Function
Function 定型文抽出(data As Variant) As Variant()

    Dim sentence() As Variant
    ReDim sentence(1 To UBound(data, 1), 1 To 3)
    
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("定型文")
    
    With file_name
        i = 1
        Do Until i = UBound(data, 1) + 1
            '1
            If data(i, 17) <> "" And data(i, 18) <> "" Then
                sentence(i, 1) = getSentence(data(i, 3), data(i, 17), data(i, 18))
            End If
            
            '2
            If data(i, 24) <> "" And data(i, 25) <> "" Then
                sentence(i, 2) = getSentence(data(i, 3), data(i, 24), data(i, 25))
            End If
            
            '3
            If data(i, 32) <> "" And data(i, 33) <> "" Then
                sentence(i, 3) = getSentence(data(i, 3), data(i, 32), data(i, 33))
            End If
            i = i + 1
        Loop
    End With
    
    定型文抽出 = sentence()

End Function
Function getSentence(year As Variant, r As Variant, c As Variant) As String
    
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("定型文")
    
    With file_name
        get_r = .Range(.Cells(3, 2), .Cells(10, 2)).Find(Left(r, 1)).Row
        get_c = .Range(.Cells(2, 3), .Cells(2, 11)).Find(Left(c, 1)).Column
        getSentence = .Cells(get_r, get_c)
        
        If get_c = 3 Or get_c = 4 Then
              separate1 = Left(getSentence, 5)
              separate2 = Mid(getSentence, 6)
              createSentence = separate1 & year - 1 & separate2
        Else
             createSentence = getSentence
        End If
        
    End With
    
    定型文判定 = createSentence
    
End Function
Function 通知一覧書込(data As Variant, sentence As Variant) As Boolean

    bool = False
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("通知一覧")
    
    With file_name
        i = 1
        Do Until i = UBound(data, 1) + 1
            .Cells(1 + i, 4) = data(i, 6) '住民コード
            .Cells(1 + i, 5) = data(i, 3) '年度
            .Cells(1 + i, 6) = data(i, 7) '名前
            .Cells(1 + i, 7) = data(i, 5) '住所
            .Cells(1 + i, 8) = data(i, 10) '指定番号
            .Cells(1 + i, 9) = data(i, 11) '事業所
            .Cells(1 + i, 10) = data(i, 36) '区分
            
            .Cells(1 + i, 14) = data(i, 12)
            .Cells(1 + i, 15) = sentence(i, 1)
            '.Cells(1 + i, 16) = data(i, 18)
            
            .Cells(1 + i, 17) = data(i, 20)
            .Cells(1 + i, 18) = sentence(i, 2)
            '.Cells(1 + i, 19) = data(i, 26)
            
            .Cells(1 + i, 20) = data(i, 28)
            .Cells(1 + i, 21) = sentence(i, 3)
            '.Cells(1 + i, 22) = data(i, 34)
            i = i + 1
        Loop
    End With
    
    bool = True
    通知一覧書込 = bool

End Function
