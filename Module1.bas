Attribute VB_Name = "Module1"
Sub 税務署提出データ抽出()

    'status = ファイル確認()
    
    'If status = False Then
        'MsgBox "対象ファイルがありません。", vbCritical, "エラー"
        'Exit Sub
    'End If
    
    index_data = 一覧ファイル抽出()
    
    detail_data = 詳細ファイル抽出()
    
    Status = ファイルペア確認(index_data, detail_data)
    
    If Status = False Then
        MsgBox "ファイルペアが異なります。", vbCritical, "エラー"
        Exit Sub
    End If
    
    param_data = パラメータ置換(detail_data)
    
    sort_data = ソート(data)
    
    print_bool = 印刷(sort_data)
    
    write_bool = 履歴書込(sort_data)
    
    If print_bool = True And write_bool = True Then
        MsgBox "終了しました。", vbInformation, "OK"
    Else
        MsgBox "失敗しました。確認が必要です。", vbCritical, "エラー"
    End If

End Sub
Function ファイル確認() As Boolean

    Dim wbn As Workbook
    Status = 0
    
    For Each wbn In Workbooks
        If wbn.Name = "一覧.csv" Then
            Status = Status + 1
        End If
        If wbn.Name = "詳細.csv" Then
            Status = Status + 1
        End If
    Next

    If Status = 2 Then
        bool = True
    Else
        bool = False
    End If
    
    ファイル確認 = bool

End Function
Function 一覧ファイル抽出() As Variant()

    Dim data() As Variant
    Dim file_name As Worksheet
    'Set file_name = Workbooks("一覧.csv").Sheets("一覧")
    Set file_name = ThisWorkbook.Sheets("一覧")
    
    With file_name
        max_row = .Cells(Rows.Count, 1).End(xlUp).Row
        max_col = .Cells(1, Columns.Count).End(xlToLeft).Column
        
        ReDim data(1 To max_row, 1 To max_col)
        
        data = .Range(.Cells(1, 1), .Cells(max_row, max_col))
    End With
    
    result = paramシートクリア()
    
    result = シート貼付(data)
    
    一覧ファイル抽出 = data()

End Function
Function 詳細ファイル抽出() As Variant()

    Dim data() As Variant
    Dim file_name As Worksheet
    'Set file_name = Workbooks("詳細.csv").Sheets("詳細")
    Set file_name = ThisWorkbook.Sheets("詳細")
    
    With file_name
        max_row = .Cells(Rows.Count, 1).End(xlUp).Row
        max_col = .Cells(1, Columns.Count).End(xlToLeft).Column
        
        ReDim data(1 To max_row, 1 To max_col)
        
        data = .Range(.Cells(1, 1), .Cells(max_row, max_col))
    End With
    
    result = シート貼付(data)
    
    詳細ファイル抽出 = data()

End Function
Function ファイルペア確認(index_data As Variant, detail_data As Variant) As Boolean
    
    index_count = UBound(index_data, 1)
    detail_count = UBound(detail_data, 1)
    
    If index_count = detail_count Then
        bool = True
    Else
        bool = False
    End If
    
    ファイルペア確認 = bool
    
End Function
Function パラメータ置換(data As Variant) As Variant()

    Dim param_data() As Variant
    ReDim param_data(1 To UBound(data, 1), 1 To 12)
    
    header_line = Array("続柄1", "区分1", "理由1", "続柄2", "区分2", "理由2", "続柄3", "区分3", "理由3", "担当", "徴収区分")
    For i = 0 To UBound(header_line, 1)
        param_data(1, i + 2) = header_line(i)
    Next

    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("param")
    
    With file_name
        For i = 1 To UBound(data, 1)
            On Error Resume Next
            param_data(i, 1) = data(i, 2) '住民コード
            param_data(i, 2) = WorksheetFunction.VLookup(data(i, 9), .Range("A:B"), 2, False) '続柄１
            param_data(i, 3) = WorksheetFunction.VLookup(data(i, 10), .Range("C:D"), 2, False) '区分１
            param_data(i, 4) = WorksheetFunction.VLookup(data(i, 11), .Range("E:F"), 2, False) '理由１
            
            param_data(i, 5) = WorksheetFunction.VLookup(data(i, 16), .Range("A:B"), 2, False) '続柄２
            param_data(i, 6) = WorksheetFunction.VLookup(data(i, 17), .Range("C:D"), 2, False) '区分２
            param_data(i, 7) = WorksheetFunction.VLookup(data(i, 18), .Range("E:F"), 2, False) '理由２
            
            param_data(i, 8) = WorksheetFunction.VLookup(data(i, 23), .Range("A:B"), 2, False) '続柄３
            param_data(i, 9) = WorksheetFunction.VLookup(data(i, 24), .Range("C:D"), 2, False) '区分３
            param_data(i, 10) = WorksheetFunction.VLookup(data(i, 25), .Range("E:F"), 2, False) '理由３
            
            param_data(i, 11) = WorksheetFunction.VLookup(Mid(data(i, 3), 7, 2), .Range("i:k"), 3, False)  '担当
            
            kubun_point = InStr(1, data(i, 4), "徴", vbTextCompare)
            param_data(i, 12) = Mid(data(i, 4), kubun_point - 1, 2) '区分
        Next
    End With
    
    result = シート貼付(param_data)
    
    パラメータ置換 = param_data()

End Function
Function paramシートクリア() As Boolean
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("sort")
    file_name.Cells.Clear
    paramシートクリア = True
End Function
Function シート貼付(data As Variant) As Boolean

    bool = False
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("sort")
    
    index1 = UBound(data, 1)
    index2 = UBound(data, 2)
    
    With file_name
        max_col = .Cells(1, Columns.Count).End(xlToLeft).Column
        
        last_col = 0
        If max_col = 1 Then
            last_col = 0
        Else
            last_col = max_col
        End If
        
        next_col = 0
        If max_col > 1 Then
            next_col = 1
        End If
        
        .Range(.Cells(1, max_col + next_col), .Cells(index1, index2 + last_col)) = data
    End With
    
    bool = True
    シート貼付 = bool
    
End Function
Function ソート(data As Variant) As Variant()

    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("sort")
    
    Dim sort_data() As Variant
    
    With file_name
        max_row = .Cells(Rows.Count, 1).End(xlUp).Row
        max_col = .Cells(1, Columns.Count).End(xlToLeft).Column
        ReDim sort_data(1 To max_row, 1 To max_col)
        .Range(.Cells(1, 1), .Cells(max_row, max_col)).Sort Key1:=.Range("AR1"), order1:=xlAscending, Header:=xlYes
        sort_data = .Range(.Cells(1, 1), .Cells(max_row, max_col))
    End With
    
    ソート = sort_data()
    
End Function
Function 印刷(data As Variant) As Boolean

    bool = False
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("様式")

    With file_name
        .Unprotect

        i = 2
        Do Until i = UBound(data, 1) + 1
            .Cells(1, 1) = "平成" & data(i, 1) '年度
            .Cells(1, 8) = data(i, 1) '年度
            .Cells(4, 1) = data(i, 12) '連絡事項
            .Cells(11, 3) = 文字化け置換(data(i, 5)) '扶養者住所
            .Cells(11, 8) = data(i, 2) '住民コード
            .Cells(13, 3) = data(i, 3) '氏名
            .Cells(12, 8) = CDate(Format(data(i, 4), "@@@@/@@/@@"))  '生年月日
            .Cells(14, 3) = data(i, 8) '支払者住所
            .Cells(14, 8) = StrConv(data(i, 6), vbNarrow) '指定番号
            .Cells(16, 3) = data(i, 7)  '支払者氏名
            
            .Cells(19, 3) = data(i, 13)  '氏名１
            .Cells(20, 8) = data(i, 16)  '合計所得１
            .Cells(19, 8) = data(i, 35)  '続柄１
            .Cells(20, 3) = data(i, 15)  '勤務先１
            .Cells(19, 6) = CDate(Format(data(i, 14), "@@@@/@@/@@"))  '生年月日１
            .Cells(21, 3) = data(i, 18) & "　" & data(i, 36) '控除区分１
            .Cells(22, 3) = data(i, 19) & "　" & data(i, 37) '否認理由１
            '.Cells(23, 3) = data(i, 54)  '備考１
            
            .Cells(24, 3) = data(i, 20)  '氏名２
            .Cells(25, 8) = data(i, 23)  '合計所得２
            .Cells(24, 8) = data(i, 38)  '続柄２
            .Cells(25, 3) = data(i, 22)  '勤務先２
            .Cells(24, 6) = CDate(Format(data(i, 21), "@@@@/@@/@@"))   '生年月日２
            .Cells(26, 3) = data(i, 25) & "　" & data(i, 39)   '控除区分２
            .Cells(27, 3) = data(i, 26) & "　" & data(i, 40)   '否認理由２
            '.Cells(28, 3) = data(i, 67)  '備考２
            
            .Cells(29, 3) = data(i, 27) '氏名３
            .Cells(30, 8) = data(i, 30)  '合計所得３
            .Cells(29, 8) = data(i, 41) '続柄３
            .Cells(30, 3) = data(i, 29) '勤務先３
            .Cells(29, 6) = CDate(Format(data(i, 28), "@@@@/@@/@@"))   '生年月日３
            .Cells(31, 3) = data(i, 32) & "　" & data(i, 42)   '控除区分３
            .Cells(32, 3) = data(i, 33) & "　" & data(i, 43)   '否認理由３
            '.Cells(33, 3) = data(i, 80)  '備考３
            '.PrintOut from:=1, to:=1
            Debug.Print i
            i = i + 1
        Loop
    
        .Protect userinterfaceonly:=True
   
    End With
    
    bool = True
    印刷 = bool

End Function
Function 履歴書込(data As Variant) As Boolean

    bool = False
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("作成履歴")
    
    max_row = file_name.Cells(Rows.Count, 2).End(xlUp).Row
    
    With file_name
        i = 2
        x = 1
        Do Until i = UBound(data, 1) + 1
            .Cells(max_row + x, 2) = Date
            .Cells(max_row + x, 3) = data(i, 1) '年度
            .Cells(max_row + x, 4) = data(i, 12) '連絡事項
            .Cells(max_row + x, 5) = 文字化け置換(data(i, 5)) '扶養者住所
            .Cells(max_row + x, 6) = data(i, 2) '住民コード
            .Cells(max_row + x, 7) = data(i, 3) '氏名
            .Cells(max_row + x, 8) = CDate(Format(data(i, 4), "@@@@/@@/@@")) '生年月日
            .Cells(max_row + x, 9) = data(i, 8) '事業所住所
            .Cells(max_row + x, 10) = StrConv(data(i, 6), vbNarrow) '指定番号
            .Cells(max_row + x, 11) = data(i, 7) '事業所名
            
            .Cells(max_row + x, 12) = data(i, 13) '氏名１
            .Cells(max_row + x, 13) = data(i, 16) '合計所得１
            .Cells(max_row + x, 14) = data(i, 35) '続柄１
            .Cells(max_row + x, 15) = data(i, 15) '勤務先１
            .Cells(max_row + x, 16) = CDate(Format(data(i, 14), "@@@@/@@/@@")) '生年月日１
            .Cells(max_row + x, 17) = data(i, 18) & "　" & data(i, 36) '控除区分１
            .Cells(max_row + x, 18) = data(i, 19) & "　" & data(i, 37) '否認理由１
            '.Cells(max_row + i, 19) = data(i, 54) '備考１
            
            .Cells(max_row + x, 20) = data(i, 20) '氏名２
            .Cells(max_row + x, 21) = data(i, 23) '合計所得２
            .Cells(max_row + x, 22) = data(i, 38) '続柄２
            .Cells(max_row + x, 23) = data(i, 22) '勤務先２
            .Cells(max_row + x, 24) = CDate(Format(data(i, 21), "@@@@/@@/@@")) '生年月日２
            .Cells(max_row + x, 25) = data(i, 25) & "　" & data(i, 39) '控除区分２
            .Cells(max_row + x, 26) = data(i, 26) & "　" & data(i, 40) '否認理由２
            '.Cells(max_row + i, 27) = data(i, 67) '備考２
            
            .Cells(max_row + x, 28) = data(i, 27) '氏名３
            .Cells(max_row + x, 29) = data(i, 30) '合計所得３
            .Cells(max_row + x, 30) = data(i, 41) '続柄３
            .Cells(max_row + x, 31) = data(i, 29) '勤務先３
            .Cells(max_row + x, 32) = CDate(Format(data(i, 28), "@@@@/@@/@@")) '生年月日３
            .Cells(max_row + x, 33) = data(i, 32) & "　" & data(i, 42) '控除区分３
            .Cells(max_row + x, 34) = data(i, 33) & "　" & data(i, 43) '否認理由３
            '.Cells(max_row + i, 35) = data(i, 80) '備考３
            .Cells(max_row + x, 36) = data(i, 45) '徴収区分
            .Cells(max_row + x, 37) = data(i, 44) '担当
            i = i + 1
            x = x + 1
        Loop
    End With
    
    bool = True
    履歴書込 = bool

End Function
Function 文字化け置換(address As Variant) As Variant

    result = address
    If Mid(address, 7, 2) = "下積" Then
        result = "下積翠" & Mid(address, 10)
    ElseIf Mid(address, 7, 2) = "上積" Then
        result = "上積翠" & Mid(address, 10)
    ElseIf Mid(address, 7, 2) = "上帯" Then
        result = "上帯那" & Mid(address, 10)
    ElseIf Mid(address, 7, 2) = "下帯" Then
        result = "下帯那" & Mid(address, 10)
    ElseIf Mid(address, 7, 2) = "千　" Then
        result = "千塚" & Mid(address, 9)
    ElseIf Mid(address, 7, 2) = "　原" Then
        result = "塚原" & Mid(address, 9)
    Else
        result = Mid(address, 7)
    End If
    
    文字化け置換 = result

End Function
