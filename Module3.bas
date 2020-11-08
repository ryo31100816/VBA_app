Attribute VB_Name = "Module3"
Sub 通知作成()

    Status = MsgBox("印刷開始します", vbOKCancel, "通知作成")

    If Status = vbOK Then
    
        data = 作成データ抽出()

        bool = 印刷実行(data)
        
        bool = 印刷済フラグ()
        
        If bool = True Then
            MsgBox "終了しました。", vbInformation, "OK"
        Else
            MsgBox "失敗しました。確認が必要です。", vbCritical, "エラー"
        End If
    
        result = initialize()
    
    End If

End Sub
Function 作成データ抽出() As Variant()

    Dim data() As Variant
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("通知一覧")
    
    max_row = 500
    max_col = file_name.Cells(1, Columns.Count).End(xlToLeft).Column
    ReDim data(1 To max_row, 1 To max_col)

    With file_name
        i = 1
        For Each col In .Range(Cells(2, 4), Cells(max_row, 4)).SpecialCells(xlCellTypeVisible)
            If col = "" Then
                Exit For
            End If
            For y = 1 To max_col
                data(i, y) = .Cells(1 + i, y)
            Next
            i = i + 1
        Next
    End With
    
    作成データ抽出 = data()

End Function
Function 印刷実行(data As Variant) As Boolean
    
    'フラグ確認
    '①　特徴 or　併徴
    '②　普徴 or　併徴→普年併徴、普通徴収のみ、年金特徴のみ
    Dim i As Long
    i = 1
    Do Until data(i, 4) = ""
        If data(i, 2) = "" Then '印刷済みフラグ
            If data(i, 10) = "特徴" Or data(i, 10) = "併徴" Then  '給与特徴か普通徴収（年金特徴含む）
                result = 特徴(i, data)
            End If
              
            If data(i, 10) = "普徴" Or data(i, 10) = "併徴" Then '普通徴収か年金特徴
                If data(i, 12) <> "" And data(i, 13) <> "" Then '普年併徴
                    result = 普徴(i, data)
                    result = 年特(i, data)
                ElseIf data(i, 12) <> "" And data(i, 13) = "" Then '普通徴収のみ
                    result = 普徴(i, data)
                ElseIf data(i, 12) = "" And data(i, 13) <> "" Then '年金特徴のみ
                    result = 年特(i, data)
                End If
            End If
        End If
        i = i + 1
    Loop
    
    If result = True Then
        bool = True
    Else
        bool = False
    End If
    
    印刷実行 = bool

End Function
Function 特徴(i As Long, data As Variant) As Boolean

    bool = False
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("特徴")
    
    With file_name
        .Cells(2, 4) = data(i, 9)
        .Cells(5, 4) = data(i, 8)
        .Cells(7, 3) = data(i, 6)
        .Cells(16, 5) = data(i, 5)
        .Cells(17, 3) = data(i, 11)
        If data(i, 16) <> "" Then
            .Cells(24, 4) = data(i, 16)
        Else
            .Cells(24, 4) = data(i, 14) & "　" & data(i, 15)
        End If
        
        If data(i, 19) <> "" Then
            .Cells(30, 4) = data(i, 19)
        Else
            .Cells(30, 4) = data(i, 17) & "　" & data(i, 18)
        End If
        
        If data(i, 22) <> "" Then
            .Cells(36, 4) = data(i, 22)
        Else
            .Cells(36, 4) = data(i, 20) & data(i, 21)
        End If
        '.PrintOut from:=1, To:=1
    End With
        
    bool = True
    特徴 = bool
    
End Function
Function 普徴(i As Long, data As Variant) As Boolean

    bool = False
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("普徴")
                
    With file_name
        .Cells(2, 3) = data(i, 7)
        .Cells(7, 3) = data(i, 6)
        .Cells(16, 5) = data(i, 5)
        .Cells(17, 3) = data(i, 12)
        
        If data(i, 16) <> "" Then
            .Cells(24, 4) = data(i, 16)
        Else
            .Cells(24, 4) = data(i, 14) & "　" & data(i, 15)
        End If
    
        If data(i, 19) <> "" Then
            .Cells(30, 4) = data(i, 19)
        Else
            .Cells(30, 4) = data(i, 17) & "　" & data(i, 18)
        End If
        
        If data(i, 22) <> "" Then
            .Cells(36, 4) = data(i, 22)
        Else
            .Cells(36, 4) = data(i, 20) & "　" & data(i, 21)
        End If
            
        .Cells(40, 3) = "既に送付してあります納付書は第" & data(i, 12) - 1 & "期まで納付していただき、第" & data(i, 12) & "期以降は同封の納付書で納めてください。（すでに全納されている場合は、税額追加分を納付書で納めてください。）"
    
        .Cells(44, 3) = "口座振替の方は、第" & data(i, 12) & "期以降は同封の通知書の税額が、指定の口座から引き落としになります。（すでに全納されている場合は、税額追加分が引落としになります）"
        '.PrintOut from:=1, To:=1
    End With

    bool = True
    普徴 = bool
    
End Function
Function 年特(i As Long, data As Variant) As Boolean

    bool = False
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("年特")
                
    With file_name
        .Cells(2, 3) = data(i, 7)
        .Cells(7, 3) = data(i, 6)
        .Cells(16, 5) = data(i, 5)
        .Cells(17, 3) = data(i, 13)
            
        If data(i, 16) <> "" Then
            .Cells(24, 4) = data(i, 16)
        Else
            .Cells(24, 4) = data(i, 14) & "　" & data(i, 15)
        End If
        
        If data(i, 19) <> "" Then
            .Cells(30, 4) = data(i, 19)
        Else
            .Cells(30, 4) = data(i, 17) & "　" & data(i, 18)
        End If
        
        If data(i, 22) <> "" Then
            .Cells(36, 4) = data(i, 22)
        Else
            .Cells(36, 4) = data(i, 20) & "　" & data(i, 21)
        End If
    
        '.PrintOut from:=1, To:=1
    End With

    bool = True
    年特 = bool
    
End Function
Function initialize() As Boolean

    bool = False
    Dim file_name As Worksheet

    Set file_name = ThisWorkbook.Sheets("特徴")
    With file_name
        .Cells(2, 4).Formula = "=VLookup(M4,通知一覧!A:V,9,FALSE)"
        .Cells(5, 4).Formula = "=VLookup(M4,通知一覧!A:V,8,FALSE)"
        .Cells(7, 3).Formula = "=VLookup(M4,通知一覧!A:V,6,FALSE)"
        .Cells(16, 5).Formula = "=VLookup(M4,通知一覧!A:V,5,FALSE)"
        .Cells(17, 3).Formula = "=VLookup(M4,通知一覧!A:V,11,FALSE)"
        .Cells(24, 4).Formula = "=IF(OFFSET(通知一覧!A1,M4,15)<>0,VLOOKUP(M4,通知一覧!A:V,16,FALSE),VLOOKUP(M4,通知一覧!A:V,14,FALSE)& ""　""&VLOOKUP(M4,通知一覧!A:V,15,FALSE))&"""""
        .Cells(30, 4).Formula = "=IF(OFFSET(通知一覧!A1,M4,18)<>0,VLOOKUP(M4,通知一覧!A:V,19,FALSE),VLOOKUP(M4,通知一覧!A:V,17,FALSE)& ""　""&VLOOKUP(M4,通知一覧!A:V,18,FALSE))&"""""
        .Cells(36, 4).Formula = "=IF(OFFSET(通知一覧!A1,M4,21)<>0,VLOOKUP(M4,通知一覧!A:V,22,FALSE),VLOOKUP(M4,通知一覧!A:V,20,FALSE)& ""　""&VLOOKUP(M4,通知一覧!A:V,21,FALSE))&"""""
    End With
    
    Set file_name = ThisWorkbook.Sheets("普徴")
    With file_name
        .Cells(2, 3).Formula = "=VLookup(M4,通知一覧!A:V,7,FALSE)"
        .Cells(7, 3).Formula = "=VLookup(M4,通知一覧!A:V,6,FALSE)"
        .Cells(16, 5).Formula = "=VLookup(M4,通知一覧!A:V,5,FALSE)"
        .Cells(17, 3).Formula = "=VLookup(M4,通知一覧!A:V,12,FALSE)"
        .Cells(24, 4).Formula = "=IF(OFFSET(通知一覧!A1,M4,15)<>0,VLOOKUP(M4,通知一覧!A:V,16,FALSE),VLOOKUP(M4,通知一覧!A:V,14,FALSE)& ""　""&VLOOKUP(M4,通知一覧!A:V,15,FALSE))&"""""
        .Cells(30, 4).Formula = "=IF(OFFSET(通知一覧!A1,M4,18)<>0,VLOOKUP(M4,通知一覧!A:V,19,FALSE),VLOOKUP(M4,通知一覧!A:V,17,FALSE)& ""　""&VLOOKUP(M4,通知一覧!A:V,18,FALSE))&"""""
        .Cells(36, 4).Formula = "=IF(OFFSET(通知一覧!A1,M4,21)<>0,VLOOKUP(M4,通知一覧!A:V,22,FALSE),VLOOKUP(M4,通知一覧!A:V,20,FALSE)& ""　""&VLOOKUP(M4,通知一覧!A:V,21,FALSE))&"""""
    End With
    
    Set file_name = ThisWorkbook.Sheets("年特")
    With file_name
        .Cells(2, 3).Formula = "=VLookup(M4,通知一覧!A:V,7,FALSE)"
        .Cells(7, 3).Formula = "=VLookup(M4,通知一覧!A:V,6,FALSE)"
        .Cells(16, 5).Formula = "=VLookup(M4,通知一覧!A:V,5,FALSE)"
        .Cells(17, 3).Formula = "=VLookup(M4,通知一覧!A:V,12,FALSE)"
        .Cells(24, 4).Formula = "=IF(OFFSET(通知一覧!A1,M4,15)<>0,VLOOKUP(M4,通知一覧!A:V,16,FALSE),VLOOKUP(M4,通知一覧!A:V,14,FALSE)& ""　""&VLOOKUP(M4,通知一覧!A:V,15,FALSE))&"""""
        .Cells(30, 4).Formula = "=IF(OFFSET(通知一覧!A1,M4,18)<>0,VLOOKUP(M4,通知一覧!A:V,19,FALSE),VLOOKUP(M4,通知一覧!A:V,17,FALSE)& ""　""&VLOOKUP(M4,通知一覧!A:V,18,FALSE))&"""""
        .Cells(36, 4).Formula = "=IF(OFFSET(通知一覧!A1,M4,21)<>0,VLOOKUP(M4,通知一覧!A:V,22,FALSE),VLOOKUP(M4,通知一覧!A:V,20,FALSE)& ""　""&VLOOKUP(M4,通知一覧!A:V,21,FALSE))&"""""
    End With
    
    bool = True
    initialize = bool

End Function
Function 印刷済フラグ() As Boolean

    bool = False
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("通知一覧")
    
    max_row = 500

    With file_name
        For Each col In .Range(Cells(2, 4), Cells(max_row, 4)).SpecialCells(xlCellTypeVisible)
            If col = "" Then
                Exit For
            End If
            If .Cells(col.Row, 2) = "" Then
                .Cells(col.Row, 2) = "○"
            End If
        Next
    End With
        
    bool = True
    印刷済フラグ = bool
    
End Function
