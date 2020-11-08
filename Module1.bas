Attribute VB_Name = "Module1"
Sub �Ŗ�����o�f�[�^���o()

    'status = �t�@�C���m�F()
    
    'If status = False Then
        'MsgBox "�Ώۃt�@�C��������܂���B", vbCritical, "�G���["
        'Exit Sub
    'End If
    
    index_data = �ꗗ�t�@�C�����o()
    
    detail_data = �ڍ׃t�@�C�����o()
    
    Status = �t�@�C���y�A�m�F(index_data, detail_data)
    
    If Status = False Then
        MsgBox "�t�@�C���y�A���قȂ�܂��B", vbCritical, "�G���["
        Exit Sub
    End If
    
    param_data = �p�����[�^�u��(detail_data)
    
    sort_data = �\�[�g(data)
    
    print_bool = ���(sort_data)
    
    write_bool = ��������(sort_data)
    
    If print_bool = True And write_bool = True Then
        MsgBox "�I�����܂����B", vbInformation, "OK"
    Else
        MsgBox "���s���܂����B�m�F���K�v�ł��B", vbCritical, "�G���["
    End If

End Sub
Function �t�@�C���m�F() As Boolean

    Dim wbn As Workbook
    Status = 0
    
    For Each wbn In Workbooks
        If wbn.Name = "�ꗗ.csv" Then
            Status = Status + 1
        End If
        If wbn.Name = "�ڍ�.csv" Then
            Status = Status + 1
        End If
    Next

    If Status = 2 Then
        bool = True
    Else
        bool = False
    End If
    
    �t�@�C���m�F = bool

End Function
Function �ꗗ�t�@�C�����o() As Variant()

    Dim data() As Variant
    Dim file_name As Worksheet
    'Set file_name = Workbooks("�ꗗ.csv").Sheets("�ꗗ")
    Set file_name = ThisWorkbook.Sheets("�ꗗ")
    
    With file_name
        max_row = .Cells(Rows.Count, 1).End(xlUp).Row
        max_col = .Cells(1, Columns.Count).End(xlToLeft).Column
        
        ReDim data(1 To max_row, 1 To max_col)
        
        data = .Range(.Cells(1, 1), .Cells(max_row, max_col))
    End With
    
    result = param�V�[�g�N���A()
    
    result = �V�[�g�\�t(data)
    
    �ꗗ�t�@�C�����o = data()

End Function
Function �ڍ׃t�@�C�����o() As Variant()

    Dim data() As Variant
    Dim file_name As Worksheet
    'Set file_name = Workbooks("�ڍ�.csv").Sheets("�ڍ�")
    Set file_name = ThisWorkbook.Sheets("�ڍ�")
    
    With file_name
        max_row = .Cells(Rows.Count, 1).End(xlUp).Row
        max_col = .Cells(1, Columns.Count).End(xlToLeft).Column
        
        ReDim data(1 To max_row, 1 To max_col)
        
        data = .Range(.Cells(1, 1), .Cells(max_row, max_col))
    End With
    
    result = �V�[�g�\�t(data)
    
    �ڍ׃t�@�C�����o = data()

End Function
Function �t�@�C���y�A�m�F(index_data As Variant, detail_data As Variant) As Boolean
    
    index_count = UBound(index_data, 1)
    detail_count = UBound(detail_data, 1)
    
    If index_count = detail_count Then
        bool = True
    Else
        bool = False
    End If
    
    �t�@�C���y�A�m�F = bool
    
End Function
Function �p�����[�^�u��(data As Variant) As Variant()

    Dim param_data() As Variant
    ReDim param_data(1 To UBound(data, 1), 1 To 12)
    
    header_line = Array("����1", "�敪1", "���R1", "����2", "�敪2", "���R2", "����3", "�敪3", "���R3", "�S��", "�����敪")
    For i = 0 To UBound(header_line, 1)
        param_data(1, i + 2) = header_line(i)
    Next

    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("param")
    
    With file_name
        For i = 1 To UBound(data, 1)
            On Error Resume Next
            param_data(i, 1) = data(i, 2) '�Z���R�[�h
            param_data(i, 2) = WorksheetFunction.VLookup(data(i, 9), .Range("A:B"), 2, False) '�����P
            param_data(i, 3) = WorksheetFunction.VLookup(data(i, 10), .Range("C:D"), 2, False) '�敪�P
            param_data(i, 4) = WorksheetFunction.VLookup(data(i, 11), .Range("E:F"), 2, False) '���R�P
            
            param_data(i, 5) = WorksheetFunction.VLookup(data(i, 16), .Range("A:B"), 2, False) '�����Q
            param_data(i, 6) = WorksheetFunction.VLookup(data(i, 17), .Range("C:D"), 2, False) '�敪�Q
            param_data(i, 7) = WorksheetFunction.VLookup(data(i, 18), .Range("E:F"), 2, False) '���R�Q
            
            param_data(i, 8) = WorksheetFunction.VLookup(data(i, 23), .Range("A:B"), 2, False) '�����R
            param_data(i, 9) = WorksheetFunction.VLookup(data(i, 24), .Range("C:D"), 2, False) '�敪�R
            param_data(i, 10) = WorksheetFunction.VLookup(data(i, 25), .Range("E:F"), 2, False) '���R�R
            
            param_data(i, 11) = WorksheetFunction.VLookup(Mid(data(i, 3), 7, 2), .Range("i:k"), 3, False)  '�S��
            
            kubun_point = InStr(1, data(i, 4), "��", vbTextCompare)
            param_data(i, 12) = Mid(data(i, 4), kubun_point - 1, 2) '�敪
        Next
    End With
    
    result = �V�[�g�\�t(param_data)
    
    �p�����[�^�u�� = param_data()

End Function
Function param�V�[�g�N���A() As Boolean
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("sort")
    file_name.Cells.Clear
    param�V�[�g�N���A = True
End Function
Function �V�[�g�\�t(data As Variant) As Boolean

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
    �V�[�g�\�t = bool
    
End Function
Function �\�[�g(data As Variant) As Variant()

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
    
    �\�[�g = sort_data()
    
End Function
Function ���(data As Variant) As Boolean

    bool = False
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("�l��")

    With file_name
        .Unprotect

        i = 2
        Do Until i = UBound(data, 1) + 1
            .Cells(1, 1) = "����" & data(i, 1) '�N�x
            .Cells(1, 8) = data(i, 1) '�N�x
            .Cells(4, 1) = data(i, 12) '�A������
            .Cells(11, 3) = ���������u��(data(i, 5)) '�}�{�ҏZ��
            .Cells(11, 8) = data(i, 2) '�Z���R�[�h
            .Cells(13, 3) = data(i, 3) '����
            .Cells(12, 8) = CDate(Format(data(i, 4), "@@@@/@@/@@"))  '���N����
            .Cells(14, 3) = data(i, 8) '�x���ҏZ��
            .Cells(14, 8) = StrConv(data(i, 6), vbNarrow) '�w��ԍ�
            .Cells(16, 3) = data(i, 7)  '�x���Ҏ���
            
            .Cells(19, 3) = data(i, 13)  '�����P
            .Cells(20, 8) = data(i, 16)  '���v�����P
            .Cells(19, 8) = data(i, 35)  '�����P
            .Cells(20, 3) = data(i, 15)  '�Ζ���P
            .Cells(19, 6) = CDate(Format(data(i, 14), "@@@@/@@/@@"))  '���N�����P
            .Cells(21, 3) = data(i, 18) & "�@" & data(i, 36) '�T���敪�P
            .Cells(22, 3) = data(i, 19) & "�@" & data(i, 37) '�۔F���R�P
            '.Cells(23, 3) = data(i, 54)  '���l�P
            
            .Cells(24, 3) = data(i, 20)  '�����Q
            .Cells(25, 8) = data(i, 23)  '���v�����Q
            .Cells(24, 8) = data(i, 38)  '�����Q
            .Cells(25, 3) = data(i, 22)  '�Ζ���Q
            .Cells(24, 6) = CDate(Format(data(i, 21), "@@@@/@@/@@"))   '���N�����Q
            .Cells(26, 3) = data(i, 25) & "�@" & data(i, 39)   '�T���敪�Q
            .Cells(27, 3) = data(i, 26) & "�@" & data(i, 40)   '�۔F���R�Q
            '.Cells(28, 3) = data(i, 67)  '���l�Q
            
            .Cells(29, 3) = data(i, 27) '�����R
            .Cells(30, 8) = data(i, 30)  '���v�����R
            .Cells(29, 8) = data(i, 41) '�����R
            .Cells(30, 3) = data(i, 29) '�Ζ���R
            .Cells(29, 6) = CDate(Format(data(i, 28), "@@@@/@@/@@"))   '���N�����R
            .Cells(31, 3) = data(i, 32) & "�@" & data(i, 42)   '�T���敪�R
            .Cells(32, 3) = data(i, 33) & "�@" & data(i, 43)   '�۔F���R�R
            '.Cells(33, 3) = data(i, 80)  '���l�R
            '.PrintOut from:=1, to:=1
            Debug.Print i
            i = i + 1
        Loop
    
        .Protect userinterfaceonly:=True
   
    End With
    
    bool = True
    ��� = bool

End Function
Function ��������(data As Variant) As Boolean

    bool = False
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("�쐬����")
    
    max_row = file_name.Cells(Rows.Count, 2).End(xlUp).Row
    
    With file_name
        i = 2
        x = 1
        Do Until i = UBound(data, 1) + 1
            .Cells(max_row + x, 2) = Date
            .Cells(max_row + x, 3) = data(i, 1) '�N�x
            .Cells(max_row + x, 4) = data(i, 12) '�A������
            .Cells(max_row + x, 5) = ���������u��(data(i, 5)) '�}�{�ҏZ��
            .Cells(max_row + x, 6) = data(i, 2) '�Z���R�[�h
            .Cells(max_row + x, 7) = data(i, 3) '����
            .Cells(max_row + x, 8) = CDate(Format(data(i, 4), "@@@@/@@/@@")) '���N����
            .Cells(max_row + x, 9) = data(i, 8) '���Ə��Z��
            .Cells(max_row + x, 10) = StrConv(data(i, 6), vbNarrow) '�w��ԍ�
            .Cells(max_row + x, 11) = data(i, 7) '���Ə���
            
            .Cells(max_row + x, 12) = data(i, 13) '�����P
            .Cells(max_row + x, 13) = data(i, 16) '���v�����P
            .Cells(max_row + x, 14) = data(i, 35) '�����P
            .Cells(max_row + x, 15) = data(i, 15) '�Ζ���P
            .Cells(max_row + x, 16) = CDate(Format(data(i, 14), "@@@@/@@/@@")) '���N�����P
            .Cells(max_row + x, 17) = data(i, 18) & "�@" & data(i, 36) '�T���敪�P
            .Cells(max_row + x, 18) = data(i, 19) & "�@" & data(i, 37) '�۔F���R�P
            '.Cells(max_row + i, 19) = data(i, 54) '���l�P
            
            .Cells(max_row + x, 20) = data(i, 20) '�����Q
            .Cells(max_row + x, 21) = data(i, 23) '���v�����Q
            .Cells(max_row + x, 22) = data(i, 38) '�����Q
            .Cells(max_row + x, 23) = data(i, 22) '�Ζ���Q
            .Cells(max_row + x, 24) = CDate(Format(data(i, 21), "@@@@/@@/@@")) '���N�����Q
            .Cells(max_row + x, 25) = data(i, 25) & "�@" & data(i, 39) '�T���敪�Q
            .Cells(max_row + x, 26) = data(i, 26) & "�@" & data(i, 40) '�۔F���R�Q
            '.Cells(max_row + i, 27) = data(i, 67) '���l�Q
            
            .Cells(max_row + x, 28) = data(i, 27) '�����R
            .Cells(max_row + x, 29) = data(i, 30) '���v�����R
            .Cells(max_row + x, 30) = data(i, 41) '�����R
            .Cells(max_row + x, 31) = data(i, 29) '�Ζ���R
            .Cells(max_row + x, 32) = CDate(Format(data(i, 28), "@@@@/@@/@@")) '���N�����R
            .Cells(max_row + x, 33) = data(i, 32) & "�@" & data(i, 42) '�T���敪�R
            .Cells(max_row + x, 34) = data(i, 33) & "�@" & data(i, 43) '�۔F���R�R
            '.Cells(max_row + i, 35) = data(i, 80) '���l�R
            .Cells(max_row + x, 36) = data(i, 45) '�����敪
            .Cells(max_row + x, 37) = data(i, 44) '�S��
            i = i + 1
            x = x + 1
        Loop
    End With
    
    bool = True
    �������� = bool

End Function
Function ���������u��(address As Variant) As Variant

    result = address
    If Mid(address, 7, 2) = "����" Then
        result = "���ϐ�" & Mid(address, 10)
    ElseIf Mid(address, 7, 2) = "���" Then
        result = "��ϐ�" & Mid(address, 10)
    ElseIf Mid(address, 7, 2) = "���" Then
        result = "��ѓ�" & Mid(address, 10)
    ElseIf Mid(address, 7, 2) = "����" Then
        result = "���ѓ�" & Mid(address, 10)
    ElseIf Mid(address, 7, 2) = "��@" Then
        result = "���" & Mid(address, 9)
    ElseIf Mid(address, 7, 2) = "�@��" Then
        result = "�ˌ�" & Mid(address, 9)
    Else
        result = Mid(address, 7)
    End If
    
    ���������u�� = result

End Function
