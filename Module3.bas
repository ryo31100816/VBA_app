Attribute VB_Name = "Module3"
Sub �ʒm�쐬()

    Status = MsgBox("����J�n���܂�", vbOKCancel, "�ʒm�쐬")

    If Status = vbOK Then
    
        data = �쐬�f�[�^���o()

        bool = ������s(data)
        
        bool = ����σt���O()
        
        If bool = True Then
            MsgBox "�I�����܂����B", vbInformation, "OK"
        Else
            MsgBox "���s���܂����B�m�F���K�v�ł��B", vbCritical, "�G���["
        End If
    
        result = initialize()
    
    End If

End Sub
Function �쐬�f�[�^���o() As Variant()

    Dim data() As Variant
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("�ʒm�ꗗ")
    
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
    
    �쐬�f�[�^���o = data()

End Function
Function ������s(data As Variant) As Boolean
    
    '�t���O�m�F
    '�@�@���� or�@����
    '�A�@���� or�@���������N�����A���ʒ����̂݁A�N�������̂�
    Dim i As Long
    i = 1
    Do Until data(i, 4) = ""
        If data(i, 2) = "" Then '����ς݃t���O
            If data(i, 10) = "����" Or data(i, 10) = "����" Then  '���^���������ʒ����i�N�������܂ށj
                result = ����(i, data)
            End If
              
            If data(i, 10) = "����" Or data(i, 10) = "����" Then '���ʒ������N������
                If data(i, 12) <> "" And data(i, 13) <> "" Then '���N����
                    result = ����(i, data)
                    result = �N��(i, data)
                ElseIf data(i, 12) <> "" And data(i, 13) = "" Then '���ʒ����̂�
                    result = ����(i, data)
                ElseIf data(i, 12) = "" And data(i, 13) <> "" Then '�N�������̂�
                    result = �N��(i, data)
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
    
    ������s = bool

End Function
Function ����(i As Long, data As Variant) As Boolean

    bool = False
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("����")
    
    With file_name
        .Cells(2, 4) = data(i, 9)
        .Cells(5, 4) = data(i, 8)
        .Cells(7, 3) = data(i, 6)
        .Cells(16, 5) = data(i, 5)
        .Cells(17, 3) = data(i, 11)
        If data(i, 16) <> "" Then
            .Cells(24, 4) = data(i, 16)
        Else
            .Cells(24, 4) = data(i, 14) & "�@" & data(i, 15)
        End If
        
        If data(i, 19) <> "" Then
            .Cells(30, 4) = data(i, 19)
        Else
            .Cells(30, 4) = data(i, 17) & "�@" & data(i, 18)
        End If
        
        If data(i, 22) <> "" Then
            .Cells(36, 4) = data(i, 22)
        Else
            .Cells(36, 4) = data(i, 20) & data(i, 21)
        End If
        '.PrintOut from:=1, To:=1
    End With
        
    bool = True
    ���� = bool
    
End Function
Function ����(i As Long, data As Variant) As Boolean

    bool = False
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("����")
                
    With file_name
        .Cells(2, 3) = data(i, 7)
        .Cells(7, 3) = data(i, 6)
        .Cells(16, 5) = data(i, 5)
        .Cells(17, 3) = data(i, 12)
        
        If data(i, 16) <> "" Then
            .Cells(24, 4) = data(i, 16)
        Else
            .Cells(24, 4) = data(i, 14) & "�@" & data(i, 15)
        End If
    
        If data(i, 19) <> "" Then
            .Cells(30, 4) = data(i, 19)
        Else
            .Cells(30, 4) = data(i, 17) & "�@" & data(i, 18)
        End If
        
        If data(i, 22) <> "" Then
            .Cells(36, 4) = data(i, 22)
        Else
            .Cells(36, 4) = data(i, 20) & "�@" & data(i, 21)
        End If
            
        .Cells(40, 3) = "���ɑ��t���Ă���܂��[�t���͑�" & data(i, 12) - 1 & "���܂Ŕ[�t���Ă��������A��" & data(i, 12) & "���ȍ~�͓����̔[�t���Ŕ[�߂Ă��������B�i���łɑS�[����Ă���ꍇ�́A�Ŋz�ǉ�����[�t���Ŕ[�߂Ă��������B�j"
    
        .Cells(44, 3) = "�����U�ւ̕��́A��" & data(i, 12) & "���ȍ~�͓����̒ʒm���̐Ŋz���A�w��̌�������������Ƃ��ɂȂ�܂��B�i���łɑS�[����Ă���ꍇ�́A�Ŋz�ǉ����������Ƃ��ɂȂ�܂��j"
        '.PrintOut from:=1, To:=1
    End With

    bool = True
    ���� = bool
    
End Function
Function �N��(i As Long, data As Variant) As Boolean

    bool = False
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("�N��")
                
    With file_name
        .Cells(2, 3) = data(i, 7)
        .Cells(7, 3) = data(i, 6)
        .Cells(16, 5) = data(i, 5)
        .Cells(17, 3) = data(i, 13)
            
        If data(i, 16) <> "" Then
            .Cells(24, 4) = data(i, 16)
        Else
            .Cells(24, 4) = data(i, 14) & "�@" & data(i, 15)
        End If
        
        If data(i, 19) <> "" Then
            .Cells(30, 4) = data(i, 19)
        Else
            .Cells(30, 4) = data(i, 17) & "�@" & data(i, 18)
        End If
        
        If data(i, 22) <> "" Then
            .Cells(36, 4) = data(i, 22)
        Else
            .Cells(36, 4) = data(i, 20) & "�@" & data(i, 21)
        End If
    
        '.PrintOut from:=1, To:=1
    End With

    bool = True
    �N�� = bool
    
End Function
Function initialize() As Boolean

    bool = False
    Dim file_name As Worksheet

    Set file_name = ThisWorkbook.Sheets("����")
    With file_name
        .Cells(2, 4).Formula = "=VLookup(M4,�ʒm�ꗗ!A:V,9,FALSE)"
        .Cells(5, 4).Formula = "=VLookup(M4,�ʒm�ꗗ!A:V,8,FALSE)"
        .Cells(7, 3).Formula = "=VLookup(M4,�ʒm�ꗗ!A:V,6,FALSE)"
        .Cells(16, 5).Formula = "=VLookup(M4,�ʒm�ꗗ!A:V,5,FALSE)"
        .Cells(17, 3).Formula = "=VLookup(M4,�ʒm�ꗗ!A:V,11,FALSE)"
        .Cells(24, 4).Formula = "=IF(OFFSET(�ʒm�ꗗ!A1,M4,15)<>0,VLOOKUP(M4,�ʒm�ꗗ!A:V,16,FALSE),VLOOKUP(M4,�ʒm�ꗗ!A:V,14,FALSE)& ""�@""&VLOOKUP(M4,�ʒm�ꗗ!A:V,15,FALSE))&"""""
        .Cells(30, 4).Formula = "=IF(OFFSET(�ʒm�ꗗ!A1,M4,18)<>0,VLOOKUP(M4,�ʒm�ꗗ!A:V,19,FALSE),VLOOKUP(M4,�ʒm�ꗗ!A:V,17,FALSE)& ""�@""&VLOOKUP(M4,�ʒm�ꗗ!A:V,18,FALSE))&"""""
        .Cells(36, 4).Formula = "=IF(OFFSET(�ʒm�ꗗ!A1,M4,21)<>0,VLOOKUP(M4,�ʒm�ꗗ!A:V,22,FALSE),VLOOKUP(M4,�ʒm�ꗗ!A:V,20,FALSE)& ""�@""&VLOOKUP(M4,�ʒm�ꗗ!A:V,21,FALSE))&"""""
    End With
    
    Set file_name = ThisWorkbook.Sheets("����")
    With file_name
        .Cells(2, 3).Formula = "=VLookup(M4,�ʒm�ꗗ!A:V,7,FALSE)"
        .Cells(7, 3).Formula = "=VLookup(M4,�ʒm�ꗗ!A:V,6,FALSE)"
        .Cells(16, 5).Formula = "=VLookup(M4,�ʒm�ꗗ!A:V,5,FALSE)"
        .Cells(17, 3).Formula = "=VLookup(M4,�ʒm�ꗗ!A:V,12,FALSE)"
        .Cells(24, 4).Formula = "=IF(OFFSET(�ʒm�ꗗ!A1,M4,15)<>0,VLOOKUP(M4,�ʒm�ꗗ!A:V,16,FALSE),VLOOKUP(M4,�ʒm�ꗗ!A:V,14,FALSE)& ""�@""&VLOOKUP(M4,�ʒm�ꗗ!A:V,15,FALSE))&"""""
        .Cells(30, 4).Formula = "=IF(OFFSET(�ʒm�ꗗ!A1,M4,18)<>0,VLOOKUP(M4,�ʒm�ꗗ!A:V,19,FALSE),VLOOKUP(M4,�ʒm�ꗗ!A:V,17,FALSE)& ""�@""&VLOOKUP(M4,�ʒm�ꗗ!A:V,18,FALSE))&"""""
        .Cells(36, 4).Formula = "=IF(OFFSET(�ʒm�ꗗ!A1,M4,21)<>0,VLOOKUP(M4,�ʒm�ꗗ!A:V,22,FALSE),VLOOKUP(M4,�ʒm�ꗗ!A:V,20,FALSE)& ""�@""&VLOOKUP(M4,�ʒm�ꗗ!A:V,21,FALSE))&"""""
    End With
    
    Set file_name = ThisWorkbook.Sheets("�N��")
    With file_name
        .Cells(2, 3).Formula = "=VLookup(M4,�ʒm�ꗗ!A:V,7,FALSE)"
        .Cells(7, 3).Formula = "=VLookup(M4,�ʒm�ꗗ!A:V,6,FALSE)"
        .Cells(16, 5).Formula = "=VLookup(M4,�ʒm�ꗗ!A:V,5,FALSE)"
        .Cells(17, 3).Formula = "=VLookup(M4,�ʒm�ꗗ!A:V,12,FALSE)"
        .Cells(24, 4).Formula = "=IF(OFFSET(�ʒm�ꗗ!A1,M4,15)<>0,VLOOKUP(M4,�ʒm�ꗗ!A:V,16,FALSE),VLOOKUP(M4,�ʒm�ꗗ!A:V,14,FALSE)& ""�@""&VLOOKUP(M4,�ʒm�ꗗ!A:V,15,FALSE))&"""""
        .Cells(30, 4).Formula = "=IF(OFFSET(�ʒm�ꗗ!A1,M4,18)<>0,VLOOKUP(M4,�ʒm�ꗗ!A:V,19,FALSE),VLOOKUP(M4,�ʒm�ꗗ!A:V,17,FALSE)& ""�@""&VLOOKUP(M4,�ʒm�ꗗ!A:V,18,FALSE))&"""""
        .Cells(36, 4).Formula = "=IF(OFFSET(�ʒm�ꗗ!A1,M4,21)<>0,VLOOKUP(M4,�ʒm�ꗗ!A:V,22,FALSE),VLOOKUP(M4,�ʒm�ꗗ!A:V,20,FALSE)& ""�@""&VLOOKUP(M4,�ʒm�ꗗ!A:V,21,FALSE))&"""""
    End With
    
    bool = True
    initialize = bool

End Function
Function ����σt���O() As Boolean

    bool = False
    Dim file_name As Worksheet
    Set file_name = ThisWorkbook.Sheets("�ʒm�ꗗ")
    
    max_row = 500

    With file_name
        For Each col In .Range(Cells(2, 4), Cells(max_row, 4)).SpecialCells(xlCellTypeVisible)
            If col = "" Then
                Exit For
            End If
            If .Cells(col.Row, 2) = "" Then
                .Cells(col.Row, 2) = "��"
            End If
        Next
    End With
        
    bool = True
    ����σt���O = bool
    
End Function
