Attribute VB_Name = "����"
Option Explicit

Dim ��ĸ�� As Variant
Dim ����� As Variant
Dim ������ As Variant
Dim ��ĸ�� As Variant
Dim ������ As Variant
Dim ������ As Variant

Function ��ĸ��(��ĸ As String) As String
    ��ĸ�� = Left(��ĸ, 1)
    If InStr("���R��֮�~��ģ�҅�Ձ��������ۺ���ʒ���Ⱥ���������������Մ�}����㕇���", ��ĸ��) Then
        ' �޵��޺�
        If InStr("�����}", ��ĸ��) Then ��ĸ�� = Left(��ĸ, 2)
        If ��ĸ = "������" Then ��ĸ�� = "��"
    ElseIf InStr("֧֬΢�R���U�ѽԉ���Ԫ�hɽ̩������������", ��ĸ��) Then
        ' �޵��к�
        If InStr("֧֬������", ��ĸ��) Then ��ĸ�� = Left(��ĸ, 2)
        ��ĸ�� = ��ĸ�� + Right(��ĸ, 1)
        If ��ĸ = "��A���_" Then ��ĸ�� = "��A"
    ElseIf InStr("�|", ��ĸ��) Then
        ' �е��޺�
        ��ĸ�� = Left(��ĸ, 2)
    Else '�е��к������������
        ��ĸ�� = ��ĸ
    End If
End Function

Function ��ĸת��(��ĸ As String) As String
End Function

Sub �������()
    ��ĸ�� = Array("�|һ", "�|��", "��", "�R", "��", "֧�_", "֧��", "֬�_", "֬��", "֮", "΢�_", "΢��", "�~", _
        "��", "ģ", "�R�_", "�R��", "���_", "����", "�U�_", "�U��", "���_", "�Ѻ�", "���_", "�Ժ�", "���_", "����", _
        "��", "��", "���_", "���", "Ձ", "��", "��", "��", "Ԫ�_", "Ԫ��", "��", "��", "��", "��", "�h�_", "�h��", _
        "ɽ�_", "ɽ��", "̩�_", "̩��", "���_", "�Ⱥ�", "���_", "�ɺ�", "ʒ", "��", "��", "��", "��", "��һ��", _
        "�����_", "������", "����_", "�����", "�����_", "ꖺ�", "��_", "���_", "�ƺ�", "�����_", "������", "�����_", _
        "������", "�����_", "������", "���_", "���", "���_", "���", "��", "��", "���_", "�Ǻ�", "��", "��", "��", _
        "��", "��", "Մ", "�}", "��", "��", "�", "��", "��")
    ����� = Array("u" & ChrW(&H14B), "iu" & ChrW(&H14B), "o" & ChrW(&H14B), "io" & ChrW(&H14B), ChrW(&H254) & _
        ChrW(&H14B), "ie", "uie", "i", "ui", "i" & ChrW(&H259), "i" & ChrW(&H259) & "i", "iu" & ChrW(&H259) & "i", "i" _
        & ChrW(&H254), "io", "o", "ei", "uei", "i" & ChrW(&H25B), "iu" & ChrW(&H25B), "i" & ChrW(&H250), "iu" & _
        ChrW(&H250), ChrW(&H25B), "u" & ChrW(&H25B), ChrW(&H25B) & "i", "u" & ChrW(&H25B) & "i", "ai", "uai", "u" & _
        ChrW(&H252) & "i", ChrW(&H252) & "i", "i" & ChrW(&H26A) & "n", "iu" & ChrW(&H26A) & "n", "uien", "ien", "iu" & _
        ChrW(&H259) & "n", "i" & ChrW(&H259) & "n", "i" & ChrW(&H250) & "n", "iu" & ChrW(&H250) & "n", "u" & ChrW(&H259) _
        & "n", ChrW(&H259) & "n", ChrW(&H251) & "n", "u" & ChrW(&H251) & "n", "an", "uan", ChrW(&H25B) & "n", "u" & _
        ChrW(&H25B) & "n", ChrW(&H251) & "i", "u" & ChrW(&H251) & "i", "en", "uen", "i" & ChrW(&H25B) & "n", "iu" & _
        ChrW(&H25B) & "n", "eu", "i" & ChrW(&H25B) & "u", "au", ChrW(&H251) & "u", ChrW(&H251), "u" & ChrW(&H251), "i" _
        & ChrW(&H251), "iu" & ChrW(&H251), "a", "ua", "ia", "ia" & ChrW(&H14B), "iua" & ChrW(&H14B), ChrW(&H251) & _
        ChrW(&H14B), "u" & ChrW(&H251) & ChrW(&H14B), ChrW(&H250) & ChrW(&H14B), "u" & ChrW(&H250) & ChrW(&H14B), "i" & _
        ChrW(&H250) & ChrW(&H14B), "iu" & ChrW(&H250) & ChrW(&H14B), ChrW(&H25B) & ChrW(&H14B), "u" & ChrW(&H25B) & _
        ChrW(&H14B), "i" & ChrW(&H25B) & ChrW(&H14B), "iu" & ChrW(&H25B) & ChrW(&H14B), "e" & ChrW(&H14B), "ue" & _
        ChrW(&H14B), "i" & ChrW(&H259) & ChrW(&H14B), "iu" & ChrW(&H259) & ChrW(&H14B), ChrW(&H259) & ChrW(&H14B), "u" _
        & ChrW(&H259) & ChrW(&H14B), ChrW(&H259) & "u", "i" & ChrW(&H259) & "u", "i" & ChrW(&H26A) & "u", "i" & _
        ChrW(&H259) & "m", ChrW(&H252) & "m", ChrW(&H251) & "m", "i" & ChrW(&H25B) & "m", "em", ChrW(&H25B) & "m", "am", _
        "i" & ChrW(&H250) & "m", "iu" & ChrW(&H250) & "m")
    ������ = Array("uq" & ChrW(&H14B), "u" & ChrW(&H302) & "q" & ChrW(&H14B), "uoq" & ChrW(&H14B), "u" & ChrW(&H302) & "oq" & _
        ChrW(&H14B), "oq" & ChrW(&H14B), "ieq", "u" & ChrW(&H302) & "eq", "iq", "uiq", "ie" & ChrW(&H302) & "q", "ie" & ChrW(&H302) _
        & "qi", "u" & ChrW(&H302) & "e" & ChrW(&H302) & "qi", "ioq", "u" & ChrW(&H302) & "oq", "uoq", "eqi", "ueqi", "ieqi", "u" & _
        ChrW(&H302) & "eqi", "iaqi", "u" & ChrW(&H302) & "aqi", "eq", "ueq", "eqi", "ueqi", "aqi", "uaqi", "uoqi", "oqi", "iqn", "u" & ChrW(&H302) & "qn", _
        "u" & ChrW(&H302) & "eqn", "ieqn", "u" & ChrW(&H302) & "e" & ChrW(&H302) & "qn", "ie" & ChrW(&H302) & "qn", "iaqn", _
        "u" & ChrW(&H302) & "aqn", "ue" & ChrW(&H302) & "qn", "e" & ChrW(&H302) & "qn", "aqn", "uaqn", "aqn", "uaqn", "eqn", _
        "ueqn", "aqi", "uaqi", "eqn", "ueqn", "ieqn", "u" & ChrW(&H302) & "eqn", "equ", "iequ", "aqu", "aqu", "aq", "uaq", "iaq", "u" _
        & ChrW(&H302) & "aq", "aq", "uaq", "iaq", "iaq" & ChrW(&H14B), "iuaq" & ChrW(&H14B), "aq" & ChrW(&H14B), "uaq" & _
        ChrW(&H14B), "aeq" & ChrW(&H14B), "oeq" & ChrW(&H14B), "ieq" & ChrW(&H14B), "u" & ChrW(&H302) & "eq" & ChrW(&H14B), _
        "eq" & ChrW(&H14B), "ueq" & ChrW(&H14B), "ieq" & ChrW(&H14B), "u" & ChrW(&H302) & "eq" & ChrW(&H14B), "eq" & _
        ChrW(&H14B), "ueq" & ChrW(&H14B), "ie" & ChrW(&H302) & "q" & ChrW(&H14B), "u" & ChrW(&H302) & "e" & ChrW(&H302) & "q" & _
        ChrW(&H14B), "e" & ChrW(&H302) & "q" & ChrW(&H14B), "ue" & ChrW(&H302) & "q" & ChrW(&H14B), "e" & ChrW(&H302) & "qu", "ie" & _
        ChrW(&H302) & "qu", "iqu", "ie" & ChrW(&H302) & "qm", "oqm", "aqm", "ieqm", "eqm", "eqm", "aqm", "iaqm", "u" & ChrW(&H302) _
        & "aqm")
    ��ĸ�� = Array("��", "��", "�K", "��", "��", "͸", "��", "��", "֪", "��", "��", "��", "��", "��", "��", "��", "а", _
        "�f", "��", "��", "��", "ٹ", "��", "��", "��", "��", "��", "Ҋ", "Ϫ", "Ⱥ", "��", "��", "ϻ", "Ӱ", "��", "��", _
        "��", "��")
    ������ = Array("p", "p" & ChrW(&H2B0), "b", "m", "t", "t" & ChrW(&H2B0), "d", "n", ChrW(&H288), ChrW(&H288) & _
        ChrW(&H2B0), ChrW(&H256), ChrW(&H273), "ts", "ts" & ChrW(&H2B0), "dz", "s", "z", "t" & ChrW(&H282), "t" & _
        ChrW(&H282) & ChrW(&H2B0), "d" & ChrW(&H290), ChrW(&H282), ChrW(&H290), "t" & ChrW(&H255), "t" & ChrW(&H255) & _
        ChrW(&H2B0), "d" & ChrW(&H291), ChrW(&H255), ChrW(&H291), "k", "k" & ChrW(&H2B0), "g", ChrW(&H14B), "x", _
        ChrW(&H263), ChrW(&H294), "", "j", "l", ChrW(&H235) & ChrW(&H291))
    ������ = Array("b", "p", "^b", "m", "d", "t", "^d", "n", "d" & ChrW(&H302), "t" & ChrW(&H302), "^d" & ChrW(&H302), _
        "n" & ChrW(&H302), ChrW(&H111), ChrW(&H167), "^" & ChrW(&H111), "s", "z", ChrW(&H111) & ChrW(&H302), ChrW(&H167) & _
        ChrW(&H302), "^" & ChrW(&H111) & ChrW(&H302), "s" & ChrW(&H302), "z" & ChrW(&H302), ChrW(&H111) & ChrW(&H303), _
        ChrW(&H167) & ChrW(&H303), "^" & ChrW(&H111) & ChrW(&H303), "s" & ChrW(&H303), "z" & ChrW(&H303), "g", "k", "^g", _
        ChrW(&H14B), "h", "x", "", "^", "j", "l", "^n" & ChrW(&H303))
End Sub

Function ��������ת����(��ĸ As String, ��ĸ As String, ���� As String) As String
    Dim S As String
    Dim i As Long
    Dim f As Boolean
    
    If ��ĸ = "��A" Then
        S = "���_"
    Else
        S = Replace(Replace(��ĸ, "A", ""), "B", "")
    End If
    f = False
    For i = 0 To 91
        If ��ĸ��(i) = S Then
            ��������ת���� = �����(i)
            f = True
            Exit For
        End If
    Next i
    If f = False Then
        MsgBox ("δ�ҵ���ĸ��" & ��ĸ)
        ��������ת���� = ""
        Exit Function
    End If
    f = False
    For i = 0 To 37
        If ��ĸ��(i) = ��ĸ Then
            ��������ת���� = ������(i) ' + ��������ת����
            f = True
            Exit For
        End If
    Next i
    If f = False Then
        MsgBox ("δ�ҵ���ĸ��" & ��ĸ)
        ��������ת���� = ""
        Exit Function
    End If
    
    If ���� = "��" Then
        Select Case Right(��������ת����, 1)
        Case "m"
            S = "p"
        Case "n"
            S = "t"
        Case ChrW(&H14B)
            S = "k"
        End Select
        ��������ת���� = Left(��������ת����, Len(��������ת����) - 1) & S
    End If
End Function

Function �ڸ���ת����(��ĸ As String, ��ĸ As String, ���� As String) As String
    Dim S As String
    Dim i As Long
    
    If ��ĸ = "��A" Then
        S = "���_"
    Else
        S = Replace(Replace(��ĸ, "A", ""), "B", "")
    End If
    For i = 0 To 91
        If ��ĸ��(i) = S Then
            �ڸ���ת���� = ������(i)
            Exit For
        End If
    Next i
    For i = 0 To 37
        If ��ĸ��(i) = ��ĸ Then
            �ڸ���ת���� = ������(i) + �ڸ���ת����
            Exit For
        End If
    Next i
    
    Select Case ����
    Case "ƽ"
        If Left(�ڸ���ת����, 1) = "^" Then
            S = ChrW(&H308)
        Else
            S = ChrW(&H304)
        End If
    Case "��"
        If Left(�ڸ���ת����, 1) = "^" Then
            S = ChrW(&H312)
        Else
            S = ChrW(&H301)
        End If
    Case "ȥ"
        If Left(�ڸ���ת����, 1) = "^" Then
            S = ChrW(&H30A)
        Else
            S = ChrW(&H300)
        End If
    Case "��"
        �ڸ���ת���� = Replace(�ڸ���ת����, "q", "")
        Select Case Right(�ڸ���ת����, 1)
        Case "m"
            S = "pb"
        Case "n"
            S = "td"
        Case ChrW(&H14B)
            S = "kg"
        End Select
        If Left(�ڸ���ת����, 1) = "^" Then
            �ڸ���ת���� = Left(�ڸ���ת����, Len(�ڸ���ת����) - 1) & Left(S, 1)
        Else
            �ڸ���ת���� = Left(�ڸ���ת����, Len(�ڸ���ת����) - 1) & Right(S, 1)
        End If
        S = ""
    End Select
    �ڸ���ת���� = Replace(Replace(�ڸ���ת����, "^", ""), "q", S)
    
    If ��ĸ <> "֬�_" Then
        If InStr("ҊϪȺ��ϻ", ��ĸ) Then
            If ��ĸ <> "֬�_" Then
                If Mid(�ڸ���ת����, 2, 2) = "u" & ChrW(&H302) Or ��ĸ = "��" Or Left(��ĸ, 1) = "��" Then
                    �ڸ���ת���� = Left(�ڸ���ת����, 1) & ChrW(&H303) & Mid(�ڸ���ת����, 2)
                ElseIf Mid(�ڸ���ת����, 2, 1) = "i" Then
                    �ڸ���ת���� = Left(�ڸ���ת����, 1) & ChrW(&H303) & Mid(�ڸ���ת����, 3)
                End If
            End If
        ElseIf InStr("�²�������", ��ĸ) Then
            If ��ĸ <> "֬�_" And ��ĸ <> "��" And Left(��ĸ, 1) <> "��" Then _
                �ڸ���ת���� = Replace(�ڸ���ת����, ChrW(&H303) & "i", ChrW(&H303))
        ElseIf ��ĸ = "��" Then
            If ��ĸ <> "֬�_" Then
                If Mid(�ڸ���ת����, 2, 2) = "u" & ChrW(&H302) Or ��ĸ = "��" Or Left(��ĸ, 1) = "��" Then
                    �ڸ���ת���� = "n" & ChrW(&H303) & Mid(�ڸ���ת����, 2)
                ElseIf Mid(�ڸ���ת����, 2, 1) = "i" Then
                    �ڸ���ת���� = "n" & ChrW(&H303) & Mid(�ڸ���ת����, 3)
                End If
            End If
        ElseIf ��ĸ = "��" Then
            If Left(�ڸ���ת����, 1) = "i" Then
                If ��ĸ = "֬�_" Or ��ĸ = "��" Or Left(��ĸ, 1) = "��" Then
                    �ڸ���ת���� = "j" & �ڸ���ת����
                Else
                    �ڸ���ת���� = "j" & Mid(�ڸ���ת����, 2)
                End If
            ElseIf Left(�ڸ���ת����, 1) = "u" Then
                �ڸ���ת���� = "w" & Mid(�ڸ���ת����, 2)
            ElseIf Left(�ڸ���ת����, 2) = "u" & ChrW(&H302) Then
                �ڸ���ת���� = "j" & �ڸ���ת����
            End If
        ElseIf ��ĸ = "��" Then
            If ��ĸ <> "֬�_" And ��ĸ <> "��" And Left(��ĸ, 1) <> "��" Then _
                �ڸ���ת���� = Replace(�ڸ���ת����, "ji", "j")
        End If
    End If
End Function

Function ����ת����(Col As Long) As Boolean
    ����ת���� = True
    Dim ��ĸ As String
    Dim ��ĸ As String
    Dim ���� As String
    Dim �������� As String
    Dim �ڸ��� As String
    
    ��ĸ = Selection.Tables(1).Cell(4, Col).Range.Text
    If Len(��ĸ) = 2 Then Exit Function
    ��ĸ = Left(��ĸ, Len(��ĸ) - 2)
    ��ĸ = Mid(��ĸ, 2, Len(��ĸ) - 2)
    ���� = Right(��ĸ, 1)
    ��ĸ = Left(��ĸ, 1)
    
    ��ĸ = ��ĸ��(��ĸ)
    �������� = ��������ת����(��ĸ, ��ĸ, ����)
    If �������� = "" Then
        ����ת���� = False
        Exit Function
    End If
    �ڸ��� = �ڸ���ת����(��ĸ, ��ĸ, ����)
    
    'MsgBox (��ĸ & vbCrLf & ��ĸ & vbCrLf & ����)
    'Selection.Tables(1).Cell(1, Col).Select
    'Selection.TypeText Text:=�ڸ���
    Selection.Tables(1).Cell(3, Col).Select
    Selection.TypeText Text:=��������
    Selection.Tables(1).Cell(4, Col).Select
    Selection.TypeText Text:=��ĸ & Left(��ĸ, 1)
    Selection.Font.Superscript = True
    Selection.TypeText Text:=Mid(��ĸ, 2)
    Selection.Font.Superscript = False
    Selection.Tables(1).Cell(5, Col).Select
    Selection.Font.Superscript = True
    Selection.TypeText Text:=����
    If Col < Selection.Tables(1).Columns.Count Then
        Selection.Tables(1).Cell(4, Col + 1).Select
    Else
        ����ת���� = False
    End If
End Function

Sub ����ת��()
    If Selection.Information(wdWithInTable) = False Then Exit Sub
    �������
    ����ת���� (Selection.Cells(1).ColumnIndex)
End Sub

Sub ����ת������()
    Dim i As Long
    If Selection.Information(wdWithInTable) = False Then Exit Sub
    �������
    Selection.Tables(1).Cell(1, 1).Select
    While ����ת����(Selection.Cells(1).ColumnIndex)
    Wend
End Sub

Sub ���и�����(Col As Long)
    Dim ��ĸ As String
    Dim ��ĸ As String
    Dim ���� As String
    Dim �������� As String
    Dim �ڸ��� As String
    
    ��ĸ = Selection.Tables(1).Cell(4, Col).Range.Text
    If Len(��ĸ) = 2 Then Exit Sub
    ��ĸ = Left(��ĸ, Len(��ĸ) - 2)
    ��ĸ = Mid(��ĸ, 2)
    ��ĸ = Left(��ĸ, 1)
    ���� = Selection.Tables(1).Cell(5, Col).Range.Text
    ���� = Left(����, Len(����) - 2)
    
    �������� = ��������ת����(��ĸ, ��ĸ, ����)
    If �������� = "" Then Exit Sub
    �ڸ��� = �ڸ���ת����(��ĸ, ��ĸ, ����)
    
    Selection.Tables(1).Cell(1, Col).Select
    Selection.TypeText Text:=�ڸ���
    Selection.Tables(1).Cell(3, Col).Select
    Selection.TypeText Text:=��������
    If Col < Selection.Tables(1).Columns.Count Then Selection.Tables(1).Cell(4, Col + 1).Select
End Sub

Sub ���и���()
    If Selection.Information(wdWithInTable) = False Then Exit Sub
    �������
    ���и����� (Selection.Cells(1).ColumnIndex)
End Sub
