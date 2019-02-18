Attribute VB_Name = "KanbunH"
Function Kaeriten(c As String)
'
' ��������ķ���
'
    Application.ScreenUpdating = False
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="EQ"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="\s\do1(" & c & ")"
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Size = Selection.Font.Size / 2
    Selection.Fields.ToggleShowCodes
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True
End Function
Function KaeritenOv(c As String)
'
' �����ص��ķ���
'
    Dim S As String
    S = Selection.Text
    Application.ScreenUpdating = False
    Selection.TypeBackspace
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeBackspace
    Selection.TypeText Text:="EQ"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="\o\al(" & S & ", \s\do1(" & c & "))"
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Size = Selection.Font.Size / 2
    Selection.Fields.ToggleShowCodes
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True
End Function
Sub Kanbun_Line()
'
' Kanbun_Line ��
' �����ּ�����
'
    Selection.TypeText Text:="�D"
End Sub
Sub Kanbun_RE()
'
' Kanbun_RE ��
' ������
'
    If Selection.Type <> wdNoSelection And Selection.Type <> wdSelectionIP Then _
        Selection.TypeBackspace
    Kaeriten c:="��"
End Sub
Sub Kanbun_1()
'
' Kanbun_1 ��
' ����һ��
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="һ"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="һ"
    End If
End Sub
Sub Kanbun_2()
'
' Kanbun_2 ��
' �������
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="��"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="��"
    End If
End Sub
Sub Kanbun_3()
'
' Kanbun_3 ��
' ��������
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="��"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="��"
    End If
End Sub
Sub Kanbun_Up()
'
' Kanbun_Up ��
' �����ϵ�
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="��"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="��"
    End If
End Sub
Sub Kanbun_Mid()
'
' Kanbun_Mid ��
' �����е�
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="��"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="��"
    End If
End Sub
Sub Kanbun_Down()
'
' Kanbun_Down ��
' �����µ�
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="��"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="��"
    End If
End Sub
Sub Kanbun_A()
'
' Kanbun_A ��
' ����׵�
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="��"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="��"
    End If
End Sub
Sub Kanbun_B()
'
' Kanbun_B ��
' �����ҵ�
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="��"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="��"
    End If
End Sub
Sub Kanbun_C()
'
' Kanbun_C ��
' �������
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="��"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="��"
    End If
End Sub
Sub Kanbun_D()
'
' Kanbun_D ��
' ���붡��
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="��"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="��"
    End If
End Sub
Sub Kanbun_Sky()
'
' Kanbun_Sky ��
' �������
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="��"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="��"
    End If
End Sub
Sub Kanbun_Earth()
'
' Kanbun_Earth ��
' ����ص�
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="��"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="��"
    End If
End Sub
Sub Kanbun_Man()
'
' Kanbun_Man ��
' �����˵�
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="��"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="��"
    End If
End Sub
Sub Kanbun_1_RE()
'
' Kanbun_1_RE ��
' ����һ���
'
    Application.ScreenUpdating = False
    If Selection.Type <> wdNoSelection And Selection.Type <> wdSelectionIP Then _
        Selection.TypeBackspace
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeBackspace
    Selection.TypeText Text:="EQ"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="\o\al(\s\up1(һ),\s\do1(��)) "
    Selection.MoveLeft Unit:=wdCharacter, Count:=3
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Size = Selection.Font.Size / 2
    Selection.MoveLeft Unit:=wdCharacter, Count:=10
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Size = Selection.Font.Size / 2
    Selection.Fields.ToggleShowCodes
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True
End Sub
Sub Kanbun_2_Line()
'
' Kanbun_2_Line ��
' ���������ּ�����
'
    Application.ScreenUpdating = False
    If Selection.Type <> wdNoSelection And Selection.Type <> wdSelectionIP Then _
        Selection.TypeBackspace
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeBackspace
    Selection.TypeText Text:="EQ"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="\o\al(\s\do1(��),�D)"
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Size = Selection.Font.Size / 2
    Selection.Fields.ToggleShowCodes
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True
End Sub
Sub Kanbun_Up_RE()
'
' Kanbun_Up_RE ��
' �����ϥ��
'
    Application.ScreenUpdating = False
    If Selection.Type <> wdNoSelection And Selection.Type <> wdSelectionIP Then _
        Selection.TypeBackspace
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeBackspace
    Selection.TypeText Text:="EQ"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="\o\al(\s\up3(��),\s\do1(��)) "
    Selection.MoveLeft Unit:=wdCharacter, Count:=3
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Size = Selection.Font.Size / 2
    Selection.MoveLeft Unit:=wdCharacter, Count:=10
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Size = Selection.Font.Size / 2
    Selection.Fields.ToggleShowCodes
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True
End Sub
