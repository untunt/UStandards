Attribute VB_Name = "ר���Ű�"
Option Explicit
Sub �Զ����ַ�����()
    Dim Temp As String
    Temp = InputBox("�����ַ����ű�����%��", "�Զ����ַ�����", Selection.Font.Scaling)
    If Temp <> "" Then Selection.Font.Scaling = Temp
End Sub
Sub �Զ����ַ�λ��()
    Dim Temp As String
    Temp = InputBox("�����ַ�λ�ã�����", "�Զ����ַ�λ��", Selection.Font.Position)
    If Temp <> "" Then Selection.Font.Position = Temp
End Sub

Sub ��һ����() ' completed F5.17
    Dim Sel As String
    Dim Pos As Long
    Dim �� As String
    Dim �� As String
    Static a As Boolean
    Application.ScreenUpdating = False
    If a Then
        ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes
        ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes
        a = False
        Application.ScreenUpdating = True
        Exit Sub
    End If
    If Selection.Type = wdSelectionIP Or Selection.Type = wdNoSelection Then Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    �� = Selection
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace
    Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend
    �� = Selection
    Selection.Delete
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.TypeText Text:="EQ \o(" + �� + ",\s\up 6("
    Selection.Font.Superscript = True
    Selection.TypeText Text:=��
    Selection.Font.Superscript = False
    Selection.TypeText Text:="))"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    a = True
    Application.ScreenUpdating = True
End Sub
Sub TimesNewRoman��()
    Selection.LanguageID = wdEnglishUS
    Selection.NoProofing = False
    Application.CheckLanguage = False
    Selection.Font.Name = "Times New Roman"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
End Sub

Sub �����öٺ�() 'F7.19
    Dim ch As String
    Application.ScreenUpdating = False
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    ch = Selection
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False
    Selection.TypeBackspace
    Selection.TypeText Text:="EQ"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:=ch & "��"
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Spacing = -4
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Spacing = -1.3
    Selection.MoveRight Unit:=wdCharacter, Count:=3
    ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes
    ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes
    Application.ScreenUpdating = True
End Sub
Sub �±���_����()
    �±��� "��"
End Sub
Sub �±���_����()
    �±��� "��"
End Sub
Sub �±���_��ʽ����()
    �±��� "��"
End Sub
Function �±���(Str As String)
    Application.ScreenUpdating = False
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="EQ \o("
    Selection.TypeText Text:=Str
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Superscript = wdToggle
    Selection.Font.Name = "����"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.Font.Superscript = wdToggle
    Selection.Font.Name = "Times New Roman"
    Selection.TypeText Text:=","
    Selection.Delete
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Subscript = wdToggle
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.Font.Subscript = wdToggle
    Selection.TypeText Text:=")"
    Selection.Delete
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes
    ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes
    Application.ScreenUpdating = True
End Function
Sub �Ƴ������()
    With Selection.Find
        .Text = ChrW(&H200B)
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .MatchFuzzy = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
