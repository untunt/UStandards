Attribute VB_Name = "专用排版"
Option Explicit
Sub 自定义字符缩放()
    Dim Temp As String
    Temp = InputBox("输入字符缩放比例（%）", "自定义字符缩放", Selection.Font.Scaling)
    If Temp <> "" Then Selection.Font.Scaling = Temp
End Sub
Sub 自定义字符位置()
    Dim Temp As String
    Temp = InputBox("输入字符位置（磅）", "自定义字符位置", Selection.Font.Position)
    If Temp <> "" Then Selection.Font.Position = Temp
End Sub

Sub 合一音标() ' completed F5.17
    Dim Sel As String
    Dim Pos As Long
    Dim 字 As String
    Dim 音 As String
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
    字 = Selection
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace
    Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend
    音 = Selection
    Selection.Delete
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.TypeText Text:="EQ \o(" + 字 + ",\s\up 6("
    Selection.Font.Superscript = True
    Selection.TypeText Text:=音
    Selection.Font.Superscript = False
    Selection.TypeText Text:="))"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    a = True
    Application.ScreenUpdating = True
End Sub
Sub TimesNewRoman化()
    Selection.LanguageID = wdEnglishUS
    Selection.NoProofing = False
    Application.CheckLanguage = False
    Selection.Font.Name = "Times New Roman"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
End Sub

Sub 着重用顿号() 'F7.19
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
    Selection.TypeText Text:=ch & "、"
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
Sub 下标字_连读()
    下标字 "︵"
End Sub
Sub 下标字_轻声()
    下标字 "·"
End Sub
Sub 下标字_显式儿化()
    下标字 "—"
End Sub
Function 下标字(Str As String)
    Application.ScreenUpdating = False
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="EQ \o("
    Selection.TypeText Text:=Str
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Superscript = wdToggle
    Selection.Font.Name = "宋体"
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
Sub 移除零宽间隔()
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
