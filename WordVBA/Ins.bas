Attribute VB_Name = "Ins"
Public Function InsertS(n As Integer)
    'Selection.InsertSSymbol CharacterNumber:=n, Font:="+西文正文", Unicode:=True
    Selection.TypeText Text:=ChrW(n) ' E12.6 修改
End Function
Public Function InsertST(n As Integer)
    InsertS (n)
    'Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    '？
    'Selection.MoveRight Unit:=wdCharacter, Count:=1
End Function
Sub Ed()
    InsertS (&H2202)
End Sub
Sub delta()
    InsertS (&H2206)
End Sub
Sub Sigma()
    InsertS (&H2211)
End Sub
Sub Minus()
    InsertS (&H2212)
End Sub
Sub Root()
    InsertS (&H221A)
End Sub
Sub Infinity()
    InsertS (&H221E)
End Sub
Sub S()
    InsertS (&H222B)
End Sub
Sub Chara1()
    InsertS (12272)
End Sub
Sub Chara2()
    InsertS (12273)
End Sub
Sub Chara3()
    InsertS (12274)
End Sub
Sub Chara4()
    InsertS (12275)
End Sub
Sub Chara5()
    InsertS (12276)
End Sub
Sub Chara6()
    InsertS (12277)
End Sub
Sub Chara7()
    InsertS (12278)
End Sub
Sub Chara8()
    InsertS (12279)
End Sub
Sub Chara9()
    InsertS (12280)
End Sub
Sub Chara10()
    InsertS (12281)
End Sub
Sub Chara11()
    InsertS (12282)
End Sub
Sub Chara12()
    InsertS (12283)
End Sub
Sub Unicode切换()
    Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend
    Selection.ToggleCharacterCode
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
End Sub
Sub Unicode切换10次()
    For i = 0 To 10
        Unicode切换
    Next i
End Sub
Public Function InsertField(Content As String) ' I6.11
    Application.ScreenUpdating = False
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False
    Selection.TypeBackspace
    Selection.Delete
    Selection.TypeText Text:=Content
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes
    ActiveWindow.View.ShowFieldCodes = Not ActiveWindow.View.ShowFieldCodes
    Application.ScreenUpdating = True
End Function
