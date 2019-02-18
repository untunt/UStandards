Attribute VB_Name = "KanbunH"
Function Kaeriten(c As String)
'
' 插入独立的返点
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
' 插入重叠的返点
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
' Kanbun_Line 宏
' 插入字间连线
'
    Selection.TypeText Text:="―"
End Sub
Sub Kanbun_RE()
'
' Kanbun_RE 宏
' 插入レ点
'
    If Selection.Type <> wdNoSelection And Selection.Type <> wdSelectionIP Then _
        Selection.TypeBackspace
    Kaeriten c:="レ"
End Sub
Sub Kanbun_1()
'
' Kanbun_1 宏
' 插入一点
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="一"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="一"
    End If
End Sub
Sub Kanbun_2()
'
' Kanbun_2 宏
' 插入二点
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="二"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="二"
    End If
End Sub
Sub Kanbun_3()
'
' Kanbun_3 宏
' 插入三点
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="三"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="三"
    End If
End Sub
Sub Kanbun_Up()
'
' Kanbun_Up 宏
' 插入上点
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="上"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="上"
    End If
End Sub
Sub Kanbun_Mid()
'
' Kanbun_Mid 宏
' 插入中点
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="中"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="中"
    End If
End Sub
Sub Kanbun_Down()
'
' Kanbun_Down 宏
' 插入下点
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="下"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="下"
    End If
End Sub
Sub Kanbun_A()
'
' Kanbun_A 宏
' 插入甲点
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="甲"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="甲"
    End If
End Sub
Sub Kanbun_B()
'
' Kanbun_B 宏
' 插入乙点
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="乙"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="乙"
    End If
End Sub
Sub Kanbun_C()
'
' Kanbun_C 宏
' 插入丙点
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="丙"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="丙"
    End If
End Sub
Sub Kanbun_D()
'
' Kanbun_D 宏
' 插入丁点
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="丁"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="丁"
    End If
End Sub
Sub Kanbun_Sky()
'
' Kanbun_Sky 宏
' 插入天点
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="天"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="天"
    End If
End Sub
Sub Kanbun_Earth()
'
' Kanbun_Earth 宏
' 插入地点
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="地"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="地"
    End If
End Sub
Sub Kanbun_Man()
'
' Kanbun_Man 宏
' 插入人点
'
    If Selection.Type = wdNoSelection Or Selection.Type = wdSelectionIP Then
        Kaeriten c:="人"
    ElseIf Len(Selection.Text) = 1 Then
        KaeritenOv c:="人"
    End If
End Sub
Sub Kanbun_1_RE()
'
' Kanbun_1_RE 宏
' 插入一レ点
'
    Application.ScreenUpdating = False
    If Selection.Type <> wdNoSelection And Selection.Type <> wdSelectionIP Then _
        Selection.TypeBackspace
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeBackspace
    Selection.TypeText Text:="EQ"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="\o\al(\s\up1(一),\s\do1(レ)) "
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
' Kanbun_2_Line 宏
' 插入二点和字间连线
'
    Application.ScreenUpdating = False
    If Selection.Type <> wdNoSelection And Selection.Type <> wdSelectionIP Then _
        Selection.TypeBackspace
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeBackspace
    Selection.TypeText Text:="EQ"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="\o\al(\s\do1(二),―)"
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Font.Size = Selection.Font.Size / 2
    Selection.Fields.ToggleShowCodes
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True
End Sub
Sub Kanbun_Up_RE()
'
' Kanbun_Up_RE 宏
' 插入上レ点
'
    Application.ScreenUpdating = False
    If Selection.Type <> wdNoSelection And Selection.Type <> wdSelectionIP Then _
        Selection.TypeBackspace
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeBackspace
    Selection.TypeText Text:="EQ"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="\o\al(\s\up3(上),\s\do1(レ)) "
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
