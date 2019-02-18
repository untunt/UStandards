Attribute VB_Name = "广韵"
Option Explicit

Dim 韵母表 As Variant
Dim 音标表 As Variant
Dim 音读表 As Variant
Dim 声母表 As Variant
Dim 音标声 As Variant
Dim 音读声 As Variant

Function 韵母简化(韵母 As String) As String
    韵母简化 = Left(韵母, 1)
    If InStr("冬鍾江之魚虞模灰咍諄臻文欣魂痕寒桓蕭宵肴豪歌蒸侯尤幽侵覃談鹽添咸銜嚴凡", 韵母简化) Then
        ' 无等无呼
        If InStr("宵侵鹽", 韵母简化) Then 韵母简化 = Left(韵母, 2)
        If 韵母 = "蒸三合" Then 韵母简化 = "職合"
    ElseIf InStr("支脂微齊祭廢佳皆夬真元刪山泰先仙陽唐清青登", 韵母简化) Then
        ' 无等有呼
        If InStr("支脂祭仙真", 韵母简化) Then 韵母简化 = Left(韵母, 2)
        韵母简化 = 韵母简化 + Right(韵母, 1)
        If 韵母 = "真A三開" Then 韵母简化 = "真A"
    ElseIf InStr("東", 韵母简化) Then
        ' 有等无呼
        韵母简化 = Left(韵母, 2)
    Else '有等有呼（戈麻庚耕）
        韵母简化 = 韵母
    End If
End Function

Function 韵母转换(韵母 As String) As String
End Function

Sub 添加数组()
    韵母表 = Array("東一", "東三", "冬", "鍾", "江", "支開", "支合", "脂開", "脂合", "之", "微開", "微合", "魚", _
        "虞", "模", "齊開", "齊合", "祭開", "祭合", "廢開", "廢合", "佳開", "佳合", "皆開", "皆合", "夬開", "夬合", _
        "灰", "咍", "真開", "真合", "諄", "臻", "文", "欣", "元開", "元合", "魂", "痕", "寒", "桓", "刪開", "刪合", _
        "山開", "山合", "泰開", "泰合", "先開", "先合", "仙開", "仙合", "蕭", "宵", "肴", "豪", "歌", "戈一合", _
        "戈三開", "戈三合", "麻二開", "麻二合", "麻三開", "陽合", "陽開", "唐開", "唐合", "庚二開", "庚二合", "庚三開", _
        "庚三合", "耕二開", "耕二合", "清開", "清合", "青開", "青合", "蒸", "職合", "登開", "登合", "侯", "尤", "幽", _
        "侵", "覃", "談", "鹽", "添", "咸", "銜", "嚴", "凡")
    音标表 = Array("u" & ChrW(&H14B), "iu" & ChrW(&H14B), "o" & ChrW(&H14B), "io" & ChrW(&H14B), ChrW(&H254) & _
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
    音读表 = Array("uq" & ChrW(&H14B), "u" & ChrW(&H302) & "q" & ChrW(&H14B), "uoq" & ChrW(&H14B), "u" & ChrW(&H302) & "oq" & _
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
    声母表 = Array("幫", "滂", "並", "明", "端", "透", "定", "泥", "知", "徹", "澄", "娘", "精", "清", "從", "心", "邪", _
        "莊", "初", "崇", "生", "俟", "章", "昌", "常", "書", "船", "見", "溪", "群", "疑", "曉", "匣", "影", "云", "以", _
        "來", "日")
    音标声 = Array("p", "p" & ChrW(&H2B0), "b", "m", "t", "t" & ChrW(&H2B0), "d", "n", ChrW(&H288), ChrW(&H288) & _
        ChrW(&H2B0), ChrW(&H256), ChrW(&H273), "ts", "ts" & ChrW(&H2B0), "dz", "s", "z", "t" & ChrW(&H282), "t" & _
        ChrW(&H282) & ChrW(&H2B0), "d" & ChrW(&H290), ChrW(&H282), ChrW(&H290), "t" & ChrW(&H255), "t" & ChrW(&H255) & _
        ChrW(&H2B0), "d" & ChrW(&H291), ChrW(&H255), ChrW(&H291), "k", "k" & ChrW(&H2B0), "g", ChrW(&H14B), "x", _
        ChrW(&H263), ChrW(&H294), "", "j", "l", ChrW(&H235) & ChrW(&H291))
    音读声 = Array("b", "p", "^b", "m", "d", "t", "^d", "n", "d" & ChrW(&H302), "t" & ChrW(&H302), "^d" & ChrW(&H302), _
        "n" & ChrW(&H302), ChrW(&H111), ChrW(&H167), "^" & ChrW(&H111), "s", "z", ChrW(&H111) & ChrW(&H302), ChrW(&H167) & _
        ChrW(&H302), "^" & ChrW(&H111) & ChrW(&H302), "s" & ChrW(&H302), "z" & ChrW(&H302), ChrW(&H111) & ChrW(&H303), _
        ChrW(&H167) & ChrW(&H303), "^" & ChrW(&H111) & ChrW(&H303), "s" & ChrW(&H303), "z" & ChrW(&H303), "g", "k", "^g", _
        ChrW(&H14B), "h", "x", "", "^", "j", "l", "^n" & ChrW(&H303))
End Sub

Function 国际音标转换子(声母 As String, 韵母 As String, 声调 As String) As String
    Dim S As String
    Dim i As Long
    Dim f As Boolean
    
    If 韵母 = "真A" Then
        S = "真開"
    Else
        S = Replace(Replace(韵母, "A", ""), "B", "")
    End If
    f = False
    For i = 0 To 91
        If 韵母表(i) = S Then
            国际音标转换子 = 音标表(i)
            f = True
            Exit For
        End If
    Next i
    If f = False Then
        MsgBox ("未找到韵母：" & 韵母)
        国际音标转换子 = ""
        Exit Function
    End If
    f = False
    For i = 0 To 37
        If 声母表(i) = 声母 Then
            国际音标转换子 = 音标声(i) ' + 国际音标转换子
            f = True
            Exit For
        End If
    Next i
    If f = False Then
        MsgBox ("未找到声母：" & 声母)
        国际音标转换子 = ""
        Exit Function
    End If
    
    If 声调 = "入" Then
        Select Case Right(国际音标转换子, 1)
        Case "m"
            S = "p"
        Case "n"
            S = "t"
        Case ChrW(&H14B)
            S = "k"
        End Select
        国际音标转换子 = Left(国际音标转换子, Len(国际音标转换子) - 1) & S
    End If
End Function

Function 于干语转换子(声母 As String, 韵母 As String, 声调 As String) As String
    Dim S As String
    Dim i As Long
    
    If 韵母 = "真A" Then
        S = "真開"
    Else
        S = Replace(Replace(韵母, "A", ""), "B", "")
    End If
    For i = 0 To 91
        If 韵母表(i) = S Then
            于干语转换子 = 音读表(i)
            Exit For
        End If
    Next i
    For i = 0 To 37
        If 声母表(i) = 声母 Then
            于干语转换子 = 音读声(i) + 于干语转换子
            Exit For
        End If
    Next i
    
    Select Case 声调
    Case "平"
        If Left(于干语转换子, 1) = "^" Then
            S = ChrW(&H308)
        Else
            S = ChrW(&H304)
        End If
    Case "上"
        If Left(于干语转换子, 1) = "^" Then
            S = ChrW(&H312)
        Else
            S = ChrW(&H301)
        End If
    Case "去"
        If Left(于干语转换子, 1) = "^" Then
            S = ChrW(&H30A)
        Else
            S = ChrW(&H300)
        End If
    Case "入"
        于干语转换子 = Replace(于干语转换子, "q", "")
        Select Case Right(于干语转换子, 1)
        Case "m"
            S = "pb"
        Case "n"
            S = "td"
        Case ChrW(&H14B)
            S = "kg"
        End Select
        If Left(于干语转换子, 1) = "^" Then
            于干语转换子 = Left(于干语转换子, Len(于干语转换子) - 1) & Left(S, 1)
        Else
            于干语转换子 = Left(于干语转换子, Len(于干语转换子) - 1) & Right(S, 1)
        End If
        S = ""
    End Select
    于干语转换子 = Replace(Replace(于干语转换子, "^", ""), "q", S)
    
    If 韵母 <> "脂開" Then
        If InStr("見溪群曉匣", 声母) Then
            If 韵母 <> "脂開" Then
                If Mid(于干语转换子, 2, 2) = "u" & ChrW(&H302) Or 韵母 = "幽" Or Left(韵母, 1) = "真" Then
                    于干语转换子 = Left(于干语转换子, 1) & ChrW(&H303) & Mid(于干语转换子, 2)
                ElseIf Mid(于干语转换子, 2, 1) = "i" Then
                    于干语转换子 = Left(于干语转换子, 1) & ChrW(&H303) & Mid(于干语转换子, 3)
                End If
            End If
        ElseIf InStr("章昌常書船", 声母) Then
            If 韵母 <> "脂開" And 韵母 <> "幽" And Left(韵母, 1) <> "真" Then _
                于干语转换子 = Replace(于干语转换子, ChrW(&H303) & "i", ChrW(&H303))
        ElseIf 声母 = "疑" Then
            If 韵母 <> "脂開" Then
                If Mid(于干语转换子, 2, 2) = "u" & ChrW(&H302) Or 韵母 = "幽" Or Left(韵母, 1) = "真" Then
                    于干语转换子 = "n" & ChrW(&H303) & Mid(于干语转换子, 2)
                ElseIf Mid(于干语转换子, 2, 1) = "i" Then
                    于干语转换子 = "n" & ChrW(&H303) & Mid(于干语转换子, 3)
                End If
            End If
        ElseIf 声母 = "云" Then
            If Left(于干语转换子, 1) = "i" Then
                If 韵母 = "脂開" Or 韵母 = "幽" Or Left(韵母, 1) = "真" Then
                    于干语转换子 = "j" & 于干语转换子
                Else
                    于干语转换子 = "j" & Mid(于干语转换子, 2)
                End If
            ElseIf Left(于干语转换子, 1) = "u" Then
                于干语转换子 = "w" & Mid(于干语转换子, 2)
            ElseIf Left(于干语转换子, 2) = "u" & ChrW(&H302) Then
                于干语转换子 = "j" & 于干语转换子
            End If
        ElseIf 声母 = "以" Then
            If 韵母 <> "脂開" And 韵母 <> "幽" And Left(韵母, 1) <> "真" Then _
                于干语转换子 = Replace(于干语转换子, "ji", "j")
        End If
    End If
End Function

Function 反切转换子(Col As Long) As Boolean
    反切转换子 = True
    Dim 声母 As String
    Dim 韵母 As String
    Dim 声调 As String
    Dim 国际音标 As String
    Dim 于干语 As String
    
    声母 = Selection.Tables(1).Cell(4, Col).Range.Text
    If Len(声母) = 2 Then Exit Function
    声母 = Left(声母, Len(声母) - 2)
    韵母 = Mid(声母, 2, Len(声母) - 2)
    声调 = Right(声母, 1)
    声母 = Left(声母, 1)
    
    韵母 = 韵母简化(韵母)
    国际音标 = 国际音标转换子(声母, 韵母, 声调)
    If 国际音标 = "" Then
        反切转换子 = False
        Exit Function
    End If
    于干语 = 于干语转换子(声母, 韵母, 声调)
    
    'MsgBox (声母 & vbCrLf & 韵母 & vbCrLf & 声调)
    'Selection.Tables(1).Cell(1, Col).Select
    'Selection.TypeText Text:=于干语
    Selection.Tables(1).Cell(3, Col).Select
    Selection.TypeText Text:=国际音标
    Selection.Tables(1).Cell(4, Col).Select
    Selection.TypeText Text:=声母 & Left(韵母, 1)
    Selection.Font.Superscript = True
    Selection.TypeText Text:=Mid(韵母, 2)
    Selection.Font.Superscript = False
    Selection.Tables(1).Cell(5, Col).Select
    Selection.Font.Superscript = True
    Selection.TypeText Text:=声调
    If Col < Selection.Tables(1).Columns.Count Then
        Selection.Tables(1).Cell(4, Col + 1).Select
    Else
        反切转换子 = False
    End If
End Function

Sub 反切转换()
    If Selection.Information(wdWithInTable) = False Then Exit Sub
    添加数组
    反切转换子 (Selection.Cells(1).ColumnIndex)
End Sub

Sub 反切转换整行()
    Dim i As Long
    If Selection.Information(wdWithInTable) = False Then Exit Sub
    添加数组
    Selection.Tables(1).Cell(1, 1).Select
    While 反切转换子(Selection.Cells(1).ColumnIndex)
    Wend
End Sub

Sub 反切更新子(Col As Long)
    Dim 声母 As String
    Dim 韵母 As String
    Dim 声调 As String
    Dim 国际音标 As String
    Dim 于干语 As String
    
    声母 = Selection.Tables(1).Cell(4, Col).Range.Text
    If Len(声母) = 2 Then Exit Sub
    声母 = Left(声母, Len(声母) - 2)
    韵母 = Mid(声母, 2)
    声母 = Left(声母, 1)
    声调 = Selection.Tables(1).Cell(5, Col).Range.Text
    声调 = Left(声调, Len(声调) - 2)
    
    国际音标 = 国际音标转换子(声母, 韵母, 声调)
    If 国际音标 = "" Then Exit Sub
    于干语 = 于干语转换子(声母, 韵母, 声调)
    
    Selection.Tables(1).Cell(1, Col).Select
    Selection.TypeText Text:=于干语
    Selection.Tables(1).Cell(3, Col).Select
    Selection.TypeText Text:=国际音标
    If Col < Selection.Tables(1).Columns.Count Then Selection.Tables(1).Cell(4, Col + 1).Select
End Sub

Sub 反切更新()
    If Selection.Information(wdWithInTable) = False Then Exit Sub
    添加数组
    反切更新子 (Selection.Cells(1).ColumnIndex)
End Sub
