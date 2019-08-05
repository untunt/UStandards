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
    Selection.TypeText Text:="a"
    Selection.TypeBackspace
    Application.ScreenUpdating = True
End Function
' 以下为I8.4新增
Sub punct_prime()
    InsertS (&H2032)
End Sub
Sub punct_double_prime()
    InsertS (&H2033)
End Sub
Sub punct_liaison()
    InsertS (&H203F)
End Sub
Sub punct_dotted_cir()
    InsertS (&H25CC)
End Sub
Sub punct_double_hyphen()
    InsertS (&H2E40)
End Sub
Sub comb_ls_rnd_sup()
    InsertS (&H351)
End Sub
Sub comb_mr_rnd_sup()
    InsertS (&H357)
End Sub
Sub comb_sub_w()
    InsertS (&H32B)
End Sub
Sub comb_sub_m()
    InsertS (&H33C)
End Sub
Sub comb_sub_atr()
    InsertS (&H318)
End Sub
Sub comb_sub_rtr()
    InsertS (&H319)
End Sub
Sub comb_sub_mid_ctr()
    InsertS (&H353)
End Sub
Sub comb_sup_n()
    InsertS (&H346)
End Sub
Sub comb_sub_up()
    InsertS (&H34E)
End Sub
Sub comb_sub_v()
    InsertS (&H32C)
End Sub
Sub comb_sup_inv_breve()
    InsertS (&H311)
End Sub
Sub comb_sub_dot()
    InsertS (&H323)
End Sub
Sub modi_6()
    InsertS (&H2BB)
End Sub
Sub modi_9()
    InsertS (&H2BC)
End Sub
Sub modi_9_mirrored()
    InsertS (&H2BD)
End Sub
Sub modi_ls_rnd_sup()
    InsertS (&H2BF)
End Sub
Sub modi_mr_rnd_sup()
    InsertS (&H2BE)
End Sub
Sub modi_ls_rnd_sub()
    InsertS (&H2D3)
End Sub
Sub modi_mr_rnd_sub()
    InsertS (&H2D2)
End Sub
Sub modi_sup_ring()
    InsertS (&H2DA)
End Sub
Sub modi_sub_ring()
    InsertS (&H2F3)
End Sub
Sub modi_sup_up()
    InsertS (&HA71B)
End Sub
Sub modi_sup_down()
    InsertS (&HA71C)
End Sub
Sub modi_sup_equal()
    InsertS (&H2ED)
End Sub
Sub modi_sup_dot()
    InsertS (&H2D9)
End Sub
Sub modi_mid_dot()
    InsertS (&HA78F)
End Sub
Sub modi_sup_breve()
    InsertS (&H2D8)
End Sub
Sub modi_sup_x()
    InsertS (&H2DF)
End Sub
Sub modi_mid_grave()
    InsertS (&H2F4)
End Sub
Sub modi_sup_e()
    InsertS (&H1D49)
End Sub
Sub modi_sup_o()
    InsertS (&H1D52)
End Sub
