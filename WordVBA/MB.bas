Attribute VB_Name = "MB"
Option Explicit


'================================
' ÑÕÉ«×ª»»Ëã·¨
' RGB2HSL
' HSL2RGB
' http://www.cnhup.com
'================================
Private Type HSL
    H As Double    ' 0-360
    S As Double    ' 0-1
    L As Double    ' 0-1
End Type

Private Type Colour
    R As Double    ' 0-1
    G As Double    ' 0-1
    B As Double    ' 0-1
End Type

' Calculate HSL from RGB
' Hue is in degrees
' Lightness is between 0 and 1
' Saturation is between 0 and 1
Private Function RGB2HSL(ByRef c1 As Colour) As HSL
    Dim themin As Double, themax As Double, delta As Double
    Dim c2 As HSL

    themin = MinD(c1.R, MinD(c1.G, c1.B))
    themax = MaxD(c1.R, MaxD(c1.G, c1.B))

    delta = themax - themin
    c2.L = (themin + themax) / 2
    c2.S = 0

    If ((c2.L > 0) And (c2.L < 1)) Then
        If (c2.L < 0.5) Then
            c2.S = delta / (2 * c2.L)
        Else
            c2.S = delta / (2 - 2 * c2.L)
        End If
    End If

    c2.H = 0

    If (delta > 0) Then
        If ((themax = c1.R) And (themax <> c1.G)) Then _
           c2.H = c2.H + (c1.G - c1.B) / delta
        If ((themax = c1.G) And (themax <> c1.B)) Then _
           c2.H = c2.H + (2 + (c1.B - c1.R) / delta)
        If ((themax = c1.B) And (themax <> c1.R)) Then _
           c2.H = c2.H + (4 + (c1.R - c1.G) / delta)

        c2.H = c2.H * 60
    End If

    RGB2HSL = c2
End Function

' Calculate RGB from HSL, reverse of RGB2HSL()
' Hue is in degrees
' Lightness is between 0 and 1
' Saturation is between 0 and 1
Private Function HSL2RGB(ByRef c1 As HSL) As Colour
    Dim c2 As Colour, sat As Colour, ctmp As Colour

    Do While (c1.H < 0)
        c1.H = c1.H + 360
    Loop

    Do While (c1.H > 360)
        c1.H = c1.H - 360
    Loop

    If (c1.H < 120) Then
        sat.R = (120 - c1.H) / 60
        sat.G = c1.H / 60
        sat.B = 0
    ElseIf (c1.H < 240) Then
        sat.R = 0
        sat.G = (240 - c1.H) / 60
        sat.B = (c1.H - 120) / 60
    Else
        sat.R = (c1.H - 240) / 60
        sat.G = 0
        sat.B = (360 - c1.H) / 60
    End If

    sat.R = MinD(sat.R, 1)
    sat.G = MinD(sat.G, 1)
    sat.B = MinD(sat.B, 1)

    ctmp.R = 2 * c1.S * sat.R + (1 - c1.S)
    ctmp.G = 2 * c1.S * sat.G + (1 - c1.S)
    ctmp.B = 2 * c1.S * sat.B + (1 - c1.S)

    If (c1.L < 0.5) Then
        c2.R = c1.L * ctmp.R
        c2.G = c1.L * ctmp.G
        c2.B = c1.L * ctmp.B
    Else
        c2.R = (1 - c1.L) * ctmp.R + 2 * c1.L - 1
        c2.G = (1 - c1.L) * ctmp.G + 2 * c1.L - 1
        c2.B = (1 - c1.L) * ctmp.B + 2 * c1.L - 1
    End If

    HSL2RGB = c2
End Function

Private Function MinD(ByVal inA As Double, ByVal inB As Double) As Double
    If (inA < inB) Then MinD = inA Else MinD = inB
End Function

Private Function MaxD(ByVal inA As Double, ByVal inB As Double) As Double
    If (inA > inB) Then MaxD = inA Else MaxD = inB
End Function

Function NewColour(aR As Double, aG As Double, aB As Double) As Colour
    With NewColour
        .R = aR
        .G = aG
        .B = aB
    End With
End Function

Sub Color1()
    Randomize
    Selection.Font.Color = RGB(90 * Rnd + 30, 90 * Rnd + 30, 90 * Rnd + 30)
End Sub

Sub Color20()
    Dim i As Integer
    Randomize
    For i = 1 To 20
        Coloring
    Next i
End Sub

Sub Coloring()
    Dim c As HSL
    Dim c2 As Colour
    c.H = 360 * Rnd
    c.L = 0.3 * Rnd + 0.15
    c.S = 0.3 * Rnd + 0.6
    c2 = HSL2RGB(c)
    Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend
    If Left(Selection, 1) = "," Or _
       Left(Selection, 1) = "." Or _
       Left(Selection, 1) = ":" Or _
       Left(Selection, 1) = ";" Or _
       Left(Selection, 1) = "(" Or _
       Left(Selection, 1) = ")" Or _
       Left(Selection, 1) = "!" Or _
       Left(Selection, 1) = "'" Or _
       Left(Selection, 1) = """" Or _
       Left(Selection, 1) = "¡°" Or _
       Left(Selection, 1) = "¡±" Or _
       Left(Selection, 1) = "¡®" Or _
       Left(Selection, 1) = "¡¯" Or _
       Left(Selection, 1) = "-" Or _
       Left(Selection, 1) = "?" Then
        Selection.Font.Color = RGB(255, 0, 0)
    Else
        Selection.Font.Color = RGB(255 * c2.R, 255 * c2.G, 255 * c2.B)
    End If
    Selection.MoveRight Unit:=wdCharacter, Count:=1
End Sub
