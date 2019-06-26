Attribute VB_Name = "modGenerator"
Public radioZona As Integer

Sub main()
    radioZona = 1
    frmGlavna.txtRazmak.Text = " "
    frmGlavna.Show
    frmGlavna.txtUlaznaDatoteka = App.Path
    frmGlavna.txtIzlaznaDatoteka = App.Path
End Sub


Public Function TockaZarez(ulazniString As String) As String
Dim j As Integer

    For j = 1 To Len(ulazniString)
        If (Mid(ulazniString, j, 1) = ",") Or (Mid(ulazniString, j, 1) = ".") Then
            Mid(ulazniString, j) = "."
        End If
    Next j
    
    TockaZarez = ulazniString
    
End Function



Public Function Trans(X As Double, Y1 As Double) As Boolean
    Const PI As Double = 3.14159265358979
    Dim ro As Double
    Dim a As Double
    Dim b As Double
    Dim y As Double
    Dim bx As Double
    Dim EMI As Double
    Dim E As Double
    Dim e1 As Double
    Dim a0 As Double
    Dim b0 As Double
    Dim c0 As Double
    Dim d0 As Double
    Dim yy As Double
    Dim js As Double
    Dim fi1 As Double
    Dim d As Double
    Dim t As Double
    Dim eta As Double
    Dim En As Double
    Dim fif1 As Double
    Dim fif2 As Double
    Dim elf1 As Double
    Dim elf2 As Double
    Dim elf3 As Double
    Dim fi As Double
    Dim el As Double
    Dim alo As Double
    Dim al As Double
    Dim yf1 As Double
    Dim yf2 As Double
    Dim yf3 As Double
    Dim xf1 As Double
    Dim xf2 As Double
    Dim E0 As Double
    Dim F0 As Double
    Dim FIF3 As Double
    Dim XF3 As Double
    Dim TR As Integer
        

    Trans = False
    ro = 180 / PI
    a = 6377397.155
    b = 6356078.96325
    EMI = 0
    If b = 0 Then b = a * (1 - EMI)
    E = Sqr((a ^ 2 - b ^ 2) / a ^ 2)
    e1 = Sqr((a ^ 2 - b ^ 2) / b ^ 2)
    a0 = 1 + 3 * E ^ 2 / 4 + 45 * E ^ 4 / 64 + 175 * E ^ 6 / 256 + 11025 * E ^ 8 / 16384 + 43659 * E ^ 10 / 65536
    b0 = 3 * E ^ 2 / 4 + 15 * E ^ 4 / 16 + 525 * E ^ 6 / 512 + 2205 * E ^ 8 / 2048 + 72765 * E ^ 10 / 65536
    c0 = 15 * E ^ 4 / 64 + 105 * E ^ 6 / 256 + 2205 * E ^ 8 / 4096 + 10395 * E ^ 10 / 16384
    d0 = 35 * E ^ 6 / 512 + 315 * E ^ 8 / 2048 + 31185 * E ^ 10 / 131072
    E0 = 315 * E ^ 8 / 16384 + 3465 * E ^ 10 / 65536
    F0 = 693 * E ^ 10 / 131072
    
    yy = Int(Y1 / 1000000)
    y = Y1 - yy * 1000000 - 500000

    
    If (yy = 5) Or (yy = 6) Or (yy = 7) Then
        y = y / 0.9999: X = X / 0.9999
    ElseIf (yy = 2) Or (yy = 3) Then
        y = y / 0.9997: X = X / 0.9997
    Else
        MsgBox "Nije moguæe izvesti odabranu transformaciju. Provjerite koordinate."
        Exit Function
    End If
    
    
    TR = 0
    If (Y1 < 6000000) And (yy = 5) Then TR = 56
    If (Y1 <= 6500000) And (Y1 >= 6000000) Then TR = 65
    If (Y1 > 6500000) And (Y1 < 7000000) Then TR = 67
    If Y1 >= 7000000 Then TR = 76
    If (yy = 2) Or (yy = 3) Then TR = 1630
    
    Select Case TR
        Case 56
            js = 15
        Case 65, 67
            js = 18
        Case 76
            js = 21
        Case 1630
            js = 16.5
        Case Else
              MsgBox "Koordinate elemenata nisu u petoj ili sestoj zoni. Prepravite koordinate u 5 XXX XXX.XXX ili 6 XXX XXX.XXX"
            Exit Function
    End Select
    

    fi1 = 2 * X / (a + b)
    Do
        bx = a * (1 - E ^ 2) * (a0 * fi1 - b0 * Sin(2 * fi1) / 2 + c0 * Sin(4 * fi1) / 4 - d0 * Sin(6 * fi1) / 6 + E0 * Sin(8 * fi1) / 8 - F0 * Sin(10 * fi1) / 10)
        d = X - bx
        fi1 = fi1 + 2 * d / (a + b)
    Loop Until Abs(d) < 0.0001

    t = Tan(fi1)
    eta = e1 * Cos(fi1)
    En = a / Sqr(1 - E ^ 2 * (Sin(fi1)) ^ 2)
    fif1 = -t * (1 + eta ^ 2) / (2 * En ^ 2)
    fif2 = t * (5 + 3 * t ^ 2 + 6 * eta ^ 2 - 6 * t ^ 2 * eta ^ 2) / (24 * En ^ 4)
    FIF3 = -t * (61 + 90 * t ^ 2 + 45 * t ^ 4 + 107 * eta ^ 2 - 162 * t ^ 2 * eta ^ 2 - 45 * t ^ 4 * eta ^ 2) / (720 * En ^ 6)
    elf1 = 1 / (En * Cos(fi1))
    elf2 = -(1 + 2 * t ^ 2 + eta ^ 2) / (6 * En ^ 3 * Cos(fi1))
    elf3 = (5 + 28 * t ^ 2 + 24 * t ^ 4 + 6 * eta ^ 2 + 8 * t ^ 2 * eta ^ 2) / (120 * En ^ 5 * Cos(fi1))
    fi = fi1 + fif1 * y ^ 2 + fif2 * y ^ 4 + FIF3 * y ^ 6
    el = elf1 * y + elf2 * y ^ 3 + elf3 * y ^ 5
    alo = js / ro
    al = alo + el


    Select Case radioZona
        Case 1
            If (TR = 56) Or (TR = 76) Then
                el = al - 18 / ro
            ElseIf TR = 65 Then
                el = al - 15 / ro
            ElseIf TR = 67 Then
                el = al - 21 / ro
            End If
        Case 2
            el = al - 16.5 / ro
        Case 3
            el = al - 15 / ro
        Case 4
            el = al - 18 / ro
        Case 5
            el = al - 21 / ro
    End Select

    bx = a * (1 - E ^ 2) * (a0 * fi - b0 * Sin(2 * fi) / 2 + c0 * Sin(4 * fi) / 4 - d0 * Sin(6 * fi) / 6 + E0 * Sin(8 * fi) / 8 - F0 * Sin(10 * fi) / 10)
    t = Tan(fi)
    eta = e1 * Cos(fi)
    En = a / Sqr(1 - E ^ 2 * (Sin(fi)) ^ 2)
    yf1 = En * Cos(fi)
    yf2 = En * (Cos(fi)) ^ 3 * (1 - t ^ 2 + eta ^ 2) / 6
    yf3 = En * (Cos(fi)) ^ 5 * (5 - 18 * t ^ 2 + t ^ 4 + 14 * eta ^ 2 - 58 * t ^ 2 * eta ^ 2) / 120
    xf1 = En * Sin(fi) * Cos(fi) / 2
    xf2 = En * Sin(fi) * (Cos(fi)) ^ 3 * (5 - t ^ 2 + 9 * eta ^ 2 + 4 * eta ^ 4) / 24
    XF3 = En * Sin(fi) * (Cos(fi)) ^ 5 * (61 - 58 * t ^ 2 + t ^ 4 + 270 * eta ^ 2 - 330 * t ^ 2 * eta ^ 2) / 720
    y = yf1 * el + yf2 * el ^ 3 + yf3 * el ^ 5
    
    Select Case radioZona
        Case 1, 3, 4, 5
            y = y * 0.9999
        Case 2
            y = y * 0.9997
    End Select
    
    X = bx + xf1 * el ^ 2 + xf2 * el ^ 4 + XF3 * el ^ 6
    
    Select Case radioZona
        Case 1, 3, 4, 5
            X = X * 0.9999
        Case 2
            X = X * 0.9997
    End Select
    

    Select Case radioZona
        Case 1
            If (TR = 56) Or (TR = 76) Then y = y + 6500000
            If TR = 65 Then y = y + 5500000
            If TR = 67 Then y = y + 7500000
        Case 2
            y = y + 3500000
        Case 3
            y = y + 5500000
        Case 4
            y = y + 6500000
        Case 5
            y = y + 7500000
    End Select
        
    
    X = X
    Y1 = y
    Trans = True

End Function

Public Sub RadiNesto()
Dim dijelovi() As String
Dim dijelovi2() As String
Dim linija As String
Dim broj As String
Dim ykoord As Double
Dim xkoord As Double
Dim zkoord As String
Dim i As Integer
Dim j As Integer
Dim xKoordinataText As String
Dim yKoordinataText As String
Dim pozicija() As Integer



On Error GoTo ErrorHandler

Open frmGlavna.txtUlaznaDatoteka.Text For Input Access Read As 1
Open frmGlavna.txtIzlaznaDatoteka.Text For Output Access Write As 2
Do While Not EOF(1)
    Input #1, linija
    dijelovi = Split(linija, " ")
    i = 0
    For j = LBound(dijelovi) To UBound(dijelovi)
        If dijelovi(j) <> "" Then
            i = i + 1
            ReDim Preserve dijelovi2(i)
            ReDim Preserve pozicija(i)
            dijelovi2(i) = dijelovi(j)
            pozicija(i) = j
        Else
            dijelovi(j) = " "
        End If
    Next j
    ykoord = Val(TockaZarez(dijelovi2(2)))
    xkoord = Val(TockaZarez(dijelovi2(3)))
    rezultat = Trans(xkoord, ykoord)

    yKoordinataText = TockaZarez(Trim(Str(Round(ykoord, 3))))
    xKoordinataText = TockaZarez(Trim(Str(Round(xkoord, 3))))
    linija = ""
    If rezultat Then
        dijelovi(pozicija(2)) = yKoordinataText
        dijelovi(pozicija(3)) = xKoordinataText
        For j = LBound(dijelovi) To UBound(dijelovi)
            linija = linija + dijelovi(j)
        Next j
    Print #2, linija
    End If
    
Loop

MsgBox "Transformacija završena"
Close #2
Close #1
Exit Sub

ErrorHandler:

MsgBox "Pogreška pri radu sa odabranim datotekama"
Close #2
Close #1
Exit Sub

End Sub

