Attribute VB_Name = "Module1"
Public Sub CenterC(frm As Form)
    Dim SpcF As Integer 'Jumlah spasi yg dapat muat
    Dim clen As Integer 'Panjang tulisan
    Dim oldc As String 'Tulisan yg lama
    Dim i As Integer
    oldc = frm.Caption
    Do While Left(oldc, 1) = Space(1)
        DoEvents
        oldc = Right(oldc, Len(oldc) - 1)
    Loop
    Do While Right(oldc, 1) = Space(1)
        DoEvents
        oldc = Left(oldc, Len(oldc) - 1)
    Loop
    clen = Len(oldc)
    If InStr(oldc, "!") <> 0 Then
        If InStr(oldc, " ") <> 0 Then
            clen = clen * 1.5
        Else
            clen = clen * 1.4
        End If
    Else
        If InStr(oldc, " ") <> 0 Then
            clen = clen * 1.4
         Else
            clen = clen * 1.3
        End If
    End If
    'Periksa berapa karakter dapat muat
    SpcF = frm.Width / 61.2244 'Berapa banyak ruang yg
    'tersedia di caption tersebut
    SpcF = SpcF - clen
    If SpcF > 1 Then
        DoEvents 'Mempercepat program
        frm.Caption = Space(Int(SpcF / 2)) + oldc
    Else 'Jika form terlalu kecil untuk spasi
        frm.Caption = oldc
    End If
End Sub

