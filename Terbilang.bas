Function TerbilangAngka(n As Long) As String 'max 2.147.483.647
    Dim satuan As Variant, Minus As Boolean
    On Error GoTo terbilang_error
    satuan = Array("", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas")
    
    If n < 0 Then
        Minus = True
        n = n * -1
    End If
    
    Select Case n
        Case 0 To 11
            TerbilangAngka = " " + satuan(Fix(n))
        Case 12 To 19
            TerbilangAngka = TerbilangAngka(n Mod 10) + " Belas"
        Case 20 To 99
            TerbilangAngka = TerbilangAngka(Fix(n / 10)) + " Puluh" + TerbilangAngka(n Mod 10)
        Case 100 To 199
            TerbilangAngka = " Seratus" + TerbilangAngka(n - 100)
        Case 200 To 999
            TerbilangAngka = TerbilangAngka(Fix(n / 100)) + " Ratus" + TerbilangAngka(n Mod 100)
        Case 1000 To 1999
            TerbilangAngka = " Seribu" + TerbilangAngka(n - 1000)
        Case 2000 To 999999
            TerbilangAngka = TerbilangAngka(Fix(n / 1000)) + " Ribu" + TerbilangAngka(n Mod 1000)
        Case 1000000 To 999999999
            TerbilangAngka = TerbilangAngka(Fix(n / 1000000)) + " Juta" + TerbilangAngka(n Mod 1000000)
        Case Else
            TerbilangAngka = TerbilangAngka(Fix(n / 1000000000)) + " Milyar" + TerbilangAngka(n Mod 1000000000)
    End Select
    
    If Minus = True Then
        TerbilangAngka = "Minus" + TerbilangAngka
    End If
    
    Exit Function
terbilang_error:
    MsgBox Err.Description, vbCritical, "TerbilangAngka Error"
End Function

Function Terbilang(n As Long) As String
    Terbilang = UCase(Trim(TerbilangAngka(n)) + " Rupiah")
End Function