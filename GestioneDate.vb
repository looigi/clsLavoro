Public Class GestioneDate
    Public Function ControllaFestivo(Datella As Date) As Boolean
        Dim Ritorno As Boolean = False
        Dim Giorno As Integer = Datella.Day
        Dim Mese As Integer = Datella.Month

        Dim DataPasqua As Date = Pasqua(Datella.Year)
        Dim DatellaDopo As Date = DataPasqua.AddDays(1)
        Dim GiornoDopo As Integer = DatellaDopo.Day
        Dim Mesedopo As Integer = DatellaDopo.Month

        Select Case Mese
            Case 1
                If Giorno = 1 Or Giorno = 6 Then
                    Ritorno = True
                End If
            Case 2
            Case 3
            Case 4
                If Giorno = 25 Then
                    Ritorno = True
                End If
            Case 5
                If Giorno = 1 Then
                    Ritorno = True
                End If
            Case 6
                If Giorno = 2 Or Giorno = 29 Then
                    Ritorno = True
                End If
            Case 7
            Case 8
                If Giorno = 15 Then
                    Ritorno = True
                End If
            Case 9
            Case 10
            Case 11
                If Giorno = 1 Then
                    Ritorno = True
                End If
            Case 12
                If Giorno = 8 Or Giorno = 25 Or Giorno = 26 Then
                    Ritorno = True
                End If
        End Select

        If Ritorno = False Then
            ' Controllo Pasquetta
            If Giorno = GiornoDopo And Mese = Mesedopo Then
                Ritorno = True
            Else
                ' Controllo sabato o domenica
                If Datella.DayOfWeek = 0 Or Datella.DayOfWeek = 6 Then
                    Ritorno = True
                End If
            End If
        End If

        Return Ritorno
    End Function

    Public Function Pasqua(Anno%) As Date
        Dim A%, b%, c%, p%, q%, R%
        Dim Pasq As Date

        A = Anno% Mod 19 : b = Anno% \ 100 : c = Anno% Mod 100
        p = (19 * A + b - (b / 4) - ((b + ((b + 8) \ 25) + 1) \ 3) + 15) Mod 30
        q = (32 + 2 * ((b Mod 4) + (c \ 4)) - p - (c Mod 4)) Mod 7
        R = (p + q - 7 * ((A + 11 * p + 22 * q) \ 451) + 114)
        Pasq = DateSerial(Anno%, R \ 31, (R Mod 31) + 1)

        Return Pasqua
    End Function
End Class
