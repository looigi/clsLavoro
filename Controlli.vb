Imports System.Web.UI
Imports System.Web.UI.WebControls

Public Class Controlli
    Public Function ControllaCampoTesto(Cosa As String, NomeCampo As String, Obbligatorio As Boolean) As String
        Dim Ritorno As String = ""

        If Cosa.Trim = "" Then
            If Obbligatorio = True Then
                Ritorno = "<li>Campo '" & NomeCampo & "' non immesso</li>"
            End If
        End If

        Return Ritorno
    End Function

    Public Function ControllaCampoCombo(Cosa As String, NomeCampo As String, Obbligatorio As Boolean) As String
        Dim Ritorno As String = ""

        If Cosa.Trim = "" Then
            If Obbligatorio = True Then
                Ritorno = "<li>Campo '" & NomeCampo & "' non immesso</li>"
            End If
        End If

        Return Ritorno
    End Function

    Public Function ControllaCampoMail(Cosa As String, NomeCampo As String, Obbligatorio As Boolean) As String
        Dim Ritorno As String = ""

        If Cosa.Trim = "" Then
            If Obbligatorio = True Then
                Ritorno += "<li>Campo '" & NomeCampo & "' non immesso</li>"
            End If
        Else
            If Cosa.Length < 5 Or Cosa.IndexOf(".") = -1 Or Cosa.IndexOf("@") = -1 Then
                Ritorno += "<li>Campo '" & NomeCampo & "' non valido</li>"
            End If
        End If

        Return Ritorno
    End Function

    Public Function ControllaCampoNumerico(Cosa As String, NomeCampo As String, Obbligatorio As Boolean) As String
        Dim Ritorno As String = ""

        If Cosa.Trim = "" Then
            If Obbligatorio = True Then
                Ritorno = "<li>Campo '" & NomeCampo & "' non immesso</li>"
            End If
        Else
            If IsNumeric(Cosa) = False Then
                Ritorno = "<li>Campo '" & NomeCampo & "' non valido</li>"
            End If
        End If

        Return Ritorno
    End Function

    Public Function ControllaCampoData(Cosa As String, NomeCampo As String, Obbligatorio As Boolean) As String
        Dim Ritorno As String = ""

        If Cosa.Trim = "" Then
            If Obbligatorio = True Then
                Ritorno = "<li>Campo '" & NomeCampo & "' non immesso</li>"
            End If
        Else
            If IsDate(Cosa) = False Then
                Ritorno = "<li>Campo '" & NomeCampo & "' non valido</li>"
            End If
        End If

        Return Ritorno
    End Function

    Private Function FindControlRecursive(ByVal list As List(Of Control), ByVal parent As Control, ByVal ctrlType As System.Type) As List(Of Control)
        If parent Is Nothing Then Return list

        'If parent.GetType Is ctrlType Then
        list.Add(parent)
        'End If
        For Each child As Control In parent.Controls
            FindControlRecursive(list, child, ctrlType)
        Next

        Return list
    End Function

    Public Function PrendeValoriDiConfronto(NomePagina As Object) As String()
        Dim Confronto() As String = {}
        Dim qConfronto As Integer = 0

        Dim allTxtD As New List(Of Control)
        For Each cmb As Object In FindControlRecursive(allTxtD, NomePagina, GetType(Object))
        Next

        For i As Integer = 0 To allTxtD.Count - 1
            Dim t As Object = allTxtD(i)
            Dim Nome As String = ""

            Try
                Nome = t.id.toupper
            Catch ex As Exception

            End Try

            If Nome <> "" And Nome.IndexOf("_") = -1 Then
                If Nome.IndexOf("CMB") > -1 Or Nome.IndexOf("TXT") > -1 Or Nome.IndexOf("CHK") > -1 Then
                    ReDim Preserve Confronto(qConfronto)
                    Confronto(qConfronto) = t.ID & ";" & t.Text
                    qConfronto += 1
                End If
            End If
        Next

        If qConfronto = 0 Then
            ReDim Preserve Confronto(qConfronto)
            Confronto(qConfronto) = "NESSUN CAMPO TESTO"
        End If

        Return Confronto
    End Function

    Public Function ControntaValoriImmessiEVecchi(Confronto() As String, Vecchivalori() As String) As String()
        Dim Variazioni() As String = {}
        Dim qVariazioni As Integer = 0
        Dim Campetti1() As String
        Dim Campetti2() As String

        For i As Integer = 0 To Confronto.Length - 1
            Confronto(i) += ";"
            Vecchivalori(i) += ";"

            Campetti1 = Confronto(i).Split(";")
            Campetti2 = Vecchivalori(i).Split(";")

            If Campetti1(1).ToUpper.Trim.Replace(Chr(13), "").Replace(Chr(10), "") <> Campetti2(1).ToUpper.Trim.Replace(Chr(13), "").Replace(Chr(10), "") Then
                ReDim Preserve Variazioni(qVariazioni)
                Variazioni(qVariazioni) = i & ";" & Campetti1(0) & ";" & Campetti1(1) & ";" & Campetti2(1) & ";"
                qVariazioni += 1
            End If
        Next
        If qVariazioni = 0 Then
            ReDim Preserve Variazioni(qVariazioni)
            Variazioni(0) = "NESSUN CAMPO VARIATO"
        End If

        Return Variazioni
    End Function
End Class
