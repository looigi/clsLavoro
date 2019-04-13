Imports System.Web

Public Class clsCookie
    Inherits System.Web.UI.Page

    ' Private NomeProc As String = "SSDCVCALCIO"

    Public Sub CreateCookies(NomeProc As String, Resp As HttpResponse, NomeCookie As String, Valore As String)
        Try
            Dim myCookie As HttpCookie = New HttpCookie(NomeProc)
            myCookie(NomeCookie) = Valore
            myCookie.Expires = DateTime.Now.AddDays(30)
            Resp.Cookies.Add(myCookie)
        Catch ex As System.Web.HttpException

        Catch ex As Exception

        End Try
    End Sub

    Public Function LeggeCookie(NomeProc As String, Richiesta As HttpRequest, NomeCookie As String) As String
        Dim Ritorno As String = ""

        Try
            If (Richiesta.Cookies(NomeProc) IsNot Nothing) Then
                If (Richiesta.Cookies(NomeProc)(NomeCookie) IsNot Nothing) Then
                    Ritorno = Richiesta.Cookies(NomeProc)(NomeCookie)
                End If
            End If
        Catch ex As System.Web.HttpException

        Catch ex As Exception

        End Try

        Return Ritorno
    End Function

    Public Sub EliminaCookie(NomeProc As String, Resp As HttpResponse, NomeCookie As String)
        Dim myCookie As HttpCookie

        Try
            myCookie = New HttpCookie(NomeProc)
            myCookie(NomeCookie) = ""
            myCookie.Expires = DateTime.Now.AddDays(-1D)
            Resp.Cookies.Add(myCookie)
        Catch ex As System.Web.HttpException

        Catch ex As Exception

        End Try
    End Sub
End Class
