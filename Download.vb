Imports System.IO
Imports System.Net
Imports System.Text
Imports System.ComponentModel
Imports System.Windows.Forms

Public Class Download
    Private TipoCollegamento As String
    Private Utenza As String
    Private Password As String
    Private Dominio As String
    Private sNomeFileConPercorso As String

    Public Sub ImpostaValori(Tc As String, U As String, P As String, D As String, nf As String)
        TipoCollegamento = Tc
        Utenza = U
        Password = P
        Dominio = D
        sNomeFileConPercorso = nf
    End Sub

    Public Sub CreaEPulisceCartellaDiLavoro()
        Dim gf As New GestioneFilesDirectory
        gf.CreaDirectoryDaPercorso(gf.TornaNomeDirectoryDaPath(sNomeFileConPercorso) & "\")

        gf.ScansionaDirectorySingola(gf.TornaNomeDirectoryDaPath(sNomeFileConPercorso) & "\")
        Dim Filetti() As String = gf.RitornaFilesRilevati
        Dim qFiles As Integer = gf.RitornaQuantiFilesRilevati
        For i As Integer = 1 To qFiles
            gf.EliminaFileFisico(Filetti(i))
        Next

        If File.Exists(sNomeFileConPercorso) = True Then
            Dim Conta As Integer = 1
            Dim Estensione As String = gf.TornaEstensioneFileDaPath(sNomeFileConPercorso)

            sNomeFileConPercorso = sNomeFileConPercorso.Replace(Estensione, "")

            Do While File.Exists(sNomeFileConPercorso & Format(Conta, "00") & Estensione) = True
                Conta += 1
            Loop

            sNomeFileConPercorso = sNomeFileConPercorso & Format(Conta, "00") & Estensione
        End If
    End Sub

    Public Function ScaricaPagina(Url As String, Optional pbDown As ProgressBar = Nothing) As Boolean
        Dim Ok As Boolean = True
        Dim sourceCode As String
        Dim gf As New GestioneFilesDirectory

        If TipoCollegamento Is Nothing = True Then TipoCollegamento = ""

        If TipoCollegamento.Trim.ToUpper = "PROXY" Then
            Dim request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(Url)
            request.Proxy.Credentials = New System.Net.NetworkCredential(Utenza, Password, Dominio)
            request.Timeout = 7000
            Try
                Dim response As System.Net.HttpWebResponse = request.GetResponse()
                'Application.DoEvents()
                Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
                'Application.DoEvents()
                sourceCode = sr.ReadToEnd()
                sr.Close()
                response.Close()
                request = Nothing

                gf.CreaAggiornaFile(sNomeFileConPercorso, sourceCode)
            Catch ex As Exception
                Ok = False
            End Try
        Else
            Try
                'Dim wreq As WebRequest = WebRequest.Create(Url)
                'Dim wres As WebResponse = wreq.GetResponse()
                'Dim iBuffer As Integer = 0
                'Dim buffer(128) As [Byte]
                'Dim stream As Stream = wres.GetResponseStream()
                'iBuffer = stream.Read(buffer, 0, 128)
                'Dim strRes As New StringBuilder("")
                'While iBuffer <> 0
                '    strRes.Append(Encoding.ASCII.GetString(buffer, 0, iBuffer))
                '    iBuffer = stream.Read(buffer, 0, 128)
                'End While
                'gf.CreaAggiornaFile(sNomeFileConPercorso, strRes.ToString())

                ScaricaFileAsincrono(Url, sNomeFileConPercorso, pbDown)
            Catch ex As Exception
                Ok = False
            End Try
        End If

        Return Ok
    End Function

    Public Function ScaricaFile(Url As String, Optional sNomeDestinazione As String = "") As Boolean
        Dim Ok As Boolean = True
        ' Dim sourceCode As String
        If sNomeDestinazione = "" Then sNomeDestinazione = sNomeFileConPercorso
        Dim gf As New GestioneFilesDirectory

        If TipoCollegamento Is Nothing = True Then TipoCollegamento = ""

        If TipoCollegamento.Trim.ToUpper = "PROXY" Then
            Dim request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(Url)
            request.Proxy.Credentials = New System.Net.NetworkCredential(Utenza, Password, Dominio)
            Dim response As System.Net.HttpWebResponse = request.GetResponse()
            'Application.DoEvents()

            Dim responseStream As Stream = response.GetResponseStream()
            Dim imageBytes() As Byte

            Using br As New BinaryReader(responseStream)
                imageBytes = br.ReadBytes(500000)
                br.Close()
            End Using
            responseStream.Close()
            response.Close()

            Dim fs As New FileStream(sNomeFileConPercorso, FileMode.Create)
            Dim bw As New BinaryWriter(fs)
            Try
                bw.Write(imageBytes)
            Finally
                fs.Close()
                bw.Close()
            End Try

            request = Nothing

            response.Close()
            response = Nothing
            request = Nothing
        Else
            Dim myWebClient As New WebClient()

            Try
                myWebClient.DownloadFile(
                            Url,
                            sNomeDestinazione)
            Catch ex As Exception

            End Try
        End If

        Return Ok
    End Function

    Private StaScaricando As Boolean
    Private pbDown As ProgressBar = Nothing
    Private Blocca As Boolean
    Private wb As New WebClient
    Private sNomeFiletto As String
    Private Inizio As Date

    Public Sub ScaricaFileAsincrono(Url As String, sNomeFile As String, Optional pbDownload As ProgressBar = Nothing)
        StaScaricando = True
        Blocca = False
        Inizio = Now

        Try
            Dim Uri As New Uri(Url)
            AddHandler wb.DownloadFileCompleted, AddressOf Completed
            AddHandler wb.DownloadProgressChanged, AddressOf Scaricando

            If pbDownload Is Nothing = False Then
                pbDownload.Value = 0
                pbDownload.Visible = True

                pbDown = pbDownload
            End If

            sNomeFiletto = sNomeFile
            wb.DownloadFileAsync(Uri, sNomeFile)

            Do While StaScaricando = True
                Application.DoEvents()
            Loop
        Catch ex As Exception

        End Try
    End Sub

    Public Sub BloccaElaborazione()
        Blocca = True
        StaScaricando = False
    End Sub

    Private Sub Scaricando(ByVal sender As Object, ByVal e As DownloadProgressChangedEventArgs)
        If pbDown Is Nothing = False Then
            Try
                If DateDiff(DateInterval.Second, Inizio, Now) > 15 Then
                    Blocca = True
                End If

                Dim bytesIn As Double = Double.Parse(e.BytesReceived.ToString())
                Dim totalBytes As Double = Double.Parse(e.TotalBytesToReceive.ToString())
                Dim percentage As Double = bytesIn / totalBytes * 100

                If Blocca = True Then
                    StaScaricando = False
                    wb.CancelAsync()
                    Exit Sub
                End If

                pbDown.Value = Int32.Parse(Math.Truncate(percentage).ToString())
                Application.DoEvents()
            Catch ex As Exception
                Application.DoEvents()
            End Try
        End If
    End Sub

    Private Sub Completed(sender As Object, e As AsyncCompletedEventArgs)
        'file downloaded
        StaScaricando = False

        If pbDown Is Nothing = False Then
            pbDown.Value = 0
        End If

        If e.Cancelled Then
            Try
                Kill(sNomeFiletto)
            Catch ex As Exception

            End Try
        End If
    End Sub
End Class
