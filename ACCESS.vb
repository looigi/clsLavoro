Public Class ACCESS
    Private Connessione As String

    ' ESEMPIO DI CHIAMATA
    'Dim g As New ACCESS
    'g.ImpostaNomeProvider("Microsoft.ACE.OLEDB.12.0")
    'g.ImpostaStringaDiConnessione("Data Source=" & PathTrasfSMS & ";Persist Security Info=False;")
    'If g.LeggeImpostazioniDiBase() Then
    '   Dim conn As Object = "ADODB.Connection"
    '   conn = g.ApreDB()
    '   Dim sql As String = "Select * From Notifiche"
    '   Dim rec As Object = "ADODB.Recordset"
    '   rec = g.LeggeQuery(conn, sql)
    '   rec.Close
    '   conn.Close
    'End If
    'g = Nothing

    Public Function ProvaConnessione() As String
        Dim Conn As Object = CreateObject("ADODB.Connection")

        Try
            Conn.Open(Connessione)
            Conn.Close()

            Conn = Nothing
            Return ""
        Catch ex As Exception

            Return ex.Message
        End Try
    End Function

    Private NomeProvider As String = ""
    Private StringaDiConnessione As String = ""

    Public Sub ImpostaNomeProvider(Nome As String)
        NomeProvider = Nome
    End Sub

    Public Sub ImpostaStringaDiConnessione(Stringa As String)
        StringaDiConnessione = Stringa
    End Sub

    Public Function LeggeImpostazioniDiBase() As Boolean
        Dim Ritorno As String
        Dim Ok As Boolean = True

        Dim Provider As String = NomeProvider
        Dim connectionString As String = StringaDiConnessione

        Connessione = "Provider=" & Provider & ";" & connectionString

        If Connessione = "" Then
            Ok = False
        Else
            Ritorno = ProvaConnessione()

            If Ritorno <> "" Then
                Ok = False
            End If
        End If

        Return Ok
    End Function

    Public Function ApreDB() As Object
        ' Routine che apre il DB e vede se ci sono errori
        Dim Conn As Object = CreateObject("ADODB.Connection")

        Try
            Conn.Open(Connessione)
            Conn.CommandTimeout = 0
        Catch ex As Exception
            'Dim H As HttpApplication = HttpContext.Current.ApplicationInstance
            'Dim StringaPassaggio As String

            'StringaPassaggio = "?Errore=Apertura DB"
            'StringaPassaggio = StringaPassaggio & "&Utente=" & H.Session("idUtente")
            'StringaPassaggio = StringaPassaggio & "&Chiamante=" & H.Request.CurrentExecutionFilePath.ToUpper.Trim
            'StringaPassaggio = StringaPassaggio & "&Sql=" & ex.Message
            'H.Response.Redirect(PercorsoApplicazione & "/Errore.aspx" & StringaPassaggio)
        End Try

        Return Conn
    End Function

    Private Function ControllaAperturaConnessione(ByRef Conn As Object) As Boolean
        Dim Ritorno As Boolean = False

        If Conn Is Nothing Then
            Ritorno = True
            Conn = ApreDB()
        End If

        Return Ritorno
    End Function

    Public Function EsegueSql(ByVal Conn As Object, ByVal Sql As String) As String
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn)
        Dim Ritorno As String = ""

        ' Routine che esegue una query sul db
        Try
            Conn.Execute(Sql)
        Catch ex As Exception
            Ritorno = "ERRORE:" & ex.Message
        End Try

        ChiudeDB(AperturaManuale, Conn)

        Return Ritorno
    End Function

    Public Function EsegueSqlSenzaTRY(ByVal Conn As Object, ByVal Sql As String) As String
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn)
        Dim Ritorno As String = ""

        Conn.Execute(Sql)

        ChiudeDB(AperturaManuale, Conn)

        Return Ritorno
    End Function

    Private Sub ChiudeDB(ByVal TipoApertura As Boolean, ByRef Conn As Object)
        If TipoApertura = True Then
            Conn.Close()
        End If
    End Sub

    Public Function LeggeQuery(ByVal Conn As Object, ByVal Sql As String) As Object
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn)
        Dim Rec As Object = CreateObject("ADODB.Recordset")

        Try
            Rec.Open(Sql, Conn)
        Catch ex As Exception
            Rec = Nothing
        End Try

        ChiudeDB(AperturaManuale, Conn)

        Return rec
    End Function

End Class
