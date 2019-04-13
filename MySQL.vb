Imports MySql.Data.MySqlClient
Imports System.Web.UI.WebControls
Imports System.Data.SqlClient

Public Class MySQL
    'Dim ConnMySQL As New GestioneMySQL

    'ConnMySQL.ImpostaConnessione("10.168.146.33", "amdb", "$amdb1", "amdb")
    'ConnMySQL.ApreConnessione()
    'ConnMySQL.ApreRecordset("Select * From Buttami")
    'Dim righe As Long = ConnMySQL.ContaRigheLette
    'Dim colonne As Integer = ConnMySQL.ContaColonneLette
    'Dim Campo() As String = ConnMySQL.RitornaRiga
    'ConnMySQL.MoveNext()
    'Campo = ConnMySQL.RitornaRiga
    'ConnMySQL.ChiudeRecordset()
    'ConnMySQL.EsegueQuery("Insert Into Buttami Values (3,'mmmm')")
    'ConnMySQL.ChiudeConnessione()

    Structure Connessione
        Dim NomeServer As String
        Dim Utente As String
        Dim Password As String
        Dim Database As String
    End Structure

    Private Conness As Connessione
    Private MysqlConn As SqlConnection
    Private recordSet As DataTable
    Private NumeroRighe As Long
    Private NumeroColonne As Long
    Private RigaAttuale As Long
    Private Errore As String

    Public Sub New(Server As String, Utente As String, Password As String, SchemaName As String)
        Conness.NomeServer = Server
        Conness.Utente = Utente
        Conness.Password = Password
        Conness.Database = SchemaName

        ApreConnessione()
    End Sub

    Public Sub ApreConnessione()
        MysqlConn = New SqlConnection
        MysqlConn.ConnectionString = "server=" & Conness.NomeServer & ";" _
            & "user id=" & Conness.Utente & ";" _
            & "password=" & Conness.Password & ";" _
            & "database=" & Conness.Database & ";"
        Try
            MysqlConn.Open()
        Catch ex As Exception

        End Try
    End Sub

    Public Sub ApreRecordset(Sql As String)
        Errore = ""

        Try
            Dim aCommand As SqlCommand = New SqlCommand(Sql, MysqlConn)
            Dim aReader As SqlDataReader = aCommand.ExecuteReader()

            recordSet = New DataTable
            recordSet.Load(aReader)

            NumeroColonne = recordSet.Columns.Count
            NumeroRighe = recordSet.Rows.Count

            If NumeroRighe > 0 Then
                RigaAttuale = 0
            Else
                RigaAttuale = -1
            End If

        Catch ex As Exception
            RigaAttuale = -2
            Errore = ex.Message
        End Try
    End Sub

    Public Sub ChiudeRecordset()
        Try
            recordSet.Dispose()
        Catch ex As Exception

        End Try
    End Sub

    Public Function RitornaRiga() As String()
        If RigaAttuale < 0 Then
            Dim Ritorno(0) As String

            Ritorno(0) = "ERRORE: FINE RECORDSET"

            Return Ritorno
        Else
            Dim Ritorno(NumeroColonne - 1) As String
            Dim Colonna As Integer = 0

            Dim row As DataRow = recordSet.Rows(RigaAttuale)

            For Each column In recordSet.Columns
                Ritorno(Colonna) = row(column).ToString

                Colonna += 1
            Next

            Return Ritorno
        End If
    End Function

    Public Sub MoveNext()
        RigaAttuale += 1
        If RigaAttuale > NumeroRighe Then
            RigaAttuale = -1
        End If
    End Sub

    Public Sub MoveFirst()
        RigaAttuale = 0
    End Sub

    Public Sub MoveLast()
        RigaAttuale = NumeroRighe
    End Sub

    Public Sub MovePrevious()
        RigaAttuale -= 1
        If RigaAttuale < 0 Then
            RigaAttuale = 0
        End If
    End Sub

    Public Function ContaRigheLette() As Long
        Return NumeroRighe
    End Function

    Public Function ContaColonneLette() As Long
        Return NumeroColonne
    End Function

    Public Function EsegueQuery(Sql As String) As String
        Try
            Dim objComm As SqlCommand = New SqlCommand(Sql, MysqlConn)
            Dim i As Integer = objComm.ExecuteNonQuery()
            Dim Messaggio As String = ""

            If (i > 0) Then
                ' "Operazione OK"
                Messaggio = "OK"
            Else
                ' "Operazione KO"
                Messaggio = "ERRORE: Codice " & i
            End If

            Return Messaggio
        Catch ex As Exception
            Return "ERRORE: " & ex.Message
        End Try
    End Function

    Public Sub ChiudeConnessione()
        MysqlConn.Close()
        MysqlConn.Dispose()
    End Sub

    Public Function RitornaQuery(Sql As String) As String()
        'ApreConnessione()
        ApreRecordset(Sql)
        Dim Campo() As String
        If RigaAttuale = -2 Then
            ReDim Preserve Campo(0)
            Campo(0) = "ERRORE: " & Errore
        Else
            Campo = RitornaRiga()
        End If
        ChiudeRecordset()
        'ChiudeConnessione()

        Return Campo
    End Function

    ' ------------------------------------------------------
    ' Routine per la gestione delle griglie
    ' ------------------------------------------------------

    Private Colonne() As DataColumn
    Private riga As DataRow
    Private dttTabella As New DataTable()
    Private QuantiCampi As Integer
    Private AggiunteRighe As Boolean = False

    Public Function ImpostaCampi(Sql As String, Griglia As GridView, Optional NonEseguireSubito As Boolean = False) As Integer
        Dim q As Integer = 0
        Dim nCampi() As String = {}
        Dim Appo As String

        For i As Integer = 0 To Griglia.Columns.Count - 1
            Try
                Appo = DirectCast(Griglia.Columns(i), System.Web.UI.WebControls.BoundField).DataField
            Catch ex As Exception
                Appo = ""
            End Try

            If Appo <> "" Then
                ReDim Preserve nCampi(q)
                nCampi(q) = DirectCast(Griglia.Columns(i), System.Web.UI.WebControls.BoundField).DataField
                q += 1
            End If
        Next
        QuantiCampi = q - 1

        ReDim Preserve Colonne(UBound(nCampi))
        QuantiCampi = UBound(nCampi)

        For i As Integer = 0 To QuantiCampi
            Colonne(i) = New DataColumn(nCampi(i))

            dttTabella.Columns.Add(Colonne(i))
        Next

        Dim RigheInserite As Integer = ImpostaValori(Sql)
        If NonEseguireSubito = False Then
            VisualizzaValori(Griglia)
            ChiudeDB()
        Else
            AggiunteRighe = True
        End If

        Return RigheInserite
    End Function

    Private Function ImpostaValori(Sql As String) As Integer
        Dim NumeroRighe As Integer = 0

        ApreRecordset(Sql)
        Dim righe As Long = ContaRigheLette() - 1
        Dim colonne As Integer = ContaColonneLette() - 1

        For i As Integer = 0 To righe
            Dim Campo() As String = RitornaRiga()

            riga = dttTabella.NewRow()
            For k As Integer = 0 To colonne
                riga(k) = "" & Campo(k).ToString.Replace("&#39;", "'").Replace("&#224;", "'")

                'If IsDate(Campo) = True And Campo.Length > 6 Then
                '    Dim dCampo As Date = Campo

                '    Campo = ConverteData(dCampo)
                'End If

                'riga(i) = Campo
            Next
            dttTabella.Rows.Add(riga)

            NumeroRighe += 1

            MoveNext()
        Next

        Return NumeroRighe
    End Function

    Private Sub ChiudeDB()
        riga = Nothing
        dttTabella = Nothing

        ChiudeConnessione()
    End Sub

    Public Sub VisualizzaValori(grdView As GridView)
        grdView.DataSource = dttTabella
        grdView.DataBind()

        If AggiunteRighe = True Then
            ChiudeDB()
        End If
    End Sub

End Class
