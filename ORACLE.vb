Public Class ORACLE
    ' Routine di esempio per richiamare questa classe
    ' -----------------------------------------------------------
    'Dim gO As New GestioneORACLE
    'Dim Campo As String

    'gO.ImpostaParametriDiConnessione(sServerName, sSchemaName, sUserName, sPassword)
    'gO.ApreConnessione()
    'gO.LeggeQuery("Select * From MV_COMUNI")
    'gO.SpostaRecordsetAllaRiga(0)
    'gO.SpostaAvanti()
    'Campo = gO.RitornaCampo("DESC_PROVINCIA")
    'Response.Write("Numero righe: " & gO.QuanteRighe & "<br />")
    'Response.Write(Campo)
    'gO.ChiudeConnessione()

    Private ActiveConnection As New ttOledbConnection
    Private Recordset As System.Data.DataTable = Nothing
    Private RigaAttuale As System.Data.DataRow
    Private NumeroRiga As Long
    Private TotRighe As Long
    Private sServerName As String
    Private sSchemaName As String
    Private sUserName As String
    Private sPassword As String

    Public Sub ImpostaParametriDiConnessione(ServerName As String, SchemaName As String, UserName As String, Password As String)
        sServerName = ServerName
        sSchemaName = SchemaName
        sUserName = UserName
        sPassword = Password
    End Sub

    Public Sub ApreConnessione()
        ActiveConnection.OpenConnection(sServerName, sSchemaName, sUserName, sPassword)
    End Sub

    Public Function EsegueSQL(Sql As String) As String
        Dim Errore As String = ""

        Try
            ActiveConnection.Execute(Sql)
        Catch ex As Exception
            Errore = "ERRORE: " & ex.Message
        End Try

        Return Errore
    End Function

    Public Sub ApreTransazione()
        Try
            ActiveConnection.BeginTransaction()
        Catch ex As Exception

        End Try
    End Sub

    Public Sub CommittaTransazione()
        Try
            ActiveConnection.CommitTransaction()
        Catch ex As Exception

        End Try
    End Sub

    Public Sub RollbackTransazione()
        Try
            ActiveConnection.RollbackTransaction()
        Catch ex As Exception

        End Try
    End Sub

    Public Sub ChiudeConnessione()
        NumeroRiga = -1
        TotRighe = -1
        RigaAttuale = Nothing
        Recordset = Nothing

        ActiveConnection.CloseConnection()
    End Sub

    Public Sub LeggeQuery(Query As String)
        NumeroRiga = 0
        Try
            Recordset = ActiveConnection.GetOledbDataSet(Query).Tables(0)
            TotRighe = Recordset.Rows.Count
        Catch ex As Exception
            NumeroRiga = -1
            TotRighe = -1
            Recordset = Nothing
        End Try
    End Sub

    Public Function QuanteRighe() As Long
        Return TotRighe
    End Function

    Public Sub SpostaInizio()
        NumeroRiga = 0
        SpostaRecordsetAllaRiga(NumeroRiga)
    End Sub

    Public Sub SpostaFine()
        NumeroRiga = TotRighe
        SpostaRecordsetAllaRiga(NumeroRiga)
    End Sub

    Public Sub SpostaAvanti()
        NumeroRiga += 1
        If NumeroRiga <= TotRighe Then
            SpostaRecordsetAllaRiga(NumeroRiga)
        Else
            NumeroRiga = TotRighe
        End If
    End Sub

    Public Sub SpostaIndietro()
        NumeroRiga -= 1
        If NumeroRiga > 0 Then
            SpostaRecordsetAllaRiga(NumeroRiga)
        Else
            NumeroRiga = 0
        End If
    End Sub

    Public Sub SpostaRecordsetAllaRiga(NumeroRiga As Long)
        Try
            RigaAttuale = Recordset.Rows(NumeroRiga)
        Catch ex As Exception

        End Try
    End Sub

    Public Function RitornaCampo(NomeCampo As String) As String
        Dim Campo As String

        If IsNumeric(NomeCampo) = True Then
            Dim Valore As Integer = Val(NomeCampo)

            Campo = RigaAttuale.Item(Valore).ToString
        Else
            Campo = RigaAttuale.Item(NomeCampo).ToString
        End If

        Return Campo
    End Function

    Public Function RitornaNumeroDiColonne() As Integer
        Return Recordset.Columns.Count
    End Function

    Public Function RitornaNomiCampi() As String
        Dim Ritorno As String = ""

        For i As Integer = 0 To Recordset.Columns.Count - 1
            Ritorno += Recordset.Columns(i).ToString & ";"
        Next

        Return Ritorno
    End Function
End Class
