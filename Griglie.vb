Imports System.Data
Imports System.Web.UI.WebControls

Public Class Griglie
    Private Colonne() As DataColumn
    Private riga As DataRow
    Private dttTabella As New DataTable()
    Private Db As New SQLSERVER
    Private ConnSQL As Object
    Private Rec As Object = CreateObject("ADODB.Recordset")
    Private QuantiCampi As Integer
    Private AggiunteRighe As Boolean = False

    Public Sub ImpostaCampi(Sql As String, Griglia As GridView, Optional NonEseguireSubito As Boolean = False)
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

        ApreDB()
        ImpostaValori(Sql)
        If NonEseguireSubito = False Then
            VisualizzaValori(Griglia)
            ChiudeDB()
        Else
            AggiunteRighe = True
        End If
    End Sub

    Private Function ConverteData(Campo As Date) As String
        Dim Ritorno As String = Format(Campo.Day, "00") & "/" & Format(Campo.Month, "00") & "/" & Campo.Year

        Return Ritorno
    End Function

    Private Sub ImpostaValori(Sql As String)
        Dim Campo As String

        Rec = Db.LeggeQuery(ConnSQL, Sql)

        Do Until Rec.Eof
            riga = dttTabella.NewRow()
            For i As Integer = 0 To QuantiCampi
                Campo = "" & Rec(i).Value

                If IsDate(Campo) = True And Campo.Length > 6 Then
                    Dim dCampo As Date = Campo

                    Campo = ConverteData(dCampo)
                End If

                riga(i) = Campo
            Next
            dttTabella.Rows.Add(riga)

            Rec.MoveNext()
        Loop

        Rec.Close()
    End Sub

    Public Sub ImpostaCampiDaQuerySQL(Sql As String, Griglia As GridView, Optional NonEseguireSubito As Boolean = False)
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

        ApreDB()
        ImpostaValoriDaQuerySQL(Sql)
        If NonEseguireSubito = False Then
            VisualizzaValori(Griglia)
            ChiudeDB()
        Else
            AggiunteRighe = True
        End If
    End Sub

    Public Sub PulisceGriglie(NomeGridView As GridView)
        Dim Dati() As String = {}

        ImpostaCampiDaCSV(Dati, NomeGridView)
    End Sub

    Public Sub ImpostaCampiDaCSV(CSV() As String, Griglia As GridView, Optional NonEseguireSubito As Boolean = False)
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

        ImpostaValoriCSV(CSV)
        If NonEseguireSubito = False Then
            VisualizzaValori(Griglia)
        Else
            AggiunteRighe = True
        End If
    End Sub

    Private Sub ApreDB()
        If Db.LeggeImpostazioniDiBase() = True Then
            ConnSQL = Db.ApreDB()
        End If
    End Sub

    Private Sub ChiudeDB()
        riga = Nothing
        dttTabella = Nothing

        ConnSQL.Close()
        Db = Nothing
    End Sub

    Public Sub AggiungeValori(Sql As String)
        Dim Campo As String

        Rec = Db.LeggeQuery(ConnSQL, Sql)

        Do Until Rec.Eof
            riga = dttTabella.NewRow()
            For i As Integer = 0 To QuantiCampi
                Campo = "" & Rec(i).Value

                If IsDate(Campo) = True And Campo.Length > 6 Then
                    Dim dCampo As Date = Campo

                    Campo = ConverteData(dCampo)
                End If

                riga(i) = Campo
            Next
            dttTabella.Rows.Add(riga)

            Rec.MoveNext()
        Loop

        Rec.Close()
    End Sub

    Private Sub ImpostaValoriDaQuerySQL(Sql As String)
        Dim Campo As String

        Rec = Db.LeggeQuery(ConnSQL, Sql)

        Do Until Rec.Eof
            riga = dttTabella.NewRow()
            For i As Integer = 0 To QuantiCampi
                Campo = "" & Rec(i).Value

                If IsDate(Campo) = True And Campo.Length > 6 Then
                    Dim dCampo As Date = Campo

                    Campo = ConverteData(dCampo)
                End If

                riga(i) = Campo
            Next
            dttTabella.Rows.Add(riga)

            Rec.MoveNext()
        Loop

        Rec.Close()
    End Sub

    Private Sub ImpostaValoriCSV(Csv() As String)
        Dim Campo() As String
        Dim sCampo As String

        For k As Integer = 1 To Csv.Length - 1
            Campo = Csv(k).Split(";")

            riga = dttTabella.NewRow()
            For i As Integer = 0 To QuantiCampi
                sCampo = "" & Campo(i)

                If IsDate(Campo) = True And Campo.Length > 6 Then
                    Dim dCampo As Date = sCampo

                    sCampo = ConverteData(dCampo)
                End If

                riga(i) = sCampo
            Next
            dttTabella.Rows.Add(riga)
        Next
    End Sub

    Public Sub VisualizzaValori(grdView As GridView)
        grdView.DataSource = dttTabella
        grdView.DataBind()

        If AggiunteRighe = True Then
            ChiudeDB()
        End If
    End Sub
End Class
