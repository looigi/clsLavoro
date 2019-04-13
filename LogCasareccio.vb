Public Class Logger
    Private idProc As Integer
    Private idMaschera As Integer
    Private FaiLog As Boolean = True
    Private NomeFileLog As String = ""
    Private gf As New GestioneFilesDirectory

    'Public Sub setIdProc(Valore As Integer, NumeroMaschera As Integer)
    '    idProc = Valore
    '    idMaschera = NumeroMaschera
    'End Sub

    'Private Sub frmLog_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    '    MascheraProcedura(idMaschera).ImpostaNomeProcedura(NomeProceduraScelta, idMaschera)
    '    MascheraProcedura(idMaschera).Show()
    '    Me.Hide()
    'End Sub

    Public Sub ImpostaFileDiLog(Nome As String)
        gf = New GestioneFilesDirectory
        NomeFileLog = Nome
    End Sub

    Public Function RitornaFileDiLog() As String
        Return NomeFileLog
    End Function

    Public Sub ScriveLogServizio(Cosa As String)
        If FaiLog Then
            Try
                Dim d As Date = Now
                Dim g As String = d.Day.ToString.Trim
                Dim m As String = d.Month.ToString.Trim
                Dim h As String = d.Hour.ToString.Trim
                Dim mm As String = d.Minute.ToString.Trim
                Dim s As String = d.Second.ToString.Trim
                If g.Length = 1 Then g = "0" & g
                If m.Length = 1 Then m = "0" & m
                If h.Length = 1 Then h = "0" & h
                If mm.Length = 1 Then mm = "0" & mm
                If s.Length = 1 Then s = "0" & s
                Dim df As String = d.Year & "-" & m & "-" & g & " " & h & ":" & mm & ":" & s

                gf.CreaDirectoryDaPercorso(gf.TornaNomeDirectoryDaPath(NomeFileLog) & "\")

                gf.ApreFileDiTestoPerScrittura(NomeFileLog)
                gf.ScriveTestoSuFileAperto(df & ";" & Cosa)
                gf.ChiudeFileDiTestoDopoScrittura()
            Catch ex As Exception
                Stop
            End Try
        End If
    End Sub

    'Private Sub frmLog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '    CaricaLog()
    'End Sub

    Private Sub CaricaLog()
        'Dim DB As New OperazioniSuFile.GestioneACCESS

        'If DB.LeggeImpostazioniDiBase("ConnDB") = True Then
        '    Dim ConnSQL As Object = DB.ApreDB(idProc)
        '    Dim Rec As Object = CreateObject("ADODB.Recordset")
        '    Dim Sql As String

        '    lstAppoggio.Items.Clear()

        '    lblNomeProc.Text = ""
        '    Sql = "Select * From NomiProcedure Where idProc=" & idProc
        '    Rec = DB.LeggeQuery(idProc, ConnSQL, Sql)
        '    If Rec.Eof = False Then
        '        lblNomeProc.Text = "Procedura " & Rec("NomeProcedura").Value
        '    End If
        '    Rec.Close()

        '    Sql = "Select Top 500 * From LogOperazioni Where idProc=" & idProc & " Order By Progressivo Desc"
        '    Rec = DB.LeggeQuery(idProc, ConnSQL, Sql)
        '    Do Until Rec.Eof
        '        lstAppoggio.Items.Add(Rec("DataOra").Value & ":" & Rec("Operazione").Value)

        '        Rec.MoveNext()
        '    Loop
        '    Rec.Close()

        '    lstLog.Items.Clear()
        '    For i As Integer = lstAppoggio.Items.Count - 1 To 0 Step -1
        '        lstLog.Items.Add(lstAppoggio.Items(i))
        '    Next
        '    lstLog.SelectedIndex = lstLog.Items.Count - 1

        '    ConnSQL.close()
        '    ConnSQL = Nothing

        '    DB.ChiudeDB(True, ConnSQL)
        'End If

        'DB = Nothing
    End Sub

    Public Function RitornaNomeFileLog() As String
        Return NomeFileLog
    End Function

    'Private Sub cmdRefresh_Click(sender As Object, e As EventArgs) Handles cmdRefresh.Click
    '    CaricaLog()
    'End Sub

    'Private Sub cmdClearLOG_Click(sender As Object, e As EventArgs) Handles cmdClearLOG.Click
    'Dim DB As New OperazioniSuFile.GestioneACCESS

    'If DB.LeggeImpostazioniDiBase("ConnDB") = True Then
    '    Dim ConnSQL As Object = DB.ApreDB(idProc)
    '    Dim Sql As String

    '    Sql = "Delete From LogOperazioni Where idProc=" & idProc
    '    DB.EsegueSql(idProc, ConnSQL, Sql)

    '    ConnSQL.close()
    '    ConnSQL = Nothing

    '    DB.CompattazioneDb()

    '    DB.ChiudeDB(True, ConnSQL)
    'End If

    'DB = Nothing

    'CaricaLog()
    'End Sub

    'Private Sub lstLog_DoubleClick(sender As Object, e As EventArgs) Handles lstLog.DoubleClick
    '    MsgBox(lstLog.Text, vbInformation)
    'End Sub
End Class
