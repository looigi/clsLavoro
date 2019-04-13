Imports System
Imports System.Collections.Specialized
Imports System.Configuration
Imports Oracle.DataAccess.Client
Imports System.Globalization

Public Class ttOledbConnection

    Public Enum tDateFormat
        ddmmyyyy
        ddmmyyyyhhmi
        ddmmyyyyhhmiss
    End Enum

    Private mvarActiveConnection As OracleConnection = Nothing
    Private Transaction As OracleTransaction
    Private mvarTransactionPending As Boolean = False

    Public ReadOnly Property TransactionPending() As Boolean
        Get
            TransactionPending = mvarTransactionPending
        End Get
    End Property

    Public ReadOnly Property State() As System.Data.ConnectionState
        Get
            Dim Result As ConnectionState
            Select Case IsNothing(mvarActiveConnection)
                Case True
                    Result = ConnectionState.Closed
                Case False
                    Result = mvarActiveConnection.State
            End Select
            Return Result
        End Get
    End Property

    Public Function GetOleDbCommand(ByVal Source As String) As OracleCommand
        Dim Command As OracleCommand
        Command = New OracleCommand(Source, mvarActiveConnection)
        Return Command
    End Function

    Public Function GetOleDbDataReader(ByVal Source As String) As OracleDataReader
        Dim myCommand As New OracleCommand(Source, mvarActiveConnection)
        GetOleDbDataReader = myCommand.ExecuteReader()
        myCommand = Nothing
    End Function

    Public Function GetOledbDataSet(ByVal Source As String) As System.Data.DataSet
        Dim Dts As New DataSet
        Dim Dta As New OracleDataAdapter(Source, mvarActiveConnection)
        Dta.Fill(Dts)
        Return Dts
    End Function

    Public Function GetDataTable(ByVal Source As String) As System.Data.DataTable
        Dim Dts As New DataSet
        Dim Dta As New OracleDataAdapter(Source, mvarActiveConnection)
        Dta.Fill(Dts)
        Return Dts.Tables(0)
    End Function

    Public Function TextToDate(ByVal dd_mm_yyyy As String) As Date
        Dim Data As Date
        Dim SplitChar As String = vbNullString
        Dim Buffer As String()
        Dim I As Integer

        Dim HH_MM_SS As String = vbNullString

        Buffer = Split(dd_mm_yyyy, Space(1))

        dd_mm_yyyy = Buffer(0)

        If Buffer.Length > 1 Then HH_MM_SS = Buffer(1)

        I = 1
        Do
            SplitChar = Mid(dd_mm_yyyy, I, 1)
            If Not IsNumeric(SplitChar) Then Exit Do
            I = I + 1
        Loop Until I > dd_mm_yyyy.Length

        Buffer = Split(dd_mm_yyyy, SplitChar)

        Data = DateSerial(CInt(Buffer(2)), _
                          CInt(Buffer(1)), _
                          CInt(Buffer(0)))

        If HH_MM_SS <> vbNullString Then Data = CDate(CStr(Data) & Space(1) & HH_MM_SS)

        TextToDate = Data

    End Function

    'Public Function TextToDate(ByVal dd_mm_yyyy As String) As Date
    '    Dim Data As Date
    '    Dim HH_MM_SS As String = vbNullString

    '    If Len(dd_mm_yyyy) > 10 Then
    '        HH_MM_SS = Trim(Right(dd_mm_yyyy, Len(dd_mm_yyyy) - 10))
    '        dd_mm_yyyy = Left(dd_mm_yyyy, 10)
    '    End If

    '    Data = DateSerial(CInt(Right(dd_mm_yyyy, 4)), _
    '                      CInt(Mid(dd_mm_yyyy, 4, 2)), _
    '                      CInt(Left(dd_mm_yyyy, 2)))

    '    If HH_MM_SS <> vbNullString Then Data = CDate(CStr(Data) & Space(1) & HH_MM_SS)
    '    TextToDate = Data
    'End Function

    Public Function TextIsValidDate(ByVal dd_mm_yyyy As String) As Boolean
        Dim Giorno As Integer
        Dim Mese As Integer
        Dim Anno As Integer
        Try
            Giorno = CType(Left(dd_mm_yyyy, 2), Integer)
            Mese = CType(Mid(dd_mm_yyyy, 4, 2), Integer)
            Anno = CType(Right(dd_mm_yyyy, 4), Integer)
            If Mese < 1 Or Mese > 12 Then
                Return False
            Else
                If Giorno < 1 Or Giorno > Day(DateSerial(Anno, Mese + 1, 0)) Then
                    Return False
                Else
                    Return True
                End If
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public ReadOnly Property GetDatabaseConnectionString(Optional ByVal ServerName As String = vbNullString, _
                                                         Optional ByVal SchemaName As String = vbNullString, _
                                                         Optional ByVal UserName As String = vbNullString, _
                                                         Optional ByVal Password As String = vbNullString) As String
        Get
            Dim strConnection As String = vbNullString
            If SchemaName = vbNullString Then SchemaName = UserName
            If UserName = vbNullString Then UserName = SchemaName
            strConnection = vbNullString
            strConnection = strConnection & "User ID=" & UserName & ";"
            strConnection = strConnection & "Password=" & Password & ";"
            strConnection = strConnection & "Data Source=" & ServerName & ";"
            Return strConnection
        End Get
    End Property

    Public Sub OpenConnection(ServerName As String, SchemaName As String, UserName As String, Password As String)
        Open(ServerName, _
             SchemaName, _
             UserName, _
             Password)
    End Sub

    Friend Sub Open(ByVal ServerName As String, _
                    ByVal SchemaName As String, _
                    ByVal UserName As String, _
                    ByVal Password As String)
        mvarTransactionPending = False
        Try
            mvarActiveConnection = New OracleConnection
            mvarActiveConnection.ConnectionString = GetDatabaseConnectionString(ServerName, _
                                                                                SchemaName, _
                                                                                UserName, _
                                                                                Password)

            mvarActiveConnection.Open()
        Catch ex As Exception
            Throw New System.Exception(ex.Message, ex.InnerException)
        End Try
    End Sub

    Public Sub CloseConnection()
        Try
            mvarTransactionPending = False
            If State <> ConnectionState.Closed Then mvarActiveConnection.Close()
            mvarActiveConnection = Nothing
        Catch ex As Exception
            Throw New System.Exception(ex.Message, ex.InnerException)
        End Try
    End Sub

    Public Function IsSelect(ByVal Query As String) As Boolean
        IsSelect = (InStr(UCase(Trim(Query)), "SELECT") = 1)
    End Function

    Public Sub Execute(ByVal TextCommand As String, Optional ByVal TimeOut As Object = Nothing)
        Dim Command As New OracleCommand(TextCommand, mvarActiveConnection)
        If Not IsNothing(TimeOut) Then Command.CommandTimeout = CType(TimeOut, Integer)
        Command.ExecuteNonQuery()
    End Sub

    Public Function ExecuteScalar(ByVal TextCommand As String) As Object
        Dim sqlCommand As New OracleCommand(TextCommand, mvarActiveConnection)
        ExecuteScalar = sqlCommand.ExecuteScalar
    End Function

    Public Sub BeginTransaction(Optional ByVal IsolationLevel As Data.IsolationLevel = Data.IsolationLevel.ReadCommitted)
        Transaction = Nothing
        Transaction = mvarActiveConnection.BeginTransaction(IsolationLevel)
        mvarTransactionPending = True
    End Sub

    Public Sub CommitTransaction()
        Transaction.Commit()
        Transaction = Nothing
        mvarTransactionPending = False
    End Sub

    Public Sub RollbackTransaction()
        Transaction.Rollback()
        Transaction = Nothing
        mvarTransactionPending = False
    End Sub

    Public Function MaxValue(ByVal Tabella As String, ByVal Key As String, Optional ByVal Where As String = Nothing) As Object
        Dim Dtb As DataTable
        Dim Source As String
        Source = "SELECT MAX(" & Key & ") FROM " & Tabella
        If Not Where Is Nothing Then Source = Source & " WHERE " & Where
        Try
            Dtb = Me.GetOledbDataSet(Source).Tables(0)
            MaxValue = Dtb.Rows(0).Item(0)
        Catch ex As Exception
            Throw New System.Exception(ex.Message, ex)
        Finally
            Dtb = Nothing
        End Try
    End Function

    Public Function NextValue(ByVal Tabella As String, ByVal Key As String, Optional ByVal Where As String = Nothing, Optional ByVal StartCounter As Integer = 0) As Long
        Dim Value As Object
        Value = MaxValue(Tabella, Key, Where)
        If Value Is DBNull.Value Then
            NextValue = StartCounter
        Else
            NextValue = CLng(Value) + 1
        End If
    End Function

    Public Function SYSDATE() As String
        Return "SELECT SYSDATE FROM DUAL"
    End Function

    Public Function GetSysDate() As Date
        Return Me.GetOleDbDataReader(SYSDATE()).GetSchemaTable.Rows(0).Item(0)
    End Function

    Public Function ConcatChar() As String
        Return "||"
    End Function

    Public Function To_Date(ByVal Source As Object, _
                            Optional ByVal DateFormat As tDateFormat = tDateFormat.ddmmyyyyhhmiss) As String
        Dim Dfi As New DateTimeFormatInfo
        Dim Data As Date
        Dim Result As String = vbNullString
        If Not IsDBNull(Source) Then
            Data = Source
            Select Case DateFormat
                Case tDateFormat.ddmmyyyy
                    Dfi.ShortDatePattern = "dd/MM/yyyy"
                    Result = "TO_DATE('" & Data.ToString("d", Dfi) & "','DD/MM/YYYY')"
                Case tDateFormat.ddmmyyyyhhmi
                    Dfi.ShortDatePattern = "dd/MM/yyyy H:mm"
                    Result = "TO_DATE('" & Data.ToString("d", Dfi) & "','DD/MM/YYYY HH24:MI')"
                Case tDateFormat.ddmmyyyyhhmiss
                    Dfi.ShortDatePattern = "dd/MM/yyyy H:mm:ss"
                    Result = "TO_DATE('" & Data.ToString("d", Dfi) & "','DD/MM/YYYY HH24:MI:SS')"
            End Select
        Else
            Result = "Null"
        End If
        Return Result
    End Function

    Public Function Upper(ByVal ColumName As String) As String
        Return "UPPER(" & ColumName & ")"
    End Function

    Public Function TextToDouble(ByVal Source As String) As Double
        Dim Format As New NumberFormatInfo
        Source = Replace(Source, Format.CurrencyDecimalSeparator, vbNullString)
        Source = Replace(Source, Format.CurrencyGroupSeparator, ".")
        Return CType(Source, Double)
    End Function

    Public Function TextToSingle(ByVal Source As String) As Double
        Dim Format As New NumberFormatInfo
        Source = Replace(Source, Format.CurrencyDecimalSeparator, vbNullString)
        Source = Replace(Source, Format.CurrencyGroupSeparator, ".")
        Return CType(Source, Single)
    End Function

    Public Function TextToInteger(ByVal Source As String) As Double
        Dim Format As New NumberFormatInfo
        Source = Replace(Source, Format.CurrencyDecimalSeparator, vbNullString)
        Return CType(Source, Integer)
    End Function

    Public Function TextToLong(ByVal Source As String) As Double
        Dim Format As New NumberFormatInfo
        Source = Replace(Source, Format.CurrencyDecimalSeparator, vbNullString)
        Return CType(Source, Long)
    End Function

    Public Function TextFormatted(ByVal Source As Object) As String
        Dim Format As New NumberFormatInfo
        Dim strAppo As String = vbNullString
        If Source Is DBNull.Value Then
            strAppo = "Null"
        ElseIf IsNothing(Source) Then
            strAppo = vbNullString
        Else
            Select Case UCase(TypeName(Source))
                Case "STRING", "CHAR"
                    Source = Source.ToString
                    strAppo = "'" & Replace(CStr(Source), "'", "''") & "'"
                Case "BOOLEAN"
                    strAppo = CStr(CInt(Source))
                Case "INTEGER", "LONG", "BYTE", "DECIMAL", "SHORT"
                    strAppo = CStr(Source)
                Case "SINGLE"
                    strAppo = Replace(CSng(Source).ToString, Format.CurrencyGroupSeparator, ".")
                Case "DOUBLE"
                    strAppo = Replace(CDbl(Source).ToString, Format.CurrencyGroupSeparator, ".")
                Case "DATE"
                    strAppo = To_Date(CDate(Source))
                Case Else
                    Throw New System.Exception("Errore di formattazione")
            End Select
        End If
        Return strAppo
    End Function

    Public Function GetDay(ByVal DateColumnName As String) As String
        Return "TO_NUMBER(TO_CHAR(" & DateColumnName & ",'DD'))"
    End Function

    Public Function GetMonth(ByVal DateColumnName As String) As String
        Return "TO_NUMBER(TO_CHAR(" & DateColumnName & ",'MM'))"
    End Function

    Public Function GetYear(ByVal DateColumnName As String) As String
        Return "TO_NUMBER(TO_CHAR(" & DateColumnName & ",'YYYY'))"
    End Function

End Class
