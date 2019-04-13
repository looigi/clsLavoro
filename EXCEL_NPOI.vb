Imports System.Data
Imports System.IO
Imports System.Data.OleDb
Imports NPOI.SS.UserModel
Imports NPOI.HSSF.UserModel
Imports Excel = Microsoft.Office.Interop.Excel

Public Class EXCEL_NPOI
    Private xlApp As Excel.Application
    Private xlWorkBook As Excel.Workbook
    Private xlWorkSheet As Excel.Worksheet

    Public Sub EseguCalcoloFormule()
        xlWorkSheet.Calculate()
    End Sub

    Public Function ApreDocumentoEXCEL(sourceFile As String) As Boolean
        Dim Ok As Boolean = True
        Dim xlAppI As Excel.Application = New Excel.Application
        Dim xlWorkBookI As Excel.Workbook

        xlWorkBookI = xlAppI.Workbooks.Open(sourceFile)

        xlApp = xlAppI
        xlWorkBook = xlWorkBookI

        Return Ok
    End Function

    Public Function CreaDocumentoEXCEL() As Boolean
        Dim Ok As Boolean = True
        Dim xlAppI As Excel.Application = New Excel.Application
        Dim xlWorkBookI As Excel.Workbook

        xlWorkBookI = xlAppI.Workbooks.Add

        xlApp = xlAppI
        xlWorkBook = xlWorkBookI

        CreaNuovoFoglioDiLavoro("Appoggio")

        Return Ok
    End Function

    'Public Function EsegueFormulaRata(sFormula As String) As String
    '    Return xlApp.WorksheetFunction.Rate(8.9 / 12, 12, -12 * 2500, , 1)
    '    ' xlWorkSheet.Range("A1:A2").FormulaR1C1 = sFormula

    '    ' Return xlWorkSheet.Cells(1, 1).Text
    'End Function

    Public Sub ScriveValoreSuCella(Y As Integer, X As Integer, Cosa As String)
        xlWorkSheet.Cells(Y, X).Value = Cosa
    End Sub

    Public Function LeggeValoreDaCella(Y As Integer, X As Integer) As String
        Return xlWorkSheet.Cells(Y, X).Text
    End Function

    Public Sub SaveFoglioExcel(Nome As String)
        xlWorkBook.SaveAs(Nome, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, System.Reflection.Missing.Value, System.Reflection.Missing.Value, False, False, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, True, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value)
    End Sub

    Public Sub ImpostaFoglioDiLavoro(NomeFoglio As String)
        Dim xlWorkSheetI As Excel.Worksheet
        xlWorkSheetI = xlWorkBook.Worksheets(NomeFoglio)

        xlWorkSheet = xlWorkSheetI
    End Sub

    Public Sub CreaNuovoFoglioDiLavoro(NomeFoglio As String)
        Dim xlWorkSheetI As Excel.Worksheet
        xlWorkSheetI = xlWorkBook.Sheets.Add
        xlWorkSheetI.Name = NomeFoglio

        xlWorkSheet = xlWorkSheetI
    End Sub

    Public Sub ChiudeFoglioEXCEL()
        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Public Sub convertExcelToCSV(ByVal sourceFile As String, ByVal targetFile As String, NumeroFoglio As Integer, Optional NumeroCella() As String = Nothing, Optional Path As String = "")
        Dim gf As New GestioneFilesDirectory

        Dim FoglioDiLavoro As IWorkbook = Nothing
        Dim Foglio As ISheet

        Using file As New FileStream(sourceFile, FileMode.Open, FileAccess.Read)
            If sourceFile.ToLower.EndsWith(".xlsx") Then
                Stop
            Else
                If sourceFile.ToLower.EndsWith(".xls") Then
                    FoglioDiLavoro = New HSSFWorkbook(file)
                End If
            End If
        End Using

        Try
            Kill(targetFile)
        Catch ex As Exception

        End Try

        Using scrittura As StreamWriter = File.AppendText(targetFile)
            Foglio = FoglioDiLavoro.GetSheetAt(NumeroFoglio)

            ' Dim Righe As System.Collections.IEnumerator = Foglio.GetRowEnumerator()
            Dim Riga As IRow
            Dim RigaDaScrivere As String
            'Dim ContenutoFile As String = ""
            Dim Ok As Boolean
            Dim Campo As String
            Dim NumeroCelle As Integer
            Dim rowCount As Integer
            Dim Colonna As Integer

            'gf.CreaAggiornaFile(Path & "\Appoggio\buttami.txt", "Controllo celle")
            If NumeroCella Is Nothing = True Then
                Dim RigaDaControllare As Integer = 0

                Riga = Foglio.GetRow(RigaDaControllare)
                Do While Riga Is Nothing And RigaDaControllare <= Foglio.LastRowNum
                    RigaDaControllare += 1
                    Riga = Foglio.GetRow(RigaDaControllare)
                Loop
                If Riga Is Nothing = False Then
                    NumeroCelle = Riga.Cells.Count - 1
                    ReDim NumeroCella(NumeroCelle)
                    For i As Integer = 0 To NumeroCelle
                        NumeroCella(i) = i
                    Next
                End If
            Else
                NumeroCelle = NumeroCella.Length - 1
            End If
            For rowCount = 0 To Foglio.LastRowNum
                'gf.CreaAggiornaFile(Path & "\Appoggio\buttami.txt", "Conversione: " & rowCount & "/" & Foglio.LastRowNum)

                Riga = Foglio.GetRow(rowCount)
                If Riga Is Nothing = False Then
                    RigaDaScrivere = ""
                    Ok = False
                    For Colonna = 0 To NumeroCelle
                        'If Colonna = 19 Then Stop
                        Campo = ""
                        Try
                            Select Case Riga.GetCell(NumeroCella(Colonna)).CellType
                                Case CellType.Numeric
                                    If (DateUtil.IsCellDateFormatted(Riga(NumeroCella(Colonna)))) Then
                                        Campo = Riga.GetCell(NumeroCella(Colonna)).DateCellValue.ToString()
                                    Else
                                        Campo = Riga.GetCell(NumeroCella(Colonna)).NumericCellValue.ToString()
                                    End If
                                Case CellType.String
                                    Campo = Riga.GetCell(NumeroCella(Colonna)).StringCellValue.Trim
                                Case CellType.Blank
                                    Campo = ""
                                Case Else
                                    Campo = ""
                            End Select
                        Catch ex As Exception
                            Campo = ""
                        End Try

                        If Campo <> "" And Ok = False Then
                            Ok = True
                        End If

                        RigaDaScrivere += Chr(34) & Campo & Chr(34) & "§"

                        'scrittura.Write(Colonna & ":" & Campo & Chr(13) & Chr(10))
                    Next
                    If Ok = True Then
                        'ContenutoFile += RigaDaScrivere & Chr(207)

                        scrittura.Write(RigaDaScrivere & Chr(207))
                        'scrittura.Write(Chr(13) & Chr(10))
                    End If
                End If
                'Catch ex As Exception
                '    Stop
                'End Try
            Next

            Foglio = Nothing
            FoglioDiLavoro.Close()
            FoglioDiLavoro = Nothing
            scrittura.Close()
        End Using

        'If ContenutoFile <> "" Then
        '    gf.CreaAggiornaFile(targetFile, ContenutoFile)
        'End If

        gf = Nothing
    End Sub

    Public Sub ConverteCSVInXLS(NomeFile As String, NomeScheletro As String, RigaInizio As Integer, sPath As String, NomeDestinazioneExcel As String, Optional Intestazione As String = "")
        Dim rdf As System.IO.StreamReader = New System.IO.StreamReader(NomeFile, System.Text.Encoding.Default)
        Dim tmp As String = rdf.ReadLine()
        Dim Campi() As String
        Dim Riga As Long = RigaInizio - 1
        Dim Colonna As Integer = 0

        Try
            Kill(NomeDestinazioneExcel)
        Catch ex As Exception

        End Try

        Dim fs As FileStream = New FileStream(sPath & "\Scheletri\" & NomeScheletro, FileMode.Open, FileAccess.Read)
        Dim templateWorkbook As HSSFWorkbook = New HSSFWorkbook(fs, True)
        Dim sheet As HSSFSheet = templateWorkbook.GetSheetAt(0)
        Dim dataCell As HSSFCell
        Dim stileNero As HSSFCellStyle = templateWorkbook.CreateCellStyle()
        Dim stileBlu As HSSFCellStyle = templateWorkbook.CreateCellStyle()
        Dim stileRosso As HSSFCellStyle = templateWorkbook.CreateCellStyle()
        Dim stileVerde As HSSFCellStyle = templateWorkbook.CreateCellStyle()
        Dim fontNero As HSSFFont = templateWorkbook.CreateFont()
        Dim fontBlu As HSSFFont = templateWorkbook.CreateFont()
        Dim fontRosso As HSSFFont = templateWorkbook.CreateFont()
        Dim fontVerde As HSSFFont = templateWorkbook.CreateFont()
        Dim Row As IRow

        fontNero.Color = NPOI.HSSF.Util.HSSFColor.Black.Index
        stileNero.SetFont(fontNero)

        fontBlu.Color = NPOI.HSSF.Util.HSSFColor.Red.Index
        stileRosso.SetFont(fontBlu)

        fontRosso.Color = NPOI.HSSF.Util.HSSFColor.Blue.Index
        stileBlu.SetFont(fontRosso)

        fontVerde.Color = NPOI.HSSF.Util.HSSFColor.Green.Index
        stileVerde.SetFont(fontVerde)

        Dim cellaColore As String

        If intestazione <> "" Then
            Row = sheet.GetRow(0)
            Row.Cells(0).SetCellValue(intestazione)
        End If

        While Not tmp Is Nothing
            Campi = tmp.Split(";")

            cellaColore = Campi(Campi.Length - 2).Replace(Chr(34), "").Trim.ToUpper

            Row = sheet.CreateRow(Riga)

            For i As Integer = 0 To Campi.Length - 3
                dataCell = Row.CreateCell(i)
                dataCell.SetCellValue(Campi(i).Replace(Chr(34), ""))
                Select Case cellaColore
                    Case "ROSSO"
                        dataCell.CellStyle = stileRosso
                    Case "TURCHESE"
                        dataCell.CellStyle = stileBlu
                    Case "VERDE", "VERDE-NOTE"
                        dataCell.CellStyle = stileVerde
                    Case Else
                        dataCell.CellStyle = stileNero
                End Select
            Next

            Riga += 1

            tmp = rdf.ReadLine()
        End While

        rdf.Close()
        rdf = Nothing

        Using fileData = New FileStream(NomeDestinazioneExcel, FileMode.Create)
            templateWorkbook.Write(fileData)
            fileData.Close()
        End Using

        Row = Nothing
        sheet = Nothing

        templateWorkbook.Close()
        fs.Close()

        fs = Nothing
        templateWorkbook = Nothing

        'Kill(NomeFile)
    End Sub

End Class
