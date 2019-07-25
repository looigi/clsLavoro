Imports System.IO

Public Class GestioneCSV
	Public Function AddFieldToCSV(sValue As String, sFieldName As String, sOriginFile As String, sDestinFile As String,
									  iWhere As Integer, sSeparator As String, Optional bHeader As Boolean = True) As String
		Dim sFunctionReturn As String = ""

		Dim bSameFile As Boolean = False

		If sOriginFile = sDestinFile Then
			bSameFile = True
			sDestinFile = sOriginFile & ".tmp"
		Else
			If File.Exists(sDestinFile) Then
				File.Delete(sDestinFile)
			End If
		End If

		If File.Exists(sOriginFile) Then
			Dim outputFile As StreamWriter

			outputFile = New StreamWriter(sDestinFile, True)

			Try
				Dim objReader As StreamReader = New StreamReader(sOriginFile)
				Dim sLine As String = ""
				Dim sReturn As String = ""
				Dim bHeaderTemp As Boolean
				Dim bAdded As Boolean = False

				If bHeader Then
					bHeaderTemp = False
				Else
					bHeaderTemp = True
				End If

				Do
					sLine = objReader.ReadLine()
					If sLine <> "" Then
						If Not sLine.Contains(sSeparator) Then
							sReturn = "ERROR: Separator char not found"
							Exit Do
						End If

						If Not bHeaderTemp Then
							sReturn = AddFieldToLine(sLine, sSeparator, iWhere, sFieldName)
							bAdded = True
							bHeaderTemp = True
						Else
							sReturn = AddFieldToLine(sLine, sSeparator, iWhere, sValue)
							bAdded = True
						End If
						If sReturn.Contains("ERROR:") Then
							Exit Do
						Else
							outputFile.WriteLine(sReturn)
						End If

					End If
				Loop Until sLine Is Nothing
				objReader.Close()

				outputFile.Flush()
				outputFile.Close()

				If sReturn.Contains("ERROR:") Then
					sFunctionReturn = sReturn
				Else
					If Not bAdded Then
						sFunctionReturn = "ERROR: No rows updated"
					Else
						sFunctionReturn = "OK"

						If bSameFile Then
							File.Delete(sOriginFile)
							File.Copy(sDestinFile, sOriginFile)
							File.Delete(sDestinFile)
						End If
					End If
				End If
			Catch ex As Exception
				sFunctionReturn = "ERROR: " & ex.Message
			End Try
		Else
			sFunctionReturn = "ERROR: Origin file not exists"
		End If

		Return sFunctionReturn
	End Function

	Private Function AddFieldToLine(sRow As String, sSeparator As String, iPosition As Integer, sValue As String) As String
		Dim sReturn As String = ""
		Dim sField() As String = sRow.Split(sSeparator)
		Dim iCounter As Integer = 0
		Dim bAdded As Boolean = False
		Dim bLast As Boolean = False

		For Each s As String In sField
			iCounter += 1
			If iCounter = iPosition Then
				sReturn &= sValue & sSeparator & s & sSeparator
				bAdded = True
			Else
				sReturn &= s & sSeparator
			End If
		Next

		If Not False And iPosition = iCounter + 1 Then
			If sReturn <> "" Then
				Dim UltimoCarattere As String = Mid(sRow, sRow.Length, sRow.Length)
				If UltimoCarattere = sSeparator Then
					sReturn = Mid(sReturn, 1, sReturn.Length - 1)
				End If
			End If
			sReturn &= sValue
			bAdded = True
			bLast = True
		End If

		If bAdded = True Then
			Dim UltimoCarattere As String = Mid(sRow, sRow.Length, sRow.Length)
			If UltimoCarattere = sSeparator Then
				sReturn = Mid(sReturn, 1, sReturn.Length - 1)
			End If
		Else
			sReturn = "ERROR: position not valid"
		End If

		Return sReturn
	End Function
End Class
