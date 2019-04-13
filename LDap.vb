Imports System.Text
Imports System.Collections
Imports System.DirectoryServices

Public Class LdapAuthentication
    Private _path As String
    Private _filterAttribute As String

    Public Sub New(ByVal path As String)
        _path = path
    End Sub

    Public Function IsAuthenticated(ByVal domain As String, ByVal username As String, ByVal pwd As String) As String
        Dim domainAndUsername As String = Convert.ToString(domain & Convert.ToString("\")) & username
        Dim entry As New DirectoryEntry(_path, domainAndUsername, pwd)
        Dim NomeUtente2 As String = ""
        Dim Mail As String = ""
        Dim DepartmentUtente2 As String = ""

        Try
            'Bind to the native AdsObject to force authentication.
            'object obj = entry.NativeObject;

            Dim search As New DirectorySearcher(entry)

            search.Filter = (Convert.ToString("(SAMAccountName=") & username) + ")"
            search.PropertiesToLoad.Add("cn")
            search.PropertiesToLoad.Add("department")
            search.PropertiesToLoad.Add("displayName")
            search.PropertiesToLoad.Add("mail")
            Dim result As SearchResult = search.FindOne()

            NomeUtente2 = result.Properties("displayName")(0).ToString()
            Mail = result.Properties("mail")(0).ToString()
            DepartmentUtente2 = result.Properties("department")(0).ToString()

            If result Is Nothing Then
                Return False
            End If

            'Update the new path to the user in the directory.
            _path = result.Path
            _filterAttribute = DirectCast(result.Properties("cn")(0), String)
        Catch dscex As DirectoryServicesCOMException
            NomeUtente2 = ""
            DepartmentUtente2 = ""
        Catch ex As Exception
            NomeUtente2 = ""
            DepartmentUtente2 = ""
            ' Program.log.[Error]("LdapAuthentication.IsAuthenticated.")
            Throw ex
        End Try

        Return NomeUtente2 & "*" & DepartmentUtente2 & "*" & Mail
    End Function

    Public Function GetGroups() As String
        Dim search As New DirectorySearcher(_path)
        search.Filter = (Convert.ToString("(cn=") & _filterAttribute) + ")"
        search.PropertiesToLoad.Add("memberOf")
        Dim groupNames As New StringBuilder()

        Try
            Dim result As SearchResult = search.FindOne()
            Dim propertyCount As Integer = result.Properties("memberOf").Count
            Dim dn As String
            Dim equalsIndex As Integer, commaIndex As Integer

            For propertyCounter As Integer = 0 To propertyCount - 1
                dn = DirectCast(result.Properties("memberOf")(propertyCounter), String)
                equalsIndex = dn.IndexOf("=", 1)
                commaIndex = dn.IndexOf(",", 1)
                If -1 = equalsIndex Then
                    Return Nothing
                End If
                groupNames.Append(dn.Substring((equalsIndex + 1), (commaIndex - equalsIndex) - 1))
                groupNames.Append("|")
            Next
        Catch ex As Exception
            Throw New Exception("Error obtaining group names. " + ex.Message)
        End Try
        Return groupNames.ToString()
    End Function
End Class


