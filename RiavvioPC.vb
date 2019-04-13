Imports System.Runtime.InteropServices

Public Class RiavvioPC
    <StructLayout(LayoutKind.Sequential, Pack:=1)> _
    Friend Structure TokPriv1Luid
        Public Count As Integer
        Public Luid As Long
        Public Attr As Integer
    End Structure

    <DllImport("kernel32.dll", ExactSpelling:=True)> _
    Friend Shared Function GetCurrentProcess() As IntPtr
    End Function

    Friend Shared Function OpenProcessToken(h As IntPtr, acc As Integer, ByRef phtok As IntPtr) As Boolean
    End Function

    Friend Const SE_PRIVILEGE_ENABLED As Integer = &H2
    Friend Const TOKEN_QUERY As Integer = &H8
    Friend Const TOKEN_ADJUST_PRIVILEGES As Integer = &H20
    Friend Const SE_SHUTDOWN_NAME As String = "SeShutdownPrivilege"
    Friend Const EWX_LOGOFF As Integer = &H0
    Friend Const EWX_SHUTDOWN As Integer = &H1
    Friend Const EWX_REBOOT As Integer = &H2
    Friend Const EWX_FORCE As Integer = &H4
    Friend Const EWX_POWEROFF As Integer = &H8
    Friend Const EWX_FORCEIFHUNG As Integer = &H10

    <DllImport("advapi32.dll", SetLastError:=True)> _
    Friend Shared Function LookupPrivilegeValue(host As String, name As String, ByRef pluid As Long) As Boolean
    End Function

    Friend Shared Function AdjustTokenPrivileges(htok As IntPtr, disall As Boolean, ByRef newst As TokPriv1Luid, len As Integer, prev As IntPtr, relen As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")> _
    Private Shared Function ExitWindowsEx(uFlags As UInteger, dwReason As UInteger) As Boolean
    End Function

    <STAThread> _
    Public Function RiavviaPC() As String
        Dim Ritorno As String = ""

        Try
            Dim ok As Boolean
            Dim tp As TokPriv1Luid
            Dim hproc As IntPtr = GetCurrentProcess()
            Dim htok As IntPtr = IntPtr.Zero
            ok = OpenProcessToken(hproc, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, htok)
            tp.Count = 1
            tp.Luid = 0
            tp.Attr = SE_PRIVILEGE_ENABLED
            ok = LookupPrivilegeValue(Nothing, SE_SHUTDOWN_NAME, tp.Luid)
            If ok = False Then
                Ritorno = "ERRORE: LookupPrivilegeValue"
            Else
                ok = AdjustTokenPrivileges(htok, False, tp, 0, IntPtr.Zero, IntPtr.Zero)
                If ok = True Then
                    Ritorno = "ERRORE: AdjustTokenPrivileges"
                Else
                    ok = ExitWindowsEx(EWX_REBOOT Or EWX_FORCEIFHUNG, 0)
                    If ok = False Then
                        Ritorno = "ERRORE: ExitWindowsEx"
                    Else
                        Ritorno = "RIAVVIO EFFETTUATO"
                    End If
                End If
            End If
        Catch ex As Exception
            Ritorno = "ERRORE: " & ex.Message
        End Try

        Return Ritorno
    End Function
End Class
