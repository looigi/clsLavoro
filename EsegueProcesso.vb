Imports System.Runtime.InteropServices
Imports System.Diagnostics

Class EsegueProcesso

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure PROCESS_INFORMATION
        Public hProcess As IntPtr
        Public hThread As IntPtr
        Public dwProcessId As System.UInt32
        Public dwThreadId As System.UInt32
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure SECURITY_ATTRIBUTES
        Public nLength As System.UInt32
        Public lpSecurityDescriptor As IntPtr
        Public bInheritHandle As Boolean
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure STARTUPINFO
        Public cb As System.UInt32
        Public lpReserved As String
        Public lpDesktop As String
        Public lpTitle As String
        Public dwX As System.UInt32
        Public dwY As System.UInt32
        Public dwXSize As System.UInt32
        Public dwYSize As System.UInt32
        Public dwXCountChars As System.UInt32
        Public dwYCountChars As System.UInt32
        Public dwFillAttribute As System.UInt32
        Public dwFlags As System.UInt32
        Public wShowWindow As Short
        Public cbReserved2 As Short
        Public lpReserved2 As IntPtr
        Public hStdInput As IntPtr
        Public hStdOutput As IntPtr
        Public hStdError As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure PROFILEINFO
        Public dwSize As Integer
        Public dwFlags As Integer
        <MarshalAs(UnmanagedType.LPTStr)> _
        Public lpUserName As String
        <MarshalAs(UnmanagedType.LPTStr)> _
        Public lpProfilePath As String
        <MarshalAs(UnmanagedType.LPTStr)> _
        Public lpDefaultPath As String
        <MarshalAs(UnmanagedType.LPTStr)> _
        Public lpServerName As String
        <MarshalAs(UnmanagedType.LPTStr)> _
        Public lpPolicyPath As String
        Public hProfile As IntPtr
    End Structure

    Friend Enum SECURITY_IMPERSONATION_LEVEL
        SecurityAnonymous = 0
        SecurityIdentification = 1
        SecurityImpersonation = 2
        SecurityDelegation = 3
    End Enum

    Friend Enum TOKEN_TYPE
        TokenPrimary = 1
        TokenImpersonation = 2
    End Enum

    <DllImport("advapi32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Private Shared Function CreateProcessAsUser(hToken As IntPtr, lpApplicationName As String, lpCommandLine As String, ByRef lpProcessAttributes As SECURITY_ATTRIBUTES, ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, bInheritHandles As Boolean, _
            dwCreationFlags As UInteger, lpEnvironment As IntPtr, lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Boolean
    End Function

    <DllImport("advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function DuplicateTokenEx(hExistingToken As IntPtr, dwDesiredAccess As UInteger, ByRef lpTokenAttributes As SECURITY_ATTRIBUTES, ImpersonationLevel As SECURITY_IMPERSONATION_LEVEL, TokenType As TOKEN_TYPE, ByRef phNewToken As IntPtr) As Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)> _
    Private Shared Function OpenProcessToken(ProcessHandle As IntPtr, DesiredAccess As Integer, ByRef TokenHandle As IntPtr) As Boolean
    End Function

    <DllImport("userenv.dll", SetLastError:=True)> _
    Private Shared Function CreateEnvironmentBlock(ByRef lpEnvironment As IntPtr, hToken As IntPtr, bInherit As Boolean) As Boolean
    End Function

    <DllImport("userenv.dll", SetLastError:=True)> _
    Private Shared Function DestroyEnvironmentBlock(lpEnvironment As IntPtr) As Boolean
    End Function

    Private Const SW_SHOW As Short = 1
    Private Const SW_SHOWMAXIMIZED As Short = 7
    Private Const TOKEN_QUERY As Integer = 8
    Private Const TOKEN_DUPLICATE As Integer = 2
    Private Const TOKEN_ASSIGN_PRIMARY As Integer = 1
    Private Const GENERIC_ALL_ACCESS As Integer = 268435456
    Private Const STARTF_USESHOWWINDOW As Integer = 1
    Private Const STARTF_FORCEONFEEDBACK As Integer = 64
    Private Const CREATE_UNICODE_ENVIRONMENT As Integer = &H400
    Private Const gs_EXPLORER As String = "explorer"

    Public Shared Sub LaunchProcess(Ps_CmdLine As String)
        Dim li_Token As IntPtr = Nothing
        Dim li_EnvBlock As IntPtr = Nothing
        Dim lObj_Processes As Process() = Process.GetProcessesByName(gs_EXPLORER)

        ' Get explorer.exe id

        ' If process not found
        If lObj_Processes.Length = 0 Then
            ' Exit routine
            Return
        End If

        ' Get primary token for the user currently logged in
        li_Token = GetPrimaryToken(lObj_Processes(0).Id)

        ' If token is nothing
        If li_Token.Equals(IntPtr.Zero) Then
            ' Exit routine
            Return
        End If

        ' Get environment block
        li_EnvBlock = GetEnvironmentBlock(li_Token)

        ' Launch the process using the environment block and primary token
        LaunchProcessAsUser(Ps_CmdLine, li_Token, li_EnvBlock)

        ' If no valid enviroment block found
        If li_EnvBlock.Equals(IntPtr.Zero) Then
            ' Exit routine
            Return
        End If

        ' Destroy environment block. Free environment variables created by the 
        ' CreateEnvironmentBlock function.
        DestroyEnvironmentBlock(li_EnvBlock)
    End Sub

    Private Shared Function GetPrimaryToken(Pi_ProcessId As Integer) As IntPtr
        Dim li_Token As IntPtr = IntPtr.Zero
        Dim li_PrimaryToken As IntPtr = IntPtr.Zero
        Dim lb_ReturnValue As Boolean = False
        Dim lObj_Process As Process = Process.GetProcessById(Pi_ProcessId)
        Dim lObj_SecurityAttributes As SECURITY_ATTRIBUTES = Nothing

        ' Get process by id
        ' Open a handle to the access token associated with a process. The access token 
        ' is a runtime object that represents a user account.
        lb_ReturnValue = OpenProcessToken(lObj_Process.Handle, TOKEN_DUPLICATE, li_Token)

        ' If successfull in opening handle to token associated with process
        If lb_ReturnValue Then

            ' Create security attributes to pass to the DuplicateTokenEx function
            lObj_SecurityAttributes = New SECURITY_ATTRIBUTES()
            lObj_SecurityAttributes.nLength = Convert.ToUInt32(Marshal.SizeOf(lObj_SecurityAttributes))

            ' Create a new access token that duplicates an existing token. This function 
            ' can create either a primary token or an impersonation token.
            lb_ReturnValue = DuplicateTokenEx(li_Token, Convert.ToUInt32(TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_QUERY), lObj_SecurityAttributes, SECURITY_IMPERSONATION_LEVEL.SecurityIdentification, TOKEN_TYPE.TokenPrimary, li_PrimaryToken)

            ' If un-successful in duplication of the token
            If Not lb_ReturnValue Then
                ' Throw exception
                Throw New Exception(String.Format("DuplicateTokenEx Error: {0}", Marshal.GetLastWin32Error()))
            End If
        Else
            ' If un-successful in opening handle for token then throw exception
            Throw New Exception(String.Format("OpenProcessToken Error: {0}", Marshal.GetLastWin32Error()))
        End If

        ' Return primary token
        Return li_PrimaryToken
    End Function

    Private Shared Function GetEnvironmentBlock(Pi_Token As IntPtr) As IntPtr
        Dim li_EnvBlock As IntPtr = IntPtr.Zero
        Dim lb_ReturnValue As Boolean = CreateEnvironmentBlock(li_EnvBlock, Pi_Token, False)

        ' Retrieve the environment variables for the specified user. 
        ' This block can then be passed to the CreateProcessAsUser function.

        ' If not successful in creation of environment block then  
        If Not lb_ReturnValue Then
            ' Throw exception
            Throw New Exception(String.Format("CreateEnvironmentBlock Error: {0}", Marshal.GetLastWin32Error()))
        End If

        ' Return the retrieved environment block
        Return li_EnvBlock
    End Function

    Private Shared Sub LaunchProcessAsUser(Ps_CmdLine As String, Pi_Token As IntPtr, Pi_EnvBlock As IntPtr)
        Dim lb_Result As Boolean = False
        Dim lObj_ProcessInformation As PROCESS_INFORMATION = Nothing
        Dim lObj_ProcessAttributes As SECURITY_ATTRIBUTES = Nothing
        Dim lObj_ThreadAttributes As SECURITY_ATTRIBUTES = Nothing
        Dim lObj_StartupInfo As STARTUPINFO = Nothing

        ' Information about the newly created process and its primary thread.
        lObj_ProcessInformation = New PROCESS_INFORMATION()

        ' Create security attributes to pass to the CreateProcessAsUser function
        lObj_ProcessAttributes = New SECURITY_ATTRIBUTES()
        lObj_ProcessAttributes.nLength = Convert.ToUInt32(Marshal.SizeOf(lObj_ProcessAttributes))
        lObj_ThreadAttributes = New SECURITY_ATTRIBUTES()
        lObj_ThreadAttributes.nLength = Convert.ToUInt32(Marshal.SizeOf(lObj_ThreadAttributes))

        ' To specify the window station, desktop, standard handles, and appearance of the 
        ' main window for the new process.
        lObj_StartupInfo = New STARTUPINFO()
        lObj_StartupInfo.cb = Convert.ToUInt32(Marshal.SizeOf(lObj_StartupInfo))
        lObj_StartupInfo.lpDesktop = Nothing
        lObj_StartupInfo.dwFlags = Convert.ToUInt32(STARTF_USESHOWWINDOW Or STARTF_FORCEONFEEDBACK)
        lObj_StartupInfo.wShowWindow = SW_SHOW

        ' Creates a new process and its primary thread. The new process runs in the 
        ' security context of the user represented by the specified token.
        lb_Result = CreateProcessAsUser(Pi_Token, Nothing, Ps_CmdLine, lObj_ProcessAttributes, lObj_ThreadAttributes, True, _
            CREATE_UNICODE_ENVIRONMENT, Pi_EnvBlock, Nothing, lObj_StartupInfo, lObj_ProcessInformation)

        ' If create process return false then
        If Not lb_Result Then
            ' Throw exception
            Throw New Exception(String.Format("CreateProcessAsUser Error: {0}", Marshal.GetLastWin32Error()))
        End If
    End Sub
End Class
