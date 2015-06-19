Attribute VB_Name = "SyncShell"
'---------------------------------------------------------------------------------------
' Module    : SyncShell
' DateTime  : 6/14/2005 10:36
' Author    : briandonorfio
' Purpose   : Provides a shell function that wait until process finishes.
'
'             Dependencies:
'               - none
'---------------------------------------------------------------------------------------

Private Const STARTF_USESHOWWINDOW& = &H1
Private Const STARTF_USESTDHANDLES = &H100
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const SW_HIDE = 0
Private Const INFINITE = -1&

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long

'---------------------------------------------------------------------------------------
' Procedure : ShellWait
' DateTime  : 6/14/2005 11:10
' Author    : briandonorfio
' Purpose   : Starts a synchronous shell. Usage similar to the normal shell command.
'---------------------------------------------------------------------------------------
'
Public Function ShellWait(PathName As String, Optional WindowStyle As VbAppWinStyle = vbNormalFocus) As Boolean
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim Ret As Long
    ' Initialize the STARTUPINFO structure:
    With start
        .cb = Len(start)
        'If Not IsMissing(WindowStyle) Then
            .dwFlags = STARTF_USESHOWWINDOW
            .wShowWindow = WindowStyle
        'End If
    End With
    ' Start the shelled application:
    Ret& = CreateProcessA(0&, PathName, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    ' Wait for the shelled application to finish:
    Ret& = WaitForSingleObject(proc.hProcess, INFINITE)
    Ret& = CloseHandle(proc.hProcess)
    ShellWait = True
End Function

'---------------------------------------------------------------------------------------
' Procedure : ShellCapture
' Author    : briandonorfio
' Date      : 1/19/2012
' Purpose   : Starts a synchronous shell, and returns STDOUT/STDERR
'---------------------------------------------------------------------------------------
'
Public Function ShellCapture(cmd As String) As String
    Dim piperead As Long, pipewrite As Long
    Dim sa As SECURITY_ATTRIBUTES
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1
    If 0 = CreatePipe(piperead, pipewrite, sa, 0) Then
        ShellCapture = ""
        Exit Function
    End If
    
    Dim si As STARTUPINFO
    si.cb = Len(si)
    si.dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
    si.wShowWindow = SW_HIDE
    si.hStdOutput = pipewrite
    si.hStdError = pipewrite
    
    Dim pi As PROCESS_INFORMATION
    
    If CreateProcessA(0&, cmd, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, si, pi) Then
        CloseHandle pipewrite
        CloseHandle pi.hThread
        pipewrite = 0
        
        Const bufsize As Long = 10240
        Dim out(10240) As Byte
        Dim bytesRead As Long
        Dim retval As String
        retval = ""
        Do
            If ReadFile(piperead, out(0), bufsize, bytesRead, ByVal 0&) = 0 Then
                Exit Do
            End If
            retval = retval & Left$(StrConv(out(), vbUnicode), bytesRead)
        Loop
        CloseHandle pi.hProcess
        
        ShellCapture = retval
    Else
        ShellCapture = ""
    End If
    
    CloseHandle piperead
    CloseHandle pipewrite
End Function
