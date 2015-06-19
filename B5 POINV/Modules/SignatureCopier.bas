Attribute VB_Name = "SignatureCopier"
'---------------------------------------------------------------------------------------
' Module    : SignatureCopier
' DateTime  : 5/5/2006 09:47
' Author    : briandonorfio
' Purpose   : Copies signature files from EMP_DIR and SHARED_DIR to the user's local
'             drive.
'---------------------------------------------------------------------------------------

Option Explicit

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Private Const EMP_DIR = "\\toolsplus01\company\employee signatures\"
Private Const SHARED_DIR = "\\toolsplus01\company\outlook signatures\"
Private TO_DIR As String

Public Sub Main()

    ConsoleAlloc

    TO_DIR = Environ("AppData") & "\Microsoft\Signatures\"
    
    Dim user As String
    user = Environ("UserName")
    
    Dim file As String
    file = Dir$(EMP_DIR & user & ".*")
    While file <> ""
        Select Case Right(file, 4)
            Case Is = "html", ".htm", ".txt", ".rtf"
                If DateDiff("s", getModifiedDate(EMP_DIR & file), getModifiedDate(TO_DIR & file)) < 0 Then
                    ConsoleWriteLn "Copying " & file
                    Call Copy(EMP_DIR & file, TO_DIR & file)
                Else
                    ConsoleWriteLn "Skipping " & file
                End If
        End Select
        file = Dir$()
    Wend
    
    file = Dir$(SHARED_DIR & "*.*")
    While file <> ""
        Select Case Right(file, 4)
            Case Is = "html", ".htm", ".txt", ".rtf"
                If DateDiff("s", getModifiedDate(SHARED_DIR & file), getModifiedDate(TO_DIR & file)) < 0 Then
                    ConsoleWriteLn "Copying " & file
                    Call Copy(SHARED_DIR & file, TO_DIR & file)
                Else
                    ConsoleWriteLn "Skipping " & file
                End If
        End Select
        file = Dir$()
    Wend
    
    ConsoleFree
    
End Sub

Public Function getModifiedDate(filename As String) As String
    Dim fh As Long, ft As FILETIME, st As SYSTEMTIME
    fh = CreateFile(filename, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    If fh = -1 Then
        getModifiedDate = "1-1-1900 12:00:00"
    Else
        GetFileTime fh, ByVal 0&, ByVal 0&, ft
        FileTimeToLocalFileTime ft, ft
        FileTimeToSystemTime ft, st
        getModifiedDate = st.wMonth & "-" & st.wDay & "-" & st.wYear & " " & st.wHour & ":" & st.wMinute & ":" & st.wSecond
    End If
    CloseHandle fh
End Function

Private Function Copy(filefrom As String, fileto As String) As Boolean
    Copy = CopyFile(filefrom, fileto, 0)
End Function

