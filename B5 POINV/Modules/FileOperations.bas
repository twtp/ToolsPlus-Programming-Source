Attribute VB_Name = "FileOperations"
'---------------------------------------------------------------------------------------
' Module    : FileOperations
' DateTime  : 6/9/2006 14:09
' Author    : briandonorfio
' Purpose   : Filesystem related functions, most of these are duplicates of what you can
'             get with FSO, just 2x or so faster.
'
'             Dependencies:
'               - none
'---------------------------------------------------------------------------------------

Option Explicit

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 255
    cAlternate As String * 14
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

Public Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type

Private Const ERROR_NO_MORE_FILES = 18&
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As FIND_DATA) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

'---------------------------------------------------------------------------------------
' Procedure : FileExists
' DateTime  : 6/9/2006 14:10
' Author    : briandonorfio
' Purpose   : Basic check for a file
'
'             This is horribly unstable when running in the IDE under win7 (random out
'             of memory errors or IDE crashes), so going the FSO route.
'---------------------------------------------------------------------------------------
'
Public Function FileExists(file As String) As Boolean
    'Dim fh As Long, filedata As FIND_DATA
    'fh = FindFirstFile(file, filedata)
    'If fh = -1 Then
    '    FileExists = False
    'Else
    '    FileExists = True
    'End If
    'Call FindClose(fh)
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    FileExists = fso.FileExists(file)
    Set fso = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : FolderExists
' DateTime  : 6/9/2006 14:10
' Author    : briandonorfio
' Purpose   : Basic check for a folder
'---------------------------------------------------------------------------------------
'
Public Function FolderExists(folder As String) As Boolean
    Dim fh As Long, filedata As FIND_DATA
    fh = FindFirstFile(folder & IIf(Right(folder, 1) = "\", "", "\") & "*.*", filedata)
    If fh = -1 Then
        FolderExists = False
    Else
        FolderExists = True
    End If
    Call FindClose(fh)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFolderContents
' DateTime  : 6/9/2006 14:11
' Author    : briandonorfio
' Purpose   : Returns the contents of the folder as a variant/array. Optionally, can
'             specify a glob to use as a filter, or limit to files/folders.
'---------------------------------------------------------------------------------------
'
Public Function GetFolderContents(ByVal folder As String, Optional fileglob As String = "*.*", Optional doFiles As Boolean = True, Optional doFolders As Boolean = True) As Variant
    If Right(folder, 1) <> "\" Then
        folder = folder & "\" & fileglob
    Else
        folder = folder & fileglob
    End If
    Dim fh As Long, filedata As FIND_DATA
    fh = FindFirstFile(folder, filedata)
    If fh = -1 Then
        GetFolderContents = Array()
    Else
        Dim count As Long, retval As Long, Filename As String, pos As Long, isfolder As Boolean
        ReDim fileList(10) As Variant
        retval = 1
        While retval <> ERROR_NO_MORE_FILES And retval <> 0
            isfolder = filedata.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY
            'If (filedata.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
                If (doFolders And isfolder) Or (doFiles And Not isfolder) Then
                    If count > UBound(fileList) Then
                        ReDim Preserve fileList(UBound(fileList) * 2) As Variant
                    End If
                    pos = InStr(filedata.cFileName, Chr(0))
                    If pos = 0 Then
                        Filename = filedata.cFileName
                    Else
                        Filename = Left(filedata.cFileName, pos - 1)
                    End If
                    If Filename <> "." And Filename <> ".." Then
                        fileList(count) = Filename
                        count = count + 1
                    End If
                End If
            'End If
            retval = FindNextFile(fh, filedata)
        Wend
        If count = 0 Then
            GetFolderContents = Array()
        Else
            If UBound(fileList) > count Then
                ReDim Preserve fileList(count - 1) As Variant
            End If
            GetFolderContents = fileList
        End If
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFileModifiedDate
' DateTime  : 6/9/2006 14:11
' Author    : briandonorfio
' Purpose   : Returns the last modified date for a file, or Jan 1, 1980 12:00:00 if the
'             file does not exist. The return value is a big-endian stringified date,
'             suitable for comparisons with < and >
'---------------------------------------------------------------------------------------
'
Public Function GetFileModifiedDate(Filename As String) As String
    Dim fh As Long, ft As FILETIME, st As SYSTEMTIME
    fh = CreateFile(Filename, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    If fh = -1 Then
        'GetFileModifiedDate = "1-1-1980 12:00:00"
        GetFileModifiedDate = "19800101120000"
    Else
        GetFileTime fh, ByVal 0&, ByVal 0&, ft
        FileTimeToLocalFileTime ft, ft
        FileTimeToSystemTime ft, st
        'GetFileModifiedDate = st.wMonth & "-" & st.wDay & "-" & st.wYear & " " & st.wHour & ":" & st.wMinute & ":" & st.wSecond
        GetFileModifiedDate = stringifyDate(st)
    End If
    CloseHandle fh
End Function

'---------------------------------------------------------------------------------------
' Procedure : Copy
' DateTime  : 6/9/2006 14:13
' Author    : briandonorfio
' Purpose   : Copy a file from one place to another
'---------------------------------------------------------------------------------------
'
Public Function Copy(filefrom As String, fileto As String, Optional overwrite As Boolean = False) As Boolean
    Copy = CopyFile(filefrom, fileto, IIf(overwrite, 0, 1))
End Function

'---------------------------------------------------------------------------------------
' Procedure : CreateFolder
' DateTime  : 6/9/2006 14:14
' Author    : briandonorfio
' Purpose   : Create a folder
'---------------------------------------------------------------------------------------
'
Public Function CreateFolder(foldername As String) As Boolean
    CreateFolder = CBool(CreateDirectory(foldername, ByVal &H0))
End Function

'---------------------------------------------------------------------------------------
' Procedure : filenameparse
' DateTime  : 7/31/2006 15:06
' Author    : briandonorfio
' Purpose   : Returns the last section of a path, probably a filename.
'---------------------------------------------------------------------------------------
'
Public Function filenameparse(path As String) As String
    filenameparse = Mid(path, InStrRev(path, "\") + 1)
End Function

'---------------------------------------------------------------------------------------
' Procedure : dirnameparse
' DateTime  : 7/31/2006 15:06
' Author    : briandonorfio
' Purpose   : Returns the directory portion of a path, up to the last slash. Optionally
'             can include the trailing slash, defaults to false.
'---------------------------------------------------------------------------------------
'
Public Function dirnameparse(path As String, Optional trailingSlash As Boolean = False) As String
    dirnameparse = Left(path, InStrRev(path, "\") - IIf(trailingSlash, 0, 1))
End Function


'---------------------------------------------------------------------------------------
' Procedure : stringifyDate
' DateTime  : 6/9/2006 14:14
' Author    : briandonorfio
' Purpose   : INTERNAL FUNCTION, stringifies the date. Year, month, day, hour, min, sec,
'             zero-padded if necessary.
'---------------------------------------------------------------------------------------
'
Private Function stringifyDate(st As SYSTEMTIME) As String
    stringifyDate = st.wYear & _
                    IIf(Len(CStr(st.wMonth)) = 1, "0", "") & st.wMonth & _
                    IIf(Len(CStr(st.wDay)) = 1, "0", "") & st.wDay & _
                    IIf(Len(CStr(st.wHour)) = 1, "0", "") & st.wHour & _
                    IIf(Len(CStr(st.wMinute)) = 1, "0", "") & st.wMinute & _
                    IIf(Len(CStr(st.wSecond)) = 1, "0", "") & st.wSecond
End Function

