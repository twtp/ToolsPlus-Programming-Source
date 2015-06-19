VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AudioPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpszCommand As String, ByVal lpszReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpszBuffer As String, ByVal uLength As Long) As Long

Private loadedFile As String

Private nowPlaying As Boolean
Private nowPaused As Boolean

Private Sub Class_Initialize()
    loadedFile = ""
End Sub

Private Sub Class_Terminate()
    CloseFile
End Sub


Public Function OpenFile(file As String) As Boolean
    If loadedFile <> "" Then
        CloseFile
    End If
    Dim retval As Long
    retval = mciSendString("OPEN """ & file & """ ALIAS AudioFile", vbNullString, 0&, 0&)
    If retval = 0 Then
        loadedFile = file
        OpenFile = True
    Else
        Call errorHandler(retval)
    End If
End Function

Public Function CloseFile() As Boolean
    If loadedFile = "" Then
        CloseFile = True
    Else
        Dim retval As Long
        If nowPlaying Then
            retval = mciSendString("STOP AudioFile", vbNullString, 0&, 0&)
            If retval = 0 Then
                nowPlaying = False
                nowPaused = False
            Else
                'skip errors here?
            End If
        End If
        retval = mciSendString("CLOSE AudioFile", vbNullString, 0&, 0&)
        If retval = 0 Then
            CloseFile = True
        Else
            CloseFile = False
        End If
    End If
End Function


Public Function Play() As Boolean
    If loadedFile = "" Then
        errorNoFile
        Play = False
    ElseIf nowPlaying Then
        'do nothing
        Play = True
    Else
        Dim retval As Long
        retval = mciSendString("PLAY AudioFile", vbNullString, 0&, 0&)
        If retval = 0 Then
            nowPlaying = True
            Play = True
        Else
            Call errorHandler(retval)
            Play = False
        End If
    End If
End Function

Public Function Pause() As Boolean
    If loadedFile = "" Then
        errorNoFile
        Pause = False
    ElseIf Not nowPlaying Then
        'no error msg
        Pause = False
    Else
        Dim retval As Long
        If nowPaused Then
            retval = mciSendString("RESUME AudioFile", vbNullString, 0&, 0&)
        Else
            retval = mciSendString("PAUSE AudioFile", vbNullString, 0&, 0&)
        End If
        If retval = 0 Then
            nowPaused = Not nowPaused
            Pause = True
        Else
            Call errorHandler(retval)
            Pause = False
        End If
    End If
End Function

Public Function StopNow(Optional autoCloseFile As Boolean = True) As Boolean
    If loadedFile = "" Then
        StopNow = True
    ElseIf Not nowPlaying Then
        StopNow = True
    Else
        Dim retval As Long
        retval = mciSendString("STOP AudioFile", vbNullString, 0&, 0&)
        If retval = 0 Then
            nowPlaying = False
            StopNow = True
        Else
            Call errorHandler(retval)
        End If
        If autoCloseFile Then
            CloseFile
        End If
    End If
End Function

Public Function Status()
    Dim retval As Long
    Dim retstr As String * 256
    retstr = String(256, " ")
    Dim retlen As Long
    retlen = Len(retstr)
    retval = mciSendString("STATUS AudioFile MODE", retstr, retlen, 0&)
    If retval = 0 Then
        Debug.Print retstr
    Else
        errorHandler retval
    End If
End Function


Private Sub errorHandler(errorCode As Long)
    CloseFile
    Dim retval As Long
    Dim errNum As Long, buf As String * 256
    retval = mciGetErrorString(errNum, buf, Len(buf))
    buf = Left(buf, Len(buf) - 1)
    MsgBox "Error: " & errNum & vbCrLf & vbCrLf & buf, vbOKOnly + vbCritical
End Sub

Private Sub errorNoFile()
    MsgBox "Error: No file loaded", vbOKOnly + vbCritical
End Sub
