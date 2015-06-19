VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private soundexlen As Long

Private Sub Class_Initialize()
    soundexlen = 4
End Sub

Public Property Get Length() As Long
    Length = soundexlen
End Property

Public Property Let Length(newlen As Long)
    soundexlen = newlen
End Property

Public Function SoundEx(ByVal word As String) As String
    word = UCase(word)
    Dim firstletter As String
    Dim i As Long
    For i = 1 To Len(word)
        If Mid(word, i, 1) Like "[!A-Z]" Then
            word = Replace(word, Mid(word, i, 1), " ")
        End If
    Next i
    word = Trim(word)
    
    If Len(word) < 1 Then
        Exit Function
    End If
    
    firstletter = Mid(word, 1, 1)
    If firstletter = "H" Or firstletter = "W" Then
        word = "-" & Mid(word, 2)
    End If

    word = tr(word, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", _
                    "01230120022455012623010202")

    Dim newword As String
    For i = 2 To Len(word)
        If Mid(word, i, 1) <> Mid(word, i - 1, 1) Then
            newword = newword & Mid(word, i, 1)
        End If
    Next i
    word = newword
    
    word = Replace(word, "0", "")
    word = firstletter & word
    If Len(word) < soundexlen Then
        word = word & String(soundexlen - Len(word), "0")
    Else
        word = Left(word, soundexlen)
    End If
    SoundEx = word
    
End Function

Public Function Difference(word1 As String, word2 As String) As Long
    Dim s1 As String, s2 As String, retval As Long, i As Long
    retval = soundexlen
    s1 = SoundEx(word1)
    s2 = SoundEx(word2)
    For i = 1 To soundexlen
        If Mid(s1, i, 1) <> Mid(s2, i, 1) Then
            retval = retval - 1
        End If
    Next i
    Difference = retval
End Function
