Attribute VB_Name = "ConfParser"
'---------------------------------------------------------------------------------------
' Module    : ConfParser
' DateTime  : 10/28/2005 15:37
' Author    : briandonorfio
' Purpose   : Parses *nix-style conf files. Comments are started with #, variables are
'             one word, space in between, values are anything afterward.
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : ParseConfFile
' DateTime  : 10/28/2005 15:43
' Author    : briandonorfio
' Purpose   : Parses a given conf file. fname should be the full path and filename.
'             Returns a dictionary object, variable => value. Dictionary is empty if
'             specified file is not found, or other error.
'---------------------------------------------------------------------------------------
'
Public Function ParseConfFile(fname As String) As Dictionary
    Dim varH As Dictionary
    Set varH = New Dictionary
    Dim nextline As String, splitline() As String
On Error GoTo errh
    Open fname For Input As #1
    While Not EOF(1)
        Line Input #1, nextline
        nextline = TrimWhitespace(nextline)
        If nextline = "" Then
            'next
        ElseIf Left(nextline, 1) = "#" Then
            'next
        Else
            splitline = Split(nextline, "#")
            splitline(0) = TrimWhitespace(splitline(0))
            If InStr(splitline(0), " ") Then
                varH.Add TrimWhitespace(Left(splitline(0), InStr(splitline(0), " ") - 1)), TrimWhitespace(Mid(splitline(0), InStr(splitline(0), " ") + 1))
            End If
        End If
    Wend
done:
    Close #1
    Set ParseConfFile = varH
    Set varH = Nothing
    Exit Function
    
errh:
    Select Case Err.Number
        Case Is = 53, 52, 76
            MsgBox "Unable to find configuration file! Are your network drives working?" & vbCrLf & vbCrLf & "Continuing with default settings..."
            Err.Clear
            GoTo done
        Case Else
            MsgBox "Error reading configuration: " & Err.Number & vbCrLf & vbCrLf & Err.description
    End Select
End Function
