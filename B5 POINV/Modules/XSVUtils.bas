Attribute VB_Name = "XSVUtils"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : ReadXSV
' DateTime  : 9/11/2006 09:25
' Author    : briandonorfio
' Purpose   : Extremely simple xSV reader. Returns an array of Dictionary objects, each
'             indexed by field name. File must contain header row. FILE MUST NOT CONTAIN
'             EMBEDDED DELIMITER CHARS/NEWLINES/ETC.
'---------------------------------------------------------------------------------------
'
Public Function ReadXSV(Filename As String, Optional delim As String = ",") As Variant
    Open Filename For Input As #1
    Dim buf As String
    Dim colNames As Variant, vals As Variant
    Dim i As Long
    Dim thisLine As Dictionary
    Dim retval() As Dictionary
    Line Input #1, buf
    colNames = Split(buf, delim)
    While Not EOF(1)
        On Error Resume Next
        ReDim Preserve retval(UBound(retval) + 1) As Dictionary
        If Err.Number = 9 Then
            Err.Clear
            ReDim Preserve retval(0) As Dictionary
        End If
        On Error GoTo 0
        Line Input #1, buf
        vals = Split(buf, delim)
        Set thisLine = New Dictionary
        For i = 0 To UBound(vals)
            thisLine.Add CStr(colNames(i)), CStr(vals(i))
        Next i
        Set retval(UBound(retval)) = thisLine
        Set thisLine = Nothing
    Wend
    Close #1
    ReadXSV = retval
End Function

