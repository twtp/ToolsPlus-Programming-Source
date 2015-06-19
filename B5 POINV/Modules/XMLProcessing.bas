Attribute VB_Name = "XMLProcessing"
Option Explicit

Public Function escapeXMLEntities(ByVal txt As String) As String
    txt = Replace(txt, "&", "&amp;")
    txt = Replace(txt, """", "&quot;")
    txt = Replace(txt, "<", "&lt;")
    txt = Replace(txt, ">", "&gt;")
    escapeXMLEntities = txt
End Function

Public Function removeUpperASCII(str As String) As String
    Dim i As Long
    Dim retval As String
    For i = 1 To Len(str)
        If Asc(Mid(str, i, 1)) < 127 Then
            retval = retval & Mid(str, i, 1)
        End If
    Next i
    removeUpperASCII = retval
End Function
