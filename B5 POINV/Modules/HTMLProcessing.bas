Attribute VB_Name = "HTMLProcessing"
'---------------------------------------------------------------------------------------
' Module    : HTMLProcessing
' DateTime  : 8/11/2006 12:34
' Author    : BrianDonorfio
' Purpose   : Convert to/from text/HTML
'---------------------------------------------------------------------------------------

Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : StripHTML
' DateTime  : 7/5/2005 17:10
' Author    : briandonorfio
' Purpose   : Removes html tags from text. Optional tagsOnly argument will strip tags,
'             but leave html entities.
'---------------------------------------------------------------------------------------
'
Public Function StripHTML(ByVal txt As String, Optional tagsOnly As Boolean = False) As String
    Dim startPos As Long, endpos As Long, tag As String, taglen As Long
    txt = Replace(txt, vbCrLf, "")
    'first, scan for any comments
    startPos = InStr(txt, "<!--")
    endpos = InStr(txt, "-->")
    taglen = 3
    If endpos = 0 Then
        endpos = InStr(txt, "--!>")
        taglen = 4
    End If
    While startPos <> 0 And endpos <> 0
        tag = Mid(txt, startPos, endpos - startPos + taglen)
        txt = Replace(txt, tag, "")
        startPos = InStr(txt, "<!--")
        endpos = InStr(txt, "-->")
        taglen = 3
        If endpos = 0 Then
            endpos = InStr(txt, "--!>")
            taglen = 4
        End If
    Wend
    'now do the other tags
    startPos = InStr(txt, "<")
    endpos = InStr(txt, ">")
    While startPos <> 0 And endpos <> 0 And endpos > startPos
        tag = Mid(txt, startPos, endpos - startPos + 1)
        Select Case LCase(tag)
            Case Is = "<li>"
                txt = Replace(txt, tag, vbCrLf & "  " & Chr(149) & " ") '149 = bullet
            Case Is = "<p>"
                txt = Replace(txt, tag, vbCrLf & vbCrLf)
            Case Is = "<br>", "<br />"
                txt = Replace(txt, tag, vbCrLf)
            Case Is = "<h1>", "<h2>", "<h3>", "<h4>", "<h5>"
                txt = Replace(txt, tag, vbCrLf)
            Case Else
                txt = Replace(txt, tag, "")
        End Select
        startPos = InStr(txt, "<")
        endpos = InStr(txt, ">")
    Wend
    If Not tagsOnly Then
        txt = Replace(txt, "&nbsp;", " ")
        txt = Replace(txt, "&amp;", "&")
        txt = Replace(txt, "&quot;", """")
        txt = Replace(txt, "&lt;", "<")
        txt = Replace(txt, "&gt;", ">")
        txt = Replace(txt, "&deg;", Chr(186))   '186 = degree symbol
        txt = Replace(txt, "&plusmn;", "+/-")
        txt = Replace(txt, "&reg;", "")
        txt = Replace(txt, "&trade;", "")
    End If
    
    While InStr(txt, vbCrLf) = 1            'remove leading cr's
        txt = Replace(txt, vbCrLf, "", , 1)
    Wend
    StripHTML = txt
End Function

'---------------------------------------------------------------------------------------
' Procedure : html2text
' DateTime  : 8/1/2006 15:33
' Author    : briandonorfio
' Purpose   : Synonym for StripHTML
'---------------------------------------------------------------------------------------
'
Public Function html2text(ByVal txt As String, Optional tagsOnly As Boolean = False) As String
    html2text = StripHTML(txt, tagsOnly)
End Function

Public Function escapeHTMLEntities(ByVal txt As String) As String
    txt = Replace(txt, "&", "&amp;")
    txt = Replace(txt, """", "&quot;")
    txt = Replace(txt, "<", "&lt;")
    txt = Replace(txt, ">", "&gt;")
    txt = Replace(txt, Chr(186), "&deg;")
    escapeHTMLEntities = txt
End Function

'---------------------------------------------------------------------------------------
' Procedure : text2html
' DateTime  : 8/1/2006 15:33
' Author    : briandonorfio
' Purpose   : Changes text to a slightly HTMLish version of the text. Converts vbcrlf-
'             delimited text to paragraphs, 4+ dashes on a line to hrules, quoted
'             full paragraphs to blockquotes. Given a regex, it will replace all found
'             instances with an image tag. First capture should be the alignment for
'             the image, second capture should be the file number in BlogImages.
'---------------------------------------------------------------------------------------
'
Public Function text2html(ByVal txt As String, Optional picRE As String = "") As String
    txt = Replace(txt, "&", "&amp;")
    txt = Replace(txt, "<", "&lt;")
    txt = Replace(txt, ">", "&gt;")
    txt = Replace(txt, """", "&quot;")
    While Left(txt, 2) = vbCrLf
        txt = Mid(txt, 3)
    Wend
    While Right(txt, 2) = vbCrLf
        txt = Left(txt, Len(txt) - 2)
    Wend
    Dim Lines As Variant
    Lines = Split(txt, vbCrLf)
    Dim re As RegExp, m As MatchCollection
    Set re = New RegExp
    Dim i As Long, j As Long, html As String, InUL As Boolean
    For i = 0 To UBound(Lines)
        re.Pattern = "^\s*\*\s*(.+)\s*$"
        If re.Test(Lines(i)) Then
            Set m = re.Execute(Lines(i))
            If Not InUL Then
                html = html & "<ul>" & vbCrLf
                InUL = True
            End If
            html = html & "<li>" & m.item(0).SubMatches(0) & "</li>" & vbCrLf
        Else
            If InUL Then
                html = html & "</ul>" & vbCrLf
                InUL = False
            End If
            re.Pattern = "^\s*[-]{4,}\s*$"
            If re.Test(Lines(i)) Then
                html = html & "<hr />" & vbCrLf
            Else
                re.Pattern = "^\s*&quot;(.+)&quot;\s*$"
                If re.Test(Lines(i)) Then
                    Set m = re.Execute(Lines(i))
                    html = html & "<blockquote><p>" & m.item(0).SubMatches(0) & "</p></blockquote>" & vbCrLf
                Else
                    Lines(i) = Trim(Lines(i))
                    'If picRE <> "" Then
                    '    re.Pattern = picRE
                    '    If re.Test(Lines(i)) Then
                    '        Set m = re.Execute(Lines(i))
                    '        For j = 0 To m.count - 1
                    '            html = html & "<img src=""" & BLOG_PIC_DIR & m.item(j).SubMatches(1) & ".jpg"" float=""" & m.item(j).SubMatches(0) & """ />" & vbCrLf
                    '        Next j
                    '    End If
                    '    Lines(i) = re.Replace(Lines(i), "")
                    'End If
                    If Lines(i) <> "" Then
                        html = html & "<p>" & Lines(i) & "</p>" & vbCrLf
                    End If
                End If
            End If
        End If
    Next i
    If InUL Then
        html = html & "</ul>" & vbCrLf
    End If
    text2html = html
End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidateHTML
' DateTime  : 9/30/2008 12:58
' Author    : briandonorfio
' Purpose   : Validates the given html, returns a variant array of errors, or a zero-
'             length array if valid. This needs to be a complete html page, including
'             doctype.
'---------------------------------------------------------------------------------------
'
Public Function ValidateHTML(html As String) As Variant
    Dim nsgmls As String, infile As String, options As String
    nsgmls = "s:\mastest\mas90-signs\utilities\nsgmls\bin\nsgmls.exe"
    infile = DESTINATION_DIR & "validate.html"
    options = "-E 100 " & _
              "-s " & _
              "-c " & qq("s:\mastest\mas90-signs\utilities\nsgmls\pubtext\html4.cat") & " " & _
              "-f " & qq(DESTINATION_DIR & "validateresults.txt")
    Open infile For Output As #1
    Print #1, html
    Close #1
    
    ShellWait nsgmls & " " & options & " " & infile, vbHide
    
    Dim out As String
    out = ""
    
    Dim buf As String
    Open DESTINATION_DIR & "validateresults.txt" For Input As #1
    While Not EOF(1)
        Line Input #1, buf
        If buf <> "" Then
            buf = Replace(buf, nsgmls & ":" & infile & ":", "", 1, 1, vbTextCompare)
            out = out & buf & vbCrLf
        End If
    Wend
    Close #1
    
    out = TrimWhitespace(out, True, True, False, True)
    ValidateHTML = Split(out, vbCrLf)
End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidateHTMLFragment
' DateTime  : 9/30/2008 12:58
' Author    : briandonorfio
' Purpose   : Same as previous function, except this will add a standard html body
'             around the given fragment. Doctype is html-401-loose, same as Yahoo.
'---------------------------------------------------------------------------------------
'
Public Function ValidateHTMLFragment(frag As String, Optional strict As Boolean = False) As Variant
    Dim html As String
    html = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01" & IIf(strict, "", " Transitional") & "//EN"">" & vbCrLf & _
           "<html>" & vbCrLf & _
           " <head><title>validate</title></head>" & _
           " <body>" & vbCrLf & _
              frag & vbCrLf & _
           " </body>" & vbCrLf & _
           "</html>"
    ValidateHTMLFragment = ValidateHTML(html)
End Function

Public Function ChangeRelativeURLsToAbsolute(html As String, base As String) As String
    If Left(base, 7) <> "http://" Then
        base = "http://" & base
    End If
    If Right(base, 1) <> "/" Then
        base = base & "/"
    End If
    
    Dim re1 As VBScript_RegExp_55.RegExp, matches As VBScript_RegExp_55.MatchCollection
    Set re1 = New VBScript_RegExp_55.RegExp
    re1.Pattern = "\bhref=""([^""]+)"""
    re1.Global = True
    Set matches = re1.Execute(html)
    
    Dim re2 As VBScript_RegExp_55.RegExp
    Set re2 = New VBScript_RegExp_55.RegExp
    re2.Pattern = "^[a-z]+://"
    
    If matches.count > 0 Then
        Dim thisMatch As VBScript_RegExp_55.Match
        For Each thisMatch In matches
            Dim url As String
            url = thisMatch.SubMatches(0)
            If re2.Test(url) Then
                'absolute
            ElseIf Left(url, 7) = "mailto:" Then
                'ignore
            Else
                html = Replace(html, """" & url & """", """" & base & IIf(Left(url, 1) = "/", Mid(url, 2), url) & """", 1, 1)
            End If
        Next thisMatch
    End If
    
    ChangeRelativeURLsToAbsolute = html
End Function
