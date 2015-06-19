Attribute VB_Name = "CSVUtils"
'---------------------------------------------------------------------------------------
' Module    : CSVUtils
' DateTime  : 9/8/2006 13:37
' Author    : BrianDonorfio
' Purpose   : Delimited file handlers
'
'             Dependencies:
'               - Microsoft ActiveX Data Objects 2.8 Library
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : ExportAsCSV
' DateTime  : 1/6/2006 11:01
' Author    : briandonorfio
' Purpose   : Given a full path and filename, this exports a CSV file with the
'             information from the sql statement. If columnnames is included, the string
'             will contain a header line with column names, otherwise no header line.
'             Columnnames should be a Variant/Array, like the return from the Array()
'             function.
'---------------------------------------------------------------------------------------
'
Public Sub ExportAsCSV(sql As String, filepathname As String, Optional columnnames As Variant = Null, Optional removeCrLfs As Boolean = False)
    Dim col As Variant, line As String
On Error Resume Next
openfile:
    Open filepathname For Output As #1
    If Err.Number = 70 Then 'permission denied
        MsgBox "Unable to open """ & filepathname & """, you probably have the file open in Excel." & vbCrLf & vbCrLf & "Close Excel and click OK."
        Err.Clear
        GoTo openfile
    End If
On Error GoTo 0
    If VarType(columnnames) And vbArray Then
        For Each col In columnnames
            line = line & FormatStrForCSV(CStr(col), True) & ","
        Next col
        Print #1, Left(line, Len(line) - 1)
    End If
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve(sql)
    While Not rst.EOF
        DoEvents
        line = ""
        For Each col In rst.fields
            line = line & FormatStrForCSV(CStr(Nz(col)), removeCrLfs) & ","
        Next col
        Print #1, Left(line, Len(line) - 1)
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Close #1
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FormatStrForCSV
' DateTime  : 4/27/2006 15:19
' Author    : briandonorfio
' Purpose   : Formats a string for insertion into a CSV file, escaping double quotes,
'             commas, carriage returns. Optionally can remove all line endings, and
'             replace with spaces.
'---------------------------------------------------------------------------------------
'
Public Function FormatStrForCSV(line As String, Optional removeCrLfs = False, Optional forceEscape As Boolean = False) As String
    Dim retval As String
    If removeCrLfs Then
        retval = Replace(Replace(Replace(line, vbCrLf, " "), vbLf, " "), vbCr, " ")
    Else
        retval = line
    End If
    If forceEscape Or InStr(retval, """") Or InStr(retval, ",") Or InStr(retval, vbCr) Or InStr(retval, vbLf) Or InStr(retval, "|") Then
        retval = """" & Replace(retval, """", """""") & """"
    Else
        retval = retval
    End If
    FormatStrForCSV = retval
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadCSV
' Author    : briandonorfio
' Date      : 9/10/2012
' Purpose   : Dumb CSV reading function, returns array of dictionaries, one per record
'             with column name/value pairs. Assumes that all entries are 1 line per
'             record, with no embedded newlines (like CSV_XS). Will probably totally die
'             on badly formatted CSVs.
'---------------------------------------------------------------------------------------
'
Public Function ReadCSV(file As String) As Dictionary()
    Open file For Input As #1
    
    Dim rawLines As PerlArray
    Set rawLines = New PerlArray
    
    Dim nextline As String
    While Not EOF(1)
        Line Input #1, nextline
        nextline = TrimWhitespace(nextline)
        If nextline = "" Then
            'next
        Else
            rawLines.Push nextline
        End If
    Wend
    
    ReDim retval(0 To rawLines.Scalar - 2) As Dictionary
    
    Dim i As Long, j As Long, currentPieces As Variant, currentColumns As PerlArray, headerNames As PerlArray
    Set headerNames = Nothing
    
    For i = 0 To rawLines.Scalar - 1
        'assuming 1 row == 1 record, no newlines in entries!
        'if we want to change that, then we'll need to keep state between
        'row changes, and not initialize currentColumns selectively depending
        'on the state of the current cell quoting.
        Set currentColumns = New PerlArray
        
        currentPieces = Split(rawLines.Elem(i), ",")
        
        Dim piece As String
        piece = ""
        For j = LBound(currentPieces) To UBound(currentPieces)
            piece = piece & currentPieces(j)
            If Len(piece) = 0 Or piece = """""" Then
                'zero length string
                currentColumns.Push ""
                piece = ""
            ElseIf Left(piece, 1) <> """" Then
                'simple non-quoted string, just add to currentColumns
                currentColumns.Push piece
                piece = ""
            Else
                'complex, we are inside a double quoted string. check for matching quotes throughout and
                'only push when the piece is okay. if not okay, then we need to join with the next piece.
                If Right(piece, 1) = """" And Right(piece, 2) <> """""" Then
                    'ends with single quote, that's good. unescape interior quotes, and add without exterior
                    piece = Replace(piece, """""", """")
                    currentColumns.Push Mid(piece, 2, Len(piece) - 2)
                    piece = ""
                Else
                    'doesn't end with quote, so save this piece for later with the literal comma
                    piece = piece & ","
                End If
            End If
        Next j
        
        If headerNames Is Nothing Then
            'header hasn't been done yet, so these current columns are the header values
            Set headerNames = currentColumns
            Set currentColumns = Nothing
        Else
            'add a new dictionary to the return value, using found header names and current
            'column values
            Set retval(i - 1) = New Dictionary
            For j = 0 To headerNames.Scalar - 1
                retval(i - 1).Add CStr(headerNames.Elem(j)), CStr(currentColumns.Elem(j))
            Next j
        End If
    Next i
    
    ReadCSV = retval
End Function
