Attribute VB_Name = "ValidationFunctions"
'---------------------------------------------------------------------------------------
' Module    : ValidationFunctions
' DateTime  : 6/14/2005 10:39
' Author    : briandonorfio
' Purpose   : Validates field contents before updating db.
'
'             Dependencies:
'               - Utilities
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : validateItemNumber
' DateTime  : 6/14/2005 10:40
' Author    : briandonorfio
' Purpose   : Validates an item number (including line code). Item numbers must be only
'             alphanumeric characters or a dash, and no more than 15 characters.
'             UPDATE: 9/12/2006: added chars period and underscore.
'             UPDATE: 9/12/2013: expanded length to 27/30
'---------------------------------------------------------------------------------------
'
Public Function validateItemNumber(item As String) As Boolean
    If Len(item) < 4 Then
        validateItemNumber = False
    ElseIf (Left(item, 3) = "XXX" Or Left(item, 3) = "ZZZ") And Len(item) > 30 Then
        validateItemNumber = False
    ElseIf Len(item) > 27 Then
        validateItemNumber = False
    Else
        Dim i As Long
        For i = 1 To Len(item)
            If IsAlphaNumeric(Mid(item, i, 1)) Or Mid(item, i, 1) = "-" Or _
                                                  Mid(item, i, 1) = "." Or _
                                                  Mid(item, i, 1) = "_" Then
                'ok
            Else
                validateItemNumber = False
                Exit Function
            End If
        Next i
        validateItemNumber = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : validateLineCode
' DateTime  : 6/14/2005 10:40
' Author    : briandonorfio
' Purpose   : Validates a line code. Must be three characters, alphanumeric or dash.
'---------------------------------------------------------------------------------------
'
Public Function validateProductLine(prodLine As String) As Boolean
    'Dim re As New RegExp
    're.Pattern = "[^-A-Z0-9]"
    're.IgnoreCase = False
    If Len(prodLine) <> 3 Then 'Or re.test(LineCode) Then
        validateProductLine = False
    Else
        Dim i As Long
        For i = 1 To Len(prodLine)
            If IsAlphaNumeric(Mid(prodLine, i, 1)) Or Mid(prodLine, i, 1) = "-" Then
                'ok
            Else
                validateProductLine = False
                Exit Function
            End If
        Next i
        validateProductLine = True
    End If
    'Set re = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : PLExists
' DateTime  : 6/15/2005 16:38
' Author    : briandonorfio
' Purpose   : Checks if a prod line exists in the ProductLine table.
'---------------------------------------------------------------------------------------
'
Public Function PLExists(PL As String) As Boolean
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ProductLine FROM ProductLine WHERE ProductLine='" & PL & "'")
    PLExists = Not rst.EOF
    rst.Close
    Set rst = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : validatePartNum
' DateTime  : 6/14/2005 10:42
' Author    : briandonorfio
' Purpose   : Validates a part number (no line code). Alphanumeric or dash, no more than
'             12 characters.
'---------------------------------------------------------------------------------------
'
Public Function validatePartNum(partNum As String) As Boolean
    'Dim re As New RegExp
    're.Pattern = "[^-A-Z0-9]"
    're.IgnoreCase = False
    If Len(partNum) < 1 Or Len(partNum) > 16 Then 'Or re.test(partNum) Then
        validatePartNum = False
    Else
        Dim i As Long
        For i = 1 To Len(partNum)
            If IsAlphaNumeric(Mid(partNum, i, 1)) Or Mid(partNum, i, 1) = "-" Then
                'ok
            Else
                validatePartNum = False
                Exit Function
            End If
        Next i
        validatePartNum = True
    End If
    'Set re = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : validateUPC
' DateTime  : 6/14/2005 10:42
' Author    : briandonorfio
' Purpose   : Validates a UPC Code. Numeric (dashes should be removed), no more than 12
'             characters. No lower limit?
'---------------------------------------------------------------------------------------
'
Public Function validateUPC(UPC As String) As Boolean
    If Not IsNumeric(UPC) Or Len(UPC) > 13 Then
        validateUPC = False
    Else
        validateUPC = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : validateItemDescription
' DateTime  : 6/14/2005 10:43
' Author    : briandonorfio
' Purpose   : Validates an Item Description. Mas200 chokes on double quote and comma
'             characters. Less than 255 characters.
'---------------------------------------------------------------------------------------
'
Public Function validateItemDescription(itemDesc As String) As Boolean
    If Len(itemDesc) > 255 Or InStr(itemDesc, """") Or InStr(itemDesc, ",") Then
        validateItemDescription = False
    Else
        validateItemDescription = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : validateEPN
' DateTime  : 6/14/2005 10:44
' Author    : briandonorfio
' Purpose   : Validates EPN. Only restriction is length, less than 512. Changed by Tom to 8000.
'---------------------------------------------------------------------------------------
'
Public Function validateEPN(EPN As String) As Boolean
    If Len(EPN) > 8000 Then
        validateEPN = False
    Else
        validateEPN = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : validateComponents
' DateTime  : 6/14/2005 10:44
' Author    : briandonorfio
' Purpose   : Validates Components. Only restriction is length, less than 255.
'---------------------------------------------------------------------------------------
'
Public Function validateComponents(comps As String) As Boolean
    If Len(comps) > 255 Then
        validateComponents = False
    Else
        validateComponents = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : validateDimensions
' DateTime  : 6/14/2005 10:44
' Author    : briandonorfio
' Purpose   : Validates Dimensions. Returns the canonicalized dimension string, or null
'             string if failure. Dimensions will be formatted something like "3x2x1" for
'             single box items, and like "3x2x1; 3x2x1" for multi-box items. Longest
'             dimension always first, must have a semicolon-space between boxes.
'---------------------------------------------------------------------------------------
'
Public Function validateDimensions(Dimensions As String) As String
    If Dimensions = "" Then
        Dimensions = "0x0x0"
    End If
    Dimensions = Replace(Dimensions, "*", "x")
    'Dimensions = Replace(Dimensions, ":", ";")
    Dimensions = UCase$(Dimensions)
    'Dim boxarray() As String, dimarray() As String
    Dim dimarray() As String
    Dim one As Double, two As Double, three As Double, temp As Double
    'boxarray = Split(Dimensions, ";")
    'Dim i As Long
    'For i = 0 To UBound(boxarray)
        'boxarray(i) = Trim(CStr(boxarray(i)))
        Dimensions = Trim(Dimensions)
        'dimarray = Split(boxarray(i), "X")
        dimarray = Split(Dimensions, "X")
        If UBound(dimarray) <> 2 Then
            validateDimensions = "ERROR:dimensions"
            Exit Function
        End If
        dimarray(0) = ConvertFracToDec(Trim$(dimarray(0)))
        dimarray(1) = ConvertFracToDec(Trim$(dimarray(1)))
        dimarray(2) = ConvertFracToDec(Trim$(dimarray(2)))
        If Not IsNumeric(dimarray(0)) Or Not IsNumeric(dimarray(1)) Or Not IsNumeric(dimarray(2)) Then
            validateDimensions = "ERROR:numeric"
            Exit Function
        End If
        one = CDbl(dimarray(0))
        two = CDbl(dimarray(1))
        three = CDbl(dimarray(2))
        If one < two Then
            temp = one
            one = two
            two = temp
        End If
        If two < three Then
            temp = two
            two = three
            three = temp
        End If
        If one < two Then
            temp = one
            one = two
            two = temp
        End If
        'boxarray(i) = CStr(one) & "x" & CStr(two) & "x" & CStr(three)
    'Next i
    'validateDimensions = Join(boxarray, "; ")
    validateDimensions = Round(CStr(one), 2) & "x" & Round(CStr(two), 2) & "x" & Round(CStr(three), 2)
End Function

'---------------------------------------------------------------------------------------
' Procedure : validateShippingClass
' DateTime  : 6/14/2005 10:46
' Author    : briandonorfio
' Purpose   : Validates Shipping Class. Must be in the ShipClass table.
'---------------------------------------------------------------------------------------
'
Public Function validateShippingClass(shipClass As String) As Boolean
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ShipClassName FROM ShipClass WHERE ShipClassName='" & shipClass & "'")
    If rst.EOF Then
        validateShippingClass = False
    Else
        validateShippingClass = True
    End If
    rst.Close
    Set rst = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : validateWeight
' DateTime  : 6/14/2005 10:47
' Author    : briandonorfio
' Purpose   : Validates Weight. Returns canonicalized weight string, unit labels etc.
'             removed. Multiple boxes handled same as dimensions, with a semicolon-space
'             between them. Should be in same order as box list for dimensions.
'---------------------------------------------------------------------------------------
'
'Public Function validateWeight(Weight As String) As String
'    Weight = LCase(Weight)
'    Weight = Replace(Weight, "lbs.", "")
'    Weight = Replace(Weight, "lbs", "")
'    Weight = Replace(Weight, "lb.", "")
'    Weight = Replace(Weight, "lb", "")
'    Weight = Replace(Weight, ":", ";")
'    Dim weightarray() As String
'    weightarray = Split(Weight, ";")
'    Dim i As Long
'    For i = 0 To UBound(weightarray)
'        weightarray(i) = Trim$(weightarray(i))
'        weightarray(i) = ConvertFracToDec(weightarray(i))
'        If Not IsNumeric(weightarray(i)) Then
'            validateWeight = ""
'            Exit Function
'        End If
'    Next i
'    validateWeight = Join(weightarray, "; ")
'End Function
Public Function validateWeight(wt As String) As String
    wt = LCase(wt)
    wt = Replace(wt, "lbs.", "")
    wt = Replace(wt, "lbs", "")
    wt = Replace(wt, "lb.", "")
    wt = Replace(wt, "lb", "")
    wt = Trim$(wt)
    wt = ConvertFracToDec(wt)
    If Left(wt, 1) = "." Then
        wt = "0" & wt
    End If
    validateWeight = IIf(IsNumeric(wt), wt, "")
End Function

'---------------------------------------------------------------------------------------
' Procedure : validateDropShipTerms
' DateTime  : 6/14/2005 10:48
' Author    : briandonorfio
' Purpose   : Validates DS Terms. Only restriction is length, less than 255.
'---------------------------------------------------------------------------------------
'
Public Function validateDropShipTerms(dsterms As String) As Boolean
    If Len(dsterms) > 255 Then
        validateDropShipTerms = False
    Else
        validateDropShipTerms = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : validateSpecs
' DateTime  : 6/14/2005 10:52
' Author    : briandonorfio
' Purpose   : Validates item specifications. Must be less than 512. Extended specs don't
'             have any restrictions (memo field).
'---------------------------------------------------------------------------------------
'
Public Function validateSpecs(Spec As String) As Boolean
    If Len(Spec) > 512 Then
        validateSpecs = False
    Else
        validateSpecs = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : validateCurrency
' DateTime  : 6/14/2005 10:53
' Author    : briandonorfio
' Purpose   : Validates prices, costs, etc. Must be numeric, greater than 0.
'---------------------------------------------------------------------------------------
'
Public Function validateCurrency(cur As String) As Boolean
    If IsNumeric(cur) Then
        If CDbl(cur) >= 0 Then
            validateCurrency = True
        Else
            validateCurrency = False
        End If
    Else
        validateCurrency = False
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : validateItemPOComment
' DateTime  : 6/14/2005 10:53
' Author    : briandonorfio
' Purpose   : Validates PO Comments for items. Less than 50 chars.
'---------------------------------------------------------------------------------------
'
Public Function validateItemPOComment(Comment As String) As Boolean
    If Len(Comment) > 50 Then
        validateItemPOComment = False
    Else
        validateItemPOComment = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : validatePOComment
' DateTime  : 6/14/2005 10:54
' Author    : briandonorfio
' Purpose   : Validates the overall PO Comment. Less than 255 chars.
'---------------------------------------------------------------------------------------
'
Public Function validatePOComment(Comment As String) As Boolean
    If Len(Comment) > 255 Then
        validatePOComment = False
    Else
        validatePOComment = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : validateManufFullName
' DateTime  : 6/16/2005 14:52
' Author    : briandonorfio
' Purpose   : Validates the ManufFullName and ManufFullNameCleaned. <32 chars.
'---------------------------------------------------------------------------------------
'
Public Function validateManufFullName(name As String) As Boolean
    If Len(name) > 32 Then
        validateManufFullName = False
    Else
        validateManufFullName = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : validateGeneralInteger
' DateTime  : 6/16/2005 15:06
' Author    : briandonorfio
' Purpose   : Validates std lead time, # of po payments, etc. Anything that just needs
'             to be an integer.
'---------------------------------------------------------------------------------------
'
Public Function validateGeneralInteger(num As String) As Boolean
    If IsNumeric(num) And Not CBool(InStr(num, ".")) Then
        validateGeneralInteger = True
    Else
        validateGeneralInteger = False
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : validateDescription
' DateTime  : 7/1/2005 13:17
' Author    : briandonorfio
' Purpose   : Validates the PartNumbers description, not to be confused with the
'             ItemDescription. Less than 512
'---------------------------------------------------------------------------------------
'
Public Function validateDescription(Desc As String) As Boolean
    If Len(Desc) <= 512 Then
        validateDescription = True
    Else
        validateDescription = False
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : validateWebCaption
' DateTime  : 7/6/2005 12:31
' Author    : briandonorfio
' Purpose   : Validates the web caption. Can be <=512 chars. If blank, on the offload it
'             cat's descs+specs for caption.
'---------------------------------------------------------------------------------------
'
Public Function validateWebCaption(Caption As String) As Boolean
    If Len(Caption) <= 512 Then
        validateWebCaption = True
    Else
        validateWebCaption = False
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : validateZipCode
' DateTime  : 3/23/2006 15:16
' Author    : briandonorfio
' Purpose   : Validates a US (in 5 or 5+4) or Canadian (6 letters/numbers, possibly a
'             space) zip code.
'---------------------------------------------------------------------------------------
'
Public Function validateZipCode(ZipCode As String) As Boolean
    'Dim uszip As RegExp, cazip As RegExp
    'Set uszip = New RegExp
    'Set cazip = New RegExp
    'uszip.Pattern = "^\d{5}(-\d{4})?$"
    'cazip.Pattern = "^[\d\w]{3}\s?[\d\w]{3}$"
    'cazip.IgnoreCase = True
    'If uszip.test(zipcode) Then
    '    validateZipCode = True
    'ElseIf cazip.test(zipcode) Then
    '    validateZipCode = True
    'Else
    '    validateZipCode = False
    'End If
    If validateUSZipCode(ZipCode) Then
        validateZipCode = True
    ElseIf validateCAZipCode(ZipCode) Then
        validateZipCode = True
    Else
        validateZipCode = False
    End If
End Function

Public Function validateUSZipCode(ZipCode As String) As Boolean
    If Len(ZipCode) = 5 Or Len(ZipCode) = 9 Then 'reg zip or zip+4 w/o dash
        If IsNumeric(ZipCode) Then
            validateUSZipCode = True
        Else
            validateUSZipCode = False
        End If
    ElseIf Len(ZipCode) = 10 Then
        If IsNumeric(Left(ZipCode, 5)) And Mid(ZipCode, 6, 1) = "-" And IsNumeric(Right(ZipCode, 4)) Then
            validateUSZipCode = True
        Else
            validateUSZipCode = False
        End If
    Else
        validateUSZipCode = False
    End If
End Function

Public Function validateCAZipCode(ByVal ZipCode As String) As Boolean
    If Len(ZipCode) = 7 And Mid(ZipCode, 4, 1) = " " Or Mid(ZipCode, 4, 1) = "-" Then
        ZipCode = Left(ZipCode, 3) & Right(ZipCode, 3)
    End If
    If Len(ZipCode) = 6 Then
        If IsAlphaNumeric(ZipCode) Then
            validateCAZipCode = True
        Else
            validateCAZipCode = False
        End If
    Else
        validateCAZipCode = False
    End If
End Function
