Attribute VB_Name = "Base64Functions"
Option Explicit

Private Const BASE64_CHARACTER_SET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

Public Function Base64Encode(data As String, Optional doCrLfs As Boolean = True) As String
    Dim retval As String
    retval = ""
    Dim i As Long
    For i = 1 To Len(data) Step 3
        Dim group As Long
        group = &H10000 * Asc(Mid(data, i + 0, 1))
        If Len(data) >= i + 1 Then
            group = group + &H100 * Asc(Mid(data, i + 1, 1))
            If Len(data) >= i + 2 Then
                group = group + Asc(Mid(data, i + 2, 1))
            End If
        End If
        
        Dim octal As String
        octal = Oct(group)
        octal = String(8 - Len(octal), "0") & octal
        
        Dim part As String
        part = Mid(BASE64_CHARACTER_SET, CLng("&o" & Mid(octal, 1, 2)) + 1, 1) & _
               Mid(BASE64_CHARACTER_SET, CLng("&o" & Mid(octal, 3, 2)) + 1, 1) & _
               Mid(BASE64_CHARACTER_SET, CLng("&o" & Mid(octal, 5, 2)) + 1, 1) & _
               Mid(BASE64_CHARACTER_SET, CLng("&o" & Mid(octal, 7, 2)) + 1, 1)
        
        retval = retval & part
        
        If doCrLfs Then
            If (i + 2) Mod 57 = 0 Then
                retval = retval & vbCrLf
            End If
        End If
    Next i
    
    Select Case Len(data) Mod 3
        Case 1:
            retval = Left(retval, Len(retval) - 2) + "=="
        Case 2:
            retval = Left(retval, Len(retval) - 1) + "="
    End Select
    
    Base64Encode = retval
End Function

Public Function Base64Decode(base64 As String) As String
    base64 = Replace(base64, vbCrLf, "")
    If Len(base64) Mod 4 <> 0 Then
        Base64Decode = ""
        Exit Function
    End If
    
    Dim retval As String
    retval = ""
    Dim i As Long
    For i = 1 To Len(base64) Step 4
        Dim group As Long
        group = 0
        Dim numBytes As Long
        numBytes = 3
        Dim j As Long
        For j = 0 To 3 Step 1
            Dim thischar As String
            thischar = Mid(base64, i + j, 1)
            
            Dim thisData As String
            If thischar = "=" Then
                thisData = 0
                numBytes = numBytes - 1
            Else
                thisData = InStr(1, BASE64_CHARACTER_SET, thischar, vbBinaryCompare) - 1
            End If
            If thisData = -1 Then
                Base64Decode = ""
                Exit Function
            End If
            
            group = 64 * group + thisData
        Next j
        
        Dim hexGroup As String
        hexGroup = Hex(group)
        hexGroup = String(6 - Len(hexGroup), "0") & hexGroup
        
        Dim part As String
        part = Chr(CByte("&H" & Mid(hexGroup, 1, 2))) & _
               Chr(CByte("&H" & Mid(hexGroup, 3, 2))) & _
               Chr(CByte("&H" & Mid(hexGroup, 5, 2)))
        
        retval = retval & Left(part, numBytes)
    Next i
    
    Base64Decode = retval
End Function
