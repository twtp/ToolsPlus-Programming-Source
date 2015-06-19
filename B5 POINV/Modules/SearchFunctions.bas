Attribute VB_Name = "SearchFunctions"
'---------------------------------------------------------------------------------------
' Module    : SearchFunctions
' DateTime  : 6/14/2005 10:36
' Author    : briandonorfio
' Purpose   : Basic searching functions, sorting functions.
'
'             Dependencies:
'               - MouseHourglass + global Mouse object
'               - Microsoft Scripting Runtime (for dictionaries)
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : bsearch
' DateTime  : 6/14/2005 11:11
' Author    : briandonorfio
' Purpose   : Performs a binary search on a list, returns the index in the array if
'             found. Otherwise, if argument closeEnough is true, then it returns the
'             one before it alphabetically (or after it, if there's nothing before it),
'             if false, then it returns -1.
'---------------------------------------------------------------------------------------
'
Public Function bsearch(key As String, list() As String, Optional closeEnough As Boolean = False) As Long
    Mouse.Hourglass True
    Dim low As Long, high As Long, found As Boolean, try As Long
    low = 0
    high = UBound(list)
    While low <= high And found = False
        try = (low + high) / 2
        If UCase$(key) = UCase$(list(try)) Then
            found = True
        ElseIf UCase$(key) > UCase$(list(try)) Then
            low = try + 1
        Else
            high = try - 1
        End If
    Wend
    If found Then
        bsearch = try
    ElseIf closeEnough Then
        If high >= 0 Then
            bsearch = high
        Else
            bsearch = low
        End If
    Else
        bsearch = -1
    End If
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : lsearch
' DateTime  : 6/15/2005 15:33
' Author    : briandonorfio
' Purpose   : Basic linear search. Returns array index or -1 if not found
'---------------------------------------------------------------------------------------
'
Public Function lsearch(key As String, list() As String) As Long
    Mouse.Hourglass True
    Dim found As Boolean
    found = False
    Dim i As Long
    i = 0
    While i <= UBound(list) And Not found
        If list(i) = key Then
            found = True
        Else
            i = i + 1
        End If
    Wend
    If found Then
        lsearch = i
    Else
        lsearch = -1
    End If
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : findFirstInLC
' DateTime  : 6/14/2005 11:12
' Author    : briandonorfio
' Purpose   : bsearch+lsearch, finds the first itemnumber in a line code. Returns index
'             if found, -1 otherwise.
'---------------------------------------------------------------------------------------
'
Public Function findFirstInLC(key As String, list() As String) As Long
    Mouse.Hourglass True
    Dim low As Long, high As Long, found As Boolean, try As Long
    low = 0
    high = UBound(list)
    While low <= high And found = False
        try = (low + high) / 2
        If UCase$(key) = UCase$(Left$(list(try), 3)) Then
            found = True
        ElseIf UCase$(key) > UCase$(Left$(list(try), 3)) Then
            low = try + 1
        Else
            high = try - 1
        End If
    Wend
    If found Then
        found = False
        While Not found
            If try > 0 Then
                If Left$(list(try), 3) <> key Then
                    try = try + 1
                    found = True
                Else
                    try = try - 1
                End If
            Else
                found = True
            End If
        Wend
        findFirstInLC = try
    Else
        findFirstInLC = -1
    End If
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : InList
' DateTime  : 7/19/2005 10:43
' Author    : briandonorfio
' Purpose   : Searches for an item in a listbox, returns true if found, false otherwise.
'             Linear search, possibly better to convert listbox contents to an array,
'             then bsearch the array, but really only if there are a lot of items. API
'             function to return list as array? Optional argument for if the list is
'             sorted alphabetically already (probably from the sorted property), will
'             fire off a hopefully-faster bsearch.
'---------------------------------------------------------------------------------------
'
Public Function InList(key As String, list As ListBox, Optional sorted As Boolean = False, Optional column As Long = -1) As Boolean
    Mouse.Hourglass True
    Dim found As Boolean
    Dim testStr As String
    If sorted Or list.sorted Then
        Dim low As Long, high As Long, try As Long
        low = 0
        high = list.ListCount - 1
        While low <= high And found = False
            try = (low + high) / 2
            If column = -1 Then
                testStr = list.list(try)
            Else
                testStr = ListBoxLineColumn(list, try, column)
            End If
            
            If UCase$(key) = UCase$(testStr) Then
                found = True
            ElseIf UCase$(key) > UCase$(testStr) Then
                low = try + 1
            Else
                high = try - 1
            End If
        Wend
    Else
        Dim i As Long
        found = False
        For i = 0 To list.ListCount - 1
            If column = -1 Then
                testStr = list.list(i)
            Else
                testStr = ListBoxLineColumn(list, i, column)
            End If

            If UCase$(key) = UCase$(testStr) Then
                found = True
                i = list.ListCount
            End If
        Next i
    End If
    InList = found
    Mouse.Hourglass False
End Function

Public Function InCombo(key As String, cbo As ComboBox, Optional sorted As Boolean = False) As Boolean
    Mouse.Hourglass True
    Dim found As Boolean
    If sorted Or cbo.sorted Then
        Dim low As Long, high As Long, try As Long
        low = 0
        high = cbo.ListCount - 1
        While low <= high And found = False
            try = (low + high) / 2
            If UCase$(key) = UCase$(cbo.list(try)) Then
                found = True
            ElseIf UCase$(key) > UCase$(cbo.list(try)) Then
                low = try + 1
            Else
                high = try - 1
            End If
        Wend
    Else
        Dim i As Long
        found = False
        For i = 0 To cbo.ListCount - 1
            If cbo.list(i) = key Then
                found = True
                i = cbo.ListCount
            End If
        Next i
    End If
    InCombo = found
    Mouse.Hourglass False
End Function

Public Function InListPosition(key As String, list As ListBox) As Long
    Mouse.Hourglass True
    Dim found As Boolean, pos As Long
    If list.sorted Then
        Dim low As Long, high As Long, try As Long
        low = 0
        high = list.ListCount - 1
        While low <= high And found = False
            try = (low + high) / 2
            If UCase$(key) = UCase$(list.list(try)) Then
                found = True
                pos = try
            ElseIf UCase$(key) > UCase$(list.list(try)) Then
                low = try + 1
            Else
                high = try - 1
            End If
        Wend
    Else
        Dim i As Long
        found = False
        For i = 0 To list.ListCount - 1
            If list.list(i) = key Then
                found = True
                pos = i
                i = list.ListCount
            End If
        Next i
    End If
    Mouse.Hourglass False
    If found Then
        InListPosition = pos
    Else
        InListPosition = -1
    End If
End Function

Public Function InComboPosition(key As String, cbo As ComboBox, Optional sorted As Boolean = False) As Long
    Mouse.Hourglass True
    Dim found As Boolean, pos As Long
    If sorted Or cbo.sorted Then
        Dim low As Long, high As Long, try As Long
        low = 0
        high = cbo.ListCount - 1
        While low <= high And found = False
            try = (low + high) / 2
            If UCase$(key) = UCase$(cbo.list(try)) Then
                found = True
                pos = try
            ElseIf UCase$(key) > UCase$(cbo.list(try)) Then
                low = try + 1
            Else
                high = try - 1
            End If
        Wend
    Else
        Dim i As Long
        found = False
        For i = 0 To cbo.ListCount - 1
            If cbo.list(i) = key Then
                found = True
                pos = i
                i = cbo.ListCount
            End If
        Next i
    End If
    Mouse.Hourglass False
    If found Then
        InComboPosition = pos
    Else
        InComboPosition = -1
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : QuickSort
' DateTime  : 7/6/2005 08:38
' Author    : briandonorfio
' Purpose   : Wrapper for qsort functions. Requires array of strings, sorts descending
'             if second argument is true, otherwise defaults to ascending.
'---------------------------------------------------------------------------------------
'
Public Sub QuickSort(stringArray() As String, Optional descending As Boolean = False)
    Mouse.Hourglass True
    qsortRecurse stringArray, LBound(stringArray), UBound(stringArray), descending
    Mouse.Hourglass False
End Sub

Private Sub qsortRecurse(sarray() As String, inLow As Long, inHigh As Long, descending As Boolean)
    Dim pivot As String, swap As String, low As Long, high As Long
    
    low = inLow
    high = inHigh
    pivot = sarray((inLow + inHigh) / 2)
    
    While (low <= high)
        If Not descending Then
            While (sarray(low) < pivot And low < inHigh)
                low = low + 1
            Wend
            While (pivot < sarray(high) And high > inLow)
                high = high - 1
            Wend
        Else
            While (sarray(low) > pivot And low < inHigh)
                low = low + 1
            Wend
            While (pivot > sarray(high) And high > inLow)
                high = high - 1
            Wend
        End If
        
        If (low <= high) Then
            swap = sarray(low)
            sarray(low) = sarray(high)
            sarray(high) = swap
            low = low + 1
            high = high - 1
        End If
    Wend
    
    If (inLow < high) Then
        qsortRecurse sarray(), inLow, high, descending
    End If
    If (low < inHigh) Then
        qsortRecurse sarray(), low, inHigh, descending
    End If
End Sub

'QuickSortHashes( [ { foo => 1, bar => 2 }, { foo => 3, bar => 4 } ], array(array("foo", true), array("bar", true)) ) means sort by foo desc, bar asc
Public Sub QuickSortHashes(hashArray() As Dictionary, keys As Variant)
    If VarType(keys) = vbString Then
        keys = Array(Array(keys, False))
    End If
    If Not VarType(keys) And vbArray Then
        Err.Raise 123, "QuickSortHashes", "Type of 'keys' should be Variant/Array"
    End If
    Dim i As Long
    For i = 0 To UBound(keys)
        If VarType(keys(i)) = vbString Then
            keys(i) = Array(keys(i), False)
        ElseIf VarType(keys(i)) And vbArray Then
            If UBound(keys(i)) = 0 Then
                keys(i) = Array(keys(i)(0), False)
            ElseIf UBound(keys(i)) = 1 Then
                keys(i) = Array(keys(i)(0), CBool(keys(i)(1)))
            Else
                Err.Raise 123, "QuickSortHashes", "Element '" & i & "' of 'keys' has too many parameters, only need string key and optionally boolean asc/desc"
            End If
        Else
            Err.Raise 123, "QuickSortHashes", "Element '" & i & "' of 'keys' should be String or Variant/Array"
        End If
    Next i
    Mouse.Hourglass True
    qsortHashRecurse hashArray, LBound(hashArray), UBound(hashArray), keys
    Mouse.Hourglass False
End Sub

Private Sub qsortHashRecurse(harray() As Dictionary, inLow As Long, inHigh As Long, keys As Variant)
    Dim pivot As Dictionary, swap As Dictionary, low As Long, high As Long
    low = inLow
    high = inHigh
    Set pivot = harray((inLow + inHigh) / 2)
    
    While (low <= high)
        'If Not descending Then
            While (qsortHashCompare(harray(low), pivot, keys) = -1 And low < inHigh)
                low = low + 1
            Wend
            While (qsortHashCompare(pivot, harray(high), keys) = -1 And high > inLow)
                high = high - 1
            Wend
        'Else
        '    While (qsortHashCompare(harray(low), pivot, keys) = 1 And low < inHigh)
        '        low = low + 1
        '    Wend
        '    While (qsortHashCompare(pivot, harray(high), keys) = 1 And high > inLow)
        '        high = high - 1
        '    Wend
        'End If
        
        If (low <= high) Then
            Set swap = harray(low)
            Set harray(low) = harray(high)
            Set harray(high) = swap
            low = low + 1
            high = high - 1
        End If
    Wend
    
    If (inLow < high) Then
        qsortHashRecurse harray(), inLow, high, keys
    End If
    If (low < inHigh) Then
        qsortHashRecurse harray(), low, inHigh, keys
    End If
End Sub

Private Function qsortHashCompare(hash1 As Dictionary, hash2 As Dictionary, keys As Variant) As Long
    Dim i As Long
    Dim retval As Long
    retval = -999
    For i = 0 To UBound(keys)
        Dim thiskey As String, descending As Boolean
        thiskey = CStr(keys(i)(0))
        descending = CBool(keys(i)(1))
        
        Dim val1 As Variant, val2 As Variant
        If IsNumeric(hash1.item(thiskey)) And IsNumeric(hash2.item(thiskey)) Then
            If IsIntegeric(hash1.item(thiskey)) And IsIntegeric(hash2.item(thiskey)) Then
                val1 = CLng(hash1.item(thiskey))
                val2 = CLng(hash2.item(thiskey))
            Else
                val1 = CDbl(hash1.item(thiskey))
                val2 = CDbl(hash2.item(thiskey))
            End If
        Else
            val1 = CStr(hash1.item(thiskey))
            val2 = CStr(hash2.item(thiskey))
        End If
        
        If val1 > val2 Then
            retval = IIf(descending, -1, 1)
            Exit For
        ElseIf val1 < val2 Then
            retval = IIf(descending, 1, -1)
            Exit For
        Else
            'equal, check next key
        End If
    Next i
    
    If retval = -999 Then
        retval = 0
    End If
    qsortHashCompare = retval
End Function
