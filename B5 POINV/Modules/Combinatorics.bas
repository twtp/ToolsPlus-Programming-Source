Attribute VB_Name = "Combinatorics"
'---------------------------------------------------------------------------------------
' Module    : Combinatorics
' DateTime  : 6/1/2006 10:16
' Author    : briandonorfio
' Purpose   : Ack! Discrete math! Provides some funky math stuff we'll probably never
'             use outside of keywords.
'---------------------------------------------------------------------------------------

Option Explicit

Dim currentSet() As Boolean

'---------------------------------------------------------------------------------------
' Procedure : Combine
' DateTime  : 6/1/2006 10:17
' Author    : briandonorfio
' Purpose   : Returns a variant array of all possible count-size combinations of the
'             data array. Return value is another variant array of strings, each element
'             selected from the original data array is separated by a space.
'
'             Array("A","B","C","D") w/ count=3 returns Array("A B C", "A B D", "B C D")
'---------------------------------------------------------------------------------------
'
Public Function Combine(data As Variant, count As Long) As Variant
    If Not VarType(data) And vbArray Then
        Err.Raise ErrorCodes.CombinatoricsArgumentTypeMismatch, "Combinatorics", "data must be a variant array!"
        Exit Function
    End If
    If count > UBound(data) + 1 Then
        Err.Raise ErrorCodes.CombinatoricsInvalidArguments, "Combinatorics", "count > data length!"
        Exit Function
    End If


    ReDim retval(findBinomialCoefficient(UBound(data) + 1, count) - 1) As Variant
    Dim retCount As Long
    
    ReDim currentSet(UBound(data)) As Boolean
    retval(retCount) = getFirstCombination(data, count)
    retCount = retCount + 1
    While setupNextCombination
        retval(retCount) = translateToString(data)
        retCount = retCount + 1
    Wend
    Combine = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : translateToString
' DateTime  : 6/1/2006 10:17
' Author    : briandonorfio
' Purpose   : Internal function, translates the currentSet global flags to the word
'             string from the data array.
'---------------------------------------------------------------------------------------
'
Private Function translateToString(data As Variant) As String
    Dim i As Long, retval As String
    For i = 0 To UBound(currentSet)
        If currentSet(i) Then
            retval = retval & data(i) & " "
        End If
    Next i
    translateToString = Left(retval, Len(retval) - 1)
End Function

'---------------------------------------------------------------------------------------
' Procedure : getFirstCombination
' DateTime  : 6/1/2006 10:17
' Author    : briandonorfio
' Purpose   : Internal function, primes the currentSet global and returns the first
'             possible combination as an already translated string.
'---------------------------------------------------------------------------------------
'
Private Function getFirstCombination(data As Variant, count As Long) As String
    Dim i As Long
    For i = 0 To count - 1
        currentSet(i) = True
    Next i
    getFirstCombination = translateToString(data)
End Function

'---------------------------------------------------------------------------------------
' Procedure : setupNextCombination
' DateTime  : 6/1/2006 10:17
' Author    : briandonorfio
' Purpose   : Internal function, checks to see if there is another possible combination
'             for the currentSet global, if there is, then it sets it up and returns
'             true, otherwise false.
'---------------------------------------------------------------------------------------
'
Private Function setupNextCombination() As Boolean
    'find rightmost 1 w/ 0 to the right
    Dim i As Long, pos As Long
    pos = -1
    For i = UBound(currentSet) - 1 To 0 Step -1
        If currentSet(i) And Not currentSet(i + 1) Then
            pos = i
            i = 0
        End If
    Next i
    'if we can't find one, we're over to the right, so we're done
    If pos = -1 Then
        setupNextCombination = False
    'otherwise, still have work to do
    Else
        'find number of words 'on' to the right of pos
        Dim rightCount As Long
        For i = pos + 1 To UBound(currentSet)
            If currentSet(i) Then
                rightCount = rightCount + 1
            End If
        Next i
        'swap
        currentSet(pos) = False
        currentSet(pos + 1) = True
        'move right side
        For i = 1 To rightCount
            currentSet(pos + i + 1) = True
        Next i
        For i = pos + 2 + rightCount To UBound(currentSet)
            currentSet(i) = False
        Next i
        'and return true
        setupNextCombination = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : findBinomialCoefficient
' DateTime  : 6/1/2006 10:17
' Author    : briandonorfio
' Purpose   : Internal function, returns the number of possible outcomes for a given
'             combine() operation.
'
'             Given by the equation:
'                     n!
'               -------------
'                k! * (n-k)!
'---------------------------------------------------------------------------------------
'
Public Function findBinomialCoefficient(n As Long, k As Long) As Double
    'Dim numerator As Long, denominator As Long
'    Dim numerator As Double, denominator As Double
'    If n = k Then
'        findBinomialCoefficient = 1
'    Else
'        numerator = factorial(n)
'        denominator = factorial(k) * factorial(n - k)
'        findBinomialCoefficient = numerator / denominator
'    End If

    'faster algorithm! doesn't overflow (as easily)!
    If k = 0 Then
        findBinomialCoefficient = 1
    Else
        Dim i As Long, temp As Double
        temp = 1
        For i = 1 To k
            temp = temp * (n + 1 - i) / i
        Next i
        findBinomialCoefficient = temp
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : factorial
' DateTime  : 6/1/2006 10:17
' Author    : briandonorfio
' Purpose   : Internal function, returns the factorial of n.
'---------------------------------------------------------------------------------------
'
'Private Function factorial(n As Long) As Double
'    Dim i As Long, retval As Double
'    retval = n
'    For i = n - 1 To 1 Step -1
'        retval = retval * i
'    Next i
'    factorial = retval
'End Function

