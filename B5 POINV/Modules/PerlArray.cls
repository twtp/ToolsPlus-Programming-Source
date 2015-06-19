VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerlArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------------------------
' Module    : PerlArray
' DateTime  : 7/5/2005 15:26
' Author    : briandonorfio
' Purpose   : Object that behaves similarly to a Perl array, push/pop and shift/unshift
'             functions to make it more like a stack/queue. The array always has an
'             empty cell at the end, since VB doesn't like zero length arrays, but this
'             is always removed for the user, unless it's attempting to return a zero
'             length array, in which case it returns an array with just "".
'
'             Only supports native types (long, string, etc), not objects. Some object
'             support for elem and push, but that's it.
'
'             Not implemented:
'               - splice
'               - grep
'               - map
'               - reverse
'               - sort
'---------------------------------------------------------------------------------------

Option Explicit

Private mvarDataArray() As Variant


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' DateTime  : 7/5/2005 14:34
' Author    : briandonorfio
' Purpose   : Class constructor
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    ReDim mvarDataArray(0) As Variant
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Elem
' DateTime  : 7/5/2005 14:34
' Author    : briandonorfio
' Purpose   : Returns a single element's value.
'---------------------------------------------------------------------------------------
'
Public Property Get Elem(pos As Long) As Variant
    If pos >= LBound(mvarDataArray) And pos <= UBound(mvarDataArray) Then
        If VarType(mvarDataArray(pos)) = vbObject Then 'allow objects for this and push for now
            Set Elem = mvarDataArray(pos)
        Else
            Let Elem = mvarDataArray(pos)
        End If
    End If
End Property

'---------------------------------------------------------------------------------------
' Procedure : DataArray
' DateTime  : 7/5/2005 14:33
' Author    : briandonorfio
' Purpose   : Returns the entire array as a variant array. If zero length array, returns
'             an array with a single empty cell.
'
'             THIS METHOD IS TOTALLY FUCKED WHEN COMPILING UNDER WIN7.
'---------------------------------------------------------------------------------------
'
'Public Property Get DataArray() As Variant()
'MsgBox "in dataarray()"
'MsgBox "mvar lbound is " & LBound(mvarDataArray)
'MsgBox "mvar ubound is " & UBound(mvarDataArray)
'    DataArray = mvarDataArray
'    If UBound(mvarDataArray) <> 0 Then
'        MsgBox "redim needed"
'        ReDim Preserve DataArray(UBound(mvarDataArray) - 1)
'        MsgBox "redim complete"
'    End If
'End Property

'---------------------------------------------------------------------------------------
' Procedure : Elem
' DateTime  : 7/5/2005 14:33
' Author    : briandonorfio
' Purpose   : Sets a single element's value.
'---------------------------------------------------------------------------------------
'
Public Property Let Elem(pos As Long, value As Variant)
    If pos >= LBound(mvarDataArray) Then
        If pos > UBound(mvarDataArray) Then
            ReDim Preserve mvarDataArray(pos + 1)
        End If
        mvarDataArray(pos) = value
    End If
End Property

'---------------------------------------------------------------------------------------
' Procedure : Push
' DateTime  : 7/5/2005 14:32
' Author    : briandonorfio
' Purpose   : Adds an item to the end of the array.
'---------------------------------------------------------------------------------------
'
Public Function Push(value As Variant) As Long
    ReDim Preserve mvarDataArray(UBound(mvarDataArray) + 1)
    If VarType(value) = vbObject Then 'allow objects for this and elem for now
        Set mvarDataArray(UBound(mvarDataArray) - 1) = value
    Else
        Let mvarDataArray(UBound(mvarDataArray) - 1) = value
    End If
    Push = UBound(mvarDataArray)
End Function

'---------------------------------------------------------------------------------------
' Procedure : Pop
' DateTime  : 7/5/2005 14:32
' Author    : briandonorfio
' Purpose   : Removes and returns the last item in the array.
'---------------------------------------------------------------------------------------
'
Public Function Pop() As Variant
    If UBound(mvarDataArray) - 1 < 0 Then
        Pop = Null
    Else
        Pop = mvarDataArray(UBound(mvarDataArray) - 1)
        mvarDataArray(UBound(mvarDataArray) - 1) = Empty
        ReDim Preserve mvarDataArray(UBound(mvarDataArray) - 1)
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : UnShift
' DateTime  : 7/5/2005 14:33
' Author    : briandonorfio
' Purpose   : Adds an item to the beginning of the array. Returns number of items in
'             array.
'---------------------------------------------------------------------------------------
'
Public Function UnShift(value As Variant) As Long
    ReDim Preserve mvarDataArray(UBound(mvarDataArray) + 1)
    Dim i As Long
    For i = UBound(mvarDataArray) - 1 To 1 Step -1
        mvarDataArray(i) = mvarDataArray(i - 1)
    Next i
    mvarDataArray(0) = value
    UnShift = UBound(mvarDataArray)
End Function

'---------------------------------------------------------------------------------------
' Procedure : Shift
' DateTime  : 7/5/2005 14:33
' Author    : briandonorfio
' Purpose   : Removes and returns the first item in the array.
'---------------------------------------------------------------------------------------
'
Public Function Shift() As Variant
    If UBound(mvarDataArray) - 1 < 0 Then
        Shift = Null
    Else
        Shift = mvarDataArray(LBound(mvarDataArray))
        Dim i As Long
        For i = 0 To UBound(mvarDataArray) - 1
            mvarDataArray(i) = mvarDataArray(i + 1)
        Next i
        ReDim Preserve mvarDataArray(UBound(mvarDataArray) - 1)
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : Find
' DateTime  : 7/25/2005 11:12
' Author    : briandonorfio
' Purpose   : Finds a given key in the array, returning the position if found, -1
'             otherwise.
'---------------------------------------------------------------------------------------
'
Public Function Find(key As Variant) As Long
    Dim found As Boolean
    found = False
    Dim pos As Long, i As Long
    For i = 0 To UBound(mvarDataArray) - 1
        If mvarDataArray(i) = key Then
            found = True
            pos = i
            i = UBound(mvarDataArray)
        End If
    Next i
    If Not found Then
        pos = -1
    End If
    Find = pos
End Function

Public Function FindString(key As String, Optional compareMethod As VbCompareMethod = VbCompareMethod.vbTextCompare) As Long
    Dim found As Boolean
    found = False
    Dim pos As Long, i As Long
    For i = 0 To UBound(mvarDataArray) - 1
        If StrComp(key, mvarDataArray(i), compareMethod) = 0 Then
            found = True
            pos = i
            Exit For
        End If
    Next i
    If Not found Then
        pos = -1
    End If
    FindString = pos
End Function

'---------------------------------------------------------------------------------------
' Procedure : Join
' DateTime  : 9/7/2005 12:22
' Author    : briandonorfio
' Purpose   : Returns the elements in the array as a string, separated by delim.
'---------------------------------------------------------------------------------------
'
Public Function Join(Optional delim As String = "") As String
    Dim i As Long, retval As String
    retval = ""
    For i = 0 To UBound(mvarDataArray) - 1
        retval = retval & mvarDataArray(i) & delim
    Next i
    If retval <> "" Then
        retval = Left(retval, Len(retval) - Len(delim))
    End If
    Join = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : Scalar
' DateTime  : 9/7/2005 12:25
' Author    : briandonorfio
' Purpose   : Returns the number of items in the array. Actual number (not the highest
'             index, so a for loop should go 0 to scalar-1
'---------------------------------------------------------------------------------------
'
Public Function Scalar() As Long

    Scalar = UBound(mvarDataArray)
End Function

Public Function First() As Variant
    If UBound(mvarDataArray) - 1 < 0 Then
        First = Null
    Else
        First = mvarDataArray(LBound(mvarDataArray))
    End If
End Function

Public Function Last() As Variant
    If UBound(mvarDataArray) - 1 < 0 Then
        Last = Null
    Else
        Last = mvarDataArray(UBound(mvarDataArray) - 1)
    End If
End Function

Public Function Reverse() As PerlArray
    Dim retval As PerlArray
    Set retval = New PerlArray
    Dim i As Long, j As Long
    j = 0
    For i = UBound(mvarDataArray) - 1 To 0 Step -1
        retval.Push mvarDataArray(i)
        j = j + 1
    Next i
    Set Reverse = retval
End Function
