Attribute VB_Name = "PointerFunctions"
Option Explicit

Private allPointers() As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)

Public Sub AddPointer(ByVal thePointer As Long)
    Call findPointer(thePointer, True)
End Sub

Public Sub RemovePointer(ByVal thePointer As Long)
    Dim item As Long, max As Long
    
    item = IsValidPointer(thePointer)
    If item > -1 Then
        max = UBound(allPointers)
        If item < max Then
            CopyMemory ByVal VarPtr(allPointers(item)), ByVal VarPtr(allPointers(item + 1)), (max - item) * PTR_SIZE
        End If
        If max Then
            ReDim Preserve allPointers(max - 1) As Long
        Else
            ReDim allPointers(0) As Long
            allPointers(0) = 0
        End If
    End If
End Sub

Public Function IsValidPointer(ByVal thePointer As Long) As Long
    IsValidPointer = findPointer(thePointer, False)
End Function

Private Function findPointer(ByVal thePointer As Long, ByVal addIfNotFound As Boolean) As Long
    'bsearch, maybe add
    Static inited As Boolean
    Dim i As Long, low As Long, high As Long, max As Long, retval As Long
    
    If Not inited Then
        If addIfNotFound Then
            retval = realAddPointer(thePointer, 0, inited)
            inited = True
        Else
            retval = -1
        End If
    Else
        high = UBound(allPointers)
    
        Do
            i = low + ((high - low) \ 2)
            
            Select Case allPointers(i)
                Case Is = thePointer
                    retval = i
                    Exit Do
                Case Is > thePointer
                    high = i - 1
                Case Is < thePointer
                    low = i + 1
            End Select
            
            If low > high Then
                retval = -1
                Exit Do
            End If
        Loop
    End If
    
    If retval = -1 And addIfNotFound Then
        retval = realAddPointer(thePointer, i, inited)
        inited = True
    End If
    
    findPointer = retval
End Function

Private Function realAddPointer(ByVal thePointer As Long, ByVal pos As Long, ByVal inited As Boolean) As Long
    If Not inited Then
        ReDim allPointers(0) As Long
    Else
        If allPointers(0) <> 0 Then
            ReDim Preserve allPointers(UBound(allPointers) + 1) As Long
        End If
        Dim max As Long
        max = UBound(allPointers)
        If allPointers(pos) < thePointer Then
            pos = pos + 1
        End If
        If pos > max Then
            pos = max
        ElseIf pos < max Then
            CopyMemory ByVal VarPtr(allPointers(pos + 1)), ByVal VarPtr(allPointers(pos)), (max - pos) * PTR_SIZE
        End If
    End If
    allPointers(pos) = thePointer
    realAddPointer = pos
End Function
