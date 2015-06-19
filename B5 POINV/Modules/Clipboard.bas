Attribute VB_Name = "Clipboard"
'---------------------------------------------------------------------------------------
' Module    : Clipboard
' DateTime  : 6/30/2005 10:06
' Author    : briandonorfio
' Purpose   : Gets and sets the clipboard contents. From the following KB articles:
'              http://support.microsoft.com/default.aspx?scid=kb;en-us;210213
'              http://support.microsoft.com/default.aspx?scid=kb;en-us;210216
'
'             Dependencies:
'               - none
'---------------------------------------------------------------------------------------

Option Explicit

Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function CloseClipboard Lib "USER32" () As Long
Declare Function OpenClipboard Lib "USER32" (ByVal hWnd As Long) As Long
Declare Function EmptyClipboard Lib "USER32" () As Long
Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function SetClipboardData Lib "USER32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function GetClipboardData Lib "USER32" (ByVal wFormat As Long) As Long

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const CF_BITMAP = 2
Public Const MAXSIZE = 4096

Function ClipBoard_SetImage(pic As Picture)
    Dim hClipMemory As Long, X As Long

    ' Open the Clipboard to copy data to.
    If OpenClipboard(0&) = 0 Then
        MsgBox "Could not open the Clipboard. Copy aborted."
        Exit Function
    End If

    ' Clear the Clipboard.
    X = EmptyClipboard()

    ' Copy the data to the Clipboard.
    hClipMemory = SetClipboardData(CF_BITMAP, pic.handle)

OutOfHere2:

    If CloseClipboard() = 0 Then
        MsgBox "Could not close Clipboard."
    End If
End Function

Function ClipBoard_SetData(MyString As String)
    Dim hGlobalMemory As Long, lpGlobalMemory As Long
    Dim hClipMemory As Long, X As Long

    ' Allocate moveable global memory.
    '-------------------------------------------
    hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

    ' Lock the block to get a far pointer
    ' to this memory.
    lpGlobalMemory = GlobalLock(hGlobalMemory)

    ' Copy the string to this global memory.
    lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

    ' Unlock the memory.
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        MsgBox "Could not unlock memory location. Copy aborted."
        GoTo OutOfHere2
    End If

    ' Open the Clipboard to copy data to.
    If OpenClipboard(0&) = 0 Then
        MsgBox "Could not open the Clipboard. Copy aborted."
        Exit Function
    End If

    ' Clear the Clipboard.
    X = EmptyClipboard()

    ' Copy the data to the Clipboard.
    hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:

    If CloseClipboard() = 0 Then
        MsgBox "Could not close Clipboard."
    End If

End Function

Function ClipBoard_GetData() As String
On Error GoTo errh
   Dim hClipMemory As Long
   Dim lpClipMemory As Long
   Dim MyString As String
   Dim retval As Long

   If OpenClipboard(0&) = 0 Then
      MsgBox "Cannot open Clipboard. Another app. may have it open"
      Exit Function
   End If

   ' Obtain the handle to the global memory
   ' block that is referencing the text.
   hClipMemory = GetClipboardData(CF_TEXT)
   If IsNull(hClipMemory) Then
      MsgBox "Could not allocate memory"
      GoTo OutOfHere
   End If

   ' Lock Clipboard memory so we can reference
   ' the actual data string.
   lpClipMemory = GlobalLock(hClipMemory)

   If Not IsNull(lpClipMemory) Then
      MyString = Space$(MAXSIZE)
      retval = lstrcpy(MyString, lpClipMemory)
      retval = GlobalUnlock(hClipMemory)

      ' Peel off the null terminating character.
      MyString = Mid(MyString, 1, InStr(1, MyString, Chr$(0), 0) - 1)
   Else
      MsgBox "Could not lock memory to copy string from."
   End If

OutOfHere:

   retval = CloseClipboard()
   ClipBoard_GetData = MyString
   Exit Function
   
errh:
    MsgBox "Error " & Err.Number & " accessing clipboard!"
    ClipBoard_GetData = ""
End Function

