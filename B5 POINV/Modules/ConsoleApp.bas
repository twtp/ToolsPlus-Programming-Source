Attribute VB_Name = "ConsoleApp"
'---------------------------------------------------------------------------------------
' Module    : ConsoleApp
' DateTime  : 5/5/2006 09:44
' Author    : briandonorfio
' Purpose   : Console input/output for VB6. Sub Main should have ConsoleAlloc and
'             ConsoleFree as the first and last things.
'
'             To actually create the exe file correctly, compile it, and then copy these
'             files to the exe directory:
'                - vc98/bin/editbin.exe
'                - vc98/bin/link.exe
'                - vb98/mspdb60.dll
'             And do the following at a command prompt: editbin /subsystem:console <exe>
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, lpBuffer As Any, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long

Private Const STD_ERROR_HANDLE = -12&
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&

Private hOutput As Long
Private hInput As Long
Private hError As Long

Function ConsoleAlloc() As Boolean
    AllocConsole

    hOutput = GetStdHandle(STD_OUTPUT_HANDLE)
    hInput = GetStdHandle(STD_INPUT_HANDLE)
    hError = GetStdHandle(STD_ERROR_HANDLE)

    ConsoleAlloc = ((hOutput <> 0) And (hInput <> 0) And (hError <> 0))
End Function

Function ConsoleFree()
    Dim lResult As Long

    lResult = FreeConsole()

    ConsoleFree = (lResult <> 0)
End Function

Function ConsoleWrite(sOutput As String) As Boolean
    Dim lOutput As Long
    Dim lWritten As Long
    Dim lResult As Long

    lOutput = Len(sOutput)
    lResult = WriteConsole(hOutput, ByVal sOutput, lOutput, lWritten, 0)

    ConsoleWrite = (lResult <> 0)
End Function

Function ConsoleWriteLn(sOutput As String) As Boolean
    ConsoleWriteLn = ConsoleWrite(sOutput & vbCrLf)
End Function

Function ConsoleError(sOutput As String) As Boolean
    Dim lOutput As Long
    Dim lWritten As Long
    Dim lResult As Long

    lOutput = Len(sOutput)
    lResult = WriteConsole(hError, ByVal sOutput, lOutput, lWritten, 0)

    ConsoleError = (lResult <> 0)
End Function

Function ConsoleErrorLn(sOutput As String) As Boolean
    ConsoleErrorLn = ConsoleError(sOutput & vbCrLf)
End Function

Function ConsoleRead(sInput As String) As Boolean
    Dim lInput As Long
    Dim lResult As Long
    Dim lRead As Long

    sInput = Space(255)
    lInput = Len(sInput)

    lResult = ReadConsole(hInput, ByVal sInput, lInput, lRead, 0)

    If lRead > 1 Then sInput = Left(sInput, lRead - 2)

    ConsoleRead = (lResult <> 0)
End Function
