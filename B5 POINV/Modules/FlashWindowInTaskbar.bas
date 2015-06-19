Attribute VB_Name = "FlashWindowInTaskbar"
Option Explicit

Private Type FLASHWINFO
  cbSize As Long
  hWnd As Long
  dwFlags As Long
  uCount As Long
  dwTimeout As Long
End Type

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function FlashWindowEx Lib "user32" (FWInfo As FLASHWINFO) As Boolean

Private Const FLASHW_TRAY = 2

Public Sub FlashWindow(hWnd As Long, Optional NumberOfFlashes As Long = 5)
    If Not APIFunctionPresent("FlashWindowEx", "user32") Then
        Exit Sub
    End If
    Dim retval As Boolean
    Dim udtFWInfo As FLASHWINFO
    With udtFWInfo
       .cbSize = 20
       .hWnd = hWnd
       .dwFlags = FLASHW_TRAY
       .uCount = NumberOfFlashes
       .dwTimeout = 0
    End With
    retval = FlashWindowEx(udtFWInfo)
End Sub

Private Function APIFunctionPresent(ByVal FunctionName As String, ByVal DllName As String) As Boolean
    Dim handle As Long, addr  As Long
    handle = LoadLibrary(DllName)
    If handle <> 0 Then
        addr = GetProcAddress(handle, FunctionName)
        FreeLibrary handle
    End If
    APIFunctionPresent = (addr <> 0)
End Function

