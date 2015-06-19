VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FolderPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : FolderPicker
' DateTime  : 5/25/2006 10:53
' Author    : briandonorfio
' Purpose   : Abstracts away the API functions to provide a modal selection dialog for
'             folders.
'---------------------------------------------------------------------------------------

Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_BROWSEINCLUDEFILES = 16384
Private Const MAX_PATH = 512

Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private info As BrowseInfo

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' DateTime  : 5/25/2006 10:54
' Author    : briandonorfio
' Purpose   : CTOR, sets a few defaults so it'll at least run with a .GetFolder if they
'             don't specify anything else.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    If Forms.count = 0 Then
        info.hwndOwner = 0
    Else
        info.hwndOwner = Forms(0).hwnd
    End If
    info.ulFlags = BIF_RETURNONLYFSDIRS
    info.lpszTitle = lstrcat("Browse for directory...", "")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetParent
' DateTime  : 5/25/2006 10:54
' Author    : briandonorfio
' Purpose   : Sets the parent window for the dialog, given that window's .hWnd, no idea
'             why it needs this. If not called, this gets set to the first form in the
'             Forms collection, or if no forms are loaded, 0.
'---------------------------------------------------------------------------------------
'
Public Sub SetParent(parentHwnd As Long)
    info.hwndOwner = parentHwnd
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetTitle
' DateTime  : 5/25/2006 10:55
' Author    : briandonorfio
' Purpose   : Sets the title for the dialog's title bar. If not called, this defaults to
'             "Browse for directory..."
'---------------------------------------------------------------------------------------
'
Public Sub SetTitle(newTitle As String)
    info.lpszTitle = lstrcat(newTitle, "")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ShowFilesToo
' DateTime  : 5/25/2006 10:55
' Author    : BrianDonorfio
' Purpose   : Shows a file/folder picker.
'---------------------------------------------------------------------------------------
'
Public Sub ShowFilesToo(yesno As Boolean)
    If (info.ulFlags And BIF_BROWSEINCLUDEFILES) = BIF_BROWSEINCLUDEFILES Then 'i hope that's a bitwise and
        info.ulFlags = info.ulFlags Or BIF_BROWSEINCLUDEFILES
    Else
        info.ulFlags = info.ulFlags Xor BIF_BROWSEINCLUDEFILES
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetFolder
' DateTime  : 5/25/2006 10:56
' Author    : briandonorfio
' Purpose   : Shows the folder selection dialog. Returns the full path to the folder,
'             or a null string if cancelled.
'---------------------------------------------------------------------------------------
'
Public Function GetFolder() As String
    Dim idlist As Long, buf As String
    idlist = SHBrowseForFolder(info)
    If idlist Then
        buf = Space(MAX_PATH)
        SHGetPathFromIDList idlist, buf
        GetFolder = Left(buf, InStr(buf, vbNullChar) - 1)
    Else
        GetFolder = ""
    End If
End Function
