VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FilePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : FilePicker
' DateTime  : 5/25/2006 13:29
' Author    : briandonorfio
' Purpose   : Abstracts away the API functions for calling an open file dialog box.
'             Only the basic functions, though, but you can use it without a form, not
'             like comdlg32.ocx
'---------------------------------------------------------------------------------------

Option Explicit

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpszFilter As String
    lpszCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpszFile As String
    nMaxFile As Long
    lpszFileTitle As String
    nMaxFileTitle As Long
    lpszInitialDir As String
    lpszTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpszDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (lpOFN As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (lpOFN As OPENFILENAME) As Long

Private info As OPENFILENAME

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' DateTime  : 5/25/2006 13:31
' Author    : briandonorfio
' Purpose   : CTOR, sets a few defaults
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    info.lStructSize = Len(info)
    If Forms.count = 0 Then
        info.hwndOwner = 0
    Else
        info.hwndOwner = Forms(0).hwnd
    End If
    info.hInstance = app.hInstance
    info.lpszFilter = ""
    info.nFilterIndex = 1
    info.lpszFile = String(255, 0)
    info.nMaxFile = Len(info.lpszFile) - 1
    info.lpszFileTitle = info.lpszFile
    info.nMaxFileTitle = info.nMaxFile
    'info.lpszInitialDir = Environ("HomeShare")
    info.lpszInitialDir = Environ("HomeDrive") & Environ("HomePath") & "\Desktop"
    info.lpszTitle = ""
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetParent
' DateTime  : 5/25/2006 13:31
' Author    : briandonorfio
' Purpose   : Sets the parent window of the dialog, given that window's hWnd.
'---------------------------------------------------------------------------------------
'
Public Sub SetParent(parentHwnd As Long)
    info.hwndOwner = parentHwnd
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetFilters
' DateTime  : 5/25/2006 13:32
' Author    : briandonorfio
' Purpose   : Returns a dictionary of the current filters for the dialog box. Visible
'             string is the hash key, glob is the value.
'---------------------------------------------------------------------------------------
'
Public Function GetFilters() As Dictionary
    Dim pieces As Variant
    pieces = Split(info.lpszFilter, Chr(0))
    Dim retval As Dictionary
    Set retval = New Dictionary
    Dim i As Long
    For i = LBound(pieces) To UBound(pieces) Step 2
        retval.Add CStr(pieces(i)), CStr(pieces(i + 1))
    Next i
    retval.Add "All Files", "*.*"
    Set GetFilters = retval
    Set retval = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : AddFilter
' DateTime  : 5/25/2006 13:32
' Author    : briandonorfio
' Purpose   : Given a visible string to show, and a fileglob, this will add a filter to
'             the dialog box.
'---------------------------------------------------------------------------------------
'
Public Sub AddFilter(textStr As String, glob As String)
    info.lpszFilter = info.lpszFilter & textStr & Chr(0) & glob & Chr(0)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InitialDir
' DateTime  : 5/25/2006 13:33
' Author    : briandonorfio
' Purpose   : Gets the directory to open to. Defaults to users homedir.
'---------------------------------------------------------------------------------------
'
Public Property Get InitialDir() As String
    InitialDir = info.lpszInitialDir
End Property

'---------------------------------------------------------------------------------------
' Procedure : InitialDir
' DateTime  : 5/25/2006 13:33
' Author    : briandonorfio
' Purpose   : Sets the directory to open to.
'---------------------------------------------------------------------------------------
'
Public Property Let InitialDir(newDir As String)
    info.lpszInitialDir = newDir
End Property

'---------------------------------------------------------------------------------------
' Procedure : Title
' DateTime  : 5/25/2006 13:33
' Author    : briandonorfio
' Purpose   : Gets the title of the dialog box
'---------------------------------------------------------------------------------------
'
Public Property Get Title() As String
    Title = info.lpszTitle
End Property

'---------------------------------------------------------------------------------------
' Procedure : Title
' DateTime  : 5/25/2006 13:33
' Author    : briandonorfio
' Purpose   : Sets the title of the dialog box
'---------------------------------------------------------------------------------------
'
Public Property Let Title(newTitle As String)
    info.lpszTitle = newTitle
End Property

'---------------------------------------------------------------------------------------
' Procedure : ShowDialog
' DateTime  : 5/25/2006 13:34
' Author    : briandonorfio
' Purpose   : Opens the dialog box, and returns the filename, or a null string if the
'             user cancels it.
'---------------------------------------------------------------------------------------
'
Public Function ShowDialogOpen() As String
    AddFilter "All Files", "*.*"
    If info.lpszTitle = "" Then
        info.lpszTitle = "Open File..."
    End If
    Dim res As Long
    res = GetOpenFileName(info)
    If res Then
        ShowDialogOpen = Left(info.lpszFile, InStr(info.lpszFile, Chr(0)) - 1)
    Else
        ShowDialogOpen = ""
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : ShowDialogSave
' DateTime  : 9/25/2008 10:02
' Author    : briandonorfio
' Purpose   : Opens a save dialog, returns the filename or a null string.
'---------------------------------------------------------------------------------------
'
Public Function ShowDialogSave() As String
    AddFilter "All Files", "*.*"
    If info.lpszTitle = "" Then
        info.lpszTitle = "Save As..."
    End If
    Dim res As Long
    res = GetSaveFileName(info)
    If res Then
        ShowDialogSave = Left(info.lpszFile, InStr(info.lpszFile, Chr(0)) - 1)
    Else
        ShowDialogSave = ""
    End If
End Function
