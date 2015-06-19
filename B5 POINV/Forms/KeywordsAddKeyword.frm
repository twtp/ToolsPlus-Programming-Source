VERSION 5.00
Begin VB.Form KeywordsAddKeyword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keywords - Add Keyword"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4500
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox source 
      Height          =   285
      Left            =   3120
      TabIndex        =   6
      Text            =   "source"
      Top             =   780
      Width           =   1275
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Cancel"
      Height          =   465
      Left            =   2370
      TabIndex        =   5
      Top             =   1020
      Width           =   1545
   End
   Begin VB.CommandButton btnAddKeyword 
      Caption         =   "Add Keyword Now"
      Height          =   465
      Left            =   660
      TabIndex        =   4
      Top             =   1020
      Width           =   1545
   End
   Begin VB.CommandButton btnLandingPage 
      Caption         =   "Landing Page"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   1155
   End
   Begin VB.TextBox LandingPage 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   450
      Width           =   3195
   End
   Begin VB.TextBox SearchPhrase 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3195
   End
   Begin VB.Label Label1 
      Caption         =   "Keyword:"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "KeywordsAddKeyword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type IEInstance
    hwnd As Long
    url As String
End Type

Private Const MAX_LEN = 255

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassname As String, ByVal nMaxCount As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private IEList() As IEInstance

Private Sub btnAddKeyword_Click()
    DB.Execute "INSERT INTO KeywordsPhrases ( SearchPhrase, " & IIf(Me.LandingPage <> "", "LandingPage, ", "") & "Source ) VALUES ( '" & EscapeSQuotes(TrimWhitespace(Me.SearchPhrase, True, True, True)) & "', " & IIf(Me.LandingPage <> "", "'" & Me.LandingPage & "', ", "") & "'ListTool' )"
    DB.Execute "INSERT INTO KeywordsAttributes ( SearchPhrase, Service ) VALUES ( '" & EscapeSQuotes(Me.SearchPhrase) & "', 'Overture' )"
    DB.Execute "INSERT INTO KeywordsAttributes ( SearchPhrase, Service ) VALUES ( '" & EscapeSQuotes(Me.SearchPhrase) & "', 'AdWords' )"
    Unload KeywordsAddKeyword
End Sub

Private Sub btnExit_Click()
    Unload KeywordsAddKeyword
End Sub

Private Sub btnLandingPage_Click()
    ReDim IEList(0) As IEInstance
    enumerateIEWindows
    If listIsEmpty Then
        Shell "c:\program files\internet explorer\iexplore.exe http://www.tools-plus.com/", vbNormalFocus
    Else
        Dim i As Long, found As Boolean
        While Not found And i <= UBound(IEList)
            If IEList(i).url = "http://www.tools-plus.com/" Or IEList(i).url = "http://www.toolsplus.com/" Then
                found = True
            ElseIf InStr(IEList(i).url, "http://www.tools-plus.com/") Then
                Me.LandingPage = Replace(IEList(i).url, "http://www.tools-plus.com/", "")
                found = True
            ElseIf InStr(IEList(i).url, "http://www.toolsplus.com/") Then
                Me.LandingPage = Replace(IEList(i).url, "http://www.toolsplus.com/", "")
                found = True
            End If
            i = i + 1
        Wend
        If Not found Then
            Shell "c:\program files\internet explorer\iexplore.exe http://www.tools-plus.com/", vbNormalFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.source = "Unknown"
End Sub

Private Sub SearchPhrase_Change()
    If Me.SearchPhrase = "" Then
        Me.btnAddKeyword.Enabled = False
    Else
        Me.btnAddKeyword.Enabled = True
    End If
End Sub

Private Sub enumerateIEWindows()
    Dim hwnd As Long, windowClass As String, childhWnd As Long
    Dim urlLen As Long, url As String
    Dim thisIE As IEInstance
    hwnd = GetWindow(GetDesktopWindow(), GW_CHILD)
    While hwnd <> 0
        If Not alreadyInList(hwnd) Then
            windowClass = getWindowClassName(hwnd)
            If windowClass = "CabinetWClass" Or windowClass = "IEFrame" Then
                'window -> worker -> rebar -> comboboxex -> combobox -> edit
                childhWnd = FindWindowEx(hwnd, 0, "WorkerW", vbNullString)
                If childhWnd > 0 Then
                    childhWnd = FindWindowEx(childhWnd, 0, "ReBarWindow32", vbNullString)
                End If
                If childhWnd > 0 Then
                    childhWnd = FindWindowEx(childhWnd, 0, "ComboBoxEx32", vbNullString)
                End If
                If childhWnd > 0 Then
                    childhWnd = FindWindowEx(childhWnd, 0, "ComboBox", vbNullString)
                End If
                If childhWnd > 0 Then
                    childhWnd = FindWindowEx(childhWnd, 0, "Edit", vbNullString)
                End If
                If childhWnd > 0 Then
                    urlLen = SendMessage(childhWnd, WM_GETTEXTLENGTH, 0, ByVal 0&)
                    url = Space(urlLen + 1)
                    urlLen = SendMessage(childhWnd, WM_GETTEXT, urlLen + 1, ByVal url)
                    thisIE.hwnd = hwnd
                    thisIE.url = Left(url, Len(url) - 1)
                    If IEList(UBound(IEList)).url <> "" Then
                        ReDim Preserve IEList(UBound(IEList) + 1) As IEInstance
                    End If
                    IEList(UBound(IEList)) = thisIE
                End If
            End If
        End If
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Wend
End Sub

Private Function getWindowClassName(hwnd As Long) As String
    Dim buf As String
    Dim i As Long
    buf = String(MAX_LEN + 1, 0)
    i = GetClassName(hwnd, buf, MAX_LEN)
    If i > 0 Then
        getWindowClassName = Left(buf, i)
    Else
        getWindowClassName = ""
    End If
End Function

Private Function alreadyInList(hwnd As Long) As Boolean
    Dim i As Long
    For i = 0 To UBound(IEList)
        If IEList(i).hwnd = hwnd Then
            alreadyInList = True
            Exit Function
        End If
    Next i
    alreadyInList = False
End Function

Private Function listIsEmpty() As Boolean
    If UBound(IEList) = 0 And IEList(0).url = "" Then
        listIsEmpty = True
    Else
        listIsEmpty = False
    End If
End Function
