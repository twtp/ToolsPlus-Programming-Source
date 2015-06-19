VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form KeywordsBidTool 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keywords - Get Bids"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   10230
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnParse 
      Caption         =   "Run Parse Script Now"
      Height          =   495
      Left            =   6090
      TabIndex        =   5
      Top             =   5880
      Width           =   1485
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "STOP! (may take a few seconds)"
      Height          =   495
      Left            =   4230
      TabIndex        =   4
      Top             =   5880
      Width           =   1515
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   525
      Left            =   8130
      TabIndex        =   3
      Top             =   5850
      Width           =   1965
   End
   Begin VB.CommandButton btnAutomate 
      Caption         =   "Go!"
      Height          =   525
      Left            =   2730
      TabIndex        =   2
      Top             =   5850
      Width           =   1305
   End
   Begin VB.ListBox phraseList 
      Height          =   6300
      Left            =   30
      TabIndex        =   1
      Top             =   60
      Width           =   2055
   End
   Begin SHDocVwCtl.WebBrowser browser 
      Height          =   5655
      Left            =   2280
      TabIndex        =   0
      Top             =   60
      Width           =   7875
      ExtentX         =   13891
      ExtentY         =   9975
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "KeywordsBidTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BID_PAGE_URL As String = "http://uv.bidtool.overture.com/d/search/tools/bidtool/index.jhtml?mkt=us&lang=en_US&Partner=userbidtool&Keywords="
Private Const SAVE_DIR As String = "s:\mastest\mas90-signs\keywords_scripts\bid_pages\"

Private complete As Boolean
Private exitnow As Boolean

Private Sub browser_DocumentComplete(ByVal pDisp As Object, url As Variant)
    complete = True
End Sub

Private Sub btnAutomate_Click()
    Dim i As Long
    i = Me.PhraseList.ListIndex
    Me.PhraseList.ListIndex = -1
    For i = i To Me.PhraseList.ListCount - 1
        Me.PhraseList.ListIndex = i
        'phraseList_Click 'apparently gets called automatically when you go to it
        If exitnow Then
            Exit Sub
        End If
    Next i
    ShellWait PERL & " " & KWD_BID_PARSE, vbNormalFocus
'On Error Resume Next
    'Kill SAVE_DIR & "*"
End Sub

Private Sub btnExit_Click()
    Unload KeywordsBidTool
End Sub

Private Sub btnParse_Click()
    ShellWait PERL & " " & KWD_BID_PARSE, vbNormalFocus
End Sub

Private Sub btnStop_Click()
    exitnow = True
End Sub

Private Sub Form_Load()
    requeryPhraseList
    Me.browser.Navigate2 Replace(BID_PAGE_URL, "&Keywords=", "")
End Sub

Private Sub requeryPhraseList()
    ReDim phrList(KeywordsPhraseList.GetListLength) As String
    phrList = KeywordsPhraseList.GetList
    Me.PhraseList.Clear
    Dim i As Long
    For i = 0 To UBound(phrList)
        Me.PhraseList.AddItem phrList(i)
    Next i
End Sub

Private Sub phraseList_Click()
    exitnow = False
    complete = False
    Dim foo As Long
    If Me.PhraseList <> "" Then
        Mouse.Hourglass True
        Me.browser.Navigate2 BID_PAGE_URL & Replace(Me.PhraseList, " ", "+")
        waitForPage
        If InStr(Me.browser.Document.body.innerText, "unauthorized automated scripts") Then
            MsgBox "Enter code, then hit go again"
            exitnow = True
        ElseIf InStr(Me.browser.Document.body.innerText, "error has occurred") Then
            foo = Me.PhraseList.ListIndex
            Me.PhraseList.ListIndex = -1
            Me.PhraseList.ListIndex = foo 'possible stack overflow? hopefully not too many errors in a row.
        Else
            SaveHTML Me.browser.Document.body.innerHTML
        End If
        Mouse.Hourglass False
    End If
End Sub

Private Sub SaveHTML(html As String)
    Dim fname As String
    fname = Me.PhraseList
    fname = Replace(fname, "/", "_fs_")     'forward slashes are bad in filenames
    fname = Replace(fname, """", "_dq_")    'so are double quotes
    Open SAVE_DIR & fname For Output As #1
    Print #1, html
    Close #1
End Sub

Private Function waitForPage(Optional timeout As Long = 60) As Boolean
    Dim starttime As Double
    starttime = Timer
    While Not complete
        DoEvents
        If Timer - starttime > timeout Then
            Me.browser.Stop
            complete = False
            waitForPage = False
            Exit Function
        End If
    Wend
    complete = False
    waitForPage = True
End Function
