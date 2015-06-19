VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form KeywordsClearOvertureFlags 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clear Overture Upload Flags"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnClear 
      Caption         =   "When done, click this!"
      Height          =   585
      Left            =   8430
      TabIndex        =   4
      Top             =   6030
      Width           =   2235
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit without clearing export flags"
      Height          =   525
      Left            =   6060
      TabIndex        =   2
      Top             =   6030
      Width           =   1755
   End
   Begin VB.CommandButton btnGo 
      Caption         =   "Log in to overture and click this button!"
      Height          =   555
      Left            =   3660
      TabIndex        =   1
      Top             =   6030
      Width           =   2325
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5895
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   10605
      ExtentX         =   18706
      ExtentY         =   10398
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
   Begin VB.Label lblStatusBar 
      Height          =   285
      Left            =   30
      TabIndex        =   3
      Top             =   6330
      Width           =   3195
   End
End
Attribute VB_Name = "KeywordsClearOvertureFlags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClear_Click()
    If vbYes = MsgBox("Clear flags", vbYesNo) Then
        Shell PERL & " " & KWD_O_CLEAR_FLAGS, vbNormalFocus
    End If
End Sub

Private Sub btnExit_Click()
    'KeywordsMenu.SetDownloadedOK False
    Unload KeywordsClearOvertureFlags
End Sub

Private Sub btnGo_Click()
    Mouse.Hourglass True
    Dim files As Variant, urls As Variant, i As Long
    'these file names are used in KeywordsMenu and the perl script, so don't change them
    files = Array("listings.tsv", _
                  "rejected.tsv", _
                  "pending.tsv")
    urls = Array("https://secure.overture.com/bdd?format=csv&category=all+listings&tagId=-1&listingType=active", _
                 "https://secure.overture.com/bdd?format=csv&listingType=declined", _
                 "https://secure.overture.com/bdd?format=csv&listingType=pending")
    Dim retryCount As Long
    For i = LBound(files) To UBound(files)
        Me.lblStatusBar.Caption = "downloading " & files(i)
        retryCount = 0
retry:
        DoEvents
        If Not Download(CStr(urls(i)), "s:\mastest\mas90-signs\keywords_scripts\overture_reports\" & files(i)) Then
            If retryCount = 3 Then
                MsgBox "Error downloading " & files(i) & ", aborting..."
                'KeywordsMenu.SetDownloadedOK False
                Me.lblStatusBar.Caption = ""
                Mouse.Hourglass False
                Exit Sub
            Else
                retryCount = retryCount + 1
                GoTo retry
            End If
        End If
    Next i
    'KeywordsMenu.SetDownloadedOK True
    Me.lblStatusBar.Caption = ""
    Mouse.Hourglass False
    'Unload KeywordsClearOvertureFlags
    MsgBox "Done downloading, click the other button!"
End Sub

Private Sub Form_Load()
    Me.WebBrowser1.Navigate2 "https://secure.overture.com/login.do?mkt=us&locale=en_US"
End Sub
