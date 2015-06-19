VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form KeywordsImportTool 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keywords - Import Data"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11250
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkRange 
      Caption         =   "90"
      Height          =   195
      Index           =   90
      Left            =   690
      TabIndex        =   14
      Top             =   6810
      Value           =   1  'Checked
      Width           =   555
   End
   Begin VB.CheckBox chkRange 
      Caption         =   "60"
      Height          =   195
      Index           =   60
      Left            =   690
      TabIndex        =   13
      Top             =   6600
      Value           =   1  'Checked
      Width           =   555
   End
   Begin VB.CheckBox chkRange 
      Caption         =   "30"
      Height          =   195
      Index           =   30
      Left            =   90
      TabIndex        =   12
      Top             =   6840
      Value           =   1  'Checked
      Width           =   555
   End
   Begin VB.CheckBox chkRange 
      Caption         =   "10"
      Height          =   195
      Index           =   10
      Left            =   90
      TabIndex        =   11
      Top             =   6600
      Value           =   1  'Checked
      Width           =   555
   End
   Begin VB.OptionButton opPage 
      Caption         =   "AdWords"
      Height          =   225
      Index           =   2
      Left            =   6120
      TabIndex        =   10
      Top             =   6390
      Width           =   1185
   End
   Begin VB.CommandButton btnRunNow 
      Caption         =   "Run Now"
      Height          =   225
      Left            =   7770
      TabIndex        =   9
      Top             =   6180
      Width           =   1215
   End
   Begin VB.CheckBox chkRunPerl 
      Caption         =   "run scripts when done"
      Height          =   255
      Left            =   7440
      TabIndex        =   8
      Top             =   5910
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin VB.TextBox loadingtimer 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Height          =   285
      Left            =   10170
      TabIndex        =   7
      Top             =   5970
      Width           =   885
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   525
      Left            =   9450
      TabIndex        =   6
      Top             =   6390
      Width           =   1635
   End
   Begin VB.CommandButton btnClearStats 
      Caption         =   "Clear Old Data"
      Height          =   525
      Left            =   240
      TabIndex        =   5
      Top             =   5940
      Width           =   2025
   End
   Begin VB.OptionButton opPage 
      Caption         =   "Overture"
      Height          =   225
      Index           =   1
      Left            =   6120
      TabIndex        =   3
      Top             =   6150
      Width           =   1185
   End
   Begin VB.OptionButton opPage 
      Caption         =   "Yahoo"
      Height          =   225
      Index           =   0
      Left            =   6120
      TabIndex        =   2
      Top             =   5910
      Value           =   -1  'True
      Width           =   1185
   End
   Begin VB.CommandButton btnGetStats 
      Caption         =   "Import Stats From Yahoo/Overture!"
      Height          =   555
      Left            =   2820
      TabIndex        =   1
      Top             =   5910
      Width           =   2925
   End
   Begin SHDocVwCtl.WebBrowser browser 
      Height          =   5805
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   11145
      ExtentX         =   19659
      ExtentY         =   10239
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
      Location        =   ""
   End
   Begin VB.Label lblInstructions2 
      Caption         =   $"KeywordsImportTool.frx":0000
      Height          =   435
      Left            =   1530
      TabIndex        =   4
      Top             =   6660
      Width           =   7845
   End
End
Attribute VB_Name = "KeywordsImportTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PAGE_DIR = "s:\mastest\mas90-signs\keywords_scripts\yahoo_overture_pages"

Private Const YAHOO As Long = 0
Private Const OVERTURE As Long = 1
Private Const ADWORDS As Long = 2

Private complete As Boolean
Private goAuto As Boolean

Private Sub browser_DocumentComplete(ByVal pDisp As Object, url As Variant)
    complete = True
End Sub

Private Sub btnClearStats_Click()
    If MsgBox("Clear the old import data?", vbYesNo) = vbYes Then
        dbExecute "TRUNCATE TABLE KeywordsStatistics"
    End If
End Sub

Private Sub btnExit_Click()
    Unload KeywordsImportTool
End Sub

Private Sub btnGetStats_Click()
    ReDim days(3) As Variant
    days = Array("10", "30", "60", "90")
    goAuto = True
    Mouse.Hourglass True
    Dim warn As Boolean
    warn = WARN_ON_IMPORT
    WARN_ON_IMPORT = False
    Select Case True
        Case Is = Me.opPage(YAHOO)
            deleteImportFiles days, YAHOO
            getYahooStats days
        Case Is = Me.opPage(OVERTURE)
            deleteImportFiles days, OVERTURE
            getOvertureStats days
        Case Is = Me.opPage(ADWORDS)
            ShellWait PERL & " " & "s:\mastest\mas90-signs\keywords_scripts\adwords_get_reports.pl", vbNormalFocus
            If Me.chkRunPerl Then
                ShellWait PERL & " " & "s:\mastest\mas90-signs\keywords_scripts\parse_adwords_import.pl", vbNormalFocus
            End If
    End Select
    WARN_ON_IMPORT = warn
    Mouse.Hourglass False
    goAuto = False
End Sub

Private Sub btnRunNow_Click()
    Select Case True
        Case Is = Me.opPage(YAHOO)
            ShellWait PERL & " " & KWD_Y_IMP_PARSE, vbNormalFocus
        Case Is = Me.opPage(OVERTURE)
            ShellWait PERL & " " & KWD_O_IMP_PARSE, vbNormalFocus
        Case Is = Me.opPage(ADWORDS)
            ShellWait PERL & " " & "s:\mastest\mas90-signs\keywords_scripts\parse_adwords_import.pl", vbNormalFocus
    End Select
End Sub

Private Sub Form_Load()
    Select Case True
        Case Is = Me.opPage(YAHOO)
            Me.browser.Navigate2 "http://store.yahoo.com/"
        Case Is = Me.opPage(OVERTURE)
            Me.browser.Navigate2 "https://secure.overture.com/login.do?mkt=us&locale=en_US"
        Case Is = Me.opPage(ADWORDS)
            Me.browser.Navigate2 "about:blank"
    End Select
End Sub

Private Sub getYahooStats(days() As Variant)
    complete = False
    Me.browser.Navigate2 "http://edit.store.yahoo.com/RT/MGR.toolsplus/"
    waitForPage
    Me.browser.Navigate2 findLink("Track Links")
    waitForPage 1200
    Me.browser.Navigate2 findLink("By Count of Orders")
    waitForPage 1200
    Me.browser.Navigate2 findLink("See All")
    waitForPage 1200
    Dim i As Long
    For i = 0 To 3
        If CBool(Me.chkRange(days(i))) Then
            DoEvents
            Dim j As Long
            With Me.browser.Document.Forms(0)
                For j = 20 To .ALL.Length 'start at 20 to skip first set of radio buttons
                    If UCase(.ALL(j).nodeName) = "INPUT" Then
                        If UCase(.ALL(j).Type) = "RADIO" And .ALL(j).value = i Then
                            .ALL(j).Checked = True
                            Exit For
                        End If
                    End If
                Next j
            End With
            clickSubmitButton 1200
            Dim html As String
            html = Me.browser.Document.body.innerHTML
            Open PAGE_DIR & "\yahoo" & days(i) & ".html" For Output As #1
            Print #1, html
            Close #1
        End If
    Next i
    If CBool(Me.chkRunPerl) Then
        ShellWait PERL & " " & KWD_Y_IMP_PARSE, vbNormalFocus
    End If
    MsgBox "Import from yahoo finished!"
End Sub

Private Sub getOvertureStats(days() As Variant)
    'assumes we're at any page with the big top bar
    complete = False
    Dim link As String, i As Long
    For i = 0 To 3
        DoEvents
        link = "https://secure.overture.com/createReport.do?currentReport.name=Keyword+Summary&fakesubmit=true&locator=external%2FDTC%2Fprecision%2Fsearch_term_summary.props&begin_date={date1}&query_source=All&end_date={date2}&account_id=8464611860&pageNum=1&rpt_level=Aggregate&submitted=Create+Report"
        link = Replace(link, "{date1}", Format(Date - days(i) - 1, "mm%2Fdd%2Fyyyy"))
        link = Replace(link, "{date2}", Format(Date - 1, "mm%2Fdd%2Fyyyy"))
        
        Me.browser.Navigate2 link
        waitForPage 600 '10 mins good?
        Me.loadingtimer.BackColor = RED
        Me.loadingtimer = "API"
        DoEvents
        Dim retryCount As Long
retry:
        If Download(findLink("Spreadsheet File"), PAGE_DIR & "\overture" & days(i) & ".csv") Then
            'ok
        Else
            If retryCount < 5 Then
                MsgBox "Error downloading " & days(i) & ", retrying..."
                DoEvents
                Sleep 5000
                DoEvents
                retryCount = retryCount + 1
                GoTo retry
            Else
                MsgBox "Can't download overture files!"
                Exit Sub
            End If
        End If
        Me.loadingtimer = ""
        Me.loadingtimer.BackColor = GREEN
    Next i

    If CBool(Me.chkRunPerl) Then
        ShellWait PERL & " " & KWD_O_IMP_PARSE, vbNormalFocus
    End If
    MsgBox "Import from overture finished!"
End Sub

Private Sub opPage_Click(Index As Integer)
    Form_Load
End Sub

Private Sub deleteImportFiles(days() As Variant, site As Long)
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim i As Long
    Dim fname As String, extension As String
    Dim fullpath As String
    Select Case site
        Case Is = YAHOO
            fname = "yahoo"
            extension = ".html"
        Case Is = OVERTURE
            fname = "overture"
            extension = ".csv"
    End Select
    For i = 0 To UBound(days)
        fullpath = PAGE_DIR & "\" & fname & days(i) & extension
        If fso.FileExists(fullpath) Then
            fso.DeleteFile fullpath
        End If
    Next i
    Set fso = Nothing
End Sub

Private Function waitForPage(Optional timeout As Long = 60) As Boolean
    Mouse.Hourglass True
    Dim starttime As Double
    starttime = Timer
    Me.loadingtimer.BackColor = RED
    While Not complete
        DoEvents
        Me.loadingtimer = Timer - starttime
        If Timer - starttime > timeout Then
            Me.browser.Stop
            complete = False
            waitForPage = False
            Exit Function
        End If
    Wend
    complete = False
    Me.loadingtimer = ""
    Me.loadingtimer.BackColor = GREEN
    waitForPage = True
    Mouse.Hourglass False
End Function

Private Function findLink(linkText As String) As String
    findLink = ""
    Dim i As Long
    For i = 0 To Me.browser.Document.Links.Length - 1
        If Me.browser.Document.Links(i).innerText = linkText Then
            findLink = Me.browser.Document.Links(i).href
            Exit For
        End If
    Next i
End Function

Private Function clickSubmitButton(Optional timeout As Long = 180)
    Dim i As Long
    With Me.browser.Document.Forms(0)
        For i = 0 To .ALL.Length
            If UCase(.ALL(i).nodeName) = "INPUT" Then
                If UCase(.ALL(i).Type) = "SUBMIT" Then
                    .ALL(i).Click
                    waitForPage timeout
                    Exit For
                End If
            End If
        Next i
    End With
End Function
