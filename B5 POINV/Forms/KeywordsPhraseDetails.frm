VERSION 5.00
Begin VB.Form KeywordsPhraseDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keywords - Details"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame generalFrame 
      Caption         =   "Shared Settings:"
      Height          =   1905
      Index           =   2
      Left            =   90
      TabIndex        =   9
      Top             =   5130
      Width           =   5235
      Begin VB.CheckBox chkSeasonal 
         Caption         =   "Seasonal"
         Height          =   285
         Left            =   270
         TabIndex        =   25
         Top             =   1170
         Width           =   1005
      End
      Begin VB.Frame SeasonsFrame 
         Caption         =   "Months Active:"
         Height          =   1215
         Left            =   1440
         TabIndex        =   12
         Top             =   690
         Width           =   3795
         Begin VB.CheckBox chkSeasonalMonth 
            Caption         =   "Jan"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   24
            Top             =   300
            Width           =   825
         End
         Begin VB.CheckBox chkSeasonalMonth 
            Caption         =   "Feb"
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   23
            Top             =   570
            Width           =   825
         End
         Begin VB.CheckBox chkSeasonalMonth 
            Caption         =   "March"
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   22
            Top             =   840
            Width           =   825
         End
         Begin VB.CheckBox chkSeasonalMonth 
            Caption         =   "April"
            Height          =   255
            Index           =   4
            Left            =   1080
            TabIndex        =   21
            Top             =   300
            Width           =   825
         End
         Begin VB.CheckBox chkSeasonalMonth 
            Caption         =   "May"
            Height          =   255
            Index           =   5
            Left            =   1080
            TabIndex        =   20
            Top             =   570
            Width           =   825
         End
         Begin VB.CheckBox chkSeasonalMonth 
            Caption         =   "June"
            Height          =   255
            Index           =   6
            Left            =   1080
            TabIndex        =   19
            Top             =   840
            Width           =   825
         End
         Begin VB.CheckBox chkSeasonalMonth 
            Caption         =   "Dec"
            Height          =   255
            Index           =   12
            Left            =   2880
            TabIndex        =   18
            Top             =   840
            Width           =   825
         End
         Begin VB.CheckBox chkSeasonalMonth 
            Caption         =   "Nov"
            Height          =   255
            Index           =   11
            Left            =   2880
            TabIndex        =   17
            Top             =   570
            Width           =   825
         End
         Begin VB.CheckBox chkSeasonalMonth 
            Caption         =   "Oct"
            Height          =   255
            Index           =   10
            Left            =   2880
            TabIndex        =   16
            Top             =   300
            Width           =   825
         End
         Begin VB.CheckBox chkSeasonalMonth 
            Caption         =   "Sept"
            Height          =   255
            Index           =   9
            Left            =   1980
            TabIndex        =   15
            Top             =   840
            Width           =   825
         End
         Begin VB.CheckBox chkSeasonalMonth 
            Caption         =   "Aug"
            Height          =   255
            Index           =   8
            Left            =   1980
            TabIndex        =   14
            Top             =   570
            Width           =   825
         End
         Begin VB.CheckBox chkSeasonalMonth 
            Caption         =   "July"
            Height          =   255
            Index           =   7
            Left            =   1980
            TabIndex        =   13
            Top             =   300
            Width           =   825
         End
      End
      Begin VB.TextBox LandingPage 
         Height          =   285
         Left            =   1500
         TabIndex        =   11
         Top             =   300
         Width           =   3165
      End
      Begin VB.CommandButton btnLandingPage 
         Caption         =   "Landing Page"
         Height          =   285
         Left            =   180
         TabIndex        =   10
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.Frame generalFrame 
      Caption         =   "AdWords:"
      Height          =   2145
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   2880
      Width           =   8415
      Begin VB.Frame AdWordsReasonsFrame 
         Caption         =   "Reason Inactive:"
         Height          =   2145
         Left            =   6540
         TabIndex        =   42
         Top             =   0
         Width           =   1875
         Begin VB.OptionButton opAdWordsInactiveReason 
            Caption         =   "Relevance"
            Height          =   255
            Index           =   6
            Left            =   150
            TabIndex        =   49
            Top             =   1440
            Width           =   1545
         End
         Begin VB.OptionButton opAdWordsInactiveReason 
            Caption         =   "Bad Site"
            Height          =   255
            Index           =   5
            Left            =   150
            TabIndex        =   48
            Top             =   1200
            Width           =   1545
         End
         Begin VB.OptionButton opAdWordsInactiveReason 
            Caption         =   "T/D Quality"
            Height          =   255
            Index           =   4
            Left            =   150
            TabIndex        =   47
            Top             =   960
            Width           =   1545
         End
         Begin VB.OptionButton opAdWordsInactiveReason 
            Caption         =   "Other / Unknown"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   46
            Top             =   1800
            Width           =   1545
         End
         Begin VB.OptionButton opAdWordsInactiveReason 
            Caption         =   "T/D Length"
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   45
            Top             =   720
            Width           =   1545
         End
         Begin VB.OptionButton opAdWordsInactiveReason 
            Caption         =   "Low Click Rate"
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   44
            Top             =   480
            Width           =   1545
         End
         Begin VB.OptionButton opAdWordsInactiveReason 
            Caption         =   "Duplicate"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   43
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.ComboBox AdWordsSetStatus 
         Height          =   315
         Left            =   4410
         TabIndex        =   40
         Text            =   "Set Status..."
         Top             =   480
         Width           =   1365
      End
      Begin VB.TextBox AdWordsTrackLink 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   480
         Width           =   1065
      End
      Begin VB.TextBox AdWordsCurrentBid 
         Height          =   285
         Left            =   180
         TabIndex        =   35
         Top             =   480
         Width           =   585
      End
      Begin VB.TextBox AdWordsSuggBid 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   480
         Width           =   585
      End
      Begin VB.ListBox AdWordsPeriodReports 
         Height          =   1035
         ItemData        =   "KeywordsPhraseDetails.frx":0000
         Left            =   120
         List            =   "KeywordsPhraseDetails.frx":0007
         TabIndex        =   8
         Top             =   960
         Width           =   5085
      End
      Begin VB.Label lblAdWordsPhraseStatus 
         Alignment       =   2  'Center
         Caption         =   "CURR_STATUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4050
         TabIndex        =   41
         Top             =   210
         Width           =   2025
      End
      Begin VB.Label generalLabel 
         Caption         =   "Track Link:"
         Height          =   255
         Index           =   6
         Left            =   2310
         TabIndex        =   39
         Top             =   270
         Width           =   945
      End
      Begin VB.Label generalLabel 
         Caption         =   "Sugg. Bid:"
         Height          =   255
         Index           =   5
         Left            =   1110
         TabIndex        =   37
         Top             =   270
         Width           =   885
      End
      Begin VB.Label generalLabel 
         Caption         =   "Curr. Bid:"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   36
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   375
      Left            =   3615
      TabIndex        =   6
      Top             =   7350
      Width           =   1305
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   810
      TabIndex        =   5
      Top             =   7320
      Width           =   405
   End
   Begin VB.CommandButton btnPrevious 
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   390
      TabIndex        =   4
      Top             =   7320
      Width           =   405
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Overture:"
      Height          =   2145
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   8415
      Begin VB.Frame Frame1 
         Caption         =   "Reason Inactive:"
         Height          =   2145
         Left            =   6540
         TabIndex        =   50
         Top             =   0
         Width           =   1875
         Begin VB.OptionButton opOvertureInactiveReason 
            Caption         =   "Duplicate"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   57
            Top             =   240
            Width           =   1545
         End
         Begin VB.OptionButton opOvertureInactiveReason 
            Caption         =   "Low Click Rate"
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   56
            Top             =   480
            Width           =   1545
         End
         Begin VB.OptionButton opOvertureInactiveReason 
            Caption         =   "T/D Length"
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   55
            Top             =   720
            Width           =   1545
         End
         Begin VB.OptionButton opOvertureInactiveReason 
            Caption         =   "Other / Unknown"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   54
            Top             =   1800
            Width           =   1545
         End
         Begin VB.OptionButton opOvertureInactiveReason 
            Caption         =   "T/D Quality"
            Height          =   255
            Index           =   4
            Left            =   150
            TabIndex        =   53
            Top             =   960
            Width           =   1545
         End
         Begin VB.OptionButton opOvertureInactiveReason 
            Caption         =   "Bad Site"
            Height          =   255
            Index           =   5
            Left            =   150
            TabIndex        =   52
            Top             =   1200
            Width           =   1545
         End
         Begin VB.OptionButton opOvertureInactiveReason 
            Caption         =   "Relevance"
            Height          =   255
            Index           =   6
            Left            =   150
            TabIndex        =   51
            Top             =   1440
            Width           =   1545
         End
      End
      Begin VB.ComboBox OvertureSetStatus 
         Height          =   315
         Left            =   4410
         TabIndex        =   32
         Text            =   "Set Status..."
         Top             =   480
         Width           =   1365
      End
      Begin VB.TextBox OvertureTrackLink 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   480
         Width           =   1065
      End
      Begin VB.TextBox OvertureCurrentBid 
         Height          =   285
         Left            =   180
         TabIndex        =   27
         Top             =   480
         Width           =   585
      End
      Begin VB.TextBox OvertureSuggBid 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   480
         Width           =   585
      End
      Begin VB.ListBox OverturePeriodReports 
         Height          =   1035
         ItemData        =   "KeywordsPhraseDetails.frx":0021
         Left            =   180
         List            =   "KeywordsPhraseDetails.frx":0028
         TabIndex        =   3
         Top             =   900
         Width           =   5085
      End
      Begin VB.Label lblOverturePhraseStatus 
         Alignment       =   2  'Center
         Caption         =   "CURR_STATUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4050
         TabIndex        =   33
         Top             =   210
         Width           =   2025
      End
      Begin VB.Label generalLabel 
         Caption         =   "Track Link:"
         Height          =   255
         Index           =   4
         Left            =   2310
         TabIndex        =   31
         Top             =   270
         Width           =   945
      End
      Begin VB.Label generalLabel 
         Caption         =   "Sugg. Bid:"
         Height          =   255
         Index           =   1
         Left            =   1110
         TabIndex        =   29
         Top             =   270
         Width           =   885
      End
      Begin VB.Label generalLabel 
         Caption         =   "Curr. Bid:"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   28
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.TextBox parentWindow 
      Height          =   285
      Left            =   6990
      TabIndex        =   1
      Text            =   "hidden parent window"
      Top             =   6870
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox searchPhrase 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   8325
   End
End
Attribute VB_Name = "KeywordsPhraseDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PROFIT_PCT = 0.2

Private Const RC_OTHER = 0

Private fillingform As Boolean
Private changed As Boolean
Private whichCtl As String

'========================
' PUBLIC
'========================
Public Sub RefreshForm()
    fillForm
End Sub


'========================
' PRIVATE
'========================
Private Sub handleChangedField()
    If changed = True Then
        Select Case whichCtl
            Case Is = "OvertureCurrentBid"
                OvertureCurrentBid_LostFocus
            Case Is = "AdWordsCurrentBid"
                AdWordsCurrentBid_LostFocus
            Case Is = "OvertureSetStatus"
                OvertureSetStatus_LostFocus
            Case Is = "AdWordsSetStatus"
                AdWordsSetStatus_LostFocus
            Case Is = "LandingPage"
                LandingPage_LostFocus
            Case Else
                MsgBox "Could not determine which control you updated!"
        End Select
    End If
End Sub

Private Sub btnExit_Click()
    Unload KeywordsPhraseDetails
End Sub

Private Sub btnLandingPage_Click()
    OpenDefaultApp "http://www.tools-plus.com/" & Me.LandingPage
End Sub

Private Sub btnNext_Click()
    handleChangedField
    Select Case Me.parentWindow
        Case Is = "KeywordsPhraseList"
            KeywordsPhraseList.GoNext Me.SearchPhrase
        Case Is = "KeywordsCostRevenue"
            KeywordsCostRevenue.GoNext Me.SearchPhrase
    End Select
End Sub

Private Sub btnPrevious_Click()
    handleChangedField
    Select Case Me.parentWindow
        Case Is = "KeywordsPhraseList"
            KeywordsPhraseList.GoPrevious Me.SearchPhrase
        Case Is = "KeywordsCostRevenue"
            KeywordsCostRevenue.GoPrevious Me.SearchPhrase
    End Select
End Sub

'Private Sub btnThisSugg_Click()
'    ShellWait PERL & " " & KWD_SUGGESTIONS & " " & qq(Me.SearchPhrase), vbNormalFocus
'    Load KeywordsSuggestDlg
'    KeywordsSuggestDlg.SearchPhrase = Me.SearchPhrase
'    KeywordsSuggestDlg.Show
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageUp Then
        btnPrevious_Click
        KeyCode = 0
    ElseIf KeyCode = vbKeyPageDown Then
        btnNext_Click
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
    ReDim tabs(5) As Long
    tabs(0) = 12
    tabs(1) = 48
    tabs(2) = 78
    tabs(3) = 108
    tabs(4) = 144
    tabs(5) = 186
    SetListTabStops Me.OverturePeriodReports.hwnd, tabs
    SetListTabStops Me.AdWordsPeriodReports.hwnd, tabs
    requerySetStatus
End Sub

Private Sub SearchPhrase_Change()
    fillForm
End Sub

Private Sub requerySetStatus()
    Me.OvertureSetStatus.Clear
    Me.OvertureSetStatus.AddItem "Set Status..."
    Me.OvertureSetStatus = "Set Status..."
    Me.OvertureSetStatus.AddItem "Active"
    Me.OvertureSetStatus.AddItem "Inactive"
    Me.OvertureSetStatus.AddItem "Pending"
    Me.OvertureSetStatus.AddItem "Delete"
    ExpandDropDownToFit Me.OvertureSetStatus
    Me.OvertureSetStatus.SelLength = 0
    Me.AdWordsSetStatus.Clear
    Me.AdWordsSetStatus.AddItem "Set Status..."
    Me.AdWordsSetStatus = "Set Status..."
    Me.AdWordsSetStatus.AddItem "Active"
    Me.AdWordsSetStatus.AddItem "Inactive"
    Me.AdWordsSetStatus.AddItem "Pending"
    Me.AdWordsSetStatus.AddItem "Delete"
    ExpandDropDownToFit Me.AdWordsSetStatus
    Me.AdWordsSetStatus.SelLength = 0
End Sub

Private Sub fillForm()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM vKeywords WHERE SearchPhrase='" & EscapeSQuotes(Me.SearchPhrase) & "'")
    Dim cb As TextBox, sb As TextBox, tl As TextBox, st As Label
    fillingform = True
    Me.LandingPage = Nz(rst("LandingPage"))
    Me.chkSeasonal = SQLBool(rst("Seasonal"))
    While Not rst.EOF
        Select Case rst("Engine")
            Case Is = "Overture"
                Set cb = Me.OvertureCurrentBid
                Set sb = Me.OvertureSuggBid
                Set tl = Me.OvertureTrackLink
                Set st = Me.lblOverturePhraseStatus
            Case Is = "AdWords"
                Set cb = Me.AdWordsCurrentBid
                Set sb = Me.AdWordsSuggBid
                Set tl = Me.AdWordsTrackLink
                Set st = Me.lblAdWordsPhraseStatus
        End Select
        
        If rst("NewBid") = 0 Then
            cb = Format(rst("CurrentBid"), "Currency")
            cb.FontBold = False
        Else
            cb = Format(rst("NewBid"), "Currency")
            cb.FontBold = True
        End If
        sb = Format(rst("SuggestedBid"), "Currency")
        tl = Nz(rst("TrackLink"))
        st.caption = GetKwdStatus(rst("SearchPhrase"), rst("Active"), rst("Seasonal"), rst("Pending"), rst("Rejected"), rst("DoNotUse"))
        
        toggleActiveFrame rst("Engine")
        toggleSeasonalFrame

        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    fillingform = False
    fillReports
    Set cb = Nothing
    Set sb = Nothing
    Set tl = Nothing
    Set st = Nothing
End Sub

Private Sub fillReports()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM vKeywordsCostRevenue WHERE SearchPhrase='" & EscapeSQuotes(Me.SearchPhrase) & "' ORDER BY Engine, TimeFrame")
    Me.OverturePeriodReports.Clear
    Me.AdWordsPeriodReports.Clear
    Dim line As String
    line = vbTab & "Total Cost" & vbTab & "Visits" & vbTab & "Orders" & vbTab & "Revenue" & vbTab & "Gross Profit" & vbTab & "Net Profit"
    Me.OverturePeriodReports.AddItem line
    Me.AdWordsPeriodReports.AddItem line
    While Not rst.EOF
        line = rst("TimeFrame") & vbTab & Format(rst("TotalCost"), "Currency") & vbTab & rst("Clicks") & vbTab & rst("Orders") & vbTab & Format(rst("Revenue"), "Currency") & vbTab & Format(rst("Revenue") * PROFIT_PCT, "Currency") & vbTab & Format(rst("Revenue") * PROFIT_PCT - rst("TotalCost"), "Currency")
        Select Case rst("Engine")
            Case Is = "Overture"
                Me.OverturePeriodReports.AddItem line
            Case Is = "AdWords"
                Me.AdWordsPeriodReports.AddItem line
        End Select
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub toggleActiveFrame(svc As String, Optional reasonCode As Long = -999)
    Dim lbl As Label, opg As Object 'is there a specific object for an option group?
    Select Case svc
        Case Is = "Overture"
            Set lbl = Me.lblOverturePhraseStatus
            Set opg = Me.opOvertureInactiveReason
        Case Is = "AdWords"
            Set lbl = Me.lblAdWordsPhraseStatus
            Set opg = Me.opAdWordsInactiveReason
        Case Else
            'FIXME
            MsgBox "toggleActiveFrame() called with unknown/invalid svc"
            Exit Sub
    End Select
    Dim frameEnabled As Boolean
    frameEnabled = CBool(lbl.caption = "Inactive") Or CBool(lbl.caption = "Rejected")
    Dim i As Long
    For i = opg.LBound To opg.UBound
        opg(i).Enabled = frameEnabled
    Next i
    If frameEnabled Then
        If reasonCode = -999 Then
            reasonCode = DLookup("ReasonCode", "KeywordsAttributes", "SearchPhrase='" & EscapeSQuotes(Me.SearchPhrase) & "'")
        End If
        If reasonCode = -1 Then
            For i = opg.LBound To opg.UBound
                opg(i) = False
            Next i
        Else
            If reasonCode >= opg.LBound And reasonCode <= opg.UBound Then
                opg(reasonCode) = True
            Else
                opg(RC_OTHER) = True
            End If
        End If
    End If
    Set lbl = Nothing
    Set opg = Nothing
End Sub

Private Sub toggleSeasonalFrame()
    Dim i As Long
    For i = Me.chkSeasonalMonth.LBound To Me.chkSeasonalMonth.UBound
        Me.chkSeasonalMonth(i) = False
        Me.chkSeasonalMonth(i).Enabled = CBool(Me.chkSeasonal)
    Next i
    If CBool(Me.chkSeasonal) Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT ActiveMonth FROM KeywordsSeasonalActiveMonths WHERE SearchPhrase='" & EscapeSQuotes(Me.SearchPhrase) & "'")
        While Not rst.EOF
            Me.chkSeasonalMonth(rst("ActiveMonth")) = 1
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
    End If
End Sub


'======================
' SHARED HANDLERS
'======================
Private Sub LandingPage_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            LandingPage_KeyPress KeyCode
        Case Is = vbKeyReturn
            LandingPage_LostFocus
    End Select
End Sub

Private Sub LandingPage_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "LandingPage"
End Sub

Private Sub LandingPage_LostFocus()
    If changed = True Then
        DB.Execute "UPDATE KeywordsPhrases SET LandingPage='" & EscapeSQuotes(Me.LandingPage) & "' WHERE SearchPhrase='" & EscapeSQuotes(Me.SearchPhrase) & "'"
        DB.Execute "UPDATE KeywordsAttributes SET LandingPageChanged=1 WHERE SearchPhrase='" & EscapeSQuotes(Me.SearchPhrase) & "'"
        If IsFormLoaded("KeywordsPhraseList") Then
            KeywordsPhraseList.RefreshLine Me.SearchPhrase
        End If
        changed = False
    End If
End Sub

Private Sub chkSeasonal_Click()
    If Not fillingform Then
        If Me.chkSeasonal Then
            DB.Execute "UPDATE KeywordsPhrases SET Seasonal=1 WHERE SearchPhrase='" & EscapeSQuotes(Me.SearchPhrase) & "'"
        Else
            DB.Execute "UPDATE KeywordsPhrases SET Seasonal=0 WHERE SearchPhrase='" & EscapeSQuotes(Me.SearchPhrase) & "'"
        End If
        toggleSeasonalFrame
        Dim i As Long
        Me.lblOverturePhraseStatus.caption = GetKwdStatusForPhrase(Me.SearchPhrase, "Overture")
        Me.lblAdWordsPhraseStatus.caption = GetKwdStatusForPhrase(Me.SearchPhrase, "AdWords")
        If IsFormLoaded("KeywordsPhraseList") Then
            KeywordsPhraseList.RefreshLine Me.SearchPhrase
        End If
    End If
End Sub

Private Sub chkSeasonalMonth_Click(Index As Integer)
    If Not fillingform Then
        If CBool(Me.chkSeasonalMonth(Index)) Then
            DB.Execute "INSERT INTO KeywordsSeasonalActiveMonths ( SearchPhrase, ActiveMonth ) VALUES ( '" & EscapeSQuotes(Me.SearchPhrase) & "', " & Index & " )"
        Else
            DB.Execute "DELETE FROM KeywordsSeasonalActiveMonths WHERE SearchPhrase='" & EscapeSQuotes(Me.SearchPhrase) & "' AND ActiveMonth=" & Index
        End If
    End If
End Sub


'======================
' SPLIT HANDLERS
'======================

'-------------- INACTIVE REASON OPTION BUTTONS ------------------
Private Sub opOvertureInactiveReason_Click(Index As Integer)
    InactiveReasonHandler "Overture", Index
End Sub

Private Sub opAdWordsInactiveReason_Click(Index As Integer)
    InactiveReasonHandler "AdWords", Index
End Sub

Private Sub InactiveReasonHandler(svc As String, i As Integer)
    If Not fillingform Then
        DB.Execute "UPDATE KeywordsAttributes SET ReasonCode=" & i & " WHERE SearchPhrase='" & EscapeSQuotes(Me.SearchPhrase) & "' AND Service='" & svc & "'"
    End If
End Sub

'--------------- SUGG. BID D-CLICK -----------------
Private Sub OvertureSuggBid_DblClick()
    Me.OvertureCurrentBid = Me.OvertureSuggBid
    changed = True
    whichCtl = "OvertureCurrentBid"
    OvertureCurrentBid_LostFocus
End Sub

Private Sub AdWordsSuggBid_DblClick()
    Me.AdWordsCurrentBid = Me.AdWordsSuggBid
    changed = True
    whichCtl = "AdWordsCurrentBid"
    AdWordsCurrentBid_LostFocus
End Sub

'-------------- CURRENT BID ENTRY -----------------
Private Sub OvertureCurrentBid_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            OvertureCurrentBid_KeyPress KeyCode
        Case Is = vbKeyReturn
            OvertureCurrentBid_LostFocus
    End Select
End Sub

Private Sub AdWordsCurrentBid_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            AdWordsCurrentBid_KeyPress KeyCode
        Case Is = vbKeyReturn
            AdWordsCurrentBid_LostFocus
    End Select
End Sub

Private Sub OvertureCurrentBid_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "OvertureCurrentBid"
End Sub

Private Sub AdWordsCurrentBid_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "AdWordsCurrentBid"
End Sub

Private Sub OvertureCurrentBid_LostFocus()
    If changed = True Then
        CurrentBidHandler "Overture"
    End If
End Sub

Private Sub AdWordsCurrentBid_LostFocus()
    If changed = True Then
        CurrentBidHandler "AdWords"
    End If
End Sub

Private Sub CurrentBidHandler(svc As String)
    Dim cb As TextBox
    Select Case svc
        Case Is = "Overture"
            Set cb = Me.OvertureCurrentBid
        Case Is = "AdWords"
            Set cb = Me.AdWordsCurrentBid
    End Select
    If cb = "" Or cb = "0" Then
        cb = Format(DLookup("CurrentBid", "vKeywords", "SearchPhrase='" & EscapeSQuotes(Me.SearchPhrase) & "' AND Service='" & svc & "'"), "Currency")
        cb.FontBold = False
        SetBid Me.SearchPhrase, svc, 0
        changed = False
        If IsFormLoaded("KeywordsPhraseList") Then
            KeywordsPhraseList.RefreshLine Me.SearchPhrase
        End If
    ElseIf validateCurrency(cb.Text) Then
        cb = Format(cb, "Currency")
        cb.FontBold = True
        SetBid Me.SearchPhrase, svc, cb.Text
        changed = False
        If IsFormLoaded("KeywordsPhraseList") Then
            KeywordsPhraseList.RefreshLine Me.SearchPhrase
        End If
    Else
        MsgBox "New bid must be a currency!"
        cb.SetFocus
    End If
    Set cb = Nothing
End Sub

'----------- STATUS CHANGER -------------
Private Sub OvertureSetStatus_Click()
    changed = True
    whichCtl = "OvertureSetStatus"
    OvertureSetStatus_LostFocus
End Sub

Private Sub AdWordsSetStatus_Click()
    changed = True
    whichCtl = "AdWordsSetStatus"
    AdWordsSetStatus_LostFocus
End Sub

Private Sub OvertureSetStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.OvertureSetStatus, KeyCode, Shift
    If KeyCode = vbKeyReturn Then
        OvertureSetStatus_LostFocus
    End If
End Sub

Private Sub AdWordsSetStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.AdWordsSetStatus, KeyCode, Shift
    If KeyCode = vbKeyReturn Then
        AdWordsSetStatus_LostFocus
    End If
End Sub

Private Sub OvertureSetStatus_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.OvertureSetStatus, KeyAscii
    changed = True
    whichCtl = "OvertureSetStatus"
End Sub

Private Sub AdWordsSetStatus_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.AdWordsSetStatus, KeyAscii
    changed = True
    whichCtl = "AdWordsSetStatus"
End Sub

Private Sub OvertureSetStatus_LostFocus()
    AutoCompleteLostFocus Me.OvertureSetStatus
    If changed = True Then
        SetStatusHandler "Overture"
    End If
End Sub

Private Sub AdWordsSetStatus_LostFocus()
    AutoCompleteLostFocus Me.AdWordsSetStatus
    If changed = True Then
        SetStatusHandler "AdWords"
    End If
End Sub

Private Sub SetStatusHandler(svc As String)
    Dim ctl As ComboBox, lbl As Label
    Select Case svc
        Case Is = "Overture"
            Set ctl = Me.OvertureSetStatus
            Set lbl = Me.lblOverturePhraseStatus
        Case Is = "AdWords"
            Set ctl = Me.AdWordsSetStatus
            Set lbl = Me.lblAdWordsPhraseStatus
    End Select
    Select Case ctl
        Case Is = "Active"
            Dim ok As String
            ok = KwdCheckGoodToGo(Me.SearchPhrase, svc)
            If ok = "OK" Then
                SetKwdStatusActive Me.SearchPhrase, svc
            Else
                MsgBox "Can't set this item active:" & vbCrLf & vbCrLf & ok
            End If
        Case Is = "Inactive"
            SetKwdStatusInactive Me.SearchPhrase, svc
        Case Is = "Pending"
            SetKwdStatusPending Me.SearchPhrase, svc
        Case Is = "Delete"
            If lbl.caption = "Pending" Then
                SetKwdStatusDeleted Me.SearchPhrase, svc
            Else
                SetKwdStatusToBeDeleted Me.SearchPhrase, svc
            End If
    End Select
    lbl.caption = GetKwdStatusForPhrase(Me.SearchPhrase, svc)
    If ctl <> "Set Status..." Then
        ctl = "Set Status..."
        toggleActiveFrame svc
        toggleSeasonalFrame
        If IsFormLoaded("KeywordsPhraseList") Then
            KeywordsPhraseList.RefreshLine Me.SearchPhrase
        End If
    End If
    changed = False
End Sub
