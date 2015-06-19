VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#3.5#0"; "TPControls.ocx"
Begin VB.Form KeywordsCostRevenue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keywords - Cost/Revenue Analysis"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   12420
   StartUpPosition =   1  'CenterOwner
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   0
      Left            =   30
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   2040
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin VB.Timer ScrollTimer 
      Interval        =   85
      Left            =   11850
      Top             =   1080
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   20
      Left            =   240
      TabIndex        =   54
      Top             =   5940
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   16
      Left            =   240
      TabIndex        =   50
      Top             =   5160
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   17
      Left            =   240
      TabIndex        =   51
      Top             =   5340
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   18
      Left            =   240
      TabIndex        =   52
      Top             =   5550
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   19
      Left            =   240
      TabIndex        =   53
      Top             =   5730
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   42
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   43
      Top             =   3780
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   10
      Left            =   240
      TabIndex        =   44
      Top             =   3990
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   11
      Left            =   240
      TabIndex        =   45
      Top             =   4170
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   12
      Left            =   240
      TabIndex        =   46
      Top             =   4380
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   13
      Left            =   240
      TabIndex        =   47
      Top             =   4560
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   14
      Left            =   240
      TabIndex        =   48
      Top             =   4770
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   15
      Left            =   240
      TabIndex        =   49
      Top             =   4950
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   41
      Top             =   3390
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   40
      Top             =   3210
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   39
      Top             =   3000
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   38
      Top             =   2820
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   37
      Top             =   2610
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   36
      Top             =   2430
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   35
      Top             =   2220
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkUseThisKwd 
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   34
      Top             =   2040
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Other Filters"
      Height          =   1185
      Index           =   5
      Left            =   6210
      TabIndex        =   32
      Top             =   60
      Width           =   4515
      Begin VB.CheckBox chkHideOffline 
         Caption         =   "Hide Offline"
         Height          =   195
         Left            =   150
         TabIndex        =   77
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.CommandButton btnEngineEdit 
      Caption         =   "Edit Service Parameters"
      Height          =   345
      Left            =   120
      TabIndex        =   31
      Top             =   810
      Width           =   1905
   End
   Begin VB.Frame generalFrame 
      Height          =   1185
      Index           =   3
      Left            =   3450
      TabIndex        =   19
      Top             =   60
      Width           =   2685
      Begin VB.CheckBox chkFilterByDollars 
         Caption         =   "Filter By $"
         Height          =   225
         Left            =   90
         TabIndex        =   30
         Top             =   30
         Width           =   1005
      End
      Begin VB.Frame generalFrame 
         Height          =   915
         Index           =   1
         Left            =   60
         TabIndex        =   21
         Top             =   210
         Width           =   1275
         Begin VB.OptionButton opDollarType 
            Caption         =   "Net Profit"
            Enabled         =   0   'False
            Height          =   195
            Index           =   6
            Left            =   60
            TabIndex        =   27
            Top             =   630
            Width           =   1035
         End
         Begin VB.OptionButton opDollarType 
            Caption         =   "Gross Profit"
            Enabled         =   0   'False
            Height          =   195
            Index           =   5
            Left            =   60
            TabIndex        =   26
            Top             =   390
            Width           =   1155
         End
         Begin VB.OptionButton opDollarType 
            Caption         =   "Revenue"
            Enabled         =   0   'False
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   25
            Top             =   150
            Width           =   1035
         End
      End
      Begin VB.Frame generalFrame 
         Height          =   915
         Index           =   4
         Left            =   1350
         TabIndex        =   20
         Top             =   210
         Width           =   1275
         Begin VB.OptionButton opDollarAmount 
            Caption         =   "Zero"
            Enabled         =   0   'False
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   24
            Top             =   630
            Width           =   1035
         End
         Begin VB.OptionButton opDollarAmount 
            Caption         =   "Negative"
            Enabled         =   0   'False
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   23
            Top             =   390
            Width           =   1035
         End
         Begin VB.OptionButton opDollarAmount 
            Caption         =   "Positive"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   22
            Top             =   150
            Width           =   1035
         End
      End
   End
   Begin VB.TextBox ProfitMargin 
      Height          =   285
      Left            =   1500
      TabIndex        =   17
      Text            =   "20%"
      Top             =   1350
      Width           =   615
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Select Service:"
      Height          =   675
      Index           =   2
      Left            =   150
      TabIndex        =   16
      Top             =   60
      Width           =   1845
      Begin VB.ComboBox cmbService 
         Height          =   315
         Left            =   60
         TabIndex        =   28
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton btnKeywordMaint 
      Caption         =   "Keyword Maint"
      Height          =   345
      Left            =   11040
      TabIndex        =   15
      Top             =   630
      Width           =   1365
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   345
      Left            =   11010
      TabIndex        =   14
      Top             =   0
      Width           =   1395
   End
   Begin VB.TextBox sumNetProfit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10140
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1410
      Width           =   1035
   End
   Begin VB.TextBox sumGrossProfit 
      Height          =   285
      Left            =   8970
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1410
      Width           =   1035
   End
   Begin VB.TextBox sumRevenue 
      Height          =   285
      Left            =   7770
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1410
      Width           =   1035
   End
   Begin VB.TextBox sumCost 
      Height          =   285
      Left            =   4500
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1410
      Width           =   1035
   End
   Begin VB.ListBox costRevenue 
      Height          =   4155
      Left            =   450
      TabIndex        =   4
      Top             =   2010
      Width           =   11895
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Select Period:"
      Height          =   675
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   60
      Width           =   1335
      Begin VB.ComboBox cmbPeriod 
         Height          =   315
         Left            =   60
         TabIndex        =   29
         Top             =   240
         Width           =   1185
      End
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   1
      Left            =   30
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   2220
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   2
      Left            =   30
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   2430
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   3
      Left            =   30
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   2610
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   4
      Left            =   30
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   2820
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   5
      Left            =   30
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   3000
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   6
      Left            =   30
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   3210
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   7
      Left            =   30
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   3390
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   8
      Left            =   30
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   3600
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   9
      Left            =   30
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   3780
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   10
      Left            =   30
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   3990
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   11
      Left            =   30
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   4170
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   12
      Left            =   30
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   4380
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   13
      Left            =   30
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   4560
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   14
      Left            =   30
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   4770
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   15
      Left            =   30
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   4950
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   16
      Left            =   30
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   5160
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   17
      Left            =   30
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   5340
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   18
      Left            =   30
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   5550
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   19
      Left            =   30
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   5730
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin TPControls.ClickyShape KwdOkWarn 
      Height          =   165
      Index           =   20
      Left            =   30
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   5940
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   291
   End
   Begin VB.Label headerLabel 
      Caption         =   "Net P. %"
      Height          =   225
      Index           =   9
      Left            =   11280
      TabIndex        =   76
      Tag             =   "NetProfitPct"
      Top             =   1770
      Width           =   885
   End
   Begin VB.Label headerLabel 
      Caption         =   "CPC"
      Height          =   225
      Index           =   8
      Left            =   3780
      TabIndex        =   33
      Tag             =   "CostPerClick"
      Top             =   1770
      Width           =   405
   End
   Begin VB.Label generalLabel 
      Caption         =   "Gross Profit Is:"
      Height          =   255
      Index           =   2
      Left            =   390
      TabIndex        =   18
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label headerLabel 
      Caption         =   "Views"
      Height          =   225
      Index           =   2
      Left            =   5610
      TabIndex        =   13
      Tag             =   "Impressions"
      Top             =   1770
      Width           =   585
   End
   Begin VB.Label headerLabel 
      Caption         =   "Net Profit"
      Height          =   225
      Index           =   7
      Left            =   10200
      TabIndex        =   8
      Tag             =   "NetProfit"
      Top             =   1770
      Width           =   885
   End
   Begin VB.Label headerLabel 
      Caption         =   "Gross Profit"
      Height          =   225
      Index           =   6
      Left            =   9030
      TabIndex        =   7
      Tag             =   "GrossProfit"
      Top             =   1770
      Width           =   1065
   End
   Begin VB.Label headerLabel 
      Caption         =   "Revenue"
      Height          =   225
      Index           =   5
      Left            =   7830
      TabIndex        =   6
      Tag             =   "Revenue"
      Top             =   1770
      Width           =   885
   End
   Begin VB.Label headerLabel 
      Caption         =   "Orders"
      Height          =   225
      Index           =   4
      Left            =   7140
      TabIndex        =   5
      Tag             =   "Orders"
      Top             =   1770
      Width           =   615
   End
   Begin VB.Label headerLabel 
      Caption         =   "Cost"
      Height          =   225
      Index           =   1
      Left            =   4740
      TabIndex        =   3
      Tag             =   "TotalCost"
      Top             =   1770
      Width           =   435
   End
   Begin VB.Label headerLabel 
      Caption         =   "Clicks"
      Height          =   225
      Index           =   3
      Left            =   6390
      TabIndex        =   2
      Tag             =   "Clicks"
      Top             =   1770
      Width           =   555
   End
   Begin VB.Label headerLabel 
      Caption         =   "Search Phrase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Tag             =   "SearchPhrase"
      Top             =   1770
      Width           =   1515
   End
End
Attribute VB_Name = "KeywordsCostRevenue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DFLT_POS = 1
Private Const DFLT_NEG = 2
Private Const DFLT_ZERO = 3
Private Const DFLT_REV = 4
Private Const DFLT_GP = 5
Private Const DFLT_NP = 6

Private orderby As String
Private order_desc As Boolean

Private currentService As String
Private currentServiceID As String 'int, but whatever
Private currentPeriod As String
Private currentMargin As String

Private isShopEng As Dictionary
Private kwdOnOff As Dictionary
Private kwdWarn As Dictionary

Private lastfirst As Long

Private fillingForm As Boolean

'******************************************************************************
'  PUBLIC FUNCTIONS
'******************************************************************************
Public Sub GoNext(SearchPhrase As String)
    'searchphrase isn't used, but pass it to keep these similar to keywordsphraselist
    If Me.costRevenue.ListIndex + 1 < Me.costRevenue.ListCount Then
        Me.costRevenue.ListIndex = Me.costRevenue.ListIndex + 1
        KeywordsPhraseDetails.SearchPhrase = getSelectedSearchPhrase(Me.costRevenue)
    End If
End Sub

Public Sub GoPrevious(SearchPhrase As String)
    'searchphrase isn't used, but pass it to keep these similar to keywordsphraselist
    If Me.costRevenue.ListIndex - 1 >= 0 Then
        Me.costRevenue.ListIndex = Me.costRevenue.ListIndex - 1
        KeywordsPhraseDetails.SearchPhrase = getSelectedSearchPhrase(Me.costRevenue)
    End If
End Sub





'******************************************************************************
'  PRIVATE FUNCTIONS
'******************************************************************************
Private Sub btnEngineEdit_Click()
    Load KeywordsEngineMaintenance
    KeywordsEngineMaintenance.LoadEngine Me.cmbService
    KeywordsEngineMaintenance.Show MODAL
End Sub

Private Sub btnExit_Click()
    Unload KeywordsCostRevenue
End Sub

Private Sub btnKeywordMaint_Click()
    Load KeywordsPhraseList
    KeywordsPhraseList.Show
End Sub

Private Sub chkFilterByDollars_Click()
    Dim i As Long
    fillingForm = True
    If Me.chkFilterByDollars Then
        For i = Me.opDollarAmount.LBound To Me.opDollarAmount.UBound
            Me.opDollarAmount(i).Enabled = True
        Next i
        Me.opDollarAmount(Me.opDollarAmount.LBound) = True
        For i = Me.opDollarType.LBound To Me.opDollarType.UBound
            Me.opDollarType(i).Enabled = True
        Next i
        Me.opDollarType(Me.opDollarType.LBound) = True
    Else
        For i = Me.opDollarAmount.LBound To Me.opDollarAmount.UBound
            Me.opDollarAmount(i) = False
            Me.opDollarAmount(i).Enabled = False
        Next i
        For i = Me.opDollarType.LBound To Me.opDollarType.UBound
            Me.opDollarType(i) = False
            Me.opDollarType(i).Enabled = False
        Next i
    End If
    fillingForm = False
    requeryCostRevenue
End Sub

Private Sub chkHideOffline_Click()
    requeryCostRevenue
End Sub

Private Sub KwdOkWarn_Click(Index As Integer)
    If Not fillingForm Then
        Dim thiskwd As String
        thiskwd = Me.costRevenue.List(Me.costRevenue.ListIndex)
        thiskwd = Left(thiskwd, InStr(thiskwd, vbTab) - 1)
        Me.KwdOkWarn(Index).tag = (Me.KwdOkWarn(Index).tag + 1) Mod 3
        kwdWarn.item(thiskwd) = Me.KwdOkWarn(Index).tag
        SetKwdWarnStatus thiskwd, currentServiceID, Me.KwdOkWarn(Index).tag
        Select Case Me.KwdOkWarn(Index).tag
            Case Is = 0
                Me.KwdOkWarn(Index).FillColor = RED
            Case Is = 1
                Me.KwdOkWarn(Index).FillColor = BG_GREY
            Case Is = 2
                Me.KwdOkWarn(Index).FillColor = GREEN
        End Select
        Me.costRevenue.Selected(Me.costRevenue.TopIndex + Index) = True
    End If
End Sub

Private Sub chkUseThisKwd_Click(Index As Integer)
    If Not fillingForm Then
        'Debug.Print "entering chkUseThisKwd_Click(" & Index & ")"
        Dim thiskwd As String
        thiskwd = Me.costRevenue.List(Me.costRevenue.ListIndex)
        thiskwd = Left(thiskwd, InStr(thiskwd, vbTab) - 1)
        If Me.chkUseThisKwd(Index) Then
            kwdOnOff.item(thiskwd) = 1
            SetKwdStatusActive thiskwd, currentServiceID, isShoppingEngine(currentService)
            Me.KwdOkWarn(Index).Visible = True
        Else
            kwdOnOff.item(thiskwd) = 0
            SetKwdStatusDeleted thiskwd, currentServiceID, 8
            Me.KwdOkWarn(Index).Visible = False
        End If
        If Index <> Me.chkUseThisKwd.UBound Then
            Me.chkUseThisKwd(Index + 1).SetFocus
        Else
            If Me.costRevenue.ListIndex = Me.costRevenue.ListCount - 1 Then
            Else
                Me.costRevenue.TopIndex = Me.costRevenue.TopIndex + 1
                Me.costRevenue.ListIndex = Me.costRevenue.ListIndex + 1
            End If
        End If
    End If
End Sub

Private Sub chkUseThisKwd_GotFocus(Index As Integer)
    'Debug.Print "entering chkUseThisKwd_GotFocus(" & Index & ")"
    Me.costRevenue.Selected(Me.costRevenue.TopIndex + Index) = True
End Sub

Private Sub cmbPeriod_Click()
    cmbPeriod_LostFocus
End Sub

Private Sub cmbPeriod_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.cmbPeriod, KeyCode, Shift
End Sub

Private Sub cmbPeriod_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.cmbPeriod, KeyAscii
End Sub

Private Sub cmbPeriod_LostFocus()
    AutoCompleteLostFocus Me.cmbPeriod
    If Me.cmbPeriod = "" Then
        Me.cmbPeriod = currentPeriod
    ElseIf Me.cmbPeriod <> currentPeriod Then
        currentPeriod = Me.cmbPeriod
        requeryCostRevenue
    End If
End Sub

Private Sub cmbService_Click()
    cmbService_LostFocus
End Sub

Private Sub cmbService_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.cmbService, KeyCode, Shift
End Sub

Private Sub cmbService_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.cmbService, KeyAscii
End Sub

Private Sub cmbService_LostFocus()
    AutoCompleteLostFocus Me.cmbService
    If Me.cmbService = "" Then
        Me.cmbService = currentService
    ElseIf Me.cmbService <> currentService Then
        currentService = Me.cmbService
        If Left(currentService, 11) = "(inactive) " Then
            currentService = Replace(currentService, "(inactive) ", "")
        End If
        currentServiceID = DLookup("ID", "KeywordsEngines", "Engine='" & EscapeSQuotes(currentService) & "'")
        requeryCostRevenue
    End If
End Sub

Private Sub costRevenue_Click()
    'If IsFormLoaded("KeywordsPhraseDetails") Then
    '    If KeywordsPhraseDetails.SearchPhrase <> Me.costRevenue Then
    '        KeywordsPhraseDetails.SearchPhrase = getSelectedSearchPhrase(Me.costRevenue)
    '    End If
    'End If
End Sub

Private Sub costRevenue_DblClick()
    'If Not IsFormLoaded("KeywordsPhraseDetails") Then
    '    Load KeywordsPhraseDetails
    '    KeywordsPhraseDetails.parentWindow = "KeywordsCostRevenue"
    '    KeywordsPhraseDetails.SearchPhrase = getSelectedSearchPhrase(Me.costRevenue)
    '    KeywordsPhraseDetails.Show
    'Else
    '    KeywordsPhraseDetails.SearchPhrase = getSelectedSearchPhrase(Me.costRevenue)
    'End If
End Sub

Private Sub costRevenue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If Me.chkUseThisKwd(Me.costRevenue.ListIndex - Me.costRevenue.TopIndex) Then
            Me.chkUseThisKwd(Me.costRevenue.ListIndex - Me.costRevenue.TopIndex) = 0
        Else
            Me.chkUseThisKwd(Me.costRevenue.ListIndex - Me.costRevenue.TopIndex) = 1
        End If
    End If
End Sub

Private Sub costRevenue_Scroll()
    'Debug.Print "entering costRevenue_Scroll"
    updateCheckBoxes
End Sub

Private Sub Form_Load()
    SetListTabs Me.costRevenue, Array(150, 42, 42, 36, 36, 30, 54, 48) 'cost per, total cost, view, clicks, orders, revenue, gp, np
    
    requeryServices
    requeryPeriods
    cacheShopEng
    
    Dim profitPct As String
    profitPct = Replace(Me.ProfitMargin, "%", "")
    profitPct = IIf(profitPct < 1, profitPct, profitPct / 100)
    
    orderby = "SearchPhrase"
    order_desc = False
    requeryCostRevenue
End Sub

Private Sub requeryCostRevenue()
    
    Me.ScrollTimer.Enabled = False 'just to be on the safe side
    
    If isShoppingEngine(currentService) Then
        Me.headerLabel(2).Visible = False
    Else
        Me.headerLabel(2).Visible = True
    End If
    
    Dim profitPct As String
    profitPct = Replace(Me.ProfitMargin, "%", "")
    profitPct = IIf(profitPct < 1, profitPct, profitPct / 100)
    
    Dim tf As String, svc As String, filter As String
    tf = Left(Me.cmbPeriod, InStr(Me.cmbPeriod, " ") - 1)
    svc = Me.cmbService
    If Me.chkFilterByDollars Then
        Select Case True
            Case Is = Me.opDollarType(DFLT_REV)
                filter = "Revenue"
            Case Is = Me.opDollarType(DFLT_GP)
                filter = "(Revenue*" & profitPct & ")"
            Case Is = Me.opDollarType(DFLT_NP)
                filter = "Revenue*" & profitPct & "-TotalCost"
        End Select
        Select Case True
            Case Is = Me.opDollarAmount(DFLT_POS)
                filter = filter & ">0"
            Case Is = Me.opDollarAmount(DFLT_NEG)
                filter = filter & "<0"
            Case Is = Me.opDollarAmount(DFLT_ZERO)
                filter = filter & "=0"
        End Select
    Else
        filter = ""
    End If
    
    If Me.chkHideOffline Then
        filter = IIf(filter = "", "", filter & " AND ") & "DoNotUse=0"
    End If
    
    Dim rst As ADODB.Recordset
    Dim cost As Double, rev As Double, gp As Double
    
    With New ADODB.Command
        .name = "spKeywordsCostRevenue"
        .CommandText = "spKeywordsCostRevenue"
        .CommandType = adCmdStoredProc
        .NamedParameters = True
        .ActiveConnection = DB.handle
        .Parameters.Append .CreateParameter("@eng", adVarChar, adParamInput, 30, CStr(svc))
        .Parameters.Append .CreateParameter("@tf", adInteger, adParamInput, , CLng(tf))
        .Parameters.Append .CreateParameter("@profitpct", adDouble, adParamInput, , CDbl(profitPct))
        .Parameters.Append .CreateParameter("@whereclause", adVarChar, adParamInput, 1024, CStr(filter))
        .Parameters.Append .CreateParameter("@orderby", adVarChar, adParamInput, 30, CStr(orderby & IIf(order_desc, " DESC", "")))
        Set rst = .Execute
    End With
    
    Me.costRevenue.Clear
    Dim badcosts As Boolean
    Me.costRevenue.Visible = False
    Set kwdOnOff = Nothing
    Set kwdOnOff = New Dictionary
    Set kwdWarn = Nothing
    Set kwdWarn = New Dictionary

    While Not rst.EOF
        Me.costRevenue.AddItem rst("SearchPhrase") & vbTab & _
                               IIf(rst("IsAvgCost"), "* ", "") & Format(rst("CostPerClick"), "Currency") & vbTab & _
                               Format(rst("TotalCost"), "Currency") & vbTab & _
                               IIf(rst("Impressions") = -1, "", rst("Impressions")) & vbTab & _
                               rst("Clicks") & vbTab & _
                               rst("Orders") & vbTab & _
                               Format(rst("Revenue"), "Currency") & vbTab & _
                               Format(rst("GrossProfit"), "Currency") & vbTab & _
                               Format(rst("NetProfit"), "Currency") & vbTab & _
                               Format(rst("NetProfitPct"), "#0.00")
        kwdOnOff.Add CStr(rst("SearchPhrase")), IIf(CLng(rst("DoNotUse")) = 0, 1, 0)
        kwdWarn.Add CStr(rst("SearchPhrase")), CStr(rst("WarnStatus"))
        cost = cost + Nz(rst("TotalCost"), 0)
        rev = rev + rst("Revenue")
        gp = gp + rst("Revenue") * profitPct
        rst.MoveNext
    Wend
    updateCheckBoxes
    Me.costRevenue.Visible = True
    rst.Close
    Set rst = Nothing
    If badcosts Then
        Me.headerLabel(1).ForeColor = RED
        Me.headerLabel(1).ToolTipText = "Some items are missing cost data!"
    Else
        Me.headerLabel(1).ForeColor = BLACK
        Me.headerLabel(1).ToolTipText = ""
    End If
    Me.sumCost = Format(cost, "Currency")
    Me.sumRevenue = Format(rev, "Currency")
    Me.sumGrossProfit = Format(gp, "Currency")
    Me.sumNetProfit = Format(gp - cost, "Currency")
    
    lastfirst = Me.costRevenue.TopIndex
    Me.ScrollTimer.Enabled = True
End Sub

Private Sub headerLabel_Click(Index As Integer)
    If Me.headerLabel(Index).tag <> "" Then
        If Me.headerLabel(Index).tag = orderby Then
            order_desc = Not order_desc
        Else
            orderby = Me.headerLabel(Index).tag
            order_desc = True
            Dim i As Long
            For i = 0 To Me.headerLabel.UBound
                Me.headerLabel(i).FontBold = False
            Next i
            Me.headerLabel(Index).FontBold = True
        End If
        requeryCostRevenue
    End If
End Sub

Private Sub requeryServices()
    Me.cmbService.Clear
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, Engine, Active FROM KeywordsEngines ORDER BY Active DESC, ID")
    While Not rst.EOF
        Me.cmbService.AddItem IIf(rst("Active"), "", "(inactive) ") & rst("Engine")
        rst.MoveNext
    Wend
    ExpandDropDownToFit Me.cmbService
    Me.cmbService = Me.cmbService.List(0) 'this breaks if we don't have any services...
    currentService = Me.cmbService.List(0)
    If Left(currentService, 11) = "(inactive) " Then
        currentService = Replace(currentService, "(inactive) ", "")
    End If
    currentServiceID = DLookup("ID", "KeywordsEngines", "Engine='" & EscapeSQuotes(currentService) & "'")
End Sub

Private Sub requeryPeriods()
    'these should all be numbers-space-whatever, otherwise change the parse in requeryCostRevenue
    Me.cmbPeriod.Clear
    Me.cmbPeriod.AddItem "10 days"
    Me.cmbPeriod.AddItem "30 days"
    Me.cmbPeriod.AddItem "60 days"
    Me.cmbPeriod.AddItem "90 days"
    ExpandDropDownToFit Me.cmbPeriod
    Me.cmbPeriod = "10 days"
    currentPeriod = "10 days"
End Sub

Private Function getSelectedSearchPhrase(ctl As ListBox) As String
    getSelectedSearchPhrase = Left(ctl, InStr(ctl, vbTab) - 1)
End Function

Private Sub opDollarAmount_Click(Index As Integer)
    If Not fillingForm Then
        requeryCostRevenue
    End If
End Sub

Private Sub opDollarType_Click(Index As Integer)
    If Not fillingForm Then
        requeryCostRevenue
    End If
End Sub

Private Sub ProfitMargin_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            ProfitMargin_LostFocus
    End Select
End Sub

Private Sub ProfitMargin_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii, True)
End Sub

Private Sub ProfitMargin_LostFocus()
    Me.ProfitMargin = Replace(Me.ProfitMargin, "%", "") & "%"
    If Me.ProfitMargin = "%" Then
        Me.ProfitMargin = currentMargin
    ElseIf Me.ProfitMargin <> currentMargin Then
        currentMargin = Me.ProfitMargin
        requeryCostRevenue
    End If
End Sub

Private Function isShoppingEngine(svc As String) As Boolean
    isShoppingEngine = CBool(isShopEng.item(LCase(svc)))
End Function

Private Sub cacheShopEng()
    If Not isShopEng Is Nothing Then
        Set isShopEng = Nothing
    End If
    Set isShopEng = New Dictionary
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Engine, ProductsOnly FROM KeywordsEngines")
    While Not rst.EOF
        isShopEng.Add LCase(rst("Engine")), CLng(rst("ProductsOnly"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub updateCheckBoxes()
    'Debug.Print "entering updateCheckBoxes"
    fillingForm = True
    Dim i As Long, max As Long, thiskwd As String
    max = IIf(Me.chkUseThisKwd.UBound > Me.costRevenue.ListCount, Me.costRevenue.ListCount - 1, Me.chkUseThisKwd.UBound)
    For i = Me.chkUseThisKwd.LBound To max
        thiskwd = Me.costRevenue.List(Me.costRevenue.TopIndex + i)
        thiskwd = Left(thiskwd, InStr(thiskwd, vbTab) - 1)
        Me.chkUseThisKwd(i).value = kwdOnOff.item(thiskwd)
        Me.chkUseThisKwd(i).Visible = True
        Me.KwdOkWarn(i).tag = kwdWarn.item(thiskwd)
        Select Case Me.KwdOkWarn(i).tag
            Case Is = 0
                Me.KwdOkWarn(i).FillColor = RED
            Case Is = 1
                Me.KwdOkWarn(i).FillColor = BG_GREY
            Case Is = 2
                Me.KwdOkWarn(i).FillColor = GREEN
        End Select
        If Me.chkUseThisKwd(i) Then
            Me.KwdOkWarn(i).Visible = True
        Else
            Me.KwdOkWarn(i).Visible = False
        End If
    Next i
    For i = i To Me.chkUseThisKwd.UBound
        Me.chkUseThisKwd(i).Visible = False
        Me.KwdOkWarn(i).Visible = False
    Next i
    lastfirst = Me.costRevenue.TopIndex
    fillingForm = False
End Sub

Private Sub ScrollTimer_Timer()
    If lastfirst <> Me.costRevenue.TopIndex Then
        lastfirst = Me.costRevenue.TopIndex
        costRevenue_Scroll
    End If
End Sub
