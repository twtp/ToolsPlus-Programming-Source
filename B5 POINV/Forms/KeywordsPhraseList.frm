VERSION 5.00
Begin VB.Form KeywordsPhraseList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keywords - Maintenance"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame generalFrame 
      Caption         =   "Service:"
      Height          =   1305
      Index           =   3
      Left            =   90
      TabIndex        =   99
      Top             =   660
      Width           =   1845
      Begin VB.CommandButton btnEngineEdit 
         Caption         =   "Edit Service Parameters"
         Height          =   495
         Left            =   60
         TabIndex        =   101
         Top             =   690
         Width           =   1695
      End
      Begin VB.ComboBox cmbService 
         Height          =   315
         Left            =   60
         TabIndex        =   100
         Top             =   270
         Width           =   1695
      End
   End
   Begin VB.CommandButton btnFilterOps 
      Caption         =   "Filter Operations"
      Height          =   285
      Left            =   8940
      TabIndex        =   97
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton btnNewKwds 
      Caption         =   "New Keywords"
      Height          =   285
      Left            =   8940
      TabIndex        =   77
      TabStop         =   0   'False
      ToolTipText     =   "See if any items need keywords"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   9240
      TabIndex        =   76
      Top             =   0
      Width           =   1305
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Types of Keywords:"
      Height          =   1275
      Index           =   2
      Left            =   1980
      TabIndex        =   72
      Top             =   660
      Width           =   1665
      Begin VB.OptionButton ShowKeywords 
         Caption         =   "Show All"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   75
         Top             =   300
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton ShowKeywords 
         Caption         =   "Show Phrases"
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   74
         Top             =   900
         Width           =   1335
      End
      Begin VB.OptionButton ShowKeywords 
         Caption         =   "Show Items"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   73
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton btnCostRevenue 
      Caption         =   "Cost/Rev"
      Height          =   285
      Left            =   8940
      TabIndex        =   57
      Top             =   1380
      Width           =   1575
   End
   Begin VB.CommandButton btnMakeTLs 
      Caption         =   "TrackLinks"
      Height          =   285
      Left            =   8940
      TabIndex        =   56
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Phrase"
      Height          =   6525
      Index           =   1
      Left            =   30
      TabIndex        =   11
      Top             =   2040
      Width           =   10485
      Begin VB.ComboBox SetStatus 
         Height          =   315
         Index           =   7
         Left            =   7230
         TabIndex        =   94
         Text            =   "Set Status..."
         Top             =   6060
         Width           =   1365
      End
      Begin VB.ComboBox SetStatus 
         Height          =   315
         Index           =   6
         Left            =   7230
         TabIndex        =   92
         Text            =   "Set Status..."
         Top             =   5280
         Width           =   1365
      End
      Begin VB.ComboBox SetStatus 
         Height          =   315
         Index           =   5
         Left            =   7230
         TabIndex        =   90
         Text            =   "Set Status..."
         Top             =   4500
         Width           =   1365
      End
      Begin VB.ComboBox SetStatus 
         Height          =   315
         Index           =   4
         Left            =   7230
         TabIndex        =   88
         Text            =   "Set Status..."
         Top             =   3720
         Width           =   1365
      End
      Begin VB.ComboBox SetStatus 
         Height          =   315
         Index           =   3
         Left            =   7230
         TabIndex        =   86
         Text            =   "Set Status..."
         Top             =   2940
         Width           =   1365
      End
      Begin VB.ComboBox SetStatus 
         Height          =   315
         Index           =   2
         Left            =   7230
         TabIndex        =   84
         Text            =   "Set Status..."
         Top             =   2160
         Width           =   1365
      End
      Begin VB.ComboBox SetStatus 
         Height          =   315
         Index           =   1
         Left            =   7230
         TabIndex        =   82
         Text            =   "Set Status..."
         Top             =   1380
         Width           =   1365
      End
      Begin VB.ComboBox SetStatus 
         Height          =   315
         Index           =   0
         Left            =   7230
         TabIndex        =   80
         Text            =   "Set Status..."
         Top             =   600
         Width           =   1365
      End
      Begin VB.CommandButton btnPhraseDetails 
         Caption         =   "Details"
         Height          =   585
         Index           =   7
         Left            =   9210
         TabIndex        =   65
         Top             =   5760
         Width           =   795
      End
      Begin VB.CommandButton btnPhraseDetails 
         Caption         =   "Details"
         Height          =   585
         Index           =   6
         Left            =   9210
         TabIndex        =   64
         Top             =   4980
         Width           =   795
      End
      Begin VB.CommandButton btnPhraseDetails 
         Caption         =   "Details"
         Height          =   585
         Index           =   5
         Left            =   9210
         TabIndex        =   63
         Top             =   4200
         Width           =   795
      End
      Begin VB.CommandButton btnPhraseDetails 
         Caption         =   "Details"
         Height          =   585
         Index           =   4
         Left            =   9210
         TabIndex        =   62
         Top             =   3420
         Width           =   795
      End
      Begin VB.CommandButton btnPhraseDetails 
         Caption         =   "Details"
         Height          =   585
         Index           =   3
         Left            =   9210
         TabIndex        =   61
         Top             =   2640
         Width           =   795
      End
      Begin VB.CommandButton btnPhraseDetails 
         Caption         =   "Details"
         Height          =   585
         Index           =   2
         Left            =   9210
         TabIndex        =   60
         Top             =   1860
         Width           =   795
      End
      Begin VB.CommandButton btnPhraseDetails 
         Caption         =   "Details"
         Height          =   585
         Index           =   1
         Left            =   9210
         TabIndex        =   59
         Top             =   1080
         Width           =   795
      End
      Begin VB.CommandButton btnPhraseDetails 
         Caption         =   "Details"
         Height          =   585
         Index           =   0
         Left            =   9210
         TabIndex        =   58
         Top             =   300
         Width           =   795
      End
      Begin VB.VScrollBar subformScrollBar 
         Height          =   6315
         LargeChange     =   8
         Left            =   10110
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   150
         Width           =   285
      End
      Begin VB.TextBox phrase 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Index           =   0
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   300
         Width           =   3015
      End
      Begin VB.TextBox CurrentBid 
         Height          =   285
         Index           =   0
         Left            =   3240
         TabIndex        =   50
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox SuggBid 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   0
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox TrackLink 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   0
         Left            =   5130
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   300
         Width           =   1065
      End
      Begin VB.TextBox LandingPage 
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   47
         Top             =   630
         Width           =   3015
      End
      Begin VB.TextBox TrackLink 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   1
         Left            =   5130
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1065
      End
      Begin VB.TextBox SuggBid 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   1
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   1080
         Width           =   585
      End
      Begin VB.TextBox CurrentBid 
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   44
         Top             =   1080
         Width           =   585
      End
      Begin VB.TextBox phrase 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Index           =   1
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox LandingPage 
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   42
         Top             =   1410
         Width           =   3015
      End
      Begin VB.TextBox TrackLink 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   2
         Left            =   5130
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1065
      End
      Begin VB.TextBox SuggBid 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   2
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1860
         Width           =   585
      End
      Begin VB.TextBox CurrentBid 
         Height          =   285
         Index           =   2
         Left            =   3240
         TabIndex        =   39
         Top             =   1860
         Width           =   585
      End
      Begin VB.TextBox phrase 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Index           =   2
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1860
         Width           =   3015
      End
      Begin VB.TextBox LandingPage 
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   37
         Top             =   2190
         Width           =   3015
      End
      Begin VB.TextBox TrackLink 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   3
         Left            =   5130
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1065
      End
      Begin VB.TextBox SuggBid 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   3
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2640
         Width           =   585
      End
      Begin VB.TextBox CurrentBid 
         Height          =   285
         Index           =   3
         Left            =   3240
         TabIndex        =   34
         Top             =   2640
         Width           =   585
      End
      Begin VB.TextBox phrase 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Index           =   3
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox LandingPage 
         Height          =   285
         Index           =   3
         Left            =   150
         TabIndex        =   32
         Top             =   2970
         Width           =   3015
      End
      Begin VB.TextBox TrackLink 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   4
         Left            =   5130
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3420
         Width           =   1065
      End
      Begin VB.TextBox SuggBid 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   4
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   3420
         Width           =   585
      End
      Begin VB.TextBox CurrentBid 
         Height          =   285
         Index           =   4
         Left            =   3240
         TabIndex        =   29
         Top             =   3420
         Width           =   585
      End
      Begin VB.TextBox phrase 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Index           =   4
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   3420
         Width           =   3015
      End
      Begin VB.TextBox LandingPage 
         Height          =   285
         Index           =   4
         Left            =   150
         TabIndex        =   27
         Top             =   3750
         Width           =   3015
      End
      Begin VB.TextBox TrackLink 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   5
         Left            =   5130
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1065
      End
      Begin VB.TextBox SuggBid 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   5
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   4200
         Width           =   585
      End
      Begin VB.TextBox CurrentBid 
         Height          =   285
         Index           =   5
         Left            =   3240
         TabIndex        =   24
         Top             =   4200
         Width           =   585
      End
      Begin VB.TextBox phrase 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Index           =   5
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   4200
         Width           =   3015
      End
      Begin VB.TextBox LandingPage 
         Height          =   285
         Index           =   5
         Left            =   150
         TabIndex        =   22
         Top             =   4530
         Width           =   3015
      End
      Begin VB.TextBox TrackLink 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   6
         Left            =   5130
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   4980
         Width           =   1065
      End
      Begin VB.TextBox SuggBid 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   6
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   4980
         Width           =   585
      End
      Begin VB.TextBox CurrentBid 
         Height          =   285
         Index           =   6
         Left            =   3240
         TabIndex        =   19
         Top             =   4980
         Width           =   585
      End
      Begin VB.TextBox phrase 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Index           =   6
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   4980
         Width           =   3015
      End
      Begin VB.TextBox LandingPage 
         Height          =   285
         Index           =   6
         Left            =   150
         TabIndex        =   17
         Top             =   5310
         Width           =   3015
      End
      Begin VB.TextBox TrackLink 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   7
         Left            =   5130
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   5760
         Width           =   1065
      End
      Begin VB.TextBox SuggBid 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   7
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   5760
         Width           =   585
      End
      Begin VB.TextBox CurrentBid 
         Height          =   285
         Index           =   7
         Left            =   3240
         TabIndex        =   14
         Top             =   5760
         Width           =   585
      End
      Begin VB.TextBox phrase 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Index           =   7
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   5760
         Width           =   3015
      End
      Begin VB.TextBox LandingPage 
         Height          =   285
         Index           =   7
         Left            =   150
         TabIndex        =   12
         Top             =   6090
         Width           =   3015
      End
      Begin VB.Label lblPhraseStatus 
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
         Index           =   7
         Left            =   6900
         TabIndex        =   95
         Top             =   5760
         Width           =   1905
      End
      Begin VB.Label lblPhraseStatus 
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
         Index           =   6
         Left            =   6900
         TabIndex        =   93
         Top             =   4980
         Width           =   1905
      End
      Begin VB.Label lblPhraseStatus 
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
         Index           =   5
         Left            =   6900
         TabIndex        =   91
         Top             =   4200
         Width           =   1905
      End
      Begin VB.Label lblPhraseStatus 
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
         Index           =   4
         Left            =   6900
         TabIndex        =   89
         Top             =   3420
         Width           =   1905
      End
      Begin VB.Label lblPhraseStatus 
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
         Index           =   3
         Left            =   6900
         TabIndex        =   87
         Top             =   2640
         Width           =   1905
      End
      Begin VB.Label lblPhraseStatus 
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
         Index           =   2
         Left            =   6900
         TabIndex        =   85
         Top             =   1860
         Width           =   1905
      End
      Begin VB.Label lblPhraseStatus 
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
         Index           =   1
         Left            =   6900
         TabIndex        =   83
         Top             =   1080
         Width           =   1905
      End
      Begin VB.Label lblPhraseStatus 
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
         Index           =   0
         Left            =   6900
         TabIndex        =   81
         Top             =   270
         Width           =   1905
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   0
         X2              =   10110
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   0
         X2              =   10110
         Y1              =   1770
         Y2              =   1770
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   0
         X2              =   10110
         Y1              =   2550
         Y2              =   2550
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   0
         X2              =   10110
         Y1              =   3330
         Y2              =   3330
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   0
         X2              =   10110
         Y1              =   4110
         Y2              =   4110
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   0
         X2              =   10110
         Y1              =   4890
         Y2              =   4890
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   0
         X2              =   10110
         Y1              =   5670
         Y2              =   5670
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   0
         X2              =   10110
         Y1              =   6450
         Y2              =   6450
      End
      Begin VB.Label generalLabel 
         Caption         =   "TrackLink"
         Height          =   255
         Index           =   3
         Left            =   5160
         TabIndex        =   55
         Top             =   0
         Width           =   735
      End
      Begin VB.Label generalLabel 
         Caption         =   "Curr. Bid"
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   54
         Top             =   0
         Width           =   675
      End
      Begin VB.Label generalLabel 
         Caption         =   "Sugg. Bid"
         Height          =   255
         Index           =   5
         Left            =   3930
         TabIndex        =   53
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Filters:"
      Height          =   1905
      Index           =   0
      Left            =   3690
      TabIndex        =   2
      Top             =   30
      Width           =   5205
      Begin VB.OptionButton opFilter 
         Caption         =   "Bid != SuggBid"
         Height          =   255
         Index           =   15
         Left            =   1320
         TabIndex        =   98
         Top             =   780
         Width           =   1395
      End
      Begin VB.OptionButton opFilter 
         Caption         =   "Need Bid"
         Height          =   255
         Index           =   13
         Left            =   1320
         TabIndex        =   96
         Top             =   1050
         Width           =   1035
      End
      Begin VB.TextBox filterByDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   4050
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   780
         Width           =   1035
      End
      Begin VB.OptionButton opFilter 
         Caption         =   "New Since:"
         Height          =   255
         Index           =   12
         Left            =   2910
         TabIndex        =   78
         Top             =   780
         Width           =   1125
      End
      Begin VB.OptionButton opFilter 
         Caption         =   "To Be Offloaded"
         Height          =   255
         Index           =   11
         Left            =   2910
         TabIndex        =   70
         Top             =   240
         Width           =   1545
      End
      Begin VB.OptionButton opFilter 
         Caption         =   "New Active"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   69
         Top             =   1050
         Width           =   1215
      End
      Begin VB.TextBox filterByText 
         Height          =   255
         Left            =   3900
         TabIndex        =   68
         Top             =   510
         Width           =   1185
      End
      Begin VB.OptionButton opFilter 
         Caption         =   "Contains:"
         Height          =   255
         Index           =   9
         Left            =   2910
         TabIndex        =   67
         Top             =   510
         Width           =   1035
      End
      Begin VB.OptionButton opFilter 
         Caption         =   "Need Land Page"
         Height          =   255
         Index           =   8
         Left            =   1320
         TabIndex        =   66
         Top             =   1590
         Width           =   1545
      End
      Begin VB.OptionButton opFilter 
         Caption         =   "Need Tracklink"
         Height          =   255
         Index           =   7
         Left            =   1320
         TabIndex        =   10
         Top             =   1320
         Width           =   1515
      End
      Begin VB.OptionButton opFilter 
         Caption         =   "Deleted"
         Height          =   255
         Index           =   6
         Left            =   1320
         TabIndex        =   9
         Top             =   510
         Width           =   1035
      End
      Begin VB.OptionButton opFilter 
         Caption         =   "Rejected"
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   1035
      End
      Begin VB.OptionButton opFilter 
         Caption         =   "Pending"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   1590
         Width           =   1035
      End
      Begin VB.OptionButton opFilter 
         Caption         =   "Inactive"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1035
      End
      Begin VB.OptionButton opFilter 
         Caption         =   "Active"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   1035
      End
      Begin VB.OptionButton opFilter 
         Caption         =   "Default"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton opFilter 
         Caption         =   "All"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   510
         Width           =   1035
      End
   End
   Begin VB.ComboBox jumpToPhrase 
      Height          =   315
      Left            =   900
      TabIndex        =   0
      Top             =   150
      Width           =   2715
   End
   Begin VB.Label lblFilterCount 
      Alignment       =   1  'Right Justify
      Caption         =   "COUNT"
      Height          =   225
      Left            =   8850
      TabIndex        =   71
      Top             =   8640
      Width           =   1305
   End
   Begin VB.Label generalLabel 
      Caption         =   "Jump To:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   210
      Width           =   765
   End
End
Attribute VB_Name = "KeywordsPhraseList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NUM_BOXES As Long = 8

Private Const KWD_F_DEFAULT As Long = 0
Private Const KWD_F_ALL As Long = 1
Private Const KWD_F_ACTIVE As Long = 2
Private Const KWD_F_INACTIVE As Long = 3
Private Const KWD_F_PENDING As Long = 4
Private Const KWD_F_REJECTED As Long = 5
Private Const KWD_F_DELETED As Long = 6
Private Const KWD_F_NEED_TRACKLINK As Long = 7
Private Const KWD_F_NEED_LANDINGPAGE As Long = 8
Private Const KWD_F_BY_TEXT As Long = 9
Private Const KWD_F_NEWLY_ACTIVE As Long = 10
Private Const KWD_F_NEED_OFFLOAD As Long = 11
Private Const KWD_F_BY_DATE As Long = 12
Private Const KWD_F_NEED_BID As Long = 13
'Private Const KWD_F_STALE_BIDS As Long = 14
Private Const KWD_F_BID_NE_SUGG As Long = 15
'Private Const KWD_F_RECENT_BIDS As Long = 16

Private Const KWD_S_ALL As Long = 0
Private Const KWD_S_ITEMS As Long = 1
Private Const KWD_S_PHRASES As Long = 2

'Private Const SVC_OVERTURE As Long = 0
'Private Const SVC_ADWORDS As Long = 1

Private changed As Boolean
Private PhraseList() As String
Private currentPos As Long

Private currentService As String
Private serviceIDLookup As Dictionary

'******************************************************************************
'  PUBLIC FUNCTIONS
'******************************************************************************
Public Sub GoNext(SearchPhrase As String)
    Dim pos As Long
    pos = bsearch(SearchPhrase, PhraseList) + 1
    If pos <= UBound(PhraseList) Then
        Me.subformScrollBar.value = IIf(pos > Me.subformScrollBar.max, Me.subformScrollBar.max, pos)
        KeywordsPhraseDetails.SearchPhrase = PhraseList(pos)
    End If
End Sub

Public Sub GoPrevious(SearchPhrase As String)
    Dim pos As Long
    pos = bsearch(SearchPhrase, PhraseList) - 1
    If pos >= 0 Then
        Me.subformScrollBar.value = IIf(pos > Me.subformScrollBar.max, Me.subformScrollBar.max, pos)
        KeywordsPhraseDetails.SearchPhrase = PhraseList(pos)
    End If
End Sub

Public Function GetListLength()
    GetListLength = UBound(PhraseList)
End Function

Public Function GetList() As String()
    GetList = PhraseList
End Function

Public Function GetListWhereClause(filter As Long) As String
    Dim whereClause As String
    Select Case filter
        Case Is = KWD_F_DEFAULT
            whereClause = "DoNotUse=0"
        Case Is = KWD_F_ALL
            whereClause = ""
        Case Is = KWD_F_ACTIVE
            whereClause = "DoNotUse=0 AND Active=1"
        Case Is = KWD_F_INACTIVE
            whereClause = "DoNotUse=0 AND Active=0 AND Pending=0 AND Rejected=0"
        Case Is = KWD_F_PENDING
            whereClause = "DoNotUse=0 AND Pending=1 AND Active=0"
        Case Is = KWD_F_REJECTED
            whereClause = "DoNotUse=0 AND Rejected=1"
        Case Is = KWD_F_DELETED
            whereClause = "DoNotUse<>0"
        Case Is = KWD_F_NEED_TRACKLINK
            whereClause = "DoNotUse=0 AND TrackLink IS NULL"
        Case Is = KWD_F_NEED_LANDINGPAGE
            whereClause = "DoNotUse=0 AND LandingPage IS NULL"
        Case Is = KWD_F_BY_TEXT
            whereClause = "SearchPhrase LIKE '%" & Me.filterByText & "%'"
        Case Is = KWD_F_NEWLY_ACTIVE
            whereClause = "DoNotUse=0 AND Active=1 AND NewlyActive=1"
        Case Is = KWD_F_NEED_OFFLOAD
            whereClause = "DoNotUse=0 AND Active=1 AND (LandingPageChanged=1 OR TrackLinkChanged=1 OR NewlyActive=1 OR NewBid<>0)"
        Case Is = KWD_F_BY_DATE
            whereClause = "StartDate>='" & Me.filterByDate & "'"
        Case Is = KWD_F_NEED_BID
            whereClause = "DoNotUse=0 AND CurrentBid=0 AND NewBid=0"
        'Case Is = KWD_F_STALE_BIDS
        '    whereclause = "(Active=1 OR Pending=1) AND BidsLastUpdated<=GETDATE()-" & Me.bidStaleDays
        Case Is = KWD_F_BID_NE_SUGG
            whereClause = "(Active=1 OR Pending=1) AND CurrentBid<>0 AND NewBid<>0 AND NewBid<>SuggestedBid AND CurrentBid<>SuggestedBid"
        'Case Is = KWD_F_RECENT_BIDS
        '    whereclause = "(Active=1 OR Pending=1) AND BidsLastUpdated>=GETDATE()-" & Me.bidNewDays / 24
        Case Else
            whereClause = ""
    End Select
    whereClause = whereClause & IIf(whereClause = "", "", " AND ") & "Engine='" & EscapeSQuotes(currentService) & "'"
    Select Case True
        Case Is = Me.ShowKeywords(KWD_S_ALL)
            'nothing to add
        Case Is = Me.ShowKeywords(KWD_S_ITEMS)
            whereClause = whereClause & " AND ItemLink IS NOT NULL AND ItemLink<>''"
        Case Is = Me.ShowKeywords(KWD_S_PHRASES)
            whereClause = whereClause & " AND (ItemLink IS NULL OR ItemLink='')"
    End Select
    GetListWhereClause = whereClause
End Function

Public Sub RefreshLine(phrase As String)
    Dim i As Long
    For i = 0 To NUM_BOXES - 1
        If Me.phrase(i) = phrase Then
            Dim rst As ADODB.Recordset
            Set rst = DB.retrieve("SELECT * FROM vKeywords WHERE SearchPhrase='" & EscapeSQuotes(phrase) & "'")
            fillLine rst, i, , phrase
            rst.Close
            Set rst = Nothing
            i = NUM_BOXES
        End If
    Next i
End Sub

Public Sub RefreshSubform()
    fillSubform Me.subformScrollBar.value
End Sub

Public Function GetCurrentService() As String
    'Select Case True
    '    Case Is = Me.opService(SVC_OVERTURE)
    '        currentService = "Overture"
    '    Case Is = Me.opService(SVC_ADWORDS)
    '        currentService = "AdWords"
    'End Select
    GetCurrentService = currentService
End Function










'******************************************************************************
'  PRIVATE FUNCTIONS
'******************************************************************************
'Private Sub bidNewDays_KeyPress(KeyAscii As Integer)
'    KeyAscii = LimitToNumbers(KeyAscii, True)
'End Sub
'
'Private Sub bidNewDays_LostFocus()
'    opFilter_Click KWD_F_RECENT_BIDS
'End Sub
'
'Private Sub bidStaleDays_KeyPress(KeyAscii As Integer)
'    KeyAscii = LimitToNumbers(KeyAscii, True)
'End Sub
'
'Private Sub bidStaleDays_LostFocus()
'    opFilter_Click KWD_F_STALE_BIDS
'End Sub

Private Sub btnCostRevenue_Click()
    Load KeywordsCostRevenue
    KeywordsCostRevenue.Show
End Sub

Private Sub btnEngineEdit_Click()
    Load KeywordsEngineMaintenance
    KeywordsEngineMaintenance.LoadEngine Me.cmbService
    KeywordsEngineMaintenance.Show MODAL
End Sub

Private Sub btnExit_Click()
    Unload KeywordsPhraseList
End Sub

Private Sub btnFilterOps_Click()
    Load KeywordsFilterOps
    KeywordsFilterOps.Show MODAL
End Sub

'Private Sub btnGetBids_Click()
'    Load KeywordsBidTool
'    KeywordsBidTool.Show
'End Sub

Private Sub btnMakeTLs_Click()
    'Load KeywordsTrackLinkTool
    'KeywordsTrackLinkTool.Show
    MsgBox "removed"
End Sub

Private Sub btnPhraseDetails_Click(Index As Integer)
    Load KeywordsPhraseDetails
    KeywordsPhraseDetails.parentWindow = "KeywordsPhraseList"
    KeywordsPhraseDetails.SearchPhrase = Me.phrase(Index)
    KeywordsPhraseDetails.Show
End Sub

Private Sub btnNewKwds_Click()
    Load KeywordsNewKeywords
    KeywordsNewKeywords.Show MODAL
    requeryForm
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
        requeryForm
    End If
End Sub

Private Sub filterByText_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            filterByText_KeyPress KeyCode
        Case Is = vbKeyReturn
            filterByText_LostFocus
    End Select
End Sub

Private Sub filterByText_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub filterByText_LostFocus()
    If changed = True Then
        changed = False
        If Me.opFilter(KWD_F_BY_TEXT) = False Then
            Me.opFilter(KWD_F_BY_TEXT) = True
        Else
            requeryForm KWD_F_BY_TEXT
        End If
    End If
End Sub

Private Sub Form_Load()
    'requeryViewAt
    requerySetStatus
    requeryServices
    requeryForm KWD_F_DEFAULT
End Sub

Private Sub jumpToPhrase_Click()
    changed = True
    jumpToPhrase_LostFocus
End Sub

Private Sub jumpToPhrase_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.jumpToPhrase, KeyCode, Shift
End Sub

Private Sub jumpToPhrase_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.jumpToPhrase, KeyAscii
    If KeyAscii = vbKeyReturn Then
        jumpToPhrase_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub jumpToPhrase_LostFocus()
    AutoCompleteLostFocus Me.jumpToPhrase
    If changed = True And Me.jumpToPhrase <> "" Then
        changed = False
        Dim pos As Long
        pos = bsearch(Me.jumpToPhrase, PhraseList)
        If pos > Me.subformScrollBar.max Then
            Me.subformScrollBar.value = Me.subformScrollBar.max
        Else
            Me.subformScrollBar.value = pos
        End If
    End If
End Sub

'Private Sub opService_Click(Index As Integer)
'    requeryForm
'End Sub

Private Sub SetStatus_Click(Index As Integer)
    SetStatus_LostFocus Index
End Sub

Private Sub SetStatus_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.SetStatus(Index), KeyCode, Shift
    If KeyCode = vbKeyReturn Then
        SetStatus_LostFocus Index
    End If
End Sub

Private Sub SetStatus_KeyPress(Index As Integer, KeyAscii As Integer)
    AutoCompleteKeyPress Me.SetStatus(Index), KeyAscii
End Sub

Private Sub SetStatus_LostFocus(Index As Integer)
    AutoCompleteLostFocus Me.SetStatus(Index)
    Dim svc As String
    svc = GetCurrentService()
    Dim svcid As String
    svcid = KwdGetServiceIDFromName(svc)
    Select Case Me.SetStatus(Index)
        Case Is = "Active"
            Dim foo As String, ignoreBidAmt As Boolean
            ignoreBidAmt = False
retryCheck:
            foo = KwdCheckGoodToGo(Me.phrase(Index), svc, ignoreBidAmt)
            If foo = "OK" Then
                SetKwdStatusActive Me.phrase(Index), svcid
                Me.lblPhraseStatus(Index).Caption = GetKwdStatusForPhrase(Me.phrase(Index), svc)
            ElseIf CBool(InStr(foo, "Missing bid")) Then
                If vbYes = MsgBox("No bid amount for this item, continue?", vbYesNo + vbDefaultButton2) Then
                    ignoreBidAmt = True
                    GoTo retryCheck
                Else
                    MsgBox "Can't set this item active:" & vbCrLf & vbCrLf & foo
                End If
            Else
                MsgBox "Can't set this item active:" & vbCrLf & vbCrLf & foo
            End If
        Case Is = "Inactive"
            SetKwdStatusInactive Me.phrase(Index), svcid
            Me.lblPhraseStatus(Index).Caption = GetKwdStatusForPhrase(Me.phrase(Index), svc)
        Case Is = "Pending"
            SetKwdStatusPending Me.phrase(Index), svcid
            Me.lblPhraseStatus(Index).Caption = GetKwdStatusForPhrase(Me.phrase(Index), svc)
        Case Is = "Delete"
            If Me.lblPhraseStatus(Index).Caption = "Pending" Then
                SetKwdStatusDeleted Me.phrase(Index), svcid
            Else
                SetKwdStatusToBeDeleted Me.phrase(Index), svcid
            End If
            Me.lblPhraseStatus(Index).Caption = GetKwdStatusForPhrase(Me.phrase(Index), svc)
    End Select
    If Me.SetStatus(Index) <> "Set Status..." Then
        Me.SetStatus(Index) = "Set Status..."
        If IsFormLoaded("KeywordsPhraseDetails") Then
            KeywordsPhraseDetails.RefreshForm
        End If
    End If
End Sub

'Private Sub ViewAt_Click(Index As Integer)
'    ViewAt_LostFocus Index
'End Sub
'
'Private Sub ViewAt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    AutoCompleteKeyDown Me.ViewAt(Index), KeyCode, Shift
'    If KeyCode = vbKeyReturn Then
'        ViewAt_LostFocus Index
'    End If
'End Sub
'
'Private Sub ViewAt_KeyPress(Index As Integer, KeyAscii As Integer)
'    AutoCompleteKeyPress Me.ViewAt(Index), KeyAscii
'End Sub
'
'Private Sub ViewAt_LostFocus(Index As Integer)
'    AutoCompleteLostFocus Me.ViewAt(Index)
'    Select Case Me.ViewAt(Index)
'        Case Is = "Tools Plus"
'            OpenDefaultApp "http://www.tools-plus.com/" & Me.LandingPage(Index)
'        Case Is = "Froogle"
'            OpenDefaultApp "http://froogle.google.com/froogle?q=" & Replace(Me.phrase(Index), " ", "+")
'        Case Is = "Yahoo"
'            OpenDefaultApp "http://search.yahoo.com/search?p=" & Replace(Me.phrase(Index), " ", "+")
'    End Select
'    Me.ViewAt(Index) = "View At..."
'End Sub
'
'Private Sub requeryViewAt()
'    Dim i As Integer
'    For i = Me.ViewAt.LBound To Me.ViewAt.UBound
'        Me.ViewAt(i).Clear
'        Me.ViewAt(i).AddItem "View At..."
'        Me.ViewAt(i) = "View At..."
'        Me.ViewAt(i).AddItem "Tools Plus"
'        Me.ViewAt(i).AddItem "Froogle"
'        Me.ViewAt(i).AddItem "Yahoo"
'        ExpandDropDownToFit Me.ViewAt(i)
'        Me.ViewAt(i).SelLength = 0
'    Next i
'End Sub

Private Sub requerySetStatus()
    Dim i As Long
    For i = Me.SetStatus.LBound To Me.SetStatus.UBound
        Me.SetStatus(i).Clear
        Me.SetStatus(i).AddItem "Set Status..."
        Me.SetStatus(i) = "Set Status..."
        Me.SetStatus(i).AddItem "Active"
        Me.SetStatus(i).AddItem "Inactive"
        Me.SetStatus(i).AddItem "Pending"
        Me.SetStatus(i).AddItem "Delete"
        ExpandDropDownToFit Me.SetStatus(i)
        Me.SetStatus(i).SelLength = 0
    Next i
End Sub

Private Sub requeryServices()
    Me.cmbService.Clear
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, Engine FROM KeywordsEngines ORDER BY ID")
    Set serviceIDLookup = New Dictionary
    While Not rst.EOF
        Me.cmbService.AddItem rst("Engine")
        serviceIDLookup.Add CStr(rst("Engine")), CStr(rst("ID"))
        rst.MoveNext
    Wend
    ExpandDropDownToFit Me.cmbService
    Me.cmbService = Me.cmbService.list(0) 'this breaks if we don't have any services...
    currentService = Me.cmbService.list(0)
End Sub

Private Sub requeryForm(Optional filter As Long = -1)
    Dim i As Long
    If filter = -1 Then
        For i = Me.opFilter.LBound To Me.opFilter.UBound
            If Me.opFilter(i) Then
                filter = i
                Exit For
            End If
        Next i
    End If
    Dim rst As ADODB.Recordset
    Dim whereClause As String
    whereClause = GetListWhereClause(filter)
    'this will filter for just the type of kwds we want
    Set rst = DB.retrieve("EXEC spKeywordsList '" & EscapeSQuotes(currentService) & "', '" & EscapeSQuotes(whereClause) & "'")
    If rst.RecordCount > 0 Then
        ReDim PhraseList(rst.RecordCount - 1) As String
        Me.jumpToPhrase.Clear
        For i = 0 To rst.RecordCount - 1
            PhraseList(i) = rst("SearchPhrase")
            Me.jumpToPhrase.AddItem rst("SearchPhrase")
            rst.MoveNext
        Next i
        Me.subformScrollBar.Min = 0
        Me.subformScrollBar.max = IIf(rst.RecordCount - NUM_BOXES < 0, 0, rst.RecordCount - NUM_BOXES)
        Me.subformScrollBar.value = 0
    Else
        ReDim PhraseList(0) As String
        PhraseList(0) = "NORECORDS"
        Me.subformScrollBar.Min = 0
        Me.subformScrollBar.max = 0
    End If
    Me.lblFilterCount.Caption = rst.RecordCount
    rst.Close
    Set rst = Nothing
    fillSubform 0
End Sub

Private Sub fillSubform(Optional offset As Long = 0)
    Dim rst As ADODB.Recordset
    Dim i As Long
    For i = 0 To NUM_BOXES - 1
        If UBound(PhraseList) >= i Then
            If PhraseList(offset + i) <> "NORECORDS" Then
                Set rst = DB.retrieve("SELECT * FROM vKeywords WHERE SearchPhrase='" & EscapeSQuotes(PhraseList(offset + i)) & "' AND Engine='" & EscapeSQuotes(currentService) & "'")
                SetSubformControlsVisible i, True
                fillLine rst, i, offset
                rst.Close
                Set rst = Nothing
            Else
                While i < NUM_BOXES
                    SetSubformControlsVisible i, False
                    i = i + 1
                Wend
            End If
        Else
            While i < NUM_BOXES
                SetSubformControlsVisible i, False
                i = i + 1
            Wend
        End If
    Next i
End Sub

Private Sub SetSubformControlsVisible(Index As Long, onoff As Boolean)
    Me.phrase(Index).Visible = onoff
    Me.LandingPage(Index).Visible = onoff
    Me.CurrentBid(Index).Visible = onoff
    Me.SuggBid(Index).Visible = onoff
    Me.TrackLink(Index).Visible = onoff
    'Me.GoogleTrackLink(Index).Visible = onoff
    'Me.OtherBids1(Index).Visible = onoff
    'Me.OtherBids2(Index).Visible = onoff
    'Me.OtherBids3(Index).Visible = onoff
    'Me.OtherBids4(Index).Visible = onoff
    'Me.OtherBids5(Index).Visible = onoff
    'Me.btnGoToFroogle(Index).Visible = onoff
    'Me.btnGoToYahoo(Index).Visible = onoff
    'Me.btnGoToWebsite(Index).Visible = onoff
    'Me.btnThisSugg(Index).Visible = onoff
    'Me.ViewAt(Index).Visible = onoff
    Me.btnPhraseDetails(Index).Visible = onoff
    Me.lblPhraseStatus(Index).Visible = onoff
    Me.SetStatus(Index).Visible = onoff
End Sub

Private Sub fillLine(rst As ADODB.Recordset, i As Long, Optional offset As Long = -1, Optional phrase As String = "")
    'must give either offset or phrase as argument!
    Me.phrase(i) = rst("SearchPhrase")
    Me.LandingPage(i) = Nz(rst("LandingPage"))
    If rst("NewBid") = 0 Then
        Me.CurrentBid(i) = Format(rst("CurrentBid"), "Currency")
        Me.CurrentBid(i).FontBold = False
    Else
        Me.CurrentBid(i) = Format(rst("NewBid"), "Currency")
        Me.CurrentBid(i).FontBold = True
    End If
    Me.SuggBid(i) = Format(rst("SuggestedBid"), "Currency")
    Me.TrackLink(i) = Nz(rst("TrackLink"))
    Me.lblPhraseStatus(i).Caption = GetKwdStatus(rst("SearchPhrase"), rst("Active"), rst("Seasonal"), rst("Pending"), rst("Rejected"), rst("DoNotUse"))
    Dim phr As String
    If offset = -1 Then
        phr = phrase
    Else
        phr = PhraseList(offset + i)
    End If
    'If currentService = "Overture" Then
    '    'TODO: how to handle services that don't give bid amounts?
    '    Dim rst2 As ADODB.Recordset
    '    Set rst2 = retrieve("SELECT TOP 5 Bidder, Amount FROM KeywordsPhraseBids WHERE SearchPhrase='" & EscapeSQuotes(phr) & "' ORDER BY Amount DESC", conn)
    '    Dim j As Long
    '    j = 1
    '    While Not rst2.EOF
    '        Select Case j
    '            Case Is = 1
    '                Me.OtherBids1(i) = Format(rst2("Amount"), "Currency")
    '                Me.OtherBids1(i).ToolTipText = rst2("Bidder")
    '            Case Is = 2
    '                Me.OtherBids2(i) = Format(rst2("Amount"), "Currency")
    '                Me.OtherBids2(i).ToolTipText = rst2("Bidder")
    '            Case Is = 3
    '                Me.OtherBids3(i) = Format(rst2("Amount"), "Currency")
    '                Me.OtherBids3(i).ToolTipText = rst2("Bidder")
    '            Case Is = 4
    '                Me.OtherBids4(i) = Format(rst2("Amount"), "Currency")
    '                Me.OtherBids4(i).ToolTipText = rst2("Bidder")
    '            Case Is = 5
    '                Me.OtherBids5(i) = Format(rst2("Amount"), "Currency")
    '                Me.OtherBids5(i).ToolTipText = rst2("Bidder")
    '        End Select
    '        j = j + 1
    '        rst2.MoveNext
    '    Wend
    '    If j < 6 Then
    '        For j = j To 5
    '            Select Case j
    '                Case Is = 1
    '                    Me.OtherBids1(i) = ""
    '                    Me.OtherBids1(i).ToolTipText = ""
    '                Case Is = 2
    '                    Me.OtherBids2(i) = ""
    '                    Me.OtherBids2(i).ToolTipText = ""
    '                Case Is = 3
    '                    Me.OtherBids3(i) = ""
    '                    Me.OtherBids3(i).ToolTipText = ""
    '                Case Is = 4
    '                    Me.OtherBids4(i) = ""
    '                    Me.OtherBids4(i).ToolTipText = ""
    '                Case Is = 5
    '                    Me.OtherBids5(i) = ""
    '                    Me.OtherBids5(i).ToolTipText = ""
    '            End Select
    '        Next j
    '    End If
    '    rst2.Close
    '    Set rst2 = Nothing
    'End If
End Sub

Private Sub opFilter_Click(Index As Integer)
    If Index = KWD_F_BY_DATE Then
        Load GenericCalendarPopUp
        GenericCalendarPopUp.SetUpCalendar IIf(Me.filterByDate = "", Date, Me.filterByDate), "KeywordsPhraseList", "FilterByDate"
        GenericCalendarPopUp.Show MODAL
    End If
    requeryForm CLng(Index)
End Sub

Private Sub ShowKeywords_Click(Index As Integer)
    requeryForm
End Sub

Private Sub subformScrollBar_Change()
    Mouse.Hourglass True
    If PhraseList(0) <> "NORECORDS" Then
        If Abs(Me.subformScrollBar.value - currentPos) > 1 Then
            fillSubform Me.subformScrollBar.value
        Else
            If Me.subformScrollBar.value > currentPos Then
                shuffleUp
                addLineAfter Me.phrase(NUM_BOXES - 2)
            Else
                shuffleDown
                addLineBefore Me.phrase(1)
            End If
        End If
        currentPos = Me.subformScrollBar.value
    End If
    Mouse.Hourglass False
End Sub

Private Sub shuffleUp()
    Dim i As Long
    For i = 0 To NUM_BOXES - 2
        shuffleLines i + 1, i
    Next i
End Sub

Private Sub shuffleDown()
    Dim i As Long
    For i = NUM_BOXES - 1 To 1 Step -1
        shuffleLines i - 1, i
    Next i
End Sub

Private Sub shuffleLines(fromIndex As Long, toIndex As Long)
    Me.phrase(toIndex) = Me.phrase(fromIndex)
    Me.LandingPage(toIndex) = Me.LandingPage(fromIndex)
    Me.CurrentBid(toIndex) = Me.CurrentBid(fromIndex)
    Me.CurrentBid(toIndex).FontBold = Me.CurrentBid(fromIndex).FontBold
    Me.SuggBid(toIndex) = Me.SuggBid(fromIndex)
    Me.TrackLink(toIndex) = Me.TrackLink(fromIndex)
    'Me.GoogleTrackLink(toIndex) = Me.GoogleTrackLink(fromIndex)
    Me.lblPhraseStatus(toIndex).Caption = Me.lblPhraseStatus(fromIndex).Caption
    'Me.OtherBids1(toIndex) = Me.OtherBids1(fromIndex)
    'Me.OtherBids2(toIndex) = Me.OtherBids2(fromIndex)
    'Me.OtherBids3(toIndex) = Me.OtherBids3(fromIndex)
    'Me.OtherBids4(toIndex) = Me.OtherBids4(fromIndex)
    'Me.OtherBids5(toIndex) = Me.OtherBids5(fromIndex)
    'Me.OtherBids1(toIndex).ToolTipText = Me.OtherBids1(fromIndex).ToolTipText
    'Me.OtherBids2(toIndex).ToolTipText = Me.OtherBids2(fromIndex).ToolTipText
    'Me.OtherBids3(toIndex).ToolTipText = Me.OtherBids3(fromIndex).ToolTipText
    'Me.OtherBids4(toIndex).ToolTipText = Me.OtherBids4(fromIndex).ToolTipText
    'Me.OtherBids5(toIndex).ToolTipText = Me.OtherBids5(fromIndex).ToolTipText
End Sub

Private Sub addLineAfter(line As String)
    Dim i As Long
    i = bsearch(line, PhraseList) + 1
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM vKeywords WHERE SearchPhrase='" & EscapeSQuotes(PhraseList(i)) & "' AND Engine='" & currentService & "'")
    Dim offset As Long
    offset = i - NUM_BOXES + 1
    fillLine rst, NUM_BOXES - 1, offset
    rst.Close
    Set rst = Nothing
End Sub

Private Sub addLineBefore(line As String)
    Dim i As Long
    i = bsearch(line, PhraseList) - 1
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM vKeywords WHERE SearchPhrase='" & EscapeSQuotes(PhraseList(i)) & "' AND Engine='" & currentService & "'")
    Dim offset As Long
    offset = i
    fillLine rst, 0, offset
    rst.Close
    Set rst = Nothing
End Sub

Private Sub setDetailWindowOnFocus(i As Integer)
    If IsFormLoaded("KeywordsPhraseDetails") Then
        If KeywordsPhraseDetails.SearchPhrase <> Me.phrase(i) Then
            KeywordsPhraseDetails.SearchPhrase = Me.phrase(i)
        End If
    End If
End Sub



'------------------
' focus stuff
'------------------
Private Sub phrase_GotFocus(Index As Integer)
    setDetailWindowOnFocus Index
End Sub

Private Sub LandingPage_GotFocus(Index As Integer)
    setDetailWindowOnFocus Index
End Sub

Private Sub CurrentBid_GotFocus(Index As Integer)
    setDetailWindowOnFocus Index
End Sub

Private Sub SuggBid_GotFocus(Index As Integer)
    setDetailWindowOnFocus Index
End Sub

Private Sub TrackLink_GotFocus(Index As Integer)
    setDetailWindowOnFocus Index
End Sub

Private Sub SetStatus_GotFocus(Index As Integer)
    setDetailWindowOnFocus Index
End Sub

Private Sub btnPhraseDetails_GotFocus(Index As Integer)
    setDetailWindowOnFocus Index
End Sub

'Private Sub OtherBids1_GotFocus(Index As Integer)
'    setDetailWindowOnFocus Index
'End Sub
'
'Private Sub OtherBids2_GotFocus(Index As Integer)
'    setDetailWindowOnFocus Index
'End Sub
'
'Private Sub OtherBids3_GotFocus(Index As Integer)
'    setDetailWindowOnFocus Index
'End Sub
'
'Private Sub OtherBids4_GotFocus(Index As Integer)
'    setDetailWindowOnFocus Index
'End Sub
'
'Private Sub OtherBids5_GotFocus(Index As Integer)
'    setDetailWindowOnFocus Index
'End Sub



'--------------
' db updating
'--------------
Private Sub LandingPage_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            LandingPage_KeyPress Index, KeyCode
        Case Is = vbKeyReturn
            LandingPage_LostFocus Index
    End Select
End Sub

Private Sub LandingPage_KeyPress(Index As Integer, KeyAscii As Integer)
    changed = True
End Sub

Private Sub LandingPage_LostFocus(Index As Integer)
    If changed = True Then
        DB.Execute "UPDATE KeywordsPhrases SET LandingPage='" & EscapeSQuotes(Me.LandingPage(Index)) & "' WHERE SearchPhrase='" & EscapeSQuotes(Me.phrase(Index)) & "'"
        DB.Execute "UPDATE KeywordsAttributes SET LandingPageChanged=1 WHERE SearchPhrase='" & EscapeSQuotes(Me.phrase(Index)) & "'"
        changed = False
        If IsFormLoaded("KeywordsPhraseDetails") Then
            KeywordsPhraseDetails.LandingPage = Me.LandingPage(Index)
        End If
    End If
End Sub

Private Sub CurrentBid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            CurrentBid_KeyPress Index, KeyCode
        Case Is = vbKeyReturn
            CurrentBid_LostFocus Index
    End Select
End Sub

Private Sub CurrentBid_KeyPress(Index As Integer, KeyAscii As Integer)
    changed = True
End Sub

Private Sub CurrentBid_LostFocus(Index As Integer)
    If changed = True Then
        If Me.CurrentBid(Index) = "" Or Me.CurrentBid(Index) = "0" Then
            Me.CurrentBid(Index) = Format(DLookup("CurrentBid", "vKeywords", "SearchPhrase='" & EscapeSQuotes(Me.phrase(Index)) & "' AND Engine='" & EscapeSQuotes(currentService) & "'"), "Currency")
            SetBid Me.phrase(Index), serviceIDLookup.item(currentService), 0
            Me.CurrentBid(Index).FontBold = False
            changed = False
            If IsFormLoaded("KeywordsPhraseDetails") Then
                KeywordsPhraseDetails.RefreshForm
            End If
        ElseIf validateCurrency(Me.CurrentBid(Index)) Then
            Me.CurrentBid(Index) = Format(Me.CurrentBid(Index), "Currency")
            SetBid Me.phrase(Index), serviceIDLookup.item(currentService), Me.CurrentBid(Index)
            Me.CurrentBid(Index).FontBold = True
            changed = False
            If IsFormLoaded("KeywordsPhraseDetails") Then
                KeywordsPhraseDetails.RefreshForm
            End If
        Else
            MsgBox "New bid must be a currency!"
            Me.CurrentBid(Index).SetFocus
        End If
    End If
End Sub

Private Sub SuggBid_DblClick(Index As Integer)
    Me.CurrentBid(Index) = Me.SuggBid(Index)
    changed = True
    CurrentBid_LostFocus Index
End Sub
