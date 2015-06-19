VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#3.5#0"; "TPControls.ocx"
Begin VB.Form SignMaintenance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sign Maintenance"
   ClientHeight    =   9405
   ClientLeft      =   2685
   ClientTop       =   2955
   ClientWidth     =   13020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   13020
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Sphere 1 Category:"
      Height          =   1095
      Left            =   30
      TabIndex        =   168
      Top             =   1950
      Width           =   2715
      Begin VB.CommandButton sphereOneEditMappingBtn 
         Caption         =   "re-map"
         Height          =   195
         Left            =   960
         TabIndex        =   170
         Top             =   840
         Width           =   735
      End
      Begin VB.ListBox SphereOneCatLst 
         Height          =   645
         Left            =   120
         TabIndex        =   169
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton btnGoToTPP 
      Caption         =   "&P"
      Height          =   255
      Left            =   4650
      TabIndex        =   165
      TabStop         =   0   'False
      ToolTipText     =   "Go to ToolsPlusParts.com website for this item"
      Top             =   30
      Width           =   285
   End
   Begin VB.ComboBox cmbItemPictureType 
      Height          =   315
      ItemData        =   "SignMaintenance.frx":0000
      Left            =   8160
      List            =   "SignMaintenance.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   164
      Top             =   4290
      Width           =   1125
   End
   Begin VB.CommandButton btnImageUpload 
      Caption         =   "Upload Image"
      Height          =   465
      Left            =   7440
      TabIndex        =   163
      Top             =   4230
      Width           =   675
   End
   Begin VB.CommandButton btnPrintScreen 
      Caption         =   "Print Screen"
      Height          =   405
      Left            =   9360
      TabIndex        =   162
      Top             =   0
      Width           =   765
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Shipping Extras:"
      Height          =   615
      Index           =   1
      Left            =   4770
      TabIndex        =   144
      Top             =   3810
      Width           =   1905
      Begin VB.CommandButton btnShippingExtrasEdit 
         Caption         =   "E"
         Height          =   255
         Left            =   1650
         TabIndex        =   146
         Top             =   270
         Width           =   195
      End
      Begin VB.ComboBox cmbShippingExtra 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   145
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton btnRefreshVendorQtys 
      Caption         =   "Refresh &Vendor Qtys"
      Height          =   255
      Left            =   11370
      TabIndex        =   161
      TabStop         =   0   'False
      Top             =   8340
      Width           =   1605
   End
   Begin VB.CommandButton btnGoToEBay 
      Caption         =   "&E"
      Height          =   255
      Left            =   5250
      TabIndex        =   159
      TabStop         =   0   'False
      ToolTipText     =   "Go to the ebay listing for this item"
      Top             =   30
      Width           =   285
   End
   Begin VB.Frame generalFrame 
      Height          =   1755
      Index           =   2
      Left            =   4770
      TabIndex        =   35
      Top             =   420
      Width           =   1905
      Begin VB.CommandButton btnGroupEditor 
         Caption         =   "GE"
         Height          =   225
         Left            =   1530
         TabIndex        =   123
         TabStop         =   0   'False
         ToolTipText     =   "Open the Group Item Editor"
         Top             =   1500
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CheckBox chkWebPointer 
         Caption         =   "Group Template"
         Height          =   225
         Left            =   90
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1665
      End
      Begin VB.CheckBox chkRequestForm 
         Caption         =   "Request"
         Height          =   225
         Left            =   90
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1665
      End
      Begin VB.CheckBox chkToBePublished 
         Caption         =   "To Be Published"
         Height          =   225
         Left            =   90
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   150
         Width           =   1485
      End
      Begin VB.CheckBox chkPublished 
         Caption         =   "Published"
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
         Left            =   90
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   360
         Width           =   1425
      End
      Begin VB.CheckBox chkWebDelete 
         Caption         =   "Web Delete?"
         Height          =   225
         Left            =   90
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   570
         Width           =   1665
      End
      Begin VB.CheckBox chkWebDeleted 
         Caption         =   "Web Deleted"
         Height          =   225
         Left            =   90
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   780
         Width           =   1665
      End
      Begin VB.CheckBox chkWebOrderable 
         Caption         =   "Web Orderable"
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
         Left            =   90
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1665
      End
      Begin TPControls.ClickyShape TBPubPriority 
         Height          =   195
         Left            =   1560
         TabIndex        =   98
         Top             =   150
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   344
      End
   End
   Begin VB.Frame generalFrame 
      Height          =   1695
      Index           =   4
      Left            =   4770
      TabIndex        =   148
      Top             =   2130
      Width           =   1905
      Begin VB.CommandButton secondaryCategoryEditBtn 
         Caption         =   "E"
         Height          =   225
         Left            =   1440
         TabIndex        =   172
         Top             =   1440
         Width           =   195
      End
      Begin VB.TextBox SecondaryCategoryTxt 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   55
         Locked          =   -1  'True
         TabIndex        =   171
         Text            =   "SecondaryCategory"
         Top             =   1400
         Width           =   1365
      End
      Begin VB.CommandButton btnEBayStoreCategorySuggest 
         Caption         =   "S"
         Height          =   225
         Left            =   1650
         TabIndex        =   160
         Top             =   1170
         Width           =   195
      End
      Begin VB.TextBox EBayStoreCategory 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   156
         Text            =   "EBayStoreCategory"
         Top             =   1130
         Width           =   1365
      End
      Begin VB.CommandButton btnEBayStoreCategoryEdit 
         Caption         =   "E"
         Height          =   225
         Left            =   1440
         TabIndex        =   155
         Top             =   1170
         Width           =   195
      End
      Begin VB.TextBox EBayEBayStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   154
         Top             =   600
         Width           =   1785
      End
      Begin VB.CommandButton btnEBayCategorySuggest 
         Caption         =   "S"
         Height          =   225
         Left            =   1650
         TabIndex        =   153
         Top             =   900
         Width           =   195
      End
      Begin VB.CommandButton btnEBayCategoryEdit 
         Caption         =   "E"
         Height          =   225
         Left            =   1440
         TabIndex        =   152
         Top             =   900
         Width           =   195
      End
      Begin VB.TextBox EBayCategory 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   151
         Text            =   "EBayCategory"
         Top             =   870
         Width           =   1365
      End
      Begin VB.CheckBox chkEBayPublished 
         Caption         =   "EBay Published"
         Height          =   225
         Left            =   90
         TabIndex        =   150
         Top             =   360
         Width           =   1665
      End
      Begin VB.CheckBox chkEBayTBPub 
         Caption         =   "EBay T/B Pub"
         Height          =   225
         Left            =   90
         TabIndex        =   149
         Top             =   150
         Width           =   1665
      End
   End
   Begin VB.CommandButton btnTBPubAlert 
      BackColor       =   &H00FFBB00&
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   9150
      Style           =   1  'Graphical
      TabIndex        =   147
      TabStop         =   0   'False
      ToolTipText     =   "To Be Published - Blue priority"
      Top             =   9030
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton btnKirkUnlockSpecs 
      Caption         =   "Unlock"
      Height          =   285
      Left            =   1110
      TabIndex        =   143
      Top             =   3570
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Frame extendedSpecsFrame 
      Height          =   3255
      Left            =   5880
      TabIndex        =   124
      Top             =   4710
      Width           =   7065
      Begin VB.TextBox ExtendedSpecs 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         Index           =   0
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   141
         Top             =   690
         Width           =   6975
      End
      Begin VB.TextBox ExtendedSpecs 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         Index           =   1
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   140
         Top             =   690
         Width           =   6975
      End
      Begin VB.TextBox ExtendedSpecs 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         Index           =   2
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   139
         Top             =   690
         Width           =   6975
      End
      Begin VB.TextBox ExtendedSpecs 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         Index           =   3
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   138
         Top             =   690
         Width           =   6975
      End
      Begin VB.TextBox ExtendedSpecs 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         Index           =   4
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   137
         Top             =   690
         Width           =   6975
      End
      Begin VB.TextBox ExtendedSpecs 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         Index           =   5
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   136
         Top             =   690
         Width           =   6975
      End
      Begin VB.CheckBox chkTabStrip 
         Caption         =   "FAQ"
         Height          =   255
         Index           =   0
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   390
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox chkTabStrip 
         Caption         =   "Features"
         Height          =   255
         Index           =   2
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   120
         Width           =   1065
      End
      Begin VB.CheckBox chkTabStrip 
         Caption         =   "Specs"
         Height          =   255
         Index           =   3
         Left            =   3300
         Style           =   1  'Graphical
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   120
         Width           =   1065
      End
      Begin VB.CheckBox chkTabStrip 
         Caption         =   "Preview"
         Height          =   345
         Index           =   6
         Left            =   4650
         Style           =   1  'Graphical
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   240
         Width           =   795
      End
      Begin VB.CheckBox chkTabStrip 
         Caption         =   "Write-Up"
         Height          =   255
         Index           =   1
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   120
         Width           =   1065
      End
      Begin VB.CheckBox chkTabStrip 
         Caption         =   "Media"
         Height          =   255
         Index           =   4
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   390
         Width           =   1065
      End
      Begin VB.CheckBox chkTabStrip 
         Caption         =   "Imp. Notes"
         Height          =   255
         Index           =   5
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   120
         Width           =   1065
      End
      Begin VB.CommandButton btnExtSpecsInNotepad 
         Caption         =   "&N"
         Height          =   255
         Left            =   6720
         TabIndex        =   128
         TabStop         =   0   'False
         ToolTipText     =   "Open tab's contents in Notepad (or default editor)"
         Top             =   150
         Width           =   285
      End
      Begin VB.CommandButton btnExtendedSpecsSearch 
         Caption         =   "Find"
         Height          =   255
         Left            =   5520
         TabIndex        =   127
         Top             =   150
         Width           =   435
      End
      Begin VB.CommandButton btnExtendedSpecsSearchContinue 
         Caption         =   "->"
         Height          =   255
         Left            =   5970
         TabIndex        =   126
         Top             =   150
         Width           =   255
      End
      Begin VB.CommandButton btnExtendedSpecsValidateHTML 
         Caption         =   "V"
         Height          =   255
         Left            =   6420
         TabIndex        =   125
         TabStop         =   0   'False
         ToolTipText     =   "Validate this tab's HTML"
         Top             =   150
         Width           =   285
      End
      Begin SHDocVwCtl.WebBrowser ExtendedSpecsPreview 
         Height          =   2505
         Left            =   60
         TabIndex        =   142
         Top             =   690
         Visible         =   0   'False
         Width           =   6975
         ExtentX         =   12303
         ExtentY         =   4419
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
   Begin VB.TextBox Website 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   5430
      Locked          =   -1  'True
      TabIndex        =   121
      Top             =   4410
      Width           =   1185
   End
   Begin VB.CommandButton btnTBPubAlert 
      BackColor       =   &H0000FFFF&
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   8130
      Style           =   1  'Graphical
      TabIndex        =   119
      TabStop         =   0   'False
      ToolTipText     =   "To Be Published - Yellow priority"
      Top             =   9030
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton btnTBPubMisc 
      Caption         =   "QOO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   116
      TabStop         =   0   'False
      ToolTipText     =   "To Be Published - quantity on order is > 0"
      Top             =   9030
      Width           =   495
   End
   Begin VB.CommandButton btnTBPubMisc 
      Caption         =   "DC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   10170
      Style           =   1  'Graphical
      TabIndex        =   115
      TabStop         =   0   'False
      ToolTipText     =   "To Be Published - the item this is replacing is discontinued completely"
      Top             =   9030
      Width           =   495
   End
   Begin VB.CommandButton btnTBPubMisc 
      Caption         =   "BDC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   9660
      Style           =   1  'Graphical
      TabIndex        =   114
      TabStop         =   0   'False
      ToolTipText     =   "To Be Published - the item this is replacing is being discontinued"
      Top             =   9030
      Width           =   495
   End
   Begin VB.CommandButton btnTBPubMisc 
      Caption         =   "QOH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   11190
      Style           =   1  'Graphical
      TabIndex        =   113
      TabStop         =   0   'False
      ToolTipText     =   "To Be Published - quantity on hand > 0"
      Top             =   9030
      Width           =   495
   End
   Begin VB.TextBox InternalNotes 
      Height          =   675
      Left            =   1980
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   111
      Top             =   8190
      Width           =   4125
   End
   Begin VB.CommandButton btnTBPubAlert 
      BackColor       =   &H000000FF&
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   7620
      Style           =   1  'Graphical
      TabIndex        =   107
      TabStop         =   0   'False
      ToolTipText     =   "To Be Published - Red priority"
      Top             =   9030
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Last Change"
      Height          =   555
      Index           =   0
      Left            =   2310
      TabIndex        =   105
      Top             =   3030
      Width           =   1125
      Begin VB.TextBox LastChanged 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ComboBox cmbSignFunctions 
      Height          =   315
      Left            =   9750
      Style           =   2  'Dropdown List
      TabIndex        =   104
      Top             =   8160
      Width           =   1515
   End
   Begin TPControls.Picture ManufLogo 
      Height          =   465
      Left            =   6720
      TabIndex        =   97
      Top             =   3690
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   820
   End
   Begin TPControls.Picture ItemPicture 
      Height          =   2355
      Left            =   6720
      TabIndex        =   96
      Top             =   1290
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   4154
   End
   Begin VB.CommandButton btnTBPubAlert 
      BackColor       =   &H0000FF00&
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   95
      TabStop         =   0   'False
      ToolTipText     =   "To Be Published - Green priority"
      Top             =   9030
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton btnGoToWebsite 
      Caption         =   "&W"
      Height          =   255
      Left            =   4950
      TabIndex        =   94
      TabStop         =   0   'False
      ToolTipText     =   "Go to the website for this item"
      Top             =   30
      Width           =   285
   End
   Begin VB.Frame CrossSellFrame 
      Caption         =   "Up Sell Links:"
      Height          =   2025
      Left            =   9450
      TabIndex        =   87
      Top             =   1290
      Width           =   3525
      Begin VB.CommandButton btnCrossSellPaste 
         Caption         =   "P"
         Height          =   225
         Left            =   1680
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton btnCrossSellCopy 
         Caption         =   "C"
         Height          =   225
         Left            =   1410
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin TPControls.SimpleListView CrossSellList 
         Height          =   1755
         Left            =   60
         TabIndex        =   99
         Top             =   210
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   3096
         MultiSelect     =   0   'False
         Sorted          =   0   'False
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton btnCrossSellPush 
         Caption         =   "P"
         Height          =   225
         Left            =   3240
         TabIndex        =   93
         TabStop         =   0   'False
         ToolTipText     =   "Push these cross sell links to similar items"
         Top             =   570
         Width           =   225
      End
      Begin VB.CommandButton btnCrossSellMoveDown 
         Caption         =   "q"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3240
         TabIndex        =   92
         TabStop         =   0   'False
         ToolTipText     =   "Move selected item lower in the list"
         Top             =   1530
         Width           =   225
      End
      Begin VB.CommandButton btnCrossSellMoveUp 
         Caption         =   "p"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3240
         TabIndex        =   91
         TabStop         =   0   'False
         ToolTipText     =   "Move selected item higher in the list"
         Top             =   1290
         Width           =   225
      End
      Begin VB.CommandButton btnCrossSellDelete 
         Caption         =   "D"
         Height          =   225
         Left            =   3240
         TabIndex        =   90
         TabStop         =   0   'False
         ToolTipText     =   "Remove selected line from the cross sell list"
         Top             =   1050
         Width           =   225
      End
      Begin VB.CommandButton btnCrossSellAdd 
         Caption         =   "A"
         Height          =   225
         Left            =   3240
         TabIndex        =   89
         TabStop         =   0   'False
         ToolTipText     =   "Add a new line to the include list, if a line is selected, it is inserted before that one, otherwise at the end"
         Top             =   810
         Width           =   225
      End
      Begin VB.CommandButton btnCrossSellSuggest 
         Caption         =   "S"
         Height          =   225
         Left            =   3240
         TabIndex        =   88
         TabStop         =   0   'False
         ToolTipText     =   "Create cross-sell links by looking at similar items"
         Top             =   330
         Width           =   225
      End
   End
   Begin VB.ComboBox cmbOtherFunctions 
      Height          =   315
      Left            =   9750
      Style           =   2  'Dropdown List
      TabIndex        =   81
      Top             =   8520
      Width           =   1515
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Triad Code"
      Height          =   555
      Index           =   11
      Left            =   3450
      TabIndex        =   77
      Top             =   3030
      Width           =   1275
      Begin VB.TextBox TriadCode 
         Height          =   255
         Left            =   30
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   210
         Width           =   615
      End
      Begin VB.CommandButton btnApplyTriadCodeToFilter 
         Caption         =   "All"
         Height          =   255
         Left            =   660
         TabIndex        =   79
         TabStop         =   0   'False
         ToolTipText     =   "Apply this Triad Code to all items in the current filter"
         Top             =   210
         Width           =   285
      End
      Begin VB.CommandButton btnClearTriadCode 
         Caption         =   "C"
         Height          =   255
         Left            =   960
         TabIndex        =   78
         TabStop         =   0   'False
         ToolTipText     =   "Clear this Triad Code from any items that have it"
         Top             =   210
         Width           =   255
      End
   End
   Begin VB.Frame IncludesFrame 
      Caption         =   "Includes:"
      Height          =   1335
      Left            =   9450
      TabIndex        =   76
      Top             =   3330
      Width           =   3525
      Begin VB.CommandButton btnIncludesPNAdd 
         Caption         =   "PN"
         Height          =   225
         Left            =   1440
         TabIndex        =   103
         ToolTipText     =   "Add Item # line, maybe"
         Top             =   0
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton btnIncludesMassAdd 
         Caption         =   "M"
         Height          =   225
         Left            =   930
         TabIndex        =   102
         TabStop         =   0   'False
         ToolTipText     =   "Add new lines by single c+p"
         Top             =   0
         Width           =   225
      End
      Begin VB.CommandButton btnIncludesMoveDown 
         Caption         =   "q"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3240
         TabIndex        =   86
         TabStop         =   0   'False
         ToolTipText     =   "Move selected item lower in the list"
         Top             =   960
         Width           =   225
      End
      Begin VB.CommandButton btnIncludesMoveUp 
         Caption         =   "p"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3240
         TabIndex        =   85
         TabStop         =   0   'False
         ToolTipText     =   "Move selected item higher in the list"
         Top             =   720
         Width           =   225
      End
      Begin VB.CommandButton btnIncludesDelete 
         Caption         =   "D"
         Height          =   225
         Left            =   3240
         TabIndex        =   84
         TabStop         =   0   'False
         ToolTipText     =   "Remove selected item from the includes list"
         Top             =   480
         Width           =   225
      End
      Begin VB.CommandButton btnIncludesAdd 
         Caption         =   "A"
         Height          =   225
         Left            =   3240
         TabIndex        =   83
         TabStop         =   0   'False
         ToolTipText     =   "Add a new item to the include list, if an item is selected, it is inserted before that one, otherwise at the end"
         Top             =   240
         Width           =   225
      End
      Begin VB.ListBox IncludesList 
         Height          =   1035
         Left            =   60
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   210
         Width           =   3135
      End
   End
   Begin VB.Frame chkFrame 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   90
      TabIndex        =   71
      Top             =   7920
      Width           =   1815
      Begin VB.CheckBox chkMetaNoIndex 
         Caption         =   "No-Index"
         Height          =   225
         Left            =   120
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   30
         Width           =   1485
      End
      Begin VB.CheckBox chkPrintSignsByQOH 
         Caption         =   "Sign by QOH"
         Height          =   225
         Left            =   120
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   750
         Width           =   1485
      End
      Begin VB.CheckBox chkDCSpecs 
         Caption         =   "D/C Specs"
         Height          =   225
         Left            =   120
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   270
         Width           =   1485
      End
      Begin VB.CheckBox chkPrintSign 
         Caption         =   "Print Sign Y/N?"
         Height          =   225
         Left            =   120
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   510
         Width           =   1485
      End
   End
   Begin VB.CommandButton btnOptionsDialog 
      Caption         =   "Opt"
      Height          =   255
      Left            =   11100
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   60
      Width           =   405
   End
   Begin VB.CommandButton btnGoToInvMaint 
      Caption         =   "Go To &Inv Maint"
      Height          =   315
      Left            =   11730
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   9030
      Width           =   1275
   End
   Begin VB.CommandButton btnFilterByForm 
      Caption         =   "&FBF"
      Height          =   255
      Left            =   10410
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   60
      Width           =   675
   End
   Begin VB.CommandButton btnNextLC 
      Caption         =   "uu"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2850
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   9030
      Width           =   495
   End
   Begin VB.CommandButton btnPrevLC 
      Caption         =   "tt"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   9030
      Width           =   495
   End
   Begin VB.CommandButton btnFirstItem 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   30
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   9030
      Width           =   315
   End
   Begin VB.CommandButton btnLastItem 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3360
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   9030
      Width           =   315
   End
   Begin VB.CheckBox chkHideImages 
      Caption         =   "Hide Pics"
      Height          =   465
      Left            =   6690
      Style           =   1  'Graphical
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   4230
      Width           =   675
   End
   Begin VB.CommandButton btnWebOffload 
      Caption         =   "Web &Offload"
      Height          =   255
      Left            =   11370
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   8610
      Width           =   1605
   End
   Begin VB.CommandButton btnRefreshQtys 
      Caption         =   "&Refresh Quantities"
      Height          =   255
      Left            =   11370
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   8070
      Width           =   1605
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   405
      Left            =   11820
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   0
      Width           =   1185
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Specials"
      Height          =   975
      Index           =   7
      Left            =   6180
      TabIndex        =   54
      Top             =   7950
      Width           =   3495
      Begin VB.CommandButton btnSpecialsFilter 
         Caption         =   "Filter"
         Height          =   285
         Left            =   2760
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   300
         Width           =   675
      End
      Begin VB.CommandButton btnSpecialsChange 
         Caption         =   "Edit"
         Height          =   285
         Left            =   2760
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   600
         Width           =   675
      End
      Begin VB.ListBox Specials 
         Height          =   645
         Left            =   90
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   240
         Width           =   2625
      End
   End
   Begin VB.CommandButton btnRemoveFilter 
      Caption         =   "&All Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7830
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   0
      Width           =   1365
   End
   Begin VB.Frame WebPathsFrame 
      Caption         =   "Web Paths:"
      Height          =   1695
      Left            =   30
      TabIndex        =   46
      Top             =   420
      Width           =   2715
      Begin VB.ListBox WebPathTree 
         Height          =   1035
         Left            =   120
         TabIndex        =   110
         Top             =   480
         Width           =   2535
      End
      Begin VB.CommandButton btnWebPathEdit 
         Caption         =   "Edit"
         Height          =   225
         Left            =   90
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   270
         Width           =   915
      End
      Begin VB.CommandButton btnWebPathPaste 
         Caption         =   "Paste"
         Height          =   225
         Left            =   1890
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   270
         Width           =   735
      End
      Begin VB.CommandButton btnWebPathCopy 
         Caption         =   "Copy"
         Height          =   225
         Left            =   1140
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Cross Sells:"
      Height          =   2415
      Index           =   3
      Left            =   2790
      TabIndex        =   41
      Top             =   420
      Width           =   1935
      Begin VB.ListBox Accessories 
         Height          =   1620
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   720
         Width           =   1755
      End
      Begin VB.CommandButton btnAddAccessories 
         Caption         =   "Add CrossSells / Accy"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Formerly known as 'Accessories'"
         Top             =   480
         Width           =   1725
      End
      Begin VB.CommandButton btnAccessoriesCopy 
         Caption         =   "Copy"
         Height          =   225
         Left            =   120
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton btnAccessoriesPaste 
         Caption         =   "Paste"
         Height          =   225
         Left            =   990
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame generalFrame 
      Height          =   1395
      Index           =   5
      Left            =   30
      TabIndex        =   31
      Top             =   3870
      Width           =   4695
      Begin VB.TextBox Desc 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   720
         TabIndex        =   12
         Top             =   1050
         Width           =   3915
      End
      Begin VB.TextBox Desc 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   9
         Top             =   150
         Width           =   3915
      End
      Begin VB.TextBox Desc 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   2
         Left            =   720
         TabIndex        =   10
         Top             =   450
         Width           =   3915
      End
      Begin VB.TextBox Desc 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   11
         Top             =   750
         Width           =   3915
      End
      Begin VB.Label lblDesc 
         Caption         =   "Desc 4:"
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   75
         ToolTipText     =   "This field should be for any extras or specials; i.e., w/ free bit set"
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblDesc 
         Caption         =   "Desc 1:"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   34
         Tag             =   "This field is for ADJECTIVES for the item; i.e., 18 gauge, 1/2'', left tilting"
         Top             =   180
         Width           =   615
      End
      Begin VB.Label lblDesc 
         Caption         =   "Desc 2:"
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   33
         ToolTipText     =   "This field is for the ITEM NUMBER"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblDesc 
         Caption         =   "Desc 3:"
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   32
         ToolTipText     =   "This field is a simple explanation of what the item is; i.e., sds rotary hammer"
         Top             =   780
         Width           =   615
      End
   End
   Begin VB.Frame generalFrame 
      Height          =   2595
      Index           =   6
      Left            =   30
      TabIndex        =   22
      Top             =   5280
      Width           =   5835
      Begin VB.TextBox Spec 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   13
         Top             =   150
         Width           =   5055
      End
      Begin VB.TextBox Spec 
         Height          =   285
         Index           =   2
         Left            =   720
         TabIndex        =   14
         Top             =   450
         Width           =   5055
      End
      Begin VB.TextBox Spec 
         Height          =   285
         Index           =   8
         Left            =   720
         TabIndex        =   20
         Top             =   2250
         Width           =   5055
      End
      Begin VB.TextBox Spec 
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   15
         Top             =   750
         Width           =   5055
      End
      Begin VB.TextBox Spec 
         Height          =   285
         Index           =   4
         Left            =   720
         TabIndex        =   16
         Top             =   1050
         Width           =   5055
      End
      Begin VB.TextBox Spec 
         Height          =   285
         Index           =   5
         Left            =   720
         TabIndex        =   17
         Top             =   1350
         Width           =   5055
      End
      Begin VB.TextBox Spec 
         Height          =   285
         Index           =   6
         Left            =   720
         TabIndex        =   18
         Top             =   1650
         Width           =   5055
      End
      Begin VB.TextBox Spec 
         Height          =   285
         Index           =   7
         Left            =   720
         TabIndex        =   19
         Top             =   1950
         Width           =   5055
      End
      Begin VB.Label lblSpec 
         Caption         =   "H-lite 1:"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   30
         Top             =   180
         Width           =   615
      End
      Begin VB.Label lblSpec 
         Caption         =   "H-lite 2:"
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   29
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblSpec 
         Caption         =   "H-lite 3:"
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   28
         Top             =   780
         Width           =   615
      End
      Begin VB.Label lblSpec 
         Caption         =   "H-lite 4:"
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   27
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblSpec 
         Caption         =   "H-lite 5:"
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   26
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label lblSpec 
         Caption         =   "H-lite 6:"
         Height          =   255
         Index           =   6
         Left            =   90
         TabIndex        =   25
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblSpec 
         Caption         =   "H-lite 7:"
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   24
         Top             =   1980
         Width           =   615
      End
      Begin VB.Label lblSpec 
         Caption         =   "H-lite 8:"
         Height          =   255
         Index           =   8
         Left            =   90
         TabIndex        =   23
         Top             =   2280
         Width           =   615
      End
   End
   Begin VB.ComboBox cmbFilters 
      Height          =   315
      ItemData        =   "SignMaintenance.frx":0040
      Left            =   5550
      List            =   "SignMaintenance.frx":0042
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
      Width           =   2265
   End
   Begin VB.TextBox ItemNumber 
      Height          =   285
      Left            =   5310
      TabIndex        =   8
      Top             =   30
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox jumpToLine 
      Height          =   315
      Left            =   1110
      TabIndex        =   1
      Top             =   0
      Width           =   765
   End
   Begin VB.ComboBox jumpToItem 
      Height          =   315
      Left            =   3030
      TabIndex        =   5
      Top             =   0
      Width           =   1545
   End
   Begin VB.TextBox txtPosition 
      Height          =   285
      Left            =   3750
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   9030
      Width           =   795
   End
   Begin VB.HScrollBar barPosition 
      Height          =   285
      LargeChange     =   10
      Left            =   840
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   9840
      Width           =   1905
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Availability:"
      Height          =   825
      Index           =   8
      Left            =   6720
      TabIndex        =   50
      Top             =   420
      Width           =   6255
      Begin VB.TextBox EBayAvailabilityLimit 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   5580
         TabIndex        =   157
         TabStop         =   0   'False
         Text            =   "-99999"
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox Availability 
         Height          =   315
         Left            =   60
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   210
         Width           =   4515
      End
      Begin VB.TextBox AvailabilityLimit 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   5580
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "-99999"
         Top             =   180
         Width           =   615
      End
      Begin VB.Label generalLabel 
         Caption         =   "EBay:"
         Height          =   255
         Index           =   0
         Left            =   5100
         TabIndex        =   158
         Top             =   510
         Width           =   465
      End
      Begin VB.Label lblAvailability 
         Caption         =   "Longish availability string goes here...."
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
         Left            =   90
         TabIndex        =   70
         Top             =   540
         Width           =   4995
      End
      Begin VB.Label generalLabel 
         Caption         =   "Displays:"
         Height          =   255
         Index           =   6
         Left            =   4890
         TabIndex        =   51
         Top             =   210
         Width           =   675
      End
   End
   Begin PoInv.HoldDownButton btnPrevItem 
      Height          =   345
      Left            =   840
      TabIndex        =   166
      TabStop         =   0   'False
      Top             =   9030
      Width           =   1000
      _ExtentX        =   1773
      _ExtentY        =   609
      TimeBeforeHoldDown=   500
      TimeBetweenRepeats=   30
      Caption         =   "t"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PoInv.HoldDownButton btnNextItem 
      Height          =   345
      Left            =   1850
      TabIndex        =   167
      TabStop         =   0   'False
      Top             =   9030
      Width           =   1000
      _ExtentX        =   1773
      _ExtentY        =   609
      TimeBeforeHoldDown=   500
      TimeBetweenRepeats=   30
      Caption         =   "u"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label generalLabel 
      Caption         =   "Website:"
      Height          =   255
      Index           =   2
      Left            =   4770
      TabIndex        =   120
      Top             =   4440
      Width           =   645
   End
   Begin VB.Label lblTitleCount 
      Caption         =   "Title count"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   525
      Left            =   4770
      TabIndex        =   117
      Top             =   4710
      Width           =   675
   End
   Begin VB.Label generalLabel 
      Caption         =   "INTERNAL NOTES:"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   112
      Top             =   7950
      Width           =   1515
   End
   Begin VB.Label lblStockStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   0
      TabIndex        =   69
      Top             =   3000
      Width           =   2235
   End
   Begin VB.Label generalLabel 
      Caption         =   "Jump To Line:"
      Height          =   285
      Index           =   13
      Left            =   60
      TabIndex        =   7
      Top             =   30
      Width           =   1035
   End
   Begin VB.Label generalLabel 
      Caption         =   "Jump To Item:"
      Height          =   255
      Index           =   12
      Left            =   1980
      TabIndex        =   6
      Top             =   30
      Width           =   1005
   End
   Begin VB.Label lblMaxAmt 
      Caption         =   "of MAX"
      Height          =   195
      Left            =   4560
      TabIndex        =   4
      Top             =   9090
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   13020
      Y1              =   8970
      Y2              =   8970
   End
   Begin VB.Label lblStatusBar 
      Height          =   345
      Left            =   5430
      TabIndex        =   3
      Top             =   9000
      Width           =   2175
   End
End
Attribute VB_Name = "SignMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ADDING A NEW EXT SPECS TAB
'   add field to PartNumbers & trigger
'   add to spPartNumbersByItem
'   add to vWebOffload
'   add to switch in SignMaintenance.barPosition_Change
'   add to switch in SignMaintenance.ExtendedSpecs_LostFocus
'   add to SignMaintenance.fillForm
'   add to WebOffload.webOffloadSingleItem

Public sphereOneFormOpen As Boolean

Private ItemList() As String
Private changed As Boolean
Private fillingForm As Boolean
Private firstLoad As Boolean
Private whichCtl As String
Private stopLineLostFocus As Boolean
Private lockedDown As Boolean

Private backButtonActive As Boolean
Private backButtonHistory As PerlArray
Private backButtonPointer As Long

Private shippingExtraIndexToIDMap As Dictionary

Private Const XSELL_EXPAND_BY = 2715

Private WithEvents MAS200IE As Mas200ImportExport
Attribute MAS200IE.VB_VarHelpID = -1
Public CurrentOffset As Long
Public CurrentStartRecord As Long
Public CurrentEndRecord As Long


Public IHATEVISUALBASIC As Boolean






Private Sub statusChangeLog(item As String, changefield As String, changedesc As String)
    DB.Execute "INSERT INTO PartNumbersStatusChangeLog ( ItemNumber, ChangedBy, ChangedField, ChangeDesc ) VALUES ( '" & item & "', " & GetCurrentUserID() & ", '" & EscapeSQuotes(changefield) & "', '" & EscapeSQuotes(changedesc) & "' )"
End Sub









Public Function EBayGetAvailableQty(item As String, fudgeFactor As Integer) As Integer
    'Dim qry As ADODB.Recordset
    'Set qry = DB.retrieve("SELECT PartNumbers.EBayAvailableQty,QtyOnSo,QtyOnBO FROM vActualWhseQty INNER JOIN PartNumbers ON PartNumbers.ItemNumber = vActualWhseQty.ItemNumber WHERE vActualWhseQty.ItemNumber='" & item & "'")
    'If qry.RecordCount > 0 Then
    ' EBayGetAvailableQty = (CInt(qry("EBayAvailableQty")) - CInt(qry("QtyOnSo")) - CInt(qry("QtyOnBO"))) - fudgeFactor
    'Else
    ' EBayGetAvailableQty = -99999
    'End If
    Dim qry As ADODB.Recordset
    Set qry = DB.retrieve("SELECT QtyOnHand,QtyOnSO,QtyOnPO,QtyOnBO FROM vActualWhseQty WHERE ItemNumber='" + Me.ItemNumber.Text + "'")
    If (qry.RecordCount > 0) Then
        'EBayGetAvailableQty = (CInt(qry("QtyOnHand")) - CInt(qry("QtyOnSO")) - CInt(qty("QtyOnBO")))
        EBayGetAvailableQty = (CInt(qry("QtyOnHand")) - CInt(qry("QtyOnSo")) - CInt(qry("QtyOnBO"))) - fudgeFactor
        
    Else
        EBayGetAvailableQty = -99999
    End If
    
End Function
Public Function EBayIsCurrentItemListedOutOfStock() As Boolean
'Dim aQty As Integer

  Dim aQty As Integer
  aQty = EBayGetAvailableQty(Me.ItemNumber.Text, CInt(Me.EBayAvailabilityLimit.Text))
  If aQty <> -99999 Then
    If aQty <= 0 Then 'And EBayEBayStatus.Text = "currently: UP" Then
        'make sure item is not dc/ed or being dc/ed with 0 qty
        Dim isDC As ADODB.Recordset
        Set isDC = DB.retrieve("SELECT ItemStatus FROM InventoryMaster WHERE ItemNumber='" + Me.ItemNumber.Text + "'")
        If (isDC.RecordCount > 0) Then
        If (isDC("ItemStatus") = 1) Or (isDC("ItemStatus") = 2) Or (isDC("ItemStatus") = 9 And aQty = 0) Then
            EBayIsCurrentItemListedOutOfStock = False
        Else
            EBayIsCurrentItemListedOutOfStock = True
        End If
        End If
      
    Else
      EBayIsCurrentItemListedOutOfStock = False
    End If
  Else
    EBayIsCurrentItemListedOutOfStock = False
  End If
End Function








Public Sub btnGoToTPP_Click()
Dim url As String
    url = "http://www.toolpartsstore.com/"
    If Me.chkWebPointer Then
        url = url & DLookup("GroupPageName", "PartNumberGroupMaster", "ItemNumberPointer='" & EscapeSQuotes(Me.ItemNumber) & "'")
    Else
        url = url & LCase(CreateYahooID(Me.ItemNumber))
    End If
    url = url & ".html"
    OpenDefaultApp url
End Sub

Private Sub btnIncludesMassAdd_Click()
    Load SignsIncludesAdd
    SignsIncludesAdd.Show MODAL
    requeryIncludesList Me.ItemNumber
End Sub

Private Sub btnIncludesPNAdd_Click()
    Dim cnt As Long, i As Long
    cnt = 0
    For i = 1 To 8
        If Me.Spec(i) <> "" Then
            cnt = cnt + 1
        End If
    Next i
    If cnt < 3 Then
        MsgBox "Less than 3 specs, not adding!"
    Else
        Dim ln As String
        ln = DLookup("ManufFullNameWeb", "ProductLine", "ProductLine='" & Left(Me.ItemNumber, 3) & "'") & " " & Mid(Me.ItemNumber, 4) & " " & Me.Desc(3)
        If Not InList(ln, Me.IncludesList, False) Then
            If Me.IncludesList.ListCount > 0 Then
                DB.Execute "UPDATE PartNumberIncludes SET SortOrder=SortOrder+1 WHERE ItemNumber='" & Me.ItemNumber & "' AND SortOrder>=0"
            End If
            DB.Execute "INSERT INTO PartNumberIncludes ( ItemNumber, IncludeText, SortOrder ) VALUES ( '" & Me.ItemNumber & "', '" & EscapeSQuotes(ln) & "', 0 )", True
            Me.IncludesList.AddItem ln, 0
        End If
    End If
End Sub

Private Sub btnKirkUnlockSpecs_Click()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("exec spPartNumbersByItem '" & Me.ItemNumber & "'")
    fillSpecs rst, True
    Me.extendedSpecsFrame.Enabled = True
    Me.ExtendedSpecs(0) = Nz(rst("ExtendedSpecsHTML"))
    Me.ExtendedSpecs(1) = Nz(rst("WriteupHTML"))
    Me.ExtendedSpecs(2) = Nz(rst("FeaturesHTML"))
    Me.ExtendedSpecs(3) = Nz(rst("TechSpecsHTML"))
    Me.ExtendedSpecs(4) = Nz(rst("MediaHTML"))
    Me.ExtendedSpecs(5) = Nz(rst("NotesHTML"))
    'Me.ExtendedSpecs(6) = Nz(rst("QuestionsHTML"))
    Dim i As Long
    For i = 0 To Me.ExtendedSpecs.UBound
        Me.chkTabStrip(i).FontBold = Len(Me.ExtendedSpecs(i)) > 0
    Next i
    'If Me.ExtendedSpecsPreview.Visible Then
    If Me.chkTabStrip(Me.chkTabStrip.UBound) Then
        loadExtendedSpecsWebpage rst("ItemNumber")
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub btnNextItem_Click()
If Me.CurrentOffset < Me.CurrentEndRecord Then
    Me.CurrentOffset = Me.CurrentOffset + 1
    barPosition_Change
End If
End Sub

Private Sub btnPrevItem_Click()
If Me.CurrentOffset > Me.CurrentStartRecord Then
    Me.CurrentOffset = Me.CurrentOffset - 1
    barPosition_Change
End If
End Sub

Private Sub btnPrintScreen_Click()
    ClipBoard_SetImage Utilities.CaptureVB6Form(Me)
    MsgBox "Screenshot captured!"
End Sub

Private Sub chkTabStrip_Click(index As Integer)
    'why is this not an option group? maybe i'm dumb.
    If Not fillingForm Then
        fillingForm = True
        If Me.chkTabStrip(index) Then
            'clicked, flip the others off and show correct textarea
            Dim i As Long
            For i = 0 To Me.chkTabStrip.UBound
                If i = index Then
                    If i = Me.chkTabStrip.UBound Then
                        loadExtendedSpecsWebpage (Me.ItemNumber)
                        Me.ExtendedSpecsPreview.Visible = True
                    Else
                        Me.ExtendedSpecs(i).Visible = True
                        Me.ExtendedSpecs(i).SetFocus
                        HandleTextBoxCursorPosition Me.ExtendedSpecs(i)
                    End If
                Else
                    If Me.chkTabStrip(i) Then
                        Me.chkTabStrip(i) = 0
                        If i = Me.chkTabStrip.UBound Then
                            Me.ExtendedSpecsPreview.Visible = False
                        Else
                            Me.ExtendedSpecs(i).Visible = False
                        End If
                    End If
                End If
            Next i
        Else
            'not unclickable
            Me.chkTabStrip(index) = 1
        End If
        fillingForm = False
    End If
End Sub

'******************************************************************************
'   EVENT HANDLERS
'******************************************************************************
Private Sub MAS200IE_StatusChanged(status As String)
    Me.lblStatusBar.Caption = status
    DoEvents
End Sub



'******************************************************************************
'  PUBLIC FUNCTIONS
'******************************************************************************
Public Sub FilterByForm(filtersql As String)
    Me.btnFilterByForm.tag = filtersql
    requeryForm "FBF"
    'If Me.barPosition.value = 0 Then
    If Me.CurrentOffset = 0 Then
        barPosition_Change
    Else
        'Me.barPosition.value = 0
        Me.CurrentOffset = 0
    End If
    Me.jumpToLine.SetFocus
End Sub

'Public Sub FilterBySearch(filtersql As String)
'    Me.btnSearch.tag = filtersql
'    requeryForm "FBS"
'    If Me.barPosition.value = 0 Then
'        barPosition_Change
'    Else
'        Me.barPosition.value = 0
'    End If
'    Me.jumpToLine.SetFocus
'End Sub

Public Sub SetFilter(filtername As String)
    'If InCombo(filtername, Me.cmbFilters) Then
        Me.cmbFilters = filtername
        requeryForm filtername
        'If Me.barPosition.value = 0 Then
        If Me.CurrentOffset = 0 Then
            barPosition_Change
        Else
            'Me.barPosition.value = 0
            Me.CurrentOffset = 0
        End If
        'Me.jumpToLine.SetFocus
    'End If
End Sub

Public Function GetFilter() As String
    If Me.cmbFilters = "Filter By Form" Then
        GetFilter = "FBF-" & Me.btnFilterByForm.tag
'    ElseIf Me.cmbFilters = "Filter By Search" Then
'        GetFilter = "FBS-" & Me.btnSearch.tag
    Else
        GetFilter = Me.cmbFilters
    End If
End Function

Public Sub GoToItem(item As String, Optional RemoveFilterIfNecessary As Boolean = False)
    If Me.ItemNumber = item Then
        Exit Sub
    End If
    
    Mouse.Hourglass True
    Dim pos As Long
    If Left(item, 3) <> "XXX" And Left(item, 3) <> "ZZZ" Then
        If Not ExistsInPartNumbers(item) Then
            If MsgBox("Item doesn't exist in Signs, do you want to create it?", vbYesNo) = vbYes Then
                DB.Execute "INSERT INTO PartNumbers ( ItemNumber ) VALUES ( '" & item & "' )"
                requeryForm "All"
                'Me.barPosition.value = bsearch(item, ItemList)
                Me.CurrentOffset = bsearch(item, ItemList)
                barPosition_Change
            End If
        Else
            pos = bsearch(item, ItemList)
            If pos = -1 Then
                If RemoveFilterIfNecessary Then
                    requeryForm "All"
                    pos = bsearch(item, ItemList)
                    If pos = -1 Then
                        'ok, maybe not. let's just give up here
                    Else
                        'Me.barPosition.value = pos
                        Me.CurrentOffset = pos
                        barPosition_Change
                    End If
                End If
            Else
                'Me.barPosition.value = pos
                Me.CurrentOffset = pos
                barPosition_Change
            End If
        End If
    End If
    Mouse.Hourglass False
End Sub

Public Sub RequerySignsInPlace()
    Dim item As String
    'If Me.barPosition.value = Me.barPosition.Min Then
    If Me.CurrentOffset = Me.CurrentStartRecord Then
        'item = ItemList(Me.barPosition.value + 1)
        item = ItemList(Me.CurrentOffset + 1)
    Else
        'item = ItemList(Me.barPosition.value - 1)
        item = ItemList(Me.CurrentOffset - 1)
    End If
    requeryForm Me.cmbFilters
    'Me.barPosition.value = bsearch(item, ItemList)
    Me.CurrentOffset = bsearch(item, ItemList)
End Sub

Public Sub refreshForm(Optional updateOther As Boolean = True)
    Dim item As ADODB.Recordset
    Set item = DB.retrieve("exec spPartNumbersByItem '" & Me.ItemNumber & "'")
    fillForm item
    item.Close
    Set item = Nothing
    If updateOther Then
        If IsFormLoaded("InventoryMaintenance") Then
            InventoryMaintenance.refreshForm False
        End If
    End If
End Sub

Public Function GetItemList() As String()
    GetItemList = ItemList
End Function

Public Function CurrentWebsiteID() As Long
    CurrentWebsiteID = CLng(Me.Website.tag)
End Function



'******************************************************************************
'  PRIVATE FUNCTIONS - three sections
'     - button + misc event handlers
'     - functions related to the form, mostly filling, etc
'     - data validation + saving
'******************************************************************************
Private Sub barPosition_Change()
    Mouse.Hourglass True
    If changed = True Then
        Select Case whichCtl
            Case Is = "Availability"
                Availability_LostFocus
            Case Is = "Desc1"
                Desc_LostFocus 1
            Case Is = "Desc2"
                Desc_LostFocus 2
            Case Is = "Desc3"
                Desc_LostFocus 3
            Case Is = "Desc4"
                Desc_LostFocus 4
            Case Is = "Spec1"
                Spec_LostFocus 1
            Case Is = "Spec2"
                Spec_LostFocus 2
            Case Is = "Spec3"
                Spec_LostFocus 3
            Case Is = "Spec4"
                Spec_LostFocus 4
            Case Is = "Spec5"
                Spec_LostFocus 5
            Case Is = "Spec6"
                Spec_LostFocus 6
            Case Is = "Spec7"
                Spec_LostFocus 7
            Case Is = "Spec8"
                Spec_LostFocus 8
            Case Is = "AvailabilityLimit"
                AvailabilityLimit_LostFocus
            Case Is = "EBayAvailabilityLimit"
                EBayAvailabilityLimit_LostFocus
            Case Is = "ExtendedSpecs0"
                ExtendedSpecs_LostFocus 0
            Case Is = "ExtendedSpecs1"
                ExtendedSpecs_LostFocus 1
            Case Is = "ExtendedSpecs2"
                ExtendedSpecs_LostFocus 2
            Case Is = "ExtendedSpecs3"
                ExtendedSpecs_LostFocus 3
            Case Is = "ExtendedSpecs4"
                ExtendedSpecs_LostFocus 4
            Case Is = "ExtendedSpecs5"
                ExtendedSpecs_LostFocus 5
            'Case Is = "ExtendedSpecs6"
            '    ExtendedSpecs_LostFocus 6
            Case Is = "TriadCode"
                TriadCode_LostFocus
            Case Is = "txtPosition"
                txtPosition_LostFocus
            Case Is = "InternalNotes"
                InternalNotes_LostFocus
            Case Else
                MsgBox "Could not determine which control you were updating, so not saving."
        End Select
    End If
    'If ItemList(Me.barPosition.value) = "NORECORDS" Then
    If ItemList(Me.CurrentOffset) = "NORECORDS" Then
        MsgBox "No records found."
    Else
        Dim rst As ADODB.Recordset
        'Set rst = DB.retrieve("exec spPartNumbersByItem '" & ItemList(Me.barPosition.value) & "'")
        Set rst = DB.retrieve("exec spPartNumbersByItem '" & ItemList(Me.CurrentOffset) & "'")
        If rst.EOF Then
            'MsgBox "Record " & ItemList(Me.barPosition.value) & " deleted by another user? requerying..."
            MsgBox "Record " & ItemList(Me.CurrentOffset) & " deleted by another user? requerying..."
            Dim item As String
            'If Me.barPosition.value = Me.barPosition.Min Then
            If Me.CurrentOffset = Me.CurrentStartRecord Then
                'item = ItemList(Me.barPosition.value + 1)
                item = ItemList(Me.CurrentOffset + 1)
            Else
                'item = ItemList(Me.barPosition.value - 1)
                item = ItemList(Me.CurrentOffset + 1)
            End If
            requeryForm Me.cmbFilters
            'Me.barPosition.value = bsearch(item, ItemList)
            Me.CurrentOffset = bsearch(item, ItemList)
            Mouse.Hourglass False
            Exit Sub
        End If
        fillForm rst
        rst.Close
        Set rst = Nothing
        'Me.txtPosition = Me.barPosition.value + 1
        Me.txtPosition = CStr(CLng(Me.CurrentOffset + 1))
        changed = False
        If Not firstLoad Then
            If IsFormLoaded("InventoryMaintenance") Then
                InventoryMaintenance.GoToItem Me.ItemNumber
            End If
        End If
        If Not Me.ActiveControl Is Nothing Then
            If TypeOf Me.ActiveControl Is TextBox Then
                HandleTextBoxCursorPosition Me.ActiveControl
            ElseIf TypeOf Me.ActiveControl Is ComboBox Then
                HandleComboBoxCursorPosition Me.ActiveControl
            End If
        End If
    End If
    Mouse.Hourglass False
End Sub

'Private Sub btnAccessoriesAddToFilter_Click()
'    Mouse.Hourglass True
'    Dim itemtocopy As String
'    itemtocopy = ClipBoard_GetData()
'    If InStr(itemtocopy, "Accessories for ") Then
'        itemtocopy = Replace(itemtocopy, "Accessories for ", "")
'        If ExistsInPartNumbers(itemtocopy) Then
'            Dim rst As ADODB.Recordset
'            Set rst = retrieve("SELECT Accessory FROM PartNumberAccessories WHERE ItemNumber='" & itemtocopy & "'")
'            Dim i As Integer
'            Dim rstLookup As ADODB.Recordset
'            Dim accH As Dictionary
'            For i = 0 To UBound(itemList)
'                rst.MoveFirst
'                Set rstLookup = retrieve("SELECT Accessory FROM PartNumberAccessories WHERE ItemNumber='" & itemList(i) & "'")
'                Set accH = New Dictionary
'                While Not rstLookup.EOF
'                    accH.Add rstLookup("Accessory").value, "1"
'                    rstLookup.MoveNext
'                Wend
'                rstLookup.Close
'                Set rstLookup = Nothing
'                While Not rst.EOF
'                    If Not accH.exists(rst("Accessory").value) Then
'                        If rst("Accessory") <> itemList(i) Then
'                            dbExecute "INSERT INTO PartNumberAccessories ( ItemNumber, Accessory ) VALUES ( '" & itemList(i) & "', '" & rst("Accessory") & "' )"
'                        Else
'                            'MsgBox "Not adding item " & itemList(i) & " as an accessory to itself. If you really want to add it, do it manually."
'                        End If
'                    End If
'                    rst.MoveNext
'                Wend
'                Set accH = Nothing
'            Next i
'            rst.Close
'            Set rst = Nothing
'            requeryAccessories Me.ItemNumber
'        Else
'            MsgBox "Invalid clipboard contents."
'        End If
'    Else
'        MsgBox "Invalid clipboard contents."
'    End If
'    Mouse.Hourglass False
'End Sub

Private Sub btnAccessoriesCopy_Click()
    ClipBoard_SetData ("Accessories for " & Me.ItemNumber)
End Sub

Private Sub btnAccessoriesPaste_Click()
    Mouse.Hourglass True
    Dim itemtocopy As String
    itemtocopy = ClipBoard_GetData()
    If InStr(itemtocopy, "Accessories for ") Then
        itemtocopy = Replace(itemtocopy, "Accessories for ", "")
        If ExistsInPartNumbers(itemtocopy) Then
            Dim rst As ADODB.Recordset
            Set rst = DB.retrieve("SELECT Accessory FROM PartNumberAccessories WHERE ItemNumber='" & itemtocopy & "'")
            While Not rst.EOF
                If rst("Accessory") = Me.ItemNumber Then
                    DB.Execute "INSERT INTO PartNumberAccessories ( ItemNumber, Accessory ) VALUES ( '" & Me.ItemNumber & "', '" & itemtocopy & "' )", True
                Else
                    DB.Execute "INSERT INTO PartNumberAccessories ( ItemNumber, Accessory ) VALUES ( '" & Me.ItemNumber & "', '" & rst("Accessory") & "' )", True
                End If
                rst.MoveNext
            Wend
            rst.Close
            Set rst = Nothing
            requeryAccessories Me.ItemNumber
        Else
            MsgBox "Invalid clipboard contents."
        End If
    Else
        MsgBox "Invalid clipboard contents."
    End If
    Mouse.Hourglass False
End Sub

Private Sub btnAddAccessories_Click()
    Load AddCrossSells
    AddCrossSells.Show
End Sub

'Private Sub btnAddAccessoriesReverse_Click()
'    Load AddAccessories
'    AddAccessories.chkReverseAccessories = 1
'    AddAccessories.Show MODAL
'    requeryAccessories Me.ItemNumber
'End Sub

'Private Sub btnAddWebPaths_Click(Index As Integer)
'    Load AddItemWebPaths
'    AddItemWebPaths.dx = Index
'    AddItemWebPaths.realFormLoad
'    AddItemWebPaths.Show MODAL
'    requeryWebPaths Me.ItemNumber
'End Sub

Private Sub btnApplyTriadCodeToFilter_Click()
    If MsgBox("Apply Triad Code " & qq(Me.TriadCode) & " to all items in current filter?", vbYesNo) = vbYes Then
        Dim i As Long
        For i = 0 To UBound(ItemList)
            DB.Execute "UPDATE InventoryMaster SET TriadCode='" & EscapeSQuotes(Me.TriadCode) & "' WHERE ItemNumber='" & ItemList(i) & "'"
        Next i
    End If
End Sub

'Private Sub btnBaseItemEdit_Click()
'    Load AddBaseItem
'    AddBaseItem.BaseItem = Me.baseItemList
'    AddBaseItem.Show MODAL
'    RefreshForm
'End Sub

'Private Sub btnBaseItemNew_Click()
'    Dim newitem As String
'    newitem = Me.ItemNumber
'    newitem = InputBox("Enter the new base item #:", , newitem)
'    If newitem = "" Then
'        Exit Sub
'    End If
'    If Not validateItemNumber(newitem) Then
'        MsgBox "invalid base item number"
'        Exit Sub
'    End If
'    If DLookup("COUNT(*)", "BaseItemList", "BaseItem='" & newitem & "'") <> 0 Then
'        MsgBox "This base item already exists"
'        Exit Sub
'    End If
'    Dim exists As Boolean
'    If ExistsInInventoryMaster(newitem) Then
'        If MsgBox("This item already exists, are you sure?", vbYesNo) = vbNo Then
'            Exit Sub
'        End If
'        exists = True
'    End If
'    'ok, looks good
'    If Not exists Then
'        CreateItem newitem, CLng(Me.Website.tag)
'    End If
'    updateInventoryMaster "ItemDescription", "BASE ITEM", newitem, "'"
'    DB.Execute "INSERT INTO BaseItemList ( BaseItem ) VALUES ( '" & newitem & "' )"
'    requeryBaseItemList
'    Me.baseItemList = newitem
'    btnBaseItemEdit_Click
'End Sub

Private Sub btnClearTriadCode_Click()
    If MsgBox("Remove Triad Code " & qq(Me.TriadCode) & " from any items that have it?", vbYesNo) = vbYes Then
        DB.Execute "UPDATE InventoryMaster SET TriadCode=NULL WHERE TriadCode='" & EscapeSQuotes(Me.TriadCode) & "'"
        Me.TriadCode = ""
    End If
End Sub

Private Sub btnCrossSellAdd_Click()
    Dim url As String, txt As String
    url = InputBox("Enter the URL to point to:")
    If url = "" Then
        Exit Sub
    End If
    If InStr(url, "/") Then
        url = Mid(url, InStrRev(url, "/") + 1)
    End If
    txt = InputBox("Enter text to show:")
    If txt = "" Then
        Exit Sub
    End If
    Dim stupidhack As Boolean
    If Not Me.CrossSellList.InList(txt) Then
        Dim sortNum As String
        If Me.CrossSellList.ListCount = 0 Then
            sortNum = -1
        ElseIf Me.CrossSellList.SelIndex = -1 Then
            sortNum = Me.CrossSellList.ListCount - 1
        Else
            sortNum = Me.CrossSellList.SelIndex - 1
            If sortNum = -1 Then
                stupidhack = True
            End If
            DB.Execute "UPDATE PartNumberCrossSell SET SortOrder=SortOrder+1 WHERE ItemNumber='" & Me.ItemNumber & "' AND SortOrder>" & sortNum
        End If
        DB.Execute "INSERT INTO PartNumberCrossSell ( ItemNumber, ShowText, URL, SortOrder ) VALUES ( '" & Me.ItemNumber & "', '" & EscapeSQuotes(txt) & "', '" & EscapeSQuotes(url) & "', " & sortNum + 1 & " )"
        Me.CrossSellList.Add Array(txt, url), CLng(sortNum)
        If stupidhack Then
            requeryCrossSellList Me.ItemNumber
        End If
    Else
        MsgBox "That link already exists!"
    End If
End Sub

Private Sub btnCrossSellDelete_Click()
    If Me.CrossSellList.SelIndex <> -1 Then
        Dim this As String
        this = Me.CrossSellList.GetTextSelected(0)
        If vbYes = MsgBox("Remove " & qq(this) & "?", vbYesNo) Then
            Dim sortNum As String
            sortNum = DLookup("SortOrder", "PartNumberCrossSell", "ItemNumber='" & Me.ItemNumber & "' AND ShowText='" & EscapeSQuotes(this) & "'")
            DB.Execute "DELETE FROM PartNumberCrossSell WHERE ItemNumber='" & Me.ItemNumber & "' AND ShowText='" & EscapeSQuotes(this) & "'"
            DB.Execute "UPDATE PartNumberCrossSell SET SortOrder=SortOrder-1 WHERE ItemNumber='" & Me.ItemNumber & "' AND SortOrder>" & sortNum
            Me.CrossSellList.RemoveSelected
        End If
    End If
End Sub

Private Sub btnCrossSellMoveDown_Click()
    If Me.CrossSellList.SelIndex <> -1 And Me.CrossSellList.SelIndex <> Me.CrossSellList.ListCount - 1 Then
        Dim this() As Variant, nextone() As Variant
        this = Me.CrossSellList.GetRowSelected()
        nextone = Me.CrossSellList.GetRow(Me.CrossSellList.SelIndex + 1)
        DB.Execute "UPDATE PartNumberCrossSell SET SortOrder=SortOrder+1 WHERE ItemNumber='" & Me.ItemNumber & "' AND ShowText='" & EscapeSQuotes(CStr(this(0))) & "'"
        DB.Execute "UPDATE PartNumberCrossSell SET SortOrder=SortOrder-1 WHERE ItemNumber='" & Me.ItemNumber & "' AND ShowText='" & EscapeSQuotes(CStr(nextone(0))) & "'"
        Me.CrossSellList.Edit CStr(nextone(0)), Me.CrossSellList.SelIndex, 0
        Me.CrossSellList.Edit CStr(nextone(1)), Me.CrossSellList.SelIndex, 1
        Me.CrossSellList.Edit CStr(this(0)), Me.CrossSellList.SelIndex + 1, 0
        Me.CrossSellList.Edit CStr(this(1)), Me.CrossSellList.SelIndex + 1, 1
        Me.CrossSellList.SelectRow Me.CrossSellList.SelIndex + 1
    End If
End Sub

Private Sub btnCrossSellMoveUp_Click()
    If Me.CrossSellList.SelIndex <> -1 And Me.CrossSellList.SelIndex <> 0 Then
        'Dim this(1) As String, prev(1) As String
        Dim this() As Variant, prev() As Variant
        this = Me.CrossSellList.GetRowSelected()
        prev = Me.CrossSellList.GetRow(Me.CrossSellList.SelIndex - 1)
        DB.Execute "UPDATE PartNumberCrossSell SET SortOrder=SortOrder-1 WHERE ItemNumber='" & Me.ItemNumber & "' AND ShowText='" & EscapeSQuotes(CStr(this(0))) & "'"
        DB.Execute "UPDATE PartNumberCrossSell SET SortOrder=SortOrder+1 WHERE ItemNumber='" & Me.ItemNumber & "' AND ShowText='" & EscapeSQuotes(CStr(prev(0))) & "'"
        Me.CrossSellList.Edit CStr(prev(0)), Me.CrossSellList.SelIndex, 0
        Me.CrossSellList.Edit CStr(prev(1)), Me.CrossSellList.SelIndex, 1
        Me.CrossSellList.Edit CStr(this(0)), Me.CrossSellList.SelIndex - 1, 0
        Me.CrossSellList.Edit CStr(this(1)), Me.CrossSellList.SelIndex - 1, 1
        Me.CrossSellList.SelectRow Me.CrossSellList.SelIndex - 1, True, False, False
    End If
End Sub

Private Sub btnCrossSellPush_Click()
    'TODO
    MsgBox "Not implemented"
End Sub

Private Sub btnCrossSellSuggest_Click()
    
    MsgBox "TODO: BROKEN! need to fix paths!"
    Exit Sub
    
    Mouse.Hourglass True
    Dim mfgrQuery As String, pathQuery As String
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT WebPathName, PathLevel FROM vPNWebPaths WHERE PathLevel<=3 AND ItemNumber='" & Me.ItemNumber & "'")
    mfgrQuery = "SELECT DISTINCT ItemNumber FROM vCrossSellItemInfo WHERE "
    pathQuery = "SELECT DISTINCT ItemNumber FROM vCrossSellItemInfo WHERE "
    mfgrQuery = mfgrQuery & "ProductLine='" & Left(Me.ItemNumber, 3) & "' AND "
    While Not rst.EOF
        If rst("PathLevel") <> 1 Then
            mfgrQuery = mfgrQuery & "ItemNumber IN (SELECT ItemNumber FROM vPNWebPaths WHERE WebPathName='" & EscapeSQuotes(rst("WebPathName")) & "' AND PathLevel=" & rst("PathLevel") & ") AND "
        End If
        pathQuery = pathQuery & "ItemNumber IN (SELECT ItemNumber FROM vPNWebPaths WHERE WebPathName='" & EscapeSQuotes(rst("WebPathName")) & "' AND PathLevel=" & rst("PathLevel") & ") AND "
        rst.MoveNext
    Wend
    rst.Close
    mfgrQuery = mfgrQuery & "LEN(URL)>8" 'drop the mfgr paths for now
    pathQuery = pathQuery & "LEN(URL)>8"
    
    'we've got two queries to find all the useful items for suggestions
    ' useful = same paths 2 and 3
    '          same manuf, or same path1
    '          published or t/b published
    '          have something in cross sell already, but not a manuf link (8 char url "foo.html")
    'do the manuf one first, since that's hopefully more relevant
    'if we can't get any there, do the path one
    Dim continuewith As String
    Set rst = DB.retrieve(mfgrQuery)
    If Not rst.EOF Then
        'ok with this query
        continuewith = "mfgr"
    Else
        rst.Close
        Set rst = DB.retrieve(pathQuery)
        If Not rst.EOF And rst.RecordCount >= 2 Then
            'ok with this query
            continuewith = "path"
        Else
            MsgBox "Can't find enough similar items to make a suggestion" & vbCrLf & vbCrLf & "Fill this one out manually, and maybe the next one will work"
            continuewith = ""
        End If
    End If
    
    Dim countitems As Long
    countitems = rst.RecordCount
    rst.Close
    
    Dim Results As PerlArray
    Set Results = New PerlArray
    
    If continuewith = "" Then
        Mouse.Hourglass False
        Exit Sub
    End If
    
    'now, we use that original query as a subquery, to get grouped info about cross sells
    Dim gpsql As String
    gpsql = "SELECT ShowText, URL, COUNT(*) AS NumOccurrences FROM PartNumberCrossSell WHERE ItemNumber IN ("
    Select Case continuewith
        Case Is = "mfgr"
            gpsql = gpsql & mfgrQuery
        Case Is = "path"
            gpsql = gpsql & pathQuery
    End Select
    gpsql = gpsql & ") AND LEN(URL)>8 GROUP BY ShowText, URL ORDER BY NumOccurrences DESC"
    Set rst = DB.retrieve(gpsql)
    'ok, now we've got the text/url and # of times that appears in similar items.
    'grab a few good ones, and throw them in a perlarray
    'kind of a funny loop, but it'll be easier to change the parameters later
    Dim textSeen As Dictionary
    Set textSeen = New Dictionary
    Dim stopNow As Boolean
    While Not stopNow
        If textSeen.exists(CStr(LCase(rst("ShowText")))) Then
            'skip
        Else
            textSeen.Add CStr(LCase(rst("ShowText"))), 1
            Results.Push Array(CStr(rst("ShowText")), CStr(rst("URL")))
        End If
        
        rst.MoveNext
        
        If rst.EOF Then                                         'that's all, folks
            stopNow = True
        ElseIf rst("NumOccurrences") / countitems < 0.3 Then    'below 30% = bail
            stopNow = True
        ElseIf Results.Scalar > 5 Then                          '5 is enough, right?
            stopNow = True
        End If
    Wend
    rst.Close
    
    'ok, now we have to reorder them. how?
    'TODO
    'Dim orderedResults As PerlArray
    'Set orderedResults = New PerlArray
    Dim i As Long
    'For i = 0 To Results.Scalar - 1
    '
    'Next i
    'Set Results = orderedResults
    'Set orderedResults = Nothing
    
    
    Set rst = DB.retrieve("SELECT TOP 1 ShowText, URL, COUNT(*) AS NumOccurrences FROM vCrossSellItemInfo WHERE ProductLine='" & Left(Me.ItemNumber, 3) & "' AND LEN(URL)=8 GROUP BY ShowText, URL ORDER BY NumOccurrences DESC")
    If rst.EOF Then
        'no manuf links found for line?
    Else
        Results.Push Array(CStr(rst("ShowText")), CStr(rst("URL")))
    End If
    rst.Close
    
    Dim msg As String
    msg = "Remove all existing cross sell links and add the following?" & vbCrLf & vbCrLf
    For i = 0 To Results.Scalar - 1
        msg = msg & Results.Elem(i)(0) & vbCrLf & "     " & Results.Elem(i)(1) & vbCrLf & vbCrLf
    Next i
    If vbYes = MsgBox(msg, vbYesNo) Then
        DB.Execute "DELETE FROM PartNumberCrossSell WHERE ItemNumber='" & Me.ItemNumber & "'"
        For i = 0 To Results.Scalar - 1
            DB.Execute "INSERT INTO PartNumberCrossSell ( ItemNumber, ShowText, URL, SortOrder ) VALUES ( '" & Me.ItemNumber & "', '" & EscapeSQuotes(CStr(Results.Elem(i)(0))) & "', '" & EscapeSQuotes(CStr(Results.Elem(i)(1))) & "', " & i & " )"
        Next i
        requeryCrossSellList Me.ItemNumber
    End If
    
    Mouse.Hourglass False
End Sub

Private Sub btnCrossSellCopy_Click()
    ClipBoard_SetData "CrossSell for " & Me.ItemNumber
End Sub

Private Sub btnCrossSellPaste_Click()
    Mouse.Hourglass True
    Dim itemtocopy As String
    itemtocopy = ClipBoard_GetData()
    If InStr(itemtocopy, "CrossSell for ") Then
        itemtocopy = Replace(itemtocopy, "CrossSell for ", "")
        If ExistsInPartNumbers(itemtocopy) Then
            DB.Execute "DELETE FROM PartNumberCrossSell WHERE ItemNumber='" & Me.ItemNumber & "'"
            Dim rst As ADODB.Recordset
            Set rst = DB.retrieve("SELECT ShowText, URL, SortOrder FROM PartNumberCrossSell WHERE ItemNumber='" & itemtocopy & "'")
            While Not rst.EOF
                DB.Execute "INSERT INTO PartNumberCrossSell ( ItemNumber, ShowText, URL, SortOrder ) VALUES ( '" & Me.ItemNumber & "', '" & EscapeSQuotes(rst("ShowText")) & "', '" & EscapeSQuotes(rst("URL")) & "', " & rst("SortOrder") & " )"
                rst.MoveNext
            Wend
            rst.Close
            Set rst = Nothing
            requeryCrossSellList Me.ItemNumber
        Else
            MsgBox "Invalid clipboard contents."
        End If
    Else
        MsgBox "Invalid clipboard contents."
    End If
    Mouse.Hourglass False
End Sub

Private Sub CrossSellList_DblClick(i As Long, j As Long, txt As String)
    Dim newField As String, newText As String
    Select Case j
        Case Is = 0
            newField = "ShowText"
            newText = InputBox("Enter new text to show:", , txt)
        Case Is = 1
            newField = "URL"
            newText = InputBox("Enter new URL:", , txt)
        Case Else
            MsgBox "Whoa, didn't expect that. You clicked on a column I don't know about."
            Exit Sub
    End Select
    If newText = "" Then
        Exit Sub
    End If
    If newField = txt Then
        Exit Sub
    End If
    If newField = "URL" Then
        If InStr(newText, "/") Then
            newText = Mid(newText, InStrRev(newText, "/") + 1)
        End If
    End If
    DB.Execute "UPDATE PartNumberCrossSell SET " & newField & "='" & EscapeSQuotes(newText) & "' WHERE ItemNumber='" & Me.ItemNumber & "' AND ShowText='" & EscapeSQuotes(Me.CrossSellList.GetTextSelected) & "'"
    Me.CrossSellList.Edit newText, i, j
End Sub

'Private Sub btnEditShipClass_Click()
'    Load AddShipClass
'    AddShipClass.Show MODAL
'End Sub

'Private Sub btnEditWebPaths_Click()
'    Load AddWebPath
'    AddWebPath.Show MODAL
'End Sub

'Private Sub btnEditWebSigns_Click()
'    Load AddWebSign
'    AddWebSign.Show MODAL
'End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnExtendedSpecsSearch_Click()
    Dim i As Long
    Dim dx As Long
    For i = 0 To Me.chkTabStrip.UBound
        If Me.chkTabStrip(i) Then
            dx = i
            Exit For
        End If
    Next i
    If dx <> Me.chkTabStrip.UBound Then
        Dim tofind As String
        tofind = LCase(InputBox("Enter search string:", , Me.btnExtendedSpecsSearch.tag))
        If tofind <> "" Then
            Dim found As Boolean
            found = False
            Me.btnExtendedSpecsSearch.tag = tofind
            Me.ExtendedSpecs(dx).SetFocus
            Dim pos As Long
            For pos = 1 To Len(Me.ExtendedSpecs(dx).Text) - Len(tofind)
                If LCase(Mid(Me.ExtendedSpecs(dx).Text, pos, Len(tofind))) = tofind Then
                    Me.ExtendedSpecs(dx).selStart = pos - 1
                    Me.ExtendedSpecs(dx).SelLength = Len(tofind)
                    found = True
                    Exit For
                End If
            Next pos
            If Not found Then
                MsgBox "Search string not found!"
            End If
        End If
    End If
End Sub

Private Sub btnExtendedSpecsSearchContinue_Click()
    Dim i As Long
    Dim dx As Long
    For i = 0 To Me.chkTabStrip.UBound
        If Me.chkTabStrip(i) Then
            dx = i
            Exit For
        End If
    Next i
    If dx <> Me.chkTabStrip.UBound Then
        If Me.btnExtendedSpecsSearch.tag <> "" Then
            Dim found As Boolean
            found = False
            Dim origpos As Long
            origpos = Me.ExtendedSpecs(dx).selStart + 2
            Dim tofind As String
            tofind = Me.btnExtendedSpecsSearch.tag
            Me.ExtendedSpecs(dx).SetFocus
            Dim pos As Long
            For pos = origpos To Len(Me.ExtendedSpecs(dx).Text) - Len(tofind)
                If LCase(Mid(Me.ExtendedSpecs(dx).Text, pos, Len(tofind))) = tofind Then
                    Me.ExtendedSpecs(dx).selStart = pos - 1
                    Me.ExtendedSpecs(dx).SelLength = Len(tofind)
                    found = True
                    Exit For
                End If
            Next pos
            If Not found Then
                For pos = 1 To origpos
                    If LCase(Mid(Me.ExtendedSpecs(dx).Text, pos, Len(tofind))) = tofind Then
                        Me.ExtendedSpecs(dx).selStart = pos - 1
                        Me.ExtendedSpecs(dx).SelLength = Len(tofind)
                        found = True
                        Exit For
                    End If
                Next pos
            End If
            If Not found Then
                MsgBox "Search string not found!"
            End If
        Else
            btnExtendedSpecsSearch_Click
        End If
    End If
End Sub

Private Sub btnExtSpecsInNotepad_Click()
    Dim dx As String
    Dim i As Long
    For i = 0 To Me.chkTabStrip.UBound
        If Me.chkTabStrip(i) Then
            dx = i
            Exit For
        End If
    Next i
    If dx <> Me.chkTabStrip.UBound Then
        Open DESTINATION_DIR & "extspec.txt" For Output As #1
        Print #1, Me.ExtendedSpecs(dx)
        Close #1
        OpenDefaultApp DESTINATION_DIR & "extspec.txt"
    End If
End Sub

Private Sub btnExtendedSpecsValidateHTML_Click()
    Dim dx As String
    Dim i As Long
    For i = 0 To Me.chkTabStrip.UBound
        If Me.chkTabStrip(i) Then
            dx = i
            Exit For
        End If
    Next i
    If dx <> Me.chkTabStrip.UBound Then
        If Len(Me.ExtendedSpecs(dx)) = 0 Then
            MsgBox "This tab is empty, douchenozzle!"
        Else
            Dim errors As Variant
            Mouse.Hourglass True
            errors = ValidateHTMLFragment(Me.ExtendedSpecs(dx))
            Mouse.Hourglass False
            If UBound(errors) = -1 Then
                MsgBox "HTML is ok!"
            Else
                MsgBox "Invalid HTML!" & vbCrLf & vbCrLf & Join(errors, vbCrLf)
            End If
        End If
    End If
End Sub

Private Sub btnFilterByForm_Click()
    Load FilterByFormDialog
    FilterByFormDialog.hiddenOwner = "SignMaintenance"
    FilterByFormDialog.Show MODAL
    If Me.btnFilterByForm.tag <> "" Then
        If IsFormLoaded("InventoryMaintenance") Then
            If MsgBox("Filter Po/Inv too?", vbYesNo) = vbYes Then
                InventoryMaintenance.FilterByForm SignMaintenance.btnFilterByForm.tag
            End If
        End If
        requeryForm "FBF"
        'If Me.barPosition.value = 0 Then
        If Me.CurrentOffset = 0 Then
            barPosition_Change
        Else
            'Me.barPosition.value = 0
            Me.CurrentOffset = 0
        End If
    End If
End Sub

Private Sub btnFirstItem_Click()
    'Me.barPosition.value = Me.barPosition.Min
    Me.CurrentOffset = Me.CurrentStartRecord
    barPosition_Change
End Sub

Private Sub btnGoToInvMaint_Click()
    If IsFormLoaded("InventoryMaintenance") Then
        If IsFormMinimized("InventoryMaintenance") Then
            UnMinimizeForm "InventoryMaintenance"
        Else
            FocusOnForm "InventoryMaintenance"
        End If
        'InventoryMaintenance.GoToItem Me.ItemNumber
    Else
        Load InventoryMaintenance
        InventoryMaintenance.GoToItem Me.ItemNumber
        InventoryMaintenance.Show
    End If
End Sub

Public Sub btnGoToWebsite_Click()
    Dim url As String
    url = "http://" & WebsiteURLHash.item(CLng(Me.Website.tag)) & "/"
    If Me.chkWebPointer Then
        url = url & DLookup("GroupPageName", "PartNumberGroupMaster", "ItemNumberPointer='" & EscapeSQuotes(Me.ItemNumber) & "'")
    Else
        url = url & LCase(CreateYahooID(Me.ItemNumber))
    End If
    url = url & ".html"
    OpenDefaultApp url
End Sub

Public Sub btnGoToEBay_Click()
    Dim id As String
    id = DLookup("EBayItemID", "PartNumbers", "ItemNumber='" & Me.ItemNumber & "'")
    If id = "" Then
        MsgBox "Item has never been listed at EBay!"
    Else
        OpenDefaultApp "http://cgi.ebay.com/ws/eBayISAPI.dll?ViewItem&item=" & id
    End If
End Sub

Private Sub btnGroupEditor_Click()
    If Not IsFormLoaded("GroupItemEditor") Then
        Load GroupItemEditor
    End If
    GroupItemEditor.GoToItem Me.ItemNumber
    If IsFormMinimized("GroupItemEditor") Then
        UnMinimizeForm "GroupItemEditor"
    End If
    If GroupItemEditor.Visible = False Then
        GroupItemEditor.Show
    End If
    FocusOnForm "GroupItemEditor"
End Sub

'Private Sub btnGroupSignFilter_Click()
'    If Me.GroupSign.ListCount = 1 Then
'        Me.GroupSign.Selected(0) = True
'    End If
'    If Me.GroupSign.SelCount <> -1 Then
'        requeryForm "Group Sign"
'        Me.barPosition.value = 0
'        barPosition_Change
'    End If
'End Sub

Private Sub btnIncludesAdd_Click()
    Dim ln As String
    ln = InputBox("Enter kit item:")
    If ln = "" Then
        Exit Sub
    End If
    ln = UnFrontpage(ln)
    If Not InList(ln, Me.IncludesList, False) Then
        Dim sortNum As String
        If Me.IncludesList.SelCount = 0 Then
            sortNum = Me.IncludesList.ListCount
        Else
            sortNum = Me.IncludesList.ListIndex
            DB.Execute "UPDATE PartNumberIncludes SET SortOrder=SortOrder+1 WHERE ItemNumber='" & Me.ItemNumber & "' AND SortOrder>=" & sortNum
        End If
        DB.Execute "INSERT INTO PartNumberIncludes ( ItemNumber, IncludeText, SortOrder ) VALUES ( '" & Me.ItemNumber & "', '" & EscapeSQuotes(ln) & "', " & sortNum & " )"
        Me.IncludesList.AddItem ln, sortNum
    Else
        MsgBox "That item is already included"
    End If
End Sub

Private Sub btnIncludesDelete_Click()
    If Me.IncludesList.SelCount <> 0 Then
        If vbYes = MsgBox("Remove " & qq(Me.IncludesList) & "?", vbYesNo) Then
            Dim sortNum As String
            sortNum = DLookup("SortOrder", "PartNumberIncludes", "ItemNumber='" & Me.ItemNumber & "' AND IncludeText='" & EscapeSQuotes(Me.IncludesList) & "'")
            DB.Execute "DELETE FROM PartNumberIncludes WHERE ItemNumber='" & Me.ItemNumber & "' AND IncludeText='" & EscapeSQuotes(Me.IncludesList) & "'"
            DB.Execute "UPDATE PartNumberIncludes SET SortOrder=SortOrder-1 WHERE ItemNumber='" & Me.ItemNumber & "' AND SortOrder>" & sortNum
            Me.IncludesList.RemoveItem Me.IncludesList.ListIndex
        End If
    End If
End Sub

Private Sub btnIncludesMoveDown_Click()
    If Me.IncludesList.SelCount <> 0 And Me.IncludesList.ListIndex <> Me.IncludesList.ListCount - 1 Then
        Dim this As String, nextone As String
        this = Me.IncludesList
        nextone = Me.IncludesList.list(Me.IncludesList.ListIndex + 1)
        DB.Execute "UPDATE PartNumberIncludes SET SortOrder=SortOrder+1 WHERE ItemNumber='" & Me.ItemNumber & "' AND IncludeText='" & EscapeSQuotes(this) & "'"
        DB.Execute "UPDATE PartNumberIncludes SET SortOrder=SortOrder-1 WHERE ItemNumber='" & Me.ItemNumber & "' AND IncludeText='" & EscapeSQuotes(nextone) & "'"
        Me.IncludesList.list(Me.IncludesList.ListIndex + 1) = this
        Me.IncludesList.list(Me.IncludesList.ListIndex) = nextone
        Me.IncludesList.ListIndex = Me.IncludesList.ListIndex + 1
    End If
End Sub

Private Sub btnIncludesMoveUp_Click()
    If Me.IncludesList.SelCount <> 0 And Me.IncludesList.ListIndex <> 0 Then
        Dim this As String, prev As String
        this = Me.IncludesList
        prev = Me.IncludesList.list(Me.IncludesList.ListIndex - 1)
        DB.Execute "UPDATE PartNumberIncludes SET SortOrder=SortOrder-1 WHERE ItemNumber='" & Me.ItemNumber & "' AND IncludeText='" & EscapeSQuotes(this) & "'"
        DB.Execute "UPDATE PartNumberIncludes SET SortOrder=SortOrder+1 WHERE ItemNumber='" & Me.ItemNumber & "' AND IncludeText='" & EscapeSQuotes(prev) & "'"
        Me.IncludesList.list(Me.IncludesList.ListIndex - 1) = this
        Me.IncludesList.list(Me.IncludesList.ListIndex) = prev
        Me.IncludesList.ListIndex = Me.IncludesList.ListIndex - 1
    End If
End Sub

Private Sub IncludesList_DblClick()
    Dim newText As String
    newText = InputBox("Enter kit item: ", , Me.IncludesList)
    If newText <> "" Then
        DB.Execute "UPDATE PartNumberIncludes SET IncludeText='" & EscapeSQuotes(newText) & "' WHERE ItemNumber='" & Me.ItemNumber & "' AND IncludeText='" & EscapeSQuotes(Me.IncludesList) & "'"
        Me.IncludesList.list(Me.IncludesList.ListIndex) = newText
    End If
End Sub

Private Sub btnLastItem_Click()
    'Me.barPosition.value = Me.barPosition.max
    Me.CurrentOffset = Me.CurrentEndRecord
    barPosition_Change
End Sub

Private Sub btnNextLC_Click()
    Dim oldLC As String
    Dim pos As Long
    Dim found As Boolean
    oldLC = Trim$(Me.jumpToLine)
    'pos = Me.barPosition.value
    pos = Me.CurrentOffset
    found = False
    While Not found
        'If pos < Me.barPosition.max Then
        If pos < Me.CurrentEndRecord Then
            If oldLC = Left$(ItemList(pos), 3) Then
                pos = pos + 1
            Else
                found = True
            End If
        Else
            found = True
        End If
    Wend
    'Me.barPosition.value = pos
    Me.CurrentOffset = pos
    barPosition_Change
End Sub

Private Sub btnPrevLC_Click()
    Mouse.Hourglass True
    Dim oldLC As String
    Dim pos As Long
    Dim found As Boolean
    oldLC = Trim$(Me.jumpToLine)
    'pos = Me.barPosition.value
    pos = Me.CurrentOffset
    found = False
    While Not found
        'If pos > Me.barPosition.Min Then
        If pos > Me.CurrentStartRecord Then
            If oldLC = Left$(ItemList(pos), 3) Then
                pos = pos - 1
            Else
                found = True
            End If
        Else
            found = True
        End If
    Wend
    Dim newLC As String
    newLC = Left$(ItemList(pos), 3)
    found = False
    While Not found
        'If pos > Me.barPosition.Min Then
        If pos > Me.CurrentStartRecord Then
            If newLC = Left$(ItemList(pos), 3) Then
                pos = pos - 1
            Else
                pos = pos + 1
                found = True
            End If
        Else
            found = True
        End If
    Wend
    'Me.barPosition.value = pos
    Me.CurrentOffset = pos
    barPosition_Change
    Mouse.Hourglass False
End Sub

Private Sub btnRefreshQtys_Click()
    'If RunningThroughVPN() Then
    '    MsgBox "You probably don't want to refresh quantities through the VPN!"
    'ElseIf IsProcessFlagSet("RefreshQtys") Then
    If IsProcessFlagSet("RefreshQtys") Then
        MsgBox "Error: the refresh quantities process is already running!"
    Else
        SetProcessFlag "RefreshQtys"
        Me.btnRefreshQtys.Enabled = False
        Mouse.Hourglass True
        Set MAS200IE = New Mas200ImportExport
        MAS200IE.RefreshQuantities
        Set MAS200IE = Nothing
        CheckForBeingDCItems
        Me.lblStatusBar.Caption = ""
        Mouse.Hourglass False
        Me.btnRefreshQtys.Enabled = True
        ClearProcessFlag "RefreshQtys"
    End If
End Sub

Private Sub btnRefreshVendorQtys_Click()
    Shell PERL & " " & REFRESH_VENDOR_QTYS, vbNormalFocus
End Sub

Private Sub btnRemoveFilter_Click()
    Dim current As String
    current = Me.ItemNumber
    Me.cmbFilters = "All"
    requeryForm "All"
    'Me.barPosition.value = bsearch(current, ItemList)
    Me.CurrentOffset = bsearch(current, ItemList)
    barPosition_Change
End Sub

'Private Sub btnSearch_Click()
'    Load SearchByPath
'    SearchByPath.Show
'End Sub

Private Sub btnOptionsDialog_Click()
    Load OptionsDialog
    OptionsDialog.Show MODAL
End Sub

Private Sub btnShippingExtrasEdit_Click()
    Load ShippingExtrasDialog
    ShippingExtrasDialog.Show MODAL
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("EXEC spPartNumbersByItem '" & Me.ItemNumber & "'")
    fillingForm = True
    requeryShippingExtras rst("PrimaryVendor"), Nz(rst("ShippingExtraID"))
    fillingForm = False
    rst.Close
    Set rst = Nothing
End Sub

Private Sub btnSpecialsChange_Click()
    Load AddSpecials
    AddSpecials.Show MODAL
    requerySpecials Me.ItemNumber
End Sub

Private Sub btnSpecialsFilter_Click()
    If Me.Specials.SelCount = 1 Then
        requeryForm "Special"
        'Me.barPosition.value = 0
        Me.CurrentOffset = 0
        barPosition_Change
    End If
End Sub

Private Sub btnTBPubAlert_Click(index As Integer)
    Select Case index
        Case Is = 0
            SetFilter "T/B Pub - RED"
        Case Is = 1
            SetFilter "T/B Pub - GREEN"
        Case Is = 2
            SetFilter "T/B Pub - YELLOW"
        Case Is = 3
            SetFilter "T/B Pub - BLUE"
    End Select
End Sub

Private Sub btnTBPubMisc_Click(index As Integer)
    Select Case index
        Case Is = 0
            SetFilter "T/B Pub - Repl BDC"
        Case Is = 1
            SetFilter "T/B Pub - Repl DC"
        Case Is = 2
            SetFilter "T/B Pub - QOO"
        Case Is = 3
            SetFilter "T/B Pub - QOH"
    End Select
End Sub

'Private Sub btnToggleExtSpecs_Click()
'    If Me.ExtendedSpecs.Visible = True Then
'        Me.ExtendedSpecs.Visible = False
'        Me.ExtendedSpecsPreview.Visible = True
'        LoadExtendedSpecsWebpage Me.ItemNumber
'    Else
'        Me.ExtendedSpecs.Visible = True
'        Me.ExtendedSpecsPreview.Visible = False
'    End If
'End Sub

Private Sub btnWebOffload_Click()
    If IsFormLoaded("WebOffload") Then
        UpdateWebOffloadChangedQuantity
        FocusOnForm "WebOffload"
    Else
        Load WebOffload
        WebOffload.Show
    End If
End Sub

Private Sub btnWebPathEdit_Click()
    Load AddItemWebPaths
    AddItemWebPaths.Show MODAL
    requeryWebPaths Me.ItemNumber
End Sub

Private Sub secondaryCategoryEditBtn_Click()
'Shell "\\TOOLSPLUS04\Databases\mastest\mas90-signs\A_Dist\toms beta\EBayCategoryMapper.exe", vbNormalFocus
 Load SignMaintenanceEBayCategory
    SignMaintenanceEBayCategory.LoadItem Me.ItemNumber, 2
    SignMaintenanceEBayCategory.Show MODAL
    'category string change is handled by dialog, nothing else to do
End Sub

Private Sub sphereOneEditMappingBtn_Click()

SphereOneCategoryMapper.Visible = True
SphereOneCategoryMapper.SetItemNumber Me.ItemNumber



''SphereOneCategoryMapper.Show False
End Sub

Private Sub WebPathTree_DblClick()
    btnWebPathEdit_Click
End Sub

Private Sub btnWebPathCopy_Click()
    ClipBoard_SetData "WebPaths for " & Me.ItemNumber
End Sub

Private Sub btnWebPathPaste_Click()
    Mouse.Hourglass True
    Dim itemtocopy As String
    itemtocopy = ClipBoard_GetData()
    If InStr(itemtocopy, "WebPaths for ") Then
        itemtocopy = Replace(itemtocopy, "WebPaths for ", "")
        If ExistsInPartNumbers(itemtocopy) Then
            If Me.WebPathTree.ListCount > 0 Then
                If vbYes = MsgBox("Remove existing paths first?", vbYesNo) Then
                    DB.Execute "DELETE FROM PartNumberWebPaths WHERE ItemNumber='" & Me.ItemNumber & "'"
                End If
            End If
            Dim rst As ADODB.Recordset
            Set rst = DB.retrieve("SELECT WebPathID FROM PartNumberWebPaths WHERE ItemNumber='" & itemtocopy & "'")
            While Not rst.EOF
                DB.Execute "INSERT INTO PartNumberWebPaths ( ItemNumber, WebPathID ) VALUES ( '" & Me.ItemNumber & "', " & rst("WebPathID") & " )", True
                rst.MoveNext
            Wend
            rst.Close
            Set rst = Nothing
            requeryWebPaths Me.ItemNumber
        Else
            MsgBox "Invalid clipboard contents."
        End If
    Else
        MsgBox "Invalid clipboard contents."
    End If
    Mouse.Hourglass False
End Sub

'Private Sub chkBaseItem_Click()
'    If Not fillingForm Then
'        If Not CBool(Me.chkBaseItem) Then
'            DB.Execute "UPDATE PartNumbers SET BaseItemID=0 WHERE ItemNumber='" & Me.ItemNumber & "'"
'        End If
'        baseItemFrameToggle CBool(Me.chkBaseItem)
'    End If
'End Sub

Private Sub chkDCSpecs_Click()
    If Not fillingForm Then
        DB.Execute "UPDATE PartNumbers SET ShowDCSpecs=" & SQLBool(Me.chkDCSpecs) & " WHERE ItemNumber='" & Me.ItemNumber & "'"
        refreshForm
    End If
End Sub

'Private Sub chkDisplayOnFrontPage_Click()
'    If Not fillingForm Then
'        If Me.chkDisplayOnFrontPage = 0 Then    'closest you can get to a locked checkbox?
'            fillingForm = True
'            Me.chkDisplayOnFrontPage = 1
'            fillingForm = False
'        Else
'            fillingForm = True
'            Me.chkDisplayOnFrontPage = 0
'            fillingForm = False
'        End If
'    End If
'End Sub

Private Sub chkHideImages_Click()
    If Not fillingForm Then
        Mouse.Hourglass True
        If Me.chkHideImages Then
            'Me.ManufLogo = LoadPicture("")
            'Me.ItemPicture = LoadPicture("")
            Me.ManufLogo.DisplayImage ""
            Me.ManufLogo.Visible = False
            Me.ItemPicture.DisplayImage ""
            Me.ItemPicture.Visible = False

            Me.CrossSellFrame.Left = Me.CrossSellFrame.Left - XSELL_EXPAND_BY
            Me.CrossSellFrame.width = Me.CrossSellFrame.width + XSELL_EXPAND_BY
            Me.CrossSellList.width = Me.CrossSellList.width + XSELL_EXPAND_BY
            Me.btnCrossSellAdd.Left = Me.btnCrossSellAdd.Left + XSELL_EXPAND_BY
            Me.btnCrossSellDelete.Left = Me.btnCrossSellDelete.Left + XSELL_EXPAND_BY
            Me.btnCrossSellMoveDown.Left = Me.btnCrossSellMoveDown.Left + XSELL_EXPAND_BY
            Me.btnCrossSellMoveUp.Left = Me.btnCrossSellMoveUp.Left + XSELL_EXPAND_BY
            Me.btnCrossSellPush.Left = Me.btnCrossSellPush.Left + XSELL_EXPAND_BY
            Me.btnCrossSellSuggest.Left = Me.btnCrossSellSuggest.Left + XSELL_EXPAND_BY
            
            Me.ItemPicture.SetCachePointer Nothing
            Me.ManufLogo.SetCachePointer Nothing
            Set ImageCache = Nothing
        Else
            'ResizeAndInsertPic GenerateLogoPicPath(Me.ItemNumber), Me.ManufLogo, MANUF_LOGO_HEIGHT, MANUF_LOGO_WIDTH
            'ResizeAndInsertPic GenerateItemPicPath(Me.ItemNumber), Me.ItemPicture, ITEM_IMG_HEIGHT, ITEM_IMG_WIDTH
            Me.ManufLogo.DisplayImage PictureDBFunctions.GenerateLogoPicturePath(Me.ItemNumber)
            Me.ItemPicture.DisplayImage PictureDBFunctions.GenerateItemPicturePathAny(Me.ItemNumber)
            
            Me.btnCrossSellSuggest.Left = Me.btnCrossSellSuggest.Left - XSELL_EXPAND_BY
            Me.btnCrossSellPush.Left = Me.btnCrossSellPush.Left - XSELL_EXPAND_BY
            Me.btnCrossSellMoveUp.Left = Me.btnCrossSellMoveUp.Left - XSELL_EXPAND_BY
            Me.btnCrossSellMoveDown.Left = Me.btnCrossSellMoveDown.Left - XSELL_EXPAND_BY
            Me.btnCrossSellDelete.Left = Me.btnCrossSellDelete.Left - XSELL_EXPAND_BY
            Me.btnCrossSellAdd.Left = Me.btnCrossSellAdd.Left - XSELL_EXPAND_BY
            Me.CrossSellList.width = Me.CrossSellList.width - XSELL_EXPAND_BY
            Me.CrossSellFrame.width = Me.CrossSellFrame.width - XSELL_EXPAND_BY
            Me.CrossSellFrame.Left = Me.CrossSellFrame.Left + XSELL_EXPAND_BY
            
            Me.ManufLogo.Visible = True
            Me.ItemPicture.Visible = True
            
            Set ImageCache = New Dictionary
            Me.ItemPicture.SetCachePointer ImageCache
            Me.ManufLogo.SetCachePointer ImageCache
        End If
        Mouse.Hourglass False
    End If
End Sub

Private Sub btnImageUpload_Click()
    'Shell PERL & " -MToolsPlus::ImageDB -e ""ToolsPlus::ImageDB::sync_specific_items('" & Me.ItemNumber & "'); system('pause');""", vbNormalFocus
    Shell PERL & " -MToolsPlus::ImageDB -e ""ToolsPlus::ImageDB::sync_specific_items('" & Me.ItemNumber & "');""", vbNormalFocus
    If Me.chkEBayPublished Then
        EBay.EBayMarkRevisionRequired Me.ItemNumber
    End If
End Sub

Private Sub cmbItemPictureType_Click()
    Me.ItemPicture.DisplayImage PictureDBFunctions.GenerateItemPicturePath(Me.ItemNumber, Array(CStr(Me.cmbItemPictureType)))
End Sub

Private Sub cmbFilters_Click()
    changed = True
    cmbFilters_LostFocus
End Sub

Private Sub cmbFilters_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.cmbFilters, KeyCode, Shift
End Sub

Private Sub cmbFilters_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.cmbFilters, KeyAscii
    If KeyAscii = vbKeyReturn Then
        cmbFilters_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub cmbFilters_LostFocus()
    AutoCompleteLostFocus Me.cmbFilters
    If changed = True Then
        changed = False
        If Me.cmbFilters = "" Then
            Me.cmbFilters = "All"
        End If
        requeryForm Me.cmbFilters
        If IsFormLoaded("InventoryMaintenance") Then
            If vbYes = MsgBox("Filter PO/Inv too?", vbYesNo) Then
                InventoryMaintenance.SetFilter Me.cmbFilters
            End If
        End If
        'If Me.barPosition.value = 0 Then
        If Me.CurrentOffset = 0 Then
            barPosition_Change
        Else
            'Me.barPosition.value = 0
            Me.CurrentOffset = 0
        End If
        Me.jumpToLine.SetFocus
    End If
End Sub

Private Sub cmbOtherFunctions_Click()
    Dim skipme As Boolean
    skipme = False
    Dim i As Long
    Select Case Me.cmbOtherFunctions
        Case Is = "<Other>"
            'do nothing
            skipme = True
        Case Is = "Edit Categories"
            'If Not IsFormLoaded("AddWebPath") Then
            '    Load AddWebPath
            '    AddWebPath.Show MODAL
            'End If
            If HasPermissionsTo("SignMaintenanceWrite") Then
                If Not IsFormLoaded("AddWebPaths2") Then
                    Load AddWebPaths2
                    AddWebPaths2.Show
                End If
            Else
                MsgBox "You don't have permission to use this form"
            End If
        'Case Is = "Edit Web Signs"
        '    If Not IsFormLoaded("AddWebSign") Then
        '        Load AddWebSign
        '        AddWebSign.Show MODAL
        '    End If
        'Case Is = "Edit Ship Classes"
        '    If Not IsFormLoaded("AddShipClass") Then
        '        Load AddShipClass
        '        AddShipClass.Show MODAL
        '    End If
        Case Is = "Barcode"
            If IsFormLoaded("Barcode") Then
                Barcode.LoadItem Me.ItemNumber
                FocusOnForm "Barcode"
            Else
                If HasPermissionsTo("BarCode") Then
                    Load Barcode
                    Barcode.LoadItem Me.ItemNumber
                    Barcode.Show
                Else
                    MsgBox "You don't have permission to use the Barcode form"
                End If
            End If
        Case Is = "Acronyms"
            'If Not IsFormLoaded("AddAcronyms") Then
            '    Load AddAcronyms
            '    AddAcronyms.Show
            'End If
            MsgBox "Broken, let Brian know if you want to use this again!"
        Case Is = "Offload Filter (Yahoo)"
            If HasPermissionsTo("WebOffload") Then
                Dim doOffload As Boolean
                doOffload = True
                If MsgBox("Offload all " & UBound(ItemList) + 1 & " items in filter?", vbYesNo) = vbNo Then
                    doOffload = False
                Else
                    If UBound(ItemList) > 1000 Then
                        If vbNo = MsgBox("This is a big filter, are you REALLY sure?", vbYesNo) Then
                            doOffload = False
                        End If
                    End If
                End If
                If doOffload Then
                    Mouse.Hourglass True
                    For i = 0 To UBound(ItemList)
                        Me.lblStatusBar.Caption = "Updating " & ItemList(i)
                        DoEvents
                        DB.Execute "UPDATE PartNumbers SET LastChanged=GETDATE() WHERE ItemNumber='" & ItemList(i) & "'"
                    Next i
                    Me.lblStatusBar.Caption = ""
                    Mouse.Hourglass False
                End If
            Else
                MsgBox "You don't have web offload permissions."
            End If
        Case Is = "Offload Filter (EBay)"
            If HasPermissionsTo("WebOffload") Then
                Mouse.Hourglass True
                For i = 0 To UBound(ItemList)
                    Me.lblStatusBar.Caption = "Updating " & ItemList(i)
                    EBay.EBayMarkRevisionRequired ItemList(i)
                    DoEvents
                Next i
                Me.lblStatusBar.Caption = ""
                Mouse.Hourglass False
            Else
                MsgBox "You don't have web offload permissions."
            End If
        Case Is = "Global Replace"
            If Not IsFormLoaded("GlobalFindReplace") Then
                Load GlobalFindReplace
                GlobalFindReplace.Show
            End If
    End Select
    If Not skipme Then
        Me.cmbOtherFunctions = "<Other>"
    End If
End Sub

Private Sub cmbShippingExtra_Click()
    If Not fillingForm Then
        Dim id As Long
        id = shippingExtraIndexToIDMap.item(CLng(Me.cmbShippingExtra.ListIndex))
        updatePartNumbers "ShippingExtraID", IIf(id = -1, "NULL", CStr(id)), Me.ItemNumber, ""
    End If
End Sub

Private Sub cmbSignFunctions_Click()
    Dim skipme As Boolean
    skipme = False
    Select Case Me.cmbSignFunctions
        Case Is = "<Signs>"
            'do nothing
            skipme = True
        'Case Is = "Design Sign"
        '    OpenSignInAccess Me.SignName, acViewDesign
        'Case Is = "Preview Sign"
        '    Load SignPreviewModeDlg
        '    SignPreviewModeDlg.hiddenItemNumber = Me.ItemNumber
        '    SignPreviewModeDlg.hiddenSignName = Me.SignName
        '    SignPreviewModeDlg.hiddenViewMode = acViewPreview
        '    SignPreviewModeDlg.Show MODAL, SignMaintenance
        'Case Is = "Print Sign"
        '    If Me.chkPrintSign Then
        '        If vbYes = MsgBox("Use the NEW signs?", vbYesNo) Then
        '            Shell "s:\mastest\mas90-signs\signs\print_single_sign.bat " & Me.ItemNumber, vbNormalFocus
        '        Else
        '
        '        'If MsgBox("Are you sure you want to PRINT this sign?", vbYesNo) = vbYes Then
        '            SignPreviewModeDlg.hiddenItemNumber = Me.ItemNumber
        '            SignPreviewModeDlg.hiddenSignName = Me.SignName
        '            SignPreviewModeDlg.hiddenViewMode = acViewNormal
        '            SignPreviewModeDlg.Show
        '            SignPreviewModeDlg.printer = "\\toolsplus04\phaser 8550dp ps"
        '            SignPreviewModeDlg.btnOK_Click
        '        'End If
        '
        '        End If
        '    Else
        '        MsgBox "This item does not need a sign. If you really want to print it, check the ""print sign y/n"" then hit this button again."
        '    End If
        Case Is = "Print Sign"
            'Load SignPreviewModeDlg
            'SignPreviewModeDlg.hiddenItemNumber = Me.ItemNumber
            'SignPreviewModeDlg.hiddenSignName = Me.SignName
            'SignPreviewModeDlg.Show MODAL, SignMaintenance
            MsgBox "This is broken! Tell Brian if you use it, and he'll fix it."
    End Select
    If Not skipme Then
        Me.cmbSignFunctions = "<Signs>"
    End If
End Sub

'Private Sub cmbSpecialChars_Click()
'    cmbSpecialChars_LostFocus
'End Sub
'
'Private Sub cmbSpecialChars_KeyDown(KeyCode As Integer, Shift As Integer)
'    AutoCompleteKeyDown Me.cmbSpecialChars, KeyCode, Shift
'    If KeyCode = vbKeyC And Shift = vbCtrlMask Then
'        ClipBoard_SetData (Me.cmbSpecialChars)
'    ElseIf KeyCode = vbKeyX And Shift = vbCtrlMask Then
'        ClipBoard_SetData (Me.cmbSpecialChars)
'        Me.cmbSpecialChars = ""
'    End If
'End Sub
'
'Private Sub cmbSpecialChars_KeyPress(KeyAscii As Integer)
'    AutoCompleteKeyPress Me.cmbSpecialChars, KeyAscii
'    If KeyAscii = vbKeyReturn Then
'        cmbSpecialChars_LostFocus
'    End If
'End Sub
'
'Private Sub cmbSpecialChars_LostFocus()
'    AutoCompleteLostFocus Me.cmbSpecialChars
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Me.jumpToLine.SetFocus
    ElseIf KeyCode = vbKeyF2 Then
        Me.jumpToItem.SetFocus
    ElseIf KeyCode = vbKeyF3 Then
        btnExtendedSpecsSearchContinue_Click
    'ElseIf (KeyCode = vbKeyF8 Or KeyCode = vbKeyPageUp) And Me.barPosition <> Me.barPosition.Min Then
    ElseIf (KeyCode = vbKeyF8 Or KeyCode = vbKeyPageUp) And Me.CurrentOffset <> Me.CurrentStartRecord Then
        'Me.barPosition = Me.barPosition - 1
        Me.CurrentOffset = Me.CurrentOffset - 1
        If TypeOf Me.ActiveControl Is TextBox Then
            HandleTextBoxCursorPosition Me.ActiveControl
        ElseIf TypeOf Me.ActiveControl Is ComboBox Then
            HandleComboBoxCursorPosition Me.ActiveControl
        End If
        KeyCode = 0
        barPosition_Change
    'ElseIf (KeyCode = vbKeyF9 Or KeyCode = vbKeyPageDown) And Me.barPosition <> Me.barPosition.max Then
    ElseIf (KeyCode = vbKeyF9 Or KeyCode = vbKeyPageDown) And Me.CurrentOffset <> Me.CurrentEndRecord Then
        'Me.barPosition = Me.barPosition + 1
        Me.CurrentOffset = Me.CurrentOffset + 1
        If TypeOf Me.ActiveControl Is TextBox Then
            HandleTextBoxCursorPosition Me.ActiveControl
        ElseIf TypeOf Me.ActiveControl Is ComboBox Then
            HandleComboBoxCursorPosition Me.ActiveControl
        End If
        KeyCode = 0
        barPosition_Change
    ElseIf KeyCode = vbKeyF5 Then
        'Debug.Print "F5 hit, checking back, # in history is " & backButtonHistory.Scalar
        'Debug.Print "history is: " & backButtonHistory.Join(", ")
        'Debug.Print "pointer is at " & backButtonPointer
        If backButtonPointer = 1 Then
            'skip
        Else
            backButtonPointer = backButtonPointer - 1
            backButtonActive = True
            SignMaintenance.GoToItem CStr(backButtonHistory.Elem(backButtonPointer - 1))
            backButtonActive = False
        End If
        KeyCode = 0
    ElseIf KeyCode = vbKeyF6 Then
        'Debug.Print "F6 hit, checking forward, # in history is " & backButtonHistory.Scalar
        'Debug.Print "history is: " & backButtonHistory.Join(", ")
        'Debug.Print "pointer is at " & backButtonPointer
        If backButtonPointer = backButtonHistory.Scalar Then
            'skip
        Else
            backButtonPointer = backButtonPointer + 1
            backButtonActive = True
            SignMaintenance.GoToItem CStr(backButtonHistory.Elem(backButtonPointer - 1))
            backButtonActive = False
        End If
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
Me.CurrentEndRecord = 50000
    'If USE_ALT_DATABASE Then
    '    Me.BackColor = &H80&
    '    Me.btnRefreshQtys.Enabled = False
    'End If
    Mouse.Hourglass True
    
    If GetCurrentUserNickname = "Kirk" Then
        Me.btnKirkUnlockSpecs.Visible = True
    End If
    
    backButtonActive = False
    Set backButtonHistory = New PerlArray
    
    If Not HasPermissionsTo("SignMaintenanceWrite") Then
        lockedDown = True
        lockForm
    End If
    If Not HasPermissionsTo("InventoryMaintenanceRead") Then
        Me.btnGoToInvMaint.Enabled = False
    End If
    If Not HasPermissionsTo("PrintSigns") Then
        'Me.btnSignPreview.Enabled = False
        'Me.btnSignPrint.Enabled = False
        'Me.btnSignDesign.Enabled = False
        Me.cmbSignFunctions.Enabled = False
    End If
    If Not HasPermissionsTo("WebOffload") Then
        Me.btnWebOffload.Enabled = False
    End If
    If Not HasPermissionsTo("Accessorize") Then
        Me.btnAccessoriesCopy.Enabled = False
        Me.btnAccessoriesPaste.Enabled = False
        Me.btnAddAccessories.Enabled = False
        'Me.btnAddAccessoriesReverse.Enabled = False
    End If
    
    If ImageCache Is Nothing Then
        Set ImageCache = New Dictionary
    End If
    Me.ItemPicture.SetCachePointer ImageCache
    Me.ManufLogo.SetCachePointer ImageCache
    
    ReDim tabs(0) As Long
    tabs(0) = 300
    SetListTabStops Me.WebPathTree.hWnd, tabs
    
    If Not lockedDown Then
        requeryTBPubButton
    End If
    
    Dim i As Long
    For i = 0 To Me.chkTabStrip.UBound - 1
        Me.ExtendedSpecs(i).Visible = False
    Next i
    Me.ExtendedSpecs(0).Visible = True
    
    'initialize cross sell headers
    Me.CrossSellList.SetColumnNames Array("Text to show", "URL")
    
    If Not lockedDown Then
        CheckForExpiredSpecials
        CheckForBeingDCItems
    End If
    requeryFilters
    requeryOtherFunctions
    requerySignFunctions
    'requeryGroupSignNew
    'requerySpecialChars
    requeryJumpToLine
    'requeryWebAccSign
    requeryAvailability
    'requeryBaseItemList
    
    Me.cmbItemPictureType.ListIndex = 0
    
    'Me.cmbFilters = "All"
    'requeryForm "All"
    If IsFormLoaded("InventoryMaintenance") Then
        Dim invfilter As String
        invfilter = InventoryMaintenance.GetFilter
        If Left(invfilter, 4) = "FBF-" Then
            Me.btnFilterByForm.tag = Mid(invfilter, 5)
            requeryForm "FBF"
'        ElseIf Left(invfilter, 4) = "FBS-" Then
'            Me.btnSearch.tag = Mid(invfilter, 5)
'            requeryForm "FBS"
        Else
            requeryForm "All"
        End If
    Else
        requeryForm "All"
    End If
    
    ExpandDropDownToFit Me.Availability
    ExpandDropDownToFit Me.cmbFilters

    firstLoad = True
    'Me.barPosition.value = 0
    Me.CurrentOffset = 0
    barPosition_Change
    firstLoad = False
    Mouse.Hourglass False
    
    Me.cmbFilters.SelLength = 0
End Sub

'Private Sub GroupSign_Click()
'    Me.GroupSignOrder = Mid(Me.GroupSign, InStr(Me.GroupSign, vbTab) + 1)
'    Enable Me.GroupSignOrder
'End Sub
'
'Private Sub GroupSign_DblClick()
'    If MsgBox("Do you want to delete this item from the selected group sign?", vbYesNo) = vbYes Then
'        DB.Execute "DELETE FROM PartNumberGroupSigns WHERE ItemNumber='" & Me.ItemNumber & "' AND GroupID=" & GroupSignHashNametoID.item(Left(Me.GroupSign, InStr(Me.GroupSign, vbTab) - 1))
'    End If
'End Sub

Private Sub ItemPicture_DblClick()
    'OpenDefaultApp GenerateItemPicPath(Me.ItemNumber)
    Dim path As String
    path = Me.ItemPicture.GetImagePath
    If path <> "" Then
        OpenDefaultApp path
    End If
End Sub

Private Sub jumpToLine_Click()
    changed = True
    jumpToLine_LostFocus
End Sub

Private Sub jumpToLine_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.jumpToLine, KeyCode, Shift
    If Shift = vbCtrlMask And KeyCode = vbKeyV Then
        Dim item As String
        item = ClipBoard_GetData
        item = Replace(item, vbCrLf, "")
        item = Replace(item, " ", "")
        If ExistsInPartNumbers(item) Then
            stopLineLostFocus = True
            Me.jumpToLine = Left(item, 3)
            'barPosition_Change
            Me.jumpToItem = Mid(item, 4)
            IHATEVISUALBASIC = True
             'barPosition_Change
            'requeryJumpToItem
           DoEvents
            jumpToItem_LostFocus
            'jumpToLine_LostFocus
            'Me.jumpToItem.SetFocus
            changed = True
            
        Else
            MsgBox "Unable to find item " & item & "!"
            KeyCode = 0
            Shift = 0
        End If
    End If
End Sub

Private Sub jumpToLine_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.jumpToLine, KeyAscii
    If KeyAscii = vbKeyReturn Then
        Me.jumpToItem.SetFocus
        'jumpToItem_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub jumpToLine_LostFocus()
    If IHATEVISUALBASIC = True Then
    IHATEVISUALBASIC = False
    Exit Sub
    End If
    
    Mouse.Hourglass True
    If changed = True And stopLineLostFocus = False Then
        changed = False
        AutoCompleteLostFocus Me.jumpToLine
        If Len(Me.jumpToLine) = 3 Then
            requeryJumpToItem
            Dim pos As Long
            pos = findFirstInLC(Me.jumpToLine, ItemList)
            If pos = -1 Then
                MsgBox "No items for line code in current filter?"
            Else
                'Me.barPosition.value = pos
                Me.CurrentOffset = pos
                'jumpToItem_LostFocus
                barPosition_Change
            End If
            'Me.jumpToItem = Mid(ItemList(Me.barPosition.value), 4)
            Me.jumpToItem = Mid(ItemList(Me.CurrentOffset), 4)
            'jumpToItem_LostFocus
            barPosition_Change
            Me.jumpToItem.SetFocus
        ElseIf Len(Me.jumpToLine) = 0 Then
            Me.jumpToLine = Left(Me.ItemNumber, 3)
        End If
    ElseIf stopLineLostFocus = True Then
        stopLineLostFocus = False
        requeryJumpToItem
        jumpToLine_LostFocus
        If changed = True Then Me.jumpToItem.ListIndex = 0
        changed = False
    End If
    Mouse.Hourglass False
End Sub

Private Sub jumpToItem_Click()
    changed = True
    jumpToItem_LostFocus
End Sub

Private Sub jumpToItem_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.jumpToItem, KeyCode, Shift
    If KeyCode = vbKeyV And Shift = vbCtrlMask Then
        Dim pn As String
        pn = ClipBoard_GetData
        If InCombo(pn, Me.jumpToItem, True) Then
            changed = False
            GoToItem Me.jumpToLine & pn, True
        Else
            MsgBox "Item not found! Are you in the correct line code?"
        End If
    End If
End Sub

Private Sub jumpToItem_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.jumpToItem, KeyAscii
    If KeyAscii = vbKeyReturn Then
        jumpToItem_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub jumpToItem_LostFocus()

    If changed = True Then
        changed = False
        If Me.jumpToItem = "" Then
            Me.jumpToItem = Mid(Me.ItemNumber, 4)
        Else
            Mouse.Hourglass True
            AutoCompleteLostFocus Me.jumpToItem
            Dim pos As Long
            If Len(Me.jumpToItem) <> 0 Then
                pos = bsearch(Me.jumpToLine & Me.jumpToItem, ItemList)
                If pos = -1 Then
                    MsgBox "Item is not in current filter?"
                Else
                    'Me.barPosition.value = pos
                    Me.CurrentOffset = pos
                    barPosition_Change
                End If
            End If
            'Me.Desc(1).SetFocus
            Me.jumpToItem.selStart = 0
            Me.jumpToItem.SelLength = Len(Me.jumpToItem)
            Mouse.Hourglass False
        End If
    End If
 
End Sub

Private Sub lblDesc_Click(index As Integer)
    Me.Desc(index).SetFocus
    Me.Desc(index).selStart = 0
    Me.Desc(index).SelLength = Len(Me.Desc(index))
End Sub

Private Sub lblSpec_Click(index As Integer)
    If Me.Spec(index).Enabled Then
        Me.Spec(index).SetFocus
        Me.Spec(index).selStart = 0
        Me.Spec(index).SelLength = Len(Me.Spec(index))
    End If
End Sub

Private Sub lblSpec_DblClick(index As Integer)
    If Not lockedDown Then
        'Dim rst As ADODB.Recordset
        'Select Case Index
        '    Case Is = 1, 2
        '        Me.chkDCSpecs = SQLBool(Not CBool(Me.chkDCSpecs))
            'Case Is = 1
            '    Set rst = retrieve("SELECT SpecialChar FROM SpecialCharacters WHERE ID=1")
            '    Me.Spec(1) = rst("SpecialChar")
            '    rst.Close
            '    Set rst = Nothing
            '    changed = True
            '    Spec_LostFocus 1
            'Case Is = 2
            '    Dim repl As String
            '    repl = DLookup("ReplacedBy", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'")
            '    If repl <> "" Then
            '        Set rst = retrieve("SELECT SpecialChar FROM SpecialCharacters WHERE ID=2")
            '        Me.Spec(2) = rst("SpecialChar")
            '        rst.Close
            '        Set rst = Nothing
            '        Me.Spec(2) = Replace(Me.Spec(2), "LINE-PART", LCase(repl), , 1)
            '        Me.Spec(2) = Replace(Me.Spec(2), "LINE-PART", IIf(Left(repl, 3) <> Left(Me.ItemNumber, 3), TitleCase(DLookup("ManufFullNameWeb", "ProductLine", "ProductLine='" & Left(repl, 3) & "'")) & " ", "") & Mid(repl, 4))
            '        changed = True
            '        Spec_LostFocus 2
            '    End If
        'End Select
    End If
End Sub




'-----------------------
' form related functions
'-----------------------
Private Sub requeryForm(filter As String)
    Mouse.Hourglass True
    requeryTBPubButton
    If Me.cmbFilters <> filter Then
        Me.cmbFilters = filter
    End If
    Dim rst As ADODB.Recordset
On Error Resume Next
    Dim amt As String
    Select Case filter
        Case Is = "All"
            Set rst = DB.retrieve("SELECT ItemNumber FROM vSignsRecSource ORDER BY ItemNumber")
        Case Is = "Special"
            'Set rst = DB.retrieve("exec spSignsRecSourceSpecial '" & Me.Specials & "'")
            Set rst = DB.retrieve("SELECT vSignsRecSource.ItemNumber FROM vSignsRecSource INNER JOIN PartNumberSpecials ON vSignsRecSource.ItemNumber=PartNumberSpecials.ItemNumber WHERE PartNumberSpecials.SpecialName='" & EscapeSQuotes(Me.Specials) & "' ORDER BY vSignsRecSource.ItemNumber")
            Me.cmbFilters = "Special - " & Me.Specials
        'Case Is = "Group Sign"
        '    Set rst = DB.retrieve("exec spSignsRecSourceGroupSign '" & Left(Me.GroupSign, InStr(Me.GroupSign, vbTab) - 1) & "'")
        '    Me.cmbFilters = "Group Sign - " & Left(Me.GroupSign, InStr(Me.GroupSign, vbTab) - 1)
        Case Is = "FBF", "Filter By Form"
            Set rst = DB.retrieve(Me.btnFilterByForm.tag)
            Me.cmbFilters = "Filter By Form"
'        Case Is = "FBS", "Filter By Search"
'            Set rst = DB.retrieve(Me.btnSearch.tag)
'            Me.cmbFilters = "Filter By Search"
        Case Is = "Accessories > X"
            amt = InputBox("Enter threshold accessories amount:")
            Set rst = DB.retrieve("SELECT vSignsRecSource.ItemNumber FROM vSignsRecSource INNER JOIN PartNumbers ON vSignsRecSource.ItemNumber=PartNumbers.ItemNumber INNER JOIN PartNumberAccessories ON vSignsRecSource.ItemNumber=PartNumberAccessories.ItemNumber WHERE PartNumbers.WebDeleted=0 GROUP BY vSignsRecSource.ItemNumber HAVING COUNT(*)>" & amt)
        Case Is = "Published no review"
            amt = InputBox("Enter number of days to go back:")
            Set rst = DB.retrieve("SELECT vSignsRecSource.ItemNumber FROM vSignsRecSource INNER JOIN PartNumbers ON vSignsRecSource.ItemNumber=PartNumbers.ItemNumber WHERE PartNumbers.WebPublished=1 AND DATEDIFF(d, PartNumbers.PublishedDate, GETDATE())<" & amt & " AND PartNumbers.ReviewedDate IS NULL ORDER BY vSignsRecSource.ItemNumber")
        Case Else
            Dim query As String
            query = DLookup("Query", "PoInvFilters", "FilterName='" & EscapeSQuotes(filter) & "'")
            If query = "" Then
                MsgBox "Signs: can't filter for " & qq(filter) & ", defaulting to ""All""."
                Set rst = DB.retrieve("SELECT vSignsRecSource.ItemNumber FROM vSignsRecSource ORDER BY vSignsRecSource.ItemNumber")
            Else
                Set rst = DB.retrieve(query)
            End If
    End Select
On Error GoTo 0
    If rst Is Nothing Then
        MsgBox "Unable to get the list of items. Probably an invalid SQL string, or timeout?"
    Else
        If rst.RecordCount > 0 Then
            ReDim ItemList(rst.RecordCount - 1) As String
            Dim i As Long
            For i = 0 To rst.RecordCount - 1
                If Left(rst("ItemNumber"), 3) = "XXX" Or Left(rst("ItemNumber"), 3) = "ZZZ" Then
                    ReDim Preserve ItemList(i - 1) As String
                    Exit For
                End If
                ItemList(i) = rst("ItemNumber")
                rst.MoveNext
            Next i
            Me.lblMaxAmt.Caption = "of " & UBound(ItemList) + 1
            'Me.lblMaxAmt.Caption = "of " & rst.re
            'Me.barPosition.Min = 0
            'Me.barPosition.max = UBound(ItemList)
            Me.CurrentEndRecord = UBound(ItemList)
        Else
            ReDim ItemList(0) As String
            ItemList(0) = "NORECORDS"
            Me.lblMaxAmt.Caption = "of 0"
            Me.CurrentEndRecord = 0
            'Me.barPosition.Min = 0
            'Me.barPosition.max = 0
        End If
        rst.Close
        Set rst = Nothing
    End If
    Mouse.Hourglass False
End Sub

Private Sub fillForm(item As ADODB.Recordset)
    Mouse.Hourglass True

    If backButtonActive = False Then
        If backButtonHistory.Elem(backButtonPointer) = item("ItemNumber") Then
            'Debug.Print "loaded item is same as forward, using existing history"
            backButtonPointer = backButtonPointer + 1
        Else
            'Debug.Print "adding " & item("ItemNumber") & " to back history"
            'Debug.Print "history is: " & backButtonHistory.Join(", ")
            'Debug.Print "pointer is at " & backButtonPointer
            While backButtonHistory.Scalar > backButtonPointer
                backButtonHistory.Pop
            Wend
            backButtonHistory.Push CStr(item("ItemNumber"))
            backButtonPointer = backButtonPointer + 1
        End If
    End If
    
    Dim i As Long
    Dim doLogoPic As Boolean
    If Me.jumpToLine = "" Or Left(item("ItemNumber"), 3) <> Left(Me.ItemNumber, 3) Then
        Me.jumpToLine = Left(item("ItemNumber"), 3)
        requeryJumpToItem
        doLogoPic = True
    Else
        doLogoPic = False
    End If
    Me.jumpToItem = Mid(item("ItemNumber"), 4)
    Me.ItemNumber = item("ItemNumber")
    Me.Website = WebsiteNameHash.item(CLng(item("WebsiteID").value))
    Me.Website.tag = item("WebsiteID").value
    'If Nz(item("BaseItem")) <> "" Then
    '    Me.baseItemList = item("BaseItem")
    '    baseItemFrameToggle True, True
    'Else
    '    baseItemFrameToggle False, True
    'End If
    Me.extendedSpecsFrame.Enabled = True
    fillDescs item
    fillSpecs item
    Me.ExtendedSpecs(0) = Nz(item("ExtendedSpecsHTML"))
    Me.ExtendedSpecs(1) = Nz(item("WriteupHTML"))
    Me.ExtendedSpecs(2) = Nz(item("FeaturesHTML"))
    Me.ExtendedSpecs(3) = Nz(item("TechSpecsHTML"))
    Me.ExtendedSpecs(4) = Nz(item("MediaHTML"))
    Me.ExtendedSpecs(5) = Nz(item("NotesHTML"))
    'Me.ExtendedSpecs(6) = Nz(item("QuestionsHTML"))
    For i = 0 To Me.ExtendedSpecs.UBound
        Me.chkTabStrip(i).FontBold = Len(Me.ExtendedSpecs(i)) > 0
    Next i
    'If Me.ExtendedSpecsPreview.Visible Then
    If Me.chkTabStrip(Me.chkTabStrip.UBound) Then
        loadExtendedSpecsWebpage item("ItemNumber")
    End If
    Me.LastChanged = Format$(item("LastChanged"), "m/d/yy")
    Me.TriadCode = Nz(item("TriadCode"))
    'Me.SignAcc = item("AccessorySign")
    'Me.SignWeb = item("MainSign")
    fillingForm = True
    Me.chkToBePublished = SQLBool(item("WebToBePublished"))
    Me.chkPublished = SQLBool(item("WebPublished"))
    Me.chkWebPointer = SQLBool(item("WebPointered"))
    Me.btnGroupEditor.Visible = CBool(Me.chkWebPointer = vbChecked)
    Me.chkWebDelete = SQLBool(item("WebDelete"))
    Me.chkWebDeleted = SQLBool(item("WebDeleted"))
    Me.chkWebOrderable = SQLBool(item("WebOrderable"))
    Me.chkRequestForm = SQLBool(item("ShowRequestForm"))
    Me.chkEBayTBPub = SQLBool(item("EBayToBePublished"))
    Me.chkEBayPublished = SQLBool(item("EBayPublished"))
    EBayCategorySetByID Nz(item("EBayCategoryID"), -1), 0
    EBayCategorySetByID Nz(item("EBayStoreCategoryID"), -1), 1
    If item("EBayToBePublished") = True Or item("EBayPublished") = True Then
        Select Case item("EBayStatusID")
            Case Is = EBayStatus.EBayStatusDown
                Me.EBayEBayStatus = "currently: DOWN"
            Case Is = EBayStatus.EbayStatusUp
                Me.EBayEBayStatus = "currently: UP"
            Case Else
                Me.EBayEBayStatus = "currently: ERROR!"
        End Select
    Else
        Me.EBayEBayStatus = ""
    End If
    'Me.chkEBayAllowBestOffer = SQLBool(item("EBayAllowBestOffer"))
    Me.chkPrintSign = SQLBool(item("PrintSign"))
    Me.chkPrintSignsByQOH = SQLBool(item("SignByQOH"))
    'Me.chkDisplayOnFrontPage = SQLBool(item("DisplayOnFront"))
    Me.chkDCSpecs = SQLBool(item("ShowDCSpecs"))
    Me.chkMetaNoIndex = SQLBool(item("MetaNoIndex"))
    fillingForm = False
    'Me.WebCaption = Nz(item("WebCaption"))
    Me.AvailabilityLimit = item("AvailabilityLimit")
    Me.EBayAvailabilityLimit = item("EBayReserveQty")
    Me.lblStockStatus.Caption = ItemStatusHashIDtoStr.item(item("ItemStatus").value)
    Me.lblAvailability = GetAvailabilityText(item("ItemStatus"), item("QtyOnHand"), item("QtyOnSO"), item("QtyOnPO"), item("QtyOrdered"), item("AvailabilityLimit"), item("ShowRequestForm"), Nz(item("AvailString")))
    Me.Availability = IIf(IsNull(item("ShortDesc")), "Regular availability", item("ShortDesc"))
    If Me.chkPublished Then
        Me.chkPublished.ToolTipText = "Date published: " & Format(item("PublishedDate"), "Short Date")
    End If
    If Me.chkToBePublished Then
        Me.TBPubPriority.Visible = True
        Me.TBPubPriority.tag = item("TBPubPriority")
        Select Case item("TBPubPriority")
            Case Is = 0
                Me.TBPubPriority.FillColor = BG_GREY
            Case Is = 1
                Me.TBPubPriority.FillColor = RED
            Case Is = 2
                Me.TBPubPriority.FillColor = YELLOW
            Case Is = 3
                Me.TBPubPriority.FillColor = GREEN
            Case Is = 4
                Me.TBPubPriority.FillColor = LT_BLUE
        End Select
    Else
        Me.TBPubPriority.Visible = False
        Me.TBPubPriority.tag = item("TBPubPriority")
    End If
    requeryWebPaths item("ItemNumber")
    requeryAccessories item("ItemNumber")
    requeryCrossSellList item("ItemNumber")
    requeryIncludesList item("ItemNumber")
    fillingForm = True
    requeryShippingExtras item("PrimaryVendor"), Nz(item("ShippingExtraID"))
    fillingForm = False
    'requeryGroupSign item("ItemNumber")
    'Me.GroupSignOrder = ""
    'Disable Me.GroupSignOrder
    requerySpecials item("ItemNumber")
    If Not CBool(Me.chkHideImages) Then
        If doLogoPic Then
            Me.ManufLogo.DisplayImage PictureDBFunctions.GenerateLogoPicturePath(Me.ItemNumber)
        End If
        'Me.ItemPicture.DisplayImage PictureDBFunctions.GenerateItemPicturePath(Me.ItemNumber, True, True)
        'Me.ItemPicture.DisplayImage PictureDBFunctions.GenerateItemPicturePathAny(Me.ItemNumber)
        Me.ItemPicture.DisplayImage PictureDBFunctions.GenerateItemPicturePath(Me.ItemNumber, Array(CStr(Me.cmbItemPictureType)))
        'MsgBox PictureDBFunctions.GenerateItemPicturePath(Me.ItemNumber, Array(CStr(Me.cmbItemPictureType)))
    End If
    Me.InternalNotes = Nz(item("InternalNotes"))
    
    fixTitleTextLengthField
    
    
    If EBayIsCurrentItemListedOutOfStock = True Then Me.EBayEBayStatus.Text = "UP - OutOfStock"
    
    'Tom's code 3-11-2015 to add in sphere 1 categories to signs
    SphereOneCatLst.Clear
    Dim sphereonecat As ADODB.Recordset
    Set sphereonecat = DB.retrieve("SELECT DISTINCT SphereOneCategories.Name FROM SphereOneCategories INNER JOIN MasterCategoryConnectors ON MasterCategoryConnectors.ConnectorID=SphereOneCategories.ID AND MasterCategoryConnectors.ConnectorType=4 INNER JOIN MasterCategories ON MasterCategories.ID=MasterCategoryConnectors.MasterID INNER JOIN PartNumberWebPaths ON PartNumberWebPaths.WebPathID=MasterCategories.ID WHERE ItemNumber='" + Me.ItemNumber + "'")
    If sphereonecat.RecordCount > 0 Then
        Do
            Me.SphereOneCatLst.AddItem sphereonecat("Name")
            sphereonecat.MoveNext
        Loop Until sphereonecat.EOF = True Or sphereonecat.BOF = True
    End If
    
    If SphereOneCategoryMapper.Visible = True Then
        Unload SphereOneCategoryMapper
        sphereOneEditMappingBtn_Click
    End If
    
    'check for overridden ebay secondary category. if none, then if the primary cat is not null, use its alternative...
    Dim SecondaryCategory As ADODB.Recordset
    Set SecondaryCategory = DB.retrieve("SELECT SecondaryCategoryID FROM EBaySecondaryCategoryOverride WHERE ItemNumber='" & Me.ItemNumber & "'")
    If SecondaryCategory.RecordCount > 0 Then
        
        Set SecondaryCategory = DB.retrieve("SELECT Name FROM EBayCategories WHERE ID=" & SecondaryCategory("SecondaryCategoryID"))
        If SecondaryCategory.RecordCount > 0 Then
            Me.SecondaryCategoryTxt.Text = SecondaryCategory("Name")
        Else
            Me.SecondaryCategoryTxt.Text = "<error finding catname>"
        End If
    Else
        If IsNull(item("EBayCategoryID")) = False And item("EBayStatusID") = EBayStatus.EbayStatusUp Then
            Set SecondaryCategory = DB.retrieve("SELECT Name FROM EBayCategories WHERE ID=dbo.AlternativeCategoryID(" & item("EBayCategoryID") & ")")
            If (SecondaryCategory.RecordCount > 0) Then
                Me.SecondaryCategoryTxt.Text = SecondaryCategory("Name")
            Else
                Me.SecondaryCategoryTxt.Text = "<none>"
            End If
        Else
            Me.SecondaryCategoryTxt.Text = "<none>"
        End If
    End If
    Mouse.Hourglass False
End Sub

Private Sub fixTitleTextLengthField()
    Dim txt As String
    If ManufWebHashPLtoName.exists(CStr(Left(Me.ItemNumber, 3))) Then
        If Not IsNull(ManufWebHashPLtoName.item(CStr(Left(Me.ItemNumber, 3)))) Then
            txt = ManufWebHashPLtoName.item(CStr(Left(Me.ItemNumber, 3)))
            txt = txt & " " & Me.Desc(2)
            txt = txt & " " & Me.Desc(1)
            txt = txt & " " & Me.Desc(3)
            txt = TrimWhitespace(txt, True, True, True)
            Me.lblTitleCount.Caption = Len(txt)
            If Len(txt) > 70 Then
                Me.lblTitleCount.ForeColor = RED
            Else
                Me.lblTitleCount.ForeColor = DARK_GREEN
            End If
        Else
            Me.lblTitleCount.Caption = "??"
            Me.lblTitleCount.ForeColor = RED
        End If
    Else
        Me.lblTitleCount.Caption = "??"
        Me.lblTitleCount.ForeColor = RED
    End If
End Sub

Private Sub fillDescs(item As ADODB.Recordset)
    Dim i As Long
    If Nz(item("BaseItem")) <> "" Then
        For i = 1 To 4
            Me.Desc(i) = FillBaseItemSpecs(Me.ItemNumber, i)
            Me.Desc(i).Enabled = False
        Next i
    Else
        For i = 1 To 4
            Me.Desc(i) = Nz(item("Desc" & i))
            Me.Desc(i).Enabled = True
        Next i
    End If
End Sub

Private Sub fillSpecs(item As ADODB.Recordset, Optional forceSkipTemplate = False)
    Dim i As Long
    'if has a base item
    'If Nz(item("BaseItem")) <> "" Then
    '    For i = 1 To 8
    '        Me.Spec(i) = FillBaseItemSpecs(Me.baseItemList, i)
    '        Me.Spec(i).Enabled = False
    '    Next i
    '    Me.extendedSpecsFrame.Enabled = False
    '    Exit Sub
    'End If
    
    'if d/c
    If item("ShowDCSpecs") Then
    'If item("ShowDCSpecs") Then
        For i = 1 To 8
            Me.Spec(i) = FillDCSpecs(Me.ItemNumber, i)
            Me.Spec(i).Enabled = False
        Next i
        Exit Sub
    End If
    
    'if grouped item
    Dim groupMaster As String
    groupMaster = Logic.GetGroupedMasteritemForSubitem(Me.ItemNumber)
    If groupMaster <> "" And forceSkipTemplate = False Then
        For i = 1 To 8
            Me.Spec(i) = FillTemplateItemSpecs(groupMaster, i)
            Me.Spec(i).Enabled = False
        Next i
        Me.extendedSpecsFrame.Enabled = False
        Exit Sub
    End If
    
    'if not d/c
    For i = 1 To 8
        Me.Spec(i) = Nz(item("Spec" & i & "HTML"))
        Me.Spec(i).Enabled = True
    Next i
End Sub

Private Sub requeryFilters()
    Me.cmbFilters.Clear
    Me.cmbFilters.AddItem "All"
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT FilterName FROM PoInvFilters WHERE ShowInSigns=1 ORDER BY FilterName")
    While Not rst.EOF
        Me.cmbFilters.AddItem rst("FilterName")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Me.cmbFilters.AddItem "Accessories > X"
    Me.cmbFilters.AddItem "Published no review"
End Sub

'Private Sub requeryGroupSignNew()
'    Mouse.Hourglass True
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT GroupName FROM GroupSigns ORDER BY GroupName")
'    Me.GroupSignNew.Clear
'    While Not rst.EOF
'        Me.GroupSignNew.AddItem rst("GroupName")
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'    Mouse.Hourglass False
'End Sub
'
'Private Sub requerySpecialChars()
'    Mouse.Hourglass True
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT SpecialChar FROM SpecialCharacters ORDER BY ID")
'    Me.cmbSpecialChars.Clear
'    While Not rst.EOF
'        Me.cmbSpecialChars.AddItem (rst("SpecialChar"))
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'    Mouse.Hourglass False
'End Sub

Private Sub requeryWebPaths(item As String)
    Mouse.Hourglass True
    If IsGroupedSubitem(item) Then
        Me.WebPathsFrame.Enabled = False
        Me.WebPathTree.Clear
        Me.WebPathTree.AddItem "Group sub-item, no paths needed"
    Else
        Me.WebPathsFrame.Enabled = True
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT WebPathID FROM PartNumberWebPaths WHERE ItemNumber='" & item & "'")
        Me.WebPathTree.Clear
        While Not rst.EOF
            Dim paths As PerlArray
            Set paths = New PerlArray
            Dim pid As String
            pid = rst("WebPathID")
            While pid <> "-1"
                'getname will auto-add the path if not there, huzzah for dwim?
                paths.UnShift Array(WebPathCache_GetName(pid), pid)
                pid = WebPathCache_GetParent(pid)
            Wend
            
            Dim i As Long
            Dim lastPos As Long
            lastPos = Me.WebPathTree.ListCount
            For i = 0 To paths.Scalar - 1
                Dim pos As Long
                pos = -1
                Dim j As Long
                For j = 0 To Me.WebPathTree.ListCount - 1
                    If paths.Elem(i)(1) = Mid(Me.WebPathTree.list(j), InStr(Me.WebPathTree.list(j), vbTab) + 1) Then
                        pos = j
                        Exit For
                    End If
                Next j
                If pos = -1 Then
                    Me.WebPathTree.AddItem Space(i * 2) & paths.Elem(i)(0) & vbTab & paths.Elem(i)(1), lastPos
                    lastPos = lastPos + 1
                Else
                    'do nothing
                    lastPos = pos + 1
                End If
                If i = paths.Scalar - 1 Then
                    Me.WebPathTree.list(lastPos - 1) = "*" & Me.WebPathTree.list(lastPos - 1)
                End If
            Next i
            
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
    End If
    Mouse.Hourglass False
End Sub

Private Sub requeryAccessories(item As String)
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    'Set rst = DB.retrieve("exec spSignsGetAccessories '" & item & "'")
    Set rst = DB.retrieve("SELECT Accessory FROM PartNumberAccessories WHERE ItemNumber='" & item & "' ORDER BY Accessory")
    Me.Accessories.Clear
    While Not rst.EOF
        Me.Accessories.AddItem rst("Accessory")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requeryOtherFunctions()
    Me.cmbOtherFunctions.Clear
    Me.cmbOtherFunctions.AddItem "<Other>"
    Me.cmbOtherFunctions.AddItem "Edit Categories"
    'Me.cmbOtherFunctions.AddItem "Edit Web Signs"
    'Me.cmbOtherFunctions.AddItem "Edit Ship Classes"
    Me.cmbOtherFunctions.AddItem "Barcode"
    Me.cmbOtherFunctions.AddItem "Acronyms"
    Me.cmbOtherFunctions.AddItem "Offload Filter (Yahoo)"
    Me.cmbOtherFunctions.AddItem "Offload Filter (EBay)"
    Me.cmbOtherFunctions.AddItem "Global Replace"
    ExpandDropDownToFit Me.cmbOtherFunctions
    Me.cmbOtherFunctions = "<Other>"
End Sub

Private Sub requerySignFunctions()
    Me.cmbSignFunctions.Clear
    Me.cmbSignFunctions.AddItem "<Signs>"
    'Me.cmbSignFunctions.AddItem "Preview Sign"
    'Me.cmbSignFunctions.AddItem "Design Sign"
    Me.cmbSignFunctions.AddItem "Print Sign"
    ExpandDropDownToFit Me.cmbSignFunctions
    Me.cmbSignFunctions = "<Signs>"
End Sub

Private Sub requeryIncludesList(item As String)
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT IncludeText FROM PartNumberIncludes WHERE ItemNumber='" & item & "' ORDER BY SortOrder")
    Me.IncludesList.Clear
    While Not rst.EOF
        Me.IncludesList.AddItem rst("IncludeText")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requeryCrossSellList(item As String)
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ShowText, URL FROM PartNumberCrossSell WHERE ItemNumber='" & item & "' ORDER BY SortOrder")
    Me.CrossSellList.Clear
    'While Not rst.EOF
        Me.CrossSellList.Add rst, -1, True
    '    rst.MoveNext
    'Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

'Private Sub requeryGroupSign(item As String)
'    Mouse.Hourglass True
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("exec spSignsGetGroupSign '" & item & "'")
'    Me.GroupSign.Clear
'    While Not rst.EOF
'        Me.GroupSign.AddItem rst("GroupName") & vbTab & rst("GroupOrder")
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'    Mouse.Hourglass False
'End Sub

Private Sub requerySpecials(item As String)
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT SpecialName FROM PartNumberSpecials WHERE ItemNumber='" & item & "'")
    Me.Specials.Clear
    While Not rst.EOF
        Me.Specials.AddItem (rst("SpecialName"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requeryShippingExtras(vendor As String, extraID As String)
    Me.cmbShippingExtra.Clear
    Set shippingExtraIndexToIDMap = New Dictionary
    Me.cmbShippingExtra.AddItem "<None>"
    shippingExtraIndexToIDMap.Add CLng(0), CLng(-1)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, ShortDescription FROM ShippingExtras WHERE VendorNumber='" & EscapeSQuotes(vendor) & "' ORDER BY ShortDescription")
    Dim toSelect As Long
    toSelect = 0
    Dim i As Long
    i = 1
    While Not rst.EOF
        Me.cmbShippingExtra.AddItem rst("ShortDescription")
        shippingExtraIndexToIDMap.Add CLng(i), CLng(rst("ID"))
        If rst("ID") = extraID Then
            toSelect = i
        End If
        i = i + 1
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Me.cmbShippingExtra.ListIndex = toSelect
    If toSelect = 0 And extraID <> "" Then
        'error, shipping extra no longer exists or item has a different vendor
        'notify and switch to none
        MsgBox "Shipping extra definition deleted, or the item vendor number has been changed! Switching this to ""none"", but you might want to fix that."
        updatePartNumbers "ShippingExtraID", "NULL", Me.ItemNumber, ""
    End If
End Sub

Private Sub requeryJumpToLine()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM vRequeryLineCodes ORDER BY ProductLine")
    Me.jumpToLine.Clear
    While Not rst.EOF
        Me.jumpToLine.AddItem (rst("ProductLine"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requeryJumpToItem()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("exec spRequeryJumpToItem '" & Me.jumpToLine & "'")
    Me.jumpToItem.Clear
    While Not rst.EOF
        Me.jumpToItem.AddItem Mid$(rst("ItemNumber"), 4)
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    ExpandDropDownToFit Me.jumpToItem
    Mouse.Hourglass False
End Sub

'Private Sub requeryWebAccSign()
'    Mouse.Hourglass True
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT SignName, SignType FROM WebAccSign")
'    Me.SignAcc.Clear
'    Me.SignWeb.Clear
'    While Not rst.EOF
'        Select Case rst("SignType")
'            Case Is = "A"
'                Me.SignAcc.AddItem (rst("SignName"))
'            Case Is = "P"
'                Me.SignWeb.AddItem (rst("SignName"))
'        End Select
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'    Mouse.Hourglass False
'End Sub

Private Sub requeryAvailability()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ShortDesc FROM AvailOverrides ORDER BY ID")
    Me.Availability.Clear
    Me.Availability.AddItem "Regular availability"
    While Not rst.EOF
        Me.Availability.AddItem rst("ShortDesc")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

'Private Sub requeryBaseItemList()
'    Mouse.Hourglass True
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT BaseItem FROM BaseItemList ORDER BY BaseItem")
'    Me.baseItemList.Clear
'    While Not rst.EOF
'        Me.baseItemList.AddItem rst("BaseItem")
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'    Mouse.Hourglass False
'End Sub

Private Sub requeryTBPubButton()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT SUM(CASE WHEN TBPubPriority=0 THEN 1 ELSE 0 END) AS PriorityGrey, " & _
                                 "SUM(CASE WHEN TBPubPriority=1 THEN 1 ELSE 0 END) AS PriorityRed, " & _
                                 "SUM(CASE WHEN TBPubPriority=2 THEN 1 ELSE 0 END) AS PriorityYellow, " & _
                                 "SUM(CASE WHEN TBPubPriority=3 THEN 1 ELSE 0 END) AS PriorityGreen, " & _
                                 "SUM(CASE WHEN TBPubPriority=4 THEN 1 ELSE 0 END) AS PriorityBlue " & _
                          "FROM PartNumbers WHERE WebToBePublished=1")
    If Nz(rst("PriorityRed")) = "" Then
        Me.btnTBPubAlert(0).Visible = False
    Else
        Me.btnTBPubAlert(0).Caption = rst("PriorityRed")
        Me.btnTBPubAlert(0).Visible = True
    End If
    
    If Nz(rst("PriorityYellow")) = "" Then
        Me.btnTBPubAlert(2).Visible = False
    Else
        Me.btnTBPubAlert(2).Caption = rst("PriorityYellow")
        Me.btnTBPubAlert(2).Visible = True
    End If
    
    If Nz(rst("PriorityGreen")) = "" Then
        Me.btnTBPubAlert(1).Visible = False
    Else
        Me.btnTBPubAlert(1).Caption = rst("PriorityGreen")
        Me.btnTBPubAlert(1).Visible = True
    End If
    
    If Nz(rst("PriorityBlue")) = "" Then
        Me.btnTBPubAlert(3).Visible = False
    Else
        Me.btnTBPubAlert(3).Caption = rst("PriorityBlue")
        Me.btnTBPubAlert(3).Visible = True
    End If
    
    rst.Close
    Set rst = Nothing
End Sub

Private Sub loadExtendedSpecsWebpage(item As String)
    Open DESTINATION_DIR & "extspec.html" For Output As #1
    Print #1, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01//EN"" ""http://www.w3.org/TR/html4/strict.dtd"">" & vbCrLf & _
              "<html>" & vbCrLf & _
              " <head>" & vbCrLf & _
              "  <title>" & item & "</title>" & vbCrLf & _
              "  <base href=""http://" & WebsiteURLHash.item(CLng(Me.Website.tag)) & "/"" target=""_blank"">" & vbCrLf & _
              "  <link rel=""stylesheet"" type=""text/css"" href=""http://p1.hostingprod.com/@tools-plus.com/solidcactus/toolsplus-styles.css"">" & vbCrLf & _
              "  <script type=""text/javascript"" src=""http://ajax.googleapis.com/ajax/libs/jquery/1.7.2/jquery.min.js""></script>" & vbCrLf & _
              "  <script type=""text/javascript"" src=""http://p1.hostingprod.com/@tools-plus.com/solidcactus/toolsplus.js""></script>" & vbCrLf & _
              "  <script>" & vbCrLf & "// Answers Cloud Services Embed Script v1.01" & vbCrLf & "// DO NOT MODIFY BELOW THIS LINE" & "*****************************************" & vbCrLf & _
              "  ;(function (g) {" & vbCrLf & "   var d = document, i, am = d.createElement('script'), h = d.head || d.getElementsByTagName(""head"")[0]," & vbCrLf & _
              "       aex = {" & vbCrLf & "        ""src"": ""//gateway.answerscloud.com/toolsplus/production/gateway.min.js""," & vbCrLf & _
              "         ""type"": ""text/javascript""," & vbCrLf & "         ""async"": ""true"","
    Print #1, "         ""data-vendor"": ""acs""," & vbCrLf & "          ""data-role"": ""gateway""" & vbCrLf & _
              "        };" & vbCrLf & _
              "     for (var attr in aex) { am[attr] = aex[attr]; }" & vbCrLf & _
              "     h.appendChild(am);" & vbCrLf & _
              "     g['acsReady'] = function () {var aT = '__acsReady__', args = Array.prototype.slice.call(arguments, 0), k = setInterval(function () {if (typeof g[aT] === 'function') {clearInterval(k);for (i = 0; i < args.length; i++) {g[aT].call(g, function(fn) { return function() {setTimeout(fn, 1)};}(args[i]));}}}, 50);};" & vbCrLf & _
              "  })(window);" & vbCrLf & vbCrLf & _
              " // DO NOT MODIFY ABOVE THIS LINE*****************************************" & vbCrLf & _
              " </script>" & vbCrLf & _
              "  <!--  <link rel=""stylesheet"" href=""http://lib.store.yahoo.net/lib/toolsplus/pdStyleReviews.css"" type=""text/css"">-->" & vbCrLf & _
              "  <!--  <script type=""text/javascript"" src=""http://lib.store.yahoo.net/lib/toolsplus/reviewsClientScript.js""></script>-->" & vbCrLf & _
              "  <script type=""text/javascript"">" & vbCrLf & _
              "  var prClientDomain = ""http://myaccount.tools-plus.com/"";" & vbCrLf & "  var prClientName = ""tools-plus"";" & vbCrLf & "  var prClientSkin = ""tools-plus"";" & vbCrLf & "  prDebug = false;" & vbCrLf & "  function pdOpenProductReviewsTab() {" & vbCrLf & "     //**********************************************" & vbCrLf & "     //*** INCLUDE OPEN TAB JS HERE - IF REQUIRED ***" & vbCrLf & "     //**********************************************" & vbCrLf & "     TP.Tabs.set_active(""Reviews"",1);" & vbCrLf & "  }" & vbCrLf & " </script>" & vbCrLf & " </head>" & vbCrLf & " <body>"
    Print #1, "  <div style=""text-align: left;"">" 'takes the place of #overall
    Print #1, "  <div style=""width: 737px; background: white;"">" 'takes of place of #container
    Print #1, "  <div class=""content"" style=""overflow: hidden;"" width=""737"">"
    Print #1, "  <div id=""item-content-pane"">"
    
    Print #1, "   <p><a href=""" & DESTINATION_DIR & "extspec.html"">[open preview in new window]</a></p>"
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM vWebOffloadBackend WHERE ItemNumber='" & item & "'")
    If rst.EOF Then
        Print #1, "<p>couldn't load record, that's weird</p>"
    Else
        'TODO: specials
        'TODO: group items
        'Print #1, GenerateItemPageContents(rst, New Dictionary, New Dictionary, New Dictionary, "", -1, "", False)
        'Print #1, GenerateItemPageContents(rst, New Dictionary, New Dictionary, New Dictionary)
        SpecialsCache_Initialize item
        Print #1, GenerateItemPageContents(rst)
        SpecialsCache_Clear
    End If
    Print #1, "   <script type=""text/javascript"">TP.Tabs.zip();</script>"
    
    Print #1, "  </div>" 'item content pane
    
    Print #1, "   <br>"
    'cross sells
    Dim i As Long
    If Me.CrossSellList.ListCount > 0 Then
        Print #1, "   <div class=""general-section-header"">Check out the rest of our...</div>"
        Print #1, "   <ul id=""related-sections"">"
        For i = 0 To Me.CrossSellList.ListCount - 1
            Print #1, "    <li><a href=""" & Me.CrossSellList.GetText(i, 1) & """>" & Me.CrossSellList.GetText(i, 0) & "</a></li>"
        Next i
        Print #1, "   </ul>"
    End If
    
    Print #1, "  </div>"
    Print #1, "  </div>"
    Print #1, "  </div>"
    Print #1, " </body>"
    Print #1, "</html>"
    Close #1
    Me.ExtendedSpecsPreview.Navigate2 DESTINATION_DIR & "extspec.html"
End Sub

Private Sub lockForm()
    Dim i As Long
    'Me.btnFilterByForm.Enabled = False
    Me.btnOptionsDialog.Enabled = False
    Me.btnWebPathCopy.Enabled = False
    Me.btnWebPathPaste.Enabled = False
    'Me.btnAccessoriesCopy.Enabled = False
    'Me.btnAccessoriesPaste.Enabled = False
    'Me.btnAccessoriesAddToFilter.Enabled = False
    'Me.btnOpenAccessorizer.Enabled = False
    Me.btnWebPathEdit.Enabled = False
    Me.btnWebPathCopy.Enabled = False
    Me.btnWebPathPaste.Enabled = False
    'Me.btnAddAccessories.Enabled = False
    'Me.btnAddAccessoriesReverse.Enabled = False
    'Me.generalFrame(9).Enabled = False  'group signs
    'Me.btnGroupSignFilter.Enabled = False
    'Me.btnPreviewGroupSign.Enabled = False
    Me.generalFrame(8).Enabled = False  'availability
    Me.generalFrame(2).Enabled = False  'pub flags
    'Me.btnEditWebPaths.Enabled = False
    'Me.btnEditWebSigns.Enabled = False
    'Me.btnEditShipClass.Enabled = False
    Me.generalFrame(7).Enabled = False  'specials
    Me.CrossSellFrame.Enabled = False
    Me.IncludesFrame.Enabled = False
    Me.btnSpecialsChange.Enabled = False
    Me.btnSpecialsFilter.Enabled = False
    For i = Me.Desc.LBound To Me.Desc.UBound
        Me.Desc(i).Locked = True
    Next i
    For i = Me.Spec.LBound To Me.Spec.UBound
        Me.Spec(i).Locked = True
    Next i
    For i = Me.ExtendedSpecs.LBound To Me.ExtendedSpecs.UBound
        Me.ExtendedSpecs(i).Locked = True
    Next i
    'Me.cmbSpecialChars.Locked = True
    Me.InternalNotes.Enabled = False
    Me.chkFrame.Enabled = False
    'Me.chkBaseItem.Enabled = False
    'Me.baseItemFrame.Enabled = False
    Me.generalFrame(11).Enabled = False 'triad code
End Sub

'Private Sub baseItemFrameToggle(onoff As Boolean, Optional skipfill As Boolean = False)
'    Dim i As Long
'    If onoff Then
'        Enable Me.baseItemList
'        Me.btnBaseItemEdit.Enabled = True
'        Me.btnBaseItemNew.Enabled = True
'        fillingForm = True
'        Me.chkBaseItem = SQLBool(True)
'        fillingForm = False
'        If Not skipfill Then
'            For i = 1 To 4
'                Me.Desc(i) = FillBaseItemSpecs(Me.ItemNumber, i)
'                Me.Desc(i).Enabled = False
'            Next i
'            For i = 1 To 8
'                Me.Spec(i) = FillBaseItemSpecs(Me.ItemNumber, i)
'                Me.Spec(i).Enabled = False
'            Next i
'        End If
'    Else
'        Me.baseItemList = ""
'        Disable Me.baseItemList
'        Me.btnBaseItemEdit.Enabled = False
'        Me.btnBaseItemNew.Enabled = False
'        fillingForm = True
'        Me.chkBaseItem = SQLBool(False)
'        fillingForm = False
'        If Not skipfill Then
'            Dim rst As ADODB.Recordset
'            Set rst = DB.retrieve("EXEC spPartNumbersByItem '" & Me.ItemNumber & "'")
'            For i = 1 To 4
'                Me.Desc(i) = Nz(rst("Desc" & i))
'                Me.Desc(i).Enabled = True
'            Next i
'            For i = 1 To 8
'                Me.Spec(i) = Nz(rst("Spec" & i & "HTML"))
'                Me.Spec(i).Enabled = True
'            Next i
'            rst.Close
'            Set rst = Nothing
'        End If
'    End If
'End Sub





'----------------
' db handling
'----------------
Private Sub Accessories_DblClick()
    'If Not lockedDown Then
    '    btnAddAccessories_Click
    'End If
    If HasPermissionsTo("Accessorize") Then
        btnAddAccessories_Click
    End If
End Sub

Private Sub Accessories_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyC And Shift = vbCtrlMask Then
        If Me.Accessories <> "" Then
            ClipBoard_SetData "Accessory|" & Me.Accessories
        End If
    ElseIf KeyCode = vbKeyV And Shift = vbCtrlMask Then
        Dim Clipboard As String
        Clipboard = ClipBoard_GetData
        If Left(Clipboard, 10) = "Accessory|" Then
            Dim acc As String
            acc = Mid(Clipboard, InStr(Clipboard, "|") + 1)
            If acc <> Me.ItemNumber And ExistsInInventoryMaster(acc) Then
                DB.Execute "INSERT INTO PartNumberAccessories ( ItemNumber, Accessory ) VALUES ( '" & Me.ItemNumber & "', '" & acc & "' )"
                requeryAccessories Me.ItemNumber
            End If
        Else
            MsgBox "Invalid clipboard contents."
        End If
    End If
End Sub

Private Sub Availability_Click()
    changed = True
    Availability_LostFocus
End Sub

Private Sub Availability_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.Availability, KeyCode, Shift
    If KeyCode = vbKeyC And Shift = vbCtrlMask Then
        ClipBoard_SetData (Me.Availability)
    ElseIf KeyCode = vbKeyX And Shift = vbCtrlMask Then
        ClipBoard_SetData (Me.Availability)
        Me.Availability = ""
    ElseIf KeyCode = vbKeyV And Shift = vbCtrlMask Then
        Me.Availability = ClipBoard_GetData
        changed = True
        whichCtl = "Availability"
    End If
    Select Case KeyCode
        Case Is = vbKeyReturn
            Availability_LostFocus
        Case Is = vbKeyDelete
            Availability_KeyPress KeyCode
    End Select
End Sub

Private Sub Availability_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.Availability, KeyAscii
    changed = True
    whichCtl = "Availability"
End Sub

Private Sub Availability_LostFocus()
    AutoCompleteLostFocus Me.Availability
    If changed = False Then
        Exit Sub
    End If
    If Me.Availability.Text = "Regular availability" Then
        updatePartNumbers "AvailabilityOverride", "Null", Me.ItemNumber, ""
    Else
        If IsItemDSable(Me.ItemNumber) Then
            If 0 = DLookup("COALESCE(Orderable,1)", "AvailOverrides", "ID=" & AvailHashStrToID.item(Me.Availability.Text)) Then
                MsgBox "WARNING!" & vbCrLf & vbCrLf & "This will override the dropshippable flag, so this item will not be orderable."
            End If
        End If
        updatePartNumbers "AvailabilityOverride", AvailHashStrToID.item(Me.Availability.Text), Me.ItemNumber, ""
    End If
    fillingForm = True
    Me.chkWebOrderable = SQLBool(ShouldBeWebOrderableForItem(Me.ItemNumber))
    fillingForm = False
    Me.lblAvailability.Caption = GetAvailabilityTextForItem(Me.ItemNumber)
    changed = False
End Sub

Private Sub AvailabilityLimit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AvailabilityLimit_LostFocus
    Else
        changed = True
        whichCtl = "AvailabilityLimit"
    End If
End Sub

Private Sub AvailabilityLimit_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "AvailabilityLimit"
End Sub

Private Sub AvailabilityLimit_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.AvailabilityLimit) Then
            updatePartNumbers "AvailabilityLimit", Me.AvailabilityLimit, Me.ItemNumber, ""
            changed = False
            'CheckInStockOrBackOrder Me.ItemNumber
            CheckInStockOrBackOrder2 Me.ItemNumber
            Dim statuscode As String, qoh As Long
            statuscode = DLookup("ItemStatus", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'")
            qoh = DLookup("QtyOnHand", "vActualWhseQty", "ItemNumber='" & Me.ItemNumber & "'")
            If qoh <= Me.AvailabilityLimit And IsItemDC(statuscode) Then
                If MsgBox("Discontinue this item on the web?", vbYesNo) = vbYes Then
                    SetWebDiscontinued Me.ItemNumber
                    refreshForm
                End If
            ElseIf qoh > Me.AvailabilityLimit And IsItemDC(statuscode) Then
                MsgBox "Un-discontinuing this item...give it a price before offloading!!"
                SetUnDiscontinued Me.ItemNumber
                refreshForm
            End If
        Else
            MsgBox "Availability limit must be an integer."
            Me.AvailabilityLimit.SetFocus
        End If
    End If
End Sub

Private Sub EBayAvailabilityLimit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        EBayAvailabilityLimit_LostFocus
    Else
        changed = True
        whichCtl = "EBayAvailabilityLimit"
    End If
End Sub

Private Sub EBayAvailabilityLimit_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "EBayAvailabilityLimit"
End Sub

Private Sub EBayAvailabilityLimit_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.EBayAvailabilityLimit) Then
            updatePartNumbers "EBayReserveQty", Me.EBayAvailabilityLimit, Me.ItemNumber, ""
            changed = False
            'CheckInStockOrBackOrder Me.ItemNumber
            CheckInStockOrBackOrder2 Me.ItemNumber
        Else
            MsgBox "EBay availability limit must be an integer."
            Me.EBayAvailabilityLimit.SetFocus
        End If
    End If
End Sub

'Private Sub BaseItemList_Click()
'    changed = True
'    BaseItemList_LostFocus
'End Sub
'
'Private Sub BaseItemList_KeyDown(KeyCode As Integer, Shift As Integer)
'    AutoCompleteKeyDown Me.baseItemList, KeyCode, Shift
'    Select Case KeyCode
'        Case Is = vbKeyDelete
'            BaseItemList_KeyPress KeyCode
'        Case Is = vbKeyReturn
'            BaseItemList_LostFocus
'    End Select
'End Sub
'
'Private Sub BaseItemList_KeyPress(KeyAscii As Integer)
'    AutoCompleteKeyPress Me.baseItemList, KeyAscii
'    changed = True
'    whichCtl = "BaseItemList"
'End Sub
'
'Private Sub BaseItemList_LostFocus()
'    AutoCompleteLostFocus Me.baseItemList
'    If changed = True Then
'        If Me.baseItemList = "" Then
'            DB.Execute "UPDATE PartNumbers SET BaseItemID=0 WHERE ItemNumber='" & Me.ItemNumber & "'"
'            baseItemFrameToggle False
'        Else
'            DB.Execute "UPDATE PartNumbers SET BaseItemID=(SELECT ID FROM BaseItemList WHERE BaseItem='" & Me.baseItemList & "') WHERE ItemNumber='" & Me.ItemNumber & "'"
'        End If
'        changed = False
'    End If
'End Sub

Private Sub chkMetaNoIndex_Click()
    If Not fillingForm Then
        Dim doChange As Boolean
        doChange = False
        If META_NOINDEX_MOLLY_ON = False Then
            doChange = True
        Else
            If Me.chkMetaNoIndex.value = vbChecked Then
                If vbYes = MsgBox("Really set this to be not indexable by google?!?", vbYesNo) Then
                    doChange = True
                End If
            Else
                doChange = True
            End If
        End If
        
        If doChange Then
            updatePartNumbers "MetaNoIndex", IIf(Me.chkMetaNoIndex, "1", "0"), Me.ItemNumber, ""
        Else
            fillingForm = True
            Me.chkMetaNoIndex.value = IIf(Me.chkMetaNoIndex, "0", "1")
            fillingForm = False
        End If
    End If
End Sub

Private Sub chkPrintSign_Click()
    If Not fillingForm Then
        updatePartNumbers "PrintSign", IIf(Me.chkPrintSign, "1", "0"), Me.ItemNumber, ""
    End If
End Sub

Private Sub chkPrintSignsByQOH_Click()
    If Not fillingForm Then
        updatePartNumbers "SignByQOH", IIf(Me.chkPrintSignsByQOH, "1", "0"), Me.ItemNumber, ""
    End If
End Sub

Private Sub chkPublished_Click()
    If Not fillingForm Then
        Mouse.Hourglass True
        If Me.chkPublished Then
            'If Me.chkBaseItem Then
            '    If "0" = DLookup("COUNT(*)", "PartNumberWebPaths", "ItemNumber='" & Me.ItemNumber & "'") Then
            '        MsgBox "No web paths, not publishing."
            '    Else
            '        updatePartNumbers "WebPublished", "1", Me.ItemNumber, ""
            '        updatePartNumbers "WebToBePublished", "0", Me.ItemNumber, ""
            '        'updatePartNumbers "WebPointered", "0", Me.ItemNumber, ""
            '        statusChangeLog Me.ItemNumber, "Published", "checked"
            '        fillingForm = True
            '        Me.chkToBePublished = False
            '        Me.chkWebPointer = False
            '        updatePartNumbers "TBPubPriority", "0", Me.ItemNumber, ""
            '        Me.TBPubPriority.Visible = False
            '        fillingForm = False
            '    End If
            'Else
                Dim rst As ADODB.Recordset, rst2 As ADODB.Recordset, rst3 As ADODB.Recordset, conflictingItem As String, isGroupMaster As Boolean, isGroupSub As Boolean
                Set rst = DB.retrieve("SELECT IDiscountMarkupPriceRate1 FROM InventoryMaster WHERE ItemNumber='" & Me.ItemNumber & "'")
                Set rst2 = DB.retrieve("SELECT COUNT(*) AS NumWP FROM PartNumberWebPaths WHERE ItemNumber='" & Me.ItemNumber & "'")
                Set rst3 = DB.retrieve("SELECT TotalWeight FROM vInventoryComponentsDisplayable WHERE ItemNumber='" & Me.ItemNumber & "'")
                conflictingItem = CheckConflictingItem(Me.ItemNumber)
                isGroupMaster = IsGroupedMasteritem(Me.ItemNumber)
                isGroupSub = IsGroupedSubitem(Me.ItemNumber)
                If conflictingItem <> "" Then
                    MsgBox "Name conflicts w/ published item " & qq(conflictingItem) & ", you need to web delete this item first!"
                ElseIf isGroupSub = False And rst2("NumWP") = 0 Then
                    MsgBox "No web paths, not publishing."
                ElseIf rst.EOF Then
                    MsgBox "Could not find part in Po/Inv, not publishing."
                ElseIf isGroupMaster = False And rst3("TotalWeight") = 0 Then
                    MsgBox "Invalid weight, not publishing."
                ElseIf isGroupMaster = False And (rst("IDiscountMarkupPriceRate1") = 0 Or rst("IDiscountMarkupPriceRate1") = 88888.88) Then
                    MsgBox "Invalid price, not publishing."
                Else
                    updatePartNumbers "WebPublished", "1", Me.ItemNumber, ""
                    updatePartNumbers "WebToBePublished", "0", Me.ItemNumber, ""
                    'updatePartNumbers "WebPointered", "0", Me.ItemNumber, ""
                    statusChangeLog Me.ItemNumber, "Published", "checked"
                    'If MsgBox("Update the published date?", vbYesNo) = vbYes Then
                        updatePartNumbers "PublishedDate", "GETDATE()", Me.ItemNumber, ""
                    'End If
                    If isGroupMaster = False Then
                        If MsgBox("Add to shopping engines?", vbYesNo) = vbYes Then
                            AddItemKeyword Me.ItemNumber, True, True
                        Else
                            AddItemKeyword Me.ItemNumber, True, False
                        End If
                    Else
                        AddItemKeyword Me.ItemNumber, True, False
                    End If
                End If
                rst.Close
                rst2.Close
                Set rst = Nothing
                Set rst2 = Nothing
                refreshForm
            'End If
        Else
            updatePartNumbers "WebPublished", "0", Me.ItemNumber, ""
            statusChangeLog Me.ItemNumber, "Published", "unchecked"
        End If
        Mouse.Hourglass False
    End If
End Sub

Private Sub chkRequestForm_Click()
    If Not fillingForm Then
        updatePartNumbers "ShowRequestForm", IIf(Me.chkRequestForm, "1", "0"), Me.ItemNumber, ""
        statusChangeLog Me.ItemNumber, "Request", IIf(Me.chkRequestForm, "checked", "unchecked")
        If Me.chkRequestForm Then
            updatePartNumbers "WebOrderable", "0", Me.ItemNumber, ""
            fillingForm = True
            Me.chkWebOrderable = 0
            fillingForm = False
        End If
    End If
End Sub

Private Sub chkToBePublished_Click()
    If Not fillingForm Then
        updatePartNumbers "WebToBePublished", IIf(Me.chkToBePublished, "1", "0"), Me.ItemNumber, ""
        statusChangeLog Me.ItemNumber, "ToBePublished", IIf(Me.chkToBePublished, "checked", "unchecked")
        If Me.chkToBePublished Then
            Me.TBPubPriority.Visible = True
        Else
            Me.TBPubPriority.Visible = False
            updatePartNumbers "TBPubPriority", "0", Me.ItemNumber, ""
        End If
    End If
End Sub

Private Sub chkWebPointer_Click()
    If Not fillingForm Then
        If Me.chkWebPointer Then
            
            If vbNo = MsgBox("Are you sure you want to turn this into a template for a group item page?", vbYesNo) Then
                fillingForm = True
                Me.chkWebPointer = 0
                fillingForm = False
                Exit Sub
            End If
            
            Dim pageURL1 As String
            pageURL1 = InputBox("Enter the url to use for this page. Make sure it doesn't conflict with any other pages, because I don't check for that!", "Group Item Setup", CreateYahooID(Me.ItemNumber))
            If pageURL1 = "" Then
                fillingForm = True
                Me.chkWebPointer = 0
                fillingForm = False
                Exit Sub
            End If
            
            Dim pageURL2 As String
            pageURL2 = URLify(pageURL1)
            If pageURL1 <> pageURL2 Then
                If vbNo = MsgBox(qq(pageURL1) & " will be url-ified as " & qq(pageURL2) & ", is this okay?", vbYesNo) Then
                    fillingForm = True
                    Me.chkWebPointer = 0
                    fillingForm = False
                    Exit Sub
                End If
            End If
            
            Dim rst As ADODB.Recordset
            Set rst = DB.retrieve("SELECT COUNT(*) FROM WebPaths WHERE dbo.WebPathCompleteURL(ID)='" & pageURL2 & "'")
            If rst(0) = 1 Then
                MsgBox "URL " & qq(pageURL2) & " conflicts with a section page! Talk to Kirk or Brian!"
                rst.Close
                Set rst = Nothing
                Exit Sub
            End If
            rst.Close
            Set rst = Nothing
            
            updatePartNumbers "WebPointered", "1", Me.ItemNumber, ""
            statusChangeLog Me.ItemNumber, "Pointered", "checked"
            
            DB.Execute "INSERT INTO PartNumberGroupMaster ( GroupPageName, ItemNumberPointer ) VALUES ( '" & pageURL2 & "', '" & Me.ItemNumber & "' )"
            updateInventoryMaster "Hide", "1", Me.ItemNumber, ""
            'MsgBox "Turned this into a template, delete it from Mas if it's already been exported!"
            
            Me.btnGroupEditor.Visible = True
        Else
            MsgBox "Can't un-template an item! Talk to Brian!"
            fillingForm = True
            Me.chkWebPointer = 1
            fillingForm = False
        End If
    End If
End Sub

Private Sub chkWebDelete_Click()
    If Not fillingForm Then
        updatePartNumbers "WebDelete", IIf(Me.chkWebDelete, "1", "0"), Me.ItemNumber, ""
        statusChangeLog Me.ItemNumber, "Delete?", IIf(Me.chkWebDelete, "checked", "unchecked")
    End If
End Sub

Private Sub chkWebDeleted_Click()
    If Not fillingForm Then
        If Me.chkWebDeleted Then
            updatePartNumbers "WebDeleted", "1", Me.ItemNumber, ""
            updatePartNumbers "WebPublished", "0", Me.ItemNumber, ""
            updatePartNumbers "WebToBePublished", "0", Me.ItemNumber, ""
            'updatePartNumbers "WebPointered", "0", Me.ItemNumber, ""
            updatePartNumbers "WebDelete", "0", Me.ItemNumber, ""
            statusChangeLog Me.ItemNumber, "Deleted", "checked"
            If Me.TBPubPriority.Visible Then
                Me.TBPubPriority.Visible = False
                updatePartNumbers "TBPubPriority", "0", Me.ItemNumber, ""
            End If
            fillingForm = True
            Me.chkPublished = 0
            Me.chkToBePublished = 0
            Me.chkWebPointer = 0
            Me.chkWebDelete = 0
            fillingForm = False
            MsgBox "You still need to delete this part from Yahoo."
        Else
            updatePartNumbers "WebDeleted", "0", Me.ItemNumber, ""
            statusChangeLog Me.ItemNumber, "Deleted", "unchecked"
        End If
    End If
End Sub

Private Sub chkWebOrderable_Click()
    If Not fillingForm Then
        updatePartNumbers "WebOrderable", IIf(Me.chkWebOrderable, "1", "0"), Me.ItemNumber, ""
        statusChangeLog Me.ItemNumber, "Orderable", IIf(Me.chkWebOrderable, "checked", "unchecked")
        If Me.chkWebOrderable Then
            updatePartNumbers "ShowRequestForm", "0", Me.ItemNumber, ""
            fillingForm = False
            Me.chkRequestForm = 0
            fillingForm = False
        End If
    End If
End Sub

Private Sub Desc_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.Desc(index).selStart = 0
                Me.Desc(index).SelLength = Len(Me.Desc(index))
                KeyCode = 0
                Shift = 0
            End If
        Case Is = vbKeyReturn
            Desc_LostFocus index
        Case Is = vbKeyDelete
            Desc_KeyPress index, KeyCode
    End Select
End Sub

Private Sub Desc_KeyPress(index As Integer, KeyAscii As Integer)
    changed = True
    whichCtl = "Desc" & index
End Sub

Private Sub Desc_LostFocus(index As Integer)
    If changed = True Then
        Me.Desc(index) = UnFrontpage(Me.Desc(index))
        Me.Desc(index) = StripHTML(Me.Desc(index))
        Me.Desc(index) = TrimWhitespace(Me.Desc(index), True, True, True)
        If validateDescription(Me.Desc(index)) Then
            updatePartNumbers "Desc" & index, Me.Desc(index), Me.ItemNumber, "'"
            changed = False
            fixTitleTextLengthField
        Else
            MsgBox "Invalid item description, too long or invalid characters."
        End If
    End If
End Sub

Private Sub ExtendedSpecs_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.ExtendedSpecs(index).selStart = 0
                Me.ExtendedSpecs(index).SelLength = Len(Me.ExtendedSpecs(index))
                KeyCode = 0
                Shift = 0
            Else
                ExtendedSpecs_KeyPress index, KeyCode
            End If
        Case Is = vbKeyReturn
            'ExtendedSpecs_LostFocus
        Case Is = vbKeyDelete
            ExtendedSpecs_KeyPress index, KeyCode
    End Select
End Sub

Private Sub ExtendedSpecs_KeyPress(index As Integer, KeyAscii As Integer)
    changed = True
    whichCtl = "ExtendedSpecs" & index
End Sub

Private Sub ExtendedSpecs_LostFocus(index As Integer)
    If changed = True And index <> Me.chkTabStrip.UBound Then
        Me.ExtendedSpecs(index) = UnFrontpage(Me.ExtendedSpecs(index))
        While Left(Me.ExtendedSpecs(index), 2) = vbCrLf
            Me.ExtendedSpecs(index) = Mid(Me.ExtendedSpecs(index), 3)
        Wend
        While Right(Me.ExtendedSpecs(index), 2) = vbCrLf
            Me.ExtendedSpecs(index) = Left(Me.ExtendedSpecs(index), Len(Me.ExtendedSpecs(index)) - 2)
        Wend
        If Len(Me.ExtendedSpecs(index)) > 0 _
          And InStr(Me.ExtendedSpecs(index), vbCrLf) = 0 _
          And Left(Me.ExtendedSpecs(index), 1) <> "<" Then
            Me.ExtendedSpecs(index) = "<p>" & Me.ExtendedSpecs(index) & "</p>"
        End If
        Select Case index
            Case Is = 0
                updatePartNumbers "ExtendedSpecsHTML", Me.ExtendedSpecs(index), Me.ItemNumber, "'"
            Case Is = 1
                updatePartNumbers "WriteupHTML", Me.ExtendedSpecs(index), Me.ItemNumber, "'"
            Case Is = 2
                updatePartNumbers "FeaturesHTML", Me.ExtendedSpecs(index), Me.ItemNumber, "'"
            Case Is = 3
                updatePartNumbers "TechSpecsHTML", Me.ExtendedSpecs(index), Me.ItemNumber, "'"
            Case Is = 4
                updatePartNumbers "MediaHTML", Me.ExtendedSpecs(index), Me.ItemNumber, "'"
            Case Is = 5
                updatePartNumbers "NotesHTML", Me.ExtendedSpecs(index), Me.ItemNumber, "'"
            'Case Is = 6
            '    updatePartNumbers "QuestionsHTML", Me.ExtendedSpecs(Index), Me.ItemNumber, "'"
        End Select
        changed = False
    End If
End Sub

Private Sub Spec_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.Spec(index).selStart = 0
                Me.Spec(index).SelLength = Len(Me.Spec(index))
                KeyCode = 0
                Shift = 0
            End If
        Case Is = vbKeyReturn
            Spec_LostFocus index
        Case Is = vbKeyDelete
            Spec_KeyPress index, KeyCode
    End Select
End Sub

Private Sub Spec_KeyPress(index As Integer, KeyAscii As Integer)
    changed = True
    whichCtl = "Spec" & index
End Sub

Private Sub Spec_LostFocus(index As Integer)
    If changed = True Then
        Me.Spec(index) = TrimWhitespace(Me.Spec(index), True, True, True)
        Me.Spec(index) = UnFrontpage(Me.Spec(index))
        If validateSpecs(Me.Spec(index)) Then
            updatePartNumbers "Spec" & index & "HTML", Me.Spec(index), Me.ItemNumber, "'"
            changed = False
        Else
            MsgBox "Invalid item spec, must be less than 512 chars."
        End If
    End If
End Sub

Private Sub TBPubPriority_Click()
    Me.TBPubPriority.tag = (Me.TBPubPriority.tag + 1) Mod 5
    updatePartNumbers "TBPubPriority", Me.TBPubPriority.tag, Me.ItemNumber, ""
    Select Case Me.TBPubPriority.tag
        Case Is = 0
            Me.TBPubPriority.FillColor = BG_GREY
        Case Is = 1
            Me.TBPubPriority.FillColor = RED
        Case Is = 2
            Me.TBPubPriority.FillColor = YELLOW
        Case Is = 3
            Me.TBPubPriority.FillColor = GREEN
        Case Is = 4
            Me.TBPubPriority.FillColor = LT_BLUE
    End Select
End Sub

Private Sub TBPubPriority_DblClick()
    TBPubPriority_Click
End Sub

Private Sub TriadCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            TriadCode_LostFocus
        Case Is = vbKeyDelete
            TriadCode_KeyPress KeyCode
    End Select
End Sub

Private Sub TriadCode_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "TriadCode"
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
        Exit Sub
    End If
    If Len(Me.TriadCode) - Me.TriadCode.SelLength >= 3 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TriadCode_LostFocus()
    If changed = True Then
        updateInventoryMaster "TriadCode", Me.TriadCode, Me.ItemNumber, "'"
        changed = False
    End If
End Sub

Private Sub txtPosition_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            txtPosition_LostFocus
        Case Is = vbKeyDelete
            txtPosition_KeyPress KeyCode
        Case Else
            If IsAlphaNumeric(Chr(KeyCode)) And Not IsNumeric(Chr(KeyCode)) Then
                KeyCode = 0
            End If
    End Select
End Sub

Private Sub txtPosition_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "txtPosition"
End Sub

Private Sub txtPosition_LostFocus()
    If changed = True Then
        If IsNumeric(Me.txtPosition) Then
            'ok
            'If Me.txtPosition - 1 < Me.barPosition.Min Then
            If Me.txtPosition - 1 < 0 Then
                'Me.txtPosition = Me.barPosition.Min
                Me.txtPosition = Me.CurrentStartRecord
            'ElseIf Me.txtPosition - 1 > Me.barPosition.max Then
            ElseIf Me.txtPosition - 1 > Me.CurrentEndRecord Then
                'Me.txtPosition = Me.barPosition.max
                Me.txtPosition = Me.CurrentEndRecord
            End If
            changed = False
            'Me.barPosition.value = Me.txtPosition - 1
            Me.CurrentOffset = Me.txtPosition - 1
        Else
            MsgBox "Position is not numeric!"
            'Me.txtPosition = Me.barPosition.value
            Me.txtPosition = Me.CurrentOffset
            changed = False
        End If
    End If
    barPosition_Change
End Sub

Private Sub InternalNotes_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.InternalNotes.selStart = 0
                Me.InternalNotes.SelLength = Len(Me.InternalNotes)
                KeyCode = 0
                Shift = 0
            End If
        Case Is = vbKeyReturn
            InternalNotes_LostFocus
        Case Is = vbKeyDelete
            InternalNotes_KeyPress KeyCode
    End Select
End Sub

Private Sub InternalNotes_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "InternalNotes"
End Sub

Private Sub InternalNotes_LostFocus()
    If changed = True Then
        updatePartNumbers "InternalNotes", Me.InternalNotes, Me.ItemNumber, "'"
        changed = False
    End If
End Sub


'---------------------------------
' EBAY
'---------------------------------

Public Sub EBayCategorySet(category As String, categoryType As Long)
    Dim ctl As TextBox
    Select Case categoryType
        Case Is = 0
            Set ctl = Me.EBayCategory
        Case Is = 1
            Set ctl = Me.EBayStoreCategory
        Case Is = 2
            Set ctl = Me.SecondaryCategoryTxt
    End Select
    
    ctl.Text = Replace(Replace(category, "Business & Industrial ", "B&I"), "Home & Garden > Tools ", "H&G>T")
End Sub

Public Sub EBayCategorySetByID(id As String, categoryType As Long)
    Dim ctl As TextBox
    Select Case categoryType
        Case Is = 0
            Set ctl = Me.EBayCategory
        Case Is = 1
            Set ctl = Me.EBayStoreCategory
    End Select
    
    If id = -1 Then
        ctl.Text = "<none>"
    Else
        If EBayCategoryHashIDtoName.exists(CStr(id)) Then
            EBayCategorySet CStr(EBayCategoryHashIDtoName.item(CStr(id))(1)), categoryType
        Else
            ctl.Text = "<lookup error>"
        End If
    End If
End Sub

Private Sub chkEBayPublished_Click()
    If Not fillingForm Then
        If Me.chkEBayPublished Then
            Dim errmsg As String
            errmsg = EBay.EBayIsItemPublishable(Me.ItemNumber)
            If errmsg = "OK" Then
                updatePartNumbers "EBayPublished", "1", Me.ItemNumber, ""
                updatePartNumbers "EBayToBePublished", "0", Me.ItemNumber, ""
                EBay.EBayMarkRevisionRequired Me.ItemNumber
                statusChangeLog Me.ItemNumber, "EBayPublished", "checked"
                'EBayUpdateChangeLog Me.ItemNumber, "add"
                fillingForm = True
                Me.chkEBayTBPub = 0
                fillingForm = False
            Else
                MsgBox errmsg
                fillingForm = True
                Me.chkEBayPublished = 0
                fillingForm = False
            End If
        Else
            updatePartNumbers "EBayPublished", "0", Me.ItemNumber, ""
            EBay.EBayMarkRevisionRequired Me.ItemNumber
            statusChangeLog Me.ItemNumber, "EBayPublished", "unchecked"
        End If
    End If
End Sub

Private Sub chkEBayTBPub_Click()
    If Not fillingForm Then
        If Me.chkEBayTBPub Then
            updatePartNumbers "EBayToBePublished", "1", Me.ItemNumber, ""
            statusChangeLog Me.ItemNumber, "EBayToBePublished", "checked"
            updatePartNumbers "EBayPublished", "0", Me.ItemNumber, ""
            statusChangeLog Me.ItemNumber, "EBayPublished", "unchecked"
            EBay.EBayMarkRevisionRequired Me.ItemNumber
            fillingForm = True
            Me.chkEBayPublished = 0
            fillingForm = False
        Else
            updatePartNumbers "EBayToBePublished", "0", Me.ItemNumber, ""
            statusChangeLog Me.ItemNumber, "EBayToBePublished", "unchecked"
        End If
    End If
End Sub

'Private Sub chkEBayAllowBestOffer_Click()
'    If Not fillingForm Then
'        If Me.chkEBayAllowBestOffer Then
'            updatePartNumbers "EBayAllowBestOffer", "1", Me.ItemNumber, ""
'        Else
'            updatePartNumbers "EBayAllowBestOffer", "0", Me.ItemNumber, ""
'        End If
'    End If
'End Sub



Private Sub btnEBayCategoryEdit_Click()
    Load SignMaintenanceEBayCategory
    SignMaintenanceEBayCategory.LoadItem Me.ItemNumber, 0
    SignMaintenanceEBayCategory.Show MODAL
    'category string change is handled by dialog, nothing else to do
End Sub

Private Sub btnEBayStoreCategoryEdit_Click()
    Load SignMaintenanceEBayCategory
    SignMaintenanceEBayCategory.LoadItem Me.ItemNumber, 1
    SignMaintenanceEBayCategory.Show MODAL
    'category string change is handled by dialog, nothing else to do
End Sub

Private Sub btnEBayCategorySuggest_Click()
    Dim id As String
    Dim groupMaster As String
    groupMaster = Logic.GetGroupedMasteritemForSubitem(Me.ItemNumber)
    id = EBay.EBayGetProbableEBayCategoryIDFor(IIf(groupMaster = "", Me.ItemNumber, groupMaster))
    If id = "" Then
        MsgBox "No other categorized items to suggest from!"
    Else
        DB.Execute "UPDATE PartNumbers SET EBayCategoryID=" & id & " WHERE ItemNumber='" & Me.ItemNumber & "'"
        EBayCategorySet CStr(EBayCategoryHashIDtoName.item(id)(1)), 0
    End If
End Sub

Private Sub btnEBayStoreCategorySuggest_Click()
    Dim id As String
    id = EBay.EBayGetProbableStoreCategoryIDFor(Me.ItemNumber)
    If id = "" Then
        MsgBox "No other items in this line code to pull from!"
    Else
        DB.Execute "UPDATE PartNumbers SET EBayStoreCategoryID=" & id & " WHERE ItemNumber='" & Me.ItemNumber & "'"
        EBayCategorySet CStr(EBayCategoryHashIDtoName.item(id)(1)), 1
    End If
End Sub
