VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form InventoryMaintenance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PO/INV - Build B5"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   12420
   Begin VB.CommandButton showHideCompareBtn 
      Caption         =   "Hide"
      Height          =   255
      Left            =   11640
      TabIndex        =   296
      Top             =   5280
      Width           =   615
   End
   Begin VB.ListBox productTagsLst 
      Height          =   840
      Left            =   11425
      TabIndex        =   294
      Top             =   1440
      Width           =   975
   End
   Begin VB.Timer priceCompareTmr 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   11880
      Top             =   5520
   End
   Begin VB.TextBox AvailabilityLimit 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11640
      TabIndex        =   288
      TabStop         =   0   'False
      Text            =   "-99999"
      Top             =   425
      Width           =   615
   End
   Begin VB.TextBox EBayAvailabilityLimit 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11640
      TabIndex        =   287
      TabStop         =   0   'False
      Text            =   "-99999"
      Top             =   750
      Width           =   615
   End
   Begin VB.CommandButton priceCompareBtn 
      Caption         =   "Price Compare"
      Height          =   495
      Left            =   11400
      TabIndex        =   286
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton blackListPriceComparisonsBtn 
      Caption         =   "BL"
      Height          =   195
      Left            =   11950
      TabIndex        =   284
      Top             =   6060
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton eOffBtn 
      Caption         =   "offload"
      Height          =   255
      Left            =   7800
      TabIndex        =   283
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton kitCreatorShortcut10 
      Caption         =   "10"
      Height          =   195
      Left            =   10815
      TabIndex        =   279
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton kitCreatorShortcut5 
      Caption         =   "5"
      Height          =   195
      Left            =   10575
      TabIndex        =   278
      Top             =   4200
      Width           =   255
   End
   Begin VB.CommandButton kitCreatorShortcut3 
      Caption         =   "3"
      Height          =   195
      Left            =   10335
      TabIndex        =   277
      Top             =   4200
      Width           =   255
   End
   Begin VB.CommandButton kitCreatorShortcut12 
      Caption         =   "12"
      Height          =   195
      Left            =   10815
      TabIndex        =   276
      Top             =   4005
      Width           =   375
   End
   Begin VB.CommandButton kitCreatorShortcut6 
      Caption         =   "6"
      Height          =   195
      Left            =   10575
      TabIndex        =   275
      Top             =   4005
      Width           =   255
   End
   Begin VB.CommandButton kitCreatorShortcut2 
      Caption         =   "2"
      Height          =   195
      Left            =   10335
      TabIndex        =   274
      Top             =   4005
      Width           =   255
   End
   Begin VB.ListBox EBayComparisonsLst 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   12960
      TabIndex        =   260
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Frame FrameMiscPrice 
      Height          =   1035
      Left            =   9740
      TabIndex        =   146
      Top             =   7940
      Width           =   2640
      Begin VB.OptionButton opFolliesStatus 
         Caption         =   "Flat Price"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   213
         Top             =   580
         Width           =   1155
      End
      Begin VB.OptionButton opFolliesStatus 
         Caption         =   "50% off"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   212
         Top             =   480
         Width           =   915
      End
      Begin VB.OptionButton opFolliesStatus 
         Caption         =   "40% off"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   211
         Top             =   360
         Width           =   1155
      End
      Begin VB.OptionButton opFolliesStatus 
         Caption         =   "25% off"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   210
         Top             =   240
         Width           =   915
      End
      Begin VB.OptionButton opFolliesStatus 
         Caption         =   "Not Follied"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   209
         Top             =   150
         Width           =   1155
      End
   End
   Begin VB.ListBox InternetComparisonLst 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   13560
      TabIndex        =   259
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox JetPrice 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9000
      TabIndex        =   255
      Top             =   5415
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "test"
      Height          =   255
      Left            =   13440
      TabIndex        =   253
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton btnViewPOs 
      Caption         =   "View All POs"
      Height          =   285
      Left            =   4920
      TabIndex        =   252
      TabStop         =   0   'False
      Top             =   8640
      Width           =   2000
   End
   Begin VB.ListBox POList 
      BackColor       =   &H8000000F&
      Height          =   645
      Left            =   4680
      TabIndex        =   251
      TabStop         =   0   'False
      Top             =   8040
      Width           =   2475
   End
   Begin VB.CommandButton dynamicPriceEditBtn 
      Caption         =   "edit"
      Height          =   255
      Left            =   14760
      TabIndex        =   250
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox dynamicPriceChk 
      Caption         =   "Dynamic Pricing"
      Height          =   255
      Left            =   13320
      TabIndex        =   249
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "balloon tips"
      Height          =   255
      Left            =   13200
      TabIndex        =   248
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer popupTmr 
      Enabled         =   0   'False
      Interval        =   1100
      Left            =   6480
      Top             =   480
   End
   Begin VB.Timer lineTmr 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   480
      Top             =   120
   End
   Begin VB.TextBox LoadedItemID 
      Height          =   285
      Left            =   2790
      TabIndex        =   239
      Text            =   "LoadedItemID"
      Top             =   240
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton btnPrintScreen 
      Caption         =   "Print Screen"
      Height          =   315
      Left            =   9120
      TabIndex        =   235
      Top             =   0
      Width           =   1005
   End
   Begin VB.CommandButton btnVendorQuantityCheck 
      Caption         =   "VQ"
      Height          =   255
      Left            =   3060
      TabIndex        =   232
      ToolTipText     =   "Check Quantity With Vendor"
      Top             =   630
      Width           =   405
   End
   Begin VB.CommandButton btnPriceCompareGoToEBay 
      Caption         =   "EBay"
      Height          =   255
      Left            =   7360
      TabIndex        =   231
      Top             =   9000
      Width           =   1485
   End
   Begin VB.CommandButton btnPriceCompareGoToYahoo 
      Caption         =   "Yahoo Shopping"
      Height          =   255
      Left            =   5950
      TabIndex        =   230
      Top             =   9000
      Width           =   1395
   End
   Begin VB.CommandButton btnPriceCompareGoToFroogle 
      Caption         =   "Google Products"
      Height          =   255
      Left            =   4560
      TabIndex        =   229
      Top             =   9000
      Width           =   1395
   End
   Begin PoInv.HoldDownButton btnNextItem 
      Height          =   285
      Left            =   1920
      TabIndex        =   220
      TabStop         =   0   'False
      Top             =   9330
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
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
   Begin PoInv.HoldDownButton btnPrevItem 
      Height          =   285
      Left            =   840
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   9330
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
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
   Begin VB.TextBox Website 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   218
      TabStop         =   0   'False
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton btnGoToDropshipItem 
      Caption         =   "&Z"
      Height          =   255
      Left            =   3390
      TabIndex        =   216
      ToolTipText     =   "Switch Dropship/Regular"
      Top             =   960
      Width           =   285
   End
   Begin VB.CommandButton btnWeightDimsExtendedInfo 
      Caption         =   "Extended Info..."
      Height          =   285
      Left            =   120
      TabIndex        =   196
      Top             =   4920
      Width           =   2445
   End
   Begin VB.Frame FramePricingInfo 
      Height          =   3165
      Left            =   4560
      TabIndex        =   172
      Top             =   1440
      Width           =   2535
      Begin VB.OptionButton opDropshipCostType 
         Caption         =   "E"
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   2
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   222
         TabStop         =   0   'False
         ToolTipText     =   "Calculate cost + EBay/PayPal fees (from web price)"
         Top             =   1320
         Value           =   -1  'True
         Width           =   225
      End
      Begin VB.OptionButton opDropshipCostType 
         Caption         =   "%"
         Height          =   225
         Index           =   1
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   195
         TabStop         =   0   'False
         ToolTipText     =   "Calculate cost + Dropship fees + EBay/PayPal fees (from web price)"
         Top             =   1320
         Width           =   225
      End
      Begin VB.OptionButton opDropshipCostType 
         Caption         =   "C"
         Height          =   225
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   194
         TabStop         =   0   'False
         ToolTipText     =   "Calculate dropship cost to a commerical location"
         Top             =   1320
         Width           =   225
      End
      Begin VB.TextBox DropshipCost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   193
         TabStop         =   0   'False
         Top             =   1500
         Width           =   885
      End
      Begin VB.TextBox StdPack 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   150
         Width           =   615
      End
      Begin VB.TextBox StdCost 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1530
         TabIndex        =   4
         Top             =   810
         Width           =   885
      End
      Begin VB.TextBox StdPrice 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1530
         TabIndex        =   5
         Top             =   2160
         Width           =   885
      End
      Begin VB.TextBox ListPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1530
         TabIndex        =   182
         TabStop         =   0   'False
         Top             =   2490
         Width           =   885
      End
      Begin VB.TextBox AvgCost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   181
         TabStop         =   0   'False
         Top             =   1110
         Width           =   885
      End
      Begin VB.TextBox NewCost 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   180
         TabStop         =   0   'False
         Top             =   480
         Width           =   885
      End
      Begin VB.CommandButton btnTNCforLine 
         Caption         =   "L"
         Height          =   225
         Left            =   840
         TabIndex        =   179
         TabStop         =   0   'False
         ToolTipText     =   "Apply TNC for all items in this line code"
         Top             =   540
         Width           =   255
      End
      Begin VB.CommandButton btnTNCforItem 
         Caption         =   "I"
         Height          =   225
         Left            =   570
         TabIndex        =   178
         TabStop         =   0   'False
         ToolTipText     =   "Apply TNC for this item"
         Top             =   540
         Width           =   255
      End
      Begin VB.TextBox MAPP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1530
         TabIndex        =   177
         TabStop         =   0   'False
         Top             =   2820
         Width           =   885
      End
      Begin VB.TextBox InventoriedDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   176
         TabStop         =   0   'False
         Top             =   1830
         Width           =   885
      End
      Begin VB.CommandButton btnTNCSwap 
         Caption         =   "S"
         Height          =   225
         Left            =   300
         TabIndex        =   175
         TabStop         =   0   'False
         ToolTipText     =   "Swap TNC and StdCost for this item"
         Top             =   540
         Width           =   255
      End
      Begin VB.CommandButton btnMAPPtoWP 
         Caption         =   "WP"
         Height          =   255
         Left            =   330
         TabIndex        =   174
         TabStop         =   0   'False
         ToolTipText     =   "Set MAPP as the web price"
         Top             =   2790
         Width           =   405
      End
      Begin VB.CommandButton btnItemHistory 
         Caption         =   "History"
         Height          =   285
         Left            =   120
         TabIndex        =   173
         TabStop         =   0   'False
         ToolTipText     =   "View price/cost/dc change history for this item"
         Top             =   810
         Width           =   645
      End
      Begin VB.Label DropshipCostLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "D/S Cost"
         Height          =   255
         Left            =   270
         TabIndex        =   192
         Top             =   1530
         Width           =   1215
      End
      Begin VB.Label generalLabel 
         Caption         =   "Std Pack"
         Height          =   255
         Index           =   13
         Left            =   990
         TabIndex        =   191
         Top             =   180
         Width           =   795
      End
      Begin VB.Label lblStdCost 
         Alignment       =   1  'Right Justify
         Caption         =   "Std Cost"
         Height          =   255
         Left            =   810
         TabIndex        =   190
         Top             =   840
         Width           =   675
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Store Price"
         Height          =   255
         Index           =   15
         Left            =   690
         TabIndex        =   189
         Top             =   2190
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "List Price"
         Height          =   255
         Index           =   16
         Left            =   690
         TabIndex        =   188
         Top             =   2520
         Width           =   795
      End
      Begin VB.Label lblAvgCost 
         Alignment       =   1  'Right Justify
         Caption         =   "Avg Cost"
         Height          =   255
         Left            =   810
         TabIndex        =   187
         Top             =   1140
         Width           =   675
      End
      Begin VB.Label lblTNC 
         Caption         =   "TNC"
         Height          =   255
         Left            =   1140
         TabIndex        =   186
         Top             =   510
         Width           =   375
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "MAPP"
         Height          =   255
         Index           =   17
         Left            =   720
         TabIndex        =   185
         Top             =   2820
         Width           =   495
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Inventoried"
         Height          =   255
         Index           =   14
         Left            =   420
         TabIndex        =   184
         Top             =   1860
         Width           =   1065
      End
   End
   Begin VB.CheckBox chkTaxExempt 
      Caption         =   "Tax Exempt"
      Height          =   225
      Left            =   7320
      TabIndex        =   157
      TabStop         =   0   'False
      ToolTipText     =   "Select this if the item is potentially tax exempt (actual tax exempt status will be calculated from the sale price)"
      Top             =   4560
      Width           =   1245
   End
   Begin VB.CommandButton btnNewBarcode 
      Caption         =   "A"
      Height          =   225
      Left            =   4140
      TabIndex        =   156
      TabStop         =   0   'False
      ToolTipText     =   "Add a new barcode for this item"
      Top             =   5400
      Width           =   315
   End
   Begin VB.CommandButton btnLoadBarcodes 
      Caption         =   "Barcode (X)"
      Height          =   225
      Left            =   3060
      TabIndex        =   155
      TabStop         =   0   'False
      ToolTipText     =   "Open the barcoding form for advanced changes"
      Top             =   5400
      Width           =   1065
   End
   Begin VB.CommandButton btnAlertsPending 
      BackColor       =   &H0000FFFF&
      Caption         =   "Alerts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   147
      TabStop         =   0   'False
      Top             =   9330
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton btnAddFilterToXSheet 
      Caption         =   "Add Filter"
      Height          =   225
      Left            =   10560
      TabIndex        =   144
      TabStop         =   0   'False
      ToolTipText     =   "Add all items in this line code the New Products Spreadsheet"
      Top             =   4560
      Width           =   825
   End
   Begin VB.TextBox components 
      Height          =   795
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   3240
      Width           =   3705
   End
   Begin VB.CommandButton btnFilterByForm 
      Caption         =   "&FBF"
      Height          =   315
      Left            =   9360
      TabIndex        =   107
      TabStop         =   0   'False
      ToolTipText     =   "Filter By Form"
      Top             =   420
      Width           =   405
   End
   Begin VB.Frame chkFrame 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   126
      Top             =   420
      Width           =   2355
      Begin VB.CheckBox chkStock 
         Caption         =   "Stock"
         Height          =   225
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   0
         Width           =   1065
      End
      Begin VB.CheckBox chkDropship 
         Caption         =   "Dropship"
         Height          =   225
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   240
         Width           =   1065
      End
      Begin VB.CheckBox chkDCbyUs 
         Caption         =   "D/C by Us"
         Height          =   225
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   0
         Width           =   1065
      End
      Begin VB.CheckBox chkDCbyMfr 
         Caption         =   "D/C by Mfr"
         Height          =   225
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.TextBox TriadCode 
      Height          =   255
      Left            =   10770
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   5370
      Width           =   615
   End
   Begin VB.CommandButton btnEmailItem 
      Caption         =   "Email Item"
      Height          =   255
      Left            =   30
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   330
      Width           =   915
   End
   Begin VB.CommandButton btnGoToWebsite 
      Caption         =   "&W"
      Height          =   255
      Left            =   3060
      TabIndex        =   116
      TabStop         =   0   'False
      ToolTipText     =   "Go to website"
      Top             =   960
      Width           =   285
   End
   Begin VB.CommandButton btnPriceCompareDetails 
      Caption         =   "Details"
      Height          =   255
      Left            =   10200
      TabIndex        =   115
      Top             =   9000
      Width           =   1215
   End
   Begin VB.TextBox freightActualAvg 
      Height          =   255
      Left            =   3450
      Locked          =   -1  'True
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   9030
      Width           =   1065
   End
   Begin VB.CommandButton btnReplaceForGoTo 
      Caption         =   "Replacement For"
      Height          =   225
      Left            =   90
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   5670
      Width           =   1455
   End
   Begin VB.CommandButton btnReplacedByGoTo 
      Caption         =   "Replaced By"
      Height          =   225
      Left            =   90
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton btnOptionsDialog 
      Caption         =   "Opt"
      Height          =   315
      Left            =   9780
      TabIndex        =   108
      TabStop         =   0   'False
      ToolTipText     =   "Set the cursor position"
      Top             =   420
      Width           =   405
   End
   Begin VB.TextBox ReplacementFor 
      Height          =   255
      Left            =   1560
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1395
   End
   Begin VB.CommandButton btnPriceCompare 
      Caption         =   "Price Compare"
      Height          =   255
      Left            =   13560
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton btnFreightEstimate 
      Caption         =   "Estimate Freight"
      Height          =   255
      Left            =   8850
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1335
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
      Height          =   285
      Left            =   3450
      TabIndex        =   102
      TabStop         =   0   'False
      ToolTipText     =   "Jump to last item"
      Top             =   9330
      Width           =   315
   End
   Begin VB.CommandButton btnAddLCToXSheet 
      Caption         =   "Add LC"
      Height          =   225
      Left            =   9870
      TabIndex        =   95
      TabStop         =   0   'False
      ToolTipText     =   "Add all items in this line code the New Products Spreadsheet"
      Top             =   4560
      Width           =   675
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
      Height          =   285
      Left            =   0
      TabIndex        =   94
      TabStop         =   0   'False
      ToolTipText     =   "Jump to first item"
      Top             =   9330
      Width           =   315
   End
   Begin VB.CommandButton btnGoToSigns 
      Caption         =   "Go To &Signs"
      Height          =   285
      Left            =   10230
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   9330
      Width           =   1185
   End
   Begin VB.CommandButton btnReports 
      Caption         =   "&Reports"
      Height          =   315
      Left            =   10200
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   420
      Width           =   795
   End
   Begin VB.CommandButton btnEditLineCode 
      Caption         =   "E"
      Height          =   255
      Left            =   1980
      TabIndex        =   91
      TabStop         =   0   'False
      ToolTipText     =   "Edit this Line Code's information"
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton btnAddLineCode 
      Caption         =   "A"
      Height          =   255
      Left            =   2280
      TabIndex        =   90
      TabStop         =   0   'False
      ToolTipText     =   "Add a new Line Code"
      Top             =   30
      Width           =   285
   End
   Begin VB.CommandButton btnGenerateXSheet 
      Caption         =   "Generate Spreadsheet"
      Height          =   225
      Left            =   9480
      TabIndex        =   89
      TabStop         =   0   'False
      ToolTipText     =   "View the New Products Spreadsheet"
      Top             =   4800
      Width           =   1905
   End
   Begin VB.CheckBox chkAddToXSheet 
      Caption         =   "Add Item"
      Height          =   225
      Left            =   8970
      Style           =   1  'Graphical
      TabIndex        =   88
      TabStop         =   0   'False
      ToolTipText     =   "Add this item to the New Products Spreadsheet"
      Top             =   4560
      Width           =   885
   End
   Begin VB.ComboBox cmbFilters 
      Height          =   315
      ItemData        =   "InventoryMaintenance.frx":0000
      Left            =   5730
      List            =   "InventoryMaintenance.frx":0002
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   1845
   End
   Begin VB.Frame frameInstorePrice 
      Height          =   1995
      Left            =   30
      TabIndex        =   46
      Top             =   5940
      Width           =   1900
      Begin VB.CommandButton btnCopyEbayToStore 
         Caption         =   "C. Ebay"
         Height          =   255
         Left            =   480
         TabIndex        =   240
         Top             =   0
         Width           =   765
      End
      Begin VB.CommandButton btnCopyInternet 
         Caption         =   "C. Inet"
         Height          =   255
         Left            =   1200
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   0
         Width           =   795
      End
      Begin VB.TextBox DiscountMarkupPriceRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   960
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1620
         Width           =   855
      End
      Begin VB.TextBox DiscountMarkupPriceRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   960
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1290
         Width           =   855
      End
      Begin VB.TextBox DiscountMarkupPriceRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   960
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox DiscountMarkupPriceRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   960
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   630
         Width           =   855
      End
      Begin VB.TextBox BreakQty 
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   960
         Width           =   675
      End
      Begin VB.TextBox BreakQty 
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   630
         Width           =   675
      End
      Begin VB.TextBox BreakQty 
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1620
         Width           =   675
      End
      Begin VB.TextBox BreakQty 
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1290
         Width           =   675
      End
      Begin VB.TextBox DiscountMarkupPriceRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox BreakQty 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   300
         Width           =   675
      End
      Begin VB.Label generalLabel 
         Caption         =   "In-Store"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   54
         Left            =   90
         TabIndex        =   131
         Top             =   0
         Width           =   1035
      End
      Begin VB.Label generalLabel 
         Caption         =   "1         "
         Height          =   225
         Index           =   30
         Left            =   90
         TabIndex        =   47
         Top             =   330
         Width           =   200
      End
      Begin VB.Label generalLabel 
         Caption         =   "2         "
         Height          =   225
         Index           =   31
         Left            =   90
         TabIndex        =   49
         Top             =   660
         Width           =   200
      End
      Begin VB.Label generalLabel 
         Caption         =   "3     "
         Height          =   225
         Index           =   32
         Left            =   90
         TabIndex        =   51
         Top             =   990
         Width           =   200
      End
      Begin VB.Label generalLabel 
         Caption         =   "4   "
         Height          =   225
         Index           =   33
         Left            =   90
         TabIndex        =   53
         Top             =   1320
         Width           =   200
      End
      Begin VB.Label generalLabel 
         Caption         =   "5    "
         Height          =   225
         Index           =   34
         Left            =   90
         TabIndex        =   55
         Top             =   1650
         Width           =   200
      End
   End
   Begin VB.CommandButton btnDeleteItem 
      Caption         =   "Delete Item"
      Height          =   315
      Left            =   8100
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   420
      Width           =   1125
   End
   Begin VB.Frame framePOs 
      Height          =   1035
      Left            =   7200
      TabIndex        =   78
      Top             =   7940
      Width           =   2565
      Begin VB.TextBox MinimumOrderAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   120
         TabStop         =   0   'False
         Text            =   "min ord"
         Top             =   210
         Width           =   1425
      End
      Begin VB.TextBox PrepaidAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   119
         TabStop         =   0   'False
         Text            =   "ppd"
         Top             =   480
         Width           =   1425
      End
      Begin VB.TextBox PrepaidSpecialAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   118
         TabStop         =   0   'False
         Text            =   "ppd"
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label generalLabel 
         Caption         =   "Total"
         Height          =   225
         Index           =   24
         Left            =   1470
         TabIndex        =   215
         Top             =   0
         Width           =   375
      End
      Begin VB.Label generalLabel 
         Caption         =   "Vendor"
         Height          =   225
         Index           =   23
         Left            =   150
         TabIndex        =   214
         Top             =   0
         Width           =   525
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Min. Order:"
         Height          =   225
         Index           =   50
         Left            =   120
         TabIndex        =   123
         Top             =   240
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "PPD @:"
         Height          =   225
         Index           =   51
         Left            =   120
         TabIndex        =   122
         Top             =   480
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "PPD spec:"
         Height          =   225
         Index           =   52
         Left            =   120
         TabIndex        =   121
         Top             =   750
         Width           =   795
      End
   End
   Begin VB.TextBox POComment 
      Height          =   255
      Left            =   4050
      TabIndex        =   10
      Top             =   5640
      Width           =   4005
   End
   Begin VB.TextBox ReplacedBy 
      Height          =   255
      Left            =   1560
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5370
      Width           =   1395
   End
   Begin VB.Frame frameInternetPrice 
      Height          =   1995
      Left            =   1920
      TabIndex        =   62
      Top             =   5940
      Width           =   5505
      Begin VB.CommandButton inetAmazonBtn 
         Caption         =   "A"
         Height          =   195
         Left            =   3360
         TabIndex        =   273
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton inetSortPriceBtn 
         Caption         =   "P"
         Height          =   195
         Left            =   4320
         TabIndex        =   269
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton inetSortShippingBtn 
         Caption         =   "S"
         Height          =   195
         Left            =   4680
         TabIndex        =   268
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton inetSortTotalBtn 
         Caption         =   "T"
         Height          =   195
         Left            =   5040
         TabIndex        =   267
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox InternetAlgorithmCmb 
         Height          =   315
         Left            =   1870
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   263
         Top             =   1560
         Width           =   3495
      End
      Begin MSComctlLib.ListView InternetComparisonLVI 
         Height          =   1215
         Left            =   1850
         TabIndex        =   258
         Top             =   300
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   2143
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.CommandButton btnCopyEbayToWeb 
         Caption         =   "Copy Ebay"
         Height          =   255
         Left            =   960
         TabIndex        =   241
         Top             =   30
         Width           =   885
      End
      Begin VB.CommandButton btnCopyInstore 
         Caption         =   "Copy Instore"
         Height          =   255
         Left            =   1860
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   30
         Width           =   1155
      End
      Begin VB.TextBox IDiscountMarkupPriceRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   960
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   1620
         Width           =   855
      End
      Begin VB.TextBox IDiscountMarkupPriceRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   960
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   1290
         Width           =   855
      End
      Begin VB.TextBox IDiscountMarkupPriceRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   960
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox IDiscountMarkupPriceRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   960
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   630
         Width           =   855
      End
      Begin VB.TextBox IBreakQty 
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   1620
         Width           =   675
      End
      Begin VB.TextBox IBreakQty 
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1290
         Width           =   675
      End
      Begin VB.TextBox IBreakQty 
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   960
         Width           =   675
      End
      Begin VB.TextBox IBreakQty 
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   630
         Width           =   675
      End
      Begin VB.TextBox IDiscountMarkupPriceRate 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox IBreakQty 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   300
         Width           =   675
      End
      Begin VB.Label generalLabel 
         Caption         =   "Internet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   90
         TabIndex        =   132
         Top             =   0
         Width           =   1035
      End
      Begin VB.Label generalLabel 
         Caption         =   "5"
         Height          =   225
         Index           =   44
         Left            =   90
         TabIndex        =   71
         Top             =   1650
         Width           =   200
      End
      Begin VB.Label generalLabel 
         Caption         =   "4"
         Height          =   225
         Index           =   43
         Left            =   90
         TabIndex        =   69
         Top             =   1320
         Width           =   200
      End
      Begin VB.Label generalLabel 
         Caption         =   "3 "
         Height          =   225
         Index           =   42
         Left            =   90
         TabIndex        =   67
         Top             =   990
         Width           =   200
      End
      Begin VB.Label generalLabel 
         Caption         =   "2  "
         Height          =   225
         Index           =   41
         Left            =   90
         TabIndex        =   65
         Top             =   660
         Width           =   200
      End
      Begin VB.Label generalLabel 
         Caption         =   "1"
         Height          =   225
         Index           =   40
         Left            =   120
         TabIndex        =   63
         Top             =   330
         Width           =   200
      End
   End
   Begin VB.ListBox priceComparison 
      BackColor       =   &H8000000F&
      Height          =   1425
      Left            =   13200
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   6240
      Width           =   2475
   End
   Begin VB.ListBox freightActual 
      BackColor       =   &H8000000F&
      Height          =   1035
      Left            =   2400
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   7980
      Width           =   2205
   End
   Begin VB.ListBox freightEstimates 
      BackColor       =   &H8000000F&
      Height          =   1230
      Left            =   60
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   7980
      Width           =   2265
   End
   Begin VB.ComboBox jumpToItem 
      Height          =   315
      Left            =   3810
      TabIndex        =   2
      Top             =   0
      Width           =   1485
   End
   Begin VB.ComboBox jumpToLine 
      Height          =   315
      Left            =   1170
      TabIndex        =   1
      Top             =   0
      Width           =   765
   End
   Begin VB.TextBox ItemNumber 
      Height          =   285
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   930
      Width           =   1695
   End
   Begin VB.TextBox ItemDescription 
      Height          =   285
      Left            =   1290
      TabIndex        =   3
      Top             =   1260
      Width           =   2550
   End
   Begin VB.TextBox prodLine 
      Height          =   285
      Left            =   1290
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1590
      Width           =   585
   End
   Begin VB.ComboBox PrimaryVendor 
      Height          =   315
      Left            =   2460
      TabIndex        =   7
      Top             =   1590
      Width           =   2055
   End
   Begin VB.TextBox EPN 
      Height          =   1305
      Left            =   720
      MaxLength       =   8000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1920
      Width           =   3705
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
      Left            =   7620
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Width           =   1365
   End
   Begin VB.CommandButton btnAddNewItem 
      Caption         =   "Add New Item"
      Height          =   315
      Left            =   6930
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   420
      Width           =   1185
   End
   Begin VB.TextBox txtPosition 
      Height          =   285
      Left            =   3810
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   9330
      Width           =   795
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
      Height          =   285
      Left            =   330
      TabIndex        =   82
      TabStop         =   0   'False
      ToolTipText     =   "Jump to beginning of previous line"
      Top             =   9330
      Width           =   495
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
      Height          =   285
      Left            =   2940
      TabIndex        =   84
      TabStop         =   0   'False
      ToolTipText     =   "Jump to beginning of next line"
      Top             =   9330
      Width           =   495
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   315
      Left            =   10620
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   0
      Width           =   795
   End
   Begin VB.Frame shipInfoFrame 
      Height          =   1275
      Left            =   60
      TabIndex        =   133
      Top             =   4080
      Width           =   4335
      Begin VB.OptionButton opShippingType 
         Caption         =   "Free Shipping"
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   219
         TabStop         =   0   'False
         Tag             =   "This item (and any others on the order) ships for free"
         Top             =   1020
         Width           =   1515
      End
      Begin VB.OptionButton opShippingType 
         Caption         =   "TF for $6.50"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   150
         TabStop         =   0   'False
         Tag             =   "This item ships via Truck Freight, for the regular $6.50 shipping charge"
         Top             =   810
         Width           =   1515
      End
      Begin VB.OptionButton opShippingType 
         Caption         =   "Truck Freight"
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   149
         TabStop         =   0   'False
         Tag             =   "This item ships via Truck Freight, customer pays full"
         Top             =   600
         Width           =   1515
      End
      Begin VB.OptionButton opShippingType 
         Caption         =   "Regular Shipping"
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   148
         TabStop         =   0   'False
         ToolTipText     =   "This item ships via regular UPS/Post Office service"
         Top             =   390
         Width           =   1515
      End
      Begin VB.CheckBox chkReconditioned 
         Caption         =   "Reconditioned"
         Height          =   195
         Left            =   240
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Weight 
         Height          =   285
         Left            =   900
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   150
         Width           =   735
      End
      Begin VB.TextBox Dimensions 
         Height          =   285
         Left            =   900
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   450
         Width           =   1485
      End
      Begin VB.Label conditionLbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "New"
         Height          =   255
         Left            =   2300
         TabIndex        =   285
         Top             =   120
         Width           =   2000
      End
      Begin VB.Label lblDimensionalWeight 
         Caption         =   "DW"
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
         Height          =   255
         Left            =   1650
         TabIndex        =   221
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label generalLabel 
         Caption         =   "Weight"
         Height          =   255
         Index           =   8
         Left            =   60
         TabIndex        =   137
         Top             =   180
         Width           =   735
      End
      Begin VB.Label generalLabel 
         Caption         =   "Dimensions"
         Height          =   255
         Index           =   10
         Left            =   60
         TabIndex        =   136
         Top             =   510
         Width           =   885
      End
   End
   Begin VB.Frame FrameEBayPrice 
      Height          =   1995
      Left            =   7440
      TabIndex        =   223
      Top             =   5940
      Width           =   4965
      Begin VB.CommandButton btnGoToTPP 
         Caption         =   "&P"
         Height          =   255
         Left            =   400
         TabIndex        =   282
         TabStop         =   0   'False
         ToolTipText     =   "Go to ToolsPlusParts.com website for this item"
         Top             =   1320
         Width           =   285
      End
      Begin VB.CommandButton btnGoToWeb 
         Caption         =   "&W"
         Height          =   255
         Left            =   720
         TabIndex        =   281
         TabStop         =   0   'False
         ToolTipText     =   "Go to the website for this item"
         Top             =   1320
         Width           =   285
      End
      Begin VB.CommandButton btnGoToEBay 
         Caption         =   "&E"
         Height          =   255
         Left            =   1020
         TabIndex        =   280
         TabStop         =   0   'False
         ToolTipText     =   "Go to the ebay listing for this item"
         Top             =   1320
         Width           =   285
      End
      Begin VB.CommandButton ebaySortTotalBtn 
         Caption         =   "T"
         Height          =   195
         Left            =   4200
         TabIndex        =   272
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton ebaySortShippingBtn 
         Caption         =   "S"
         Height          =   195
         Left            =   3840
         TabIndex        =   271
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton ebaySortPriceBtn 
         Caption         =   "P"
         Height          =   195
         Left            =   3480
         TabIndex        =   270
         Top             =   120
         Width           =   255
      End
      Begin VB.ComboBox EBayAlgorithmCmb 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   266
         Top             =   1560
         Width           =   3375
      End
      Begin VB.CommandButton btnCopyStoreToEbay 
         Caption         =   "Copy Store"
         Height          =   225
         Left            =   1850
         TabIndex        =   242
         Top             =   30
         Width           =   945
      End
      Begin VB.TextBox EBayBestOfferAutoAccept 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   580
         TabIndex        =   238
         Top             =   840
         Width           =   765
      End
      Begin VB.CommandButton btnEBayPriceCopyMAPP 
         Caption         =   "MAPP"
         Height          =   225
         Left            =   1250
         TabIndex        =   228
         ToolTipText     =   "Set EBay Price to the MAPP"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton btnEBayPriceCopyInternet 
         Caption         =   "Inet"
         Height          =   225
         Left            =   840
         TabIndex        =   227
         ToolTipText     =   "Set EBay Price to the current Internet Price"
         Top             =   30
         Width           =   405
      End
      Begin VB.TextBox EBayPrice 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   520
         TabIndex        =   226
         Top             =   300
         Width           =   855
      End
      Begin VB.CheckBox chkEBayAllowBestOffer 
         Caption         =   "Best Offer?"
         Height          =   255
         Left            =   120
         TabIndex        =   236
         Top             =   600
         Width           =   1125
      End
      Begin MSComctlLib.ListView EBayComparisonsLVI 
         Height          =   1215
         Left            =   1420
         TabIndex        =   261
         Top             =   300
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   2143
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label generalLabel 
         Caption         =   "Auto $"
         Height          =   255
         Index           =   21
         Left            =   105
         TabIndex        =   237
         Top             =   840
         Width           =   525
      End
      Begin VB.Label generalLabel 
         Caption         =   "Price:"
         Height          =   225
         Index           =   27
         Left            =   90
         TabIndex        =   225
         Top             =   300
         Width           =   495
      End
      Begin VB.Label generalLabel 
         Caption         =   "EBay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   90
         TabIndex        =   224
         Top             =   0
         Width           =   645
      End
   End
   Begin VB.CheckBox JetChk 
      Caption         =   "Jet.com"
      Height          =   195
      Left            =   13440
      TabIndex        =   257
      Top             =   8640
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ListView InternetDynamicAlgorithmLVI 
      Height          =   375
      Left            =   12960
      TabIndex        =   264
      Top             =   3600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView EBayDynamicAlgorithmLVI 
      Height          =   375
      Left            =   12960
      TabIndex        =   265
      Top             =   3960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame FrameSubtitle 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4560
      TabIndex        =   243
      Top             =   4680
      Width           =   6765
      Begin VB.CommandButton AddSelectedSubtitleBtn 
         Caption         =   "A"
         Height          =   255
         Left            =   6240
         TabIndex        =   247
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton DeleteSelectedSubtitleBtn 
         Caption         =   "D"
         Height          =   255
         Left            =   6480
         TabIndex        =   246
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox SubTitleChk 
         Caption         =   "Override SubTitle Field"
         Height          =   255
         Left            =   0
         TabIndex        =   244
         Top             =   0
         Width           =   2175
      End
      Begin VB.ComboBox SubtitleCmb 
         Height          =   315
         Left            =   0
         TabIndex        =   245
         Text            =   "-Choose Subtitle-"
         Top             =   360
         Width           =   6255
      End
      Begin VB.Label Label1 
         Caption         =   "Chars Remaining:"
         Height          =   255
         Left            =   3120
         TabIndex        =   293
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label subtitleCharCountLbl 
         Caption         =   "55"
         Height          =   255
         Left            =   4440
         TabIndex        =   292
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame frameSalesHistory 
      Height          =   3450
      Left            =   7200
      TabIndex        =   28
      Top             =   960
      Width           =   4245
      Begin VB.CommandButton btnEBayKitPriceManager 
         Caption         =   "Kit Pricing"
         Height          =   465
         Left            =   2760
         TabIndex        =   262
         ToolTipText     =   "Open the EBay Kit Price Manager"
         Top             =   2520
         Width           =   765
      End
      Begin VB.CommandButton btnRefreshQuantitiesForItem 
         Caption         =   "On Order"
         Height          =   255
         Left            =   90
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Refresh quantities for this item only"
         Top             =   1470
         Width           =   885
      End
      Begin VB.CommandButton btnKitCreator 
         Caption         =   "Kit Creator"
         Height          =   465
         Left            =   3540
         TabIndex        =   171
         TabStop         =   0   'False
         Top             =   2520
         Width           =   645
      End
      Begin VB.CommandButton btnViewStockBOAlerts 
         Caption         =   "On BO"
         Height          =   255
         Left            =   90
         TabIndex        =   170
         TabStop         =   0   'False
         ToolTipText     =   "View ""back in stock"" alert information"
         Top             =   1770
         Width           =   885
      End
      Begin VB.CommandButton btnSalesHistoryColumnToggle 
         Caption         =   "Combine"
         Height          =   285
         Left            =   3420
         TabIndex        =   169
         TabStop         =   0   'False
         ToolTipText     =   "Toggle between split P2 and SO information and combined overall sales info"
         Top             =   180
         Width           =   765
      End
      Begin VB.CommandButton btnSalesHistoryPeriodTypeToggle 
         Caption         =   "Rolling"
         Height          =   285
         Left            =   2700
         TabIndex        =   168
         TabStop         =   0   'False
         ToolTipText     =   "Toggle between static periods (this month, last month, etc) and rolling periods (last 30 days, etc)"
         Top             =   180
         Width           =   705
      End
      Begin VB.TextBox SalesSO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   163
         TabStop         =   0   'False
         Text            =   "SOThisMonth"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox SalesSO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   162
         TabStop         =   0   'False
         Text            =   "SOCurPer"
         Top             =   1380
         Width           =   375
      End
      Begin VB.TextBox SalesSO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   161
         TabStop         =   0   'False
         Text            =   "SOLastPer"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox SalesSO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   160
         TabStop         =   0   'False
         Text            =   "SOYTD"
         Top             =   1980
         Width           =   375
      End
      Begin VB.TextBox SalesSO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   159
         TabStop         =   0   'False
         Text            =   "SOLYR"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox SalesSO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   158
         TabStop         =   0   'False
         Text            =   "SOLastMonth"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox DropShips 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   1
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   154
         TabStop         =   0   'False
         Text            =   "ZLastMonth"
         Top             =   1080
         Width           =   285
      End
      Begin VB.TextBox Returns 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   3570
         Locked          =   -1  'True
         TabIndex        =   153
         TabStop         =   0   'False
         Text            =   "ReturnsLastMonth"
         Top             =   1080
         Width           =   285
      End
      Begin VB.TextBox SalesP2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   152
         TabStop         =   0   'False
         Text            =   "P2LastMonth"
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton btnViewInvQtyTriggers 
         Caption         =   "On Hand"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   143
         TabStop         =   0   'False
         ToolTipText     =   "Set/View inventory quantity triggers on this item"
         Top             =   870
         Width           =   885
      End
      Begin VB.TextBox DropShips 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   0
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   142
         TabStop         =   0   'False
         Text            =   "ZThisMonth"
         Top             =   780
         Width           =   285
      End
      Begin VB.TextBox DropShips 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   2
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   141
         TabStop         =   0   'False
         Text            =   "ZThisPeriod"
         Top             =   1380
         Width           =   285
      End
      Begin VB.TextBox DropShips 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   3
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   140
         TabStop         =   0   'False
         Text            =   "ZLastPeriod"
         Top             =   1680
         Width           =   285
      End
      Begin VB.TextBox DropShips 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   4
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   139
         TabStop         =   0   'False
         Text            =   "ZYTD"
         Top             =   1980
         Width           =   285
      End
      Begin VB.TextBox DropShips 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00008000&
         Height          =   255
         Index           =   5
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   138
         TabStop         =   0   'False
         Text            =   "ZLYR"
         Top             =   2280
         Width           =   285
      End
      Begin VB.TextBox Returns 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   5
         Left            =   3570
         Locked          =   -1  'True
         TabIndex        =   101
         TabStop         =   0   'False
         Text            =   "ReturnsLYR"
         Top             =   2280
         Width           =   285
      End
      Begin VB.TextBox Returns 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   4
         Left            =   3570
         Locked          =   -1  'True
         TabIndex        =   100
         TabStop         =   0   'False
         Text            =   "ReturnsYTD"
         Top             =   1980
         Width           =   285
      End
      Begin VB.TextBox Returns 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   3
         Left            =   3570
         Locked          =   -1  'True
         TabIndex        =   99
         TabStop         =   0   'False
         Text            =   "ReturnsLastPeriod"
         Top             =   1680
         Width           =   285
      End
      Begin VB.TextBox Returns 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   3570
         Locked          =   -1  'True
         TabIndex        =   98
         TabStop         =   0   'False
         Text            =   "ReturnsThisPeriod"
         Top             =   1380
         Width           =   285
      End
      Begin VB.TextBox Returns 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   3570
         Locked          =   -1  'True
         TabIndex        =   97
         TabStop         =   0   'False
         Text            =   "ReturnsThisMonth"
         Top             =   780
         Width           =   285
      End
      Begin VB.ComboBox cmbWhse 
         Height          =   315
         Left            =   180
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   180
         Width           =   2475
      End
      Begin VB.TextBox QtyBackorder 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "QtyBackorder"
         ToolTipText     =   "Customer Back Orders"
         Top             =   1800
         Width           =   705
      End
      Begin VB.TextBox QtyOnPO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "QtyOnPO"
         ToolTipText     =   "Tools Plus Purchase Orders"
         Top             =   1500
         Width           =   705
      End
      Begin VB.TextBox SalesP2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "P2LYR"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox SalesP2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "P2YTD"
         Top             =   1980
         Width           =   375
      End
      Begin VB.TextBox QtyOnSO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "QtyOnSO"
         ToolTipText     =   "Customer Sales Orders"
         Top             =   1200
         Width           =   705
      End
      Begin VB.TextBox QtyOnHand 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "QtyOnHand"
         Top             =   900
         Width           =   705
      End
      Begin VB.CommandButton btnSuggestQtyToOrder 
         Caption         =   "S"
         Height          =   255
         Left            =   1470
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Suggest a quantity to order"
         Top             =   600
         Width           =   285
      End
      Begin VB.TextBox QtyOrdered 
         Height          =   255
         Left            =   990
         TabIndex        =   6
         Top             =   600
         Width           =   465
      End
      Begin VB.TextBox SalesP2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "P2LastPer"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox SalesP2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "P2CurPer"
         Top             =   1380
         Width           =   375
      End
      Begin VB.TextBox SalesP2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "P2ThisMonth"
         Top             =   780
         Width           =   375
      End
      Begin VB.Frame FrameOPLS 
         Height          =   945
         Left            =   60
         TabIndex        =   197
         Top             =   2190
         Width           =   2685
         Begin VB.TextBox OrderPointTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   255
            Left            =   750
            Locked          =   -1  'True
            TabIndex        =   206
            TabStop         =   0   'False
            Top             =   90
            Width           =   555
         End
         Begin VB.TextBox OrderPointStore 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   750
            TabIndex        =   203
            TabStop         =   0   'False
            Top             =   390
            Width           =   555
         End
         Begin VB.TextBox OrderPointWhse 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   750
            TabIndex        =   202
            TabStop         =   0   'False
            Top             =   690
            Width           =   555
         End
         Begin VB.CheckBox chkOrderPointLockStore 
            Caption         =   "Lock"
            Height          =   255
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   201
            TabStop         =   0   'False
            ToolTipText     =   "Lock order point for the store at the specified amount"
            Top             =   390
            Width           =   585
         End
         Begin VB.CheckBox chkOrderPointLockWhse 
            Caption         =   "Lock"
            Height          =   255
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   200
            TabStop         =   0   'False
            ToolTipText     =   "Lock order point for the warehouse at the specified amount"
            Top             =   690
            Width           =   585
         End
         Begin VB.CheckBox chkStockedHereStore 
            Caption         =   "Stock"
            Height          =   255
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   199
            TabStop         =   0   'False
            ToolTipText     =   "Stock this item at the store?"
            Top             =   390
            Width           =   585
         End
         Begin VB.CheckBox chkStockedHereWhse 
            Caption         =   "Stock"
            Height          =   255
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   198
            TabStop         =   0   'False
            ToolTipText     =   "Stock this item at the warehouse?"
            Top             =   690
            Width           =   585
         End
         Begin VB.Label salesHistoryRowLabel 
            Caption         =   "Sales LYR"
            Height          =   225
            Index           =   5
            Left            =   1770
            TabIndex        =   208
            Top             =   90
            Width           =   825
         End
         Begin VB.Label generalLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Total OP"
            Height          =   225
            Index           =   9
            Left            =   -90
            TabIndex        =   207
            Top             =   90
            Width           =   795
         End
         Begin VB.Label generalLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "P2 OP"
            Height          =   225
            Index           =   12
            Left            =   -90
            TabIndex        =   205
            Top             =   420
            Width           =   795
         End
         Begin VB.Label generalLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "SO OP"
            Height          =   225
            Index           =   22
            Left            =   -90
            TabIndex        =   204
            Top             =   720
            Width           =   795
         End
      End
      Begin VB.Label salesHistoryHeaderLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "DS"
         Height          =   225
         Index           =   3
         Left            =   3930
         TabIndex        =   167
         Top             =   510
         Width           =   285
      End
      Begin VB.Label salesHistoryHeaderLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Ret"
         Height          =   225
         Index           =   2
         Left            =   3540
         TabIndex        =   166
         Top             =   510
         Width           =   285
      End
      Begin VB.Label salesHistoryHeaderLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "SO"
         Height          =   225
         Index           =   1
         Left            =   3210
         TabIndex        =   165
         Top             =   510
         Width           =   315
      End
      Begin VB.Label salesHistoryHeaderLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "P2"
         Height          =   225
         Index           =   0
         Left            =   2850
         TabIndex        =   164
         Top             =   510
         Width           =   315
      End
      Begin VB.Label salesHistoryRowLabel 
         Caption         =   "Last Month"
         Height          =   225
         Index           =   1
         Left            =   1830
         TabIndex        =   151
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "On SO"
         Height          =   225
         Index           =   20
         Left            =   120
         TabIndex        =   32
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "New Order"
         Height          =   225
         Index           =   18
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   795
      End
      Begin VB.Label salesHistoryRowLabel 
         Caption         =   "Sales YTD"
         Height          =   225
         Index           =   4
         Left            =   1830
         TabIndex        =   43
         Top             =   1980
         Width           =   825
      End
      Begin VB.Label salesHistoryRowLabel 
         Caption         =   "Last Per"
         Height          =   225
         Index           =   3
         Left            =   1830
         TabIndex        =   41
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label salesHistoryRowLabel 
         Caption         =   "Current Per"
         Height          =   225
         Index           =   2
         Left            =   1830
         TabIndex        =   39
         Top             =   1380
         Width           =   825
      End
      Begin VB.Label salesHistoryRowLabel 
         Caption         =   "This Month"
         Height          =   225
         Index           =   0
         Left            =   1830
         TabIndex        =   37
         Top             =   780
         Width           =   825
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Tags"
      Height          =   255
      Left            =   11640
      TabIndex        =   295
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label generalLabel 
      Caption         =   "EBay:"
      Height          =   255
      Index           =   37
      Left            =   11160
      TabIndex        =   291
      Top             =   780
      Width           =   465
   End
   Begin VB.Label generalLabel 
      Caption         =   "Displays:"
      Height          =   255
      Index           =   36
      Left            =   11040
      TabIndex        =   290
      Top             =   465
      Width           =   675
   End
   Begin VB.Label generalLabel 
      Caption         =   "EBay:"
      Height          =   255
      Index           =   35
      Left            =   210
      TabIndex        =   289
      Top             =   330
      Width           =   465
   End
   Begin VB.Label generalLabel 
      Caption         =   "$"
      Height          =   225
      Index           =   29
      Left            =   8880
      TabIndex        =   256
      Top             =   5445
      Width           =   525
   End
   Begin VB.Label generalLabel 
      Caption         =   "Jet.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   7920
      TabIndex        =   254
      Top             =   5400
      Width           =   1005
   End
   Begin VB.Label lblIsKit 
      Caption         =   "KIT"
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
      Height          =   225
      Left            =   3870
      TabIndex        =   234
      Top             =   1290
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblVendorQuantity 
      Caption         =   "VQ"
      Height          =   255
      Left            =   3510
      TabIndex        =   233
      Top             =   660
      Width           =   405
   End
   Begin VB.Label generalLabel 
      Caption         =   "Website:"
      Height          =   255
      Index           =   25
      Left            =   30
      TabIndex        =   217
      Top             =   630
      Width           =   1245
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Triad Code:"
      Height          =   225
      Index           =   53
      Left            =   9690
      TabIndex        =   125
      Top             =   5400
      Width           =   1065
   End
   Begin VB.Label generalLabel 
      Caption         =   "Avg:"
      Height          =   195
      Index           =   11
      Left            =   2910
      TabIndex        =   111
      Top             =   9060
      Width           =   495
   End
   Begin VB.Label lblManufName 
      Caption         =   "ManufFullNameCleaned"
      Height          =   225
      Left            =   960
      TabIndex        =   103
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label lblDCDS 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   3870
      TabIndex        =   21
      Top             =   1080
      Width           =   3225
   End
   Begin VB.Label lblPOComment 
      Caption         =   "PO Comment"
      Height          =   225
      Index           =   50
      Left            =   3060
      TabIndex        =   22
      ToolTipText     =   "Double-click to add '1 free w/ 12'"
      Top             =   5670
      Width           =   975
   End
   Begin VB.Label generalLabel 
      Caption         =   "Jump To Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   12
      Top             =   30
      Width           =   1155
   End
   Begin VB.Label generalLabel 
      Caption         =   "Jump To Line"
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
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   1245
   End
   Begin VB.Label generalLabel 
      Caption         =   "Item Number:"
      Height          =   255
      Index           =   2
      Left            =   30
      TabIndex        =   13
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label generalLabel 
      Caption         =   "Item Description:"
      Height          =   255
      Index           =   3
      Left            =   30
      TabIndex        =   15
      Top             =   1290
      Width           =   1245
   End
   Begin VB.Label generalLabel 
      Caption         =   "Product Line:"
      Height          =   255
      Index           =   4
      Left            =   30
      TabIndex        =   16
      Top             =   1620
      Width           =   1245
   End
   Begin VB.Label generalLabel 
      Caption         =   "Vendor:"
      Height          =   255
      Index           =   5
      Left            =   1890
      TabIndex        =   17
      Top             =   1620
      Width           =   585
   End
   Begin VB.Label generalLabel 
      Caption         =   "A. Info"
      Height          =   435
      Index           =   6
      Left            =   60
      TabIndex        =   18
      Top             =   1950
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Notes:"
      Height          =   285
      Index           =   7
      Left            =   30
      TabIndex        =   19
      Top             =   3180
      Width           =   585
   End
   Begin VB.Label lblStatusBar 
      Height          =   225
      Left            =   5700
      TabIndex        =   87
      Top             =   9390
      Width           =   3375
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11430
      Y1              =   9300
      Y2              =   9300
   End
   Begin VB.Label lblMaxAmt 
      Caption         =   "of MAX"
      Height          =   195
      Left            =   4650
      TabIndex        =   86
      Top             =   9390
      Width           =   975
   End
End
Attribute VB_Name = "InventoryMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public priceCompareEnabled As Boolean


'adding a new ctl
'  - add to fillForm
'  - add keydown/lostfocus handlers
'  - add to barPosition_Change
'  - add to security locking fns

Private Const NEW_PROD_XSHEET_LOCATION As String = "h:\Common Documents\New Products Spreadsheet (Empty).xls"

Private Const WHSE_ALL As String = "All Warehouses"
Private Const WHSE_ACTUAL As String = "Physical Warehouses"

Private Const SHIP_REG As Long = 0
Private Const SHIP_TF As Long = 1
Private Const SHIP_TF_CHEAP As Long = 2
Private Const SHIP_FREE As Long = 3


Private OriginalIDiscountPrice As String
Private OriginalEBayPrice As String

Private IMtt As FormTooltips

Private currentFilterItemList() As String
Private currentFilterPosition As Long

Private changed As Boolean
Private whichCtl As String
Private firstLoad As Boolean
Private fillingForm As Boolean
Dim xsheetarray() As String
Private stopLineLostFocus As Boolean
Private lockedDown As Boolean
Private lockedDownPC As Boolean
Private lockedDownFE As Boolean
Private lockedDownWD As Boolean

Public DynamicDontMessWith As Boolean

'Tom's Code 10-2-2014
'variables....
Public InvQtyTriggerLoaded As Boolean
Public InvQtyTriggerFrmX As Integer
Public InvQtyTriggerFrmY As Integer


'Tom's Code 9-29-2014 to hold balloon popup data for pop up form upgrade.
Public CurrentTNCdata As String
Public CurrentSTDdata As String
Public CurrentSTDPricedata As String
Public CurrentAVGdata As String
Public CurrentDSCostdata As String
Public CurrentINVdata As String
Public CurrentLISTdata As String
Public CurrentSTOREdata As String
Public CurrentMAPPdata As String
Public HidePopUp As Boolean
Public textBoxBackColor As Long
Public CurrentIDiscountRate As String
Public CurrentEBayPrice As String
Public currentItemNumber As String
Public CurrentDiscountRate As String
Public CurrentSubtitle As String
Public lineDoubleEnterFix As Boolean



'Tom's Code 10-15-2014 to tell if current cell has been clicked then no popup...
Public PopUpNewCost As Boolean
Public PopUpSTDCost As Boolean

'Public inetShowAmazon As Boolean
'Public inetShowFroogle As Boolean
'Public inetShowOthers As Boolean
Public orderAscending As Boolean

Public subtitleNotClicked As Boolean

Public LastComparedItem As String






'******************************************************************************
'  PUBLIC FUNCTIONS
'******************************************************************************
Public Function GoToItem(item As String, Optional RemoveFilterIfNecessary As Boolean = False, Optional suppressErrorMessage As Boolean = False)
    If Me.ItemNumber = item Then
        GoToItem = True
        Exit Function
    End If
    
    Dim pos As Long
    If Not ExistsInInventoryMaster(item) Then
        'If MsgBox("Item doesn't exist in Po/Inv, do you want to create it?", vbYesNo) = vbYes Then
        '    DB.Execute "INSERT INTO InventoryMaster ( ItemNumber ) VALUES ( '" & item & "' )"
        '    requeryForm "All"
        '    Me.barPosition.value = bsearch(item, currentFilterItemList)
        'End If
        If Not suppressErrorMessage Then
            MsgBox "Item '" & item & "' does not exist in Po/Inv!"
        End If
        GoToItem = False
    Else
        pos = bsearch(item, currentFilterItemList)
        If pos = -1 Then
            If RemoveFilterIfNecessary Then
                requeryForm "All"
                pos = bsearch(item, currentFilterItemList)
                If pos = -1 Then
                    'ok, maybe not. let's just give up here
                    GoToItem = False
                Else
                    'Me.barPosition.value = pos
                    loadItemByFilterIndex pos
                    GoToItem = True
                End If
            Else
                GoToItem = False
            End If
        Else
            'Me.barPosition.value = pos
            loadItemByFilterIndex pos
            GoToItem = True
        End If
    End If
End Function

Public Sub FilterByForm(filtersql As String)
    Me.btnFilterByForm.tag = filtersql
    requeryForm "FBF"
    'If Me.barPosition.value = 0 Then
    '    barPosition_Change
    'Else
    '    Me.barPosition.value = 0
    'End If
    loadItemByFilterIndex 0
    Me.jumpToLine.SetFocus
End Sub

'Public Sub FilterBySearch(filtersql As String)
'    Me.btnSearch.tag = filtersql
'    requeryForm "FBS"
'    'If Me.barPosition.value = 0 Then
'    '    barPosition_Change
'    'Else
'    '    Me.barPosition.value = 0
'    'End If
'    loadItemByFilterIndex 0
'    Me.jumpToLine.SetFocus
'End Sub

Public Sub SetFilter(filtername As String)
    'If InCombo(filtername, Me.cmbFilters) Then
        requeryForm filtername
        'If Me.barPosition.value = 0 Then
        '    barPosition_Change
        'Else
        '    Me.barPosition.value = 0
        'End If
        loadItemByFilterIndex 0
        
        Me.ClearCurrentVars

If InventoryQuantityTriggersV2.Visible = True Then
Unload InventoryQuantityTriggersV2
    If InventoryQuantityTriggersV2.Visible = True Then
        'do nothing, there was something to save...
        Exit Sub
    Else
        btnViewInvQtyTriggers_Click
    End If


End If
        'Me.jumpToLine.SetFocus
    'End If
End Sub

Public Sub SetFilterByItemList(newItemList() As String)
    currentFilterItemList = newItemList
    currentFilterPosition = 0
    requeryForm "<List>"
    loadItemByFilterIndex currentFilterPosition
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

Public Sub RequeryPoInvInPlace()
    Dim item As String
    'item = currentFilterItemList(Me.barPosition.value)
    item = currentFilterItemList(currentFilterPosition)
    requeryForm Me.cmbFilters
    'Me.barPosition.value = bsearch(item, currentFilterItemList, True)
    loadItemByFilterIndex bsearch(item, currentFilterItemList, True)
End Sub

Public Sub refreshForm(Optional updateOther As Boolean = True)
    Dim item As ADODB.Recordset
    Set item = DB.retrieve("exec spInventoryMasterByItem '" & Me.ItemNumber & "'")
    fillForm item
    item.Close
    Set item = Nothing
    If updateOther Then
        If IsFormLoaded("SignMaintenance") Then
            SignMaintenance.refreshForm False
        End If
    End If
End Sub

Public Sub LockForExport(tf As Boolean)
    If tf Then
        'lock everything, maybe work on this later
        Me.lblStatusBar.caption = "FORM LOCKED FOR EXPORT!!"
        Me.lblStatusBar.FontBold = True
        Me.lblStatusBar.ForeColor = 255
        'anything that we send to mas90 should go here
        lockFormCtls True
        Me.shipInfoFrame.Enabled = False
        lockedDownWD = True
        lockedDownPC = True
        lockedDown = True
    Else
        'unlock, but keep permissions
        Me.lblStatusBar.caption = ""
        Me.lblStatusBar.FontBold = False
        Me.lblStatusBar.ForeColor = 0
        If HasPermissionsTo("InventoryMaintenanceWrite") Then
            lockedDown = False
            lockFormCtls False
        End If
        If HasPermissionsTo("PriceComparison") Then
            lockedDownPC = False
        End If
        If HasPermissionsTo("WeightDims") Then
            lockedDownWD = False
            Me.shipInfoFrame.Enabled = True
            Me.components.Locked = False
        End If
    End If
End Sub

'Public Sub LockForImport(tf As Boolean)
'    Me.lblStatusBar = "LOCKED FOR IMPORT!!"
'End Sub

Public Sub ClearXSheetItemList()
    ReDim xsheetarray(10) As String
    setXSheetCheckStatus Me.ItemNumber
End Sub

Public Function GetXSheetItemList() As PerlArray
    ''this returns the raw xsheetarray, any calling functions
    ''need to check for "" strings in the middle and filter
    ''them out
    'GetXSheetItemList = xsheetarray
    Dim temp As PerlArray
    Set temp = New PerlArray
    Dim i As Long
    For i = 0 To UBound(xsheetarray)
        If xsheetarray(i) <> "" Then
            temp.Push CStr(xsheetarray(i))
        End If
    Next i
    Set GetXSheetItemList = temp
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
'Private Sub barPosition_Change()
'    Mouse.Hourglass True
'    handleChangedField
'    If currentFilterItemList(Me.barPosition.value) = "NORECORDS" Then
'        MsgBox "No records found."
'    Else
'        Dim pos As Long
'        Dim rst As ADODB.Recordset
'        Set rst = DB.retrieve("exec spInventoryMasterByItem '" & currentFilterItemList(Me.barPosition.value) & "'")
'        If rst.EOF Then
'            MsgBox "Record " & currentFilterItemList(Me.barPosition.value) & " deleted by another user? requerying..."
'            Dim item As String
'retry:
'            If Me.barPosition.value = Me.barPosition.Min Then
'                item = currentFilterItemList(Me.barPosition.value + 1)
'            Else
'                item = currentFilterItemList(Me.barPosition.value - 1)
'            End If
'            requeryForm Me.cmbFilters
'            pos = bsearch(item, currentFilterItemList)
'            If pos = -1 Then
'                GoTo retry
'            Else
'                Me.barPosition.value = pos
'            End If
'            Mouse.Hourglass False
'            Exit Sub
'        End If
'        If IsNull(rst("QtyOnHand")) Then
'            DB.Execute "INSERT INTO InventoryQuantities ( ItemNumber, WhseCode ) VALUES ( '" & currentFilterItemList(Me.barPosition.value) & "', '000' )"
'            rst.Close
'            Set rst = DB.retrieve("exec spInventoryMasterByItem'" & currentFilterItemList(Me.barPosition.value) & "'")
'        End If
'        fillForm rst
'        rst.Close
'        Set rst = Nothing
'        Me.txtPosition = Me.barPosition.value + 1
'        If Not firstLoad Then
'            If IsFormLoaded("SignMaintenance") Then
'                SignMaintenance.GoToItem currentFilterItemList(Me.barPosition.value)
'            End If
'        End If
'        If Not Me.ActiveControl Is Nothing Then
'            If TypeOf Me.ActiveControl Is TextBox Then
'                HandleTextBoxCursorPosition Me.ActiveControl
'            ElseIf TypeOf Me.ActiveControl Is ComboBox Then
'                HandleComboBoxCursorPosition Me.ActiveControl
'            End If
'        End If
'    End If
'    changed = False
'    Mouse.Hourglass False
'End Sub
Private Sub loadItemByFilterIndex(newdx As Long)


'see if our subtitle needs inserting, updating, removing...
Dim subtitle As String
subtitle = Me.SubtitleCmb.Text
Dim checkSub As ADODB.Recordset
Set checkSub = DB.retrieve("SELECT ID FROM Subtitles WHERE Subtitle='" & Replace(subtitle, "'", "''") & "'")
If checkSub.RecordCount > 0 Then
    If Me.SubTitleChk.value = vbChecked Then
        'update item and set offload flag
        Dim itemCheck As ADODB.Recordset
        Set itemCheck = DB.retrieve("SELECT ItemNumber FROM SubtitledItems WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
        If itemCheck.RecordCount > 0 Then
            'item exists in subbed itms db..
            DB.Execute "UPDATE SubtitledItems SET SubtitleID=" & checkSub("ID") & " WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
            
        Else
            'item doesnt exist in subb itms db
            DB.Execute "INSERT INTO SubtitledItems (ItemNumber,SubtitleID) VALUES('" & Me.ItemNumber.Text & "'," & CStr(checkSub("ID")) & ")"
            
        End If
    Else
        'not checked.. if the item exists in subtitleditems then remove it...
        Dim noSubCheck As ADODB.Recordset
        Set noSubCheck = DB.retrieve("SELECT ItemNumber FROM SubtitledItems WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
        If noSubCheck.RecordCount > 0 Then
        'it did exist... now we must remove...
        DB.Execute "DELETE FROM SubtitledItems WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
        End If
    End If
Else
    'create new sub then give it to current item
    If Me.SubTitleChk.value = vbChecked And Len(Me.SubtitleCmb.Text) > 0 Then
    DB.Execute "INSERT INTO Subtitles (Subtitle) VALUES('" & Replace(Me.SubtitleCmb.Text, "'", "''") & "')"
    Set checkSub = DB.retrieve("SELECT ID FROM Subtitles WHERE Subtitle='" & subtitle & "'")
    Dim subID As Integer
    subID = checkSub("ID")
    Set checkSub = DB.retrieve("SELECT ItemNumber FROM SubtitledItems WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
    If checkSub.RecordCount = 0 Then
        DB.Execute "INSERT INTO SubtitledItems (ItemNumber,SubtitleID) VALUES('" & Me.ItemNumber.Text & "'," & subID & ")"
    Else
        DB.Execute "UPDATE SubtitledItems SET SubtitleID=" & CStr(subID) & " WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
        
    End If
    
    End If
    
End If

    currentFilterPosition = newdx
    Mouse.Hourglass True
    If HasPermissionsTo("InventoryMaintenanceWrite") Then
        If Me.StdCost <> "" And Me.StdCost <> "$0.00" Then
            Dim priceAlerts As PerlArray
            Set priceAlerts = New PerlArray
            If Me.DiscountMarkupPriceRate(1) <> "" And Me.DiscountMarkupPriceRate(1) <> "$0.00" Then
                If priceFieldIsLessThanPriceField(Me.DiscountMarkupPriceRate(1), Me.StdCost) Then
                    InventoryQuantityTriggersV2.Visible = False
                     priceAlerts.Push "store price is less than standard cost!"
                End If
            End If
            If Me.IDiscountMarkupPriceRate(1) <> "" And Me.IDiscountMarkupPriceRate(1) <> "$0.00" Then
                If priceFieldIsLessThanPriceField(Me.IDiscountMarkupPriceRate(1), Me.StdCost) Then
                    InventoryQuantityTriggersV2.Visible = False
                    priceAlerts.Push "internet price is less than standard cost!"
                End If
            End If
            If Me.EBayPrice <> "" And Me.EBayPrice <> "$0.00" Then
                If priceFieldIsLessThanPriceField(Me.EBayPrice, Me.StdCost) Then
                    InventoryQuantityTriggersV2.Visible = False
                    priceAlerts.Push "ebay price is less than standard cost!"
                End If
            End If
            If priceAlerts.Scalar > 0 Then
                InventoryQuantityTriggersV2.Visible = False
                MsgBox "Warning for " & Me.ItemNumber & ":" & vbCrLf & vbCrLf & priceAlerts.Join(vbCrLf), vbExclamation
            End If
            Set priceAlerts = Nothing
        End If
    End If
    handleChangedField
    If UBound(currentFilterItemList) = -1 Then
        MsgBox "No records found."
    ElseIf currentFilterPosition > UBound(currentFilterItemList) Then
        MsgBox "Error, tell Brian how you did this!"
    Else
        Dim rst As ADODB.Recordset
        'Tom's Code 10-2-2014 - adding in a couple vars for subtitles...
        Dim itmEBayPrice As Double
        Dim itmItemNumber As String
        
        
        'CUT FROM HERE
         Set rst = DB.retrieve("exec spInventoryMasterByItem '" & currentFilterItemList(currentFilterPosition) & "'")
        If rst.EOF Then
            MsgBox "Unable to load item " & currentFilterItemList(currentFilterPosition) & ", maybe try reloading your filter?"
        Else
            If IsNull(rst("QtyOnHand")) Then
                DB.Execute "INSERT INTO InventoryQuantities ( ItemNumber, WhseCode ) VALUES ( '" & currentFilterItemList(currentFilterPosition) & "', '000' )"
                rst.Close
                Set rst = DB.retrieve("exec spInventoryMasterByItem'" & currentFilterItemList(currentFilterPosition) & "'")
            End If
            fillForm rst
            Me.txtPosition = currentFilterPosition + 1
            If Not firstLoad Then
                If IsFormLoaded("SignMaintenance") Then
                    SignMaintenance.GoToItem currentFilterItemList(currentFilterPosition)
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
        
        'Toms' Code 10-2-2014 to save rst data before it's wiped for checking subtitle...
        On Error GoTo damnit
        itmEBayPrice = rst("EBayPrice")
        itmItemNumber = rst("ItemNumber")
damnit:
        On Error GoTo 0
        rst.Close
        Set rst = Nothing
    End If
    
    
    'Tom's Code 10-2-2014 for subtitles
    If itmEBayPrice > 200 Then
        'check to make sure item doesn't have a subtitle assigned to it, otherwise assign the default...
        Dim sqlQuery As ADODB.Recordset
        Set sqlQuery = DB.retrieve("SELECT Subtitles.SubTitle FROM Subtitles, SubtitledItems WHERE SubtitledItems.SubtitleID=Subtitles.ID AND ItemNumber='" & itmItemNumber & "'")
        
        If sqlQuery.RecordCount > 0 Then
        Me.SubtitleCmb.Text = sqlQuery("SubTitle")
        Else
        'doesn't have a subtitle.. give it the default...
        Set sqlQuery = DB.retrieve("SELECT Subtitles.SubTitle FROM Subtitles WHERE ID=1")
        If sqlQuery.RecordCount > 0 Then
            Me.SubtitleCmb.Text = sqlQuery("SubTitle")
        Else
            MsgBox "Ok, well this is embarrassing.. I cant find the default subtitle. Go yell at Tom. He can fix it..."
            
        End If
        End If
    
    End If
    
    'check into dynamic pricing for this item...
Dim checkDynamic As ADODB.Recordset
Set checkDynamic = DB.retrieve("SELECT NewPrice,MinPrice,Enabled,EBay,TPStore,TPWeb,JetDotCom FROM PriceWiserItems WHERE ItemNumber='" & itmItemNumber & "'")
If checkDynamic.RecordCount > 0 Then
Me.dynamicPriceEditBtn.Visible = True
Me.dynamicPriceChk.caption = "Dynamic Price: $" & CStr(checkDynamic("NewPrice"))
'DynamicPricingFrm.applyBtn.Enabled = False
'DynamicPricingFrm.Enabled = False
'DynamicPricingFrm.dynamicPriceLbl.Caption = Me.dynamicPriceChk.Caption


'Me.lastPriceLbl.Visible = True
'Me.lastPriceLbl.Caption = "Last Price: $" + CStr(checkDynamic("NewPrice")) + "  Min. Price: $" & CStr(checkDynamic("MinPrice"))
If checkDynamic("Enabled") = True Then
Me.dynamicPriceChk.value = vbChecked
Else
Me.dynamicPriceChk.value = vbUnchecked
Me.dynamicPriceChk.caption = "Dynamic Price"
'Me.lastPriceLbl.Visible = False
Me.dynamicPriceEditBtn.Visible = False
End If
Else
'Me.lastPriceLbl.Caption = "Last Price: $"
Me.dynamicPriceEditBtn.Visible = False
'Me.lastPriceLbl.Visible = False
Me.dynamicPriceChk.value = vbUnchecked
Me.dynamicPriceChk.caption = "Dynamic Price"
End If
    
    'PASTED HERE
   
    
    changed = False
    Mouse.Hourglass False
End Sub

Private Sub handleChangedField()
    If changed = True Then
        Select Case whichCtl
            Case Is = "itemDescription"
                ItemDescription_LostFocus
            Case Is = "prodLine"
                prodLine_LostFocus
            Case Is = "PrimaryVendor"
                PrimaryVendor_LostFocus
            Case Is = "epn"
                EPN_LostFocus
            Case Is = "components"
                Components_LostFocus
            Case Is = "Weight"
                Weight_LostFocus
            Case Is = "Dimensions"
                Dimensions_LostFocus
            Case Is = "ReplacedBy"
                ReplacedBy_LostFocus
            Case Is = "ReplacementFor"
                ReplacementFor_LostFocus
            Case Is = "StdPack"
                StdPack_LostFocus
            Case Is = "NewCost"
                NewCost_LostFocus
            Case Is = "StdCost"
                StdCost_LostFocus
            Case Is = "StdPrice"
                StdPrice_LostFocus
            Case Is = "ListPrice"
                ListPrice_LostFocus
            Case Is = "MAPP"
                MAPP_LostFocus
            Case Is = "POComment"
                POComment_LostFocus
            Case Is = "QtyOrdered"
                QtyOrdered_LostFocus
            'Case Is = "CurrentOP"
            '    CurrentOP_LostFocus
            'Case Is = "NewOP"
            '    NewOP_LostFocus
'            Case Is = "Class"
'                Class_LostFocus
            Case Is = "OrderPointStore"
                OrderPointStore_LostFocus
            Case Is = "OrderPointWhse"
                OrderPointWhse_LostFocus
            'Case Is = "MfrCost"
            '    InventoriedDate_LostFocus
            Case Is = "DiscountMarkupPriceRate(1)"
                DiscountMarkupPriceRate_LostFocus 1
            Case Is = "DiscountMarkupPriceRate(2)"
                DiscountMarkupPriceRate_LostFocus 2
            Case Is = "DiscountMarkupPriceRate(3)"
                DiscountMarkupPriceRate_LostFocus 3
            Case Is = "DiscountMarkupPriceRate(4)"
                DiscountMarkupPriceRate_LostFocus 4
            Case Is = "DiscountMarkupPriceRate(5)"
                DiscountMarkupPriceRate_LostFocus 5
            Case Is = "IDiscountMarkupPriceRate(1)"
                IDiscountMarkupPriceRate_LostFocus 1
            Case Is = "IDiscountMarkupPriceRate(2)"
                IDiscountMarkupPriceRate_LostFocus 2
            Case Is = "IDiscountMarkupPriceRate(3)"
                IDiscountMarkupPriceRate_LostFocus 3
            Case Is = "IDiscountMarkupPriceRate(4)"
                IDiscountMarkupPriceRate_LostFocus 4
            Case Is = "IDiscountMarkupPriceRate(5)"
                IDiscountMarkupPriceRate_LostFocus 5
            Case Is = "BreakQty(1)"
                BreakQty_LostFocus 1
            Case Is = "BreakQty(2)"
                BreakQty_LostFocus 2
            Case Is = "BreakQty(3)"
                BreakQty_LostFocus 3
            Case Is = "BreakQty(4)"
                BreakQty_LostFocus 4
            Case Is = "BreakQty(5)"
                BreakQty_LostFocus 5
            Case Is = "IBreakQty(1)"
                IBreakQty_LostFocus 1
            Case Is = "IBreakQty(2)"
                IBreakQty_LostFocus 2
            Case Is = "IBreakQty(3)"
                IBreakQty_LostFocus 3
            Case Is = "IBreakQty(4)"
                IBreakQty_LostFocus 4
            Case Is = "IBreakQty(5)"
                IBreakQty_LostFocus 5
            Case Is = "EBayPrice"
                EBayPrice_LostFocus
            'Case Is = "EBayBestOfferAutomationLevel(0)"
            '    EBayBestOfferAutomationLevel_LostFocus 0
            'Case Is = "EBayBestOfferAutomationLevel(1)"
            '    EBayBestOfferAutomationLevel_LostFocus 1
            Case Is = "EBayBestOfferAutoAccept"
                EBayBestOfferAutoAccept_LostFocus
            Case Is = "TriadCode"
                TriadCode_LostFocus
            'Case Is = "EbayProstorePrice"
            '    EbayProstorePrice_LostFocus
            'Case Is = "EbayAuctionPrice"
            '    EbayAuctionPrice_LostFocus
            'Case Is = "EbayAuctionReservePrice"
            '    EbayAuctionReservePrice_LostFocus
            Case Else
                MsgBox "Couldn't figure out which control you just updated, so not saving..."
        End Select
    End If
End Sub

Private Sub AddSelectedSubtitleBtn_Click()
Dim testSub As ADODB.Recordset
Set testSub = DB.retrieve("SELECT ID FROM SubtitledItems WHERE Subtitle='" & Me.SubtitleCmb.Text & "'")
If testSub.RecordCount > 0 Then
    MsgBox "The subtitle with ID of " & CStr(testSub("ID")) & " has the same subtitle text. Please use that subtitle instead of creating a duplicate."
    
Else
    DB.Execute "INSERT INTO Subtitles (Subtitle) VALUES('" & Me.SubtitleCmb.Text & "')"
End If
End Sub

Private Sub addToJetBtn_Click()
JetDotCom.OffloadItemToJet Me.ItemNumber.Text
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

Private Sub AvgCost_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.HidePopUp = True
Me.HidePopUpFrm
Sleep 300
AvgCost.ZOrder 0
AvgCost.SetFocus
AvgCost.Refresh
End Sub

Private Sub blackListPriceComparisonsBtn_Click()
Shell "S:\mastest\mas90-signs\A_Dist\toms beta\CompetitorLister.exe"
End Sub

Private Sub btnAddFilterToXSheet_Click()
    Dim continue As Boolean
    continue = True
    If UBound(currentFilterItemList) > 200 Then
        continue = vbYes = MsgBox("There's a lot of items in the filter (" & (UBound(currentFilterItemList) + 1) & "), you might want to filter by form first? Continue anyway?", vbYesNo)
    End If
    If continue Then
        Dim i As Long
        For i = 0 To UBound(currentFilterItemList)
            addToXSheet currentFilterItemList(i)
        Next i
        setXSheetCheckStatus Me.ItemNumber
    End If
End Sub

Private Sub btnAddLCToXSheet_Click()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber FROM InventoryMaster WHERE ProductLine='" & Me.prodLine & "'")
    While Not rst.EOF
        addToXSheet rst("ItemNumber")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    setXSheetCheckStatus Me.ItemNumber
End Sub

Private Sub btnAddLineCode_Click()
    Dim lc As String
    lc = InputBox("Enter the three letter line code:")
    If lc <> "" Then
        lc = UCase(lc)
        If validateProductLine(lc) Then
            If Not PLExists(lc) Then
                If CBool(InStr(lc, "-")) Then
                    MsgBox "WARNING: as of Mas 4.20, there is a bug in VI so line codes with a dash won't import properly" & vbCrLf & vbCrLf & "You can make this line code, but let Brian know that he needs to import the line manually!"
                End If
                DB.Execute ("INSERT INTO ProductLine ( ProductLine, ManufFullName, ManufFullNameCleaned ) VALUES ( '" & lc & "', '" & lc & "', '" & lc & "' )")
                Load AddLineCode
                'AddLineCode.ProductLine = lc
                'AddLineCode.UpdateNow
                AddLineCode.LoadProductLine lc
                'AddLineCode.Show MODAL, InventoryMaintenance
                AddLineCode.Show
            Else
                If MsgBox("Line code already exists, would you like to edit it?", vbYesNo) = vbYes Then
                    Load AddLineCode
                    'AddLineCode.ProductLine = lc
                    'AddLineCode.UpdateNow
                    AddLineCode.LoadProductLine lc
                    'AddLineCode.Show MODAL, InventoryMaintenance
                    AddLineCode.Show
                End If
            End If
            requeryJumpToLine
        Else
            MsgBox "Line Code is invalid, must be 3 characters, only alphanumeric or dash."
        End If
    End If
End Sub

Private Sub btnAddNewItem_Click()
    Dim msg As String
    msg = "Enter line code and ItemNumber:" & vbCrLf & vbCrLf
    msg = msg & "Item will be added to the '" & WebsiteNameHash.item(CURRENT_WEBSITE_ID) & "' website"
    
    Dim newitem As String
    newitem = UCase(InputBox(msg, "Add New Item"))
    If newitem = "" Then
        Exit Sub
    Else
        'to make this easiest for margie, new items always get created for TP, fuck other websites
        'importing through spreadsheets or something can prompt or whatever
        If CreateItem(newitem, CURRENT_WEBSITE_ID) Then
            requeryForm "All"
            If Me.jumpToLine = Left(newitem, 3) Then
                requeryJumpToItem
            End If
            'Me.barPosition.value = bsearch(newitem, currentFilterItemList)
            loadItemByFilterIndex bsearch(newitem, currentFilterItemList)
            If IsFormLoaded("SignMaintenance") Then
                SignMaintenance.RequerySignsInPlace
                SignMaintenance.GoToItem newitem
            End If
            Me.ItemDescription.SetFocus
            Me.ItemDescription.selStart = 0
            Me.ItemDescription.SelLength = Len(Me.ItemDescription)
        End If
    End If
End Sub

'Private Sub btnAddShipClass_Click()
'    Load AddShipClass
'    AddShipClass.Show MODAL, InventoryMaintenance
'    requeryClass
'End Sub

Private Sub btnAlertsPending_Click()
    Load InventoryAlerts
    InventoryAlerts.Show MODAL
    Me.btnAlertsPending.Visible = False
    setAlertsButton
End Sub

Private Sub btnCopyEbayToStore_Click()
    LogItemStorePriceChanged Me.ItemNumber, DLookup("EBayPrice", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'")
    UpdateNotificationEmail "store price", Me.ItemNumber
    DB.Execute "EXEC spInvMaintCopyEbayToStore '" & Me.ItemNumber & "'"
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("exec spInventoryMasterByItem '" & Me.ItemNumber & "'")
    fillStorePrices rst
    rst.Close
    Set rst = Nothing
End Sub

Private Sub btnCopyEbayToWeb_Click()
    LogItemStorePriceChanged Me.ItemNumber, DLookup("EBayPrice", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'")
    UpdateNotificationEmail "web price", Me.ItemNumber
    DB.Execute "EXEC spInvMaintCopyEbayToWeb '" & Me.ItemNumber & "'"
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("exec spInventoryMasterByItem '" & Me.ItemNumber & "'")
    fillInetPrices rst
    rst.Close
    Set rst = Nothing
    updateDropshipItemWebPrice Me.ItemNumber, Me.IDiscountMarkupPriceRate(1)
End Sub

Private Sub btnCopyInstore_Click()
    LogItemWebPriceChanged Me.ItemNumber, DLookup("StdPrice", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'")
    UpdateNotificationEmail "web price", Me.ItemNumber
    DB.Execute "EXEC spInvMaintCopyStoreToWeb '" & Me.ItemNumber & "'"
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("exec spInventoryMasterByItem '" & Me.ItemNumber & "'")
    fillInetPrices rst
    rst.Close
    Set rst = Nothing
    updateDropshipItemWebPrice Me.ItemNumber, Me.IDiscountMarkupPriceRate(1)
End Sub

Private Sub btnCopyInternet_Click()
    LogItemStorePriceChanged Me.ItemNumber, DLookup("IDiscountMarkupPriceRate1", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'")
    UpdateNotificationEmail "store price", Me.ItemNumber
    DB.Execute "EXEC spInvMaintCopyWebToStore '" & Me.ItemNumber & "'"
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("exec spInventoryMasterByItem '" & Me.ItemNumber & "'")
    fillStorePrices rst
    Me.StdPrice = Format(rst("StdPrice"), "Currency")
    rst.Close
    Set rst = Nothing
End Sub

Private Sub btnCopyStoreToEbay_Click()
    LogItemEBayPriceChanged Me.ItemNumber, DLookup("EBayPrice", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'")
    UpdateNotificationEmail "ebay price", Me.ItemNumber
    DB.Execute "UPDATE InventoryMaster Set EBayPrice = StdPrice WHERE ItemNumber='" & Me.ItemNumber & "'"
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("exec spInventoryMasterByItem '" & Me.ItemNumber & "'")
    fillEbayPrices rst
    rst.Close
    Set rst = Nothing
End Sub

Private Sub btnDeleteItem_Click()
    If Mas200Helpers.ItemExistsInMas(Me.ItemNumber) Then
        MsgBox "This item still exists in Mas! Delete there first, then come back here to delete."
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete the item:" & vbCrLf & vbCrLf & Me.ItemNumber, vbYesNo) = vbYes Then
        'deleting an item will potentially delete components that are referenced
        'in the warehouse. that's bad. check for inventory on hand, only delete
        'if zero.
        Dim canDelete As Boolean
        canDelete = False
        If DLookup("IsMASKit", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'") Then
            'kits can be deleted with no problems
            canDelete = True
        ElseIf DLookup("COUNT(*)", "InventoryQuantities", "ItemNumber='" & Me.ItemNumber & "' AND QuantityOnHand<>0") = 0 Then
            'item is not in mas inventory, can be deleted without screwing up the
            'warehouse. we can't delete an item with Mas history in Mas, so that's
            'a problem, but that's harder to check? not sure what the definitive
            'thing is for "has history" for mas.
            canDelete = True
        End If
        
        If canDelete = False Then
            MsgBox "Can't delete an item that has inventory in Mas!"
        Else
            Dim rst As ADODB.Recordset
            Set rst = DB.retrieve("SELECT ComponentID FROM InventoryComponentMap WHERE ItemID=" & Me.LoadedItemID)
            While Not rst.EOF
                If "1" = DLookup("COUNT(ItemID)", "InventoryComponentMap", "ComponentID=" & rst("ComponentID")) Then
                    DB.Execute "DELETE FROM InventoryComponents WHERE ID=" & rst("ComponentID")
                    DB.Execute "DELETE FROM InventoryComponentBarcodes WHERE ComponentID=" & rst("ComponentID")
                End If
                rst.MoveNext
            Wend
            rst.Close
            Set rst = Nothing
            DB.Execute "EXEC __spItemDelete '" & Me.ItemNumber & "'"
            MsgBox "Delete successful. Remember to delete it from MAS 200. Also, check to see if it's on any POs, otherwise you'll get batch errors."
            requeryForm Me.cmbFilters
            'barPosition_Change
            If currentFilterPosition > UBound(currentFilterItemList) Then
                currentFilterPosition = UBound(currentFilterItemList)
            End If
            loadItemByFilterIndex currentFilterPosition
        End If
    End If
End Sub

Private Sub btnEBayKitPriceManager_Click()
    Load EBayKitPriceManager
    EBayKitPriceManager.Show
End Sub

Private Sub btnEditLineCode_Click()
    Load AddLineCode
    'AddLineCode.ProductLine = Me.jumpToLine
    'AddLineCode.UpdateNow
    AddLineCode.LoadProductLine Me.jumpToLine
    'AddLineCode.Show MODAL, InventoryMaintenance
    AddLineCode.Show
    'Dim rst As ADODB.Recordset
    'Set rst = DB.retrieve("SELECT MinimumOrderAmount, PrepaidAmount, PrepaidSpecialAmount FROM ProductLine WHERE ProductLine='" & Me.prodLine & "'")
    'Me.MinimumOrderAmount = IIf(IsNumeric(rst("MinimumOrderAmount")), Format(rst("MinimumOrderAmount"), "Currency"), Nz(rst("MinimumOrderAmount")))
    'Me.PrepaidAmount = IIf(IsNumeric(rst("PrepaidAmount")), Format(rst("PrepaidAmount"), "Currency"), Nz(rst("PrepaidAmount")))
    'Me.PrepaidSpecialAmount = IIf(IsNumeric(rst("PrepaidSpecialAmount")), Format(rst("PrepaidSpecialAmount"), "Currency"), Nz(rst("PrepaidSpecialAmount")))
    'rst.Close
    'Set rst = Nothing
End Sub

Private Sub btnEmailItem_Click()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Desc1, Desc2, Desc3 FROM PartNumbers WHERE ItemNumber='" & Me.ItemNumber & "'")
    If rst.EOF Then
        OpenEmailTo "", Me.lblManufName & " " & Left(Me.ItemNumber, 4)
    Else
        OpenEmailTo "", Me.lblManufName & " " & Nz(rst("Desc2"), Left(Me.ItemNumber, 4)) & " " & Nz(rst("Desc1"), "") & " " & Nz(rst("Desc3"), "")
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnFilterByForm_Click()
    Load FilterByFormDialog
    FilterByFormDialog.hiddenOwner = "InventoryMaintenance"
    FilterByFormDialog.Show MODAL
    If Me.btnFilterByForm.tag <> "" Then
        If IsFormLoaded("SignMaintenance") Then
            If MsgBox("Filter Signs too?", vbYesNo) = vbYes Then
                SignMaintenance.FilterByForm InventoryMaintenance.btnFilterByForm.tag
            End If
        End If
        requeryForm "FBF"
        'If Me.barPosition.value = 0 Then
        '    barPosition_Change
        'Else
        '    Me.barPosition.value = 0
        'End If
        loadItemByFilterIndex 0
    End If
End Sub

Private Sub btnFirstItem_Click()
    'Me.barPosition.value = Me.barPosition.Min
    If currentFilterPosition <> 0 Then
        loadItemByFilterIndex 0
    End If
    If InventoryQuantityTriggersV2.Visible = True Then
Unload InventoryQuantityTriggersV2


    If InventoryQuantityTriggersV2.Visible = True Then
        'do nothing, there was something to save...
        Exit Sub
    Else
        btnViewInvQtyTriggers_Click
    End If


End If

    
End Sub

Private Sub btnFreightEstimate_Click()
    Mouse.Hourglass True
    Dim ZipCode As String
    ZipCode = InputBox("Enter zip code:" & vbCrLf & vbCrLf & "Testing zip codes:" & vbCrLf & "  - 89005 - Nevada" & vbCrLf & "  - 99521 - Alaska")
    If ZipCode <> "" Then
        Mouse.Hourglass True
        Dim itemHash As Dictionary
        Set itemHash = New Dictionary
        itemHash.Add CStr(Me.ItemNumber.Text), CLng(1)
        Dim Results As Dictionary
        Set Results = ShippingDB.EstimateFreightForShipment(itemHash, ZipCode)
        If Not Results Is Nothing Then
            Dim iter As Variant
            For Each iter In Results.keys
                If iter = "UPS Ground" Then
                    DB.Execute "DELETE FROM FreightEstimates WHERE ItemID=" & Me.LoadedItemID & " AND ZipCode='" & EscapeSQuotes(ZipCode) & "' AND Service='U'"
                    DB.Execute "INSERT INTO FreightEstimates (ItemID, ItemNumber, ZipCode, Cost, Service) VALUES ( " & Me.LoadedItemID & ", '" & Me.ItemNumber & "', '" & EscapeSQuotes(ZipCode) & "', " & CStr(Results.item(CStr(iter))) & ", 'U' )"
                ElseIf iter = "USPS Priority Mail" Then
                    DB.Execute "DELETE FROM FreightEstimates WHERE ItemID=" & Me.LoadedItemID & " AND ZipCode='" & EscapeSQuotes(ZipCode) & "' AND Service='P'"
                    DB.Execute "INSERT INTO FreightEstimates (ItemID, ItemNumber, ZipCode, Cost, Service) VALUES ( " & Me.LoadedItemID & ", '" & Me.ItemNumber & "', '" & EscapeSQuotes(ZipCode) & "', " & CStr(Results.item(CStr(iter))) & ", 'P' )"
                ElseIf iter = "USPS First-Class Mail" Then
                    DB.Execute "DELETE FROM FreightEstimates WHERE ItemID=" & Me.LoadedItemID & " AND ZipCode='" & EscapeSQuotes(ZipCode) & "' AND Service='FC'"
                    DB.Execute "INSERT INTO FreightEstimates (ItemID, ItemNumber, ZipCode, Cost, Service) VALUES ( " & Me.LoadedItemID & ", '" & Me.ItemNumber & "', '" & EscapeSQuotes(ZipCode) & "', " & CStr(Results.item(CStr(iter))) & ", 'FC' )"
                End If
            Next iter
        End If
        Mouse.Hourglass False
    End If
    requeryFreightEstimates
    FlashWindow Me.hwnd, 3
    Mouse.Hourglass False
End Sub

Private Sub btnGenerateXSheet_Click()
    Load InventorySpreadsheetDialog
    'InventorySpreadsheetDialog.LoadItems xsheetarray
    InventorySpreadsheetDialog.Show MODAL
End Sub

Private Sub btnGoToDropshipItem_Click()
    Dim item As String
    If Left(Me.ItemNumber, 3) <> "ZZZ" Then
        item = ZZZifyItemNumber(Me.ItemNumber)
        If InventoryMaintenance.GoToItem(item, True, True) Then
            'done
        Else
            'zzz not found
            If vbYes = MsgBox("Dropship item '" & item & "' not found, create it now?", vbYesNo) Then
                'If Me.chkDropship.Enabled = False Then
                '    MsgBox "Can't make this dropshippable! Maybe it's discontinued?"
                'ElseIf Me.chkDropship = 1 Then
                    CreateZZZClone Me.ItemNumber
                    InventoryMaintenance.GoToItem item, True, False
                'Else
                '    Me.chkDropship = 1 'this should create the item on the fly
                '    InventoryMaintenance.GoToItem item, True, False
                'End If
            End If
        End If
    Else
        'any way to get this working better? zzz-ify is pretty destructive when necessary
        item = Mid(Me.ItemNumber, 4)
        InventoryMaintenance.GoToItem item, True, False
    End If
End Sub

Private Sub btnGoToEBay_Click()
SignMaintenance.ItemNumber.Text = Me.ItemNumber.Text
SignMaintenance.btnGoToEBay_Click
If SignMaintenance.Visible = False Then
    Unload SignMaintenance
End If
End Sub

Private Sub btnGoToSigns_Click()
    If IsFormLoaded("SignMaintenance") Then
        If IsFormMinimized("SignMaintenance") Then
            UnMinimizeForm "SignMaintenance"
        Else
            FocusOnForm "SignMaintenance"
        End If
        'SignMaintenance.GoToItem Me.ItemNumber
    Else
        Load SignMaintenance
        SignMaintenance.GoToItem Me.ItemNumber
        SignMaintenance.Show
    End If
End Sub

Private Sub btnGoToTPP_Click()
SignMaintenance.ItemNumber.Text = Me.ItemNumber.Text
SignMaintenance.btnGoToTPP_Click
If SignMaintenance.Visible = False Then
    Unload SignMaintenance
End If

End Sub

Private Sub btnGoToWeb_Click()
SignMaintenance.ItemNumber.Text = Me.ItemNumber.Text
SignMaintenance.btnGoToWebsite_Click
If SignMaintenance.Visible = False Then
    Unload SignMaintenance
End If
End Sub

Private Sub btnGoToWebsite_Click()
OpenDefaultApp "http://" & WebsiteURLHash.item(CLng(Me.Website.tag)) & "/" & LCase(CreateYahooID(Me.ItemNumber)) & ".html"

End Sub

Private Sub btnItemHistory_Click()
    Load InventoryItemHistory
    InventoryItemHistory.Show MODAL
End Sub

Private Sub btnLastItem_Click()
    'Me.barPosition.value = Me.barPosition.max
    If currentFilterPosition <> UBound(currentFilterItemList) Then
        loadItemByFilterIndex UBound(currentFilterItemList)
    End If
    
    If InventoryQuantityTriggersV2.Visible = True Then
Unload InventoryQuantityTriggersV2


    If InventoryQuantityTriggersV2.Visible = True Then
        'do nothing, there was something to save...
        Exit Sub
    Else
        btnViewInvQtyTriggers_Click
    End If


End If

    
End Sub

Private Sub btnLoadBarcodes_Click()
    If HasPermissionsTo("Barcode") Then
        If IsFormLoaded("Barcode") Then
            Barcode.LoadItem Me.ItemNumber
            FocusOnForm "Barcode"
        Else
            Load Barcode
            Barcode.LoadItem Me.ItemNumber
            Barcode.Show
        End If
    End If
End Sub

Private Sub btnMAPPtoWP_Click()
    If MsgBox("Set web price for this item to the MAPP?", vbYesNo) = vbYes Then
        LogItemWebPriceChanged Me.ItemNumber, Me.MAPP
        DB.Execute "UPDATE InventoryMaster SET IDiscountMarkupPriceRate1=MAPP WHERE ItemNumber='" & Me.ItemNumber & "'"
        Me.IDiscountMarkupPriceRate(1) = Me.MAPP
    End If
End Sub

'Private Sub btnMAPPtoWPforLine_Click()
'    If MsgBox("Set web price to the MAPP price for ALL items in line?", vbYesNo) = vbYes Then
'        Dim uponly As Boolean
'        uponly = CBool(MsgBox("Only change price if MAPP is higher?", vbYesNo) = vbYes)
'        Dim rst As ADODB.Recordset
'        Set rst = DB.retrieve("SELECT ItemNumber, MAPP FROM InventoryMaster WHERE ProductLine='" & Me.prodLine & "' AND MAPP<>0" & IIf(uponly, " AND MAPP>IDiscountMarkupPriceRate1", ""))
'        While Not rst.EOF
'            Me.lblStatusBar.caption = "Processing " & rst("ItemNumber")
'            DoEvents
'            LogItemWebPriceChanged rst("ItemNumber"), rst("MAPP")
'            updateInventoryMaster "IDiscountMarkupPriceRate1", rst("MAPP"), rst("ItemNumber"), ""
'            rst.MoveNext
'        Wend
'        Me.lblStatusBar.caption = ""
'        rst.Close
'        Set rst = Nothing
'    End If
'    RefreshForm
'End Sub

Private Sub btnNewBarcode_Click()
    If AddNewBarcodeForItem(Me.ItemNumber, Me.LoadedItemID) <> "" Then
        Me.btnLoadBarcodes.caption = "Barcode (" & DLookup("COUNT(Barcode)", "vBarcodesComponents", "ItemNumber='" & Me.ItemNumber & "'") & ")"
    End If
End Sub

Public Sub ClearCurrentVars()

    Me.CurrentAVGdata = ""
    Me.CurrentDSCostdata = ""
    Me.CurrentINVdata = ""
    Me.CurrentLISTdata = ""
    Me.CurrentMAPPdata = ""
    Me.CurrentSTDdata = ""
    Me.CurrentSTOREdata = ""
    Me.CurrentTNCdata = ""
    Me.CurrentSTDPricedata = ""
    Me.CurrentIDiscountRate = ""
    Me.HidePopUp = False
   
    

End Sub


Public Sub btnNextItem_Click()
'    If currentFilterPosition <> UBound(currentFilterItemList) Then
'        loadItemByFilterIndex currentFilterPosition + 1
'    End If




'If AlertSaveChanges.Visible = False Then
    If currentFilterPosition <> UBound(currentFilterItemList) Then
        If InventoryQuantityTriggersV2.IsSaving = False Then
               
            loadItemByFilterIndex currentFilterPosition + 1
        End If
        
    End If
'End If


Me.ClearCurrentVars

If InventoryQuantityTriggersV2.Visible = True Then
Unload InventoryQuantityTriggersV2
    If InventoryQuantityTriggersV2.Visible = True Then
        'do nothing, there was something to save...
        Exit Sub
    Else
        btnViewInvQtyTriggers_Click
    End If


End If




End Sub


Private Sub btnNextLC_Click()
    Mouse.Hourglass True
    Dim oldLC As String
    Dim pos As Long
    Dim found As Boolean
    oldLC = Trim$(Me.prodLine)
    'pos = Me.barPosition.value
    pos = currentFilterPosition
    found = False
    While Not found
        'If pos < Me.barPosition.max Then
        If pos < UBound(currentFilterItemList) - 1 Then
            If oldLC = Left$(currentFilterItemList(pos), 3) Then
                pos = pos + 1
            Else
                found = True
            End If
        Else
            found = True
        End If
    Wend
    'Me.barPosition.value = pos
    If currentFilterPosition <> pos Then
        loadItemByFilterIndex pos
    End If
    Mouse.Hourglass False
    
    If InventoryQuantityTriggersV2.Visible = True Then
Unload InventoryQuantityTriggersV2


    If InventoryQuantityTriggersV2.Visible = True Then
        'do nothing, there was something to save...
        Exit Sub
    Else
        btnViewInvQtyTriggers_Click
    End If


End If

End Sub

Public Sub btnPrevItem_Click()
'    If currentFilterPosition <> 0 Then
'        loadItemByFilterIndex currentFilterPosition - 1
'    End If



'If AlertSaveChanges.Visible = False Then
    If currentFilterPosition <> 0 Then
        If InventoryQuantityTriggersV2.IsSaving = False Then
        loadItemByFilterIndex currentFilterPosition - 1
        End If
    End If
'End If


Me.ClearCurrentVars


If InventoryQuantityTriggersV2.Visible = True Then
    Unload InventoryQuantityTriggersV2
        If InventoryQuantityTriggersV2.Visible = True Then
            Exit Sub
        Else
            btnViewInvQtyTriggers_Click
        End If
End If

End Sub

'Private Sub btnOpenSoundRec_Click()
'    OpenSoundRecorder
'End Sub

Private Sub btnPrevLC_Click()
    Mouse.Hourglass True
    Dim oldLC As String
    Dim pos As Long
    Dim found As Boolean
    oldLC = Trim$(Me.prodLine)
    'pos = Me.barPosition.value
    pos = currentFilterPosition
    found = False
    While Not found
        'If pos > Me.barPosition.Min Then
        If pos > 0 Then
            If oldLC = Left$(currentFilterItemList(pos), 3) Then
                pos = pos - 1
            Else
                found = True
            End If
        Else
            found = True
        End If
    Wend
    Dim newLC As String
    newLC = Left$(currentFilterItemList(pos), 3)
    found = False
    While Not found
        'If pos > Me.barPosition.Min Then
        If pos > 0 Then
            If newLC = Left$(currentFilterItemList(pos), 3) Then
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
    If currentFilterPosition <> pos Then
        loadItemByFilterIndex pos
    End If
    Mouse.Hourglass False
    
    
    If InventoryQuantityTriggersV2.Visible = True Then
Unload InventoryQuantityTriggersV2


    If InventoryQuantityTriggersV2.Visible = True Then
        'do nothing, there was something to save...
        Exit Sub
    Else
        btnViewInvQtyTriggers_Click
    End If


End If

End Sub

Private Sub btnPriceCompare_Click()
    PriceCompareSingleItemOrLC Me.ItemNumber, True, False
    requeryPriceComparisons
    FlashWindow Me.hwnd, 3
End Sub

Private Sub btnPriceCompareDetails_Click()
    Load PriceComparisonDlg
    PriceComparisonDlg.Show MODAL
End Sub

Private Sub btnPriceCompareGoToEBay_Click()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Query, StdCost, TNC FROM vPriceComparisonInfo WHERE ItemNumber='" & Me.ItemNumber & "'")
    If rst.EOF Then
        MsgBox "Error generating price compare query for this item!"
    ElseIf Nz(rst(0)) = "" Then
        MsgBox "Error generating search query! This line might be missing the manufacturer web name!"
    Else
        Dim minimum As Variant
        If CDec(rst("StdCost")) = 0 And CDec(rst("TNC")) = 0 Then
            minimum = 0
        ElseIf CDec(rst("StdCost")) > 0 And CDec(rst("TNC")) = 0 Then
            minimum = CDec(rst("StdCost"))
        ElseIf CDec(rst("StdCost")) = 0 And CDec(rst("TNC")) > 0 Then
            minimum = CDec(rst("TNC"))
        Else
            minimum = IIf(CDec(rst("StdCost")) < CDec(rst("TNC")), CDec(rst("StdCost")), CDec(rst("TNC")))
        End If
        minimum = CDec(Round(CDec(minimum) * 0.8, 2))
        'LH_BIN=1              == buy it now only
        '_sop=15               == sort price + shipping: lowest first
        'LH_ItemCondition=3    == item condition new
        'LH_SellerWithStore=1  == store only results, also puts the store name in the results
        '_udlo=minimum         == price minimum threshold 80% of std/tnc, whichever is lower
        '_nkw                  == search term
        OpenDefaultApp "http://www.ebay.com/sch/i.html?LH_BIN=1&_sop=15&LH_ItemCondition=3&LH_SellerWithStore=1" & IIf(CDec(minimum) = 0, "", "&_udlo=" & minimum) & "&_nkw=" & URLEncode(rst(0))
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub btnPriceCompareGoToFroogle_Click()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Query, StdCost, TNC FROM vPriceComparisonInfo WHERE ItemNumber='" & Me.ItemNumber & "'")
    If rst.EOF Then
        MsgBox "Error generating price compare query for this item!"
    ElseIf Nz(rst(0)) = "" Then
        MsgBox "Error generating search query! This line might be missing the manufacturer web name!"
    Else
        OpenDefaultApp "http://www.google.com/products?oe=utf-8&um=1&ie=UTF-8&sa=N&hl=en&tab=wf&q=" & URLEncode(rst(0))
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub btnPriceCompareGoToYahoo_Click()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Query FROM vPriceComparisonInfo WHERE ItemNumber='" & Me.ItemNumber & "'")
    If rst.EOF Then
        MsgBox "Error generating price compare query for this item!"
    ElseIf Nz(rst(0)) = "" Then
        MsgBox "Error generating search query! This line might be missing the manufacturer web name!"
    Else
        OpenDefaultApp "http://shopping.yahoo.com/search?p=" & URLEncode(rst(0))
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub btnPrintScreen_Click()
    ClipBoard_SetImage Utilities.CaptureVB6Form(Me)
    MsgBox "Screenshot captured!"
End Sub

Private Sub btnRefreshQuantitiesForItem_Click()
    Mouse.Hourglass True
    Dim m As Mas200ImportExport
    Set m = New Mas200ImportExport
    m.RefreshQuantities Me.ItemNumber
    Set m = Nothing
    CheckForBeingDCItems
    Mouse.Hourglass False
    
    fillSalesHistory Me.ItemNumber, Me.chkDropship
End Sub

Private Sub btnRemoveFilter_Click()
    Dim current As String
    current = Me.ItemNumber
    Me.cmbFilters = "All"
    requeryForm "All"
    setAlertsButton
    'Me.barPosition.value = bsearch(current, currentFilterItemList)
    ''Me.barPosition.value = 0
    'barPosition_Change
    loadItemByFilterIndex bsearch(current, currentFilterItemList)
End Sub

Private Sub btnReplacedByGoTo_Click()
    If Me.ReplacedBy <> "" Then
        Dim pos As Long
retry:
        pos = bsearch(Me.ReplacedBy, currentFilterItemList)
        If pos = -1 Then
            If MsgBox("Item is not in current filter?" & vbCrLf & vbCrLf & "Remove filter and try again?", vbYesNo) = vbYes Then
                Dim lcitem As String
                lcitem = Me.ReplacedBy
                requeryForm "All"
                Me.jumpToLine = Left$(lcitem, 3)
                Me.jumpToItem = Mid$(lcitem, 4)
                GoTo retry
            End If
        Else
            'Me.barPosition.value = pos
            loadItemByFilterIndex pos
        End If
    End If
End Sub

Private Sub btnReplaceForGoTo_Click()
    If Me.ReplacementFor <> "" Then
        Dim pos As Long
retry:
        pos = bsearch(Me.ReplacementFor, currentFilterItemList)
        If pos = -1 Then
            If MsgBox("Item is not in current filter?" & vbCrLf & vbCrLf & "Remove filter and try again?", vbYesNo) = vbYes Then
                Dim lcitem As String
                lcitem = Me.ReplacementFor
                requeryForm "All"
                Me.jumpToLine = Left$(lcitem, 3)
                Me.jumpToItem = Mid$(lcitem, 4)
                GoTo retry
            End If
        Else
            'Me.barPosition.value = pos
            loadItemByFilterIndex pos
        End If
    End If
End Sub

Private Sub btnReports_Click()
    OpenReportsDB
End Sub

Private Sub btnSalesHistoryColumnToggle_Click()
    If Me.btnSalesHistoryColumnToggle.caption = "Combine" Then
        Me.btnSalesHistoryColumnToggle.caption = "Split"
        Me.salesHistoryHeaderLabel(1).caption = "Sold"
    Else
        Me.btnSalesHistoryColumnToggle.caption = "Combine"
        Me.salesHistoryHeaderLabel(1).caption = "SO"
    End If
    
    Me.salesHistoryHeaderLabel(0).Visible = CBool(Me.btnSalesHistoryColumnToggle.caption = "Combine")
    Dim i As Long
    For i = Me.SalesP2.LBound To Me.SalesP2.UBound
        Me.SalesP2(i).Visible = CBool(Me.btnSalesHistoryColumnToggle.caption = "Combine")
    Next i
    
    fillSalesHistory Me.ItemNumber, CBool(Me.chkDropship)
End Sub

Private Sub btnSalesHistoryPeriodTypeToggle_Click()
    If Me.btnSalesHistoryPeriodTypeToggle.caption = "Rolling" Then
        Me.btnSalesHistoryPeriodTypeToggle.caption = "Static"
        Me.salesHistoryRowLabel(0).caption = "Last 30"
        Me.salesHistoryRowLabel(1).caption = "30 to 60"
        Me.salesHistoryRowLabel(2).caption = "Last 90"
        Me.salesHistoryRowLabel(3).caption = "90 to 180"
        Me.salesHistoryRowLabel(4).caption = "Last 365"
        Me.salesHistoryRowLabel(5).caption = "365 to 730"
    Else
        Me.btnSalesHistoryPeriodTypeToggle.caption = "Rolling"
        Me.salesHistoryRowLabel(0).caption = "This Month"
        Me.salesHistoryRowLabel(1).caption = "Last Month"
        Me.salesHistoryRowLabel(2).caption = "Current Per"
        Me.salesHistoryRowLabel(3).caption = "Last Per"
        Me.salesHistoryRowLabel(4).caption = "Sales YTD"
        Me.salesHistoryRowLabel(5).caption = "Sales LYR"
    End If
    fillSalesHistory Me.ItemNumber, CBool(Me.chkDropship)
End Sub

'Private Sub btnSearch_Click()
'    Load SearchByPath
'    SearchByPath.Show
'End Sub

Private Sub btnOptionsDialog_Click()
    Load OptionsDialog
    OptionsDialog.Show MODAL
End Sub

Private Sub btnSuggestQtyToOrder_Click()
    If Me.prodLine <> "ZZZ" Then
        'suggest button functionality
        Dim op As Long, diff As Long
        op = DLookup("SUM(OrderPoint)", "InventoryLocationInfo", "ItemNumber='" & Me.ItemNumber & "'")
        If Me.QtyOnHand < op Then
            diff = op - Me.QtyOnHand
            If diff < Me.StdPack Then
                Me.QtyOrdered = Me.StdPack
            ElseIf diff Mod Me.StdPack = 0 Then
                Me.QtyOrdered = diff
            Else
                Me.QtyOrdered = Me.StdPack * (Int(diff / Me.StdPack) + 1)
            End If
            changed = True
            QtyOrdered_LostFocus
        Else
            Me.QtyOrdered = 0
        End If
    Else
        'po list functionality
        Load PurchaseOrderSelectionDialog
        PurchaseOrderSelectionDialog.Show MODAL
        'update qtyordered, dupe of fillForm()
        Dim poRst As ADODB.Recordset
        Set poRst = DB.retrieve("SELECT SUM(Quantity), COUNT(*) FROM PurchaseOrders INNER JOIN PurchaseOrderLines ON PurchaseOrders.ID=PurchaseOrderLines.HeaderID WHERE Exported=0 AND ItemNumber='" & Me.ItemNumber & "'")
        If poRst(1) = "0" Then
            Me.QtyOrdered = "0"
        ElseIf poRst(1) = "1" Then
            Me.QtyOrdered = poRst(1)
        Else
            Me.QtyOrdered = poRst(0) & "/" & poRst(1)
        End If
        requeryPOList lookupPrimaryVendorNumber(Me.PrimaryVendor)
    End If
End Sub

Private Sub btnTNCforItem_Click()
    If DLookup("TNC", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'") <> 0 Then
        If MsgBox("Apply TNC for this item?", vbYesNo) = vbYes Then
            LogItemCostChanged Me.ItemNumber, Me.NewCost
            DB.Execute "UPDATE InventoryMaster SET StdCost=TNC, TNC=0 WHERE ItemNumber='" & Me.ItemNumber & "'"
DatabaseFunctions.ModifyItemCost Me.LoadedItemID, "Std Cost", IIf(Me.NewCost = "", "0", Me.NewCost)
DatabaseFunctions.ModifyItemCost Me.LoadedItemID, "New Cost", "0"
If Left(Me.ItemNumber, 3) <> "XXX" And Left(Me.ItemNumber, 3) <> "ZZZ" Then
Dim zzz As String
zzz = ZZZifyItemNumber(Me.ItemNumber)
If zzz <> "" Then
Dim zzzID As Long
zzzID = Utilities.GetItemIDByItemCode(zzz)
If zzzID <> -1 Then
updateInventoryMaster "StdCost", IIf(Me.NewCost = "", "0", Me.NewCost), zzz, ""
updateInventoryMaster "TNC", "0", zzz, ""
DatabaseFunctions.ModifyItemCost zzzID, "Std Cost", IIf(Me.NewCost = "", "0", Me.NewCost)
DatabaseFunctions.ModifyItemCost zzzID, "New Cost", "0"
End If
End If
End If
        End If
        refreshForm
    End If
End Sub

Private Sub btnTNCforLine_Click()
    If Left(Me.ItemNumber, 3) = "XXX" Or Left(Me.ItemNumber, 3) = "ZZZ" Then
        MsgBox "Can't do XXX/ZZZ!"
        Exit Sub
    End If
    If MsgBox("Apply TNC for all items in this line?", vbYesNo) = vbYes Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT ItemID, ItemNumber, TNC FROM InventoryMaster WHERE TNC<>0 AND (LEFT(ItemNumber,3)='" & Left(Me.ItemNumber, 3) & "')")
        While Not rst.EOF
            LogItemCostChanged rst("ItemNumber"), rst("TNC")
            DB.Execute "UPDATE InventoryMaster SET StdCost=TNC, TNC=0 WHERE TNC<>0 AND ItemNumber='" & rst("ItemNumber") & "'"
DatabaseFunctions.ModifyItemCost rst("ItemID"), "Std Cost", rst("TNC")
DatabaseFunctions.ModifyItemCost rst("ItemID"), "New Cost", "0"
Dim zzz As String
zzz = ZZZifyItemNumber(Me.ItemNumber)
If zzz <> "" Then
Dim zzzID As Long
zzzID = Utilities.GetItemIDByItemCode(zzz)
If zzzID <> -1 Then
updateInventoryMaster "StdCost", rst("TNC"), zzz, ""
updateInventoryMaster "TNC", "0", zzz, ""
DatabaseFunctions.ModifyItemCost zzzID, "Std Cost", rst("TNC")
DatabaseFunctions.ModifyItemCost zzzID, "New Cost", "0"
End If
End If
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
    End If
    refreshForm
End Sub

Private Sub btnTNCSwap_Click()
    If MsgBox("Swap TNC and StdCost for this item?", vbYesNo) = vbYes Then
        LogItemCostSwapped Me.ItemNumber
        DB.Execute "UPDATE InventoryMaster SET TNC=StdCost, StdCost=TNC WHERE ItemNumber='" & Me.ItemNumber & "'"
DatabaseFunctions.ModifyItemCost Me.LoadedItemID, "Std Cost", IIf(Me.NewCost = "", "0", Me.NewCost)
DatabaseFunctions.ModifyItemCost Me.LoadedItemID, "New Cost", Me.StdCost
If Left(Me.ItemNumber, 3) <> "XXX" And Left(Me.ItemNumber, 3) <> "ZZZ" Then
Dim zzz As String
zzz = ZZZifyItemNumber(Me.ItemNumber)
If zzz <> "" Then
Dim zzzID As Long
zzzID = Utilities.GetItemIDByItemCode(zzz)
If zzzID <> -1 Then
updateInventoryMaster "StdCost", IIf(Me.NewCost = "", "0", Me.NewCost), zzz, ""
updateInventoryMaster "TNC", Me.StdCost, zzz, ""
DatabaseFunctions.ModifyItemCost zzzID, "Std Cost", IIf(Me.NewCost = "", "0", Me.NewCost)
DatabaseFunctions.ModifyItemCost zzzID, "New Cost", Me.StdCost
End If
End If
End If
        Dim temp As String
        temp = Me.NewCost
        Me.NewCost = Me.StdCost
        Me.StdCost = temp
        If Me.StdCost = "" Then
            Me.StdCost = "$0.00"
        End If
    End If
    'Tom's code to hopefully auto-update the d/s field...
    fillDropshipCostField
End Sub

Private Sub btnVendorQuantityCheck_Click()
    If Left(Me.ItemNumber, 3) = "XXX" Or Left(Me.ItemNumber, 3) = "ZZZ" Then
        MsgBox "XXX/ZZZ items unsupported!"
        Exit Sub
    End If
    
    Select Case Left(Me.ItemNumber, 3)
        Case Is = "JEA", "JET", "POA", "POW", "PER", "WIL"
            Dim req As MSXML2.XMLHTTP
            Set req = New MSXML2.XMLHTTP
            req.Open "GET", "http://toolsplus04/whse/vendor-qoh.plex?item1=" & Me.ItemNumber & "&format=SOLO"
            req.Send
            While req.ReadyState <> 4
                Sleep 100
                DoEvents
            Wend
            If req.status <> 200 Then
                MsgBox "Error connecting to WMH"
            Else
                Me.lblVendorQuantity.caption = TrimWhitespace(req.responseText, True, True, False, True)
                Me.lblVendorQuantity.ToolTipText = Now()
                'MsgBox "WMH has " & Me.lblVendorQuantity.Caption & " in stock"
            End If
        Case Else
            MsgBox "Vendor stock checking not implemented for this line!"
    End Select
End Sub

Private Sub btnViewInvQtyTriggers_Click()
   'Original code for inv qty trig #1
   ' Load InventoryQuantityTriggers
   ' InventoryQuantityTriggers.Show MODAL
   ' If DLookup("ItemNumber", "InventoryQuantityTriggers", "ItemNumber='" & Me.ItemNumber & "'") = Me.ItemNumber Then
   '     Me.btnViewInvQtyTriggers.FontBold = True
   ' Else
   '     Me.btnViewInvQtyTriggers.FontBold = False
   ' End If
   
   
       Load InventoryQuantityTriggersV2
    
   InventoryQuantityTriggersV2.qtyItemNumberLbl.caption = Me.jumpToLine.Text & Me.jumpToItem.Text 'Me.ItemNumber.Text
    
    'InventoryQuantityTriggersV2.dateItemNumberLbl.Caption = Me.ItemNumber.Text
    
    
    
    Dim itemTriggers As ADODB.Recordset
    Set itemTriggers = DB.retrieve("SELECT TriggerType,QtyThreshold,Below,DateThreshold,Notes FROM InventoryQuantityTriggers WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
    
    If itemTriggers.RecordCount = 0 Then
    'No triggers were set for this item...
    'InventoryQuantityTriggersV2.qtyNotActiveLbl.Visible = True
    'InventoryQuantityTriggersV2.dateNotActiveLbl.Visible = True
    'InventoryQuantityTriggersV2.dateNotActiveLbl.ZOrder 1
    'InventoryQuantityTriggersV2.qtyNotActiveLbl.ZOrder 1
    'InventoryQuantityTriggersV2.qtyActiveLbl.ZOrder 0
    'InventoryQuantityTriggersV2.dateActiveLbl.ZOrder 0
    'InventoryQuantityTriggersV2.qtyAlertLbl.ZOrder 0
    'InventoryQuantityTriggersV2.dateAlertLbl.ZOrder 0
    'InventoryQuantityTriggersV2.DisableDateSide
    'InventoryQuantityTriggersV2.DisableQtySide
    
    
    End If
    
    
    
    If itemTriggers.RecordCount <= 2 Then
    Dim anotherDate As Boolean
    Dim anotherQty As Boolean
        While (itemTriggers.EOF = False And itemTriggers.BOF = False)
            If itemTriggers("TriggerType") = "date" Then
                If anotherDate = False Then
                    Dim theDate As String
                    
                    If Len(itemTriggers("DateThreshold")) > 5 Then
                    theDate = itemTriggers("DateThreshold")
                   Else
                   theDate = Date
                   End If
                    Dim theDateNotes As String
                    If Len(itemTriggers("Notes")) > 0 Then
                    theDateNotes = itemTriggers("Notes")
                    Else
                    theDateNotes = ""
                    End If
                    Dim isDateAlerted As Boolean
                    'MsgBox "alert date is " & CDate(theDate) & " vs today which is " & Date
                    
                    If CDate(theDate) <= Date Then
                        isDateAlerted = True
                    Else
                        isDateAlerted = False
                    End If
                    'MsgBox CStr(isDateAlerted)

                    InventoryQuantityTriggersV2.SetDateInfo theDate, theDateNotes, isDateAlerted, False
                    InventoryQuantityTriggersV2.dateNotActiveLbl.Visible = False
                    InventoryQuantityTriggersV2.dateNotActiveLbl.ZOrder 1
                    anotherDate = True
                                       
                    
                Else
                    MsgBox "The item " & Me.ItemNumber.Text & " has multiple Date Triggers. Call Tom to remove them as they shouldn't exist."
                    
                End If
            Else
                'InventoryQuantityTriggersV2.SetDateInfo Date, "", False, True
                'InventoryQuantityTriggersV2.dateNotActiveLbl.Visible = True
                'InventoryQuantityTriggersV2.dateNotActiveLbl.ZOrder 0
                
            End If
     
            If itemTriggers("TriggerType") = "qty" Then
                If anotherQty = False Then
                    Dim theQty As String
                    theQty = itemTriggers("QtyThreshold")
                    Dim isBelow As String
                    isBelow = itemTriggers("Below")
                    Dim theQtyNotes As String
                    If Len(itemTriggers("Notes")) > 0 Then
                    theQtyNotes = itemTriggers("Notes")
                    Else
                    theQtyNotes = ""
                    End If
                    Dim isQtyAlerted As Boolean
                    If CInt(theQty) > (CInt(QtyOnHand.Text) - CInt(QtyOnSO.Text)) Then
                    'alert qty is greater than qoh... did we want above or below tho?
                        If isBelow = True Then
                            'yes, qoh is below alert qty so throw alert
                            isQtyAlerted = True
                        Else
                            'nope, qoh is still below the alert so no alert
                            isQtyAlerted = False
                        End If
                    
                    Else
                        If CInt(theQty) < (CInt(QtyOnHand.Text) - CInt(QtyOnSO.Text)) Then
                            'here, QOH is greater than alert qty...
                             If isBelow = True Then
                            'no, qoh is greater than alert qty... no alert
                                isQtyAlerted = False
                             Else
                            'yep, qoh is still greater than alert qty, so alert
                                isQtyAlerted = True
                             End If
                        
                        End If
                        
                        
                        If CInt(theQty) = (CInt(QtyOnHand.Text) - CInt(QtyOnSO.Text)) Then
                            'either above or below, set the alert cause it's value is equal to.
                            isQtyAlerted = True
                        End If
                        
                    
                    End If
                    InventoryQuantityTriggersV2.SetQtyInfo theQty, isBelow, theQtyNotes, isQtyAlerted, CStr(CInt(Me.QtyOnHand.Text) - CInt(Me.QtyOnSO.Text)), False
                    anotherQty = True
                Else
                    MsgBox "The item " & Me.ItemNumber.Text & " has multiple Quantity Triggers. Call Tom to remove them as they shouldn't exist."
                    
                End If
            Else
                'InventoryQuantityTriggersV2.SetQtyInfo "", "", "", False, Me.QtyOnHand.Text, True
            End If
        itemTriggers.MoveNext
        Wend
    End If
    
    If itemTriggers.RecordCount > 2 Then
    
        MsgBox "The item " & Me.ItemNumber.Text & " has more than two triggers in the database. This can only happen if there was a glitch in the Matrix. Call Tom."
        
    End If
    
    
    
    'this statement below will get deprecated soon by the statement directly above this...
    If DLookup("ItemNumber", "InventoryQuantityTriggers", "ItemNumber='" & Me.ItemNumber & "'") = Me.ItemNumber Then
        Me.btnViewInvQtyTriggers.FontBold = True
    Else
        Me.btnViewInvQtyTriggers.FontBold = False
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    If anotherDate = False Then InventoryQuantityTriggersV2.DisableDateSide
    If anotherQty = False Then InventoryQuantityTriggersV2.DisableQtySide
    
    Unload InventoryAlerts
    InventoryQuantityTriggersV2.LoadForm
    InventoryQuantityTriggersV2.Show
    
    If Me.InvQtyTriggerLoaded = False Then
   
        InventoryQuantityTriggersV2.Left = Me.Left + Me.width + 200
        InventoryQuantityTriggersV2.Top = Me.Top + 50
    
    Else

    InventoryQuantityTriggersV2.Left = InvQtyTriggerFrmX
    InventoryQuantityTriggersV2.Top = InvQtyTriggerFrmY
    End If
    
    
        InventoryQuantityTriggersV2.Command1.Enabled = False
   
       
    InventoryMaintenance.SetFocus
    
       
    
   
End Sub

Private Sub btnViewPOs_Click()
    Load PurchaseOrdersForm
    PurchaseOrdersForm.Show
End Sub

Private Sub btnKitCreator_Click()
    Load KitCreator
    KitCreator.SelectItem Me.ItemNumber
    KitCreator.Show
End Sub

Private Sub btnViewStockBOAlerts_Click()
    Load InventoryStockAlertRequests
    InventoryStockAlertRequests.Show
End Sub

Private Sub btnWeightDimsExtendedInfo_Click()
    Load InventoryComponents
    InventoryComponents.LoadItem Me.LoadedItemID
    InventoryComponents.Show MODAL
    If HasPermissionsTo("WeightDims") Then
        fillWeightDimensions Me.ItemNumber, Me.LoadedItemID
        ExternalComponentDBSync Me.ItemNumber
    End If
End Sub

Private Sub chkAddToXSheet_Click()
    If Not fillingForm Then
        If Me.chkAddToXSheet Then
            addToXSheet Me.ItemNumber
        Else
            addToXSheet Me.ItemNumber, True
        End If
    End If
End Sub

Private Sub chkDCbyMfr_Click()
    If Not fillingForm Then
        Dim oldstatus As String
        oldstatus = DLookup("ItemStatus", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'")
        If Me.chkDCbyMfr Then
            updateInventoryMaster "ItemStatus", SetItemDCbyMfr(oldstatus, DLookup("QtyOnHand", "vActualWhseQty", "ItemNumber='" & Me.ItemNumber & "'")), Me.ItemNumber, ""
            HandleDiscontinuation ItemNumber
        Else
            updateInventoryMaster "ItemStatus", UnSetItemDCbyMfr(oldstatus), Me.ItemNumber, ""
            If Not Me.chkDCbyUs Then
                SetUnDiscontinued ItemNumber
            End If
        End If
        refreshForm
    End If
End Sub

Private Sub chkDCbyUs_Click()
    If Not fillingForm Then
        Dim oldstatus As String
        oldstatus = DLookup("ItemStatus", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'")
        If Me.chkDCbyUs Then
            updateInventoryMaster "ItemStatus", SetItemDCbyUs(oldstatus, DLookup("QtyOnHand", "vActualWhseQty", "ItemNumber='" & Me.ItemNumber & "'")), Me.ItemNumber, ""
            HandleDiscontinuation ItemNumber
        Else
            updateInventoryMaster "ItemStatus", UnSetItemDCbyUs(oldstatus), Me.ItemNumber, ""
            If Not Me.chkDCbyMfr Then
                SetUnDiscontinued ItemNumber
            End If
        End If
        refreshForm
    End If
End Sub

Private Sub chkDropship_Click()
    If Not fillingForm Then
        If Me.chkDropship Then
            If 0 = DLookup("COALESCE(Orderable,1)", "PartNumbers LEFT OUTER JOIN AvailOverrides ON PartNumbers.AvailabilityOverride=AvailOverrides.ID", "PartNumbers.ItemNumber='" & Me.ItemNumber & "'") Then
                MsgBox "WARNING!" & vbCrLf & vbCrLf & "This item will still be unavailable on the web, since the availability string is overriding the d/s flag!" & vbCrLf & vbCrLf & "Just email this to Kirk if you don't know what it means."
            End If
            SetDropShippable Me.ItemNumber
        Else
            SetNotDropShippable Me.ItemNumber
        End If
        refreshForm
    End If
End Sub

Private Sub chkOrderPointLockStore_Click()
    If Not fillingForm Then
        chkLockedHereGenericHandler Me.chkOrderPointLockStore, 1
    End If
End Sub

Private Sub chkOrderPointLockWhse_Click()
    If Not fillingForm Then
        chkLockedHereGenericHandler Me.chkOrderPointLockWhse, 5
    End If
End Sub

Private Sub chkLockedHereGenericHandler(ctl As VB.CheckBox, whseid As Long)
    If Not fillingForm Then
        DB.Execute "UPDATE InventoryLocationInfo SET LockOrderPoint=" & ctl.value & " WHERE WarehouseID=" & whseid & " AND ItemNumber='" & Me.ItemNumber & "'"
        If ctl.value Then
            'things to do when locking?
        Else
            If vbYes = MsgBox("Calculate order point now?", vbYesNo) Then
                Dim m As Mas200ImportExport
                Set m = New Mas200ImportExport
                m.CalculateReorderPoints Me.ItemNumber
                Set m = Nothing
            End If
        End If
    End If
End Sub

Private Sub chkStockedHereStore_Click()
    If Not fillingForm Then
        chkStockedHereGenericHandler Me.chkStockedHereStore, 1, "store", Me.OrderPointStore, Me.chkOrderPointLockStore
    End If
End Sub

Private Sub chkStockedHereWhse_Click()
    If Not fillingForm Then
        chkStockedHereGenericHandler Me.chkStockedHereWhse, 5, "warehouse", Me.OrderPointWhse, Me.chkOrderPointLockWhse
    End If
End Sub

Private Sub chkStockedHereGenericHandler(ctl As VB.CheckBox, whseid As Long, whseDesc As String, opBox As VB.TextBox, lockCheck As VB.CheckBox)
    If Not fillingForm Then
        'fillingForm = True
        'If vbYes = MsgBox("Are you sure you want to " & IIf(ctl.value, "start", "stop") & " stocking this item at the " & whseDesc & "?", vbYesNo) Then
            'Dim newOP As String
            'newOP = IIf(ctl.value, "0", "NULL")
            'DB.Execute "UPDATE InventoryLocationInfo SET StockHere=" & ctl.value & ", OrderPoint=" & newOP & ", LockOrderPoint=0 WHERE WarehouseID=" & whseid & " AND ItemNumber='" & Me.ItemNumber & "'"
            'opBox.Enabled = CBool(ctl.value)
            'lockCheck.Enabled = CBool(ctl.value)
            'If ctl.value Then
            '    'things to do when stocking? calc op?
            '    opBox = "0"
            '    If vbYes = MsgBox("Calculate order point now?", vbYesNo) Then
            '        Dim m As Mas200ImportExport
            '        Set m = New Mas200ImportExport
            '        m.CalculateReorderPoints Me.ItemNumber
            '        Set m = Nothing
            '        opBox = DLookup("OrderPoint", "InventoryLocationInfo", "WarehouseID=" & whseid & " AND ItemNumber='" & Me.ItemNumber & "'")
            '    End If
            'Else
            '    opBox = "--"
            '    'Me.OrderPointTotal = CLng(Me.OrderPointStore) + CLng(Me.OrderPointWhse)
            '    fillTotalOrderPoint
            'End If
            DB.Execute "UPDATE InventoryLocationInfo SET StockHere=" & ctl.value & " WHERE WarehouseID=" & whseid & " AND ItemNumber='" & Me.ItemNumber & "'"
        'Else
        '    ctl.value = IIf(ctl.value, 0, 1)
        'End If
        'fillingForm = False
    End If
End Sub

'Private Sub chkNeedsBox_Click()
'    If Not fillingForm Then
'        updateInventoryMaster "NeedsBox", Me.chkNeedsBox, Me.ItemNumber, ""
'        'updateExternalShippingDB Me.ItemNumber
'        'updateShippingDB Me.ItemNumber
'        UpdateASODatabase ItemNumber
'    End If
'End Sub

Private Sub chkReconditioned_Click()
    If Not fillingForm Then
        updateInventoryMaster "IsReconditioned", Me.chkReconditioned, Me.ItemNumber, ""
    End If
End Sub

Private Sub chkStock_Click()
    If Not fillingForm Then
        Dim oldstatus As String
        oldstatus = DLookup("ItemStatus", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'")
        If Me.chkStock Then
            updateInventoryMaster "ItemStatus", SetItemStocked(oldstatus), Me.ItemNumber, ""
        Else
            Dim qoh As Long, qoo As Long
            qoh = DLookup("QtyOnHand", "vActualWhseQty", "ItemNumber='" & Me.ItemNumber & "'")
            qoo = DLookup("QtyOnPO", "vActualWhseQty", "ItemNumber='" & Me.ItemNumber & "'")
            If qoh > 0 Then
                If MsgBox("Are you sure you want to un-stock an item with qty on hand? You should probably just let it discontinue itself when it gets to 0.", vbYesNo + vbDefaultButton2) = vbNo Then
                    fillingForm = True
                    Me.chkStock = 1
                    fillingForm = False
                    Exit Sub
                End If
            ElseIf qoo > 0 Then
                If vbNo = MsgBox("Are you sure you want to un-stock an item with qty on order? You should probably run a refresh qtys and let it discontinue itself.", vbYesNo + vbDefaultButton2) Then
                    fillingForm = True
                    Me.chkStock = 1
                    fillingForm = False
                    Exit Sub
                End If
            End If
            updateInventoryMaster "ItemStatus", UnSetItemStocked(oldstatus), Me.ItemNumber, ""
            If IsItemDC(oldstatus) Then
                Dim availLimit As Long
                availLimit = DLookup("AvailabilityLimit", "PartNumbers", "ItemNumber='" & Me.ItemNumber & "'")
                If qoh = 0 Then
                    SetDiscontinued Me.ItemNumber
                ElseIf qoh <= availLimit Then
                    SetWebDiscontinued Me.ItemNumber
                'else not d/c yet, so we're good
                End If
            End If
        End If
        refreshForm
    End If
End Sub

Private Sub opPeriodType_Click(index As Integer)
    fillSalesHistory Me.ItemNumber, Me.chkDropship
End Sub

Private Sub Command1_Click()
IMtt.AddToolTipToCtl Me.Command1, "this should show"
MsgBox "Now hover over this button and should show balloon tip"

End Sub

Private Sub Command2_Click()
JetDotCom.RemoveOldItemsFromJet

End Sub

Private Sub Command3_Click()
ExternalComponentDBSync Me.ItemNumber.Text


End Sub

Private Sub Command5_Click()
If Me.EBayComparisonsLVI.SortKey = 4 Then
    If Me.EBayComparisonsLVI.SortOrder = lvwAscending Then
        Me.EBayComparisonsLVI.SortOrder = lvwDescending
        SortNumericEBay
    Else
        Me.EBayComparisonsLVI.SortOrder = lvwAscending
        SortNumericEBay
    End If

Else
    Me.EBayComparisonsLVI.SortKey = 4
    Me.EBayComparisonsLVI.SortOrder = lvwAscending
    Me.EBayComparisonsLVI.sorted = True
    SortNumericEBay
    
End If
End Sub

Private Sub Command4_Click()

End Sub

Private Sub conditionLbl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim direction As Integer

If Button = 2 Then direction = -1 Else direction = 1

'get next condition
'update condition in db
'set offloadrequired = true

'next condition is either current+1 or 1 (if current is highest id #)

Dim SetOffload As Boolean
Dim ErrorMessage As Boolean
Dim NewConditionID As Integer

Dim nextCondition As ADODB.Recordset
Set nextCondition = DB.retrieve("SELECT ConditionID FROM ItemConditionMaster WHERE ConditionName='" & Me.conditionLbl.caption & "'")
If nextCondition.RecordCount > 0 Then
 Dim id As Integer
 id = CInt(nextCondition("ConditionID"))
 Set nextCondition = DB.retrieve("SELECT ConditionID,ConditionName FROM ItemConditionMaster WHERE ConditionID=" & CStr(id + direction))
 If nextCondition.RecordCount > 0 Then
    conditionLbl.caption = nextCondition("ConditionName")
    NewConditionID = CInt(nextCondition("ConditionID"))
    'set offload required...
    SetOffload = True
 Else
    If direction > 0 Then
    Set nextCondition = DB.retrieve("SELECT ConditionID,ConditionName FROM ItemConditionMaster WHERE ConditionID=1")
    Else
    Set nextCondition = DB.retrieve("SELECT TOP 1 ConditionID,ConditionName FROM ItemConditionMaster ORDER BY ConditionID DESC")
    End If
    If nextCondition.RecordCount > 0 Then
        conditionLbl.caption = nextCondition("ConditionName")
        NewConditionID = CInt(nextCondition("ConditionID"))
        'set offload required...
        SetOffload = True
    Else
        'error...
        ErrorMessage = True
    End If
 End If
Else
 'error....
 ErrorMessage = True
End If

If ErrorMessage = True Then
    MsgBox "Something bad happened with the Condition of this item. Tell Tom its his fault somehow even though he wont believe you."
Exit Sub
End If
If SetOffload = True Then
    DB.Execute "UPDATE ItemConditionLines SET ConditionID=" & CStr(NewConditionID) & " WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
    
    Dim isOnEbay As ADODB.Recordset
    Set isOnEbay = DB.retrieve("SELECT EBayPublished FROM PartNumbers WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND EbayPublished=1")
    If isOnEbay.RecordCount > 0 Then
        DB.Execute "UPDATE PartNumbers SET EBayOffloadRequired=1 WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
    Else
     'error or not on ebay... do nothing..
    End If
    
    
End If
End Sub

Private Sub dbughandleBtn_Click()
'PriceCompareCSharp.GetPriceComparerEBayHandle
'GetPriceComparerStatusHandles



'PriceCompareCSharp.GetPriceComparerEBayStatusHandle

'If PriceCompareCSharp.EBayIEHandle > 0 Then
' Dim testDoc As HTMLDocument
'
'
' Set testDoc = IEDOMFromhWnd(PriceCompareCSharp.EBayIEHandle)
'
' MsgBox testDoc.documentElement.outerHTML
 
'
'End If

End Sub

Private Sub DeleteSelectedSubtitleBtn_Click()

'need to find subtitle in db then find how many items use that specific subtitle...
Dim records As ADODB.Recordset
Set records = DB.retrieve("SELECT ID FROM Subtitles WHERE Subtitle='" & Me.SubtitleCmb.Text & "'")

Dim recCount As Integer
recCount = records.RecordCount


Dim res As VbMsgBoxResult
res = MsgBox("Are you sure you want to delete this subtitle? " & CStr(recCount) & " items use this current subtitle and will revert to default subtitle.", vbYesNo, "Delete Subtitle")
If res = vbYes Or res = vbOK Then
If Me.SubtitleCmb.Text <> "99.9%fb-Auth Seller-Warranty-Free Ship-$$ Back Guar" Then
Dim SelectedSubtitle As String
SelectedSubtitle = Me.SubtitleCmb.Text
'we need to default all subs back to default...
Dim defSub As ADODB.Recordset
Set defSub = DB.retrieve("SELECT ID FROM Subtitles WHERE Subtitle='" & Me.SubtitleCmb.Text & "'")
Dim subID As Integer
subID = CInt(defSub("ID"))
DB.retrieve "DELETE FROM Subtitles WHERE Subtitle='" & SelectedSubtitle & "'"
DB.Execute "UPDATE SubtitledItems SET SubtitleID=1 WHERE SubtitleID=" & CStr(subID)
'now to repopulate combobox...
Me.SubtitleCmb.Clear

Dim SubtitleResults As ADODB.Recordset
Set SubtitleResults = DB.retrieve("SELECT Subtitle FROM Subtitles")
Do

    Me.SubtitleCmb.AddItem (SubtitleResults("Subtitle"))
    SubtitleResults.MoveNext

Loop Until SubtitleResults.EOF = True Or SubtitleResults.BOF = True
Else
MsgBox "Hey you! Don't attempt to remove the default subtitle... if it needs to be altered let Tom know."
End If
End If
End Sub

Private Sub DiscountMarkupPriceRate_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.HidePopUp = True
Me.HidePopUpFrm
Sleep 300
DiscountMarkupPriceRate(index).ZOrder 0
DiscountMarkupPriceRate(index).SetFocus
DiscountMarkupPriceRate(index).Refresh
End Sub

Private Sub DropshipCost_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.HidePopUp = True
Me.HidePopUpFrm
Sleep 300
DropshipCost.ZOrder 0
DropshipCost.SetFocus
DropshipCost.Refresh
End Sub

Private Sub DropshipCost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'DropshipCost.BackColor = RGB(0, 255, 255)
            DropshipCost.Refresh
    
            'If Me.HidePopUp = False Then
                'PopUpFrm.Visible = True
                'Me.popupTmr.Enabled = True
                'PopUpFrm.InfoLbl.Caption = Me.CurrentDSCostdata
                'PopUpFrm.ZOrder 0
                'PopUpFrm.SetFocus
                'PopUpFrm.Move Me.Left + Me.FramePricingInfo.Left - 500, Me.Top + Me.FramePricingInfo.Top + 50 + DropshipCost.Top   'ScaleX(NewCost.Left, vbPixels, vbTwips), ScaleY(NewCost.Top, vbPixels, vbTwips) ', PopUpFrm.width, PopUpFrm.Height
                'PopUpFrm.Refresh
            'End If
End Sub

Private Sub dynamicPriceChk_Click()
Dim itemExists As Boolean
Dim itemCheck As ADODB.Recordset
Set itemCheck = DB.retrieve("SELECT ID,MinPrice,NewPrice,EBay,TPStore,TPWeb,JetDotCom FROM PriceWiserItems WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
If itemCheck.RecordCount > 0 Then
itemExists = True
'If itemCheck("EBay") = True Then DynamicPricingFrm.dynamicEBayChk.value = vbChecked
'If itemCheck("TPStore") = True Then DynamicPricingFrm.dynamicStoreChk.value = vbChecked
'If itemCheck("TPWeb") = True Then DynamicPricingFrm.dynamicSiteChk.value = vbChecked
'If itemCheck("JetDotCom") = True Then DynamicPricingFrm.dynamicJetChk.value = vbChecked
Else
itemExists = False
End If


If dynamicPriceChk.value = vbChecked Then
'enable this item in dynamic pricing...
If itemExists = True Then
    'DynamicPricingFrm.ListView1.Enabled = True
    'DynamicPricingFrm.reserveMinTxt.Enabled = True
    'DynamicPricingFrm.lastCheckedLbl.Enabled = True
    'DynamicPricingFrm.itemNumberLbl.Enabled = True
    'DynamicPricingFrm.e
    
    
    'it's there already... now make sure it's enabled...
    DB.Execute "UPDATE PriceWiserItems SET Enabled=1 WHERE ID=" & itemCheck("ID")
    'lastPriceLbl.Visible = True
    dynamicPriceEditBtn.Visible = True
    dynamicPriceChk.caption = "Dynamic Price: $" & itemCheck("NewPrice")
    'DynamicPricingFrm.applyBtn.Enabled = False
    dynamicPriceEditBtn.Left = dynamicPriceChk.Left + dynamicPriceChk.width
    dynamicPriceEditBtn.Refresh
    
    
Else
    'DynamicPricingFrm.dynamicEBayChk.Enabled = False
    'DynamicPricingFrm.dynamicJetChk.Enabled = False
    'DynamicPricingFrm.dynamicPriceLbl.Enabled = False
    'DynamicPricingFrm.dynamicSiteChk.Enabled = False
    'DynamicPricingFrm.dynamicStoreChk.Enabled = False
 'Me.dynEBayLbl.Visible = False
 ' Me.dynTPStoreLbl.Visible = False
  'Me.dynTPSiteLbl.Visible = False
  ' Me.dynJetDotComLbl.Visible = False
    'it's not there... but we dont want to insert. notify the user this item is not on pricewiser...
    'DB.Execute "INSERT INTO PriceWiserItems (ItemNumber,MinPrice,MaxPrice,Price,NewPrice,LastUpdate,Enabled) VALUES('" & Me.ItemNumber.Text & "',0,0,0,0,'" & Now & "')"
    MsgBox "This item is currently not on PriceWiser, or has not returned any data yet. This means you cannot choose to dynamic price this item until it's been populated/added on PriceWiser."
    dynamicPriceChk.value = vbUnchecked
    'lastPriceLbl.Visible = False
    dynamicPriceEditBtn.Visible = False
    Exit Sub
End If

Else
'disable this item in dynamic pricing...
'lastPriceLbl.Visible = False
dynamicPriceEditBtn.Visible = False
'Me.dynEBayLbl.Visible = False
'Me.dynTPSiteLbl.Visible = False
'Me.dynTPStoreLbl.Visible = False
'Me.dynJetDotComLbl.Visible = False
If itemExists = True Then
    'disable the item..
    DB.Execute "UPDATE PriceWiserItems SET Enabled=0 WHERE ID=" & itemCheck("ID")
Else
    'it's not there and we don't want it... so do nothing!
End If

End If
End Sub

Private Sub dynamicPriceEditBtn_Click()
Me.CheckDynamicPriceEnabled
'DynamicPricingFrm.Visible = True
'fillDynamicPricing

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

Private Sub EBayComparisonsLst_Click()
Me.InternetComparisonLst.Selected(0) = True
Me.priceComparison.Selected(0) = True
'12/12/05 - access to single item compare for everyone
'    If Not lockedDownPC Then
        If Me.EBayComparisonsLst.Selected(0) Then
            Me.btnPriceCompareDetails.Enabled = False
        Else
            Me.btnPriceCompareDetails.Enabled = True
        End If
'    End If
End Sub

Private Sub EBayComparisonsLVI_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim anyCheck As Boolean
Dim itm As ListItem
For Each itm In EBayComparisonsLVI.ListItems
    If itm.Checked = True Then anyCheck = True
    
    If itm.SubItems(1) <> item.SubItems(1) Then
        itm.Checked = False
    Else
        If itm.Checked = True Then
            Dim selAlg As ListItem
            For Each selAlg In Me.EBayDynamicAlgorithmLVI.ListItems
                If selAlg.Checked = True Then
                    DB.Execute "UPDATE DynamicPricing SET Enabled=1,AlgorithmID=" & selAlg.tag & " WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPEBay'"
                End If
            Next
        End If
    End If

Next

If anyCheck = False Then
    DB.Execute "UPDATE DynamicPricing SET Enabled = 0,CompetitorDataID=0 WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPEBay'"
    Dim price As ADODB.Recordset
    Set price = DB.retrieve("SELECT OriginalPrice1 FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPEBay'")
        If price.RecordCount > 0 Then
            Me.EBayPrice.Text = Format(price("OriginalPrice1"), "Currency")
            DB.Execute "UPDATE PartNumbers SET EBayToBePublished=1 WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
        End If
End If

CheckForEBayDynamicPrice



End Sub

Private Sub EBayDynamicAlgorithmLVI_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim hasACheck As Boolean
Dim checkCheck As ListItem
For Each checkCheck In Me.EBayDynamicAlgorithmLVI.ListItems
    If checkCheck.Checked = True Then hasACheck = True
Next
If hasACheck = False Then

    DB.Execute "UPDATE DynamicPricing SET AlgorithmID=0,Enabled=0,CompetitorDataID=0 WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPEBay'"
'now reload the orig. price...
Dim price As ADODB.Recordset
Set price = DB.retrieve("SELECT OriginalPrice1 FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPEBay'")
If price.RecordCount > 0 Then
    Me.EBayPrice.Text = Format(price("OriginalPrice1"), "Currency")
    DB.Execute "UPDATE PartNumbers SET EBayToBePublished=1 WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
End If
End If


Dim itm As ListItem
For Each itm In EBayDynamicAlgorithmLVI.ListItems
    If itm.SubItems(1) <> item.SubItems(1) Then
        itm.Checked = False
    End If
Next

CheckForEBayDynamicPrice
End Sub

Private Sub EBayPrice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.HidePopUp = True
Me.HidePopUpFrm
Sleep 300
EBayPrice.ZOrder 0
EBayPrice.SetFocus
EBayPrice.Refresh
End Sub

Private Sub ebaySortPriceBtn_Click()
If Me.EBayComparisonsLVI.SortKey = 1 Then
    If Me.EBayComparisonsLVI.SortOrder = lvwAscending Then
        Me.EBayComparisonsLVI.SortOrder = lvwDescending
        SortNumericEBay
    Else
        Me.EBayComparisonsLVI.SortOrder = lvwAscending
        SortNumericEBay
    End If
Else
    Me.EBayComparisonsLVI.SortKey = 1
    Me.EBayComparisonsLVI.SortOrder = lvwAscending
    Me.EBayComparisonsLVI.sorted = True
    SortNumericEBay
End If
End Sub

Private Sub ebaySortShippingBtn_Click()
If Me.EBayComparisonsLVI.SortKey = 2 Then
    If Me.EBayComparisonsLVI.SortOrder = lvwAscending Then
        Me.EBayComparisonsLVI.SortOrder = lvwDescending
        SortNumericEBay
    Else
        Me.EBayComparisonsLVI.SortOrder = lvwAscending
        SortNumericEBay
    End If
Else
    Me.EBayComparisonsLVI.SortKey = 2
    Me.EBayComparisonsLVI.SortOrder = lvwAscending
    Me.EBayComparisonsLVI.sorted = True
    SortNumericEBay
End If
End Sub

Public Sub ebaySortTotalBtn_Click()
If Me.EBayComparisonsLVI.SortKey = 4 Then
    If Me.EBayComparisonsLVI.SortOrder = lvwAscending Then
        Me.EBayComparisonsLVI.SortOrder = lvwDescending
        SortNumericEBay
    Else
        Me.EBayComparisonsLVI.SortOrder = lvwAscending
        SortNumericEBay
    End If
Else
    'Me.EBayComparisonsLVI.SortKey = 4
    Me.EBayComparisonsLVI.SortKey = 2
    Me.EBayComparisonsLVI.SortOrder = lvwAscending
    Me.EBayComparisonsLVI.sorted = True
    SortNumericEBay
End If
End Sub

Private Sub ebaystatusbtn_Click()
MsgBox PriceCompareCSharp.GetObjectText(PriceCompareCSharp.EBayStatusHandle)
End Sub

Private Sub eOffBtn_Click()
WebOffload.btnEBayOffload_Click 4
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.popupTmr.Enabled = False
'HidePopUpFrm
'Me.HidePopUp = False
End Sub

Private Sub FrameEBayPrice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.popupTmr.Enabled = False
'HidePopUpFrm
'Me.HidePopUp = False
End Sub

Private Sub frameInstorePrice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.popupTmr.Enabled = False
'HidePopUpFrm
'Me.HidePopUp = False
End Sub

Private Sub frameInternetPrice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.popupTmr.Enabled = False
'HidePopUpFrm
'Me.HidePopUp = False
End Sub

Private Sub FrameMiscPrice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.popupTmr.Enabled = False
'HidePopUpFrm
'Me.HidePopUp = False
End Sub

Private Sub framePOs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.popupTmr.Enabled = False
'HidePopUpFrm
'Me.HidePopUp = False
End Sub

Private Sub FramePricingInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.popupTmr.Enabled = False
'Me.HidePopUp = False
'Me.HidePopUpFrm

End Sub

Private Sub frameSalesHistory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.popupTmr.Enabled = False
'HidePopUpFrm
'Me.HidePopUp = False
End Sub

Private Sub IDiscountMarkupPriceRate_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.popupTmr.Enabled = False
'Me.HidePopUp = True
'Me.HidePopUpFrm
Sleep 300
IDiscountMarkupPriceRate(index).ZOrder 0
IDiscountMarkupPriceRate(index).SetFocus
IDiscountMarkupPriceRate(index).Refresh
End Sub

Private Sub inetAmazonBtn_Click()
If inetAmazonBtn.caption = "G" Then
    inetAmazonBtn.caption = "A"
Else
    inetAmazonBtn.caption = "G"
End If
requeryPriceComparisons
    Me.inetSortAscendingTotalBtn_Click
End Sub

Private Sub inetFroogleBtn_Click()

End Sub

Private Sub inetSortPriceBtn_Click()
If Me.InternetComparisonLVI.SortKey = 1 Then
    If Me.InternetComparisonLVI.SortOrder = lvwAscending Then
        Me.InternetComparisonLVI.SortOrder = lvwDescending
        SortNumeric
    Else
        Me.InternetComparisonLVI.SortOrder = lvwAscending
        SortNumeric
    End If
Else
    Me.InternetComparisonLVI.SortKey = 1
    Me.InternetComparisonLVI.SortOrder = lvwAscending
    Me.InternetComparisonLVI.sorted = True
    SortNumeric
End If
End Sub

Private Sub inetSortShippingBtn_Click()
If Me.InternetComparisonLVI.SortKey = 2 Then
    If Me.InternetComparisonLVI.SortOrder = lvwAscending Then
        Me.InternetComparisonLVI.SortOrder = lvwDescending
        SortNumeric
    Else
        Me.InternetComparisonLVI.SortOrder = lvwAscending
        SortNumeric
    End If

Else
    Me.InternetComparisonLVI.SortKey = 2
    Me.InternetComparisonLVI.SortOrder = lvwAscending
    Me.InternetComparisonLVI.sorted = True
    SortNumeric
End If
End Sub

Public Sub inetSortTotalBtn_Click()
If Me.InternetComparisonLVI.SortKey = 4 Then
    If Me.InternetComparisonLVI.SortOrder = lvwAscending Then
        Me.InternetComparisonLVI.SortOrder = lvwDescending
        SortNumeric
        Me.orderAscending = False
    Else
        Me.InternetComparisonLVI.SortOrder = lvwAscending
        SortNumeric
        Me.orderAscending = True
    End If

Else
    Me.InternetComparisonLVI.SortKey = 2
    Me.InternetComparisonLVI.SortOrder = lvwAscending
    Me.InternetComparisonLVI.sorted = True
    SortNumeric
    
    
End If
End Sub
Public Sub inetSortAscendingTotalBtn_Click()

    Me.InternetComparisonLVI.SortKey = 2
    Me.InternetComparisonLVI.SortOrder = lvwAscending
    Me.InternetComparisonLVI.sorted = True
    SortNumeric
    

End Sub

Private Sub InternetAlgorithmCmb_Click()
 If InternetAlgorithmCmb.Text <> "-Choose an Algorithm-" And InternetAlgorithmCmb.Text <> "Do Not Dynamically Price" Then
     InetDynamicPriceCheck
 Else
    Dim comparisons As ListItem
    For Each comparisons In Me.InternetComparisonLVI.ListItems
        'comparisons.SubItems(3) = "N/A"
        If comparisons.Checked = True Then
        
            comparisons.Checked = False
            Dim origPrice As ADODB.Recordset
            Set origPrice = DB.retrieve("SELECT OriginalPrice1 FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPWeb'")
            If origPrice.RecordCount > 0 Then
                Me.IDiscountMarkupPriceRate(1).Text = Format(origPrice("OriginalPrice1"), "Currency")
            End If
        End If
    Next
 End If
 
End Sub

Private Sub InternetComparisonLst_Click()
Me.EBayComparisonsLst.Selected(0) = True
Me.priceComparison.Selected(0) = True
'12/12/05 - access to single item compare for everyone
'    If Not lockedDownPC Then
        If Me.InternetComparisonLst.Selected(0) Then
            Me.btnPriceCompareDetails.Enabled = False
        Else
            Me.btnPriceCompareDetails.Enabled = True
        End If
'    End If
End Sub

Private Sub InternetComparisonLVI_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim anyChecked As Boolean
Dim itm As ListItem
For Each itm In InternetComparisonLVI.ListItems
    If itm.Checked = True Then anyChecked = True

    If itm.SubItems(1) <> item.SubItems(1) Or itm.SubItems(2) <> item.SubItems(2) Then
        itm.Checked = False
    Else
        If itm.Checked = True Then
            Dim selAlg As ListItem
            'For Each selAlg In Me.InternetDynamicAlgorithmLVI.ListItems
            '    If selAlg.Checked = True Then
                    Dim competitorID As ADODB.Recordset
                    Set competitorID = DB.retrieve("SELECT ID FROM PriceComparisons WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND StoreName='" & itm.SubItems(1) & "' AND Price=" & itm.SubItems(2))
                    If competitorID.RecordCount > 0 Then
                        DB.Execute "UPDATE DynamicPricing SET Enabled=1,CompetitorDataID=" & competitorID("ID") & " WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPWeb'"
                        'MsgBox "here"
                    Else
                        MsgBox "This item should have a dynamic pricing db entry but doesn't. Tell Tom."
                        'DB.Execute "INSERT INTO DynamicPricing (ItemNumber,Enabled,CompetitorDataID,PriceFieldToChange,OriginalPrice1,AlgorithmPercentageOverride) VALUES('" & Me.ItemNumber.Text & "',1,"
                    End If
                    
            '    End If
            'Next
        
        
            
            
        End If
    End If
Next

If anyChecked = False Then
    DB.Execute "UPDATE DynamicPricing SET Enabled=0, CompetitorDataID=0 WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPWeb'"
    Dim price As ADODB.Recordset
Set price = DB.retrieve("SELECT OriginalPrice1 FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPWeb'")
If price.RecordCount > 0 Then
    Me.IDiscountMarkupPriceRate(1).Text = Format(price("OriginalPrice1"), "Currency")
    
End If
End If


CheckForInetDynamicPrice
'SetInetDynamicPrice
End Sub
Private Sub SetInetDynamicPrice()



End Sub


Private Sub CheckForEBayDynamicPrice()
If Me.DynamicDontMessWith = False Then
    Dim compFlag As Boolean
    Dim algFlag As Boolean
    Dim comparisons As ListItem
    Dim itm As ListItem
    
    For Each comparisons In Me.EBayComparisonsLVI.ListItems
        If comparisons.Checked = True Then
            compFlag = True
        End If
        For Each itm In Me.EBayDynamicAlgorithmLVI.ListItems
            If itm.Checked = True Then
                algFlag = True
                Dim dynCheck As ADODB.Recordset
                Set dynCheck = DB.retrieve("DECLARE @input DECIMAL(7,3) SET @input=" & Format(comparisons.SubItems(2), "Currency") & " SELECT TOP 1 DynamicPricingAlgorithmLines.ID,HeaderID,DifferentialPercentage,DollarAmount,MaxPercentAbove,MaxPercentBelow FROM DynamicPricingAlgorithmLines INNER JOIN DynamicPricingAlgorithmHeader ON DynamicPricingAlgorithmHeader.ID=HeaderID WHERE HeaderID=" & itm.tag & " ORDER BY ABS(@input - DollarAmount)")
                If dynCheck.RecordCount > 0 Then
                  Dim NewMin As Currency
                  Dim NewMax As Currency
                  Dim CurrPrice As Currency
                  
                  NewMin = Format(comparisons.SubItems(2), "Currency") - (Format(comparisons.SubItems(2), "Currency") * (dynCheck("MaxPercentBelow") / 100))
                  NewMax = Format(comparisons.SubItems(2), "Currency") + (Format(comparisons.SubItems(2), "Currency") * (dynCheck("MaxPercentAbove") / 100))
                  
                  Dim ovrChk As ADODB.Recordset
                  Set ovrChk = DB.retrieve("SELECT AlgorithmPercentage FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPEBay' AND AlgorithmPercentageOverride=1")
                  If ovrChk.RecordCount > 0 Then
                        CurrPrice = Format(comparisons.SubItems(2), "Currency") + (Format(comparisons.SubItems(2), "Currency") * (ovrChk("AlgorithmPercentage") / 100))
                  Else
                        CurrPrice = Format(comparisons.SubItems(2), "Currency") + (Format(comparisons.SubItems(2), "Currency") * (dynCheck("DifferentialPercentage") / 100))
                  End If
                  If CurrPrice > NewMin And CurrPrice < NewMax Then
                    'comparisons.SubItems(3) = Format(CurrPrice, "Currency")
                  Else
                    If CurrPrice <= NewMin Then
                     'comparisons.SubItems(3) = Format(NewMin, "Currency")
                     CurrPrice = NewMin
                    Else
                    'comparisons.SubItems(3) = Format(NewMax, "Currency")
                    CurrPrice = NewMax
                    End If
                  End If
                End If
            End If
        Next
    Next

    If compFlag = False Then
        Dim origData As ADODB.Recordset
        Set origData = DB.retrieve("SELECT OriginalPrice1 FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPEBay'")
        If origData.RecordCount > 0 Then
            DB.Execute "UPDATE InventoryMaster SET EBayPrice=" & Format(origData("OriginalPrice1"), "Currency") & " WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
            
        Else
        End If
    End If

    If algFlag = False Then
        For Each comparisons In Me.EBayComparisonsLVI.ListItems
            'comparisons.SubItems(3) = "N/A"
            comparisons.Checked = False
        Next
    End If

    Dim compareFlag As Boolean
    For Each comparisons In Me.EBayComparisonsLVI.ListItems
        If comparisons.Checked = True Then
            compareFlag = True
            If algFlag = True Then
                OriginalEBayPrice = Me.EBayPrice.Text
                Me.EBayPrice = Format(CurrPrice, "Currency")
                DB.Execute "UPDATE InventoryMaster SET EBayPrice=" & Format(CurrPrice, "Currency") & " WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
                DB.Execute "UPDATE PartNumbers SET EBayToBePublished=1 WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
                Dim competitorDataID As String
                Dim AlgorithmID As String
                Dim algorithmPercentage As String
                Dim priceFieldToChange As String
                Dim priceOne As Currency
                competitorDataID = CStr(comparisons.tag)
                Dim algList As ListItem
                For Each algList In Me.EBayDynamicAlgorithmLVI.ListItems
                    If algList.Checked = True Then
                        
                        AlgorithmID = CStr(algList.tag)
                        
                        Dim algPerc As ADODB.Recordset
                        Set algPerc = DB.retrieve("SELECT AlgorithmPercentage FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPEBay'")
                        If algPerc.RecordCount > 0 Then
                        'use the set percentage
                            algorithmPercentage = CStr(Format(algPerc("AlgorithmPercentage"), "Currency"))
                        Else
                        'use the default
                            algorithmPercentage = CStr(algList.SubItems(2))
                        End If
                    End If
                Next
                
                priceFieldToChange = "TPEBay"
                priceOne = Format(Me.EBayPrice.Text, "Currency")
                If InStr(algorithmPercentage, "%") > 0 Then algorithmPercentage = Replace(algorithmPercentage, "%", "")
                Dim existsOrNot As ADODB.Recordset
                Set existsOrNot = DB.retrieve("SELECT ID FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPEBay'")
                If existsOrNot.RecordCount > 0 Then
                    DB.Execute "UPDATE DynamicPricing SET Enabled=1,AlgorithmID=" & AlgorithmID & ", CompetitorDataID=" & competitorDataID & ",AlgorithmPercentage=" & algorithmPercentage & " WHERE ID=" & existsOrNot("ID")
                Else
                    DB.Execute "INSERT INTO DynamicPricing (ItemNumber,Enabled,AlgorithmID,CompetitorDataID,AlgorithmPercentage,PriceFieldToChange,OriginalPrice1) VALUES('" & Me.ItemNumber.Text & "',1," & AlgorithmID & "," & competitorDataID & "," & algorithmPercentage & ",'TPEBay'," & OriginalEBayPrice & ")"
                End If
                
                competitorDataID = CStr(comparisons.tag)
                
                changed = True
                
                
            End If
        End If
    Next
    If compareFlag = False Then
        'Me.EBayPrice.Text = OriginalEBayPrice
        changed = False
    End If
End If

End Sub



Private Sub CheckForInetDynamicPrice()
If Me.DynamicDontMessWith = False Then
 Dim compFlag As Boolean
 Dim algFlag As Boolean
 Dim comparisons As ListItem
 Dim HeaderID As String
 Dim getHeader As ADODB.Recordset
 Set getHeader = DB.retrieve("SELECT ID FROM DynamicPricingAlgorithmHeader WHERE AlgorithmName='" & Split(Me.InternetAlgorithmCmb.list(Me.InternetAlgorithmCmb.ListIndex), ":")(0) & "'")
 If (getHeader.RecordCount > 0) Then
    HeaderID = getHeader("ID")
 End If
 

 For Each comparisons In Me.InternetComparisonLVI.ListItems
    If comparisons.Checked = True Then
        compFlag = True
    End If
    If Me.InternetAlgorithmCmb.Text = "Do Not Dynamically Price" Then
        algFlag = False
    Else
        algFlag = True
    End If
       If comparisons.Checked = True Then
           algFlag = False
           Dim dynCheck As ADODB.Recordset
           Set dynCheck = DB.retrieve("DECLARE @input DECIMAL(7,3) SET @input=" & Format(comparisons.SubItems(2), "Currency") & " SELECT TOP 1 DynamicPricingAlgorithmLines.ID,HeaderID,DifferentialPercentage,DollarAmount,MaxPercentAbove,MaxPercentBelow FROM DynamicPricingAlgorithmLines INNER JOIN DynamicPricingAlgorithmHeader ON DynamicPricingAlgorithmHeader.ID=HeaderID WHERE HeaderID=" & HeaderID & " ORDER BY ABS(@input - DollarAmount)")
            If dynCheck.RecordCount > 0 Then
                 'NewPrice = Format(CompetitorPrice, "Currency") + (Format(CompetitorPrice, "Currency") * (DynPrice("DifferentialPercentage") / 100))
                 Dim NewMin As Currency
                 Dim NewMax As Currency
                 Dim CurrPrice As Currency
                 NewMin = Format(comparisons.SubItems(2), "Currency") - (Format(comparisons.SubItems(2), "Currency") * (dynCheck("MaxPercentBelow") / 100))
                 NewMax = Format(comparisons.SubItems(2), "Currency") + (Format(comparisons.SubItems(2), "Currency") * (dynCheck("MaxPercentAbove") / 100))
                 
                 Dim ovrChk As ADODB.Recordset
                 Set ovrChk = DB.retrieve("SELECT AlgorithmPercentage FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPWeb' AND AlgorithmPercentageOverride=1")
                 If ovrChk.RecordCount > 0 Then
                    CurrPrice = Format(comparisons.SubItems(2), "Currency") + (Format(comparisons.SubItems(2), "Currency") * (ovrChk("AlgorithmPercentage") / 100))
                 
                 Else
                 
                    CurrPrice = Format(comparisons.SubItems(2), "Currency") + (Format(comparisons.SubItems(2), "Currency") * (dynCheck("DifferentialPercentage") / 100))
                 End If
                 If CurrPrice > NewMin And CurrPrice < NewMax Then
                    'comparisons.SubItems(3) = Format(CurrPrice, "Currency")
                 Else
                    If CurrPrice <= NewMin Then
                        'comparisons.SubItems(3) = Format(NewMin, "Currency")
                        CurrPrice = NewMin
                    Else
                        'comparisons.SubItems(3) = Format(NewMax, "Currrency")
                        CurrPrice = NewMax
                    End If
                 End If
            End If
       End If

 Next
 
 If compFlag = False Then
    Dim origData As ADODB.Recordset
    Set origData = DB.retrieve("SELECT OriginalPrice1 FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPWeb'")
    If origData.RecordCount > 0 Then
        DB.Execute "UPDATE InventoryMaster SET IDiscountMarkupPriceRate1=" & Format(origData("OriginalPrice1"), "Currency") & " WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
        
    Else
        
    End If
    
    
 End If
 
 If algFlag = False Then
    For Each comparisons In Me.InternetComparisonLVI.ListItems
        'comparisons.SubItems(3) = "N/A"
        comparisons.Checked = False
    Next
 
 End If
 
 Dim compareFlag As Boolean
 For Each comparisons In Me.InternetComparisonLVI.ListItems
    If comparisons.Checked = True Then
        compareFlag = True
        If algFlag = True Then
        OriginalIDiscountPrice = Me.IDiscountMarkupPriceRate(1)
        Me.IDiscountMarkupPriceRate(1) = Format(CurrPrice, "Currency")
        DB.Execute "UPDATE InventoryMaster SET IDiscountMarkupPriceRate1=" & Format(CurrPrice, "Currency") & " WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
        DB.Execute "UPDATE PartNumbers SET WebToBePublished=1 WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
        'MsgBox "TAG: " & comparisons.tag
        Dim competitorDataID As String
        Dim AlgorithmID As String
        Dim algorithmPercentage As String
        Dim priceFieldToChange As String
        Dim priceOne As Currency
        competitorDataID = CStr(comparisons.tag)
        Dim algList As ListItem
        For Each algList In Me.InternetDynamicAlgorithmLVI.ListItems
            If algList.Checked = True Then
                AlgorithmID = CStr(algList.tag)
                
                Dim algPerc As ADODB.Recordset
                Set algPerc = DB.retrieve("SELECT AlgorithmPercentage FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TBWeb'")
                If algPerc.RecordCount > 0 Then
                    'use the set percentage..
                    algorithmPercentage = CStr(Format(algPerc("AlgorithmPercentage"), "Currency"))
                Else
                    'use the default
                    algorithmPercentage = CStr(algList.SubItems(2))
                End If
            End If
        Next
        priceFieldToChange = "TPWeb"
        priceOne = Format(Me.IDiscountMarkupPriceRate(1).Text, "Currency")
        If InStr(algorithmPercentage, "%") > 0 Then algorithmPercentage = Replace(algorithmPercentage, "%", "")
        Dim existsOrNot As ADODB.Recordset
        Set existsOrNot = DB.retrieve("SELECT ID FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPWeb'")
        If existsOrNot.RecordCount > 0 Then
            Dim findAlgID As ADODB.Recordset
            Set findAlgID = DB.retrieve("SELECT ID,DifferentialPercentage FROM DynamicPricingAlgorithmHeader WHERE AlgorithmName='" & Split(Me.InternetAlgorithmCmb.list(Me.InternetAlgorithmCmb.ListIndex), ":")(0) & "'")
            If (findAlgID.RecordCount > 0) Then
                AlgorithmID = findAlgID("ID")
                algorithmPercentage = findAlgID("DifferentialPercentage")
            End If
            
            DB.Execute "UPDATE DynamicPricing SET Enabled=1,AlgorithmID=" & AlgorithmID & ", CompetitorDataID=" & competitorDataID & ",AlgorithmPercentage=" & algorithmPercentage & " WHERE ID=" & existsOrNot("ID")
            
        Else
            DB.Execute "INSERT INTO DynamicPricing (ItemNumber,Enabled,AlgorithmID,CompetitorDataID,AlgorithmPercentage,PriceFieldToChange,OriginalPrice1) VALUES('" & Me.ItemNumber.Text & "',1," & AlgorithmID & "," & competitorDataID & "," & algorithmPercentage & ",'TPWeb'," & OriginalIDiscountPrice & ")"
            
        End If
        
        
        competitorDataID = CStr(comparisons.tag)
        
        changed = True
        End If
        
    End If
 Next
 If compareFlag = False Then
    Me.IDiscountMarkupPriceRate(1) = OriginalIDiscountPrice
    DB.Execute "UPDATE PartNumbers SET WebToBePublished=1 WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
    changed = False
 End If
 End If
End Sub




Private Sub InternetDynamicAlgorithmLVI_ItemCheck(ByVal item As MSComctlLib.ListItem)

Dim hasACheck As Boolean
Dim checkCheck As ListItem
For Each checkCheck In InternetDynamicAlgorithmLVI.ListItems
    If checkCheck.Checked = True Then hasACheck = True
Next
If hasACheck = False Then
    DB.Execute "UPDATE DynamicPricing SET AlgorithmID=0,Enabled=0,CompetitorDataID=0 WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPWeb'"
    Dim price As ADODB.Recordset
Set price = DB.retrieve("SELECT OriginalPrice1 FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPWeb'")
If price.RecordCount > 0 Then
    Me.IDiscountMarkupPriceRate(1).Text = Format(price("OriginalPrice1"), "Currency")
    DB.Execute "UPDATE PartNumbers SET WebToBePublished=1 WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
    changed = True
End If
End If
    
    





Dim itm As ListItem

For Each itm In InternetDynamicAlgorithmLVI.ListItems
    If itm.SubItems(1) <> item.SubItems(1) Then
        itm.Checked = False
    Else
        'DB.Execute "UPDATE DynamicPricing SET AlgorithmID=" & itm.tag & ",Enabled=1 WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPWeb'"
    End If
Next
'CheckForInetDynamicPrice
InetDynamicPriceCheck

End Sub

Private Sub InetDynamicPriceCheck()
'Dim inetItems As String
'For Each inetItems In Me.InternetAlgorithmCmb.list
'Next

Dim algorithmName As String
algorithmName = Split(Me.InternetAlgorithmCmb.list(Me.InternetAlgorithmCmb.ListIndex), ":")(0)

Dim algInfo As ADODB.Recordset
Set algInfo = DB.retrieve("SELECT ID,DifferentialPercentage FROM DynamicPricingAlgorithmHeader WHERE AlgorithmName='" & algorithmName & "'")

If (algInfo.RecordCount > 0) Then
    Dim AlgorithmID As String
    Dim DifferentialPercent As Currency
    
    AlgorithmID = algInfo("ID")
    DifferentialPercent = Format(algInfo("DifferentialPercentage"), "Currency")
    
    Dim dynInfo As ADODB.Recordset
    Set dynInfo = DB.retrieve("SELECT ID FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPWeb'")
    If dynInfo.RecordCount > 0 Then
        DB.Execute "UPDATE DynamicPricing SET AlgorithmID=" & AlgorithmID & ", AlgorithmPercentage=" & DifferentialPercent & " WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPWeb'"
        
    Else
        DB.Execute "INSERT INTO DynamicPricing (ItemNumber,Enabled,AlgorithmID,CompetitorDataID,AlgorithmPercentage,PriceFieldToChange,OriginalPrice1,AlgorithmPercentageOverride) VALUES('" & Me.ItemNumber.Text & "',0," & AlgorithmID & ",0," & DifferentialPercent & ",'TPWeb'," & Me.IDiscountMarkupPriceRate(1).Text & ",0)"
    End If
    

    
    
    Dim comparisons As ListItem
    For Each comparisons In Me.InternetComparisonLVI.ListItems
        Dim NewMin As Currency
        Dim NewMax As Currency
        Dim CurrPrice As Currency
        Dim dynCheck As ADODB.Recordset
        Set dynCheck = DB.retrieve("DECLARE @input DECIMAL(7,3) SET @input=" & Format(comparisons.SubItems(2), "Currency") & " SELECT TOP 1 DynamicPricingAlgorithmLines.ID,HeaderID,DifferentialPercentage,DollarAmount,MaxPercentAbove,MaxPercentBelow FROM DynamicPricingAlgorithmLines INNER JOIN DynamicPricingAlgorithmHeader ON DynamicPricingAlgorithmHeader.ID=HeaderID WHERE HeaderID=" & AlgorithmID & " ORDER BY ABS(@input - DollarAmount)")
        
        If dynCheck.RecordCount > 0 Then
                Dim ovrChk As ADODB.Recordset
                Set ovrChk = DB.retrieve("SELECT AlgorithmPercentage FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPWeb' AND AlgorithmPercentageOverride=1")
                NewMin = Format(comparisons.SubItems(2), "Currency") - (Format(comparisons.SubItems(2), "Currency") * (dynCheck("MaxPercentBelow") / 100))
                NewMax = Format(comparisons.SubItems(2), "Currency") + (Format(comparisons.SubItems(2), "Currency") * (dynCheck("MaxPercentAbove") / 100))
                If ovrChk.RecordCount > 0 Then
                CurrPrice = Format(comparisons.SubItems(2), "Currency") + (Format(comparisons.SubItems(2), "Currency") * (ovrChk("AlgorithmPercentage") / 100))
                Else
                CurrPrice = Format(comparisons.SubItems(2), "Currency") + (Format(comparisons.SubItems(2), "Currency") * (dynCheck("DifferentialPercentage") / 100))
                End If
               
             If CurrPrice > NewMin And CurrPrice < NewMax Then
                'comparisons.SubItems(3) = CStr(Format(CurrPrice, "Currency"))
             Else
                If CurrPrice <= NewMin Then
                    'comparisons.SubItems(3) = "$" & CStr(CDbl(Format(NewMin, "Currency")))
                    CurrPrice = NewMin
                Else
                    'comparisons.SubItems(3) = "$" & CStr(CDbl(Format(NewMax, "Currency")))
                    CurrPrice = NewMax
                End If
             End If
        
        End If
    Next
    
End If



End Sub

Private Sub InetDynamicPriceCheck2()
Dim algFlag As Boolean
 Dim comparisons As ListItem
 Dim itm As ListItem
 
 Dim shouldProceed As Boolean
 Dim testItm As ListItem
 For Each testItm In Me.InternetComparisonLVI.ListItems
    If testItm.Checked = True Then shouldProceed = True
 Next
 
 
 
 For Each comparisons In Me.InternetComparisonLVI.ListItems
    For Each itm In InternetDynamicAlgorithmLVI.ListItems
       If itm.Checked = True Then
           algFlag = True
           Dim dynCheck As ADODB.Recordset
           Set dynCheck = DB.retrieve("DECLARE @input DECIMAL(7,3) SET @input=" & Format(comparisons.SubItems(2), "Currency") & " SELECT TOP 1 DynamicPricingAlgorithmLines.ID,HeaderID,DifferentialPercentage,DollarAmount,MaxPercentAbove,MaxPercentBelow FROM DynamicPricingAlgorithmLines INNER JOIN DynamicPricingAlgorithmHeader ON DynamicPricingAlgorithmHeader.ID=HeaderID WHERE HeaderID=" & itm.tag & " ORDER BY ABS(@input - DollarAmount)")
            If dynCheck.RecordCount > 0 Then
                 'NewPrice = Format(CompetitorPrice, "Currency") + (Format(CompetitorPrice, "Currency") * (DynPrice("DifferentialPercentage") / 100))
                 Dim NewMin As Currency
                 Dim NewMax As Currency
                 Dim CurrPrice As Currency
                 
                 'need to check if the defaults here get overridden... basically if so, get values there, else use the next 3 lines down...\
                    
                    Dim ovrChk As ADODB.Recordset
                    Set ovrChk = DB.retrieve("SELECT AlgorithmPercentage FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPWeb' AND AlgorithmPercentageOverride=1")
                    NewMin = Format(comparisons.SubItems(2), "Currency") - (Format(comparisons.SubItems(2), "Currency") * (dynCheck("MaxPercentBelow") / 100))
                    NewMax = Format(comparisons.SubItems(2), "Currency") + (Format(comparisons.SubItems(2), "Currency") * (dynCheck("MaxPercentAbove") / 100))
                    If ovrChk.RecordCount > 0 Then
                    CurrPrice = Format(comparisons.SubItems(2), "Currency") + (Format(comparisons.SubItems(2), "Currency") * (ovrChk("AlgorithmPercentage") / 100))
                    Else
                        
                    CurrPrice = Format(comparisons.SubItems(2), "Currency") + (Format(comparisons.SubItems(2), "Currency") * (dynCheck("DifferentialPercentage") / 100))
                
                    End If
                
                   
                 If CurrPrice > NewMin And CurrPrice < NewMax Then
                    'comparisons.SubItems(3) = Format(CurrPrice, "Currency")
                 Else
                    If CurrPrice <= NewMin Then
                        'comparisons.SubItems(3) = Format(NewMin, "Currency")
                        CurrPrice = NewMin
                    Else
                        'comparisons.SubItems(3) = Format(NewMax, "Currrency")
                        CurrPrice = NewMax
                    End If
                 End If
            End If
       End If
    Next
 Next

 If algFlag = False Then
    For Each comparisons In Me.InternetComparisonLVI.ListItems
        'comparisons.SubItems(3) = "N/A"
        comparisons.Checked = False
        Dim origPrice As ADODB.Recordset
        Set origPrice = DB.retrieve("SELECT OriginalPrice1 FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
        If origPrice.RecordCount > 0 And shouldProceed = True Then
            Me.IDiscountMarkupPriceRate(1).Text = Format(origPrice("OriginalPrice1"), "Currency")
            DB.Execute "UPDATE InventoryMaster SET IDiscountMarkupPriceRate1=" & Format(Me.IDiscountMarkupPriceRate(1).Text, "Currency") & " WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
            
        End If
    Next
 Else
    If Format(Me.IDiscountMarkupPriceRate(1).Text, "Currency") <> Format(CurrPrice, "Currency") And shouldProceed = True Then
        Me.IDiscountMarkupPriceRate(1).Text = Format(CurrPrice, "Currency")
        DB.Execute "UPDATE InventoryMaster SET IDiscountMarkupPriceRate1=" & Format(Me.IDiscountMarkupPriceRate(1).Text, "Currency") & " WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
    End If
 End If

 
End Sub


Private Sub InventoriedDate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.popupTmr.Enabled = False
'Me.HidePopUp = True
'Me.HidePopUpFrm
Sleep 300
InventoriedDate.ZOrder 0
InventoriedDate.SetFocus
InventoriedDate.Refresh
End Sub

Private Sub ItemNumber_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.popupTmr.Enabled = False
'Me.HidePopUp = True
'Me.HidePopUpFrm
Sleep 300
ItemNumber.ZOrder 0
ItemNumber.SetFocus
ItemNumber.Refresh
End Sub

Private Sub ItemNumber_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'ItemNumber.BackColor = RGB(0, 255, 255)
            ItemNumber.Refresh
    
            'If Me.HidePopUp = False Then
                'PopUpFrm.Visible = True
                'Me.popupTmr.Enabled = True
                'PopUpFrm.InfoLbl.Caption = Me.currentItemNumber
                'PopUpFrm.ZOrder 0
                'PopUpFrm.SetFocus
                'PopUpFrm.Move Me.Left - 1250, Me.Top + 50 + ItemNumber.Top   'ScaleX(NewCost.Left, vbPixels, vbTwips), ScaleY(NewCost.Top, vbPixels, vbTwips) ', PopUpFrm.width, PopUpFrm.Height
                'PopUpFrm.Refresh
            'End If
End Sub

Private Sub JetPrice_KeyDown(KeyCode As Integer, Shift As Integer)
'changed = True
End Sub

Private Sub kitCreatorShortcut10_Click()
Dim getDesc As ADODB.Recordset
Set getDesc = DB.retrieve("SELECT (ManufFullNameWeb + ' ' + COALESCE(Desc2,'') + ' ' + COALESCE(Desc1,'') + ' ' + COALESCE(Desc3,'')) AS Description FROM ProductLine,PartNumbers WHERE ProductLine='" & Me.prodLine.Text & "' AND ItemNumber='" & Me.ItemNumber.Text & "'")
If getDesc.RecordCount > 0 Then
    KitCreator.NewKitList.Add (Array(Me.ItemNumber, "10", IIf(IsNull(getDesc("Description")) = True, "", getDesc("Description"))))
Else
    KitCreator.NewKitList.Add (Array(Me.ItemNumber, "10", ""))
End If
KitCreator.btnSaveKit_Click
btnPrevItem_Click
End Sub

Private Sub kitCreatorShortcut12_Click()
Dim getDesc As ADODB.Recordset
Set getDesc = DB.retrieve("SELECT (ManufFullNameWeb + ' ' + COALESCE(Desc2,'') + ' ' + COALESCE(Desc1,'') + ' ' + COALESCE(Desc3,'')) AS Description FROM ProductLine,PartNumbers WHERE ProductLine='" & Me.prodLine.Text & "' AND ItemNumber='" & Me.ItemNumber.Text & "'")
If getDesc.RecordCount > 0 Then
    KitCreator.NewKitList.Add (Array(Me.ItemNumber, "12", IIf(IsNull(getDesc("Description")) = True, "", getDesc("Description"))))
Else
    KitCreator.NewKitList.Add (Array(Me.ItemNumber, "12", ""))
End If
KitCreator.btnSaveKit_Click
btnPrevItem_Click
End Sub

Private Sub kitCreatorShortcut2_Click()
Dim getDesc As ADODB.Recordset
Set getDesc = DB.retrieve("SELECT (ManufFullNameWeb + ' ' + COALESCE(Desc2,'') + ' ' + COALESCE(Desc1,'') + ' ' + COALESCE(Desc3,'')) AS Description FROM ProductLine,PartNumbers WHERE ProductLine='" & Me.prodLine.Text & "' AND ItemNumber='" & Me.ItemNumber.Text & "'")
If getDesc.RecordCount > 0 Then
    KitCreator.NewKitList.Add (Array(Me.ItemNumber, "2", IIf(IsNull(getDesc("Description")) = True, "", getDesc("Description"))))
Else
    KitCreator.NewKitList.Add (Array(Me.ItemNumber, "2", ""))
End If

KitCreator.btnSaveKit_Click

btnPrevItem_Click

End Sub

Private Sub kitCreatorShortcut3_Click()
Dim getDesc As ADODB.Recordset
Set getDesc = DB.retrieve("SELECT (ManufFullNameWeb + ' ' + COALESCE(Desc2,'') + ' ' + COALESCE(Desc1,'') + ' ' + COALESCE(Desc3,'')) AS Description FROM ProductLine,PartNumbers WHERE ProductLine='" & Me.prodLine.Text & "' AND ItemNumber='" & Me.ItemNumber.Text & "'")
If getDesc.RecordCount > 0 Then
    KitCreator.NewKitList.Add (Array(Me.ItemNumber, "3", IIf(IsNull(getDesc("Description")) = True, "", getDesc("Description"))))
Else
    KitCreator.NewKitList.Add (Array(Me.ItemNumber, "3", ""))
End If
KitCreator.btnSaveKit_Click
btnPrevItem_Click
End Sub

Private Sub kitCreatorShortcut5_Click()
Dim getDesc As ADODB.Recordset
Set getDesc = DB.retrieve("SELECT (ManufFullNameWeb + ' ' + COALESCE(Desc2,'') + ' ' + COALESCE(Desc1,'') + ' ' + COALESCE(Desc3,'')) AS Description FROM ProductLine,PartNumbers WHERE ProductLine='" & Me.prodLine.Text & "' AND ItemNumber='" & Me.ItemNumber.Text & "'")
If getDesc.RecordCount > 0 Then
    KitCreator.NewKitList.Add (Array(Me.ItemNumber, "5", IIf(IsNull(getDesc("Description")) = True, "", getDesc("Description"))))
Else
    KitCreator.NewKitList.Add (Array(Me.ItemNumber, "5", ""))
End If
KitCreator.btnSaveKit_Click
btnPrevItem_Click
End Sub

Private Sub kitCreatorShortcut6_Click()
Dim getDesc As ADODB.Recordset
Set getDesc = DB.retrieve("SELECT (ManufFullNameWeb + ' ' + COALESCE(Desc2,'') + ' ' + COALESCE(Desc1,'') + ' ' + COALESCE(Desc3,'')) AS Description FROM ProductLine,PartNumbers WHERE ProductLine='" & Me.prodLine.Text & "' AND ItemNumber='" & Me.ItemNumber.Text & "'")
If getDesc.RecordCount > 0 Then
   KitCreator.NewKitList.Add (Array(Me.ItemNumber, "6", IIf(IsNull(getDesc("Description")) = True, "", getDesc("Description"))))
Else
    KitCreator.NewKitList.Add (Array(Me.ItemNumber, "6", ""))
End If
KitCreator.btnSaveKit_Click
btnPrevItem_Click
End Sub

Private Sub lblIsKit_DblClick()
    Load InventoryKitPopup
    If Me.lblIsKit.caption = "KIT" Then
        InventoryKitPopup.LoadKit Me.ItemNumber
    Else
        InventoryKitPopup.LoadComponent Me.ItemNumber
    End If
    InventoryKitPopup.Show



End Sub

Private Sub lineTmr_Timer()
    InventoryMaintenance.lineDoubleEnterFix = False
    InventoryMaintenance.lineTmr.Enabled = False
End Sub

Private Sub ListPrice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.HidePopUp = True
Me.HidePopUpFrm
Sleep 300
ListPrice.ZOrder 0
ListPrice.SetFocus
ListPrice.Refresh
End Sub

Private Sub MAPP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.HidePopUp = True
Me.HidePopUpFrm
Sleep 300
MAPP.ZOrder 0
MAPP.SetFocus
MAPP.Refresh
End Sub

Private Sub NewCost_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.HidePopUp = True
Me.HidePopUpFrm
Sleep 300
NewCost.ZOrder 0
NewCost.SetFocus
NewCost.Refresh
End Sub

Private Sub opDropshipCostType_Click(index As Integer)
    Dim i As Long
    For i = Me.opDropshipCostType.LBound To Me.opDropshipCostType.UBound
        Me.opDropshipCostType(i).ForeColor = COLOR_BUTTON_TEXT
    Next i
    Select Case index
        Case Is = 0
            Me.DropshipCostLabel = "D/S Cost:"
            Me.opDropshipCostType(index).ForeColor = BLUE
        Case Is = 1
            Me.DropshipCostLabel = "C+D/S+12%:"
            Me.opDropshipCostType(index).ForeColor = YELLOW
        Case Is = 2
            Me.DropshipCostLabel = "EBay $:"
            Me.opDropshipCostType(index).ForeColor = RED
        Case Else
            Err.Raise "123"
    End Select
    fillDropshipCostField
End Sub

Private Sub opShippingType_Click(index As Integer)
    If Not fillingForm Then
        'Select Case True
        '    Case Is = Me.opShippingType(SHIP_REG)
        '        updateInventoryMaster "TruckFreight", "0", Me.ItemNumber
        '        updateInventoryMaster "TruckFreightCheap", "0", Me.ItemNumber
        '    Case Is = Me.opShippingType(SHIP_TF)
        '        updateInventoryMaster "TruckFreight", "1", Me.ItemNumber
        '        updateInventoryMaster "TruckFreightCheap", "0", Me.ItemNumber
        '    Case Is = Me.opShippingType(SHIP_TF_CHEAP)
        '        updateInventoryMaster "TruckFreight", "1", Me.ItemNumber
        '        updateInventoryMaster "TruckFreightCheap", "1", Me.ItemNumber
        'End Select
        updateInventoryMaster "ShippingType", CLng(index), Me.ItemNumber
        Me.opShippingType(3).ForeColor = IIf(Me.opShippingType(3) = True, RED, BLACK)
        'updateExternalShippingDB Me.ItemNumber
        'ExternalComponentDBSync Me.ItemNumber
        'ExternalItemDBSync Me.ItemNumber
    End If
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
        'If Me.barPosition.value = 0 Then
        '    barPosition_Change
        'Else
        '    Me.barPosition.value = 0
        'End If
        loadItemByFilterIndex 0
        
        'signs can filter inventory, but inventory can't filter signs
        'since signs will be missing xxx/zzz records from the IM filters
        'If IsFormLoaded("SignMaintenance") Then
        '    If vbYes = MsgBox("Filter Signs too?", vbYesNo) Then
        '        SignMaintenance.SetFilter Me.cmbFilters
        '    End If
        'End If
        Me.jumpToLine.SetFocus
        
        
        Me.ClearCurrentVars

If InventoryQuantityTriggersV2.Visible = True Then
Unload InventoryQuantityTriggersV2
    If InventoryQuantityTriggersV2.Visible = True Then
        'do nothing, there was something to save...
        Exit Sub
    Else
        btnViewInvQtyTriggers_Click
    End If


End If
        
    End If
End Sub

Private Sub cmbWhse_Click()
    cmbWhse_LostFocus
End Sub

Private Sub cmbWhse_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.cmbWhse, KeyCode, Shift
End Sub

Private Sub cmbWhse_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.cmbWhse, KeyAscii
    If KeyAscii = vbKeyReturn Then
        cmbWhse_LostFocus
    End If
End Sub

Private Sub cmbWhse_LostFocus()
    AutoCompleteLostFocus Me.cmbWhse
    fillSalesHistory Me.ItemNumber, IsItemDSable(Me.ItemNumber)
End Sub

'Private Sub CurrentOP_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case Is = vbKeyReturn
'            CurrentOP_LostFocus
'        Case Is = vbKeyDelete
'            CurrentOP_KeyPress KeyCode
'    End Select
'End Sub
'
'Private Sub CurrentOP_KeyPress(KeyAscii As Integer)
'    changed = True
'    whichCtl = "CurrentOP"
'End Sub
'
'Private Sub CurrentOP_LostFocus()
'    If changed = True Then
'        DB.Execute ("UPDATE InventoryMaster SET NewReorderPoint=" & Me.CurrentOP & ", OldReorderPoint=" & Me.CurrentOP & ", ReorderPoint=" & Me.CurrentOP & " WHERE ItemNumber='" & Me.ItemNumber & "'")
'        Me.NewOP.Visible = True
'        Me.lblNewOP.Visible = True
'        Me.chkUseNewOP.Visible = True
'        changed = False
'    End If
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'priceCompareEnabled = False
    If KeyCode = vbKeyF1 Then
        Me.jumpToLine.SetFocus
    ElseIf KeyCode = vbKeyF2 Then
        Me.jumpToItem.SetFocus
    'ElseIf (KeyCode = vbKeyPageUp Or KeyCode = vbKeyF8) And Me.barPosition <> Me.barPosition.Min Then
    ElseIf (KeyCode = vbKeyPageUp Or KeyCode = vbKeyF8) And currentFilterPosition <> 0 Then
    Me.SetFocus
        'Me.barPosition = Me.barPosition - 1
        loadItemByFilterIndex currentFilterPosition - 1
        If TypeOf Me.ActiveControl Is TextBox Then
            HandleTextBoxCursorPosition Me.ActiveControl
        ElseIf TypeOf Me.ActiveControl Is ComboBox Then
            HandleComboBoxCursorPosition Me.ActiveControl
        End If
        KeyCode = 0
        
        If InventoryQuantityTriggersV2.Visible = True Then
        Unload InventoryQuantityTriggersV2
        If InventoryQuantityTriggersV2.Visible = True Then
            Exit Sub
        Else
            btnViewInvQtyTriggers_Click
        End If
        End If
        
    'ElseIf (KeyCode = vbKeyPageDown Or KeyCode = vbKeyF9) And Me.barPosition <> Me.barPosition.max Then
    ElseIf (KeyCode = vbKeyPageDown Or KeyCode = vbKeyF9) And currentFilterPosition <> UBound(currentFilterItemList) Then
    Me.SetFocus
        'Me.barPosition = Me.barPosition + 1
        loadItemByFilterIndex currentFilterPosition + 1
        If TypeOf Me.ActiveControl Is TextBox Then
            HandleTextBoxCursorPosition Me.ActiveControl
        ElseIf TypeOf Me.ActiveControl Is ComboBox Then
            HandleComboBoxCursorPosition Me.ActiveControl
        End If
        KeyCode = 0
        If InventoryQuantityTriggersV2.Visible = True Then
        Unload InventoryQuantityTriggersV2
        If InventoryQuantityTriggersV2.Visible = True Then
            Exit Sub
        Else
            btnViewInvQtyTriggers_Click
        End If
        End If
    ElseIf KeyCode = vbKeyF4 Then
        btnAddNewItem_Click
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()


    
    'If USE_ALT_DATABASE Then
    '    Me.BackColor = &H80&
    '    Me.btnFreightEstimate.Enabled = False
    '    Me.btnPriceCompare.Enabled = False
    'End If
    setupInternetComparisonLVIHeaders
    setupEBayComparisonsLVIHeaders
    
    inetSortPriceBtn_Click
    ebaySortPriceBtn_Click
    SortNumeric
    'Me.inetSortTotalBtn_Click
    'Me.inetSortTotalBtn_Click
    'Me.ebaySortTotalBtn_Click
    'Me.ebaySortTotalBtn_Click
    
    
    Mouse.Hourglass True
    
    Me.textBoxBackColor = Me.BackColor
    
    Set IMtt = New FormTooltips
    
    Me.FrameOPLS.BorderStyle = 0
    
    If Not HasPermissionsTo("InventoryMaintenanceWrite") Then
        lockedDown = True
        lockForm
    End If
    If Not HasPermissionsTo("SignMaintenanceRead") Then
        Me.btnGoToSigns.Enabled = False
    End If
    If Not HasPermissionsTo("PriceComparison") Then
        lockedDownPC = True
        '12/12/05 - access to single item compare for everyone
        'Me.btnPriceCompare.Enabled = False
    End If
    If Not HasPermissionsTo("FreightEstimation") Then
        lockedDownFE = True
        Me.btnFreightEstimate.Enabled = False
    End If
    If Not HasPermissionsTo("WeightDims") Then
        lockedDownWD = True
        Me.shipInfoFrame.Enabled = False
        Me.components.Locked = True
    End If
    'If HasPermissionsTo("RecordAnnouncements") Then
    '    Me.btnOpenSoundRec.Visible = True
    'End If
    If Not HasPermissionsTo("Barcode") Then
        'Load barcodes can be view-only, so ok to open
        'Me.btnLoadBarcodes.Enabled = False
        Me.btnNewBarcode.Enabled = False
    End If
    
    If Not lockedDown Then
        setAlertsButton
    End If
    
    If WARN_ON_IMPORT Then
        'If IMPORT_GOING Then
        '    LockForImport True
        'End If
        If EXPORT_GOING Then
            LockForExport True
        End If
    End If
    
    ReDim xsheetarray(10) As String
    requeryVendor
    requeryWhse
    requeryFilters
    requeryJumpToLine
'    requeryClass
    
    'Me.cmbFilters = "All"
    'requeryForm "All"
    If IsFormLoaded("SignMaintenance") Then
        Dim signfilter As String
        signfilter = SignMaintenance.GetFilter
        If Left(signfilter, 4) = "FBF-" Then
            Me.btnFilterByForm.tag = Mid(signfilter, 5)
            requeryForm "FBF"
'        ElseIf Left(signfilter, 4) = "FBS-" Then
'            Me.btnSearch.tag = Mid(signfilter, 5)
'            requeryForm "FBS"
        Else
            requeryForm "All"
            'requeryForm "Ex Disc QOH=0 + OLD"
        End If
    Else
        requeryForm "All"
        'requeryForm "Ex Disc QOH=0 + OLD"
    End If

    SetListTabs Me.freightEstimates, Array(36, 36)
    SetListTabs Me.freightActual, Array(36, 84) '120=off screen
    SetListTabs Me.POList, Array(60, 60) '120=off screen
    SetListTabs Me.priceComparison, Array(108, 156, 186, 300), False '300=off screen
    
    ExpandDropDownToFit Me.PrimaryVendor
    ExpandDropDownToFit Me.cmbFilters

    firstLoad = True
    'Me.barPosition.value = 0
    'barPosition_Change
    loadItemByFilterIndex 0
    firstLoad = False
    Mouse.Hourglass False
    
    
     'Tom's Code 9-8-2014
    'this code is to check security permissions for subtitle overriding...
    If HasPermissionsTo("SubtitleOverride") = True Then
        'idk guess nothing..
    Else
    
        SubTitleChk.Enabled = False
        SubtitleCmb.Enabled = False
        DeleteSelectedSubtitleBtn.Enabled = False
        AddSelectedSubtitleBtn.Enabled = False
    End If
    
    If HasPermissionsTo("JetDotCom") = True Then
        Me.JetChk.Enabled = True
    Else
        Me.JetChk.Enabled = False
    End If
    
    If HasPermissionsTo("DynamicPricing") = True Then
        Me.dynamicPriceChk.Enabled = True
        
    Else
        Me.dynamicPriceChk.Enabled = False
    End If
    
    
    
    Me.cmbFilters.SelLength = 0 'something's making this selected even when focus isn't there...
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload InventoryQuantityTriggersV2
     Unload PopUpFrm
    
    Dim found As Boolean
    Dim i As Long
    For i = 0 To UBound(xsheetarray)
        If xsheetarray(i) <> "" Then
            found = True
            i = UBound(xsheetarray)
        End If
    Next i
    If found Then
        If MsgBox("You have items in the ""Generate New Products XSheet"" array, close anyway?", vbYesNo) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    
    Set IMtt = Nothing
End Sub

Private Sub freightActual_DblClick()
    If Me.freightActual.ListIndex > 0 Then
        Dim rst As ADODB.Recordset
        'Set rst = DB.retrieve("exec spFreightActualDblClick '" & Me.ItemNumber & "', '" & Mid(Me.freightActual, InStrRev(Me.freightActual, vbTab) + 1) & "'")
        Set rst = DB.retrieve("SELECT F.SalesOrderNumber, F.DateOfSale, F.DateOfShipment, F.ItemUnitPrice, L.City, L.State, L.Country " & _
                              "FROM FreightActual AS F LEFT OUTER JOIN LookupCityZipCode AS L ON F.ZipCode=L.ZipCode " & _
                              "WHERE F.ItemNumber='" & Me.ItemNumber & "' " & _
                              "  AND F.SalesOrderNumber='" & Utilities.ListBoxLineColumn(Me.freightActual, Me.freightActual.ListIndex, 2) & "'")
        If rst("City") = "" Then
            MsgBox "Item Unit Price: " & Format(rst("ItemUnitPrice"), "Currency") & vbCrLf & "Sales Order #" & rst("SalesOrderNumber") & vbCrLf & "Date of Sale: " & rst("DateOfSale") & vbCrLf & "Date of Shipment: " & rst("DateOfShipment") & vbCrLf & "Location: not found in lookup table."
        Else
            MsgBox "Item Unit Price: " & Format(rst("ItemUnitPrice"), "Currency") & vbCrLf & "Sales Order #" & rst("SalesOrderNumber") & vbCrLf & "Date of Sale: " & rst("DateOfSale") & vbCrLf & "Date of Shipment: " & rst("DateOfShipment") & vbCrLf & "Location: " & rst("City") & ", " & rst("State") & " " & rst("Country")
        End If
        rst.Close
        Set rst = Nothing
    End If
End Sub

Private Sub freightEstimates_DblClick()
    If Me.freightEstimates.ListIndex > 0 Then
        Dim rst As ADODB.Recordset
        'Set rst = DB.retrieve("exec spFreightEstimatesDblClick '" & Me.ItemNumber & "', '" & freightA(0) & "'")
        Set rst = DB.retrieve("SELECT L.City, L.State, L.Country, F.LastUpdated " & _
                              "FROM FreightEstimates AS F LEFT OUTER JOIN LookupCityZipCode AS L ON F.ZipCode=L.ZipCode " & _
                              "WHERE F.ItemNumber='" & Me.ItemNumber & "' " & _
                              "  AND F.ZipCode='" & Utilities.ListBoxLineColumn(Me.freightEstimates, Me.freightEstimates.ListIndex, 0) & "'")
        Dim msg As String
        If rst.EOF Then
            msg = "Can't find row, that's weird"
        Else
            If Nz(rst("Country")) = "" Then
                msg = "Location: unknown"
            Else
                msg = "Location: " & Nz(rst("City")) & ", " & Nz(rst("State")) & " " & Nz(rst("Country"))
            End If
            msg = msg & vbCrLf & vbCrLf
            msg = msg & "Queried on: " & rst("LastUpdated")
            MsgBox msg
        End If
        rst.Close
        Set rst = Nothing
    End If
End Sub

Private Sub jumpToLine_Click()
    changed = True
    jumpToLine_LostFocus
End Sub

Private Sub jumpToLine_KeyDown(KeyCode As Integer, Shift As Integer)
If Me.lineDoubleEnterFix = True Then Exit Sub
If Me.lineTmr.Enabled = True Then Exit Sub
    
    AutoCompleteKeyDown Me.jumpToLine, KeyCode, Shift
    If Shift = vbCtrlMask And KeyCode = vbKeyV Then
        Dim item As String
        item = ClipBoard_GetData
        item = Replace(item, vbCrLf, "")
        item = Replace(item, " ", "")
        If ExistsInInventoryMaster(item) Then
            stopLineLostFocus = True
            Me.jumpToLine = Left(item, 3)
            requeryJumpToItem
            Me.jumpToItem = Mid(item, 4)
            Me.jumpToItem.SetFocus
            changed = True
            jumpToItem_LostFocus
        Else
            MsgBox "Unable to find item " & item & "!"
        End If
    End If
End Sub

Private Sub jumpToLine_KeyPress(KeyAscii As Integer)
If Me.lineDoubleEnterFix = True Then Exit Sub
If Me.lineTmr.Enabled = True Then Exit Sub
    
    AutoCompleteKeyPress Me.jumpToLine, KeyAscii
    If KeyAscii = vbKeyReturn Then
        Me.jumpToItem.SetFocus
    Else
        changed = True
    End If
End Sub

Private Sub jumpToLine_LostFocus()
    AutoCompleteLostFocus Me.jumpToLine
    If changed = True And stopLineLostFocus = False Then
        changed = False
        If Len(Me.jumpToLine) = 3 Then
            requeryJumpToItem
            Dim pos As Long
retry:
            pos = findFirstInLC(Me.jumpToLine, currentFilterItemList)
            If pos = -1 Then
                If MsgBox("No items for line code in current filter?" & vbCrLf & vbCrLf & "Remove filter and try again?", vbYesNo) = vbYes Then
                    Dim lc As String
                    lc = Me.jumpToLine
                    requeryForm "All"
                    Me.jumpToLine = lc
                    GoTo retry
                End If
            Else
                'Me.barPosition.value = pos
                loadItemByFilterIndex pos
            End If
            Me.jumpToItem.SetFocus
        ElseIf Len(Me.jumpToLine) = 0 Then
            Me.jumpToLine = Left(Me.ItemNumber, 3)
        End If
    ElseIf stopLineLostFocus = True Then
        stopLineLostFocus = False
        changed = False
    End If
End Sub

Private Sub jumpToItem_Click()
    changed = True
    jumpToItem_LostFocus
End Sub

Private Sub jumpToItem_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.jumpToItem, KeyCode, Shift
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
    AutoCompleteLostFocus Me.jumpToItem
    If changed = True Then
        changed = False
        If Len(Me.jumpToItem) = 0 Then
            Me.jumpToItem = Mid(Me.ItemNumber, 4)
        Else
            Dim pos As Long
retry:
            pos = bsearch(Me.jumpToLine & Me.jumpToItem, currentFilterItemList)
            If pos = -1 Then
                If MsgBox("Item is not in current filter?" & vbCrLf & vbCrLf & "Remove filter and try again?", vbYesNo) = vbYes Then
                    Dim lcitem As String
                    lcitem = Me.jumpToLine & Me.jumpToItem
                    requeryForm "All"
                    Me.jumpToLine = Left$(lcitem, 3)
                    Me.jumpToItem = Mid$(lcitem, 4)
                    GoTo retry
                End If
            Else
                'Me.barPosition.value = pos
                loadItemByFilterIndex pos
            End If
            Me.jumpToItem.selStart = 0
            Me.jumpToItem.SelLength = Len(Me.jumpToItem)
        End If
    End If
End Sub

Private Sub lblPOComment_DblClick(index As Integer)
    handleChangedField 'doesn't call the lostfocus from the prev. ctl, so manually do it
    If Not lockedDown Then
        'this is completely ridiculous, but it works. don't change anything
        changed = True 'may be changed back later
        Dim theDate As String
        theDate = Format(Date, "m/d/yy")
        Dim re As RegExp
        Set re = New RegExp
        'test for the basic format
        re.Pattern = "(.*?)\d+f/w\d+ \d{1,2}/\d{1,2}/\d{2}$"
        If Not re.test(Me.POComment) Then
            'if nothing's there, then just add the default 1f/w12
            If Me.POComment = "" Then
                Me.POComment = "1f/w12 " & theDate
            Else
                'if we have something there, see if we can fit it.
                If Len(Me.POComment & " 1f/w12 " & theDate) <= 50 Then
                    Me.POComment = Me.POComment & " 1f/w12 " & theDate
                Else
                    MsgBox "PO comment too long to add free items!"
                    changed = False
                End If
            End If
        Else
            'f/w tag already there, figure out what it is
            Dim matches As MatchCollection
            Set matches = re.Execute(Me.POComment)
            Dim keepstr As String
            'save the beginning of the string, if it's there
            If matches(0).SubMatches.count Then
                keepstr = matches(0).SubMatches(0)
            End If
            Dim comm As String
            comm = Replace(Me.POComment, keepstr, "")
            'if the date is wrong, then start from 1f/w12
            If Mid(comm, InStrRev(comm, " ") + 1) <> theDate Then
                comm = "1f/w12 " & theDate
            Else
                'otherwise split it apart and increment
                comm = Replace(comm, " " & theDate, "")
                Dim free As String, buy As String
                free = Left(comm, InStr(comm, "f/w") - 1)
                buy = Mid(comm, InStr(comm, "f/w") + 3)
                free = CStr(CInt(free) + 1)
                buy = CStr(CInt(buy) + 12)
                comm = free & "f/w" & buy & " " & theDate
            End If
            'now check for length after putting it back together
            If Len(keepstr & comm) <= 50 Then
                Me.POComment = keepstr & comm
            Else
                MsgBox "PO comment too long to add free items!"
                changed = False
            End If
        End If
        Set re = Nothing
        POComment_LostFocus
    End If
End Sub

Private Sub popupTmr_Timer()
'PopUpFrm.Visible = True
'Me.popupTmr.Enabled = False
End Sub

Private Sub priceCompareAsync_Timer()

End Sub

Private Sub priceCompareBtn_Click()

Dim pcHandle As Long
pcHandle = IsPriceComparerOpen

If pcHandle > 0 Then
    priceCompareEnabled = True
    
Else
    On Error GoTo cantRun
    Shell """\\TOOLSPLUS04\Databases\mastest\mas90-signs\A_Dist\toms beta\FollowMePricing.exe""", vbNormalFocus
    priceCompareEnabled = True

    'priceCompareBtn_Click
    Exit Sub
    On Error GoTo 0
cantRun:
    MsgBox "Cant load Price Comparer. Are you remoting in?"
    priceCompareEnabled = False
    Exit Sub
End If

PriceCompareCSharp.GetPriceComparerEBayHandle
PriceCompareCSharp.GetPriceComparerGoogleHandle
GetPriceComparerStatusHandles
priceCompareTmr.Enabled = True


End Sub
Private Sub priceCompare()

Dim pcHandle As Long
pcHandle = IsPriceComparerOpen

If pcHandle > 0 Then

'    priceCompareAsync.Enabled = True


'priceCompareAsync.Enabled = False
'first create search term... which is MFN cleaned + desc2
Dim qry As ADODB.Recordset
Set qry = DB.retrieve("SELECT (ProductLine.ManufFullNameWeb + ' ' + PartNumbers.Desc2) AS Product FROM InventoryMaster INNER JOIN ProductLine ON ProductLine.ProductLine = InventoryMaster.ProductLine INNER JOIN PartNumbers ON PartNumbers.ItemNumber= InventoryMaster.ItemNumber WHERE InventoryMaster.ItemNumber='" & Me.ItemNumber.Text & "'")
If qry.RecordCount > 0 Then
    Dim product As String
    product = qry("Product")
    
    If product = LastComparedItem Then
        Exit Sub
    End If
    setupEBayComparisonsLVIHeaders
    setupEBayDynamicAlgorithmLVIHeaders
    
    LastComparedItem = product
    PriceCompareCSharp.SetPriceComparerItem product, pcHandle
    

'    Dim Results As String
'    Results = ConsoleCommand.GetCommandOutput("""\\TOOLSPLUS04\Databases\mastest\mas90-signs\A_Dist\toms beta\RealTimePricesConsole.exe"" """ & Product & """")
'    PriceCompareFrm.priceBrowser.Navigate ("about:blank")
'    Do
'    DoEvents
'    Loop Until PriceCompareFrm.priceBrowser.ReadyState = READYSTATE_COMPLETE Or READYSTATE_LOADED
'    PriceCompareFrm.priceBrowser.Document.Write Results
'    If PriceCompareFrm.Visible = False Then PriceCompareFrm.Visible = True
'    priceComparisonData.GetCommandOutput ("""\\TOOLSPLUS04\Databases\mastest\mas90-signs\A_Dist\toms beta\RealTimePricesConsole.exe"" """ & Product & """")
End If

Else
'MsgBox "You must open the price comparer first... Too much code in too little timespan to program it all nicely and automatic like."
End If
End Sub


Private Sub priceCompareTmr_Timer()
If priceCompareEnabled = True And PriceCompareCSharp.EBayIEHandle > 0 And PriceCompareCSharp.EBayStatusHandle > 0 Then
    Dim EBayStatus As String
    Dim EBayResults As String
    EBayStatus = PriceCompareCSharp.GetObjectText(PriceCompareCSharp.EBayStatusHandle)
    If EBayStatus = "eYes" Then
        Dim testDoc As HTMLDocument
        Set testDoc = IEDOMFromhWnd(PriceCompareCSharp.EBayIEHandle)
        EBayResults = testDoc.documentElement.outerHTML
        'change status to "eNo"...
        PriceCompareCSharp.SetObjectText PriceCompareCSharp.EBayStatusHandle, "eNo"
        Dim Results As PriceCompareCSharp.PriceCompareInfo
        Results = PriceCompareCSharp.GetEBayComparisons(EBayResults)
        Dim xcount As Integer
        Do
            Dim li As ListItem
            With Me.EBayComparisonsLVI
            
                Set li = .ListItems.Add(, , CStr(xcount))
                li.SubItems(1) = Results.CompetitorPrice(xcount)
                li.SubItems(2) = Results.CompetitorName(xcount)
                'li.SubItems(3) = results.CompetitorURL(xcount)
            
            End With
            xcount = xcount + 1
        Loop Until xcount >= Results.ResultsCount
        
        ebaySortPriceBtn_Click
        
        
    End If
End If

If priceCompareEnabled = True And PriceCompareCSharp.GoogleIEHandle > 0 And PriceCompareCSharp.GoogleStatusHandle > 0 Then
    Dim GoogleStatus As String
    Dim googleResults As String
    GoogleStatus = PriceCompareCSharp.GetObjectText(PriceCompareCSharp.GoogleStatusHandle)
    If GoogleStatus = "gYes" Then
        Dim testGoogleDoc As HTMLDocument
        Set testGoogleDoc = IEDOMFromhWnd(PriceCompareCSharp.GoogleIEHandle)
        googleResults = testGoogleDoc.documentElement.outerHTML
        PriceCompareCSharp.SetObjectText PriceCompareCSharp.GoogleStatusHandle, "gNo"
        Dim CGoogleResults As PriceCompareCSharp.PriceCompareInfo
        CGoogleResults = PriceCompareCSharp.GetGoogleComparisons(googleResults)
        Dim gcount As Integer
        Do
            Dim gli As ListItem
            With Me.InternetComparisonLVI
                Set gli = .ListItems.Add(, , CStr(gcount))
                gli.SubItems(1) = CGoogleResults.CompetitorPrice(gcount)
                gli.SubItems(2) = CGoogleResults.CompetitorName(gcount)
            
            End With
        gcount = gcount + 1
        Loop Until gcount >= CGoogleResults.ResultsCount
        
    inetSortPriceBtn_Click
    End If



End If


End Sub

Private Sub priceComparison_Click()
Me.InternetComparisonLst.Selected(0) = True
Me.EBayComparisonsLst.Selected(0) = True

'12/12/05 - access to single item compare for everyone
'    If Not lockedDownPC Then
        If Me.priceComparison.Selected(0) Then
            Me.btnPriceCompareDetails.Enabled = False
        Else
            Me.btnPriceCompareDetails.Enabled = True
        End If
'    End If
End Sub

Private Sub priceComparison_DblClick()
    If Not lockedDownPC Then
        If Not Me.priceComparison.Selected(0) Then
            Dim newprice As String
            newprice = Format(DLookup("Price", "PriceComparisons", "ID=" & Mid(Me.priceComparison, InStrRev(Me.priceComparison, vbTab) + 1)), "Currency")
            If MsgBox("Change the WEB price for " & Me.ItemNumber & " to " & newprice & "?", vbYesNo) = vbYes Then
                LogItemWebPriceChanged Me.ItemNumber, newprice
                UpdateNotificationEmail "web price", Me.ItemNumber
                updateInventoryMaster "IDiscountMarkupPriceRate1", newprice, Me.ItemNumber, ""
                Me.IDiscountMarkupPriceRate(1) = newprice
                updateDropshipItemWebPrice Me.ItemNumber, Me.IDiscountMarkupPriceRate(1)
            End If
        End If
    End If
End Sub





'------------------------------------------------------------
' stuff that needs the form, but not related to a control
'------------------------------------------------------------
Private Sub requeryForm(filter As String)
    Mouse.Hourglass True
    If Me.cmbFilters <> filter Then
        Me.cmbFilters = filter
    End If
    Dim sqlstr As String
    Dim c As Long
    Dim rst As ADODB.Recordset
On Error Resume Next
    Select Case filter
        Case Is = "<List>"
            sqlstr = ""
        Case Is = "All"
            sqlstr = "SELECT ItemNumber FROM vInvMaintRecSource ORDER BY ItemNumber"
        Case Is = "FBF", "Filter By Form"
            sqlstr = Me.btnFilterByForm.tag
            Me.cmbFilters = "Filter By Form"
'        Case Is = "FBS", "Filter By Search"
'            sqlstr = Me.btnSearch.tag
'            Me.cmbFilters = "Filter By Search"
        Case Is = "Items On XSheet"
            'can't do this b/c it's a sparse array
            'sqlstr = ListAsCommaSep(xsheetarray, "'", True)
            For c = 0 To UBound(xsheetarray)
                If xsheetarray(c) <> "" Then
                    sqlstr = sqlstr & "'" & xsheetarray(c) & "', "
                End If
            Next c
            If sqlstr = "" Then
                sqlstr = "SELECT ItemNumber FROM vInvMaintRecSource WHERE 1=2"
            Else
                sqlstr = "SELECT ItemNumber FROM vInvMaintRecSource WHERE ItemNumber IN ( " & Left(sqlstr, Len(sqlstr) - 2) & " )"
            End If
        Case Is = "Filter By Sales"
            sqlstr = ListAsCommaSep(InventorySalesView.GetItemSelection(), "'", True, False)
            If sqlstr = "" Then
                sqlstr = "SELECT ItemNumber FROM vInvMaintRecSource WHERE 1=2"
            Else
                sqlstr = "SELECT ItemNumber FROM vInvMaintRecSource WHERE ItemNumber IN ( " & sqlstr & " )"
            End If
        'Case Is = "TNC ~= StdCost"
            'sqlstr = "SELECT ItemNumber FROM vInvMaintRecSource WHERE ItemNumber IN (SELECT ItemNumber,TNC,StdCost FROM InventoryMaster WHERE StdCost > (TNC * 0.98) AND StdCost < (TNC * 1.02))"
        
        Case Else
            Dim query As String
            query = DLookup("Query", "PoInvFilters", "FilterName='" & EscapeSQuotes(filter) & "'")
            If query = "" Then
                MsgBox "Inventory: can't filter for " & qq(filter) & ", defaulting to ""All""."
                sqlstr = "SELECT ItemNumber FROM vInvMaintRecSource ORDER BY ItemNumber"
            Else
                sqlstr = query
            End If
    End Select
    If sqlstr <> "" Then
        Set rst = DB.retrieve(sqlstr)
On Error GoTo 0
        If rst Is Nothing Then
            MsgBox "Unable to get the list of items. Probably an invalid SQL string, or timeout?" & vbCrLf & vbCrLf & "SQL follows:" & vbCrLf & sqlstr
            Mouse.Hourglass False
            Exit Sub
        End If
        If rst.RecordCount > 0 Then
            ReDim currentFilterItemList(rst.RecordCount - 1) As String
            Dim i As Long
            For i = 0 To rst.RecordCount - 1
                currentFilterItemList(i) = rst("ItemNumber")
                rst.MoveNext
            Next i
            Me.lblMaxAmt.caption = "of " & rst.RecordCount
            'Me.barPosition.Min = 0
            'Me.barPosition.max = rst.RecordCount - 1
        Else
            'ReDim currentFilterItemList(0) As String
            'currentFilterItemList(0) = "NORECORDS"
            'Me.lblMaxAmt.caption = "of 0"
            'Me.barPosition.Min = 0
            'Me.barPosition.max = 0
            currentFilterItemList = Split("", "zerolengtharray")
        End If
        rst.Close
        Set rst = Nothing
    Else
        Me.lblMaxAmt.caption = "of " & UBound(currentFilterItemList) + 1
    End If
    Mouse.Hourglass False
End Sub

Private Sub requeryVendor()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM vRequeryVendor ORDER BY VendorName")
    Me.PrimaryVendor.Clear
    While Not rst.EOF
        Me.PrimaryVendor.AddItem (Trim$(rst("VendorName")))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requeryFilters()
    Me.cmbFilters.Clear
    Me.cmbFilters.AddItem "All"
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT FilterName FROM PoInvFilters WHERE ShowInPoInv=1 ORDER BY FilterName")
    While Not rst.EOF
        Me.cmbFilters.AddItem rst("FilterName")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Me.cmbFilters.AddItem "Items On XSheet"
End Sub

'Private Sub requeryClass()
'    Mouse.Hourglass True
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT * FROM vRequeryShippingClass ORDER BY ShipClassName")
'    Me.Class.Clear
'    While Not rst.EOF
'        Me.Class.AddItem (rst("ShipClassName"))
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'    Mouse.Hourglass False
'End Sub

Private Sub requeryWhse()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT WarehouseDesc FROM IM_Warehouse ORDER BY WarehouseCode")
    Me.cmbWhse.Clear
    Me.cmbWhse.AddItem (WHSE_ALL)
    Me.cmbWhse.AddItem (WHSE_ACTUAL)
    While Not rst.EOF
        Me.cmbWhse.AddItem (Trim$(rst("WarehouseDesc")))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Me.cmbWhse = WHSE_ACTUAL
    Mouse.Hourglass False
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
        Me.jumpToItem.AddItem (Mid$(rst("ItemNumber"), 4))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    ExpandDropDownToFit Me.jumpToItem
    Mouse.Hourglass False
End Sub

Private Sub requeryFreightEstimates()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    'Set rst = DB.retrieve("exec spRequeryFreightEstimates '" & Me.ItemNumber & "'")
    Set rst = DB.retrieve("SELECT ZipCode, Cost, Service FROM FreightEstimates WHERE ItemNumber='" & Me.ItemNumber & "'")
    Me.freightEstimates.Clear
    Me.freightEstimates.AddItem "Zip Code" & vbTab & "Est Cost" & vbTab & "Svc"
    While Not rst.EOF
        Me.freightEstimates.AddItem rst("ZipCode") & vbTab & Format(rst("Cost"), "Currency") & vbTab & rst("Service")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requeryFreightActual()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    'Set rst = DB.retrieve("exec spRequeryFreightActual '" & Me.ItemNumber & "'")
    Set rst = DB.retrieve("SELECT ZipCode, Cost, SalesOrderNumber FROM FreightActual WHERE DateOfShipment > '" & Format(DateAdd("yyyy", -1, DateTime.Now), "yyyy-MM-dd") & "' AND ItemNumber='" & Me.ItemNumber & "' ORDER BY DateOfShipment DESC")
    Me.freightActual.Clear
    Me.freightActual.AddItem ("Zip Code" & vbTab & "Actual Cost")
    While Not rst.EOF
        Me.freightActual.AddItem (rst("ZipCode") & vbTab & Format(rst("Cost"), "Currency") & vbTab & rst("SalesOrderNumber"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub
Private Sub setupEBayDynamicAlgorithmLVIHeaders()

    EBayDynamicAlgorithmLVI.CheckBoxes = True
    EBayDynamicAlgorithmLVI.View = lvwReport
    EBayDynamicAlgorithmLVI.LabelEdit = lvwManual
    EBayDynamicAlgorithmLVI.ListItems.Clear
    EBayDynamicAlgorithmLVI.ColumnHeaders.Clear
    EBayDynamicAlgorithmLVI.HideColumnHeaders = True
    Dim checkCol As ColumnHeader
    Set checkCol = Me.EBayDynamicAlgorithmLVI.ColumnHeaders.Add()
    checkCol.Text = "Use"
    checkCol.width = EBayDynamicAlgorithmLVI.width / 8
    Dim algorithmCol As ColumnHeader
    Set algorithmCol = EBayDynamicAlgorithmLVI.ColumnHeaders.Add()
    algorithmCol.Text = "Algorithm"
    algorithmCol.width = EBayDynamicAlgorithmLVI.width * 0.6
    Dim percentCol As ColumnHeader
    Set percentCol = EBayDynamicAlgorithmLVI.ColumnHeaders.Add()
    percentCol.Text = "%"
    percentCol.width = EBayDynamicAlgorithmLVI.width / 6
    
   ' Dim testItem As ListItem
   ' Set testItem = InternetDynamicAlgorithmLVI.ListItems.Add()
   ' testItem.Text = ""
   ' testItem.SubItems(1) = "Algorithm 1"
    
End Sub

Public Sub setupEBayComparisonsLVIHeaders()

    'EBayComparisonsLVI.CheckBoxes = True
    EBayComparisonsLVI.View = lvwReport
    EBayComparisonsLVI.ColumnHeaders.Clear
    EBayComparisonsLVI.ListItems.Clear
    EBayComparisonsLVI.LabelEdit = lvwManual
    EBayComparisonsLVI.HideColumnHeaders = True
    Dim checkCol As ColumnHeader
    Set checkCol = EBayComparisonsLVI.ColumnHeaders.Add()
    checkCol.Text = "Use"
    checkCol.width = EBayComparisonsLVI.width / 1000
    
    Dim storeCol As ColumnHeader
    Dim priceCol As ColumnHeader
    Dim shippingCol As ColumnHeader
    Dim totalCol As ColumnHeader
    Set storeCol = EBayComparisonsLVI.ColumnHeaders.Add()
    storeCol.Text = "Store"
    storeCol.width = Me.EBayComparisonsLVI.width / 4
    Set priceCol = EBayComparisonsLVI.ColumnHeaders.Add()
    priceCol.Text = "Price"
    priceCol.width = Me.EBayComparisonsLVI.width / 2
    'Set shippingCol = EBayComparisonsLVI.ColumnHeaders.Add()
    'shippingCol.Text = "Shipping"
    'shippingCol.width = Me.EBayComparisonsLVI.width / 6
    'Set totalCol = Me.EBayComparisonsLVI.ColumnHeaders.Add()
    'totalCol.Text = "Total"
    'totalCol.width = Me.EBayComparisonsLVI.width / 5
    
    
    
    
    'Dim siteCol As ColumnHeader
    'Set siteCol = EBayComparisonsLVI.ColumnHeaders.Add()
    'siteCol.Text = "Site"
    'siteCol.width = EBayComparisonsLVI.width / 4
    'Dim sitePriceCol As ColumnHeader
    'Set sitePriceCol = EBayComparisonsLVI.ColumnHeaders.Add()
    'sitePriceCol.Text = "Price"
    'sitePriceCol.width = EBayComparisonsLVI.width / 4.8
    'Dim sitePriceDynCol As ColumnHeader
    'Set sitePriceDynCol = EBayComparisonsLVI.ColumnHeaders.Add()
    'sitePriceDynCol.Text = "Dyn Price"
    'sitePriceDynCol.width = EBayComparisonsLVI.width / 4.8
   
End Sub
Public Sub setupInternetComparisonLVIHeaders()
    'InternetComparisonLVI.CheckBoxes = True
    InternetComparisonLVI.View = lvwReport
    InternetComparisonLVI.ColumnHeaders.Clear
    InternetComparisonLVI.ListItems.Clear
    InternetComparisonLVI.LabelEdit = lvwManual
    Dim checkCol As ColumnHeader
    Set checkCol = InternetComparisonLVI.ColumnHeaders.Add()
    checkCol.Text = "Use"
    checkCol.width = InternetComparisonLVI.width / 1000
    InternetComparisonLVI.HideColumnHeaders = True
    
    Dim storeCol As ColumnHeader
    Dim priceCol As ColumnHeader
    Dim shippingCol As ColumnHeader
    Dim totalCol As ColumnHeader
    Set storeCol = InternetComparisonLVI.ColumnHeaders.Add()
    storeCol.Text = "Store"
    storeCol.width = InternetComparisonLVI.width / 4
    Set priceCol = InternetComparisonLVI.ColumnHeaders.Add()
    priceCol.Text = "Price"
    priceCol.width = InternetComparisonLVI.width / 2
    'Set shippingCol = InternetComparisonLVI.ColumnHeaders.Add()
    'shippingCol.Text = "Shipping"
    'shippingCol.width = InternetComparisonLVI.width / 6
    'Set totalCol = InternetComparisonLVI.ColumnHeaders.Add()
    'totalCol.Text = "Total"
    'totalCol.width = InternetComparisonLVI.width / 5
    
    
    'Dim siteCol As ColumnHeader
    'Set siteCol = InternetComparisonLVI.ColumnHeaders.Add()
    'siteCol.Text = "Site"
    'siteCol.width = InternetComparisonLVI.width / 2
    'Dim sitePriceCol As ColumnHeader
    'Set sitePriceCol = InternetComparisonLVI.ColumnHeaders.Add()
    'sitePriceCol.Text = "Price"
    'sitePriceCol.width = InternetComparisonLVI.width / 3.8
    'Dim sitePriceDynCol As ColumnHeader
    'Set sitePriceDynCol = InternetComparisonLVI.ColumnHeaders.Add()
    'sitePriceDynCol.Text = "Dyn Price"
    'sitePriceDynCol.width = InternetComparisonLVI.width / 3.8
    
    
    'Dim testItem As ListItem
    'Set testItem = InternetComparisonLVI.ListItems.Add()
    'testItem.Text = ""
    'testItem.SubItems(1) = "Test.com"
    'testItem.SubItems(2) = "$100.00"
    'testItem.SubItems(3) = "$010.00"
    'Dim testItem2 As ListItem
    'Set testItem2 = InternetComparisonLVI.ListItems.Add()
    'testItem2.Text = ""
    'testItem2.SubItems(1) = "Test.com"
    'testItem2.SubItems(2) = "$200.00"
    'testItem2.SubItems(3) = "$020.00"
    'Dim testItem3 As ListItem
    'Set testItem3 = InternetComparisonLVI.ListItems.Add()
    'testItem3.Text = ""
    'testItem3.SubItems(1) = "Test.com"
    'testItem3.SubItems(2) = "$300.00"
    'testItem3.SubItems(3) = "$030.00"
    
End Sub
Private Sub setupInternetDynamicAlgorithmLVI()
    InternetDynamicAlgorithmLVI.CheckBoxes = True
    InternetDynamicAlgorithmLVI.View = lvwReport
    InternetDynamicAlgorithmLVI.LabelEdit = lvwManual
    Me.InternetDynamicAlgorithmLVI.ListItems.Clear
    Me.InternetDynamicAlgorithmLVI.ColumnHeaders.Clear
    Me.InternetDynamicAlgorithmLVI.HideColumnHeaders = True
    Dim checkCol As ColumnHeader
    Set checkCol = Me.InternetDynamicAlgorithmLVI.ColumnHeaders.Add()
    checkCol.Text = "Use"
    checkCol.width = InternetDynamicAlgorithmLVI.width / 8
    Dim algorithmCol As ColumnHeader
    Set algorithmCol = InternetDynamicAlgorithmLVI.ColumnHeaders.Add()
    algorithmCol.Text = "Algorithm"
    algorithmCol.width = InternetDynamicAlgorithmLVI.width * 0.6
    Dim percentCol As ColumnHeader
    Set percentCol = InternetDynamicAlgorithmLVI.ColumnHeaders.Add()
    percentCol.Text = "%"
    percentCol.width = InternetDynamicAlgorithmLVI.width / 6
    
   ' Dim testItem As ListItem
   ' Set testItem = InternetDynamicAlgorithmLVI.ListItems.Add()
   ' testItem.Text = ""
   ' testItem.SubItems(1) = "Algorithm 1"
    
    
End Sub





Private Sub requeryPriceComparisons()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Dim competitorInfo As ADODB.Recordset
    
    
    Dim sqlQuery As String
    sqlQuery = "SELECT * FROM PriceComparisons2 WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
    
    'Set rst = DB.retrieve("exec spRequeryPriceComparisons '" & Me.ItemNumber & "'")
    
       ' Set rst = DB.retrieve("SELECT * FROM PriceComparisons2 WHERE ItemNumber='" & Me.ItemNumber & "'")

'If Me.inetAmazonBtn.FontBold = True And Me.inetFroogleBtn.FontBold = False Then
'    sqlQuery = sqlQuery & " AND StoreName LIKE 'Amazon%'"
'End If
 
'If Me.inetFroogleBtn.FontBold = True And Me.inetAmazonBtn.FontBold = False Then
'    sqlQuery = sqlQuery & " AND StoreName LIKE 'Froogle%'"
'End If
'If Me.inetAmazonBtn.FontBold = False And Me.inetFroogleBtn.FontBold = False Then
'    sqlQuery = sqlQuery & " AND StoreName NOT LIKE 'Froogle%' AND StoreName NOT LIKE 'Amazon%'"
'End If
'If Me.inetFroogleBtn.FontBold = True And Me.inetAmazonBtn.FontBold = True Then
'    sqlQuery = sqlQuery & " AND StoreName LIKE 'Froogle%' OR ItemNumber='" & Me.ItemNumber.Text & "' AND StoreName LIKE 'Amazon%'"
'End If


Set rst = DB.retrieve(sqlQuery)

    
    Me.priceComparison.Clear
    Me.InternetComparisonLst.Clear
    Me.EBayComparisonsLst.Clear

    setupInternetComparisonLVIHeaders
    setupInternetDynamicAlgorithmLVI
    
    'setupEBayComparisonsLVIHeaders
    'setupEBayDynamicAlgorithmLVIHeaders
    
    Me.InternetAlgorithmCmb.Clear
    
    
    
    Me.btnPriceCompareDetails.Enabled = False
    Me.priceComparison.AddItem ("Store" & vbTab & "Price" & vbTab & "Higher?" & vbTab & "Lower?" & vbTab & "ID#")
    Me.EBayComparisonsLst.AddItem ("Store" & vbTab & vbTab & "Price")
    Me.InternetComparisonLst.AddItem ("Store" & vbTab & vbTab & "Price")
    
    'Dim isBaseItem As Boolean
    'isBaseItem = CBool(InStr(1, Me.ItemDescription, "BASE ITEM", vbTextCompare))
    While Not rst.EOF
            Dim allowFlag As Boolean
            allowFlag = True
            Set competitorInfo = DB.retrieve("SELECT Enabled,Blacklisted,Whitelisted,AddTax FROM Competitors WHERE CompetitorName='" & Replace(rst("StoreName"), "'", "''") & "'")
            If competitorInfo.RecordCount > 0 Then
                If competitorInfo("Enabled") = False Then
                   allowFlag = False
                End If
                If competitorInfo("Blacklisted") = True Then
                    allowFlag = False
                End If
            End If
        'If rst("Display") Or isBaseItem Then
        If allowFlag = True Then
        'If rst("Display") Then
            'Me.priceComparison.AddItem (Left$(rst("StoreName"), 26) & IIf(Len(Trim$(rst("StoreName"))) < 27, "", "...") & vbTab & Format(rst("Price"), "Currency") & vbTab & rst("HigherPrice") & vbTab & rst("LowerPrice") & vbTab & rst("ID"))
            If (InStr(UCase(rst("StoreName")), "EBAY") > 0 Or InStr(UCase(rst("StoreName")), "E-BAY") > 0) Then
                If (UCase(rst("Source")) = "Wiser EBay") Then
                
                    Me.EBayComparisonsLst.AddItem Left$(Replace(rst("StoreName"), "eBay", ""), 16) & IIf(Len(Trim$(rst("StoreName"))) < 17, "", "...") & IIf(Len(Trim$(rst("StoreName"))) < 11, vbTab, "") & vbTab & Format(rst("Price"), "Currency") & vbTab & vbTab & rst("ID")
                    Dim newListItem2 As ListItem
                    Set newListItem2 = Me.EBayComparisonsLVI.ListItems.Add()
                    newListItem2.Text = ""
                    Dim nameItm As String
                    nameItm = Replace(rst("StoreName"), "eBay -", "E- ")
                    newListItem2.SubItems(1) = Replace(nameItm, "E-  E-", "")
                    newListItem2.SubItems(1) = Replace(newListItem2.SubItems(1), "E-", "")
                    newListItem2.SubItems(2) = Format(rst("Price"), "Currency")
                    'newListItem2.SubItems(3) = "N/A"
                    newListItem2.SubItems(3) = Format(rst("Shipping"), "Currency")
                    
                    newListItem2.SubItems(4) = Format(CDbl(rst("Shipping")) + CDbl(rst("Price")), "Currency")
                    newListItem2.tag = rst("ID")
                End If
            Else
                If Me.inetAmazonBtn.caption = "A" And InStr(UCase(rst("StoreName")), "Amazon") <= 0 Then
                               
                    Dim newListItem As ListItem
                    Set newListItem = Me.InternetComparisonLVI.ListItems.Add()
                    newListItem.Text = ""
                    newListItem.SubItems(1) = Replace(rst("StoreName"), "Amazon -", "A- ")
                    newListItem.SubItems(2) = Format(rst("Price"), "Currency")
                If (competitorInfo.EOF = False And competitorInfo.BOF = False) Then
                If CBool(competitorInfo("AddTax")) = True Then
                    Dim calcPrice As Currency
                    calcPrice = Format(CDec((CDec(rst("Price")) + CDec(rst("Shipping"))) * 0.0635), "Currency")
                    newListItem.SubItems(3) = Format(calcPrice, "Currency")
                Else
                    newListItem.SubItems(3) = Format(rst("Shipping"), "Currency")
                End If
                Else
                    newListItem.SubItems(3) = Format(rst("Shipping"), "Currency")
                End If
                    newListItem.SubItems(4) = Format(CDbl(newListItem.SubItems(2)) + CDbl(newListItem.SubItems(3)), "Currency")
                    newListItem.tag = rst("ID")
                    Me.InternetComparisonLst.AddItem Left$(rst("StoreName"), 16) & IIf(Len(Trim$(rst("StoreName"))) < 17, "", "...") & IIf(Len(Trim$(rst("StoreName"))) < 11, vbTab, "") & vbTab & Format(rst("Price"), "Currency") & vbTab & vbTab & rst("ID")
                
                End If
                If (Me.inetAmazonBtn.caption = "G" And InStr(UCase(rst("StoreName")), "Amazon") > 0) Then
                    Dim newListItem3 As ListItem
                    Set newListItem3 = Me.InternetComparisonLVI.ListItems.Add()
                    newListItem3.Text = ""
                    newListItem3.SubItems(1) = Replace(rst("StoreName"), "Amazon -", "A- ")
                    newListItem3.SubItems(2) = Format(rst("Price"), "Currency")
                    newListItem3.SubItems(3) = Format(rst("Shipping"), "Currency")
                    newListItem3.SubItems(4) = Format(CDbl(rst("Shipping")) + CDbl(rst("Price")), "Currency")
                    newListItem3.tag = rst("ID")
                    Me.InternetComparisonLst.AddItem Left$(rst("StoreName"), 16) & IIf(Len(Trim$(rst("StoreName"))) < 17, "", "...") & IIf(Len(Trim$(rst("StoreName"))) < 11, vbTab, "") & vbTab & Format(rst("Price"), "Currency") & vbTab & vbTab & rst("ID")
            
                End If
            
            
            
            
            End If
       ' End If
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requeryPOList(VendorNumber As String)
    
    
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT PurchaseOrders.ID, VendorNumber, CoopVendor, ProductLineCode, SUM(PurchaseOrderLines.Quantity * InventoryMaster.StdCost) AS POAmount, ShiptoCode, ShipToName " & _
                          "FROM PurchaseOrders " & _
                          "          INNER JOIN PurchaseOrderLines ON PurchaseOrders.ID=PurchaseOrderLines.HeaderID " & _
                          "          INNER JOIN InventoryMaster ON PurchaseOrderLines.ItemNumber=InventoryMaster.ItemNumber " & _
                          "WHERE PurchaseOrders.VendorNumber='" & VendorNumber & "' " & _
                          "  AND PurchaseOrders.Exported=0 " & _
                          "  AND PurchaseOrders.ProductLineCode=(SELECT CASE WHEN RealLineCode IS NULL THEN ProductLine ELSE RealLineCode END AS ProductLineCode " & _
                          "                                      FROM ProductLine LEFT OUTER JOIN ProductLineSubForPOs ON ProductLine.ProductLine=ProductLineSubForPOs.AlternateLineCode " & _
                          "                                      WHERE ProductLine='" & IIf(Me.prodLine = "XXX", Mid(Me.ItemNumber, 4, 3), Me.prodLine) & "') " & _
                          "GROUP BY PurchaseOrders.ID, PurchaseOrders.VendorNumber, PurchaseOrders.CoopVendor, PurchaseOrders.ProductLineCode, PurchaseOrders.ShiptoCode, PurchaseOrders.ShipToName")
    
    If rst.RecordCount > 1 Then
        Disable Me.QtyOrdered
    Else
        Enable Me.QtyOrdered
    End If
    'Me.QtyOrdered.Locked = lockedDown Or rst.RecordCount > 1
    
    Me.POList.Clear
    'Me.POList.AddItem ("Vendor" & vbTab & "Total" & vbTab & "Minimum" & vbTab & "PPD @")
    'Me.POList.AddItem ("Vendor" & vbTab & "Total")
    'Dim vendor As String, minimum As String, ppd As String
    While Not rst.EOF
        'vendor = rst("VendorNumber") & IIf(rst("VendorNumber") = rst("CoopVendor"), "", " / " & rst("CoopVendor"))
        'minimum = FormatCurrency(rst("MinimumOrderAmount"), 0)
        'ppd = IIf(IsNumeric(rst("PrepaidAmount")), FormatCurrency(rst("PrepaidAmount"), 0), rst("PrepaidAmount")) & IIf(rst("PrepaidSpecialAmount") = 0 Or Nz(rst("PrepaidSpecialAmount")) = "", "", " (" & FormatCurrency(rst("PrepaidSpecialAmount"), 0) & ")")
        'Me.POList.AddItem vendor & vbTab & Format(rst("POAmount"), "Currency") & vbTab & minimum & vbTab & ppd
        Dim ln As String
        If rst("ShiptoCode") = "1234" Then
            ln = rst("ShipToName")
            If Len(ln) > 15 Then
                ln = Left(ln, 15)
            End If
        Else
            ln = rst("VendorNumber") & IIf(rst("VendorNumber") = rst("CoopVendor"), "", " / " & rst("CoopVendor"))
        End If
        ln = ln & vbTab & Format(rst("POAmount"), "Currency") & vbTab & rst("ID")
        Me.POList.AddItem ln
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub
Public Sub CheckDynamicPriceEnabled()
    'lets check for dynamic price for internet (TPWeb)...
    Dim checkWeb As ADODB.Recordset
    Set checkWeb = DB.retrieve("SELECT * FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPWeb' AND Enabled=1")
    
    Dim AlgorithmID As String
    If checkWeb.RecordCount > 0 Then
        AlgorithmID = checkWeb("AlgorithmID")
        
        Dim getAlg As ADODB.Recordset
        Set getAlg = DB.retrieve("SELECT * FROM DynamicPricingAlgorithmHeader WHERE ID=" & AlgorithmID)
        If getAlg.RecordCount > 0 Then
            Dim findAlg As String
            Dim xcount As Integer
            For xcount = 0 To Me.InternetAlgorithmCmb.ListCount - 1
            If InStr(Me.InternetAlgorithmCmb.list(xcount), ":") > 0 Then
                If Split(Me.InternetAlgorithmCmb.list(xcount), ":")(0) = getAlg("AlgorithmName") Then
                    Me.InternetAlgorithmCmb.ListIndex = xcount
                    'now that we are here, we have to check to see if a specific competitor is checked... and check em off if so...
                    Dim findCompetitorID As String
                        findCompetitorID = checkWeb("CompetitorDataID")
                    If checkWeb("CompetitorDataID") <> 0 Then
                        'houston, we have a dynamically set price!
                        
                        Dim checkLVI As ListItem
                        For Each checkLVI In Me.InternetComparisonLVI.ListItems
                            If checkLVI.tag = findCompetitorID Then
                                checkLVI.Checked = True
                            End If
                        Next
                    End If
                End If
            End If
            Next xcount

        
        Else
        
        
        End If
        
    Else
        'Me.InternetAlgorithmCmb.ListIndex = 0
    End If
    

    
    
    
    'lets check for dynamic price for ebay (TPEBay)...
    'Dim algorithmID As ListItem
'    Dim comparisons As ListItem
'    Dim checkEBay As ADODB.Recordset
'    Set checkEBay = DB.retrieve("SELECT * FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "' AND PriceFieldToChange='TPEBay' AND Enabled=1")
'    If checkEBay.RecordCount > 0 Then
'        'Dim algorithmID As ListItem
'        For Each algorithmID In Me.EBayDynamicAlgorithmLVI.ListItems
'            If algorithmID.tag = checkEBay("AlgorithmID") Then
'                Me.DynamicDontMessWith = True
'                algorithmID.Checked = True
'                algorithmID.SubItems(2) = CStr(checkEBay("AlgorithmPercentage")) & "%"
'                'Dim comparisons As ListItem
'                For Each comparisons In Me.EBayComparisonsLVI.ListItems
'                    If comparisons.tag = checkEBay("CompetitorDataID") Then
'                        comparisons.Checked = True
'                    End If
'
'                    If comparisons.Checked = True Then
'                        Dim dynCheck As ADODB.Recordset
'                        Set dynCheck = DB.retrieve("DECLARE @input DECIMAL(7,3) SET @input=" & Format(comparisons.SubItems(2), "Currency") & " SELECT TOP 1 DynamicPricingAlgorithmLines.ID,HeaderID,DifferentialPercentage,DollarAmount,MaxPercentAbove,MaxPercentBelow FROM DynamicPricingAlgorithmLines INNER JOIN DynamicPricingAlgorithmHeader ON DynamicPricingAlgorithmHeader.ID=HeaderID WHERE HeaderID=" & algorithmID.tag & " ORDER BY ABS(@input - DollarAmount)")
'                        If dynCheck.RecordCount > 0 Then
'                            Dim minPrice As Currency
'                            Dim maxPrice As Currency
'                            Dim dynPrice As Currency
'                            minPrice = Format((Format(dynCheck("MaxPercentBelow"), "Currency")) / 100 * (Format(dynCheck("DollarAmount"), "Currency")), "Currency")
'                            maxPrice = Format(Format((1 + (Format(dynCheck("MaxPercentAbove"), "Currency") / 100)), "Currency") * Format(dynCheck("DollarAmount"), "Currency"), "Currency")
'                            dynPrice = Format(Replace(comparisons.SubItems(2), "%", "") * (1 + (CDec(checkEBay("AlgorithmPercentage"))) / 100), "Currency")
'                            If dynPrice < maxPrice And dynPrice > minPrice Then
'                                comparisons.SubItems(3) = Format(dynPrice, "Currency")
'                            Else
'                                If dynPrice > maxPrice Then
'                                    dynPrice = maxPrice
'                                    comparisons.SubItems(3) = CStr(Format(dynPrice, "Currency"))
'                                Else
'                                    dynPrice = minPrice
'                                    comparisons.SubItems(3) = Format(dynPrice, "Currency")
'                                End If
'                            End If
'                            comparisons.SubItems(3) = Format(dynPrice, "Currency")
'                        End If
'
'                    End If
'                Next
'                Me.DynamicDontMessWith = False
'            Else
'                algorithmID.Checked = False
'            End If
'        Next
'
'        Dim compFlag As Boolean
'        compFlag = False
'        Dim competitorID As ListItem
'        For Each competitorID In Me.EBayComparisonsLVI.ListItems
'            If competitorID.tag = checkEBay("CompetitorDataID") Then
'                compFlag = True
'                competitorID.Checked = True
'                If Me.EBayPrice.Text <> Format(dynPrice, "Currency") Then
'                    Me.EBayPrice.Text = Format(dynPrice, "Currency")
'                    DB.Execute "UPDATE InventoryMaster SET EBayPrice=" & Format(dynPrice, "Currency") & " WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
'                    DB.Execute "UPDATE PartNumbers SET EBayToBePublished=1 WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
'
'                End If
'            Else
'                competitorID.Checked = False
'            End If
'        Next
'        If compFlag = False Then
'            Me.EBayPrice.Text = Format(checkWeb("OriginalPrice1"), "Currency")
'            DB.Execute "UPDATE InventoryMaster SET EBayPrice=" & Format(checkEBay("OriginalPrice1"), "Currency") & " WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
'            DB.Execute "UPDATE PartNumbers SET EBayToBePublished=1 WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
'        End If
'    Else
'        Me.EBayAlgorithmCmb.ListIndex = 0
'    End If
'
    
    
    
End Sub
Public Sub CheckDynamicPriceEnabled2()

    Dim check As ADODB.Recordset
    Set check = DB.retrieve("SELECT MinPrice,MaxPrice,Price,NewPrice,Enabled,LastUpdate,EBay,TPStore,TPWeb,JetDotCom FROM PriceWiserItems WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
    If check.RecordCount > 0 Then
        If CBool(check("Enabled")) = True Then
            Me.dynamicPriceChk.value = vbChecked
        End If
        If check("EBay") = True Then
            'DynamicPricingFrm.dynamicEBayChk.value = vbChecked
        Else
            'DynamicPricingFrm.dynamicEBayChk.value = vbUnchecked
        End If
        If check("TPStore") = True Then
            'DynamicPricingFrm.dynamicStoreChk.value = vbChecked
        Else
            'DynamicPricingFrm.dynamicStoreChk.value = vbUnchecked
        End If
        If check("TPWeb") = True Then
            'DynamicPricingFrm.dynamicSiteChk.value = vbChecked
        Else
            'DynamicPricingFrm.dynamicSiteChk.value = vbUnchecked
        End If
        If check("JetDotCom") = True Then
            'DynamicPricingFrm.dynamicJetChk.value = vbChecked
            'perfect time to grab jet.com price from sqldb
        Else
            'DynamicPricingFrm.dynamicJetChk.value = vbUnchecked
        End If
        'DynamicPricingFrm.dynamicPriceLbl.Caption = "Dynamic Price: $" & CStr(check("NewPrice"))
        'DynamicPricingFrm.SetCurrentDynamicPrice CStr(check("NewPrice"))
        'DynamicPricingFrm.dynamicPriceLbl.Enabled = True
        'DynamicPricingFrm.itemNumberLbl.Caption = Me.ItemNumber.Text
        'DynamicPricingFrm.ebayPriceLbl.Caption = "EBay Price: " & Me.EBayPrice.Text
        'DynamicPricingFrm.storePriceLbl.Caption = "TP-Store Price: " & Me.DiscountMarkupPriceRate(1).Text
        'DynamicPricingFrm.tpSitePriceLbl.Caption = "TP-Site Price: " & Me.IDiscountMarkupPriceRate(1).Text
        'DynamicPricingFrm.jetDotComPriceLbl.Caption = "Jet.com Price: " & Me.EBayPrice.Text
        'DynamicPricingFrm.jetDotComPriceLbl.Caption = "Jet.com Price: $" & Me.JetPrice.Text
        'DynamicPricingFrm.dynamicEBayChk.Enabled = True
        'DynamicPricingFrm.dynamicSiteChk.Enabled = True
        'DynamicPricingFrm.dynamicStoreChk.Enabled = True
        'DynamicPricingFrm.dynamicJetChk.Enabled = True
        'DynamicPricingFrm.reserveMinLbl.Caption = "Reserve Min: $"
        'DynamicPricingFrm.reserveMinTxt.Text = CStr(check("MinPrice"))
        'DynamicPricingFrm.currentMinPrice = CStr(check("MinPrice"))
        
        'DynamicPricingFrm.lastCheckedLbl.Caption = "Checked: " & check("LastUpdate")
        

        
        
    Else
        Me.dynamicPriceChk.value = vbUnchecked
        'DynamicPricingFrm.dynamicEBayChk.value = vbUnchecked
        'DynamicPricingFrm.dynamicSiteChk.value = vbUnchecked
        'DynamicPricingFrm.dynamicStoreChk.value = vbUnchecked
        'DynamicPricingFrm.dynamicJetChk.value = vbUnchecked
        'DynamicPricingFrm.dynamicPriceLbl.Caption = "Dynamic Price: N/A"
        'DynamicPricingFrm.dynamicPriceLbl.Enabled = False
        'DynamicPricingFrm.itemNumberLbl.Caption = "ItemNumber"
        'DynamicPricingFrm.ebayPriceLbl.Caption = "EBay Price: "
        'DynamicPricingFrm.storePriceLbl.Caption = "TP-Store Price: "
        'DynamicPricingFrm.tpSitePriceLbl.Caption = "TP-Site Price: "
        'DynamicPricingFrm.jetDotComPriceLbl.Caption = "Jet.com Price: $"
        'DynamicPricingFrm.dynamicEBayChk.Enabled = False
        'DynamicPricingFrm.dynamicSiteChk.Enabled = False
        'DynamicPricingFrm.dynamicStoreChk.Enabled = False
        'DynamicPricingFrm.dynamicJetChk.Enabled = False
        'DynamicPricingFrm.reserveMinLbl.Caption = "Reserve Min: N/A"
        'DynamicPricingFrm.lastCheckedLbl.Caption = "Checked: N/A"
        
    End If

'we are going to populate dynamic form with competitive pricing...
'DynamicPricingFrm.marketLst.Clear
'DynamicPricingFrm.sitesLst.Clear

Dim Competition As ADODB.Recordset
Set Competition = DB.retrieve("SELECT StoreName, Price FROM PriceComparisons WHERE Source='Price2Spy' AND ItemNumber='" & Me.ItemNumber.Text & "'")
If Competition.RecordCount > 0 Then
    Do
        Dim StoreName As String
        Dim StorePrice As Double
        StoreName = Competition("StoreName")
        StorePrice = CDec(Competition("Price"))
        PriceCompareStoreOrMarketplace StoreName, StorePrice
        Competition.MoveNext
    Loop Until Competition.EOF = True Or Competition.BOF = True
    

End If



End Sub
Private Sub PriceCompareStoreOrMarketplace(StoreName As String, StorePrice As Double)
    Dim realStoreName As String
    realStoreName = Mid(StoreName, 2, Len(StoreName) - 1)
    'If realStoreName = "Cpooutlets" Then DynamicPricingFrm.marketLst.AddItem StoreName & vbTab & "$" & CStr(StorePrice)
    'If realStoreName = "toolup.com" Then DynamicPricingFrm.sitesLst.AddItem StoreName & vbTab & "$" & CStr(StorePrice)
    
    
End Sub
Public Sub fillDynamicPricing()



'DynamicPricingFrm.SetupColumns

    Dim iNumber As String
    iNumber = Me.ItemNumber.Text
    
    Dim dynCheck As ADODB.Recordset
    Set dynCheck = DB.retrieve("SELECT * FROM PriceWiserItems WHERE ItemNumber='" & iNumber & "'")
    'MinPrice,MaxPrice,Price,NewPrice,LastUpdate,Enabled,EBay,TPStore,TPWeb,JetDotCom
    If dynCheck.RecordCount > 0 Then
        'lets iterate thru the results and parse out all webplaces instead of manually coding them in...
        'DynamicPricingFrm.ListView1.ListItems.Clear
        Dim tmp As field
        For Each tmp In dynCheck.fields
            'MsgBox tmp.name
            If tmp.name <> "ID" And tmp.name <> "ItemNumber" And tmp.name <> "MinPrice" And tmp.name <> "MaxPrice" And tmp.name <> "Price" And tmp.name <> "NewPrice" And tmp.name <> "LastUpdate" And tmp.name <> "Enabled" Then
                'Dim newStore As ListItem
                'Set newStore = DynamicPricingFrm.ListView1.ListItems.Add()
                'newStore.Text = ""
                'newStore.Checked = CBool(dynCheck(tmp.name))
                'newStore.SubItems(1) = tmp.name
                
                If tmp.name = "JetDotCom" Then
                    'newStore.SubItems(2) = IIf(CBool(dynCheck("Enabled") = True) And CBool(newStore.Checked) = True, IIf(Format(dynCheck("NewPrice"), "Currency") > Format(Me.JetPrice.Text, "Currency"), Format(dynCheck("NewPrice"), "Currency"), Format(dynCheck("MinPrice"), "Currency")), Format(Me.JetPrice.Text, "Currency"))
                    If CBool(dynCheck("Enabled")) = True Then
                        'If CBool(newStore.Checked) = True Then
                            'Me.JetPrice.BackColor = vbCyan
                            'If Format(dynCheck("NewPrice"), "Currency") > Me.JetPrice.Text Then
                            '    newStore.SubItems(2) = Format(dynCheck("NewPrice"), "Currency")
                            'Else
                            '    newStore.SubItems(2) = Format(Me.JetPrice.Text, "Currency")
                            'End If
                        'Else
                            'Me.JetPrice.BackColor = vbWhite
                            'newStore.SubItems(2) = Format(Me.JetPrice.Text, "Currency")
                        'End If
                    End If
                    
                End If
                If tmp.name = "EBay" Then
                    'newStore.SubItems(2) = IIf(CBool(dynCheck("Enabled") = True) And CBool(newStore.Checked) = True, "N/A Now!", Format(Me.EBayPrice.Text, "Currency"))
                End If
                If tmp.name = "TPStore" Then
                    'newStore.SubItems(2) = IIf(CBool(dynCheck("Enabled") = True) And CBool(newStore.Checked) = True, "N/A Now!", Format(Me.DiscountMarkupPriceRate(1), "Currency"))
                End If
                If tmp.name = "TPWeb" Then
                    'newStore.SubItems(2) = IIf(CBool(dynCheck("Enabled") = True) And CBool(newStore.Checked) = True, "N/A Now!", Format(Me.IDiscountMarkupPriceRate(1), "Currency"))
                End If
                
                
            End If
            
        Next
        
        
        'there are records indicating dynamic pricing for this item. let's dig a bit deeper...
        'so we will ask some question, if the answer is false then just move on.
        'first let's see if the dynamic rule is even on to begin with.
        If dynCheck("Enabled") = True Then
            'so the dynamic frm should have data now, but possibly nothing checked off..
            'DynamicPricingFrm.EnableFormControls
            If dynCheck("EBay") = True Then
                'set ebay price to dynamic price. Check on both InventoryMaintenance and on DynamicPricingFrm.
                
            End If
            If dynCheck("TPStore") = True Then
                'set tp store price to dynamic price. Check on both InventoryMaintenance and on DynamicPricingFrm.
                
            End If
            If dynCheck("TPWeb") = True Then
                'set tpweb price to dynamic price. Check on both InventoryMaintenance and on DynamicPricingFrm
                
            End If
            If dynCheck("JetDotCom") = True Then
                'set JetDotCom price to dynamic price. Check on both InventoryMaintenance and on DynamicPricingFrm
                
                If Len(InventoryMaintenance.JetPrice.Text) > 0 Then
                    If CDec(InventoryMaintenance.JetPrice.Text) <> CDec(dynCheck("NewPrice")) Then
                        InventoryMaintenance.JetPrice.Text = IIf(dynCheck("NewPrice") > dynCheck("MinPrice"), Format(dynCheck("NewPrice"), "Currency"), Format(dynCheck("MinPrice"), "Currency"))
                        
                        InventoryMaintenance.JetPrice.BackColor = vbCyan
                        'If CDec(dynCheck("NewPrice")) <> Split(DynamicPricingFrm.dynamicPriceLbl.Caption, "$")(1) Then
                        '     DynamicPricingFrm.dynamicPriceLbl.Caption = "Dynamic Price: $" & dynCheck("NewPrice")
                        '     DynamicPricingFrm.ListView1.ListItems(4).Checked = True
                        'End If
                    End If
                Else
                    InventoryMaintenance.JetPrice.Text = Format(dynCheck("NewPrice"), "Currency")
                End If
            Else
                'DynamicPricingFrm.ListView1.ListItems(4).Checked = False
            End If
        Else
            'so dynamic pricing is disabled but there is data, so disable the dyn frm but populate it too.
            'DynamicPricingFrm.DisableFormControls
            InventoryMaintenance.JetPrice.BackColor = vbWhite
            
        End If
    Else
        'there is no dynamic info for this item, disable all controls on dyn frm.
        'DynamicPricingFrm.DisableFormControls
        
    End If
    
    setupInternetDynamicAlgorithmLVI
    
    
    Dim getAlgs As ADODB.Recordset
    Set getAlgs = DB.retrieve("SELECT ID,AlgorithmName,DifferentialPercentage FROM DynamicPricingAlgorithmHeader")
    If getAlgs.RecordCount > 0 Then
        Me.InternetAlgorithmCmb.Clear
        Me.EBayAlgorithmCmb.Clear
        'Me.InternetAlgorithmCmb.Text = "-Choose an Algorithm-"
        'Me.EBayAlgorithmCmb.Text = "-Choose an Algorithm-"
        Me.InternetAlgorithmCmb.AddItem ("Do Not Dynamically Price")
        Me.EBayAlgorithmCmb.AddItem ("Do Not Dynamically Price")
        Do
            Me.InternetAlgorithmCmb.AddItem (getAlgs("AlgorithmName") & ":   " & getAlgs("DifferentialPercentage") & "%")
            Me.EBayAlgorithmCmb.AddItem (getAlgs("AlgorithmName") & ":   " & getAlgs("DifferentialPercentage") & "%")
            
            'Dim newItm As ListItem
            'Set newItm = Me.InternetDynamicAlgorithmLVI.ListItems.Add()
            'newItm.Text = ""
            'newItm.SubItems(1) = getAlgs("AlgorithmName")
            'newItm.SubItems(2) = getAlgs("DifferentialPercentage") & "%"
            'newItm.tag = getAlgs("ID")
            'Dim newItm2 As ListItem
            'Set newItm2 = Me.EBayDynamicAlgorithmLVI.ListItems.Add()
            'newItm2.Text = ""
            'newItm2.SubItems(1) = getAlgs("AlgorithmName")
            'newItm2.SubItems(2) = getAlgs("DifferentialPercentage") & "%"
            'newItm2.tag = getAlgs("ID")
            getAlgs.MoveNext
        Loop Until getAlgs.EOF = True Or getAlgs.BOF = True
    
    End If
    
    
End Sub

Public Sub fillForm(item As ADODB.Recordset)
    Mouse.Hourglass True
    


OriginalIDiscountPrice = ""
OriginalEBayPrice = ""
    
    If Me.jumpToLine = "" Or Left(item("ItemNumber"), 3) <> Left(Me.ItemNumber, 3) Then
        Me.jumpToLine = Left(item("ItemNumber"), 3)
        requeryJumpToItem
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT MinimumOrderAmount, PrepaidAmount, PrepaidSpecialAmount FROM ProductLine WHERE ProductLine='" & item("ProductLine") & "'")
        Me.MinimumOrderAmount = IIf(IsNumeric(rst("MinimumOrderAmount")), Format(rst("MinimumOrderAmount"), "Currency"), Nz(rst("MinimumOrderAmount")))
        Me.PrepaidAmount = IIf(IsNumeric(rst("PrepaidAmount")), Format(rst("PrepaidAmount"), "Currency"), Nz(rst("PrepaidAmount")))
        Me.PrepaidSpecialAmount = IIf(IsNumeric(rst("PrepaidSpecialAmount")), Format(rst("PrepaidSpecialAmount"), "Currency"), Nz(rst("PrepaidSpecialAmount")))
        rst.Close
        Set rst = Nothing
    End If
    Me.jumpToItem = Mid(item("ItemNumber"), 4)
    Me.Website = WebsiteNameHash.item(CLng(item("WebsiteID").value))
    Me.Website.tag = item("WebsiteID").value
    Me.LoadedItemID = item("ID")
    Me.ItemNumber = item("ItemNumber")
    
    If IsNumeric(item("AvailabilityLimit")) = False Then
        Me.AvailabilityLimit = "0"
    Else
        Me.AvailabilityLimit = item("AvailabilityLimit")
    End If
    If IsNumeric(item("EBayReserveQty")) = False Then
        Me.EBayAvailabilityLimit = "0"
    Else
        Me.EBayAvailabilityLimit = item("EBayReserveQty")
    End If
    
    
    
    'Me.AvailabilityLimit = item("AvailabilityLimit")
    'Me.EBayAvailabilityLimit = item("EBayReserveQty")
    'DynamicPricingFrm.itemNumberLbl.Caption = item("ItemNumber")
    'DynamicPricingFrm.dynamicPriceLbl.Caption = "Dynamic Price: N/A"
    'DynamicPricingFrm.applyBtn.Enabled = False
    'DynamicPricingFrm.Enabled = False
    'DynamicPricingFrm.dynamicEBayChk.Enabled = False
    'DynamicPricingFrm.dynamicJetChk.Enabled = False
    'DynamicPricingFrm.dynamicPriceLbl.Enabled = False
    'DynamicPricingFrm.dynamicSiteChk.Enabled = False
    'DynamicPricingFrm.dynamicStoreChk.Enabled = False
    
    Me.ItemDescription = item("ItemDescription")
    If item("IsMASKit") = True Then
        Me.lblIsKit.caption = "KIT"
        Me.lblIsKit.Visible = True
    ElseIf DLookup("COUNT(*)", "IM_SalesKitDetail", "ComponentItemCode='" & item("ItemNumber") & "'") <> "0" Then
        Me.lblIsKit.caption = "COMP"
        Me.lblIsKit.Visible = True
    Else
        Me.lblIsKit.Visible = False
    End If
    Me.prodLine = item("ProductLine")
    Me.lblManufName.caption = Replace(ManufHashPLtoName.item(item("ProductLine").value), "&", "&&")
    Me.PrimaryVendor = Trim$(VendorHash(item("PrimaryVendor").value))
    Me.EPN = Nz(item("EPN"))
    Me.components = Nz(item("Components"))
'    Me.Weight = item("Weight")
'    Me.Dimensions = item("Dimensions")
    Me.ReplacedBy = Nz(item("ReplacedBy"))
    Me.ReplacementFor = Nz(item("ReplacementFor"))
'    Me.Class = Nz(item("Class"))
    'Me.Barcode = Nz(item("BarCode"))
    'Me.btnLoadBarcodes.caption = "Barcode (" & item("NumBarcodes") & ")"
    Me.btnLoadBarcodes.caption = "Barcode (" & DLookup("COUNT(Barcode)", "InventoryComponentMap INNER JOIN InventoryComponentBarcodes ON InventoryComponentMap.ComponentID=InventoryComponentBarcodes.ComponentID", "ItemID='" & item("ID") & "'") & ")"
    Me.POComment = Nz(item("POComment"))
    Me.StdPack = item("StdPack")
    If item("IsMASKit") = 0 Then
        Me.StdCost = Format(item("StdCost"), "Currency")
    Else
        If item("StdCost") = 0 Then
            Me.StdCost = Format(DLookup("SUM(QuantityPerAssembly*StdCost)", "IM_SalesKitDetail INNER JOIN InventoryMaster ON IM_SalesKitDetail.ComponentItemCode=InventoryMaster.ItemNumber", "SalesKitNo='" & item("ItemNumber") & "'"), "Currency")
        Else
            Me.StdCost = Format(item("StdCost"), "Currency")
        End If
    End If
    Me.AvgCost = Format(item("AvgCost"), "Currency")
    'Me.InventoriedDate = Format(item("DateLastMfrCostChange"), "Short Date")
    Me.StdPrice = Format(item("StdPrice"), "Currency")
    Me.ListPrice = Format(item("ListPrice"), "Currency")
    fillStorePrices item
    fillInetPrices item
    fillEbayPrices item
    
    Dim stockAlertsOutstanding As Long
    stockAlertsOutstanding = DLookup("COUNT(*)", "CustomerAlerts", "AlertType='stock' AND SentNotification=0 AND ItemNumber='" & item("ItemNumber") & "'")
    If stockAlertsOutstanding = 0 Then
        Me.btnViewStockBOAlerts.caption = "On BO"
    Else
        Me.btnViewStockBOAlerts.caption = "On BO (" & stockAlertsOutstanding & ")"
    End If
    
    fillSalesHistory item("ItemNumber"), IsItemDSable(item("ItemStatus"))
    fillOrderPoints item("ItemNumber")
    fillWeightDimensions item("ItemNumber"), item("ID")
    fillInventoriedDate item("ID")

    'Me.QtyOrdered = item("QtyOrdered")
    Dim poRst As ADODB.Recordset
    Set poRst = DB.retrieve("SELECT SUM(Quantity), COUNT(*) FROM PurchaseOrders INNER JOIN PurchaseOrderLines ON PurchaseOrders.ID=PurchaseOrderLines.HeaderID WHERE Exported=0 AND ItemNumber='" & Me.ItemNumber & "'")
    If poRst(1) = "0" Then
        Me.QtyOrdered = "0"
        'Me.QtyOrdered.Locked = lockedDown
    ElseIf poRst(1) = "1" Then
        Me.QtyOrdered = poRst(0)
        'Me.QtyOrdered.Locked = lockedDown
    Else
        Me.QtyOrdered = poRst(0) & "/" & poRst(1)
        'Me.QtyOrdered.Locked = True
    End If
    
    
    
    If priceCompareEnabled = True Then
        priceCompare
    End If
    
    
    'Me.CurrentOP = Nz(item("OldReorderPoint"))
    'Me.NewOP = Nz(item("NewReorderPoint"))
    'If Me.NewOP <> -1 Then
    '    Me.NewOP.Visible = True
    '    Me.lblNewOP.Visible = True
    '    Me.chkUseNewOP.Visible = True
    '    If Me.NewOP > Me.CurrentOP Then
    '        Me.NewOP.ForeColor = RGB(255, 0, 0)
    '        Me.lblNewOP.ForeColor = RGB(255, 0, 0)
    '    ElseIf Me.NewOP < Me.CurrentOP Then
    '        Me.NewOP.ForeColor = RGB(0, 180, 0)
    '        Me.lblNewOP.ForeColor = RGB(0, 180, 0)
    '    Else
    '        Me.NewOP.ForeColor = RGB(0, 0, 0)
    '        Me.lblNewOP.ForeColor = RGB(0, 0, 0)
    '    End If
    'Else
    '    Me.NewOP.Visible = False
    '    Me.lblNewOP.Visible = False
    'End If
    fillingForm = True
    'Me.chkLockOP = SQLBool(item("LockOP"))
    'Me.chkUseNewOP = SQLBool(item("UpdateOrderPoint"))
    'setBoxWarningFlag item("ItemStatus"), item("BoxID")
    Me.chkTaxExempt = SQLBool(item("TaxExempt"))
    
    setXSheetCheckStatus Me.ItemNumber
    
    Me.chkStock.Enabled = True
    Me.chkDropship.Enabled = True
    Me.chkDCbyUs.Enabled = True
    Me.chkDCbyMfr.Enabled = True
    Me.chkStock = SQLBool(IsItemStocked(item("ItemStatus")))
    Me.chkDropship = SQLBool(IsItemDSable(item("ItemStatus")))
    Me.chkDCbyUs = SQLBool(IsItemDCbyUs(item("ItemStatus")))
    Me.chkDCbyMfr = SQLBool(IsItemDCbyMfr(item("ItemStatus")))
    Me.lblDCDS.caption = ItemStatusHashIDtoStr.item(item("ItemStatus").value)
    Me.chkDropship.Enabled = Not CBool(Me.chkDCbyMfr)
    
    'If item("TruckFreight") = False Then
    '    Me.opShippingType(SHIP_REG) = True
    'ElseIf item("TruckFreightCheap") = False Then
    '    Me.opShippingType(SHIP_TF) = True
    'Else
    '    Me.opShippingType(SHIP_TF_CHEAP) = True
    'End If
    Me.opShippingType(CLng(item("ShippingType"))) = True
    Me.opShippingType(3).ForeColor = IIf(Me.opShippingType(3) = True, RED, BLACK)
'    Me.chkNeedsBox = SQLBool(item("NeedsBox"))
    Me.chkReconditioned = SQLBool(item("IsReconditioned"))
    
    'Me.chkEBayFreeShipping = SQLBool(item("EBayFreeShipping"))
    
    fillingForm = False
    requeryFreightEstimates
    requeryFreightActual
    requeryPriceComparisons
    requeryPOList Nz(item("PrimaryVendor"))

    IMtt.RemoveToolTipFromCtl Me.StdCost
    IMtt.RemoveToolTipFromCtl Me.StdPrice
    IMtt.RemoveToolTipFromCtl Me.DiscountMarkupPriceRate(1)
    IMtt.RemoveToolTipFromCtl Me.IDiscountMarkupPriceRate(1)
    IMtt.RemoveToolTipFromCtl Me.EBayPrice
    IMtt.RemoveToolTipFromCtl Me.NewCost
    Me.StdCost.tag = "N" & IIf(item("IsMASKit") = 0, "N", "K") & IIf(item("StdCost") = 0, "Z", "N") 'tooltip flag Y/N, kit flag N/K, cost flag Z/N
    Me.StdPrice.tag = "N"
    Me.DiscountMarkupPriceRate(1).tag = "N"
    Me.IDiscountMarkupPriceRate(1).tag = "N"
    Me.EBayPrice.tag = "N"
    Me.MAPP.tag = "N"
    Me.NewCost.tag = "N"
    
    IMtt.AddToolTipToCtl Me.ItemNumber, "Introduced on: " & Format(item("PartIntroducedDate"), "Short Date")
    Me.currentItemNumber = "Introduced on: " & Format(item("PartIntroducedDate"), "Short Date")
    
    Me.NewCost = IIf(item("TNC") = 0, "", Format(item("TNC"), "Currency"))
    Me.MAPP = IIf(item("MAPP") = 0, "", Format(item("MAPP"), "Currency"))
    
    Dim avgFreight As Double
    avgFreight = GetAvgFreightCost(item("ItemNumber"))
    Me.freightActualAvg = Format(avgFreight, "Currency")
    
    IMtt.AddToolTipToCtl Me.AvgCost, "AvgCost+AvgFreight = " & Format(item("AvgCost") + avgFreight, "Currency")
    Me.CurrentAVGdata = "AvgCost+AvgFreight = " & Format(item("AvgCost") + avgFreight, "Currency")
    
    
    Me.TriadCode = Nz(item("TriadCode"))
    
    Me.btnViewInvQtyTriggers.FontBold = item("HasQtyTrigger")
    
    'Me.GMROI = FormatNumber(Nz(item("GMROI"), "0"), 2, vbTrue, vbFalse, vbFalse)
    
    fillDropshipCostField
    
    Me.lblVendorQuantity.caption = Nz(item("VendorQuantityOnHand"))
    If Me.lblVendorQuantity.caption <> "" Then
        Me.lblVendorQuantity.ToolTipText = "Last checked: " & item("VendorQuantityTimeChecked")
    Else
        Me.lblVendorQuantity.ToolTipText = ""
    End If
    
    fillingForm = True
    Select Case item("FolliesStatus")
        Case Is = 0
            Me.opFolliesStatus(0).value = True
        Case Is = 1
            Me.opFolliesStatus(1).value = True
        Case Is = 2
            Me.opFolliesStatus(2).value = True
        Case Is = 3
            Me.opFolliesStatus(3).value = True
        Case Is = 4
            Me.opFolliesStatus(4).value = True
        Case Else
            MsgBox "Invalid FolliesStatus for item!"
    End Select
    
    
    
    'Tom's Code 9-5-2014 To include searching for Overridden Subtitle Fields...
    Dim PopulateCombo As ADODB.Recordset
    Set PopulateCombo = DB.retrieve("SELECT Subtitle FROM Subtitles ORDER BY Subtitle")
    Dim CustomSubtitlesCount As Integer
    CustomSubtitlesCount = 0
    SubtitleCmb.Clear
    Do
    
    SubtitleCmb.AddItem (PopulateCombo("Subtitle"))
    PopulateCombo.MoveNext
    
    
    Loop Until PopulateCombo.EOF = True Or PopulateCombo.BOF = True
    
    
    
    
    
    
    Dim qry As ADODB.Recordset
    Set qry = DB.retrieve("SELECT ItemNumber,SubtitleID FROM SubtitledItems WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
    Dim SubtitleID As String
    
    If (qry.RecordCount > 0) Then
        SubtitleID = qry("SubtitleID")
        Set qry = DB.retrieve("SELECT Subtitle FROM Subtitles WHERE ID=" & SubtitleID)
     If qry.RecordCount > 0 Then
     subtitleNotClicked = True
     Me.SubTitleChk.value = vbChecked
     If HasPermissionsTo("SubtitleOverride") = True Then
        Me.SubtitleCmb.Enabled = True
        DeleteSelectedSubtitleBtn.Enabled = True
        AddSelectedSubtitleBtn.Enabled = True
     End If
     
     Me.SubtitleCmb.Text = qry("Subtitle")
     Else
     'MsgBox "This particular item has an attached subtitle that was removed from the database. Relisting/Revising this item may lose it's current subtitle or end up with the default. If this message gets too annoying we can remove it."
     'just give the item the default subtitle...
     Set qry = DB.retrieve("SELECT Subtitle FROM Subtitles WHERE ID=1")
     End If
     
    Else
    subtitleNotClicked = True
     Me.SubTitleChk.value = vbUnchecked
     Me.SubtitleCmb.Enabled = False
     DeleteSelectedSubtitleBtn.Enabled = False
     AddSelectedSubtitleBtn.Enabled = False
     Me.SubtitleCmb.Text = ""
          
     Dim SalesRollingLast365Total As ADODB.Recordset
     Set SalesRollingLast365Total = DB.retrieve("SELECT TotalSalesLast365Days FROM vInventorySpreadsheet WHERE ITEMNUMBER='" & Me.ItemNumber.Text & "'")
     
         If Me.EBayPrice > 150 Or (CInt(SalesRollingLast365Total(0))) > 50 Then
        'check to see if this item has been stamped with a sub. if not, we will add it to the db and set offload flag
        Set qry = DB.retrieve("SELECT SubtitleID FROM SubtitledItems WHERE ItemNumber='" & item("ItemNumber") & "'")
        If qry.RecordCount = 0 Then
        Dim res As Boolean
        
        Dim check As ADODB.Recordset
        Set check = DB.retrieve("SELECT EBayPublished FROM vWebOffloadEBay WHERE ItemNumber='" & item("ItemNumber") & "' AND EBayPublished=1")
        If (check.RecordCount > 0) Then
        
         res = DB.Execute("INSERT INTO SubtitledItems (ItemNumber,SubtitleID) VALUES('" & item("ItemNumber") & "',1)")
         res = DB.Execute("UPDATE vWebOffloadEBay SET EBayOffloadRequired=1 WHERE ItemNumber='" & item("ItemNumber") & "'")
        End If
         
            
        End If
        
        'give this default sub...
        Set qry = DB.retrieve("SELECT SubTitle FROM Subtitles WHERE ID=1")
        Me.SubtitleCmb.Text = qry("Subtitle")
        subtitleNotClicked = True
        Me.SubTitleChk.value = vbChecked
     End If
     'MsgBox "TOTAL SALES: " & CStr((CInt(Me.SalesSO(5)) + CInt(Me.SalesP2(5))))
     
    End If
    
    
'    Dim JetOffload As ADODB.Recordset
'    Set JetOffload = DB.retrieve("SELECT ItemNumber,Price FROM JetOffload WHERE ItemNumber='" & item("ItemNumber") & "' AND ToBeDeleted=0")
'    If JetOffload.RecordCount > 0 Then
'        JetChk.value = vbChecked
'        Dim getDynamicPrice As ADODB.Recordset
'        Set getDynamicPrice = DB.retrieve("SELECT NewPrice FROM PriceWiserItems WHERE ItemNumber='" & item("ItemNumber") & "' AND Enabled=1 AND JetDotCom=1")
'        If getDynamicPrice.RecordCount > 0 And DynamicPricingFrm.ListView1.ListItems(4).Checked = True Then
'         JetPrice.Text = getDynamicPrice("NewPrice")
'
'         If getDynamicPrice("NewPrice") <> JetOffload("Price") And DynamicPricingFrm.ListView1.ListItems(4).Checked = True Then
'            DB.Execute "UPDATE JetOffload SET Price=" & getDynamicPrice("NewPrice") & " WHERE ItemNumber='" & item("ItemNumber") & "'"
'         End If
'        Else
'        JetPrice.Text = JetOffload("Price")
'        End If
'    Else
'        JetChk.value = vbUnchecked
'        JetPrice.Text = ""
'        'JetPrice.Text = EBayPrice.Text
'    End If
    
    
'     DynamicPricingFrm.jetDotComPriceLbl.Caption = "Jet.com Price: $" & Me.JetPrice.Text
'     DynamicPricingFrm.refreshForm
     
'     DynamicPricingFrm.ListView1.ListItems.Clear
'     DynamicPricingFrm.ListView1.Refresh
    
    Dim isDyn As ADODB.Recordset
    Set isDyn = DB.retrieve("SELECT JetDotCom FROM PriceWiserItems WHERE ItemNumber='" & Me.JetPrice.Text & "' AND Enabled=1")
    If isDyn.RecordCount > 0 Then
        If CBool(isDyn("JetDotCom")) = True Then
            Me.JetPrice.BackColor = vbCyan
        Else
            Me.JetPrice.BackColor = vbWhite
        End If
    End If
    Dim jPrice As ADODB.Recordset
    Set jPrice = DB.retrieve("SELECT Price FROM JetOffload WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
    'Me.JetPrice.BackColor = vbWhite
    If jPrice.RecordCount > 0 Then
        Me.JetPrice.Text = Format(jPrice("Price"), "Currency")
    Else
        Me.JetPrice.Text = ""
    End If
    'fillDynamicPricing
    
    
    Dim findOrigPrice As ADODB.Recordset
    Set findOrigPrice = DB.retrieve("SELECT OriginalPrice1 FROM DynamicPricing WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
    If findOrigPrice.RecordCount > 0 Then
    OriginalIDiscountPrice = Format(findOrigPrice("OriginalPrice1"), "Currency")
    Me.IDiscountMarkupPriceRate(1).Text = Format(OriginalIDiscountPrice, "Currency")
    Else
    OriginalIDiscountPrice = Me.IDiscountMarkupPriceRate(1).Text
    End If
    
    OriginalEBayPrice = Me.EBayPrice.Text
    CheckDynamicPriceEnabled
    
    inetSortTotalBtn_Click
    ebaySortTotalBtn_Click
    If Me.orderAscending = False Then
        inetSortTotalBtn_Click
        ebaySortTotalBtn_Click
    End If
    
    
    
    'Tom's Code 4-21-2015 to allow for dynamic product conditions
    Dim condition As ADODB.Recordset
    Set condition = DB.retrieve("SELECT ConditionName FROM ItemConditionLines INNER JOIN ItemConditionMaster ON ItemConditionMaster.ConditionID=ItemConditionLines.ConditionID WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
    If condition.RecordCount > 0 Then
        Me.conditionLbl.caption = condition("ConditionName")
    Else
    
    End If
    
    
    'Tom's Code 5-26-2015 to allow for tooltip on receiving info
    'Now while this is the right way to go about this, MAS sucks so bad it takes
    'literally a minute to execute this block of code. So lets 'hack' job it...
    'Dim masDetails As ADODB.Recordset
    'Set masDetails = MASDB.retrieve("SELECT ItemCode,ReceiptDate FROM IM_ItemCost")
    'If masDetails.RecordCount > 0 Then
    '    IMtt.AddToolTipToCtl QtyOnHand, "Last Receipt: " & CStr(masDetails("ReceiptDate"))
    'Else
    '    IMtt.AddToolTipToCtl QtyOnHand, "N/A"
    'End If
    
    Dim itemRxHist As ADODB.Recordset
    Set itemRxHist = DB.retrieve("SELECT TOP 10 DateExported, Quantity FROM PurchaseOrders LEFT OUTER JOIN PurchaseOrderLines ON PurchaseOrders.ID=PurchaseOrderLines.HeaderID WHERE ItemNumber='" & Me.ItemNumber.Text & "' ORDER BY DateExported DESC")
    If itemRxHist.RecordCount > 0 Then
    Dim hist As String
    hist = Left$("Date Received" & "                                     ", 25) & " - " & "Qty" & vbCrLf & Left$("------------------------------------------------------", 32) & vbCrLf
     Do
        Dim test As String
        test = Left$(itemRxHist("DateExported") & "                    ", 25)
        'MsgBox Len(test)
        hist = hist & test & " - " & CStr(itemRxHist("Quantity")) & vbCrLf
        itemRxHist.MoveNext
     Loop Until itemRxHist.EOF = True Or itemRxHist.BOF = True
     IMtt.AddToolTipToCtl QtyOnHand, hist
    Else
     IMtt.AddToolTipToCtl QtyOnHand, "N/A"
    End If
    
    productTagsLst.Clear
    Dim productTags As ADODB.Recordset
    Set productTags = DB.retrieve("SELECT ProductTag FROM ProductTags WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
    If productTags.RecordCount > 0 Then
        Do
            productTagsLst.AddItem productTags("ProductTag")
            productTags.MoveNext
        Loop Until productTags.EOF = True Or productTags.BOF = True
    End If
    
    
    
    fillingForm = False

    Mouse.Hourglass False
End Sub
Private Sub UncheckDynamicFrm()
    'DynamicPricingFrm.dynamicEBayChk.value = vbUnchecked
    'DynamicPricingFrm.dynamicJetChk.value = vbUnchecked
    'DynamicPricingFrm.dynamicSiteChk.value = vbUnchecked
    'DynamicPricingFrm.dynamicStoreChk.value = vbUnchecked
End Sub

Private Sub fillStorePrices(item As ADODB.Recordset)
    Dim i As Long
    For i = 1 To 5
        Me.DiscountMarkupPriceRate(i) = Format(item("DiscountMarkupPriceRate" & i), "Currency")
        Me.BreakQty(i).Enabled = True
    Next i
    Me.BreakQty(1).Enabled = False
    For i = 2 To 5
        If item("BreakQty" & i - 1) = 999999 Then
            Me.BreakQty(i) = item("BreakQty" & i - 1)
            Disable Me.DiscountMarkupPriceRate(i)
            i = i + 1
            While i <= 5
                Me.BreakQty(i) = ""
                Me.BreakQty(i).Enabled = False
                Disable Me.DiscountMarkupPriceRate(i)
                i = i + 1
            Wend
        Else
            Me.BreakQty(i) = item("BreakQty" & i - 1) + 1
            Me.BreakQty(i - 1).Enabled = False
            Enable Me.DiscountMarkupPriceRate(i)
        End If
    Next i
    
    If priceFieldIsLessThanPriceField(Me.DiscountMarkupPriceRate(1), Me.StdCost) Then
        Me.DiscountMarkupPriceRate(1).FontBold = True
        Me.DiscountMarkupPriceRate(1).ForeColor = RED
    Else
        Me.DiscountMarkupPriceRate(1).FontBold = False
        Me.DiscountMarkupPriceRate(1).ForeColor = BLACK
    End If
    'DynamicPricingFrm.storePriceLbl.Caption = "Current Store Price: " & Me.DiscountMarkupPriceRate(1)
    
End Sub

Private Sub fillInetPrices(item As ADODB.Recordset)
    Dim i As Long
    For i = 1 To 5
        Me.IDiscountMarkupPriceRate(i) = Format(item("IDiscountMarkupPriceRate" & i), "Currency")
        Me.IBreakQty(i).Enabled = True
    Next i
    Me.IBreakQty(1).Enabled = False
    For i = 2 To 5
        If item("IBreakQty" & i - 1) = 999999 Or item("IBreakQty" & i - 1) = "99999999" Then
            Me.IBreakQty(i) = item("IBreakQty" & i - 1)
            Disable Me.IDiscountMarkupPriceRate(i)
            i = i + 1
            While i <= 5
                Me.IBreakQty(i) = ""
                Me.IBreakQty(i).Enabled = False
                Disable Me.IDiscountMarkupPriceRate(i)
                i = i + 1
            Wend
        Else
            Me.IBreakQty(i) = item("IBreakQty" & i - 1) + 1
            Me.IBreakQty(i - 1).Enabled = False
            Enable Me.IDiscountMarkupPriceRate(i)
        End If
    Next i
    
    If Me.IDiscountMarkupPriceRate(1) <> "$0.00" Then
        'DynamicPricingFrm.tpSitePriceLbl.Caption = "Current TP-Site Price: " & Me.IDiscountMarkupPriceRate(1)
        If priceFieldIsLessThanPriceField(Me.IDiscountMarkupPriceRate(1), Me.StdCost) Then
            Me.IDiscountMarkupPriceRate(1).FontBold = True
            Me.IDiscountMarkupPriceRate(1).ForeColor = RED
        Else
            Me.IDiscountMarkupPriceRate(1).FontBold = False
            Me.IDiscountMarkupPriceRate(1).ForeColor = BLACK
        End If
    Else
        Me.IDiscountMarkupPriceRate(1).FontBold = False
        Me.IDiscountMarkupPriceRate(1).ForeColor = BLACK
    End If
End Sub

Private Sub fillEbayPrices(item As ADODB.Recordset)
    fillingForm = True
    Me.EBayPrice = Format(item("EBayPrice"), "Currency")
    'DynamicPricingFrm.ebayPriceLbl.Caption = "Current EBay Price: " & Me.EBayPrice
    'DynamicPricingFrm.jetDotComPriceLbl.Caption = "Current Jet.com Price: " & Me.EBayPrice

    If Me.EBayPrice <> "$0.00" Then
        If priceFieldIsLessThanPriceField(Me.EBayPrice, Me.StdCost) Then
            Me.EBayPrice.FontBold = True
            Me.EBayPrice.ForeColor = RED
        Else
            Me.EBayPrice.FontBold = False
            Me.EBayPrice.ForeColor = BLACK
        End If
    Else
        Me.EBayPrice.FontBold = False
        Me.EBayPrice.ForeColor = BLACK
    End If
    
    Me.chkEBayAllowBestOffer = SQLBool(item("EBayAllowBestOffer"))
    If Me.chkEBayAllowBestOffer Then
        'Me.EBayBestOfferAutomationLevel(0) = IIf(item("EBayBestOfferAutoDecline") = 0, "", Format(item("EBayBestOfferAutoDecline"), "Currency"))
        'Me.EBayBestOfferAutomationLevel(1) = IIf(item("EBayBestOfferAutoAccept") = 0, "", Format(item("EBayBestOfferAutoAccept"), "Currency"))
        'Enable Me.EBayBestOfferAutomationLevel(0)
        'Enable Me.EBayBestOfferAutomationLevel(1)
        Me.EBayBestOfferAutoAccept = IIf(item("EBayBestOfferAutoAccept") = 0, "", Format(item("EBayBestOfferAutoAccept"), "Currency"))
        Enable Me.EBayBestOfferAutoAccept
    Else
        'Me.EBayBestOfferAutomationLevel(0) = ""
        'Me.EBayBestOfferAutomationLevel(1) = ""
        'Disable Me.EBayBestOfferAutomationLevel(0)
        'Disable Me.EBayBestOfferAutomationLevel(1)
        Me.EBayBestOfferAutoAccept = ""
        Disable Me.EBayBestOfferAutoAccept
    End If
    
    fillingForm = False
End Sub

Private Sub fillSalesHistory(item As String, dsable As Boolean)
    Dim rst As ADODB.Recordset
    
    If Me.cmbWhse = WHSE_ALL Then
        Set rst = DB.retrieve("SELECT QtyOnHand, QtyOnSO, QtyOnPO, QtyOnBO FROM vAllWhseQty WHERE ItemNumber='" & item & "'")
    ElseIf Me.cmbWhse = WHSE_ACTUAL Then
        Set rst = DB.retrieve("SELECT QtyOnHand, QtyOnSO, QtyOnPO, QtyOnBO FROM vActualWhseQty WHERE ItemNumber='" & item & "'")
    Else
        Set rst = DB.retrieve("SELECT QuantityOnHand, QuantityOnSalesOrder, QuantityOnPurchaseOrder, QuantityOnBackOrder FROM InventoryQuantities WHERE ItemNumber='" & item & "' AND WhseCode='" & DLookup("WarehouseCode", "IM_Warehouse", "WarehouseDesc='" & Me.cmbWhse & "'") & "'")
    End If
    If rst.EOF Then
        Me.QtyOnHand = "0"
        Me.QtyOnSO = "0"
        Me.QtyOnPO = "0"
        Me.QtyBackorder = "0"
    Else
        Me.QtyOnHand = rst(0)
        Me.QtyOnSO = rst(1)
        Me.QtyOnPO = rst(2)
        Me.QtyBackorder = rst(3)
    End If
    rst.Close
    
    Dim i As Long
    For i = 0 To Me.SalesP2.UBound
        Me.SalesP2(i) = "0"
        Me.SalesSO(i) = "0"
        Me.Returns(i) = "0"
        Me.DropShips(i) = "0"
    Next i
    
    Dim whseClause As String
    If Me.cmbWhse = WHSE_ALL Then
        whseClause = ""
    ElseIf Me.cmbWhse = WHSE_ACTUAL Then
        whseClause = "WhseCode NOT LIKE '9%'"
    Else
        whseClause = "WhseCode='" & DLookup("WarehouseCode", "IM_Warehouse", "WarehouseDesc='" & Me.cmbWhse & "'") & "'"
    End If
    
    Dim periodTypeClause As String
    'If Me.opPeriodType(0) Then
    If Me.btnSalesHistoryPeriodTypeToggle.caption = "Rolling" Then 'this is reverse
        periodTypeClause = "AND PeriodType=0" 'static
    Else
        periodTypeClause = "AND PeriodType=1" 'rolling
    End If
    'MsgBox whseClause
    Set rst = DB.retrieve("SELECT Period, SUM(P2SalesAmount), SUM(SOSalesAmount), SUM(ReturnsAmount), SUM(DropshipsAmount) FROM InventoryStatistics WHERE ItemNumber='" & item & "'" & periodTypeClause & IIf(whseClause = "", "", " AND " & whseClause) & " GROUP BY Period")
    While Not rst.EOF
        If Me.btnSalesHistoryColumnToggle.caption = "Combine" Then 'also reverse
            Me.SalesP2(rst(0)) = rst(1)
            Me.SalesSO(rst(0)) = rst(2)
        Else
            Me.SalesSO(rst(0)) = rst(1) + rst(2)
        End If
        Me.Returns(rst(0)) = rst(3)
        Me.DropShips(rst(0)) = rst(4)
        rst.MoveNext
    Wend
    rst.Close
    
'    '2008/11/25 - show d/s stats for all items, even if not d/s-able
'    If Left(item, 3) <> "XXX" And Left(item, 3) <> "ZZZ" Then
'    'If dsable And Left(item, 3) <> "XXX" And Left(item, 3) <> "ZZZ" Then
'        Dim zzz As String
'        zzz = ZZZifyItemNumber(item)
'        Set rst = DB.retrieve("SELECT Period, SUM(P2SalesAmount)+SUM(SOSalesAmount) FROM InventoryStatistics WHERE ItemNumber='" & zzz & "'" & periodTypeClause & IIf(whseClause = "", "", " AND " & whseClause) & " GROUP BY Period")
'        While Not rst.EOF
'            Me.DropShips(rst(0)) = rst(1)
'            rst.MoveNext
'        Wend
'        rst.Close
'    End If
    
    For i = 0 To Me.SalesP2.UBound
        Me.Returns(i).Visible = CBool(Me.Returns(i) <> "0")
        Me.DropShips(i).Visible = CBool(Me.DropShips(i) <> "0")
    Next i
    
    Set rst = Nothing
End Sub

Private Sub fillOrderPoints(item As String)
    '2010-02-17 - order points are fucking awful, this should be totally redesigned
    fillingForm = True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT WarehouseID, StockHere, OrderPoint, LockOrderPoint FROM InventoryLocationInfo WHERE ItemNumber='" & item & "'")
    Me.OrderPointStore = "?"
    Me.chkOrderPointLockStore = 0
    Me.OrderPointWhse = "?"
    Me.chkOrderPointLockWhse = 0
    Me.OrderPointTotal = "?"
    While Not rst.EOF
        If rst("WarehouseID") = 1 Then
            Me.chkStockedHereStore = SQLBool(rst("StockHere"))
            Me.OrderPointStore = Nz(rst("OrderPoint"), "--")
            Me.chkOrderPointLockStore = SQLBool(rst("LockOrderPoint"))
            'Me.OrderPointStore.Enabled = CBool(Me.chkStockedHereStore)
            'Me.chkOrderPointLockStore.Enabled = CBool(Me.chkStockedHereStore)
        ElseIf rst("WarehouseID") = 5 Then
            Me.chkStockedHereWhse = SQLBool(rst("StockHere"))
            Me.OrderPointWhse = Nz(rst("OrderPoint"), "--")
            Me.chkOrderPointLockWhse = SQLBool(rst("LockOrderPoint"))
            'Me.OrderPointWhse.Enabled = CBool(Me.chkStockedHereWhse)
            'Me.chkOrderPointLockWhse.Enabled = CBool(Me.chkStockedHereWhse)
        Else
            'skip for now
        End If
        rst.MoveNext
    Wend
    rst.Close
    
    If Me.OrderPointStore = "?" Then
        DB.Execute "INSERT INTO InventoryLocationInfo ( ItemNumber, WarehouseID, StockHere ) VALUES ( '" & item & "', 1, 0 )"
        Me.OrderPointStore = "0"
        Me.chkStockedHereWhse = 0
    End If
    If Me.OrderPointWhse = "?" Then
        DB.Execute "INSERT INTO InventoryLocationInfo ( ItemNumber, WarehouseID, StockHere ) VALUES ( '" & item & "', 5, 1 )"
        Me.chkStockedHereWhse = 1
        Me.OrderPointWhse = "0"
    End If
    'Me.OrderPointTotal = CLng(Me.OrderPointStore) + CLng(Me.OrderPointWhse)
    fillTotalOrderPoint
    fillingForm = False
End Sub

Private Sub fillTotalOrderPoint()
    Dim vals As Variant, valsIter As Variant
    vals = Array(Me.OrderPointStore.Text, Me.OrderPointWhse.Text)
    
    Dim sum As Long
    sum = 0
    For Each valsIter In vals
        'sum = sum + IIf(CStr(valsIter) = "--" Or CStr(valsIter) = "?", CLng(0), CLng(valsIter))
        If CStr(valsIter) = "--" Then
            'zero, skip
        ElseIf CStr(valsIter) = "?" Then
            'this shouldn't happen, should have been caught already
        Else
            sum = sum + CLng(valsIter)
        End If
    Next valsIter
    
    Me.OrderPointTotal = sum
End Sub

Private Sub fillInventoriedDate(itemID As Long)
    Me.InventoriedDate = ""
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT MIN(LastInventoriedDate) FROM LocationContents WHERE ComponentID IN (SELECT ComponentID FROM InventoryComponentMap WHERE ItemID=" & itemID & ")")
    If Not rst.EOF Then
        If Nz(rst(0)) <> "" Then
            Me.InventoriedDate = Format(rst(0), "M/d/yy")
        End If
    End If
End Sub

Private Sub fillWeightDimensions(item As String, itemID As Long)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, Weight, Length, Width, Height FROM InventoryComponents WHERE ID IN (SELECT ComponentID FROM InventoryComponentMap WHERE ItemID=" & itemID & ")")
    If rst.RecordCount = 0 Then
        'ADD - this should rarely happen
        'search first?
        DB.Execute "INSERT INTO InventoryComponents ( Component ) VALUES ( '" & item & "' )"
        Dim cid As String
        cid = DLookup("@@IDENTITY", "InventoryComponents")
        DB.Execute "INSERT INTO InventoryComponentMap ( ItemID, ItemNumber, ComponentID ) VALUES ( " & itemID & ", '" & item & "', " & cid & " )"
        Me.Weight = "0"
        Me.Dimensions = "0x0x0"
        Enable Me.Weight
        Enable Me.Dimensions
    ElseIf rst.RecordCount = 1 Then
        'SINGLE
        Me.Weight = rst("Weight")
        Me.Dimensions = rst("Length") & "x" & rst("Width") & "x" & rst("Height")
        Enable Me.Weight
        Enable Me.Dimensions
    Else
        'MULTI
        Me.Weight = "multi-box"
        Me.Dimensions = "see extended"
        Disable Me.Weight
        Disable Me.Dimensions
    End If
    
    Dim dimProblems As PerlArray
    Set dimProblems = New PerlArray
    While Not rst.EOF
        Dim cube As Long
        'arithmetic round, not Round()
        cube = Fix(CDbl(rst("Length")) + 0.5) * Fix(CDbl(rst("Width")) + 0.5) * Fix(CDbl(rst("Height")) + 0.5)
        If cube > 5184 Then
            Dim dimWt As Double
            dimWt = cube / 166
            If dimWt > CDbl(rst("Weight")) Then
                'Me.lblDimensionalWeight.Caption = Round(dimWt, 1)
                dimProblems.Push CStr(Round(dimWt, 1))
            End If
        End If
        Dim lengthAndGirth As Long
        lengthAndGirth = Fix(CDbl(rst("Length")) + 0.5) + (2 * Fix(CDbl(rst("Width")) + 0.5)) + (2 * Fix(CDbl(rst("Height")) + 0.5))
        If lengthAndGirth > 135 Then
            dimProblems.Push "$55"
        End If
        rst.MoveNext
    Wend
    If dimProblems.Scalar > 0 Then
        Me.lblDimensionalWeight = dimProblems.Join(" ")
    Else
        Me.lblDimensionalWeight = ""
    End If
    
    rst.Close
    Set rst = Nothing
End Sub

Private Sub fillDropshipCostField()
    Mouse.Hourglass True
    IMtt.RemoveToolTipFromCtl Me.DropshipCost
    
    Dim costAmt As String
    costAmt = Replace(Replace(Me.StdCost, "$", ""), ",", "")
    
    If Me.lblIsKit = "KIT" And Me.lblIsKit.Visible = True And costAmt = "" Then
        MsgBox "KIT CONTENTS ARE MESSED UP!"
        Mouse.Hourglass False
        Exit Sub
    End If
    
    Dim dsCost As Variant, ppfCost As Variant, pppCost As Variant, fvfCost As Variant, fvfDisc As Variant
    
    Select Case True
        Case Is = Me.opDropshipCostType(0), Me.opDropshipCostType(1)
            'commercial
            'cost + commercial dropship fees + ebay fees
            Dim totalWeight As Double
            If CBool(InStr(Me.Weight, "box")) Then
                totalWeight = CDbl(DLookup("SUM(CAST(Weight as float))", "InventoryComponents", "ID IN (SELECT ComponentID FROM InventoryComponentMap WHERE ItemID=" & Me.LoadedItemID & ")"))
            Else
                totalWeight = Me.Weight
            End If
            dsCost = CalculateDropshipChargesFor(Me.ItemNumber, 0, costAmt, totalWeight)
            Select Case True
                Case Is = Me.opDropshipCostType(0)
                    'cost + commercial dropship fees
                    Me.DropshipCost = Format(costAmt + dsCost(0), "Currency")
                    IMtt.AddToolTipToCtl Me.DropshipCost, TrimWhitespace(CStr(dsCost(1)), False, True, False, True)
                    Me.CurrentDSCostdata = TrimWhitespace(CStr(dsCost(1)), False, True, False, True)
                Case Is = Me.opDropshipCostType(1)
                    'cost + commercial dropship fees + paypal fees + ebay fees
                    ppfCost = EBay.EBayCalculatePayPalFlatFee(IIf(Me.EBayPrice <> Format(0, "Currency"), Me.EBayPrice, Me.IDiscountMarkupPriceRate(1)))
                    pppCost = EBay.EBayCalculatePayPalPercentageFee(IIf(Me.EBayPrice <> Format(0, "Currency"), Me.EBayPrice, Me.IDiscountMarkupPriceRate(1)))
                    fvfCost = EBay.EBayCalculateFinalValueFee(IIf(Me.EBayPrice <> Format(0, "Currency"), Me.EBayPrice, Me.IDiscountMarkupPriceRate(1)))
                    fvfDisc = EBay.EBayCalculateFinalValueDiscount(fvfCost)
                    Me.DropshipCost = Format(costAmt + ppfCost + pppCost + fvfCost - fvfDisc + CDec(dsCost(0)), "Currency")
                    IMtt.AddToolTipToCtl Me.DropshipCost, "Cost (" & Me.StdCost & ") + " & _
                                                          "PayPal Flat Fee (" & Format(ppfCost, "Currency") & ") + " & _
                                                          "PayPal Percentage Fee (" & Format(pppCost, "Currency") & ") + " & _
                                                          "EBay Final Value Fee (" & Format(fvfCost, "Currency") & ") - " & _
                                                          "EBay Final Value Discount (-" & Format(fvfDisc, "Currency") & ") + " & _
                                                          "D/S Fees (" & Format(dsCost(0), "Currency") & ")" & _
                                                          vbCrLf & vbCrLf & _
                                                          "D/S Fee breakdown:" & vbCrLf & TrimWhitespace(CStr(dsCost(1)), False, True, False, True)
                     Me.CurrentDSCostdata = "Cost (" & Me.StdCost & ") + " & _
                                                          "PayPal Flat Fee (" & Format(ppfCost, "Currency") & ") + " & _
                                                          "PayPal Percentage Fee (" & Format(pppCost, "Currency") & ") + " & _
                                                          "EBay Final Value Fee (" & Format(fvfCost, "Currency") & ") - " & _
                                                          "EBay Final Value Discount (-" & Format(fvfDisc, "Currency") & ") + " & _
                                                          "D/S Fees (" & Format(dsCost(0), "Currency") & ")" & _
                                                          vbCrLf & vbCrLf & _
                                                          "D/S Fee breakdown:" & vbCrLf & TrimWhitespace(CStr(dsCost(1)), False, True, False, True)
            End Select
        Case Is = Me.opDropshipCostType(2)
            'cost + paypal fees + ebay fees
            ppfCost = EBay.EBayCalculatePayPalFlatFee(IIf(Me.EBayPrice <> Format(0, "Currency"), Me.EBayPrice, Me.IDiscountMarkupPriceRate(1)))
            pppCost = EBay.EBayCalculatePayPalPercentageFee(IIf(Me.EBayPrice <> Format(0, "Currency"), Me.EBayPrice, Me.IDiscountMarkupPriceRate(1)))
            fvfCost = EBay.EBayCalculateFinalValueFee(IIf(Me.EBayPrice <> Format(0, "Currency"), Me.EBayPrice, Me.IDiscountMarkupPriceRate(1)))
            fvfDisc = EBay.EBayCalculateFinalValueDiscount(fvfCost)
            Me.DropshipCost = Format(costAmt + ppfCost + pppCost + fvfCost - fvfDisc, "Currency")
            IMtt.AddToolTipToCtl Me.DropshipCost, "Cost (" & Me.StdCost & ") + " & _
                                                  "PayPal Flat Fee (" & Format(ppfCost, "Currency") & ") + " & _
                                                  "PayPal Percentage Fee (" & Format(pppCost, "Currency") & ") + " & _
                                                  "EBay Final Value Fee (" & Format(fvfCost, "Currency") & ") - " & _
                                                  "EBay Final Value Discount (-" & Format(fvfDisc, "Currency") & ")"
            Me.CurrentDSCostdata = "Cost (" & Me.StdCost & ") + " & _
                                                  "PayPal Flat Fee (" & Format(ppfCost, "Currency") & ") + " & _
                                                  "PayPal Percentage Fee (" & Format(pppCost, "Currency") & ") + " & _
                                                  "EBay Final Value Fee (" & Format(fvfCost, "Currency") & ") - " & _
                                                  "EBay Final Value Discount (-" & Format(fvfDisc, "Currency") & ")"
        Case Else
            Err.Raise "123"
    End Select
    Mouse.Hourglass False
End Sub

Private Function addToXSheet(item As String, Optional Remove As Boolean = False) As Boolean
    If Not fillingForm Then
        Dim pos As Long
        If Not Remove Then
            If bsearch(item, xsheetarray) = -1 Then
                pos = lsearch("", xsheetarray)
                If pos = -1 Then
                    ReDim Preserve xsheetarray(UBound(xsheetarray) * 2) As String
                    pos = lsearch("", xsheetarray)
                    xsheetarray(pos) = item
                Else
                    xsheetarray(pos) = item
                    addToXSheet = True
                End If
            Else
                addToXSheet = False
            End If
        Else
            pos = lsearch(item, xsheetarray)
            If pos <> -1 Then
                xsheetarray(pos) = ""
                addToXSheet = True
            End If
        End If
    End If
End Function

Private Sub setXSheetCheckStatus(item As String)
    If lsearch(item, xsheetarray) <> -1 Then
        Me.chkAddToXSheet = 1
    Else
        Me.chkAddToXSheet = 0
    End If
End Sub

Private Function lockForm()
    lockFormCtls True
    Me.NewCost.Locked = True
    Me.MAPP.Locked = True
    Me.QtyOrdered.Locked = True
    'Me.NewOP.Locked = True
    Me.OrderPointStore.Locked = True
    Me.OrderPointWhse.Locked = True
    Me.POComment.Locked = True
    'Me.btnEditLineCode.Enabled = False
    Me.btnAddLineCode.Enabled = False
    Me.btnOptionsDialog.Enabled = False
    Me.btnAddNewItem.Enabled = False
    Me.btnDeleteItem.Enabled = False
    Me.btnSuggestQtyToOrder.Enabled = False
    Me.btnViewInvQtyTriggers.Enabled = False
    Me.btnViewPOs.Enabled = False
    'Me.chkFrame(1).Enabled = False
    Me.TriadCode.Enabled = False
End Function

Private Sub lockFormCtls(tf As Boolean)
    Me.frameInstorePrice.Enabled = Not tf
    Me.frameInternetPrice.Enabled = Not tf
    Me.FrameEBayPrice.Enabled = Not tf
    Me.FrameMiscPrice.Enabled = Not tf
    Me.StdPrice.Locked = tf
    Me.ItemDescription.Locked = tf
    Me.ListPrice.Locked = tf
    Me.PrimaryVendor.Locked = tf
    Me.prodLine.Locked = tf
    Me.StdCost.Locked = tf
    'Me.CurrentOP.Locked = tf
    Me.OrderPointStore.Locked = tf
    Me.OrderPointWhse.Locked = tf
    Me.FrameOPLS.Enabled = Not tf
    Me.EPN.Locked = tf
    Me.ReplacedBy.Locked = tf
    Me.ReplacementFor.Locked = tf
    Me.StdPack.Locked = tf
    Me.chkFrame(0).Enabled = Not tf 'item status toggles
    Me.btnTNCforItem.Enabled = Not tf
    Me.btnTNCforLine.Enabled = Not tf
    Me.btnTNCSwap.Enabled = Not tf
    Me.btnMAPPtoWP.Enabled = Not tf
    'Me.btnMAPPtoWPforLine.Enabled = Not tf
    Me.btnCopyInstore.Enabled = Not tf
    Me.btnCopyInternet.Enabled = Not tf
    Me.chkTaxExempt.Enabled = Not tf
    'Me.chkMetaNoIndex.Enabled = Not tf
End Sub

'Private Sub setBoxWarningFlag(Optional StatusBitString As String = "", Optional Box As String = "")
'    Dim isDC As Boolean
'    isDC = IsItemCompletelyDC(IIf(StatusBitString = "" Or Box = "", Me.ItemNumber, StatusBitString))
'    If isDC Then
'        Me.lblInvalidWtDims.Caption = ""
'    Else
'        If Box = "" Then
'            Me.lblInvalidWtDims.Caption = IIf(DLookup("BoxID", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'") <= 0, "!", "")
'        Else
'            Me.lblInvalidWtDims.Caption = IIf(Box <= 0, "!", "")
'        End If
'    End If
'End Sub

Private Sub setAlertsButton(Optional onoff As Long = -1)
    'this function expects a -1 (don't know), 0 (off), 1 (on)
    If onoff = -1 Then
        'this should be a general order of probability of occurrence
        If CountInventoryQuantityTriggers(1, "") <> 0 Then
            Me.btnAlertsPending.Visible = True
        ElseIf CountBeingDCItems() <> 0 Then
            Me.btnAlertsPending.Visible = True
        ElseIf CountPriceDifferent() <> 0 Then
            Me.btnAlertsPending.Visible = True
        ElseIf CountCustomerInStockRequestsAged() <> 0 Then
            Me.btnAlertsPending.Visible = True
        Else
            Me.btnAlertsPending.Visible = False
        End If
    Else
        Me.btnAlertsPending.Visible = CBool(onoff)
    End If
End Sub






'--------------------------
' db updating
'--------------------------
'Private Sub Barcode_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case Is = vbKeyReturn
'            Barcode_LostFocus
'        Case Is = vbKeyDelete
'            Barcode_KeyPress KeyCode
'    End Select
'End Sub
'
'Private Sub Barcode_KeyPress(KeyAscii As Integer)
'    changed = True
'    whichCtl = "Barcode"
'End Sub
'
'Private Sub Barcode_LostFocus()
'    If changed = True Then
'        If Me.barcode = "" Then
'            Me.barcode = GetUPC(Me.ItemNumber)
'        '    MsgBox "Deleting barcodes can only be done through the barcode form!"
'        '    Me.Barcode = GetBarcode(Me.ItemNumber)
'        Else
'            HandleNewBarcode Me.ItemNumber, Me.barcode
'        End If
'        changed = False
'    End If
'End Sub

Private Sub BreakQty_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            BreakQty_LostFocus index
        Case Is = vbKeyDelete
            BreakQty_KeyPress index, KeyCode
    End Select
End Sub

Private Sub BreakQty_KeyPress(index As Integer, KeyAscii As Integer)
    changed = True
    whichCtl = "BreakQty(" & index & ")"
End Sub

Private Sub BreakQty_LostFocus(index As Integer)
    If changed = True And whichCtl <> "BreakQty(1)" Then
        Mouse.Hourglass True
        If IsNumeric(Me.BreakQty(index)) Then
            If CLng(Me.BreakQty(index)) > CLng(Me.BreakQty(index - 1)) Then
                updateInventoryMaster "BreakQty" & index - 1, Me.BreakQty(index) - 1, Me.ItemNumber, ""
                Me.BreakQty(index - 1).Enabled = False
                If index < 5 Then
                    If Me.BreakQty(index + 1) = "" Then
                        updateInventoryMaster "BreakQty" & index, "999999", Me.ItemNumber, ""
                        Me.BreakQty(index + 1) = "999999"
                        Me.BreakQty(index + 1).Enabled = True
                    End If
                End If
                Enable Me.DiscountMarkupPriceRate(index)
                changed = False
            Else
                MsgBox "Break qty is less than the break qty before it."
            End If
        ElseIf Me.BreakQty(index) = "" Then
            If Me.DiscountMarkupPriceRate(index).Enabled = False Then
                Me.BreakQty(index) = 999999
            Else
                Me.BreakQty(index) = "999999"
                If index <> 5 Then
                    Me.BreakQty(index + 1) = ""
                    Me.BreakQty(index + 1).Enabled = False
                End If
                Me.DiscountMarkupPriceRate(index) = Format(0, "Currency")
                DiscountMarkupPriceRate_LostFocus index
                Disable Me.DiscountMarkupPriceRate(index)
                updateInventoryMaster "BreakQty" & index - 1, "999999", Me.ItemNumber, ""
'                Dim i As Integer
                updateInventoryMaster "BreakQty" & index, "0", Me.ItemNumber, ""
                If index <> 2 Then
                    Me.BreakQty(index - 1).Enabled = True
                End If
            End If
        Else
            MsgBox "Break qty must be a number."
        End If
        Mouse.Hourglass False
    End If
End Sub

'Private Sub chkLockOP_Click()
'    If Not fillingForm Then
'        updateInventoryMaster "LockOP", Me.chkLockOP, Me.ItemNumber, ""
'    End If
'End Sub
'
'Private Sub chkUseNewOP_Click()
'    If Not fillingForm Then
'        updateInventoryMaster "UpdateOrderPoint", Me.chkUseNewOP, Me.ItemNumber, ""
'    End If
'End Sub

Private Sub chkTaxExempt_Click()
    If Not fillingForm Then
        If vbYes = MsgBox("Are you sure you want to change tax exempt status for this item?", vbYesNo) Then
            updateInventoryMaster "TaxExempt", Me.chkTaxExempt, Me.ItemNumber, ""
        Else
            fillingForm = True
            Me.chkTaxExempt = IIf(Me.chkTaxExempt = 1, 0, 1)
            fillingForm = False
        End If
    End If
End Sub

'Private Sub Class_Click()
'    changed = True
'    Class_LostFocus
'End Sub
'
'Private Sub Class_KeyDown(KeyCode As Integer, Shift As Integer)
'    AutoCompleteKeyDown Me.Class, KeyCode, Shift
'    Select Case KeyCode
'        Case Is = vbKeyReturn
'            Class_LostFocus
'        Case Is = vbKeyDelete
'            Class_KeyPress KeyCode
'    End Select
'End Sub
'
'Private Sub Class_KeyPress(KeyAscii As Integer)
'    AutoCompleteKeyPress Me.Class, KeyAscii
'    changed = True
'    whichCtl = "Class"
'End Sub
'
'Private Sub Class_LostFocus()
'    AutoCompleteLostFocus Me.Class
'    If changed = True Then
'        updateInventoryMaster "Class", Me.Class, Me.ItemNumber, "'"
'        'updateExternalShippingDB Me.ItemNumber
'        'updateShippingDB Me.ItemNumber
'        UpdateASODatabase ItemNumber
'        changed = False
'    End If
'End Sub

Private Sub Components_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            Components_LostFocus
        Case Is = vbKeyDelete
            Components_KeyPress KeyCode
    End Select
End Sub

Private Sub Components_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "components"
End Sub

Private Sub Components_LostFocus()
    If changed = True Then
        If validateComponents(Me.components) Then
            updateInventoryMaster "Components", Me.components, Me.ItemNumber, "'"
            changed = False
        Else
            MsgBox "Components is invalid. Must be less than 255 chars."
            Me.components.SetFocus
        End If
    End If
End Sub

Private Sub Dimensions_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.Dimensions.selStart = 0
                Me.Dimensions.SelLength = Len(Me.Dimensions)
                KeyCode = 0
                Shift = 0
            Else
                Dimensions_KeyPress KeyCode
            End If
        Case Is = vbKeyC
            If Shift = vbCtrlMask Then
                ClipBoard_SetData Me.Dimensions.SelText
            Else
                Dimensions_KeyPress KeyCode
            End If
        Case Is = vbKeyV
            If Shift = vbCtrlMask Then
                Dim newdims As String
                newdims = ProcessTextBoxPaste(Me.Dimensions)
                Dim i As Long, bad As Boolean
                bad = False
                For i = 1 To Len(newdims)
                    Dim c As Long
                    c = Asc(Mid(newdims, i, 1))
                    If c = Asc("*") Or c = Asc("x") Or c = Asc("X") Or LimitToNumbers(CInt(c), True, True) Then
                        'ok
                    Else
                        bad = True
                        Exit For
                    End If
                Next i
                If bad Then
                    'msgbox?
                Else
                    Me.Dimensions = newdims
                    changed = True
                    whichCtl = "Dimensions"
                End If
            Else
                Dimensions_KeyPress KeyCode
            End If
        Case Is = vbKeyReturn
            Dimensions_LostFocus
        Case Is = vbKeyDelete
            Dimensions_KeyPress KeyCode
    End Select
End Sub

Private Sub Dimensions_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("*") Or KeyAscii = Asc("x") Or KeyAscii = Asc("X") Then
        'ok
    Else
        KeyAscii = LimitToNumbers(KeyAscii, True, True)
    End If
    If KeyAscii <> 0 Then
        changed = True
        whichCtl = "Dimensions"
    End If
End Sub

Private Sub Dimensions_LostFocus()
    If changed = True Then
        Dim dims As String
        dims = validateDimensions(Me.Dimensions)
        If Left(dims, 5) <> "ERROR" Then
            Me.Dimensions = dims
            'updateInventoryMaster "Dimensions", Me.Dimensions, Me.ItemNumber, "'"
            'If Me.Weight <> "" And Me.Weight <> "0" And Me.Dimensions <> "0x0x0" Then
            '    updateInventoryMaster "BoxID", DetermineBoxSize(Me.Dimensions, Me.Weight), Me.ItemNumber, "'"
            'Else
            '    updateInventoryMaster "BoxID", "-1", Me.ItemNumber, "'"
            'End If
            Dim cid As String
            cid = DLookup("ComponentID", "InventoryComponentMap", "ItemID=" & Me.LoadedItemID)
            Dim dimA As Variant
            dimA = Split(Me.Dimensions, "x")
            DB.Execute "UPDATE InventoryComponents SET Length=" & CStr(dimA(0)) & ", Width=" & CStr(dimA(1)) & ", Height=" & CStr(dimA(2)) & " WHERE ID=" & cid
            'DB.Execute "UPDATE InventoryComponents SET LastModifiedBy=" & GetCurrentUserID() & " WHERE ID=" & cid
            If Me.Dimensions <> "0x0x0" Then
                Dim rst As ADODB.Recordset
                Set rst = DB.retrieve("SELECT PackingType, ShippingBoxID FROM InventoryComponents WHERE ID=" & cid)
                If rst(0) = 0 Then 'brown box
                    If IsNull(rst(1)) Then
                        Dim newBoxSizeID As String
                        newBoxSizeID = DetermineBoxSize2(CStr(dimA(0)), CStr(dimA(1)), CStr(dimA(2)), Me.Weight)
                        If newBoxSizeID <> "-1" Then
                            DB.Execute "UPDATE InventoryComponents SET ShippingBoxID=" & newBoxSizeID & " WHERE ID=" & cid
                        End If
                    End If
                End If
                rst.Close
                Set rst = Nothing
            End If
            'updateExternalShippingDB Me.ItemNumber
            'updateShippingDB Me.ItemNumber
            ExternalComponentDBSync Me.ItemNumber
            changed = False
        ElseIf dims = "ERROR:numeric" Then
            MsgBox "Invalid dimensions, dimensions must be numeric."
            'Me.Dimensions.SetFocus
        ElseIf dims = "ERROR:dimensions" Then
            MsgBox "Invalid dimensions, must have 3 dimensions formatted like: 1x2x3"
            'Me.Dimensions.SetFocus
        ElseIf Left(dims, 5) = "ERROR" Then
            MsgBox "Error validating dimensions"
            'Me.Dimensions.SetFocus
        End If
        'setBoxWarningFlag
    End If
End Sub

Private Sub DiscountMarkupPriceRate_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            DiscountMarkupPriceRate_LostFocus index
        Case Is = vbKeyDelete
            DiscountMarkupPriceRate_KeyPress index, KeyCode
    End Select
End Sub

Private Sub DiscountMarkupPriceRate_KeyPress(index As Integer, KeyAscii As Integer)
    changed = True
    whichCtl = "DiscountMarkupPriceRate(" & index & ")"
End Sub

Private Sub DiscountMarkupPriceRate_LostFocus(index As Integer)
    If changed = True Then
        If validateCurrency(Me.DiscountMarkupPriceRate(index)) Then
            Me.DiscountMarkupPriceRate(index) = Format(Me.DiscountMarkupPriceRate(index), "Currency")
            If index = 1 Then
                LogItemStorePriceChanged Me.ItemNumber, Me.DiscountMarkupPriceRate(1)
                'updateInventoryMaster "OldPrice", "StdPrice", Me.ItemNumber, ""
                If Left(Me.ItemNumber, 3) <> "XXX" And Left(Me.ItemNumber, 3) <> "ZZZ" Then
                    UpdateNotificationEmail "store price", Me.ItemNumber
                End If
            End If
            updateInventoryMaster "DiscountMarkupPriceRate" & index, Me.DiscountMarkupPriceRate(index), Me.ItemNumber, ""
            If index = 1 Then
                updateInventoryMaster "StdPrice", Me.DiscountMarkupPriceRate(1), Me.ItemNumber, ""
                Me.StdPrice = Me.DiscountMarkupPriceRate(1)
                If priceFieldIsLessThanPriceField(Me.StdPrice, Me.StdCost) Then
                    MsgBox "Warning: store price is below standard cost!", vbCritical
                End If
            End If
            If CBool(Me.chkDropship) Then
                If Left(Me.ItemNumber, 3) <> "XXX" And Left(Me.ItemNumber, 3) <> "ZZZ" Then
                    Dim zzz As String
                    zzz = ZZZifyItemNumber(Me.ItemNumber)
                    If zzz <> "" Then
                        updateInventoryMaster "DiscountMarkupPriceRate1", Me.DiscountMarkupPriceRate(1), zzz, ""
                        updateInventoryMaster "StdPrice", Me.StdPrice, zzz, ""
                    End If
                End If
            End If
            changed = False
        Else
            MsgBox "Price is invalid, must be a number."
            Me.DiscountMarkupPriceRate(index).SetFocus
        End If
    End If
End Sub

Private Sub DiscountMarkupPriceRate_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.StdCost = "" Then
        Exit Sub
    End If
    If index = 1 Then
        If Me.DiscountMarkupPriceRate(1).tag = "N" Then
            Me.DiscountMarkupPriceRate(1).tag = "Y"
            Dim rst As ADODB.Recordset
            Set rst = DB.retrieve("SELECT TOP 1 TimeChanged, OldPrice FROM InventoryPriceLog WHERE ItemNumber='" & Me.ItemNumber & "' AND IsStorePrice=1 ORDER BY TimeChanged DESC")
            Dim txt As String
            If rst.EOF Then
                'Me.StdPrice.ToolTipText = "No change info for this item."
                txt = "No price change info for this item"
            Else
                'Me.StdPrice.ToolTipText = "Last Price Change: " & Format(rst("TimeChanged"), "Short Date") & ", was " & Format(rst("OldPrice"), "Currency")
                txt = "Last Price Change: " & Format(rst("TimeChanged"), "Short Date") & ", was " & Format(rst("OldPrice"), "Currency")
            End If
            rst.Close
            Set rst = Nothing
            
            txt = txt & vbCrLf
            If validateCurrency(Me.StdCost) = False Or validateCurrency(Me.DiscountMarkupPriceRate(1)) = False Then
                txt = txt & vbCrLf & "Gross Profit: invalid input field"
            Else
            If Me.StdCost = 0 Or Me.DiscountMarkupPriceRate(1) = 0 Then
                txt = txt & vbCrLf & "Gross Profit: undefined"
            Else
                txt = txt & vbCrLf & "Gross Profit: " & Round(Me.DiscountMarkupPriceRate(1) / Me.StdCost * 100 - 100, 1) & "%"
                txt = txt & vbCrLf & "Pre-tax Incl Amt: $" & Round(Me.DiscountMarkupPriceRate(1) / 1.0635, 2)
            End If
            End If
            IMtt.AddToolTipToCtl Me.DiscountMarkupPriceRate(1), txt
            Me.CurrentDiscountRate = txt
        End If
    End If
    
    
    
    
            'DiscountMarkupPriceRate(Index).BackColor = RGB(0, 255, 255)
            DiscountMarkupPriceRate(index).Refresh
    
            'If Me.HidePopUp = False Then
                'Me.popupTmr.Enabled = True
                'PopUpFrm.Visible = True
                'PopUpFrm.InfoLbl.Caption = Me.CurrentDiscountRate
                'PopUpFrm.ZOrder 0
                'PopUpFrm.SetFocus
                'PopUpFrm.Move Me.Left + Me.frameInstorePrice.Left - 500, Me.Top + Me.frameInstorePrice.Top + 50 + DiscountMarkupPriceRate(Index).Top    'ScaleX(NewCost.Left, vbPixels, vbTwips), ScaleY(NewCost.Top, vbPixels, vbTwips) ', PopUpFrm.width, PopUpFrm.Height
                'PopUpFrm.Refresh
            'End If
End Sub

'Private Sub EbayAuctionPrice_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case Is = vbKeyReturn
'            EbayAuctionPrice_LostFocus
'        Case Is = vbKeyDelete
'            EbayAuctionPrice_KeyPress KeyCode
'    End Select
'End Sub

'Private Sub EbayAuctionPrice_KeyPress(KeyAscii As Integer)
'    changed = True
'    whichCtl = "EbayAuctionPrice"
'End Sub

'Private Sub EbayAuctionPrice_LostFocus()
    'If changed = True Then
    '    If Me.EbayAuctionPrice = "" Or Me.EbayAuctionPrice = 0 Then
    '        updateInventoryMaster "EbayAuctionPrice", "0", Me.ItemNumber, ""
    '        Me.EbayAuctionPrice = ""
    '        changed = False
    '        'auction removing automation goes here
    '    ElseIf validateCurrency(Me.EbayAuctionPrice) Then
    '        Me.EbayAuctionPrice = Format(Me.EbayAuctionPrice, "Currency")
    '        updateInventoryMaster "EbayAuctionPrice", Me.EbayAuctionPrice, Me.ItemNumber, ""
    '        changed = False
    '        'auction starting automation goes here
    '        'Load EBayImportTool
    '        'EBayImportTool.Show
    '        'EBayImportTool.ListItemInEbay Me.ItemNumber, Me.EbayAuctionPrice
    '    Else
    '        MsgBox "Invalid price!"
    '    End If
    'End If
'End Sub

'EbayAuctionReservePrice is read-only, only take input from questions
'
'Private Sub EbayAuctionReservePrice_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case Is = vbKeyReturn
'            EbayAuctionReservePrice_LostFocus
'        Case Is = vbKeyDelete
'            EbayAuctionReservePrice_KeyPress
'    End Select
'End Sub
'
'Private Sub EbayAuctionReservePrice_KeyPress(KeyAscii As Integer)
'    changed = True
'    whichCtl = "EbayAuctionReservePrice"
'End Sub
'
'Private Sub EbayAuctionReservePrice_LostFocus()
'    If changed = True Then
'        If Me.EbayAuctionReservePrice = "" Or Me.EbayAuctionReservePrice = 0 Then
'            updateInventoryMaster "EbayAuctionReservePrice", "0", Me.ItemNumber, ""
'            changed = False
'        If validateCurrency(Me.EbayAuctionReservePrice) Then
'            Me.EbayAuctionPrice = Format(Me.EbayAuctionReservePrice, "Currency")
'            updateInventoryMaster "EbayAuctionReservePrice", Me.EbayAuctionReservePrice, Me.ItemNumber, ""
'            changed = False
'        Else
'            MsgBox "Invalid price!"
'        End If
'    End If
'End Sub

'Private Sub EbayProstorePrice_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case Is = vbKeyReturn
'            EbayProstorePrice_LostFocus
'        Case Is = vbKeyDelete
'            EbayProstorePrice_KeyPress KeyCode
'    End Select
'End Sub

'Private Sub EbayProstorePrice_KeyPress(KeyAscii As Integer)
'    changed = True
'    whichCtl = "EbayProstorePrice"
'End Sub

'Private Sub EbayProstorePrice_LostFocus()
'    If changed = True Then
'        If Me.EbayProstorePrice = "" Or Me.EbayProstorePrice = 0 Then
'            updateInventoryMaster "EbayProstorePrice", "0", Me.ItemNumber, ""
'            Me.EbayProstorePrice = Format("0", "Currency")
'            changed = False
'        ElseIf validateCurrency(Me.EbayProstorePrice) Then
'            Me.EbayProstorePrice = Format(Me.EbayProstorePrice, "Currency")
'            updateInventoryMaster "EbayProstorePrice", Me.EbayProstorePrice, Me.ItemNumber, ""
'            changed = False
'        Else
'            MsgBox "Invalid price!"
'        End If
'    End If
'End Sub


Private Sub EPN_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            EPN_LostFocus
        Case Is = vbKeyDelete
            EPN_KeyPress KeyCode
    End Select
End Sub

Private Sub EPN_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "epn"
End Sub

Private Sub EPN_LostFocus()
    If changed = True Then
        If validateEPN(Me.EPN) Then
            updateInventoryMaster "EPN", Me.EPN, Me.ItemNumber, "'"
            changed = False
        Else
            MsgBox "EPN is invalid. Must be less than 512 chars."
        End If
    End If
End Sub

Private Sub IBreakQty_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            IBreakQty_LostFocus index
        Case Is = vbKeyDelete
            IBreakQty_KeyPress index, KeyCode
    End Select
End Sub

Private Sub IBreakQty_KeyPress(index As Integer, KeyAscii As Integer)
    changed = True
    whichCtl = "IBreakQty(" & index & ")"
End Sub

Private Sub IBreakQty_LostFocus(index As Integer)
    If changed = True And whichCtl <> "IBreakQty(1)" Then
        Mouse.Hourglass True
        If IsNumeric(Me.IBreakQty(index)) Then
            If CLng(Me.IBreakQty(index)) > CLng(Me.IBreakQty(index - 1)) Then
                updateInventoryMaster "IBreakQty" & index - 1, Me.IBreakQty(index) - 1, Me.ItemNumber, ""
                Me.IBreakQty(index - 1).Enabled = False
                If index < 5 Then
                    If Me.IBreakQty(index + 1) = "" Then
                        updateInventoryMaster "IBreakQty" & index, "999999", Me.ItemNumber, ""
                        Me.IBreakQty(index + 1) = "999999"
                        Me.IBreakQty(index + 1).Enabled = True
                    End If
                End If
                Enable Me.IDiscountMarkupPriceRate(index)
                changed = False
            Else
                MsgBox "Break qty is less than the break qty before it."
            End If
        ElseIf Me.IBreakQty(index) = "" Then
            If Me.IDiscountMarkupPriceRate(index).Enabled = False Then
                Me.IBreakQty(index) = 999999
            Else
                Me.IBreakQty(index) = "999999"
                If index <> 5 Then
                    Me.IBreakQty(index + 1) = ""
                    Me.IBreakQty(index + 1).Enabled = False
                End If
                Me.IDiscountMarkupPriceRate(index) = Format(0, "Currency")
                IDiscountMarkupPriceRate_LostFocus index
                Disable Me.IDiscountMarkupPriceRate(index)
                updateInventoryMaster "IBreakQty" & index - 1, "999999", Me.ItemNumber, ""
'                Dim i As Integer
                updateInventoryMaster "IBreakQty" & index, "0", Me.ItemNumber, ""
                If index <> 2 Then
                    Me.IBreakQty(index - 1).Enabled = True
                End If
            End If
        Else
            MsgBox "Break qty must be a number."
        End If
        Mouse.Hourglass False
    End If
End Sub

Private Sub IDiscountMarkupPriceRate_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            IDiscountMarkupPriceRate_LostFocus index
        Case Is = vbKeyDelete
            IDiscountMarkupPriceRate_KeyPress index, KeyCode
    End Select
End Sub

Private Sub IDiscountMarkupPriceRate_KeyPress(index As Integer, KeyAscii As Integer)
    changed = True
    whichCtl = "IDiscountMarkupPriceRate(" & index & ")"
End Sub

Private Sub IDiscountMarkupPriceRate_LostFocus(index As Integer)
    If changed = True Then
        If validateCurrency(Me.IDiscountMarkupPriceRate(index)) Then
            Me.IDiscountMarkupPriceRate(index) = Format(Me.IDiscountMarkupPriceRate(index), "Currency")
            If index = 1 Then
                LogItemWebPriceChanged Me.ItemNumber, Me.IDiscountMarkupPriceRate(1)
                'updateInventoryMaster "OldInetPrice", "IDiscountMarkupPriceRate1", Me.ItemNumber, ""
                If Left(Me.ItemNumber, 3) <> "XXX" And Left(Me.ItemNumber, 3) <> "ZZZ" Then
                    UpdateNotificationEmail "web price", Me.ItemNumber
                End If
            End If
            updateInventoryMaster "IDiscountMarkupPriceRate" & index, Me.IDiscountMarkupPriceRate(index), Me.ItemNumber, ""
            changed = False
            If index = 1 Then
                updateDropshipItemWebPrice Me.ItemNumber, Me.IDiscountMarkupPriceRate(1)
            End If
            If index = 1 And Me.MAPP <> "" Then
                'If priceFieldIsLessThanPriceField(Me.IDiscountMarkupPriceRate(Index), Me.MAPP) Then
                '    If "1" = DLookup("MappViolationWarning", "ProductLine", "ProductLine='" & Me.prodLine & "'") Then
                '        MsgBox "WHOA! The price you entered is less than the MAPP!" & vbCrLf & vbCrLf & "You might want to double check this!"
                '    End If
                'End If
                If Me.IDiscountMarkupPriceRate(1) <> "$0.00" Then
                    If priceFieldIsLessThanPriceField(Me.IDiscountMarkupPriceRate(1), Me.StdCost) Then
                        MsgBox "Warning: internet price is below standard cost!", vbExclamation
                    End If
                End If
            End If
        Else
            MsgBox "Price is invalid, must be a number."
        End If
    End If
End Sub

Private Sub IDiscountMarkupPriceRate_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.StdCost = "" Then
        Exit Sub
    End If
    If index = 1 Then
        If Me.IDiscountMarkupPriceRate(1).tag = "N" Then
            Me.IDiscountMarkupPriceRate(1).tag = "Y"
            Dim rst As ADODB.Recordset
            Set rst = DB.retrieve("SELECT TOP 1 TimeChanged, OldPrice FROM InventoryPriceLog WHERE ItemNumber='" & Me.ItemNumber & "' AND IsStorePrice=0 ORDER BY TimeChanged DESC")
            Dim txt As String
            If rst.EOF Then
                'Me.IDiscountMarkupPriceRate(1).ToolTipText = "No change info for this item."
                txt = "No price change info for this item."
            Else
                'Me.IDiscountMarkupPriceRate(1).ToolTipText = "Last Price Change: " & Format(rst("TimeChanged"), "Short Date") & ", was " & Format(rst("OldPrice"), "Currency")
                txt = "Last Price Change: " & Format(rst("TimeChanged"), "Short Date") & ", was " & Format(rst("OldPrice"), "Currency")
            End If
            rst.Close
            Set rst = Nothing
            
            txt = txt & vbCrLf
            
            If validateCurrency(Me.StdCost) = False Or validateCurrency(Me.IDiscountMarkupPriceRate(1)) = False Then
                txt = txt & vbCrLf & "Gross Profit: invalid input field"
            
            Else
            If Me.StdCost = 0 Or Me.IDiscountMarkupPriceRate(1) = 0 Then
                txt = txt & vbCrLf & "Gross Profit: undefined"
            Else
                txt = txt & vbCrLf & "Gross Profit: " & Round(Me.IDiscountMarkupPriceRate(1) / Me.StdCost * 100 - 100, 1) & "%"
            End If
            End If
            
            IMtt.AddToolTipToCtl Me.IDiscountMarkupPriceRate(1), txt
            Me.CurrentIDiscountRate = txt
        End If
    End If
    
    
                'IDiscountMarkupPriceRate(Index).BackColor = RGB(0, 255, 255)
            IDiscountMarkupPriceRate(index).Refresh
    
            'If Me.HidePopUp = False Then
                'PopUpFrm.Visible = True
                'Me.popupTmr.Enabled = True
                'PopUpFrm.InfoLbl.Caption = Me.CurrentIDiscountRate
                'PopUpFrm.ZOrder 0
                'PopUpFrm.SetFocus
                'PopUpFrm.Move Me.Left + Me.frameInternetPrice.Left - 500, Me.Top + Me.frameInternetPrice.Top + 50 + IDiscountMarkupPriceRate(Index).Top   'ScaleX(NewCost.Left, vbPixels, vbTwips), ScaleY(NewCost.Top, vbPixels, vbTwips) ', PopUpFrm.width, PopUpFrm.Height
                'PopUpFrm.Refresh
            'End If
            
    
    
End Sub

Private Sub EBayPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            EBayPrice_LostFocus
        Case Is = vbKeyDelete
            EBayPrice_KeyPress KeyCode
    End Select
End Sub

Private Sub EBayPrice_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "EBayPrice"
End Sub


Private Sub JetPrice_LostFocus()
'if the text is empty
   'delete from jet.com
'if the text is numerical
   'insert or update jet.com
'if the text is non-numerical
   'error message, set focus, clear textbox
If JetPrice.Text = "" Then
    DB.Execute "UPDATE JetOffload SET ToBeDeleted=1 WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
    Exit Sub
End If

If IsNumeric(JetPrice.Text) = True Then
    Dim jetData As ADODB.Recordset
    Set jetData = DB.retrieve("SELECT ID FROM JetOffload WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
    If jetData.RecordCount > 0 Then
        DB.Execute "UPDATE JetOffload SET OffloadRequired=1, ToBeDeleted=0, Price=" & Me.JetPrice.Text & " WHERE ID=" & jetData("ID")
    Else
        Dim check As Boolean
        check = DB.Execute("INSERT INTO JetOffload (ItemNumber,OffloadRequired,ToBeDeleted,Price) VALUES('" & Me.ItemNumber.Text & "',1,0," & Me.JetPrice.Text & ")")
        If check = False Then
            MsgBox "There was an issue adding this item to the JetOffload SQL Database. Please inform Tom as this shouldn't really happen."
        End If
    End If
Else
    MsgBox "The value in Jet.com's price field is not empty and not a number. Please correct it to continue."
    JetPrice.SetFocus
End If

'Me.fillDynamicPricing

End Sub
Private Sub EBayPrice_LostFocus()
    If changed = True Then
        If validateCurrency(Me.EBayPrice) Then
            Me.EBayPrice = Format(Me.EBayPrice, "Currency")
            LogItemEBayPriceChanged Me.ItemNumber, Me.EBayPrice
            
            'TODO: EBAY PRICE CHANGE EMAIL?
            
            'If Left(Me.ItemNumber, 3) <> "XXX" And Left(Me.ItemNumber, 3) <> "ZZZ" Then
            '    UpdateNotificationEmail "web price", Me.ItemNumber
            'End If
            updateInventoryMaster "EBayPrice", Me.EBayPrice, Me.ItemNumber, ""
            changed = False
            
            updateDropshipItemEBayPrice Me.ItemNumber, Me.EBayPrice
            'If Me.MAPP <> "" Then
            '    If priceFieldIsLessThanPriceField(Me.EBayPrice, Me.MAPP) Then
            '        If "1" = DLookup("MappViolationWarning", "ProductLine", "ProductLine='" & Me.prodLine & "'") Then
            '            MsgBox "WHOA! The price you entered is less than the MAPP!" & vbCrLf & vbCrLf & "You might want to double check this!"
            '        End If
            '    End If
            'End If
            If Me.EBayPrice <> "$0.00" Then
                If priceFieldIsLessThanPriceField(Me.EBayPrice, Me.StdCost) Then
                    MsgBox "Warning: ebay price is below standard cost!", vbExclamation
                End If
                
                'If priceFieldIsLessThanPriceField(Me.EBayPrice, Me.EBayBestOfferAutomationLevel(0)) Then
                '    MsgBox "Warning: ebay price is below best offer conditions! Removing best offer conditions."
                '    updateInventoryMaster "EBayBestOfferAutoDecline", "0", Me.ItemNumber
                '    Me.EBayBestOfferAutomationLevel(0) = ""
                '    updateInventoryMaster "EBayBestOfferAutoAccept", "0", Me.ItemNumber
                '    Me.EBayBestOfferAutomationLevel(1) = ""
                'ElseIf priceFieldIsLessThanPriceField(Me.EBayPrice, Me.EBayBestOfferAutomationLevel(1)) Then
                If priceFieldIsLessThanPriceField(Me.EBayPrice, Me.EBayBestOfferAutoAccept) Then
                    MsgBox "Warning: ebay price is below best offer condition! Removing best offer condition."
                    updateInventoryMaster "EBayBestOfferAutoAccept", "0", Me.ItemNumber
                '    Me.EBayBestOfferAutomationLevel(1) = ""
                    Me.EBayBestOfferAutoAccept = ""
                End If
            End If
            
            Dim SalesRollingLast365Total As ADODB.Recordset
            Dim qry As ADODB.Recordset
     Set SalesRollingLast365Total = DB.retrieve("SELECT TotalSalesLast365Days FROM vInventorySpreadsheet WHERE ITEMNUMBER='" & Me.ItemNumber.Text & "'")
     
         If Me.EBayPrice > 150 Or (CInt(SalesRollingLast365Total(0))) > 50 Then
        'check to see if this item has been stamped with a sub. if not, we will add it to the db and set offload flag
        Set qry = DB.retrieve("SELECT SubtitleID FROM SubtitledItems WHERE ItemNumber='" & Me.ItemNumber & "'")
        If qry.RecordCount = 0 Then
        Dim res As Boolean
        Dim check As ADODB.Recordset
        Set check = DB.retrieve("SELECT EBayPublished FROM vWebOffloadEBay WHERE ItemNumber='" & Me.ItemNumber & "' AND EBayItemID IS NOT NULL")
        If (check.RecordCount > 0) Then
         res = DB.Execute("INSERT INTO SubtitledItems (ItemNumber,SubtitleID) VALUES('" & Me.ItemNumber & "',1)")
         res = DB.Execute("UPDATE vWebOffloadEBay SET EBayOffloadRequired=1 WHERE ItemNumber='" & Me.ItemNumber & "'")
        End If
         
            
        End If
        
        'give this default sub...
        Set qry = DB.retrieve("SELECT SubTitle FROM Subtitles WHERE ID=1")
        Me.SubtitleCmb.Text = qry("Subtitle")
        subtitleNotClicked = True
        Me.SubTitleChk.value = vbChecked
     End If
     'MsgBox "TOTAL SALES: " & CStr((CInt(Me.SalesSO(5)) + CInt(Me.SalesP2(5))))
            
        Else
            MsgBox "Price is invalid, must be a number."
        End If
    End If
End Sub

Private Sub EBayPrice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.StdCost = "" Then
        Exit Sub
    End If
    If Me.EBayPrice.tag = "N" Then
        Me.EBayPrice.tag = "Y"
        
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT TOP 1 TimeChanged, OldPrice FROM InventoryPriceLog WHERE ItemNumber='" & Me.ItemNumber & "' AND IsStorePrice=2 ORDER BY TimeChanged DESC")
        Dim txt As String
        If rst.EOF Then
            txt = "No price change info for this item."
        Else
            txt = "Last Price Change: " & Format(rst("TimeChanged"), "Short Date") & ", was " & Format(rst("OldPrice"), "Currency")
        End If
        rst.Close
        Set rst = Nothing
        
        txt = txt & vbCrLf
        
        If validateCurrency(Me.StdCost.Text) = False Then
            txt = txt & vbCrLf & "Gross Profit: Invalid input field"
            Exit Sub
        End If
        
        If Me.StdCost = 0 Or Me.EBayPrice = 0 Then
            txt = txt & vbCrLf & "Gross Profit: undefined"
        Else
            txt = txt & vbCrLf & "Gross Profit: " & Round(Me.EBayPrice / Me.StdCost * 100 - 100, 1) & "%"
        End If
        
        IMtt.AddToolTipToCtl Me.EBayPrice, txt
        Me.CurrentEBayPrice = txt
        
    End If
    
    
        
            'EBayPrice.BackColor = RGB(0, 255, 255)
            EBayPrice.Refresh
    
            'If Me.HidePopUp = False Then
                'PopUpFrm.Visible = True
                'Me.popupTmr.Enabled = True
                'PopUpFrm.InfoLbl.Caption = Me.CurrentEBayPrice
                'PopUpFrm.ZOrder 0
                'PopUpFrm.SetFocus
                'PopUpFrm.Move Me.Left + Me.FrameEBayPrice.Left - 500, Me.Top + Me.FrameEBayPrice.Top + 50 + StdCost.Top   'ScaleX(NewCost.Left, vbPixels, vbTwips), ScaleY(NewCost.Top, vbPixels, vbTwips) ', PopUpFrm.width, PopUpFrm.Height
                'PopUpFrm.Refresh
            'End If
            
    
    
End Sub

Private Sub chkEBayAllowBestOffer_Click()
    If Not fillingForm Then
        If Me.chkEBayAllowBestOffer Then
            updateInventoryMaster "EBayAllowBestOffer", "1", Me.ItemNumber, ""
            'Enable Me.EBayBestOfferAutomationLevel(0)
            'Enable Me.EBayBestOfferAutomationLevel(1)
            Enable Me.EBayBestOfferAutoAccept
        Else
            updateInventoryMaster "EBayAllowBestOffer", "0", Me.ItemNumber, ""
            'Disable Me.EBayBestOfferAutomationLevel(0)
            'Disable Me.EBayBestOfferAutomationLevel(1)
            Disable Me.EBayBestOfferAutoAccept
        End If
    End If
End Sub

'Private Sub EBayBestOfferAutomationLevel_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case Is = vbKeyReturn
'            EBayBestOfferAutomationLevel_LostFocus Index
'        Case Is = vbKeyDelete
'            EBayBestOfferAutomationLevel_KeyPress Index, KeyCode
'    End Select
'End Sub
'
'Private Sub EBayBestOfferAutomationLevel_KeyPress(Index As Integer, KeyAscii As Integer)
'    changed = True
'    whichCtl = "EBayBestOfferAutomationLevel(" & Index & ")"
'End Sub
'
'Private Sub EBayBestOfferAutomationLevel_LostFocus(Index As Integer)
'    If changed = True Then
'        If Me.EBayBestOfferAutomationLevel(Index) = "" Then
'            Me.EBayBestOfferAutomationLevel(Index) = 0
'        End If
'        If validateCurrency(Me.EBayBestOfferAutomationLevel(Index)) Then
'            Me.EBayBestOfferAutomationLevel(Index) = Format(Me.EBayBestOfferAutomationLevel(Index), "Currency")
'
'            If priceFieldIsLessThanPriceField(Me.EBayPrice, Me.EBayBestOfferAutomationLevel(Index)) Then
'                MsgBox "Best offer price cannot be higher than sell price!"
'                changed = False
'                Me.EBayBestOfferAutomationLevel(Index).SetFocus
'            Else
'                updateInventoryMaster IIf(Index = 0, "EBayBestOfferAutoDecline", "EBayBestOfferAutoAccept"), Me.EBayBestOfferAutomationLevel(Index), Me.ItemNumber, ""
'                changed = False
'
'                If Me.EBayBestOfferAutomationLevel(Index) = "$0.00" Then
'                    Me.EBayBestOfferAutomationLevel(Index) = ""
'                End If
'            End If
'        End If
'    End If
'End Sub

Private Sub EBayBestOfferAutoAccept_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            EBayBestOfferAutoAccept_LostFocus
        Case Is = vbKeyDelete
            EBayBestOfferAutoAccept_KeyPress KeyCode
    End Select
End Sub

Private Sub EBayBestOfferAutoAccept_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "EBayBestOfferAutoAccept"
End Sub

Private Sub EBayBestOfferAutoAccept_LostFocus()
    If changed = True Then
        If Me.EBayBestOfferAutoAccept = "" Then
            Me.EBayBestOfferAutoAccept = 0
        End If
        If validateCurrency(Me.EBayBestOfferAutoAccept) Then
            Me.EBayBestOfferAutoAccept = Format(Me.EBayBestOfferAutoAccept, "Currency")
        End If
        
        If priceFieldIsLessThanPriceField(Me.EBayPrice, Me.EBayBestOfferAutoAccept) Then
            MsgBox "Best offer price cannot be higher than sell price!"
            changed = False
            Me.EBayBestOfferAutoAccept.SetFocus
        Else
            updateInventoryMaster "EBayBestOfferAutoAccept", Me.EBayBestOfferAutoAccept, Me.ItemNumber, ""
            changed = False
            
            If Me.EBayBestOfferAutoAccept = "$0.00" Then
                Me.EBayBestOfferAutoAccept = ""
            End If
        End If
    End If
End Sub






Private Sub ItemDescription_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            ItemDescription_LostFocus
        Case Is = vbKeyDelete
            ItemDescription_KeyPress KeyCode
    End Select
End Sub

Private Sub ItemDescription_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "itemDescription"
End Sub

Private Sub ItemDescription_LostFocus()
    If changed = True Then
        Me.ItemDescription = Replace(Me.ItemDescription, """", "''")
        Me.ItemDescription = Replace(Me.ItemDescription, ",", "")
        If validateItemDescription(Me.ItemDescription) Then
            updateInventoryMaster "ItemDescription", Me.ItemDescription, Me.ItemNumber, "'"
            If DLookup("Desc3", "PartNumbers", "ItemNumber='" & Me.ItemNumber & "'") = "" Then
                updatePartNumbers "Desc3", Me.ItemDescription, Me.ItemNumber, "'"
            End If
            changed = False
        Else
            MsgBox "Item Description is invalid. Must be less than 255 characters."
        End If
    End If
End Sub

Private Sub ListPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            ListPrice_LostFocus
        Case Is = vbKeyDelete
            ListPrice_KeyPress KeyCode
    End Select
End Sub

Private Sub ListPrice_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "ListPrice"
End Sub

Private Sub ListPrice_LostFocus()
    If changed = True Then
        If validateCurrency(Me.ListPrice) Then
            Me.ListPrice = Format(Me.ListPrice, "Currency")
            updateInventoryMaster "ListPrice", Me.ListPrice, Me.ItemNumber, ""
DatabaseFunctions.ModifyItemCost Me.LoadedItemID, "List Price", Me.ListPrice
            changed = False
        Else
            MsgBox "Invalid list price. Must be numeric."
        End If
    End If
End Sub

Private Sub MAPP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            MAPP_LostFocus
        Case Is = vbKeyDelete
            MAPP_KeyPress KeyCode
    End Select
End Sub

Private Sub MAPP_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "MAPP"
End Sub

Private Sub MAPP_LostFocus()
    If changed = True Then
        If Me.MAPP = "" Then
            Me.MAPP = "0"
        End If
        If validateCurrency(Me.MAPP) Then
            Me.MAPP = Format(Me.MAPP, "Currency")
            
            LogItemMAPPChanged Me.ItemNumber, Me.MAPP, False
            updateInventoryMaster "MAPP", Me.MAPP, Me.ItemNumber, ""
DatabaseFunctions.ModifyItemCost Me.LoadedItemID, "MAPP", Me.MAPP 'already formatted from "" => "$0.00"
            If Me.MAPP = "$0.00" Then
                Me.MAPP = ""
            End If
            changed = False
        Else
            MsgBox "Invalid MAPP, must be numeric."
        End If
    End If
End Sub

Private Sub MAPP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.MAPP.tag = "N" Then
        Me.MAPP.tag = "Y"
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT TOP 1 TimeChanged, OldMAPP FROM InventoryMAPPLog WHERE IgnoreFlag=0 AND ItemNumber='" & Me.ItemNumber & "' ORDER BY TimeChanged DESC")
        Dim txt As String
        If rst.EOF Then
            txt = "No mapp change info for this item."
        Else
            txt = "Last MAPP Change: " & Format(rst("TimeChanged"), "Short Date") & ", was " & Format(rst("OldMAPP"), "Currency")
        End If
        rst.Close
        Set rst = Nothing
        If Me.MAPP <> "" Then
            txt = txt & vbCrLf & vbCrLf & "MAPP - Sales Tax: " & Format(Me.MAPP / 1.0635, "Currency")
        End If
        'nothing else
        IMtt.AddToolTipToCtl Me.MAPP, txt
        Me.CurrentMAPPdata = txt
        
    End If
    
    'MAPP.BackColor = RGB(0, 255, 255)
            MAPP.Refresh
    
            'If Me.HidePopUp = False Then
                'PopUpFrm.Visible = True
                'Me.popupTmr.Enabled = True
                'PopUpFrm.InfoLbl.Caption = Me.CurrentMAPPdata
                'PopUpFrm.ZOrder 0
                'PopUpFrm.SetFocus
                'PopUpFrm.Move Me.Left + Me.FramePricingInfo.Left - 500, Me.Top + Me.FramePricingInfo.Top + 50 + MAPP.Top   'ScaleX(NewCost.Left, vbPixels, vbTwips), ScaleY(NewCost.Top, vbPixels, vbTwips) ', PopUpFrm.width, PopUpFrm.Height
                'PopUpFrm.Refresh
            'End If
    
End Sub

'Private Sub InventoriedDate_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case Is = vbKeyReturn
'            InventoriedDate_LostFocus
'        Case Is = vbKeyDelete
'            InventoriedDate_KeyPress KeyCode
'    End Select
'End Sub
'
'Private Sub InventoriedDate_KeyPress(KeyAscii As Integer)
'    changed = True
'    whichCtl = "MfrCost"
'End Sub
'
'Private Sub InventoriedDate_LostFocus()
'    If changed = True Then
'        If IsDate(Me.InventoriedDate) Then
'            Me.InventoriedDate = Format(Me.InventoriedDate, "Short Date")
'            If MsgBox("Update entire line code?", vbYesNo) = vbYes Then
'                DB.Execute "UPDATE InventoryMaster SET DateLastMfrCostChange='" & Me.InventoriedDate & "' WHERE ProductLine='" & Me.prodLine & "'"
'            Else
'                updateInventoryMaster "DateLastMfrCostChange", Me.InventoriedDate, Me.ItemNumber, "'"
'            End If
'            changed = False
'        Else
'            MsgBox "Invalid date!"
'        End If
'    End If
'End Sub

Private Sub NewCost_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            NewCost_LostFocus
        Case Is = vbKeyDelete
            NewCost_KeyPress KeyCode
    End Select
End Sub

Private Sub NewCost_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "NewCost"
End Sub



Private Sub NewCost_LostFocus()

    If changed = True Then
        If Me.NewCost = "" Then
            Me.NewCost = "0.00"
        End If
        If validateCurrency(Me.NewCost) Then
            Me.NewCost = Format(Me.NewCost, "Currency")
            LogItemTNCChanged Me.ItemNumber, Me.NewCost
            updateInventoryMaster "TNC", Me.NewCost, Me.ItemNumber, ""
DatabaseFunctions.ModifyItemCost Me.LoadedItemID, "New Cost", Me.NewCost
            IMtt.RemoveToolTipFromCtl Me.NewCost
            Me.NewCost.tag = "N"
            'If CBool(Me.chkDropship) Then
                If Left(Me.ItemNumber, 3) <> "XXX" And Left(Me.ItemNumber, 3) <> "ZZZ" Then
                    Dim zzz As String
                    zzz = ZZZifyItemNumber(Me.ItemNumber)
                    If zzz <> "" Then
                        updateInventoryMaster "TNC", Me.NewCost, zzz, ""
Dim zzzID As Long
zzzID = Utilities.GetItemIDByItemCode(zzz)
If zzzID <> -1 Then
DatabaseFunctions.ModifyItemCost zzzID, "New Cost", Me.NewCost
End If
                    End If
                End If
            'End If
            If Me.NewCost = "$0.00" Then
                Me.NewCost = ""
            End If
            changed = False
        Else
            MsgBox "Invalid TNC. Must be numeric."
        End If
        'If Me.NewCost = "" Then
        '    LogItemTNCChanged Me.ItemNumber, "0"
        '    updateInventoryMaster "TNC", 0, Me.ItemNumber, ""
        '    'Me.NewCost.ToolTipText = ""
        '    IMtt.RemoveToolTip Me.NewCost.hWnd
        '    Me.NewCost.tag = "N"
        '    changed = False
        '    If CBool(Me.chkDropship) Then
        '        If Left(Me.ItemNumber, 3) <> "XXX" And Left(Me.ItemNumber, 3) <> "ZZZ" Then
        '            Dim zzz As String
        '            zzz = ZZZifyItemNumber(Me.ItemNumber)
        '            If zzz <> "" Then
        '                updateInventoryMaster "TNC", "0", zzz, ""
        '            End If
        '        End If
        '    End If
        'ElseIf validateCurrency(Me.NewCost) Then
        '    Me.NewCost = Format(Me.NewCost, "Currency")
        '    LogItemTNCChanged Me.ItemNumber, Me.NewCost
        '    updateInventoryMaster "TNC", Me.NewCost, Me.ItemNumber, ""
        '    'Me.NewCost.ToolTipText = "TNC Entered on: " & Format(Now(), "Short Date")
        '    'IMtt.AddToolTip Me.NewCost.hWnd, Me.hWnd, "TNC Entered on: " & Format(Now(), "Short Date")
        '    changed = False
        'Else
        '    MsgBox "Invalid TNC. Must be numeric."
        'End If
    End If
End Sub

'Private Sub NewOP_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case Is = vbKeyReturn
'            NewOP_LostFocus
'        Case Is = vbKeyDelete
'            NewOP_KeyPress KeyCode
'    End Select
'End Sub
'
'Private Sub NewOP_KeyPress(KeyAscii As Integer)
'    changed = True
'    whichCtl = "NewOP"
'End Sub
'
'Private Sub NewOP_LostFocus()
'    If changed = True Then
'        DB.Execute ("UPDATE InventoryMaster SET ReorderPoint=" & Me.NewOP & ", NewReorderPoint=" & Me.NewOP & " WHERE ItemNumber='" & Me.ItemNumber & "'")
'        changed = False
'    End If
'End Sub

Private Sub NewCost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.NewCost <> "" And Me.NewCost.tag = "N" Then
        Me.NewCost.tag = "Y"
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT TOP 1 TimeChanged, OldCost, NewCost, IsSwap FROM InventoryCostLog WHERE ItemNumber='" & Me.ItemNumber & "' AND (IsRegularCost=0 OR IsSwap=1) ORDER BY TimeChanged DESC")
        Dim txt As String
        If rst.EOF Then
            'Me.NewCost.ToolTipText = "No change info for this item."
            txt = "No cost change info for this item."
        Else
            If rst("IsSwap") Then
                'Me.NewCost.ToolTipText = "TNC entered on: " & Format(rst("TimeChanged"), "Short Date") & ", was " & rst("NewCost")
                txt = "TNC entered on: " & Format(rst("TimeChanged"), "Short Date") & ", was " & rst("NewCost")
            Else
                'Me.NewCost.ToolTipText = "TNC entered on: " & Format(rst("TimeChanged"), "Short Date") & ", was " & rst("OldCost")
                txt = "TNC entered on: " & Format(rst("TimeChanged"), "Short Date") & ", was " & rst("OldCost")
            End If
        End If
        rst.Close
        Set rst = Nothing
        If validateCurrency(Me.NewCost) Then
            txt = txt & vbCrLf & vbCrLf
            txt = txt & "Baker's Dozen: " & Format(Format(Me.NewCost, "General Number") * 12 / 13, "Currency")
            txt = txt & vbCrLf & "11+1: " & Format(Format(Me.NewCost, "General Number") * 11 / 12, "Currency")
            txt = txt & vbCrLf & "10+1: " & Format(Format(Me.NewCost, "General Number") * 10 / 11, "Currency")
            txt = txt & vbCrLf & "9+1: " & Format(Format(Me.NewCost, "General Number") * 9 / 10, "Currency")
            txt = txt & vbCrLf & "8+1: " & Format(Format(Me.NewCost, "General Number") * 8 / 9, "Currency")
            txt = txt & vbCrLf & "7+1: " & Format(Format(Me.NewCost, "General Number") * 7 / 8, "Currency")
            txt = txt & vbCrLf & "6+1: " & Format(Format(Me.NewCost, "General Number") * 6 / 7, "Currency")
            txt = txt & vbCrLf & "5+1: " & Format(Format(Me.NewCost, "General Number") * 5 / 6, "Currency")
            txt = txt & vbCrLf & "4+1: " & Format(Format(Me.NewCost, "General Number") * 4 / 5, "Currency")
            txt = txt & vbCrLf & "3+1: " & Format(Format(Me.NewCost, "General Number") * 3 / 4, "Currency")
            txt = txt & vbCrLf & "2+1: " & Format(Format(Me.NewCost, "General Number") * 2 / 3, "Currency")
            txt = txt & vbCrLf
            'Tom's Code 9-16-2014 to include -1%,-2%,-3%,-6%, and -9%
            txt = txt & vbCrLf & "-1%: " & Format(Format(Me.NewCost, "General Number") * 0.99, "Currency")
            txt = txt & vbCrLf & "-2%: " & Format(Format(Me.NewCost, "General Number") * 0.98, "Currency")
            txt = txt & vbCrLf & "-3%: " & Format(Format(Me.NewCost, "General Number") * 0.97, "Currency")
            txt = txt & vbCrLf & "-5%: " & Format(Format(Me.NewCost, "General Number") * 0.95, "Currency")
            txt = txt & vbCrLf & "-6%: " & Format(Format(Me.NewCost, "General Number") * 0.94, "Currency")
            txt = txt & vbCrLf & "-7.5%: " & Format(Format(Me.NewCost, "General Number") * 0.925, "Currency")
            txt = txt & vbCrLf & "-8%: " & Format(Format(Me.NewCost, "General Number") * 0.92, "Currency")
            txt = txt & vbCrLf & "-9%: " & Format(Format(Me.NewCost, "General Number") * 0.91, "Currency")
            txt = txt & vbCrLf & "-10%: " & Format(Format(Me.NewCost, "General Number") * 0.9, "Currency")
            txt = txt & vbCrLf & "-12%: " & Format(Format(Me.NewCost, "General Number") * 0.88, "Currency")
            txt = txt & vbCrLf & "-15%: " & Format(Format(Me.NewCost, "General Number") * 0.85, "Currency")
            txt = txt & vbCrLf & "-20%: " & Format(Format(Me.NewCost, "General Number") * 0.8, "Currency")
            txt = txt & vbCrLf
            txt = txt & vbCrLf & "+50%: " & Format(Format(Me.NewCost, "General Number") * 1.5, "Currency")
            txt = txt & vbCrLf & "+33%: " & Format(Format(Me.NewCost, "General Number") * 1.33, "Currency")
            txt = txt & vbCrLf & "+25%: " & Format(Format(Me.NewCost, "General Number") * 1.25, "Currency")
            txt = txt & vbCrLf & "+15%: " & Format(Format(Me.NewCost, "General Number") * 1.15, "Currency")
            txt = txt & vbCrLf & "+10%: " & Format(Format(Me.NewCost, "General Number") * 1.1, "Currency")
            txt = txt & vbCrLf & "+5%: " & Format(Format(Me.NewCost, "General Number") * 1.05, "Currency")
            txt = txt & vbCrLf
            If Format(Me.StdCost, "General Number") = 0 Then
                txt = txt & vbCrLf & "TNC % above STD: undefined"
            Else
                If validateCurrency(Me.NewCost.Text) = False Then
                txt = txt & vbCrLf & "TNC % above STD: Invalid field input; cannot parse"
                Else
                txt = txt & vbCrLf & "TNC % above STD: " & Format(((Format(Me.NewCost, "General Number") / Format(Me.StdCost, "General Number")) - 1) * 100, "#.##") & "%"
                End If
            End If
            txt = txt & vbCrLf
            txt = txt & vbCrLf & "+ Yahoo 3%: " & Format(Format(Me.NewCost, "General Number") * 1.03, "Currency")
            IMtt.AddToolTipToCtl Me.NewCost, txt
            
            Me.CurrentTNCdata = txt

            
            
            
            
            
        End If
    End If
    
           'Me.NewCost.BackColor = RGB(0, 255, 255)
            NewCost.Refresh
    
            'If Me.HidePopUp = False Then
                'PopUpFrm.Visible = True
                'Me.popupTmr.Enabled = True
                'PopUpFrm.InfoLbl.Caption = Me.CurrentTNCdata
                'PopUpFrm.ZOrder 0
                'PopUpFrm.SetFocus
                'PopUpFrm.Move Me.Left + Me.FramePricingInfo.Left - 500, Me.Top + Me.FramePricingInfo.Top + 50 + NewCost.Top   'ScaleX(NewCost.Left, vbPixels, vbTwips), ScaleY(NewCost.Top, vbPixels, vbTwips) ', PopUpFrm.width, PopUpFrm.Height
                'PopUpFrm.Refresh
            'End If
    
End Sub
Public Sub HidePopUpFrm()
    PopUpFrm.Visible = False
    'Me.NewCost.BackColor = vbWhite
    'Me.StdCost.BackColor = vbWhite
    'Me.AvgCost.BackColor = Me.textBoxBackColor
    'Me.ListPrice.BackColor = Me.textBoxBackColor
    'Me.InventoriedDate.BackColor = Me.textBoxBackColor
    'Me.DropshipCost.BackColor = Me.textBoxBackColor
    'Me.MAPP.BackColor = Me.textBoxBackColor
    'Me.StdPrice.BackColor = vbWhite
    'Me.EBayPrice.BackColor = vbWhite
    'Me.ItemNumber.BackColor = vbWhite
    'Me.SubtitleCmb.BackColor = vbWhite
    
    'Dim X As Integer
    'Do
    '    X = X + 1
    'Me.IDiscountMarkupPriceRate(X).BackColor = vbWhite

    'Loop Until X = Me.IDiscountMarkupPriceRate.count
    
    'Dim Y As Integer
    'Do
    '    Y = Y + 1
    'Me.DiscountMarkupPriceRate(Y).BackColor = vbWhite

    'Loop Until Y = Me.DiscountMarkupPriceRate.count
    
    
End Sub
Private Sub OrderPointStore_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            OrderPointStore_LostFocus
        Case Is = vbKeyDelete
            OrderPointStore_KeyPress KeyCode
    End Select
End Sub

Private Sub OrderPointStore_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
    If KeyAscii <> 0 Then
        changed = True
        whichCtl = "OrderPointStore"
    End If
End Sub

Private Sub OrderPointStore_LostFocus()
    If changed Then
        If Me.OrderPointStore = "" Then
            Me.OrderPointStore = "0"
        End If
        If IsIntegeric(Me.OrderPointStore) Then
            'Dim continue As Boolean
            'If Me.chkOrderPointLockStore Then
            '    continue = True
            'ElseIf vbYes = MsgBox("Lock order point for the store at " & Me.OrderPointStore & "?", vbYesNo) Then
            '    continue = True
            'Else
            '    continue = False
            'End If
            'If continue Then
                fillingForm = True
                Me.chkOrderPointLockStore = 1
                fillingForm = False
                DB.Execute "UPDATE InventoryLocationInfo SET LockOrderPoint=1, OrderPoint=" & Me.OrderPointStore & " WHERE WarehouseID=1 AND ItemNumber='" & Me.ItemNumber & "'"
                'Me.OrderPointTotal = CLng(Me.OrderPointStore) + CLng(Me.OrderPointWhse)
                fillTotalOrderPoint
                changed = False
            'Else
            '    'should switch back, probably
            'End If
        Else
            MsgBox "Invalid order point! WTF!"
            Me.OrderPointStore.SetFocus
        End If
    End If
End Sub

Private Sub OrderPointWhse_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            OrderPointWhse_LostFocus
        Case Is = vbKeyDelete
            OrderPointWhse_KeyPress KeyCode
    End Select
End Sub

Private Sub OrderPointWhse_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
    If KeyAscii <> 0 Then
        changed = True
        whichCtl = "OrderPointWhse"
    End If
End Sub

Private Sub OrderPointWhse_LostFocus()
    If changed Then
        If Me.OrderPointWhse = "" Then
            Me.OrderPointWhse = "0"
        End If
        If IsIntegeric(Me.OrderPointWhse) Then
            'Dim continue As Boolean
            'If Me.chkOrderPointLockWhse Then
            '    continue = True
            'ElseIf vbYes = MsgBox("Lock order point for the warehouse at " & Me.OrderPointWhse & "?", vbYesNo) Then
            '    continue = True
            'Else
            '    continue = False
            'End If
            'If continue Then
                fillingForm = True
                Me.chkOrderPointLockWhse = 1
                fillingForm = False
                DB.Execute "UPDATE InventoryLocationInfo SET LockOrderPoint=1, OrderPoint=" & Me.OrderPointWhse & " WHERE WarehouseID=5 AND ItemNumber='" & Me.ItemNumber & "'"
                'Me.OrderPointTotal = CLng(Me.OrderPointStore) + CLng(Me.OrderPointWhse)
                fillTotalOrderPoint
                changed = False
            'Else
            '    'should switch back, probably
            'End If
        Else
            MsgBox "Invalid order point! WTF!"
            Me.OrderPointWhse.SetFocus
        End If
    End If
End Sub

Private Sub POComment_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            POComment_LostFocus
        Case Is = vbKeyDelete
            POComment_KeyPress KeyCode
    End Select
End Sub

Private Sub POComment_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "POComment"
End Sub

Private Sub POComment_LostFocus()
    If changed = True Then
        If validateItemPOComment(Me.POComment) Then
            updateInventoryMaster "POComment", Me.POComment, Me.ItemNumber, "'"
            changed = False
        Else
            MsgBox "Invalid PO Comment, must be less than 50 chars."
            Me.POComment.SetFocus
        End If
    End If
End Sub

Private Sub poList_DblClick()
    If Not lockedDown Then
        'If Not Me.POList.Selected(0) Then
            Load PurchaseOrdersForm
            PurchaseOrdersForm.GoToPurchaseOrder Mid(Me.POList, InStrRev(Me.POList, vbTab) + 1)
            PurchaseOrdersForm.Show
        'End If
    End If
End Sub

Private Sub priceComparisonData_ReceivedData(rxData As String)
'Call MsgBox("hot damn, it worked... " + rxData)
Call PriceCompareFrm.priceBrowser.Document.Write(rxData)
End Sub

Private Sub PrimaryVendor_Click()
    changed = True
    PrimaryVendor_LostFocus
End Sub

Private Sub PrimaryVendor_KeyDown(KeyCode As Integer, Shift As Integer)
    If lockedDown Then
        KeyCode = 0
    Else
        AutoCompleteKeyDown Me.PrimaryVendor, KeyCode, Shift
        Select Case KeyCode
            Case Is = vbKeyReturn
                PrimaryVendor_LostFocus
            Case Is = vbKeyDelete
                PrimaryVendor_KeyPress KeyCode
        End Select
    End If
End Sub

Private Sub PrimaryVendor_KeyPress(KeyAscii As Integer)
    If lockedDown Then
        KeyAscii = 0
    Else
        AutoCompleteKeyPress Me.PrimaryVendor, KeyAscii
        changed = True
        whichCtl = "PrimaryVendor"
    End If
End Sub

Private Sub PrimaryVendor_LostFocus()
    If Not lockedDown Then
        AutoCompleteLostFocus Me.PrimaryVendor
        If changed = True Then
            updateInventoryMaster "PrimaryVendor", lookupPrimaryVendorNumber(Me.PrimaryVendor), Me.ItemNumber, "'"
            changed = False
        End If
    End If
End Sub

Private Sub prodLine_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            prodLine_LostFocus
        Case Is = vbKeyDelete
            prodLine_KeyPress KeyCode
    End Select
End Sub

Private Sub prodLine_KeyPress(KeyAscii As Integer)
    changed = True
    KeyAscii = ForceUppercase(KeyAscii)
    whichCtl = "prodLine"
End Sub

Private Sub prodLine_LostFocus()
    prodLine = Trim(UCase(Me.prodLine))
    If changed = True Then
        Me.prodLine = UCase(Me.prodLine)
        If validateProductLine(Me.prodLine) Then
            updateInventoryMaster "ProductLine", Me.prodLine, Me.ItemNumber, "'"
            changed = False
        Else
            MsgBox "Line code is invalid. Must be 3 chars, alphanumeric or -"
        End If
    End If
End Sub



Private Sub QtyOrdered_GotFocus()
    Me.QtyOrdered.selStart = 0
    Me.QtyOrdered.SelLength = Len(Me.QtyOrdered)
End Sub

Private Sub QtyOrdered_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            QtyOrdered_LostFocus
        Case Is = vbKeyDelete
            QtyOrdered_KeyPress KeyCode
    End Select
End Sub

Private Sub QtyOrdered_KeyPress(KeyAscii As Integer)
    If Me.QtyOrdered.Locked = False Then
        KeyAscii = LimitToNumbers(KeyAscii)
        changed = True
        whichCtl = "QtyOrdered"
    End If
End Sub

Private Sub QtyOrdered_LostFocus()
    If changed = True Then
        Dim isFakeItem As String
        isFakeItem = DLookup("WebPointered", "PartNumbers", "ItemNumber='" & Me.ItemNumber & "'")
        If Me.PrimaryVendor = "" Then
            MsgBox "Item must have a primary vendor before putting it on an order."
            Me.QtyOrdered = "0"
            changed = False
        ElseIf isFakeItem = "True" Then
            MsgBox "This is a fake template item, and shouldn't go on an order."
            Me.QtyOrdered = "0"
            changed = False
        ElseIf validateGeneralInteger(Me.QtyOrdered) = False Then
            MsgBox "Invalid qty to order, must be an integer."
            Me.QtyOrdered = "0"
            changed = False
        Else
            Dim POID As String
            'fill poid
            Dim pokeys As Variant
            pokeys = GetPurchaseOrderKeys(Me.ItemNumber)
            Dim rst As ADODB.Recordset
            Set rst = DB.retrieve("SELECT ID FROM PurchaseOrders WHERE Exported=0 AND VendorNumber='" & pokeys(0) & "' AND CoopVendor='" & pokeys(1) & "' AND ProductLineCode='" & pokeys(2) & "'")
            Select Case rst.RecordCount
                Case 0:
                    POID = ""
                Case 1:
                    POID = rst("ID")
                Case Else:
                    MsgBox "Error: multiple POs found, this should not have been enabled!"
                    rst.Close
                    Set rst = Nothing
                    Exit Sub
            End Select
            rst.Close
            Set rst = Nothing
            If Me.QtyOrdered = "0" Then
                'updateInventoryMaster "QtyOrdered", Me.QtyOrdered, Me.ItemNumber, ""
                'updateInventoryMaster "POID", "Null", Me.ItemNumber, ""
                If POID <> "" Then
                    PurchaseOrderLineEdit CLng(POID), Me.ItemNumber, 0
                End If
            Else
                If POID = "" Then
                    POID = CreatePurchaseOrder(Me.ItemNumber)
                    'if creating a new d/s po, ask for order to fill address, etc
                    If Left(Me.ItemNumber, 3) = "ZZZ" Then
                        ConvertToDropshipPO CLng(POID)
                    End If
                End If
                If POID <> "" Then
                    If DLookup("Finalized", "PurchaseOrders", "ID=" & POID) = True Then
                        MsgBox "This PO has been finalized! If you want to add this item, un-finalize it first."
                        'Me.QtyOrdered = DLookup("QtyOrdered", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'")
                        Me.QtyOrdered = DLookup("Quantity", "PurchaseOrderLines", "HeaderID=" & POID & " AND ItemNumber='" & Me.ItemNumber & "'")
                    Else
                        'updateInventoryMaster "QtyOrdered", Me.QtyOrdered, Me.ItemNumber, ""
                        'updateInventoryMaster "POID", POID, Me.ItemNumber, ""
                        PurchaseOrderLineEdit CLng(POID), Me.ItemNumber, Me.QtyOrdered
                    End If
                End If
            End If
            changed = False
            requeryPOList DLookup("PrimaryVendor", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'")
        End If
    End If
End Sub

Private Sub removeFromJetBtn_Click()

End Sub

Private Sub ReplacedBy_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            ReplacedBy_LostFocus
        Case Is = vbKeyDelete
            ReplacedBy_KeyPress KeyCode
    End Select
End Sub

Private Sub ReplacedBy_KeyPress(KeyAscii As Integer)
    changed = True
    KeyAscii = ForceUppercase(KeyAscii)
    whichCtl = "ReplacedBy"
End Sub

Private Sub ReplacedBy_LostFocus()
    If changed = True Then
        Me.ReplacedBy = UCase(Me.ReplacedBy)
        If Me.ReplacedBy = "" Then
            updateInventoryMaster "ReplacedBy", "Null", Me.ItemNumber, ""
            changed = False
        ElseIf Left(Me.ReplacedBy, 1) = "/" Or Left(Me.ReplacedBy, 1) = "\" Then
            updateInventoryMaster "ReplacedBy", Me.ReplacedBy, Me.ItemNumber, "'"
            changed = False
        ElseIf validateItemNumber(Me.ReplacedBy) And PLExists(Left$(Me.ReplacedBy, 3)) Then
            If ExistsInInventoryMaster(Me.ReplacedBy) Then
                updateInventoryMaster "ReplacedBy", Me.ReplacedBy, Me.ItemNumber, "'"
                updateInventoryMaster "ReplacementFor", Me.ItemNumber, Me.ReplacedBy, "'"
            ElseIf CreateReplacePart(Me.ReplacedBy, Me.ItemNumber) Then
                updateInventoryMaster "ReplacedBy", Me.ReplacedBy, Me.ItemNumber, "'"
                requeryForm Me.cmbFilters
            Else
                'error message already given
                Me.ReplacedBy = ""
            End If
            changed = False
        Else
            MsgBox "Invalid item number, or line code doesn't exist."
        End If
    End If
End Sub

Private Sub ReplacementFor_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            ReplacementFor_LostFocus
        Case Is = vbKeyDelete
            ReplacementFor_KeyPress KeyCode
    End Select
End Sub

Private Sub ReplacementFor_KeyPress(KeyAscii As Integer)
    changed = True
    KeyAscii = ForceUppercase(KeyAscii)
    whichCtl = "ReplacementFor"
End Sub

Private Sub ReplacementFor_LostFocus()
    If changed = True Then
        Me.ReplacementFor = Me.ReplacementFor
        If Me.ReplacementFor = "" Then
            updateInventoryMaster "ReplacementFor", "Null", Me.ItemNumber, ""
            changed = False
        ElseIf Left(Me.ReplacementFor, 1) = "/" Or Left(Me.ReplacementFor, 1) = "\" Then
            updateInventoryMaster "Replacementfor", Me.ReplacementFor, Me.ItemNumber, "'"
            changed = False
        ElseIf Not ExistsInInventoryMaster(Me.ReplacementFor) Then
            MsgBox "Couldn't find old item # in inventory master."
        Else
            updateInventoryMaster "ReplacementFor", Me.ReplacementFor, Me.ItemNumber, "'"
            updateInventoryMaster "ReplacedBy", Me.ItemNumber, Me.ReplacementFor, "'"
            changed = False
        End If
    End If
End Sub

Private Sub reviseOnJetBtn_Click()

End Sub

Private Sub shipInfoFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HidePopUpFrm
Me.HidePopUp = False
End Sub

Private Sub showHideCompareBtn_Click()
Dim handle As Long
If showHideCompareBtn.caption = "Hide" Then
    handle = IsPriceComparerOpen
    If handle > 0 Then
       PriceCompareCSharp.SetPriceComparerItem "<consume>", handle
       showHideCompareBtn.caption = "Show"
    Else
        MsgBox "Cant find Follow Me Pricing!"
        Exit Sub
    End If
    
    
Else
    handle = IsPriceComparerOpen
    If handle > 0 Then
        PriceCompareCSharp.SetPriceComparerItem "<puke>", handle
        showHideCompareBtn.caption = "Hide"
    Else
        MsgBox "Cant find Follow Me Pricing!"
        Exit Sub
    End If
End If
End Sub

Private Sub StdCost_GotFocus()
    Me.StdCost.selStart = 0
    Me.StdCost.SelLength = Len(Me.StdCost)
End Sub

Private Sub StdCost_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            StdCost_LostFocus
        Case Is = vbKeyDelete
            StdCost_KeyPress KeyCode
    End Select
End Sub

Private Sub StdCost_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "StdCost"
End Sub

Private Sub StdCost_LostFocus()
    If changed = True Then
        If Me.StdCost = "" Then
            Me.StdCost = "0"
        End If
        If validateCurrency(Me.StdCost) Then
            Me.StdCost = Format(Me.StdCost, "Currency")
            'UpdateNotificationEmail "cost", Me.ItemNumber
            'updateInventoryMaster "OldCost", "StdCost", Me.ItemNumber, ""
            LogItemCostChanged Me.ItemNumber, Me.StdCost
            updateInventoryMaster "StdCost", Me.StdCost, Me.ItemNumber, ""
DatabaseFunctions.ModifyItemCost Me.LoadedItemID, "Std Cost", Me.StdCost
            'Me.lblStdCost.ToolTipText = "Baker's Dozen: " & Format(Format(Me.StdCost, "General Number") * 12 / 13, "Currency") & "     " & "w/ freight: " & Format(Format(Me.StdCost, "General Number") * 12 / 13 + Format(Me.freightActualAvg, "General Number"), "Currency")
            Me.StdCost.tag = "N" & IIf(Me.lblIsKit.Visible, "K", "N") & IIf(Me.StdCost = "$0.00", "Z", "N")
            'If CBool(Me.chkDropship) Then
                If Left(Me.ItemNumber, 3) <> "XXX" And Left(Me.ItemNumber, 3) <> "ZZZ" Then
                    Dim zzz As String
                    zzz = ZZZifyItemNumber(Me.ItemNumber)
                    If zzz <> "" Then
                        updateInventoryMaster "StdCost", Me.StdCost, zzz, ""
Dim zzzID As Long
zzzID = Utilities.GetItemIDByItemCode(zzz)
If zzzID <> -1 Then
DatabaseFunctions.ModifyItemCost zzzID, "Std Cost", Me.StdCost
End If
                    End If
                End If
            'End If
            changed = False
        Else
            MsgBox "Invalid standard cost. Must be numeric."
        End If
    End If
End Sub

Private Sub StdCost_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.HidePopUp = True
Me.HidePopUpFrm
Sleep 300
StdCost.ZOrder 0
StdCost.SetFocus
StdCost.Refresh
End Sub

Private Sub StdCost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.StdCost = "" Then
        Exit Sub
    End If
    If Left(Me.StdCost.tag, 1) = "N" Then
        Me.StdCost.tag = "Y" & Mid(Me.StdCost.tag, 2)
        Dim txt As String
        txt = ""
        
        If Mid(Me.StdCost.tag, 2, 1) = "K" Then
            If Mid(Me.StdCost.tag, 3, 1) = "Z" Then
                txt = IIf(txt = "", "", txt & vbCrLf & vbCrLf) & "Kit cost computed from kit contents!"
            Else
                txt = IIf(txt = "", "", txt & vbCrLf & vbCrLf) & "Kit cost has been overridden! Accumulated cost of pieces is " & Format(DLookup("SUM(QuantityPerAssembly*StdCost)", "IM_SalesKitDetail INNER JOIN InventoryMaster ON IM_SalesKitDetail.ComponentItemCode=InventoryMaster.ItemNumber", "SalesKitNo='" & Me.ItemNumber & "'"), "Currency")
            End If
        End If
        
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT TOP 1 TimeChanged, OldCost FROM InventoryCostLog WHERE ItemNumber='" & Me.ItemNumber & "' AND (IsRegularCost=1 OR IsSwap=1) ORDER BY TimeChanged DESC")
        If rst.EOF Then
            'Me.StdCost.ToolTipText = "No change info for this item."
            txt = IIf(txt = "", "", txt & vbCrLf & vbCrLf) & "No cost change info for this item."
        Else
            'Me.StdCost.ToolTipText = "Last Cost Change: " & Format(rst("TimeChanged"), "Short Date") & ", was " & Format(rst("OldCost"), "Currency")
            txt = IIf(txt = "", "", txt & vbCrLf & vbCrLf) & "Last Cost Change: " & Format(rst("TimeChanged"), "Short Date") & ", was " & Format(rst("OldCost"), "Currency")
        End If
        rst.Close
        Set rst = Nothing
        
        txt = txt & vbCrLf & vbCrLf
        txt = txt & "Baker's Dozen: " & Format(Format(Me.StdCost, "General Number") * 12 / 13, "Currency")
        txt = txt & vbCrLf & "11+1: " & Format(Format(Me.StdCost, "General Number") * 11 / 12, "Currency")
        txt = txt & vbCrLf & "10+1: " & Format(Format(Me.StdCost, "General Number") * 10 / 11, "Currency")
        txt = txt & vbCrLf & "9+1: " & Format(Format(Me.StdCost, "General Number") * 9 / 10, "Currency")
        txt = txt & vbCrLf & "8+1: " & Format(Format(Me.StdCost, "General Number") * 8 / 9, "Currency")
        txt = txt & vbCrLf & "7+1: " & Format(Format(Me.StdCost, "General Number") * 7 / 8, "Currency")
        txt = txt & vbCrLf & "6+1: " & Format(Format(Me.StdCost, "General Number") * 6 / 7, "Currency")
        txt = txt & vbCrLf & "5+1: " & Format(Format(Me.StdCost, "General Number") * 5 / 6, "Currency")
        txt = txt & vbCrLf & "4+1: " & Format(Format(Me.StdCost, "General Number") * 4 / 5, "Currency")
        txt = txt & vbCrLf & "3+1: " & Format(Format(Me.StdCost, "General Number") * 3 / 4, "Currency")
        txt = txt & vbCrLf & "2+1: " & Format(Format(Me.StdCost, "General Number") * 2 / 3, "Currency")
        txt = txt & vbCrLf & "1+1: " & Format(Format(Me.StdCost, "General Number") * 1 / 2, "Currency")
        txt = txt & vbCrLf
        txt = txt & vbCrLf & "-1%: " & Format(Format(Me.StdCost, "General Number") * 0.99, "Currency")
        txt = txt & vbCrLf & "-2%: " & Format(Format(Me.StdCost, "General Number") * 0.98, "Currency")
        txt = txt & vbCrLf & "-3%: " & Format(Format(Me.StdCost, "General Number") * 0.97, "Currency")
        txt = txt & vbCrLf & "-5%: " & Format(Format(Me.StdCost, "General Number") * 0.95, "Currency")
        txt = txt & vbCrLf & "-6%: " & Format(Format(Me.StdCost, "General Number") * 0.94, "Currency")
        txt = txt & vbCrLf & "-7.5%: " & Format(Format(Me.StdCost, "General Number") * 0.925, "Currency")
        txt = txt & vbCrLf & "-8%: " & Format(Format(Me.StdCost, "General Number") * 0.92, "Currency")
        txt = txt & vbCrLf & "-9%: " & Format(Format(Me.StdCost, "General Number") * 0.91, "Currency")
        txt = txt & vbCrLf & "-10%: " & Format(Format(Me.StdCost, "General Number") * 0.9, "Currency")
        txt = txt & vbCrLf & "-12%: " & Format(Format(Me.StdCost, "General Number") * 0.88, "Currency")
        txt = txt & vbCrLf & "-15%: " & Format(Format(Me.StdCost, "General Number") * 0.85, "Currency")
        txt = txt & vbCrLf & "-20%: " & Format(Format(Me.StdCost, "General Number") * 0.8, "Currency")
        txt = txt & vbCrLf
        txt = txt & vbCrLf & "+50%: " & Format(Format(Me.StdCost, "General Number") * 1.5, "Currency")
        txt = txt & vbCrLf & "+33%: " & Format(Format(Me.StdCost, "General Number") * 1.33, "Currency")
        txt = txt & vbCrLf & "+25%: " & Format(Format(Me.StdCost, "General Number") * 1.25, "Currency")
        txt = txt & vbCrLf & "+15%: " & Format(Format(Me.StdCost, "General Number") * 1.15, "Currency")
        txt = txt & vbCrLf & "+10%: " & Format(Format(Me.StdCost, "General Number") * 1.1, "Currency")
        txt = txt & vbCrLf & "+5%: " & Format(Format(Me.StdCost, "General Number") * 1.05, "Currency")
        txt = txt & vbCrLf
        If validateCurrency(Me.NewCost) = False Or validateCurrency(Me.StdCost) = False Then
            txt = txt & vbCrLf & "TNC % above STD: invalid input field"
        Else
        If Me.NewCost = "" Or Format(Me.StdCost, "General Number") = 0 Then
            txt = txt & vbCrLf & "TNC % above STD: undefined"
        Else
            txt = txt & vbCrLf & "TNC % above STD: " & Format(((Format(Me.NewCost, "General Number") / Format(Me.StdCost, "General Number")) - 1) * 100, "#.##") & "%"
        End If
        End If
        txt = txt & vbCrLf
        txt = txt & vbCrLf & "+ Yahoo 3%: " & Format(Format(Me.StdCost, "General Number") * 1.03, "Currency")
        IMtt.AddToolTipToCtl Me.StdCost, txt
        
        Me.CurrentSTDdata = txt
        
        
    End If
    
               'Me.StdCost.BackColor = RGB(0, 255, 255)
            StdCost.Refresh
    
            'If Me.HidePopUp = False Then
                'PopUpFrm.Visible = True
                'Me.popupTmr.Enabled = True
                'PopUpFrm.InfoLbl.Caption = Me.CurrentSTDdata
                'PopUpFrm.ZOrder 0
                'PopUpFrm.SetFocus
                'PopUpFrm.Move Me.Left + Me.FramePricingInfo.Left - 500, Me.Top + Me.FramePricingInfo.Top + 50 + StdCost.Top   'ScaleX(NewCost.Left, vbPixels, vbTwips), ScaleY(NewCost.Top, vbPixels, vbTwips) ', PopUpFrm.width, PopUpFrm.Height
                'PopUpFrm.Refresh
            'End If
    
End Sub

Private Sub StdPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            StdPack_LostFocus
        Case Is = vbKeyDelete
            StdPack_KeyPress KeyCode
    End Select
End Sub

Private Sub StdPack_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
    changed = True
    whichCtl = "StdPack"
End Sub

Private Sub StdPack_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.StdPack) Then
            updateInventoryMaster "StdPack", Me.StdPack, Me.ItemNumber, ""
            changed = False
        Else
            MsgBox "Invalid Std Pack. Must be a number."
        End If
    End If
End Sub

Private Sub StdPrice_GotFocus()
    Me.StdPrice.selStart = 0
    Me.StdPrice.SelLength = Len(Me.StdPrice)
End Sub

Private Sub StdPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            StdPrice_LostFocus
        Case Is = vbKeyDelete
            StdPrice_KeyPress KeyCode
    End Select
End Sub

Private Sub StdPrice_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "StdPrice"
End Sub

Private Sub StdPrice_LostFocus()
    If changed = True Then
        If validateCurrency(Me.StdPrice) Then
            Me.StdPrice = Format(Me.StdPrice, "Currency")
            If Left(Me.ItemNumber, 3) <> "XXX" And Left(Me.ItemNumber, 3) <> "ZZZ" Then
                UpdateNotificationEmail "store price", ItemNumber
            End If
            'updateInventoryMaster "OldPrice", "StdPrice", Me.ItemNumber, ""
            LogItemStorePriceChanged Me.ItemNumber, Me.StdPrice
            updateInventoryMaster "StdPrice", Me.StdPrice, Me.ItemNumber, ""
            updateInventoryMaster "DiscountMarkupPriceRate1", Me.StdPrice, Me.ItemNumber, ""
            Me.DiscountMarkupPriceRate(1) = Me.StdPrice
            If priceFieldIsLessThanPriceField(Me.StdPrice, Me.StdCost) Then
                MsgBox "Warning: store price is below standard cost!", vbCritical
            End If
            If CBool(Me.chkDropship) Then
                If Left(Me.ItemNumber, 3) <> "XXX" And Left(Me.ItemNumber, 3) <> "ZZZ" Then
                    Dim zzz As String
                    zzz = ZZZifyItemNumber(Me.ItemNumber)
                    If zzz <> "" Then
                        updateInventoryMaster "StdPrice", Me.StdPrice, zzz, ""
                        updateInventoryMaster "DiscountMarkupPriceRate1", Me.DiscountMarkupPriceRate(1), zzz, ""
                    End If
                End If
            End If
            changed = False
        Else
            MsgBox "Invalid std price. Must be numeric."
        End If
    End If
End Sub

Private Sub StdPrice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.HidePopUp = True
Me.HidePopUpFrm
Sleep 300
StdPrice.ZOrder 0
StdPrice.SetFocus
StdPrice.Refresh
End Sub

Private Sub StdPrice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.StdPrice.tag = "N" Then
        Me.StdPrice.tag = "Y"
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT TOP 1 TimeChanged, OldPrice FROM InventoryPriceLog WHERE ItemNumber='" & Me.ItemNumber & "' AND IsStorePrice=1 ORDER BY TimeChanged DESC")
        Dim txt As String
        If rst.EOF Then
            'Me.StdPrice.ToolTipText = "No change info for this item."
            txt = "No price change info for this item."
        Else
            'Me.StdPrice.ToolTipText = "Last Price Change: " & Format(rst("TimeChanged"), "Short Date") & ", was " & Format(rst("OldPrice"), "Currency")
            txt = "Last Price Change: " & Format(rst("TimeChanged"), "Short Date") & ", was " & Format(rst("OldPrice"), "Currency")
        End If
        rst.Close
        Set rst = Nothing
        'Me.DiscountMarkupPriceRate(1).ToolTipText = Me.StdPrice.ToolTipText
        IMtt.AddToolTipToCtl Me.StdPrice, txt
        IMtt.AddToolTipToCtl Me.DiscountMarkupPriceRate(1), txt
        Me.CurrentSTDPricedata = txt
        Me.CurrentDiscountRate = txt
    End If
    
            'StdPrice.BackColor = RGB(0, 255, 255)
            StdPrice.Refresh
    
            'If Me.HidePopUp = False Then
                'PopUpFrm.Visible = True
                'Me.popupTmr.Enabled = True
                'PopUpFrm.InfoLbl.Caption = Me.CurrentSTDPricedata
                'PopUpFrm.ZOrder 0
                'PopUpFrm.SetFocus
                'PopUpFrm.Move Me.Left + Me.FramePricingInfo.Left - 500, Me.Top + Me.FramePricingInfo.Top + 50 + StdPrice.Top   'ScaleX(NewCost.Left, vbPixels, vbTwips), ScaleY(NewCost.Top, vbPixels, vbTwips) ', PopUpFrm.width, PopUpFrm.Height
                'PopUpFrm.Refresh
            'End If
    
End Sub

Private Sub SubTitleChk_Click()
If subtitleNotClicked = False Then
  If SubTitleChk.value = vbChecked Then
    'changed = True
    'whichCtl = "SubtitleCmb"
    If HasPermissionsTo("SubtitleOverride") = True Then
        SubtitleCmb.Enabled = True
        DeleteSelectedSubtitleBtn.Enabled = True
        AddSelectedSubtitleBtn.Enabled = True
    End If
        SubtitleCmb.ListIndex = 0
        DB.retrieve "UPDATE PartNumbers SET EBayOffloadRequired=1 WHERE ItemNumber='" & Me.ItemNumber.Text & "'"
        
    Else
    'If whichCtl = "SubtitleCmb" Then
    'changed = False
    'whichCtl = ""
            SubtitleCmb.Enabled = False
        DeleteSelectedSubtitleBtn.Enabled = False
        AddSelectedSubtitleBtn.Enabled = False
    End If

    'End If
    'SubtitleCmb_LostFocus
    'changed = False
End If
subtitleNotClicked = False
End Sub

Private Sub SubtitleCmb_Change()

If Len(SubtitleCmb.Text) > 55 Then
    SubtitleCmb.Text = Mid(SubtitleCmb.Text, 1, 55)
    SubtitleCmb.selStart = Len(SubtitleCmb.Text)
End If

subtitleCharCountLbl.caption = CStr(55 - Len(SubtitleCmb.Text))



'If Len(SubtitleCmb.Text) > 80 Then
'    SubtitleCmb.Text = Mid(SubtitleCmb.Text, 1, 80)
'End If
'    If SubTitleChk.value = vbChecked Then
'
'         If SubtitleCmb.Text <> "-Choose Subtitle-" And SubtitleCmb.Text <> "" Then
'
'            Dim isAlreadySubbed As ADODB.Recordset
'            Set isAlreadySubbed = DB.retrieve("SELECT ID FROM Subtitles WHERE Subtitle='" & SubtitleCmb.Text & "'")
'            If isAlreadySubbed.RecordCount > 0 Then
'                'check to see if this is a build-in-progess..
'                Dim testStr As String
'                testStr = Mid(SubtitleCmb.Text, 1, Len(SubtitleCmb.Text) - 1)
'                Set isAlreadySubbed = DB.retrieve("SELECT ID FROM Subtitles WHERE Subtitle='" & testStr & "'")
'                If isAlreadySubbed.RecordCount > 0 Then
'                    'this is an in-progress addition, so update it...
'                    DB.Execute "UPDATE Subtitles SET Subtitle='" & SubtitleCmb.Text & "' WHERE ID=" & isAlreadySubbed("ID")
'
'                Else
'                    'doesnt exist... make it...
'                    DB.Execute "INSERT INTO Subtitles (Subtitle) VALUES('" & SubtitleCmb.Text & "')"
'                End If
'
'            Else
'                'doesnt exist... make it...
'                                Dim testStr2 As String
'                testStr2 = Mid(SubtitleCmb.Text, 1, Len(SubtitleCmb.Text) - 1)
'                Set isAlreadySubbed = DB.retrieve("SELECT ID FROM Subtitles WHERE Subtitle='" & testStr2 & "'")
'                If isAlreadySubbed.RecordCount > 0 Then
'                    'this is an in-progress addition, so update it...
'                    DB.Execute "UPDATE Subtitles SET Subtitle='" & SubtitleCmb.Text & "' WHERE ID=" & isAlreadySubbed("ID")
'
'                Else
'                    'doesnt exist... make it...
'                    DB.Execute "INSERT INTO Subtitles (Subtitle) VALUES('" & SubtitleCmb.Text & "')"
'                End If
'            End If
'
'
'
'         End If
'
'
'    End If
End Sub

Private Sub SubtitleCmb_LostFocus()


'  If SubTitleChk.value = vbChecked Then
'
'        'update SubtitledItems in sql db with ItemNumber and SubtitleString
'
'        Dim rst As ADODB.Recordset
'        Set rst = DB.retrieve("SELECT ID,Subtitle FROM Subtitles WHERE Subtitle='" & SubtitleCmb.Text & "'")
'        Dim SubtitleID As String
'
'        If rst.RecordCount > 0 Then
'            'record exists... so now we check if item is currently in SubtitledItems.
'            'if it is, then we update it, if not, we create it...
'            SubtitleID = rst("ID")
'            Set rst = DB.retrieve("SELECT ItemNumber, SubtitleID FROM SubtitledItems WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
'            If rst.RecordCount > 0 Then
'            'so it does exist, now lets update it
'            DB.retrieve ("UPDATE SubtitledItems SET SubtitleID=" & SubtitleID & " WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
'            Else
'            'doesn't exist, so create it...
'            DB.retrieve ("INSERT INTO SubtitledItems VALUES('" & Me.ItemNumber.Text & "'," & SubtitleID & ")")
'            End If
'            'DB.retrieve ("UPDATE SubtitledItems SET SubtitleString='" & SubTitleTxt.Text & "' WHERE ItemNumber='" & ItemNumber.Text & "'")
'
'        Else
'            'we get blank and -Choose Subtitle- entries in subtitles... lets avoid these..
'            If SubtitleCmb.Text <> "-Choose Subtitle-" And SubtitleCmb.Text <> "" Then
'            'no records exist... so create a new one..
'             'Dim mymsg  As VbMsgBoxResult
'             'mymsg = MsgBox("Do you wish to add this as a new SubTitle?", vbYesNo, "Inserting Subtitle")
'            'If mymsg = vbYes Or mymsg = vbOK Then
'            '    DB.retrieve ("INSERT INTO Subtitles VALUES ('" & SubtitleCmb.Text & "')")
'            'Else
'            '    SubtitleCmb.ListIndex = 1
'            'End If
'            Set rst = DB.retrieve("SELECT ID FROM Subtitles WHERE Subtitle='" & SubtitleCmb.Text & "'")
'            If rst.RecordCount = 0 Then
'                'this subtitle must be new...
'                DB.Execute "INSERT INTO Subtitles (Subtitle) VALUES('" & SubtitleCmb.Text & "')"
'                Set rst = DB.retrieve("SELECT ID FROM Subtitles WHERE Subtitle='" & SubtitleCmb.Text & "'")
'            End If
'
'            SubtitleID = rst("ID")
'            'ok now we must check if item is in database and if so update it else insert it...
'            Set rst = DB.retrieve("SELECT ItemNumber,SubtitleID FROM SubtitledItems WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
'
'            If rst.RecordCount > 0 Then
'                'yea its there... let's update...
'                DB.retrieve ("UPDATE SubtitledItems SET SubtitleID=" & SubtitleID & " WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
'
'            Else
'                'ugh more work.. must insert a new row.
'                DB.retrieve ("INSERT INTO SubtitledItems VALUES('" & Me.ItemNumber.Text & "'," & SubtitleID & ")")
'
'
'            End If
'            End If
'        End If
'
'    Else
'        SubTitleChk.value = vbUnchecked
'        SubtitleCmb.Text = ""
'        DB.retrieve ("DELETE FROM SubtitledItems WHERE ItemNumber='" & Me.ItemNumber.Text & "'")
'    End If
   
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

Private Sub opFolliesStatus_Click(index As Integer)
    If Not fillingForm Then
        If INV_MAINT_FOLLIES_EDIT Then
            Select Case True
                Case Is = Me.opFolliesStatus(0) 'NOT
                    updateInventoryMaster "FolliesStatus", "0", Me.ItemNumber
                Case Is = Me.opFolliesStatus(1) '25%
                    updateInventoryMaster "FolliesStatus", "1", Me.ItemNumber
                Case Is = Me.opFolliesStatus(2) '40%
                    updateInventoryMaster "FolliesStatus", "2", Me.ItemNumber
                Case Is = Me.opFolliesStatus(3) '50%
                    updateInventoryMaster "FolliesStatus", "3", Me.ItemNumber
                Case Is = Me.opFolliesStatus(4) 'FLAT
                    updateInventoryMaster "FolliesStatus", "4", Me.ItemNumber
            End Select
        Else
            MsgBox "Follies aren't editable! Turn on in options first"
            Dim fs As Long
            fs = CLng(DLookup("FolliesStatus", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'"))
            fillingForm = True
            Select Case fs
                Case Is = 0
                    Me.opFolliesStatus(0).value = True
                Case Is = 1
                    Me.opFolliesStatus(1).value = True
                Case Is = 2
                    Me.opFolliesStatus(2).value = True
                Case Is = 3
                    Me.opFolliesStatus(3).value = True
                Case Is = 4
                    Me.opFolliesStatus(4).value = True
                Case Else
                    MsgBox "Invalid FolliesStatus for item!"
            End Select
            fillingForm = False
        End If
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
'do not remove asshole!
End Sub

Private Sub txtPosition_LostFocus()
    If changed = True Then
        If IsNumeric(Me.txtPosition) Then
            'ok
            'If Me.txtPosition - 1 < Me.barPosition.Min Then
            If Me.txtPosition - 1 < 0 Then
                'Me.txtPosition = Me.barPosition.Min
                Me.txtPosition = 0
            'ElseIf Me.txtPosition - 1 > Me.barPosition.max Then
            ElseIf Me.txtPosition - 1 > UBound(currentFilterItemList) Then
                'Me.txtPosition = Me.barPosition.max
                Me.txtPosition = UBound(currentFilterItemList)
            End If
            changed = False
            'Me.barPosition.value = Me.txtPosition - 1
            'Me.CurrentOffset = Me.txtPosition - 1
            'loadItemByFilterIndex Me.txtPosition - 1
        Else
            MsgBox "Position is not numeric!"
            'Me.txtPosition = Me.barPosition.value
            Me.txtPosition = 0
            changed = False
        End If
    End If
    If CLng(Me.txtPosition) < UBound(currentFilterItemList) And CLng(Me.txtPosition) > 0 Then
        If InventoryQuantityTriggersV2.IsSaving = False Then
            loadItemByFilterIndex Me.txtPosition - 1
        End If
    End If
End Sub

Private Sub Weight_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            Weight_LostFocus
        Case Is = vbKeyDelete
            Weight_KeyPress KeyCode
    End Select
End Sub

Private Sub Weight_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii, True, True)
    If KeyAscii <> 0 Then
        changed = True
        whichCtl = "Weight"
    End If
End Sub

Private Sub Weight_LostFocus()
    If changed = True Then
        Me.Weight = validateWeight(Me.Weight)
        If Me.Weight <> "" Then
            Me.Weight = Round(Me.Weight, 3)
            'updateInventoryMaster "Weight", Me.Weight, Me.ItemNumber, "'"
'            If Me.Weight <> "0" And Me.Dimensions <> "" And Me.Dimensions <> "0x0x0" Then
'                updateInventoryMaster "BoxID", DetermineBoxSize(Me.Dimensions, Me.Weight), Me.ItemNumber, ""
'            Else
'                updateInventoryMaster "BoxID", "-1", Me.ItemNumber, ""
'            End If
            Dim cid As String
            cid = DLookup("ComponentID", "InventoryComponentMap", "ItemID=" & Me.LoadedItemID)
            DB.Execute "UPDATE InventoryComponents SET Weight=" & Me.Weight & " WHERE ID=" & cid
            'DB.Execute "UPDATE InventoryComponents SET LastModifiedBy=" & GetCurrentUserID() & " WHERE ID=" & cid
            changed = False
            'If IsNumeric(Me.Weight) Then
            '    If CDbl(Me.Weight) > 150 Then
            '        If vbYes = MsgBox("Over 150 lbs! Should this be set to ""Truck Freight""?", vbYesNo) Then
            '            Me.opShippingType(SHIP_TF) = True
            '        End If
            '    End If
            'End If
            'updateExternalShippingDB Me.ItemNumber
            'updateShippingDB Me.ItemNumber
            ExternalComponentDBSync Me.ItemNumber
        Else
            MsgBox "Invalid weight. Not sure exactly what's wrong with it, though."
        End If
        'setBoxWarningFlag
    End If
End Sub


Private Function priceFieldIsLessThanPriceField(priceField1 As String, priceField2 As String) As Boolean
    Dim temp1 As String, temp2 As String
    temp1 = Replace(Replace(priceField1, "$", ""), ",", "")
    temp2 = Replace(Replace(priceField2, "$", ""), ",", "")
    
   If temp1 <> "" And temp2 <> "" Then
   If validateCurrency(temp1) = False Or validateCurrency(temp2) = False Then
    MsgBox "Invalid currency value entered."
    Exit Function
   End If
    End If
    
    If temp1 = "" Or temp2 = "" Then
        priceFieldIsLessThanPriceField = False
    Else
        priceFieldIsLessThanPriceField = CBool(CDbl(temp1) < CDbl(temp2))
    End If
End Function

'----------------
' EBAY
'-------------------

Private Sub btnEBayPriceCopyInternet_Click()
    Dim newprice As String
    newprice = DLookup("IDiscountMarkupPriceRate1", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'")
    LogItemEBayPriceChanged Me.ItemNumber, newprice
    updateInventoryMaster "EBayPrice", newprice, Me.ItemNumber
    Me.EBayPrice = Format(newprice, "Currency")
End Sub

'Private Sub btnEBayPriceCopyCostPlusPercentage_Click()
'    If Me.opDropshipCostType(2) <> True Then
'        MsgBox "You need to have the ""E"" button selected for this to work!"
'        Exit Sub
'    End If
'
'    Dim newprice As String
'    newprice = Replace(Replace(Me.DropshipCost, "$", ""), ",", "")
'    newprice = Round(CDec(newprice) * 1.1, 2)
'    LogItemEBayPriceChanged Me.ItemNumber, newprice
'    updateInventoryMaster "EBayPrice", newprice, Me.ItemNumber
'    Me.EBayPrice = Format(newprice, "Currency")
'    fillDropshipCostField
'End Sub

Private Sub btnEBayPriceCopyMAPP_Click()
    Dim newprice As String
    newprice = DLookup("MAPP", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'")
    LogItemEBayPriceChanged Me.ItemNumber, newprice
    updateInventoryMaster "EBayPrice", newprice, Me.ItemNumber
    Me.EBayPrice = Format(newprice, "Currency")
End Sub

'Private Sub chkEBayFreeShipping_Click()
'    If Not fillingForm Then
'        updateInventoryMaster "EBayFreeShipping", Me.chkEBayFreeShipping, Me.ItemNumber, ""
'    End If
'End Sub


Sub SortNumeric()
    Dim s As String * 10, n As Integer
    With Me.InternetComparisonLVI
        'justify the text using padding spaces
        For n = 1 To .ListItems.count
            s = vbNullString
            If .SortKey Then
                RSet s = .ListItems(n).SubItems(.SortKey)
                .ListItems(n).SubItems(.SortKey) = s
            Else
                RSet s = .ListItems(n)
                .ListItems(n).Text = s
            End If
        Next
        'sort column using "justified" text
        .sorted = True
        'trim spaces from the text
        For n = 1 To .ListItems.count
            If .SortKey Then
                .ListItems(n).SubItems(.SortKey) = _
                LTrim$(.ListItems(n).SubItems(.SortKey))
            Else
                .ListItems(n).Text = LTrim$(.ListItems(n))
            End If
        Next
    End With
End Sub
Sub SortNumericEBay()
    Dim s As String * 10, n As Integer
    With Me.EBayComparisonsLVI
        'justify the text using padding spaces
        For n = 1 To .ListItems.count
            s = vbNullString
            If .SortKey Then
                RSet s = .ListItems(n).SubItems(.SortKey)
                .ListItems(n).SubItems(.SortKey) = s
            Else
                RSet s = .ListItems(n)
                .ListItems(n).Text = s
            End If
        Next
        'sort column using "justified" text
        .sorted = True
        'trim spaces from the text
        For n = 1 To .ListItems.count
            If .SortKey Then
                .ListItems(n).SubItems(.SortKey) = _
                LTrim$(.ListItems(n).SubItems(.SortKey))
            Else
                .ListItems(n).Text = LTrim$(.ListItems(n))
            End If
        Next
    End With
End Sub
