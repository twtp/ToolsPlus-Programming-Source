VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#3.5#0"; "TPControls.ocx"
Begin VB.Form PurchaseOrdersForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Purchase Orders"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11010
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox QtyOnSO 
      Height          =   285
      Index           =   4
      Left            =   2340
      TabIndex        =   117
      Top             =   6030
      Width           =   645
   End
   Begin VB.TextBox QtyOnSO 
      Height          =   285
      Index           =   3
      Left            =   2340
      TabIndex        =   116
      Top             =   5130
      Width           =   645
   End
   Begin VB.TextBox QtyOnSO 
      Height          =   285
      Index           =   2
      Left            =   2340
      TabIndex        =   115
      Top             =   4230
      Width           =   645
   End
   Begin VB.TextBox QtyOnSO 
      Height          =   285
      Index           =   1
      Left            =   2340
      TabIndex        =   114
      Top             =   3330
      Width           =   645
   End
   Begin VB.TextBox QtyOnHand 
      Height          =   285
      Index           =   4
      Left            =   3000
      TabIndex        =   110
      Top             =   6030
      Width           =   585
   End
   Begin VB.TextBox QtyOnHand 
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   109
      Top             =   5130
      Width           =   585
   End
   Begin VB.TextBox QtyOnHand 
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   108
      Top             =   4230
      Width           =   585
   End
   Begin VB.TextBox QtyOnHand 
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   107
      Top             =   3330
      Width           =   585
   End
   Begin VB.TextBox QtyBackOrdered 
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   105
      Top             =   6030
      Width           =   495
   End
   Begin VB.TextBox QtyBackOrdered 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   104
      Top             =   5130
      Width           =   495
   End
   Begin VB.TextBox QtyBackOrdered 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   103
      Top             =   4230
      Width           =   495
   End
   Begin VB.TextBox QtyBackOrdered 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   102
      Top             =   3330
      Width           =   495
   End
   Begin VB.TextBox ReadyMessage 
      Height          =   525
      Left            =   5970
      MultiLine       =   -1  'True
      TabIndex        =   99
      Top             =   1770
      Width           =   2235
   End
   Begin VB.CommandButton btnSendReadyNotification 
      Caption         =   "Send ""Ready"" to Sherri"
      Height          =   555
      Left            =   6990
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   1200
      Width           =   1155
   End
   Begin VB.CommandButton btnPreviewPONew 
      Caption         =   "Preview (new)"
      Height          =   435
      Left            =   8820
      TabIndex        =   97
      Top             =   6990
      Width           =   825
   End
   Begin VB.CommandButton btnDropshipify 
      Caption         =   "D/S"
      Height          =   495
      Left            =   5910
      TabIndex        =   94
      Top             =   4920
      Width           =   585
   End
   Begin VB.CommandButton btnExcelExportPO 
      Caption         =   "Excel"
      Height          =   435
      Left            =   7140
      TabIndex        =   86
      Top             =   6990
      Width           =   825
   End
   Begin TPControls.DateMonthview PODueDateCal 
      Height          =   2370
      Left            =   8250
      TabIndex        =   85
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
   End
   Begin VB.TextBox MiscNotes 
      Height          =   675
      Left            =   6810
      MultiLine       =   -1  'True
      TabIndex        =   83
      Top             =   3060
      Width           =   4125
   End
   Begin VB.CommandButton btnLineCodeSubCfg 
      Caption         =   "LC Sub Config"
      Height          =   525
      Left            =   5880
      TabIndex        =   82
      Top             =   6060
      Width           =   675
   End
   Begin VB.TextBox POID 
      Height          =   285
      Left            =   2880
      TabIndex        =   81
      Text            =   "POID"
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox PODueDate 
      Enabled         =   0   'False
      Height          =   315
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   80
      Top             =   2700
      Width           =   1365
   End
   Begin VB.TextBox NumPayments 
      Height          =   285
      Left            =   6840
      TabIndex        =   78
      Top             =   390
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox DeferredList 
      Height          =   315
      Left            =   4470
      TabIndex        =   76
      Top             =   390
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkDeferred 
      Caption         =   "Deferred"
      Height          =   255
      Left            =   3480
      TabIndex        =   75
      Top             =   390
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ComboBox POTerms 
      Height          =   315
      Left            =   4290
      TabIndex        =   74
      Top             =   30
      Width           =   3075
   End
   Begin VB.ComboBox VendorPOAddr 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   72
      Top             =   1050
      Width           =   2205
   End
   Begin VB.ComboBox VendorNumber 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   71
      Top             =   720
      Width           =   2205
   End
   Begin VB.CommandButton btnFinalize 
      Caption         =   "Finalize PO"
      Height          =   285
      Left            =   6870
      TabIndex        =   68
      Top             =   2760
      Width           =   1485
   End
   Begin VB.CommandButton btnPreviewPO 
      Caption         =   "Preview (old)"
      Height          =   435
      Left            =   7980
      TabIndex        =   67
      Top             =   6990
      Width           =   825
   End
   Begin VB.Frame frameFreight 
      Caption         =   "Freight:"
      Height          =   945
      Left            =   4590
      TabIndex        =   66
      Top             =   780
      Width           =   1755
      Begin VB.OptionButton chkFreightPay 
         Caption         =   "We Pay Freight"
         Height          =   255
         Left            =   150
         TabIndex        =   70
         Top             =   570
         Width           =   1455
      End
      Begin VB.OptionButton chkFreightFree 
         Caption         =   "Freight Free"
         Height          =   255
         Left            =   150
         TabIndex        =   69
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.CommandButton btnDeletePO 
      Caption         =   "Delete"
      Height          =   435
      Left            =   5880
      TabIndex        =   52
      Top             =   6990
      Width           =   705
   End
   Begin VB.Frame frameItems 
      Caption         =   "Purchase Order Items:"
      Height          =   5445
      Left            =   75
      TabIndex        =   14
      Top             =   1800
      Width           =   5760
      Begin VB.TextBox QtyOnSO 
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   113
         Top             =   630
         Width           =   645
      End
      Begin VB.TextBox QtyOnHand 
         Height          =   285
         Index           =   0
         Left            =   2940
         TabIndex        =   106
         Top             =   630
         Width           =   585
      End
      Begin VB.TextBox QtyBackOrdered 
         Height          =   285
         Index           =   0
         Left            =   1740
         TabIndex        =   101
         Top             =   630
         Width           =   495
      End
      Begin VB.TextBox OrderTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   5070
         Width           =   1245
      End
      Begin VB.VScrollBar itemScrollBar 
         Height          =   4365
         Left            =   5490
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   570
         Width           =   255
      End
      Begin VB.TextBox ItemNumber 
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox QtyOrdered 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1230
         TabIndex        =   38
         Top             =   630
         Width           =   465
      End
      Begin VB.TextBox StdCost 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   3510
         TabIndex        =   37
         Top             =   630
         Width           =   855
      End
      Begin VB.TextBox Subtotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   630
         Width           =   975
      End
      Begin VB.TextBox ItemNumber 
         Height          =   285
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1530
         Width           =   1065
      End
      Begin VB.TextBox ItemNumber 
         Height          =   285
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2430
         Width           =   1065
      End
      Begin VB.TextBox ItemNumber 
         Height          =   285
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   3330
         Width           =   1065
      End
      Begin VB.TextBox ItemNumber 
         Height          =   285
         Index           =   4
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   4230
         Width           =   1095
      End
      Begin VB.TextBox QtyOrdered 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1230
         TabIndex        =   31
         Top             =   1530
         Width           =   465
      End
      Begin VB.TextBox QtyOrdered 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1230
         TabIndex        =   30
         Top             =   2430
         Width           =   465
      End
      Begin VB.TextBox QtyOrdered 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1230
         TabIndex        =   29
         Top             =   3330
         Width           =   465
      End
      Begin VB.TextBox QtyOrdered 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1230
         TabIndex        =   28
         Top             =   4230
         Width           =   465
      End
      Begin VB.TextBox StdCost 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3510
         TabIndex        =   27
         Top             =   1530
         Width           =   855
      End
      Begin VB.TextBox StdCost 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   3510
         TabIndex        =   26
         Top             =   2430
         Width           =   855
      End
      Begin VB.TextBox StdCost 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3510
         TabIndex        =   25
         Top             =   3330
         Width           =   855
      End
      Begin VB.TextBox StdCost 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   3510
         TabIndex        =   24
         Top             =   4230
         Width           =   855
      End
      Begin VB.TextBox Subtotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   4410
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1530
         Width           =   975
      End
      Begin VB.TextBox Subtotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   4410
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2430
         Width           =   975
      End
      Begin VB.TextBox Subtotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   4410
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   3330
         Width           =   975
      End
      Begin VB.TextBox Subtotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   4410
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   4230
         Width           =   975
      End
      Begin VB.TextBox POComment 
         Height          =   285
         Index           =   0
         Left            =   900
         TabIndex        =   19
         Top             =   960
         Width           =   4560
      End
      Begin VB.TextBox POComment 
         Height          =   285
         Index           =   1
         Left            =   900
         TabIndex        =   18
         Top             =   1860
         Width           =   4560
      End
      Begin VB.TextBox POComment 
         Height          =   285
         Index           =   2
         Left            =   900
         TabIndex        =   17
         Top             =   2760
         Width           =   4560
      End
      Begin VB.TextBox POComment 
         Height          =   285
         Index           =   3
         Left            =   900
         TabIndex        =   16
         Top             =   3660
         Width           =   4560
      End
      Begin VB.TextBox POComment 
         Height          =   285
         Index           =   4
         Left            =   900
         TabIndex        =   15
         Top             =   4560
         Width           =   4560
      End
      Begin VB.Label Label3 
         Caption         =   "Qty SO"
         Height          =   255
         Left            =   2340
         TabIndex        =   112
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "On Hand"
         Height          =   195
         Left            =   2940
         TabIndex        =   111
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Qty BO"
         Height          =   225
         Left            =   1800
         TabIndex        =   100
         Top             =   270
         Width           =   525
      End
      Begin VB.Label generalLabel 
         Caption         =   "Order Total:"
         Height          =   255
         Index           =   13
         Left            =   2580
         TabIndex        =   50
         Top             =   5100
         Width           =   975
      End
      Begin VB.Line Line2 
         X1              =   5640
         X2              =   30
         Y1              =   4950
         Y2              =   4950
      End
      Begin VB.Line Line1 
         X1              =   5640
         X2              =   30
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label generalLabel 
         Caption         =   "ItemNumber"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   48
         Top             =   300
         Width           =   975
      End
      Begin VB.Label generalLabel 
         Caption         =   "Ordered"
         Height          =   255
         Index           =   15
         Left            =   1170
         TabIndex        =   47
         Top             =   270
         Width           =   615
      End
      Begin VB.Label generalLabel 
         Caption         =   "Cost"
         Height          =   255
         Index           =   16
         Left            =   3630
         TabIndex        =   46
         Top             =   270
         Width           =   705
      End
      Begin VB.Label generalLabel 
         Caption         =   "Subtotal"
         Height          =   255
         Index           =   17
         Left            =   4440
         TabIndex        =   45
         Top             =   270
         Width           =   945
      End
      Begin VB.Label lblComment 
         Caption         =   "Comment:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   990
         Width           =   825
      End
      Begin VB.Label lblComment 
         Caption         =   "Comment:"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   42
         Top             =   1890
         Width           =   825
      End
      Begin VB.Label lblComment 
         Caption         =   "Comment:"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   41
         Top             =   2790
         Width           =   825
      End
      Begin VB.Label lblComment 
         Caption         =   "Comment:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   40
         Top             =   3690
         Width           =   825
      End
      Begin VB.Label lblComment 
         Caption         =   "Comment:"
         Height          =   255
         Index           =   4
         Left            =   150
         TabIndex        =   39
         Top             =   4590
         Width           =   825
      End
   End
   Begin VB.TextBox Comment 
      Height          =   285
      Left            =   930
      TabIndex        =   13
      Top             =   1380
      Width           =   3255
   End
   Begin VB.TextBox POAmount 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   390
      Width           =   1305
   End
   Begin VB.TextBox PONumber 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   60
      Width           =   1305
   End
   Begin VB.TextBox txtPosition 
      Height          =   285
      Left            =   3270
      TabIndex        =   3
      Top             =   7500
      Width           =   795
   End
   Begin VB.HScrollBar barPosition 
      Height          =   255
      LargeChange     =   10
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7530
      Width           =   3255
   End
   Begin VB.Frame frameShipTo 
      Caption         =   "Ship To:"
      Height          =   3135
      Left            =   6570
      TabIndex        =   1
      Top             =   3750
      Width           =   4365
      Begin VB.CheckBox chkDropshipResidential 
         Caption         =   "Residential?"
         Height          =   285
         Left            =   2820
         TabIndex        =   96
         Top             =   2520
         Width           =   1185
      End
      Begin VB.CheckBox chkDropshipLiftGate 
         Caption         =   "Lift Gate?"
         Height          =   285
         Left            =   2820
         TabIndex        =   95
         Top             =   2250
         Width           =   1185
      End
      Begin VB.TextBox DropshipCustomerEmailAddress 
         Height          =   285
         Left            =   960
         TabIndex        =   92
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox DropshipTelephoneNo 
         Height          =   285
         Left            =   960
         TabIndex        =   91
         Top             =   2460
         Width           =   1575
      End
      Begin VB.TextBox DropshipSalesOrderNo 
         Height          =   285
         Left            =   960
         TabIndex        =   89
         Top             =   2160
         Width           =   1005
      End
      Begin VB.CommandButton btnFlipAddress 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   15.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4050
         TabIndex        =   87
         Top             =   1050
         Width           =   255
      End
      Begin VB.TextBox shipAddr2 
         Height          =   285
         Left            =   960
         TabIndex        =   57
         Top             =   1260
         Width           =   3045
      End
      Begin VB.TextBox shipZipCode 
         Height          =   285
         Left            =   2580
         TabIndex        =   64
         Top             =   1860
         Width           =   1425
      End
      Begin VB.TextBox shipState 
         Height          =   285
         Left            =   960
         TabIndex        =   62
         Top             =   1860
         Width           =   765
      End
      Begin VB.TextBox shipCity 
         Height          =   285
         Left            =   960
         TabIndex        =   60
         Top             =   1560
         Width           =   3045
      End
      Begin VB.TextBox shipAddr1 
         Height          =   285
         Left            =   960
         TabIndex        =   56
         Top             =   960
         Width           =   3045
      End
      Begin VB.TextBox shipName 
         Height          =   285
         Left            =   960
         TabIndex        =   55
         Top             =   660
         Width           =   3045
      End
      Begin VB.ComboBox cmbWhseCodes 
         Height          =   315
         Left            =   180
         TabIndex        =   53
         Top             =   270
         Width           =   4095
      End
      Begin VB.Label generalLabel 
         Caption         =   "Email"
         Height          =   255
         Index           =   21
         Left            =   150
         TabIndex        =   93
         Top             =   2790
         Width           =   675
      End
      Begin VB.Label generalLabel 
         Caption         =   "Phone #"
         Height          =   255
         Index           =   20
         Left            =   150
         TabIndex        =   90
         Top             =   2490
         Width           =   675
      End
      Begin VB.Label generalLabel 
         Caption         =   "SO #"
         Height          =   255
         Index           =   19
         Left            =   150
         TabIndex        =   88
         Top             =   2190
         Width           =   675
      End
      Begin VB.Label generalLabel 
         Caption         =   "Zip Code"
         Height          =   255
         Index           =   12
         Left            =   1830
         TabIndex        =   63
         Top             =   1890
         Width           =   675
      End
      Begin VB.Label generalLabel 
         Caption         =   "State"
         Height          =   255
         Index           =   11
         Left            =   150
         TabIndex        =   61
         Top             =   1890
         Width           =   675
      End
      Begin VB.Label generalLabel 
         Caption         =   "City"
         Height          =   255
         Index           =   10
         Left            =   150
         TabIndex        =   59
         Top             =   1590
         Width           =   675
      End
      Begin VB.Label generalLabel 
         Caption         =   "Address"
         Height          =   255
         Index           =   9
         Left            =   150
         TabIndex        =   58
         Top             =   990
         Width           =   675
      End
      Begin VB.Label generalLabel 
         Caption         =   "Name"
         Height          =   255
         Index           =   8
         Left            =   150
         TabIndex        =   54
         Top             =   690
         Width           =   675
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   10080
      TabIndex        =   0
      Top             =   6990
      Width           =   885
   End
   Begin VB.Label generalLabel 
      Caption         =   "Notes:"
      Height          =   255
      Index           =   18
      Left            =   6270
      TabIndex        =   84
      Top             =   3150
      Width           =   525
   End
   Begin VB.Label generalLabel 
      Caption         =   "Est Received By:"
      Height          =   255
      Index           =   7
      Left            =   9540
      TabIndex        =   79
      Top             =   2430
      Width           =   1335
   End
   Begin VB.Label generalLabel 
      Caption         =   "# of Payments:"
      Height          =   255
      Index           =   6
      Left            =   5670
      TabIndex        =   77
      Top             =   390
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label generalLabel 
      Caption         =   "POTerms:"
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   73
      Top             =   60
      Width           =   795
   End
   Begin VB.Label lblFinalized 
      Alignment       =   2  'Center
      Caption         =   "FINALIZED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6495
      TabIndex        =   65
      Top             =   2370
      Width           =   2265
   End
   Begin VB.Label generalLabel 
      Caption         =   "Comment:"
      Height          =   255
      Index           =   4
      Left            =   90
      TabIndex        =   12
      Top             =   1380
      Width           =   825
   End
   Begin VB.Label generalLabel 
      Caption         =   "Purchase Address:"
      Height          =   255
      Index           =   3
      Left            =   90
      TabIndex        =   11
      Top             =   1050
      Width           =   1425
   End
   Begin VB.Label generalLabel 
      Caption         =   "Vendor:"
      Height          =   255
      Index           =   2
      Left            =   90
      TabIndex        =   10
      Top             =   720
      Width           =   1425
   End
   Begin VB.Label generalLabel 
      Caption         =   "PO Amount:"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   8
      Top             =   390
      Width           =   1425
   End
   Begin VB.Label generalLabel 
      Caption         =   "PO Number:"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   60
      Width           =   1425
   End
   Begin VB.Label lblMaxAmt 
      Caption         =   "of MAX"
      Height          =   195
      Left            =   4110
      TabIndex        =   5
      Top             =   7530
      Width           =   975
   End
   Begin VB.Line Line3 
      X1              =   -390
      X2              =   10620
      Y1              =   7470
      Y2              =   7470
   End
   Begin VB.Label lblStatusBar 
      Height          =   285
      Left            =   5400
      TabIndex        =   4
      Top             =   7500
      Width           =   5565
   End
End
Attribute VB_Name = "PurchaseOrdersForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private POList() As String
Private ItemList() As String
Private changed As Boolean
Private whichCtl As String
Private oldItemPos As Long
Private fillingForm As Boolean

Public Sub LockForExport(tf As Boolean)
    If tf Then
        Me.lblStatusBar.Caption = "EXPORT RUNNING, FINALIZED POs LOCKED!!"
        Me.lblStatusBar.FontBold = True
        Me.lblStatusBar.ForeColor = 255
        Me.btnFinalize.Enabled = False
    Else
        Me.lblStatusBar.Caption = ""
        Me.lblStatusBar.FontBold = False
        Me.lblStatusBar.ForeColor = 0
        Me.btnFinalize.Enabled = True
    End If
End Sub

Private Sub barPosition_Change()
    Mouse.Hourglass True
    If changed = True Then
        Select Case whichCtl
            Case Is = "QtyOrdered(0)"
                QtyOrdered_LostFocus 0
            Case Is = "QtyOrdered(1)"
                QtyOrdered_LostFocus 1
            Case Is = "QtyOrdered(2)"
                QtyOrdered_LostFocus 2
            Case Is = "QtyOrdered(3)"
                QtyOrdered_LostFocus 3
            Case Is = "QtyOrdered(4)"
                QtyOrdered_LostFocus 4
            Case Is = "StdCost(0)"
                StdCost_LostFocus 0
            Case Is = "StdCost(1)"
                StdCost_LostFocus 1
            Case Is = "StdCost(2)"
                StdCost_LostFocus 2
            Case Is = "StdCost(3)"
                StdCost_LostFocus 3
            Case Is = "StdCost(4)"
                StdCost_LostFocus 4
            Case Is = "POComment(0)"
                POComment_LostFocus 0
            Case Is = "POComment(1)"
                POComment_LostFocus 1
            Case Is = "POComment(2)"
                POComment_LostFocus 2
            Case Is = "POComment(3)"
                POComment_LostFocus 3
            Case Is = "POComment(4)"
                POComment_LostFocus 4
            Case Is = "Comment"
                Comment_LostFocus
            Case Is = "NumPayments"
                NumPayments_LostFocus
            Case Is = "ShipName"
                shipName_LostFocus
            Case Is = "ShipAddr1"
                shipAddr1_LostFocus
            Case Is = "ShipAddr2"
                shipAddr2_LostFocus
            Case Is = "ShipCity"
                shipCity_LostFocus
            Case Is = "ShipState"
                shipState_LostFocus
            Case Is = "ShipZipCode"
                shipZipCode_LostFocus
            Case Is = "DropshipSalesOrderNo"
                DropshipSalesOrderNo_LostFocus
            Case Is = "DropshipTelephoneNo"
                DropshipTelephoneNo_LostFocus
            Case Is = "DropshipCustomerEmailAddress"
                DropshipCustomerEmailAddress_LostFocus
            Case Is = "MiscNotes"
                MiscNotes_LostFocus
            Case Else
                MsgBox "Couldn't figure out which control you just updated, so not saving..."
        End Select
    End If
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("exec spPurchOrderByID '" & POList(Me.barPosition.value) & "'")
    If rst.EOF Then
        MsgBox "Record " & POList(Me.barPosition.value) & " deleted by another user? requerying..."
        Dim POID As String
        If Me.barPosition.value = Me.barPosition.Min Then
            POID = POList(Me.barPosition.value + 1)
        Else
            POID = POList(Me.barPosition.value - 1)
        End If
        requeryForm "All"
        Me.barPosition.value = bsearch(POID, POList)
    Else
        fillForm rst
        Me.txtPosition = Me.barPosition.value + 1
    End If
    rst.Close
    Set rst = Nothing
    changed = False
    Mouse.Hourglass False
End Sub

Private Sub btnDeletePO_Click()
    If MsgBox("Are you sure you want to delete this PO?", vbYesNo + vbDefaultButton2) = vbYes Then
        DB.Execute "DELETE FROM PurchaseOrders WHERE ID=" & Me.POID
        'DB.Execute "UPDATE InventoryMaster SET POID=Null, QtyOrdered=0 WHERE POID=" & Me.POID
        DB.Execute "DELETE FROM PurchaseOrderLines WHERE HeaderID=" & Me.POID
        Dim id As String
        If Me.barPosition.value = Me.barPosition.Min Then
            id = POList(Me.barPosition.value + 1)
        Else
            id = POList(Me.barPosition.value - 1)
        End If
        requeryForm "All"
        Dim pos As Long
        pos = bsearch(id, POList)
        If pos = -1 Then
            MsgBox "Error seeking, you should close the PO window and then reopen it, otherwise things might be broken"
        Else
            Me.barPosition.value = pos
        End If
    End If
End Sub

Private Sub btnDropshipify_Click()
'deprecated by Logic.ConvertToDropshipPO(), but still used a/o 2013-03-05 when
'margie enters the wrong S/O number and needs to fix
    Dim soNum As String
    soNum = InputBox("Enter Sales Order number:")
    If soNum = "" Then
        Exit Sub
    End If
    soNum = UCase(soNum)
    If IsNumeric(Mid(soNum, 1, 1)) Then
        While Len(soNum) < 7
            soNum = "0" & soNum
        Wend
    End If

    Dim m As Mas200ImportExport
    Set m = New Mas200ImportExport
    Dim so As Variant
    so = m.GetSalesOrderInfo(soNum)
    Set m = Nothing

    If IsNull(so) Then
        MsgBox "Can't find sales order " + soNum + " in Mas S/O Entry!"
        Exit Sub
    End If

    If CDbl(so(8)) = 0 Then
        MsgBox "Warning: no deposit has been taken for this order!" & vbCrLf & vbCrLf & "TODO: brian should probably check the payment type here too."
    End If
    
    Me.cmbWhseCodes = "Drop Ship Address/Other"
    cmbWhseCodes_LostFocus
    Me.shipName = so(1)
    changed = True
    shipName_LostFocus
    Me.shipAddr1 = so(2)
    changed = True
    shipAddr1_LostFocus
    Me.shipAddr2 = so(3)
    changed = True
    shipAddr2_LostFocus
    Me.shipCity = so(4)
    changed = True
    shipCity_LostFocus
    Me.shipState = so(5)
    changed = True
    shipState_LostFocus
    Me.shipZipCode = so(6)
    changed = True
    shipZipCode_LostFocus
    
    'Me.MiscNotes = "CUSTOMER PHONE NUMBER " & so(7) & vbCrLf & _
    '               "ORDER NUMBER " + soNum
    ' margie doesn't want the phone number entered, always manual?
    'Me.MiscNotes = "CUSTOMER PHONE NUMBER " & vbCrLf & _
    '               "ORDER NUMBER " + soNum
    'changed = True
    'MiscNotes_LostFocus
    Me.DropshipSalesOrderNo = soNum
    changed = True
    DropshipSalesOrderNo_LostFocus
    Me.DropshipTelephoneNo = so(7)
    changed = True
    DropshipTelephoneNo_LostFocus
    Me.DropshipCustomerEmailAddress = so(9)
    changed = True
    DropshipCustomerEmailAddress_LostFocus
    Me.chkDropshipLiftGate = 1
    Me.chkDropshipResidential = 1
End Sub

Private Sub btnExcelExportPO_Click()
    If DLookup("PONumber", "PurchaseOrders", "ID=" & Me.POID) = "" Then
        Me.PONumber = CreatePONumber(Me.POID)
    End If
On Error GoTo errh
    Dim excelapp As Excel.Application
    Set excelapp = New Excel.Application
    excelapp.Workbooks.Add
    
    Dim rownum As Long
    
    rownum = 1
    excelapp.ActiveWorkbook.ActiveSheet.Range("A" & rownum).value = "PO #:"
    excelapp.ActiveWorkbook.ActiveSheet.Range("B" & rownum).value = Me.PONumber
    rownum = rownum + 1
    
    excelapp.ActiveWorkbook.ActiveSheet.Range("A" & rownum).value = "Freight Free:"
    excelapp.ActiveWorkbook.ActiveSheet.Range("B" & rownum).value = IIf(Me.chkFreightFree, "Yes", "No")
    rownum = rownum + 1
    rownum = rownum + 1
    
    excelapp.ActiveWorkbook.ActiveSheet.Range("A" & rownum).value = "Ship To:"
    excelapp.ActiveWorkbook.ActiveSheet.Range("B" & rownum).value = Me.shipName
    rownum = rownum + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range("B" & rownum).value = Me.shipAddr1
    rownum = rownum + 1
    If Me.shipAddr2 <> "" Then
        excelapp.ActiveWorkbook.ActiveSheet.Range("B" & rownum).value = Me.shipAddr2
        rownum = rownum + 1
    End If
    excelapp.ActiveWorkbook.ActiveSheet.Range("B" & rownum).value = Me.shipCity & ", " & Me.shipState & " " & Me.shipZipCode
    rownum = rownum + 1
    
    If Me.DropshipSalesOrderNo <> "" Then
        excelapp.ActiveWorkbook.ActiveSheet.Range("A" & rownum).value = "Customer Order #:"
        excelapp.ActiveWorkbook.ActiveSheet.Range("B" & rownum).value = Me.DropshipSalesOrderNo
        rownum = rownum + 1
    End If
    
    If Me.DropshipTelephoneNo <> "" Then
        excelapp.ActiveWorkbook.ActiveSheet.Range("A" & rownum).value = "Customer Telephone #:"
        excelapp.ActiveWorkbook.ActiveSheet.Range("B" & rownum).value = Me.DropshipTelephoneNo
        rownum = rownum + 1
    End If
    
    If Me.MiscNotes <> "" Then
        rownum = rownum + 1
        excelapp.ActiveWorkbook.ActiveSheet.Range("A" & rownum).value = "Notes:"
        Dim noteLines As Variant
        noteLines = Split(Me.MiscNotes, vbCrLf)
        Dim noteIter As Variant
        For Each noteIter In noteLines
            excelapp.ActiveWorkbook.ActiveSheet.Range("B" & rownum).value = CStr(noteIter)
            rownum = rownum + 1
        Next noteIter
    End If
    
    rownum = rownum + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range("A" & rownum).value = "Product Line"
    excelapp.ActiveWorkbook.ActiveSheet.Range("B" & rownum).value = "Item Number"
    excelapp.ActiveWorkbook.ActiveSheet.Range("C" & rownum).value = "Description"
    excelapp.ActiveWorkbook.ActiveSheet.Range("D" & rownum).value = "Qty Ordered"
    excelapp.ActiveWorkbook.ActiveSheet.Range("E" & rownum).value = "Unit Cost"
    excelapp.ActiveWorkbook.ActiveSheet.Range("F" & rownum).value = "Subtotal"
    excelapp.ActiveWorkbook.ActiveSheet.Range("G" & rownum).value = "Comments"
    rownum = rownum + 1
    Dim rst As ADODB.Recordset
    'Set rst = DB.retrieve("SELECT ItemNumber, ItemDescription, QtyOrdered, StdCost, POComment FROM InventoryMaster WHERE POID=" & Me.POID & " ORDER BY ItemNumber")
    Set rst = DB.retrieve("SELECT InventoryMaster.ItemNumber, ItemDescription, Quantity, StdCost, POComment FROM InventoryMaster INNER JOIN PurchaseOrderLines ON InventoryMaster.ItemNumber=PurchaseOrderLines.ItemNumber WHERE HeaderID=" & Me.POID & " ORDER BY InventoryMaster.ItemNumber")
    While Not rst.EOF
        excelapp.ActiveWorkbook.ActiveSheet.Range("A" & rownum).value = Left(rst("ItemNumber"), 3)
        excelapp.ActiveWorkbook.ActiveSheet.Range("B" & rownum).value = Mid(rst("ItemNumber"), 4)
        excelapp.ActiveWorkbook.ActiveSheet.Range("C" & rownum).value = rst("ItemDescription")
        'excelapp.ActiveWorkbook.ActiveSheet.Range("D" & rownum).value = rst("QtyOrdered")
        excelapp.ActiveWorkbook.ActiveSheet.Range("D" & rownum).value = rst("Quantity")
        excelapp.ActiveWorkbook.ActiveSheet.Range("E" & rownum).value = rst("StdCost")
        'excelapp.ActiveWorkbook.ActiveSheet.Range("F" & rownum).value = rst("QtyOrdered") * rst("StdCost")
        excelapp.ActiveWorkbook.ActiveSheet.Range("F" & rownum).value = rst("Quantity") * rst("StdCost")
        excelapp.ActiveWorkbook.ActiveSheet.Range("G" & rownum).value = Nz(rst("POComment"))
        rownum = rownum + 1
        
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    excelapp.ActiveWorkbook.ActiveSheet.Range("E" & rownum).value = "Total:"
    excelapp.ActiveWorkbook.ActiveSheet.Range("F" & rownum).value = "=SUM(F4:F" & (rownum - 1) & ")"
    
    excelapp.ActiveWorkbook.ActiveSheet.Columns("B").HorizontalAlignment = xlLeft
    excelapp.ActiveWorkbook.ActiveSheet.Columns("E:F").NumberFormat = "$#,##0.00"
    excelapp.ActiveWorkbook.ActiveSheet.rows("3:3").NumberFormat = "General"
    excelapp.ActiveWorkbook.ActiveSheet.cells.EntireColumn.AutoFit
    excelapp.ActiveWorkbook.ActiveSheet.Range("A1").Select
    
    excelapp.Visible = True
    Set excelapp = Nothing
    Exit Sub
errh:
    MsgBox "ERROR, Excel might be messed up until you reboot: " & vbCrLf & vbCrLf & Err.Number & vbCrLf & vbCrLf & Err.Description
    If Not excelapp Is Nothing Then
        excelapp.Visible = True
        Set excelapp = Nothing
    End If
End Sub

Private Sub btnExit_Click()
    Unload PurchaseOrdersForm
End Sub

Private Sub btnFinalize_Click()
    Mouse.Hourglass True
    If Me.btnFinalize.Caption = "Finalize PO" Then
        If Me.chkFreightFree Or Me.chkFreightPay Then
            updatePurchaseOrders "Finalized", 1, Me.POID, ""
            updatePurchaseOrders "POAmount", Replace(Me.POAmount, ",", ""), Me.POID, ""
            Me.btnFinalize.Caption = "Un-Finalize PO"
            enableItemsSubform False
            Me.btnPreviewPO.Enabled = True
            handleFinalizedText True
            Me.VendorNumber.Enabled = True
            Me.VendorNumber.Locked = False
            Me.VendorPOAddr.Enabled = True
            Me.VendorPOAddr.Locked = False
            If DLookup("PONumber", "PurchaseOrders", "ID=" & Me.POID) = "" Then
                Me.PONumber = CreatePONumber(Me.POID)
            End If
        Else
            MsgBox "You need to select a freight payment choice."
        End If
    Else
        updatePurchaseOrders "Finalized", "0", Me.POID, ""
        updatePurchaseOrders "POAmount", "Null", Me.POID, ""
        Me.btnFinalize.Caption = "Finalize PO"
        enableItemsSubform True
        Me.btnPreviewPO.Enabled = True
        handleFinalizedText False
        Me.VendorNumber.Enabled = False
        Me.VendorNumber.Locked = True
        Me.VendorPOAddr.Enabled = False
        Me.VendorPOAddr.Locked = True
    End If
    Mouse.Hourglass False
End Sub

Private Sub btnFlipAddress_Click()
    If Me.shipAddr1.Enabled = False Then
        MsgBox "Can't edit this address!"
        Exit Sub
    End If
    
    Dim temp As String
    temp = Me.shipAddr1
    Me.shipAddr1 = Me.shipAddr2
    changed = True
    shipAddr1_LostFocus
    Me.shipAddr2 = temp
    changed = True
    shipAddr2_LostFocus
End Sub

Private Sub btnLineCodeSubCfg_Click()
    Load LCSubConfig
    LCSubConfig.tableName = "ProductLineSubForPOs"
    LCSubConfig.Show MODAL
End Sub

Private Sub btnPreviewPO_Click()
    If DLookup("PONumber", "PurchaseOrders", "ID=" & Me.POID) = "" Then
        Me.PONumber = CreatePONumber(Me.POID)
    End If
    OpenPurchaseOrderReport Me.POID
End Sub

Private Sub btnPreviewPONew_Click()
    If DLookup("PONumber", "PurchaseOrders", "ID=" & Me.POID) = "" Then
        Me.PONumber = CreatePONumber(Me.POID)
    End If
On Error GoTo errh
    Shell "s:\mastest\mas90-signs\purchasing\PrintPO\PrintPO.exe " & Me.POID, vbNormalFocus
    Exit Sub
errh:
    MsgBox "Margie, don't use this until it's ready."
    Err.Clear
    Exit Sub
End Sub

Private Sub btnSendReadyNotification_Click()
    Dim msg As String
    msg = Me.VendorNumber & " PO #" & IIf(Me.PONumber = "", "<not assigned>", Me.PONumber) & " is ready to go."
    If SendEmailTo("sherri@tools-plus.com", msg, msg & IIf(Me.ReadyMessage <> "", vbCrLf & vbCrLf & Me.ReadyMessage, "")) Then
        Me.btnSendReadyNotification.BackColor = GREEN
    Else
        Me.btnSendReadyNotification.BackColor = RED
    End If
End Sub

Private Sub cmbWhseCodes_Click()
    cmbWhseCodes_LostFocus
End Sub

Private Sub cmbWhseCodes_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.cmbWhseCodes, KeyCode, Shift
End Sub

Private Sub cmbWhseCodes_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.cmbWhseCodes, KeyAscii
    If KeyAscii = vbKeyReturn Then
        cmbWhseCodes_LostFocus
    End If
End Sub

Private Sub cmbWhseCodes_LostFocus()
    AutoCompleteLostFocus Me.cmbWhseCodes
    If Me.cmbWhseCodes = "Drop Ship Address/Other" Then
        If Me.shipName.Enabled = True Then
            'do nothing
        Else
            clearShipping
        End If
    Else
        fillShippingFromCombo
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageUp Then
        If Me.barPosition <> Me.barPosition.Min Then
            Me.barPosition = Me.barPosition - 1
        End If
    ElseIf KeyCode = vbKeyPageDown Then
        If Me.barPosition <> Me.barPosition.max Then
            Me.barPosition = Me.barPosition + 1
        End If
    ElseIf KeyCode = vbKeyF8 And Me.barPosition.value <> Me.barPosition.Min Then
        Me.barPosition = Me.barPosition - 1
    ElseIf KeyCode = vbKeyF9 And Me.barPosition.value <> Me.barPosition.max Then
        Me.barPosition = Me.barPosition + 1
    End If
End Sub

Private Sub Form_Load()
    Mouse.Hourglass True
    If EXPORT_GOING Then
        LockForExport True
    End If
    requeryVendorNumber
    requeryWhseCodes
    requeryPOTerms
    requeryDeferredList
    requeryForm "All"
    Me.barPosition.value = 0
    barPosition_Change
    Mouse.Hourglass False
    
    Dim i As Integer
    For i = 0 To 4
        Me.QtyOnHand(i).Enabled = False
        Me.QtyBackOrdered(i).Enabled = False
        Me.QtyOnSO(i).Enabled = False
   Next
End Sub

Private Sub itemScrollBar_Change()
    Mouse.Hourglass True
    If ItemList(0) <> "NORECORDS" Then
        If Abs(Me.itemScrollBar.value - oldItemPos) > 1 Then 'drag up or down
            fillItemSubform Me.itemScrollBar.value
        ElseIf Me.itemScrollBar.value > oldItemPos Then 'down arrow
            shuffleUp
            If Not addItemAfter(Me.ItemNumber(3)) Then
                Me.itemScrollBar = 0
                fillItemSubform
            End If
        Else 'uparrow
            shuffleDown
            If Not addItemBefore(Me.ItemNumber(1)) Then
                Me.itemScrollBar = 0
                fillItemSubform
            End If
        End If
        oldItemPos = Me.itemScrollBar.value
    End If
    Mouse.Hourglass False
End Sub









Private Sub requeryWhseCodes()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ShipToCode, ShipToCodeDesc FROM PO_ShipToAddress")
    Me.cmbWhseCodes.Clear
    While Not rst.EOF
        If rst("ShipToCode") <> "1234" Then
            Me.cmbWhseCodes.AddItem (Trim$(rst("ShipToCodeDesc")))
        End If
        rst.MoveNext
    Wend
    Me.cmbWhseCodes.AddItem ("Drop Ship Address/Other")
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requeryVendorNumber()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT DISTINCT AP.VendorName, AP.VendorNo FROM ProductLine AS PL, AP_Vendor AS AP WHERE PL.PrimaryVendorNumber=AP.VendorNo ORDER BY AP.VendorName")
    Me.VendorNumber.Clear
    While Not rst.EOF
        Me.VendorNumber.AddItem CStr(rst("VendorNo"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requeryVendorPOAddress()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT PurchaseAddressCode FROM PO_VendorPurchaseAddress WHERE VendorNo='" & lookupPrimaryVendorNumber(Me.VendorNumber) & "'")
    Me.VendorPOAddr.Clear
    While Not rst.EOF
        Me.VendorPOAddr.AddItem CStr(rst("PurchaseAddressCode"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requeryPOTerms()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT TermsCode, TermsCodeDesc, DiscountRate FROM AP_TermsCode")
    Me.POTerms.Clear
    While Not rst.EOF
        Me.POTerms.AddItem (Trim$(rst("TermsCodeDesc")))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requeryDeferredList()
    Me.DeferredList.Clear
    Me.DeferredList.AddItem ("30")
    Me.DeferredList.AddItem ("60")
    Me.DeferredList.AddItem ("90")
    Me.DeferredList.AddItem ("120")
End Sub

Private Sub requeryForm(filter As String)
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Select Case filter
        Case Is = "All"
            Set rst = DB.retrieve("SELECT ID FROM PurchaseOrders WHERE Exported=0")
        Case Else
            MsgBox "Unknown filter type """ & filter & """, defaulting to ""All""."
            Set rst = DB.retrieve("SELECT ID FROM PurchaseOrders")
    End Select
    If rst.RecordCount > 0 Then
        ReDim POList(rst.RecordCount - 1) As String
        Dim i As Long
        For i = 0 To rst.RecordCount - 1
            POList(i) = rst("ID")
            rst.MoveNext
        Next i
        Me.lblMaxAmt.Caption = "of " & rst.RecordCount
        Me.barPosition.Min = 0
        Me.barPosition.max = rst.RecordCount - 1
    Else
        ReDim POList(0) As String
        POList(0) = "NORECORDS"
        Me.lblMaxAmt.Caption = "of 0"
        Me.barPosition.Min = 0
        Me.barPosition.max = 0
    End If
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requeryItems(POID As String)
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    'Set rst = DB.retrieve("SELECT ItemNumber FROM InventoryMaster WHERE POID=" & POID & " ORDER BY ItemNumber")
    Set rst = DB.retrieve("SELECT ItemNumber FROM PurchaseOrderLines WHERE HeaderID=" & POID & " ORDER BY ItemNumber")
    If rst.RecordCount > 0 Then
        ReDim ItemList(rst.RecordCount - 1) As String
        Dim i As Long
        For i = 0 To rst.RecordCount - 1
            ItemList(i) = rst("ItemNumber")
            rst.MoveNext
        Next i
        Me.itemScrollBar.Min = 0
        Me.itemScrollBar.max = IIf(rst.RecordCount - 5 < 0, 0, rst.RecordCount - 5)
    Else
        ReDim ItemList(0) As String
        ItemList(0) = "NORECORDS"
        Me.itemScrollBar.Min = 0
        Me.itemScrollBar.max = 0
    End If
    oldItemPos = 0
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub fillForm(PO As ADODB.Recordset)
    Mouse.Hourglass True
    Me.POID = PO("ID")
    Me.PONumber = Nz(PO("PONumber"))
    'Me.POAmount = Format(po("POAmount"), "Currency")
    Me.VendorNumber = PO("VendorNumber")
    requeryVendorPOAddress
    Me.VendorPOAddr = Nz(PO("VendorPOAddr"))
    Me.Comment = Nz(PO("POComment"))
    requeryItems PO("ID")
    fillItemSubform
    fillOrderTotal PO("ID")
    fillShipping PO
    If PO("Finalized") Then
        Me.btnFinalize.Caption = "Un-Finalize PO"
        enableItemsSubform False
        Me.VendorNumber.Locked = False
        Me.VendorNumber.Enabled = True
        Me.VendorPOAddr.Locked = False
        Me.VendorPOAddr.Enabled = True
        handleFinalizedText True
    Else
        Me.btnFinalize.Caption = "Finalize PO"
        enableItemsSubform True
        Me.VendorNumber.Locked = True
        Me.VendorNumber.Enabled = False
        Me.VendorPOAddr.Locked = True
        Me.VendorPOAddr.Enabled = False
        handleFinalizedText False
    End If
    fillingForm = True 'flags checkboxes to not run click event
    If PO("FreightFree") Then
        Me.chkFreightFree.value = True
    Else
        Me.chkFreightPay.value = True
    End If
    Me.chkDeferred = SQLBool(PO("Deferrd"))
    If Me.chkDeferred Then
        Enable Me.DeferredList
    Else
        Disable Me.DeferredList
    End If
    fillingForm = False
    Me.DeferredList = Nz(PO("DeferredTime"))
    Me.POTerms = lookupPOTermDescByNum(PO("POTerms"))
    If PO("POTerms") = "00" Then
        Enable Me.NumPayments
        Me.NumPayments = PO("NumberOfPayments")
    Else
        Disable Me.NumPayments
        Me.NumPayments = ""
    End If
    Me.PODueDate = Format(PO("DueDate"), "Short Date")
    Me.PODueDateCal.value = Format(PO("DueDate"), "Short Date")
    Me.MiscNotes = Nz(PO("MiscNotes"))
    
    Me.btnSendReadyNotification.BackColor = &H8000000F
    
    Mouse.Hourglass False
End Sub

Private Sub fillShippingFromCombo()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ShipToCode, ShipToName, ShipToAddress1, ShipToAddress2, ShipToCity, ShipToState, ShipToZipCode FROM PO_ShipToAddress WHERE ShipToCodeDesc='" & Me.cmbWhseCodes & "'")
    Me.shipName = CStr(rst("POName"))
    Me.shipAddr1 = CStr(rst("POAddress1"))
    Me.shipAddr2 = CStr(rst("POAddress2"))
    Me.shipCity = CStr(rst("POCity"))
    Me.shipState = CStr(rst("POState"))
    Me.shipZipCode = CStr(rst("POZipCode"))
    'note: this keeps sales order and other dropship info...is that correct?
    DB.Execute "UPDATE PurchaseOrders SET ShipToCode='" & CStr(rst("ShipToCode")) & _
                                      "', ShipToName='" & Me.shipName & _
                                      "', ShipToAddress1='" & Me.shipAddr1 & _
                                      "', ShipToAddress2='" & Me.shipAddr2 & _
                                      "', ShipToCity='" & Me.shipCity & _
                                      "', ShipToState='" & Me.shipState & _
                                      "', ShipToZip='" & Me.shipZipCode & _
                                      "' WHERE ID=" & Me.POID
    enableAllShipping False
End Sub

Private Sub clearShipping()
    Me.shipName = ""
    Me.shipAddr1 = ""
    Me.shipAddr2 = ""
    Me.shipCity = ""
    Me.shipState = ""
    Me.shipZipCode = ""
    'ship code 1234 for d/s is also used when calculating time-to-po, etc.
    'note: this keeps sales order and other dropship info...is that correct?
    DB.Execute "UPDATE PurchaseOrders SET ShipToCode='1234" & _
                                      "', ShipToName='" & Me.shipName & _
                                      "', ShipToAddress1='" & Me.shipAddr1 & _
                                      "', ShipToAddress2='" & Me.shipAddr2 & _
                                      "', ShipToCity='" & Me.shipCity & _
                                      "', ShipToState='" & Me.shipState & _
                                      "', ShipToZip='" & Me.shipZipCode & _
                                      "' WHERE ID=" & Me.POID
    enableAllShipping True
End Sub

Private Sub fillShipping(rst As ADODB.Recordset)
    Dim rst1 As ADODB.Recordset
    Set rst1 = DB.retrieve("SELECT * FROM PO_ShipToAddress WHERE ShipToCode='" & rst("ShipToCode") & "'")
    If rst1.EOF Then
        Me.cmbWhseCodes = "Drop Ship Address/Other"
        enableAllShipping True
    ElseIf rst("ShipToCode") = "1234" Then
        Me.cmbWhseCodes = "Drop Ship Address/Other"
        enableAllShipping True
    Else
        Me.cmbWhseCodes = Trim(rst1("ShipToCodeDesc"))
        enableAllShipping False
    End If
    rst1.Close
    Set rst1 = Nothing
    'If rst("ShipToCode") = "0000" Then
    '    Me.cmbWhseCodes = "Tools Plus"
    '    enableAllShipping False
    'ElseIf rst("ShipToCode") = "0001" Then
    '    Me.cmbWhseCodes = "Central warehouse"
    '    enableAllShipping False
    'ElseIf rst("ShipToCode") = "0002" Then
    '    Me.cmbWhseCodes = "Thomaston Avenue Warehouse"
    '    enableAllShipping False
    'Else
    Me.shipName = Nz(rst("ShipToName"))
    Me.shipAddr1 = Nz(rst("ShipToAddress1"))
    Me.shipAddr2 = Nz(rst("ShipToAddress2"))
    Me.shipCity = Nz(rst("ShipToCity"))
    Me.shipState = Nz(rst("ShipToState"))
    Me.shipZipCode = Nz(rst("ShipToZip"))
    fillingForm = True
    If rst("ShipToCode") = "1234" Then
        Me.DropshipSalesOrderNo = Nz(rst("SalesOrderNo"))
        Me.DropshipTelephoneNo = Nz(rst("ShipToTelephoneNo"))
        Me.DropshipCustomerEmailAddress = Nz(rst("CustomerEmailAddress"))
        Me.chkDropshipLiftGate = SQLBool(Nz(rst("LiftGateRequired"), False))
        Me.chkDropshipResidential = SQLBool(Nz(rst("AddressZoningType"), False))
    Else
        Me.DropshipSalesOrderNo = ""
        Me.DropshipTelephoneNo = ""
        Me.DropshipCustomerEmailAddress = ""
        Me.chkDropshipLiftGate = 0
        Me.chkDropshipResidential = 0
    End If
    fillingForm = False
End Sub

Private Sub enableAllShipping(tf As Boolean)
    If tf = True Then
        Enable Me.shipName
        Enable Me.shipAddr1
        Enable Me.shipAddr2
        Enable Me.shipCity
        Enable Me.shipState
        Enable Me.shipZipCode
        Enable Me.DropshipSalesOrderNo
        Enable Me.DropshipTelephoneNo
        Enable Me.DropshipCustomerEmailAddress
        Enable Me.chkDropshipLiftGate
        Enable Me.chkDropshipResidential
    Else
        Disable Me.shipName
        Disable Me.shipAddr1
        Disable Me.shipAddr2
        Disable Me.shipCity
        Disable Me.shipState
        Disable Me.shipZipCode
        Disable Me.DropshipSalesOrderNo
        Disable Me.DropshipTelephoneNo
        Disable Me.DropshipCustomerEmailAddress
        Disable Me.chkDropshipLiftGate
        Disable Me.chkDropshipResidential
    End If
End Sub

Private Sub enableItemsSubform(tf As Boolean)
    Dim i As Long
    If tf = True Then
        For i = 0 To 4
            Me.ItemNumber(i).Enabled = True
            Me.QtyOrdered(i).Enabled = True
            Me.QtyBackOrdered(i).Enabled = False
            Me.QtyOnHand(i).Enabled = False
            Me.QtyOnSO(i).Enabled = False
            Me.StdCost(i).Enabled = True
            Me.Subtotal(i).Enabled = True
            Me.POComment(i).Enabled = True
        Next i
    Else
        For i = 0 To 4
            Me.ItemNumber(i).Enabled = False
            Me.QtyOrdered(i).Enabled = False
            Me.QtyBackOrdered(i).Enabled = False
            Me.QtyOnHand(i).Enabled = False
            Me.QtyOnSO(i).Enabled = False
            Me.StdCost(i).Enabled = False
            Me.Subtotal(i).Enabled = False
            Me.POComment(i).Enabled = False
        Next i
    End If
End Sub

Private Sub fillOrderTotal(POID As Long)
    Dim rst As ADODB.Recordset
    'Set rst = DB.retrieve("SELECT SUM(QtyOrdered*StdCost) AS OrderTotal FROM InventoryMaster WHERE POID=" & POID)
    Set rst = DB.retrieve("SELECT SUM(Quantity*StdCost) AS OrderTotal FROM InventoryMaster INNER JOIN PurchaseOrderLines ON InventoryMaster.ItemNumber=PurchaseOrderLines.ItemNumber WHERE HeaderID=" & POID)
    Me.OrderTotal = IIf(IsNull(rst("OrderTotal")), Format(0, "Currency"), Format(rst("OrderTotal"), "Currency"))
    rst.Close
    Set rst = Nothing
    Me.POAmount = Me.OrderTotal
End Sub

Private Sub fillItemSubform(Optional offset As Long = 0)
    Dim rst As ADODB.Recordset
    Dim i As Long
    For i = 0 To 4
        If UBound(ItemList) >= i Then
            If ItemList(offset + i) <> "NORECORDS" Then
                'Set rst = DB.retrieve("SELECT ItemNumber, QtyOrdered, StdCost, POComment FROM InventoryMaster WHERE ItemNumber='" & ItemList(offset + i) & "'")
                'Set rst = DB.retrieve("SELECT InventoryMaster.ItemNumber, InventoryMaster.StdCost, InventoryMaster.POComment, PurchaseOrderLines.Quantity FROM InventoryMaster INNER JOIN PurchaseOrderLines ON InventoryMaster.ItemNumber=PurchaseOrderLines.ItemNumber AND PurchaseOrderLines.HeaderID=" & Me.POID & " WHERE InventoryMaster.ItemNumber='" & ItemList(offset + i) & "'")
                Set rst = DB.retrieve("SELECT InventoryMaster.ItemNumber, InventoryMaster.StdCost, InventoryMaster.POComment, PurchaseOrderLines.Quantity, vActualWhseQty.QtyOnSO, vActualWhseQty.QtyOnBO, vActualWhseQty.QtyOnHand FROM InventoryMaster INNER JOIN PurchaseOrderLines ON InventoryMaster.ItemNumber=PurchaseOrderLines.ItemNumber AND PurchaseOrderLines.HeaderID=" & Me.POID & " INNER JOIN vActualWhseQty ON vActualWhseQty.ItemNumber=PurchaseOrderLines.ItemNumber WHERE InventoryMaster.ItemNumber='" & ItemList(offset + i) & "'")
                Me.ItemNumber(i).Visible = True
                Me.ItemNumber(i) = rst("ItemNumber")
                Me.QtyOrdered(i).Visible = True
                'Me.QtyOrdered(i) = rst("QtyOrdered")
                
                Me.QtyBackOrdered(i).Visible = True
                Me.QtyBackOrdered(i) = rst("QtyOnBO")
                
                Me.QtyOnHand(i).Visible = True
                Me.QtyOnHand(i) = rst("QtyOnHand")
                
                Me.QtyOnSO(i).Visible = True
                Me.QtyOnSO(i) = rst("QtyOnSO")
                
                Me.QtyOrdered(i) = rst("Quantity")
                Me.StdCost(i).Visible = True
                Me.StdCost(i) = Format(rst("StdCost"), "Currency")
                Me.Subtotal(i).Visible = True
                Me.Subtotal(i) = Format(Me.StdCost(i) * Me.QtyOrdered(i), "Currency")
                Me.lblComment(i).Visible = True
                Me.POComment(i).Visible = True
                Me.POComment(i) = Nz(rst("POComment"))
                rst.Close
            Else
                While i < 5
                    Me.ItemNumber(i).Visible = False
                    Me.QtyOrdered(i).Visible = False
                    Me.QtyBackOrdered(i).Visible = False
                    Me.QtyOnHand(i).Visible = False
                    Me.QtyOnSO(i).Visible = False
                    Me.StdCost(i).Visible = False
                    Me.Subtotal(i).Visible = False
                    Me.lblComment(i).Visible = False
                    Me.POComment(i).Visible = False
                    i = i + 1
                Wend
            End If
        Else
            While i < 5
                Me.ItemNumber(i).Visible = False
                Me.QtyOrdered(i).Visible = False
                Me.QtyBackOrdered(i).Visible = False
                Me.QtyOnHand(i).Visible = False
                Me.QtyOnSO(i).Visible = False
                Me.StdCost(i).Visible = False
                Me.Subtotal(i).Visible = False
                Me.lblComment(i).Visible = False
                Me.POComment(i).Visible = False
                i = i + 1
            Wend
        End If
    Next i
End Sub

Private Sub shuffleUp()
    shuffleItems 1, 0
    shuffleItems 2, 1
    shuffleItems 3, 2
    shuffleItems 4, 3
End Sub

Private Sub shuffleDown()
    shuffleItems 3, 4
    shuffleItems 2, 3
    shuffleItems 1, 2
    shuffleItems 0, 1
End Sub

Private Sub shuffleItems(fromIndex As Long, toIndex As Long)
    Me.ItemNumber(toIndex) = Me.ItemNumber(fromIndex)
    Me.QtyOrdered(toIndex) = Me.QtyOrdered(fromIndex)
    Me.QtyBackOrdered(toIndex) = Me.QtyBackOrdered(fromIndex)
    Me.QtyOnHand(toIndex) = Me.QtyOnHand(fromIndex)
    Me.QtyOnSO(toIndex) = Me.QtyOnSO(fromIndex)
    Me.StdCost(toIndex) = Me.StdCost(fromIndex)
    Me.Subtotal(toIndex) = Me.Subtotal(fromIndex)
    Me.POComment(toIndex) = Me.POComment(fromIndex)
End Sub

Private Function addItemAfter(item As String) As Boolean
On Error GoTo errh
    Dim i As Long
    i = bsearch(item, ItemList) + 1
    Dim rst As ADODB.Recordset
    'Set rst = DB.retrieve("SELECT ItemNumber, QtyOrdered, StdCost, POComment FROM InventoryMaster WHERE ItemNumber='" & ItemList(i) & "'")
    'Set rst = DB.retrieve("SELECT InventoryMaster.ItemNumber, Quantity, StdCost, POComment FROM InventoryMaster INNER JOIN PurchaseOrderLines ON InventoryMaster.ItemNumber=PurchaseOrderLines.ItemNumber WHERE InventoryMaster.ItemNumber='" & ItemList(i) & "' AND HeaderID=" & Me.POID)
    Set rst = DB.retrieve("SELECT InventoryMaster.ItemNumber, InventoryMaster.StdCost, InventoryMaster.POComment, PurchaseOrderLines.Quantity, vActualWhseQty.QtyOnSO, vActualWhseQty.QtyOnBO, vActualWhseQty.QtyOnHand FROM InventoryMaster INNER JOIN PurchaseOrderLines ON InventoryMaster.ItemNumber=PurchaseOrderLines.ItemNumber AND PurchaseOrderLines.HeaderID=" & Me.POID & " INNER JOIN vActualWhseQty ON vActualWhseQty.ItemNumber=PurchaseOrderLines.ItemNumber WHERE InventoryMaster.ItemNumber='" & ItemList(i) & "'")
    Me.ItemNumber(4) = rst("ItemNumber")
    'Me.QtyOrdered(4) = rst("QtyOrdered")
    Me.QtyOrdered(4) = rst("Quantity")
    Me.QtyBackOrdered(4) = rst("QtyOnBO")
    Me.QtyOnHand(4) = rst("QtyOnHand")
    Me.QtyOnSO(4) = rst("QtyOnSO")
    Me.StdCost(4) = Format(rst("StdCost"), "Currency")
    Me.Subtotal(4) = Format(Me.StdCost(4) * Me.QtyOrdered(4), "Currency")
    Me.POComment(4) = Nz(rst("POComment"))
    rst.Close
    Set rst = Nothing
    addItemAfter = True
    Exit Function
    
errh:
    Err.Clear
    addItemAfter = False
End Function

Private Function addItemBefore(item As String) As Boolean
On Error GoTo errh
    Dim i As Long
    i = bsearch(item, ItemList) - 1
    Dim rst As ADODB.Recordset
    'Set rst = DB.retrieve("SELECT ItemNumber, QtyOrdered, StdCost, POComment FROM InventoryMaster WHERE ItemNumber='" & ItemList(i) & "'")
    'Set rst = DB.retrieve("SELECT InventoryMaster.ItemNumber, Quantity, StdCost, POComment FROM InventoryMaster INNER JOIN PurchaseOrderLines ON InventoryMaster.ItemNumber=PurchaseOrderLines.ItemNumber WHERE InventoryMaster.ItemNumber='" & ItemList(i) & "' AND HeaderID=" & Me.POID)
    Set rst = DB.retrieve("SELECT InventoryMaster.ItemNumber, InventoryMaster.StdCost, InventoryMaster.POComment, PurchaseOrderLines.Quantity, vActualWhseQty.QtyOnSO, vActualWhseQty.QtyOnBO, vActualWhseQty.QtyOnHand FROM InventoryMaster INNER JOIN PurchaseOrderLines ON InventoryMaster.ItemNumber=PurchaseOrderLines.ItemNumber AND PurchaseOrderLines.HeaderID=" & Me.POID & " INNER JOIN vActualWhseQty ON vActualWhseQty.ItemNumber=PurchaseOrderLines.ItemNumber WHERE InventoryMaster.ItemNumber='" & ItemList(i) & "'")
    
    Me.ItemNumber(0) = rst("ItemNumber")
    'Me.QtyOrdered(0) = rst("QtyOrdered")
    Me.QtyOrdered(0) = rst("Quantity")
    Me.QtyBackOrdered(0) = rst("QtyOnBO")
    Me.QtyOnHand(0) = rst("QtyOnHand")
    Me.QtyOnSO(0) = rst("QtyOnSO")
    Me.StdCost(0) = Format(rst("StdCost"), "Currency")
    Me.Subtotal(0) = Format(Me.StdCost(0) * Me.QtyOrdered(0), "Currency")
    Me.POComment(0) = Nz(rst("POComment"))
    rst.Close
    Set rst = Nothing
    addItemBefore = True
    Exit Function
    
errh:
    Err.Clear
    addItemBefore = False
End Function

Private Sub handleFinalizedText(tf As Boolean)
    If tf Then
        Me.lblFinalized.Caption = "Finalized"
    Else
        Me.lblFinalized.Caption = "Not Finalized"
    End If
End Sub

Public Sub GoToPurchaseOrder(POID As String)
    Me.barPosition.value = lsearch(POID, POList)
End Sub





'-------------------
' db updating
'-------------------
Private Sub chkDeferred_Click()
    If Not fillingForm Then
        updatePurchaseOrders "Deferrd", SQLBool(Me.chkDeferred), Me.POID, ""
        If Me.chkDeferred Then
            Me.DeferredList = "30"
            DeferredList_LostFocus
            Enable Me.DeferredList
        Else
            Me.DeferredList = ""
            updatePurchaseOrders "DeferredTime", "Null", Me.POID, ""
            Disable Me.DeferredList
        End If
    End If
End Sub

Private Sub chkFreightFree_Click()
    If Not fillingForm Then
        updatePurchaseOrders "FreightFree", "1", Me.POID, ""
    End If
End Sub

Private Sub chkFreightPay_Click()
    If Not fillingForm Then
        updatePurchaseOrders "FreightFree", "0", Me.POID, ""
    End If
End Sub

Private Sub Comment_KeyPress(KeyAscii As Integer) 'overall po comment
    If KeyAscii = vbKeyReturn Then
        Comment_LostFocus
    Else
        changed = True
        whichCtl = "Comment"
    End If
End Sub

Private Sub Comment_LostFocus()
    If changed = True Then
        If Not validatePOComment(Me.Comment) Then
            MsgBox "Comment must be less than 255 chars."
            Me.Comment.SetFocus
        Else
            updatePurchaseOrders "POComment", Me.Comment, Me.POID, "'"
            changed = False
        End If
    End If
End Sub

Private Sub DeferredList_Click()
    DeferredList_LostFocus
End Sub

Private Sub DeferredList_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.DeferredList, KeyCode, Shift
End Sub

Private Sub DeferredList_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.DeferredList, KeyAscii
    If KeyAscii = vbKeyReturn Then
        DeferredList_LostFocus
    End If
End Sub

Private Sub DeferredList_LostFocus()
    AutoCompleteLostFocus Me.DeferredList
    If Me.DeferredList = "" Then
        Me.chkDeferred = 0
    Else
        updatePurchaseOrders "DeferredTime", Me.DeferredList, Me.POID, ""
    End If
End Sub

Private Sub MiscNotes_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            MiscNotes_KeyPress KeyCode
        Case Is = vbKeyReturn
            MiscNotes_LostFocus
    End Select
End Sub

Private Sub MiscNotes_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "MiscNotes"
End Sub

Private Sub MiscNotes_LostFocus()
    If changed = True Then
        updatePurchaseOrders "MiscNotes", Me.MiscNotes, Me.POID, "'"
        changed = False
    End If
End Sub

Private Sub NumPayments_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        NumPayments_LostFocus
    Else
        changed = True
        whichCtl = "NumPayments"
    End If
End Sub

Private Sub NumPayments_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.NumPayments) Then
            updatePurchaseOrders "NumberOfPayments", Me.NumPayments, Me.POID, ""
            changed = False
        Else
            MsgBox "Number of payments must be integer."
            Me.NumPayments.SetFocus
        End If
    End If
End Sub

Private Sub POComment_KeyPress(index As Integer, KeyAscii As Integer) 'indiv. item comment
    If KeyAscii = vbKeyReturn Then
        POComment_LostFocus index
    Else
        changed = True
        whichCtl = "POComment(" & index & ")"
    End If
End Sub

Private Sub POComment_LostFocus(index As Integer)
    If changed = True Then
        If Not validateItemPOComment(Me.POComment(index)) Then
            MsgBox "PO Comment must be less than 50 chars."
            Me.POComment(index).SetFocus
        Else
            updateInventoryMaster "POComment", Me.POComment(index), Me.ItemNumber(index), "'"
            changed = False
        End If
    End If
End Sub

Private Sub PODueDateCal_Click(selectedDate As String)
    Me.PODueDate = selectedDate
    updatePurchaseOrders "DueDate", Format(Me.PODueDate, "Short Date"), Me.POID, "'"
End Sub

Private Sub POTerms_Click()
    POTerms_LostFocus
End Sub

Private Sub POTerms_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.POTerms, KeyCode, Shift
End Sub

Private Sub POTerms_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.POTerms, KeyAscii
    If KeyAscii = vbKeyReturn Then
        POTerms_LostFocus
    End If
End Sub

Private Sub POTerms_LostFocus()
    AutoCompleteLostFocus Me.POTerms
    Dim termscode As String
    termscode = lookupPOTermsByDesc(Me.POTerms)
    updatePurchaseOrders "POTerms", termscode, Me.POID, "'"
    If termscode = "00" Then
        Enable Me.NumPayments
    Else
        Disable Me.NumPayments
    End If
End Sub

Private Sub QtyOrdered_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        QtyOrdered_LostFocus index
    Else
        changed = True
        whichCtl = "OtyOrdered(" & index & ")"
    End If
End Sub

Private Sub QtyOrdered_LostFocus(index As Integer)
    If changed = True Then
        If Not IsNumeric(Me.QtyOrdered(index)) Then
            MsgBox "Qty ordered must be numeric."
            Me.QtyOrdered(index).SetFocus
        ElseIf Me.QtyOrdered(index) = "0" Or Me.QtyOrdered(index) = "" Then
            If MsgBox("Do you want to remove this item from the Purchase Order?", vbYesNo) = vbYes Then
                'updateInventoryMaster "POID", "Null", Me.ItemNumber(Index), ""
                'updateInventoryMaster "QtyOrdered", "0", Me.ItemNumber(Index), ""
                PurchaseOrderLineEdit CLng(Me.POID), Me.ItemNumber(index), 0
                changed = False
                requeryItems Me.POID
                fillItemSubform
            Else
                Me.QtyOrdered(index).SetFocus
            End If
            fillOrderTotal CInt(Me.POID)
        Else
            'updateInventoryMaster "QtyOrdered", Me.QtyOrdered(Index), Me.ItemNumber(Index), ""
            PurchaseOrderLineEdit CLng(Me.POID), Me.ItemNumber(index), CLng(Me.QtyOrdered(index))
            changed = False
            Me.Subtotal(index) = Format(Me.QtyOrdered(index) * Me.StdCost(index), "Currency")
            fillOrderTotal CInt(Me.POID)
        End If
    End If
End Sub

Private Sub shipAddr1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        shipAddr1_LostFocus
    Else
        changed = True
        whichCtl = "ShipAddr1"
    End If
End Sub

Private Sub shipAddr1_LostFocus()
    If changed = True Then
        updatePurchaseOrders "ShipToAddress1", Me.shipAddr1, Me.POID, "'"
        changed = False
    End If
End Sub

Private Sub shipAddr2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        shipAddr2_LostFocus
    Else
        changed = True
        whichCtl = "ShipAddr2"
    End If
End Sub

Private Sub shipAddr2_LostFocus()
    If changed = True Then
        updatePurchaseOrders "ShipToAddress2", Me.shipAddr2, Me.POID, "'"
        changed = False
    End If
End Sub

Private Sub shipCity_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        shipCity_LostFocus
    Else
        changed = True
        whichCtl = "ShipCity"
    End If
End Sub

Private Sub shipCity_LostFocus()
    If changed = True Then
        updatePurchaseOrders "ShipToCity", Me.shipCity, Me.POID, "'"
        changed = False
    End If
End Sub

Private Sub shipName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        shipName_LostFocus
    Else
        changed = True
        whichCtl = "ShipName"
    End If
End Sub

Private Sub shipName_LostFocus()
    If changed = True Then
        updatePurchaseOrders "ShipToName", Me.shipName, Me.POID, "'"
        changed = False
    End If
End Sub

Private Sub shipState_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        shipState_LostFocus
    Else
        changed = True
        whichCtl = "ShipState"
    End If
End Sub

Private Sub shipState_LostFocus()
    If changed = True Then
        updatePurchaseOrders "ShipToState", Me.shipState, Me.POID, "'"
        changed = False
    End If
End Sub

Private Sub shipZipCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        shipZipCode_LostFocus
    Else
        changed = True
        whichCtl = "ShipZipCode"
    End If
End Sub

Private Sub shipZipCode_LostFocus()
    If changed = True Then
        updatePurchaseOrders "ShipToZip", Me.shipZipCode, Me.POID, "'"
        changed = False
    End If
End Sub

Private Sub DropshipSalesOrderNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            DropshipSalesOrderNo_LostFocus
        Case Is = vbKeyDelete
            DropshipSalesOrderNo_KeyPress KeyCode
    End Select
End Sub

Private Sub DropshipSalesOrderNo_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "DropshipSalesOrderNo"
End Sub

Private Sub DropshipSalesOrderNo_LostFocus()
    If changed = True Then
        updatePurchaseOrders "SalesOrderNo", Me.DropshipSalesOrderNo, Me.POID, "'"
        changed = False
    End If
End Sub

Private Sub DropshipTelephoneNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            DropshipTelephoneNo_LostFocus
        Case Is = vbKeyDelete
            DropshipTelephoneNo_KeyPress KeyCode
    End Select
End Sub

Private Sub DropshipTelephoneNo_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "DropshipTelephoneNo"
End Sub

Private Sub DropshipTelephoneNo_LostFocus()
    If changed = True Then
        updatePurchaseOrders "ShipToTelephoneNo", Me.DropshipTelephoneNo, Me.POID, "'"
        changed = False
    End If
End Sub

Private Sub DropshipCustomerEmailAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            DropshipCustomerEmailAddress_LostFocus
        Case Is = vbKeyDelete
            DropshipCustomerEmailAddress_KeyPress KeyCode
    End Select
End Sub

Private Sub DropshipCustomerEmailAddress_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "DropshipCustomerEmailAddress"
End Sub

Private Sub DropshipCustomerEmailAddress_LostFocus()
    If changed = True Then
        If Len(Me.DropshipCustomerEmailAddress) > 35 Then
            Me.DropshipCustomerEmailAddress = Left(Me.DropshipCustomerEmailAddress, 35)
        End If
        updatePurchaseOrders "CustomerEmailAddress", Me.DropshipCustomerEmailAddress, Me.POID, "'"
        changed = False
    End If
End Sub

Private Sub chkDropshipLiftGate_Click()
    If Not fillingForm Then
        updatePurchaseOrders "LiftGateRequired", Me.chkDropshipLiftGate, Me.POID, ""
    End If
End Sub

Private Sub chkDropshipResidential_Click()
    If Not fillingForm Then
        updatePurchaseOrders "AddressZoningType", IIf(Me.chkDropshipResidential, "1", "0"), Me.POID, ""
    End If
End Sub

Private Sub StdCost_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        StdCost_LostFocus index
    Else
        changed = True
        whichCtl = "StdCost(" & index & ")"
    End If
End Sub

Private Sub StdCost_LostFocus(index As Integer)
    If changed = True Then
        If Not validateCurrency(Me.StdCost(index)) Then
            MsgBox "Cost must be numeric."
            Me.StdCost(index).SetFocus
        Else
            Dim item As String, cost As String, itemID As String
            item = Me.ItemNumber(index)
            cost = Me.StdCost(index)
            itemID = Utilities.GetItemIDByItemCode(item)
            
            LogItemCostChanged item, cost
            updateInventoryMaster "StdCost", Format(cost, "Currency"), item, ""
            DatabaseFunctions.ModifyItemCost CLng(itemID), "Std Cost", cost
            
            If Left(item, 3) <> "XXX" And Left(item, 3) <> "ZZZ" Then
                Dim zzz As String
                zzz = ZZZifyItemNumber(item)
                If zzz <> "" Then
                    updateInventoryMaster "StdCost", cost, zzz, ""
                    Dim zzzID As String
                    zzzID = Utilities.GetItemIDByItemCode(zzz)
                    If zzzID <> -1 Then
                        DatabaseFunctions.ModifyItemCost CLng(zzzID), "Std Cost", cost
                    End If
                End If
            End If
            changed = False
            Me.Subtotal(index) = Format(Me.QtyOrdered(index) * Me.StdCost(index), "Currency")
            fillOrderTotal CInt(Me.POID)
        End If
    End If
End Sub

Private Sub VendorNumber_Click()
    VendorNumber_LostFocus
End Sub

Private Sub VendorNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.VendorNumber, KeyCode, Shift
End Sub

Private Sub VendorNumber_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.VendorNumber, KeyAscii
    If KeyAscii = vbKeyReturn Then
        VendorNumber_LostFocus
    End If
End Sub

Private Sub VendorNumber_LostFocus()
    AutoCompleteLostFocus Me.VendorNumber
    updatePurchaseOrders "VendorNumber", Me.VendorNumber, Me.POID, "'"
End Sub

Private Sub VendorPOAddr_Click()
    VendorPOAddr_LostFocus
End Sub

Private Sub VendorPOAddr_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.VendorPOAddr, KeyCode, Shift
End Sub

Private Sub VendorPOAddr_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.VendorPOAddr, KeyAscii
    If KeyAscii = vbKeyReturn Then
        VendorPOAddr_LostFocus
    End If
End Sub

Private Sub VendorPOAddr_LostFocus()
    AutoCompleteLostFocus Me.VendorPOAddr
    updatePurchaseOrders "VendorPOAddr", Me.VendorPOAddr, Me.POID, "'"
End Sub
