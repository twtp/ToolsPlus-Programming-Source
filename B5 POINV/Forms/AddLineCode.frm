VERSION 5.00
Begin VB.Form AddLineCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Edit Line Code"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11640
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame dropshipInfoFrame 
      Height          =   2565
      Left            =   4470
      TabIndex        =   108
      Top             =   3840
      Width           =   7155
      Begin VB.CheckBox chkDropshipAdderPercent 
         Caption         =   "%"
         Height          =   225
         Index           =   1
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   1980
         Value           =   1  'Checked
         Width           =   225
      End
      Begin VB.CheckBox chkDropshipAdderPercent 
         Caption         =   "%"
         Height          =   225
         Index           =   0
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   840
         Value           =   1  'Checked
         Width           =   225
      End
      Begin VB.CheckBox chkDropshipAdderDollars 
         Caption         =   "$"
         Height          =   225
         Index           =   1
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   1980
         Width           =   225
      End
      Begin VB.CheckBox chkDropshipAdderDollars 
         Caption         =   "$"
         Height          =   225
         Index           =   0
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   840
         Width           =   225
      End
      Begin VB.TextBox DropshipAdderMaximum 
         Height          =   255
         Index           =   1
         Left            =   2610
         TabIndex        =   160
         Top             =   2220
         Width           =   855
      End
      Begin VB.TextBox DropshipAdderMaximum 
         Height          =   255
         Index           =   0
         Left            =   2610
         TabIndex        =   157
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox DropshipAdderThreshold 
         Height          =   255
         Index           =   1
         Left            =   2610
         TabIndex        =   153
         Top             =   1950
         Width           =   855
      End
      Begin VB.TextBox DropshipAdder 
         Height          =   255
         Index           =   1
         Left            =   1140
         TabIndex        =   151
         Top             =   1950
         Width           =   555
      End
      Begin VB.TextBox LiftGateWeightThreshold 
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   147
         Top             =   1650
         Width           =   585
      End
      Begin VB.TextBox LiftGateCharge 
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   145
         Top             =   1650
         Width           =   855
      End
      Begin VB.TextBox DropshipAdderThreshold 
         Height          =   255
         Index           =   0
         Left            =   2610
         TabIndex        =   143
         Top             =   810
         Width           =   855
      End
      Begin VB.TextBox DropshipAdder 
         Height          =   255
         Index           =   0
         Left            =   1140
         TabIndex        =   141
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox LiftGateWeightThreshold 
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   137
         Top             =   510
         Width           =   585
      End
      Begin VB.TextBox LiftGateCharge 
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   135
         Top             =   510
         Width           =   855
      End
      Begin VB.Frame dropshipChargesFrame 
         Height          =   2415
         Left            =   3840
         TabIndex        =   109
         Top             =   120
         Width           =   3315
         Begin VB.OptionButton opTierCalcType 
            Caption         =   "$"
            Height          =   225
            Index           =   1
            Left            =   3030
            Style           =   1  'Graphical
            TabIndex        =   171
            Top             =   0
            Width           =   225
         End
         Begin VB.OptionButton opTierCalcType 
            Caption         =   "%"
            Height          =   225
            Index           =   0
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   170
            Top             =   0
            Width           =   225
         End
         Begin VB.TextBox dsCost 
            Height          =   255
            Index           =   6
            Left            =   2220
            TabIndex        =   169
            Top             =   1830
            Width           =   975
         End
         Begin VB.TextBox dscHigh 
            Enabled         =   0   'False
            Height          =   255
            Index           =   6
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   168
            Top             =   1830
            Width           =   885
         End
         Begin VB.TextBox dscLow 
            Height          =   255
            Index           =   6
            Left            =   90
            TabIndex        =   167
            Top             =   1830
            Width           =   885
         End
         Begin VB.TextBox dscLow 
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   129
            Text            =   "$0.00"
            Top             =   210
            Width           =   885
         End
         Begin VB.TextBox dscHigh 
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   128
            Top             =   210
            Width           =   885
         End
         Begin VB.TextBox dsCost 
            Height          =   255
            Index           =   0
            Left            =   2220
            TabIndex        =   127
            Top             =   210
            Width           =   975
         End
         Begin VB.TextBox dscLow 
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   126
            Top             =   480
            Width           =   885
         End
         Begin VB.TextBox dscLow 
            Height          =   255
            Index           =   2
            Left            =   90
            TabIndex        =   125
            Top             =   750
            Width           =   885
         End
         Begin VB.TextBox dscLow 
            Height          =   255
            Index           =   3
            Left            =   90
            TabIndex        =   124
            Top             =   1020
            Width           =   885
         End
         Begin VB.TextBox dscLow 
            Height          =   255
            Index           =   4
            Left            =   90
            TabIndex        =   123
            Top             =   1290
            Width           =   885
         End
         Begin VB.TextBox dscLow 
            Height          =   255
            Index           =   5
            Left            =   90
            TabIndex        =   122
            Top             =   1560
            Width           =   885
         End
         Begin VB.TextBox dscHigh 
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   121
            Top             =   480
            Width           =   885
         End
         Begin VB.TextBox dscHigh 
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   120
            Top             =   750
            Width           =   885
         End
         Begin VB.TextBox dscHigh 
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   119
            Top             =   1020
            Width           =   885
         End
         Begin VB.TextBox dscHigh 
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   118
            Top             =   1290
            Width           =   885
         End
         Begin VB.TextBox dscHigh 
            Enabled         =   0   'False
            Height          =   255
            Index           =   5
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   117
            Top             =   1560
            Width           =   885
         End
         Begin VB.TextBox dsCost 
            Height          =   255
            Index           =   1
            Left            =   2220
            TabIndex        =   116
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox dsCost 
            Height          =   255
            Index           =   2
            Left            =   2220
            TabIndex        =   115
            Top             =   750
            Width           =   975
         End
         Begin VB.TextBox dsCost 
            Height          =   255
            Index           =   3
            Left            =   2220
            TabIndex        =   114
            Top             =   1020
            Width           =   975
         End
         Begin VB.TextBox dsCost 
            Height          =   255
            Index           =   4
            Left            =   2220
            TabIndex        =   113
            Top             =   1290
            Width           =   975
         End
         Begin VB.TextBox dsCost 
            Height          =   255
            Index           =   5
            Left            =   2220
            TabIndex        =   112
            Top             =   1560
            Width           =   975
         End
         Begin VB.CommandButton btnSaveDSCharges 
            Caption         =   "Save Changes"
            Height          =   225
            Left            =   150
            TabIndex        =   111
            Top             =   2130
            Width           =   1545
         End
         Begin VB.CommandButton btnCancelDSCharges 
            Caption         =   "Cancel"
            Height          =   225
            Left            =   1830
            TabIndex        =   110
            Top             =   2130
            Width           =   1275
         End
         Begin VB.Label generalLabel 
            Caption         =   "From:"
            Height          =   225
            Index           =   39
            Left            =   120
            TabIndex        =   132
            Top             =   30
            Width           =   435
         End
         Begin VB.Label generalLabel 
            Caption         =   "To:"
            Height          =   225
            Index           =   40
            Left            =   1080
            TabIndex        =   131
            Top             =   30
            Width           =   315
         End
         Begin VB.Label generalLabel 
            Caption         =   "Cost:"
            Height          =   225
            Index           =   41
            Left            =   2220
            TabIndex        =   130
            Top             =   0
            Width           =   435
         End
      End
      Begin VB.Label generalLabel 
         Caption         =   "Adder Max:"
         Height          =   225
         Index           =   62
         Left            =   1800
         TabIndex        =   162
         Top             =   2250
         Width           =   855
      End
      Begin VB.Label generalLabel 
         Caption         =   "$"
         Height          =   225
         Index           =   61
         Left            =   3480
         TabIndex        =   161
         Top             =   2250
         Width           =   225
      End
      Begin VB.Label generalLabel 
         Caption         =   "Adder Max:"
         Height          =   225
         Index           =   60
         Left            =   1800
         TabIndex        =   159
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label generalLabel 
         Caption         =   "$"
         Height          =   225
         Index           =   59
         Left            =   3480
         TabIndex        =   158
         Top             =   1110
         Width           =   225
      End
      Begin VB.Label generalLabel 
         Caption         =   "$"
         Height          =   225
         Index           =   58
         Left            =   3480
         TabIndex        =   155
         Top             =   1980
         Width           =   225
      End
      Begin VB.Label generalLabel 
         Caption         =   "$"
         Height          =   225
         Index           =   57
         Left            =   3480
         TabIndex        =   154
         Top             =   840
         Width           =   225
      End
      Begin VB.Label generalLabel 
         Caption         =   "up to:"
         Height          =   225
         Index           =   55
         Left            =   2190
         TabIndex        =   152
         Top             =   1980
         Width           =   465
      End
      Begin VB.Label generalLabel 
         Caption         =   "Cost Adder:"
         Height          =   225
         Index           =   53
         Left            =   270
         TabIndex        =   150
         Top             =   1980
         Width           =   915
      End
      Begin VB.Label generalLabel 
         Caption         =   "Residential:"
         Height          =   225
         Index           =   52
         Left            =   90
         TabIndex        =   149
         Top             =   1410
         Width           =   1335
      End
      Begin VB.Label generalLabel 
         Caption         =   "lbs"
         Height          =   225
         Index           =   51
         Left            =   3480
         TabIndex        =   148
         Top             =   1680
         Width           =   225
      End
      Begin VB.Label generalLabel 
         Caption         =   "Min:"
         Height          =   225
         Index           =   50
         Left            =   2520
         TabIndex        =   146
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label generalLabel 
         Caption         =   "Lift Gate Charge:"
         Height          =   225
         Index           =   49
         Left            =   300
         TabIndex        =   144
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label generalLabel 
         Caption         =   "up to:"
         Height          =   225
         Index           =   48
         Left            =   2190
         TabIndex        =   142
         Top             =   840
         Width           =   465
      End
      Begin VB.Label generalLabel 
         Caption         =   "Cost Adder:"
         Height          =   225
         Index           =   46
         Left            =   270
         TabIndex        =   140
         Top             =   840
         Width           =   915
      End
      Begin VB.Label generalLabel 
         Caption         =   "Commercial:"
         Height          =   225
         Index           =   45
         Left            =   90
         TabIndex        =   139
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label generalLabel 
         Caption         =   "lbs"
         Height          =   225
         Index           =   31
         Left            =   3480
         TabIndex        =   138
         Top             =   540
         Width           =   225
      End
      Begin VB.Label generalLabel 
         Caption         =   "Min:"
         Height          =   225
         Index           =   24
         Left            =   2520
         TabIndex        =   136
         Top             =   540
         Width           =   345
      End
      Begin VB.Label generalLabel 
         Caption         =   "Lift Gate Charge:"
         Height          =   225
         Index           =   17
         Left            =   300
         TabIndex        =   134
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label generalLabel 
         Caption         =   "Dropship Charges:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   38
         Left            =   120
         TabIndex        =   133
         Top             =   0
         Width           =   2025
      End
   End
   Begin VB.TextBox ProductLineID 
      Height          =   285
      Left            =   90
      TabIndex        =   105
      Text            =   "PL_ID"
      Top             =   720
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.CommandButton btnNewVendor 
      Caption         =   "N"
      Height          =   285
      Left            =   4530
      TabIndex        =   94
      Top             =   420
      Width           =   285
   End
   Begin VB.CommandButton btnEmailCompany 
      Caption         =   "Email:"
      Height          =   255
      Left            =   5010
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   3420
      Width           =   825
   End
   Begin VB.CommandButton btnGoToWebsite 
      Caption         =   "WWW:"
      Height          =   255
      Left            =   5010
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   3150
      Width           =   825
   End
   Begin VB.Frame basicInfoFrame 
      Height          =   2685
      Left            =   60
      TabIndex        =   76
      Top             =   1080
      Width           =   4845
      Begin VB.CheckBox chkMappViolationWarning 
         Alignment       =   1  'Right Justify
         Caption         =   "MAPP Violation Warning"
         Height          =   525
         Left            =   3240
         TabIndex        =   156
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox AvailLimitDefault 
         Height          =   285
         Left            =   4050
         TabIndex        =   99
         Top             =   1500
         Width           =   705
      End
      Begin VB.TextBox BOBelow 
         Height          =   285
         Left            =   4050
         TabIndex        =   96
         Top             =   1200
         Width           =   705
      End
      Begin VB.CheckBox chkDefaultDropshippable 
         Caption         =   "New items default d/s-able"
         Height          =   255
         Left            =   60
         TabIndex        =   93
         Top             =   1500
         Width           =   2205
      End
      Begin VB.CheckBox chkCoop 
         Caption         =   "Coop"
         Height          =   255
         Left            =   60
         TabIndex        =   83
         Top             =   1860
         Width           =   735
      End
      Begin VB.ComboBox realVendor 
         Height          =   315
         Left            =   1770
         TabIndex        =   82
         Top             =   1830
         Width           =   795
      End
      Begin VB.TextBox StdLeadTime 
         Height          =   285
         Left            =   1230
         TabIndex        =   81
         Top             =   2250
         Width           =   615
      End
      Begin VB.TextBox ManufFullName 
         Height          =   315
         Left            =   1500
         TabIndex        =   80
         Top             =   150
         Width           =   3255
      End
      Begin VB.TextBox ManufFullNameCleaned 
         Height          =   285
         Left            =   1500
         TabIndex        =   79
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox ManufFullNameWeb 
         Height          =   285
         Left            =   1500
         TabIndex        =   78
         Top             =   780
         Width           =   3255
      End
      Begin VB.CheckBox chkHide 
         Caption         =   "Hide"
         Height          =   255
         Left            =   60
         TabIndex        =   77
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label generalLabel 
         Caption         =   "Default Avail Limit:"
         Height          =   255
         Index           =   44
         Left            =   2730
         TabIndex        =   100
         ToolTipText     =   "New items for this line will default to this availability limit"
         Top             =   1530
         Width           =   1335
      End
      Begin VB.Label generalLabel 
         Caption         =   "Backorder at/below $:"
         Height          =   255
         Index           =   26
         Left            =   2460
         TabIndex        =   95
         ToolTipText     =   "If an item's cost is below this amount, it will show as backordered on the site, even if it's dropshippable"
         Top             =   1230
         Width           =   1605
      End
      Begin VB.Label generalLabel 
         Caption         =   "Real Vendor:"
         Height          =   225
         Index           =   4
         Left            =   810
         TabIndex        =   88
         Top             =   1890
         Width           =   975
      End
      Begin VB.Label generalLabel 
         Caption         =   "Std Lead Time:"
         Height          =   225
         Index           =   5
         Left            =   60
         TabIndex        =   87
         Top             =   2280
         Width           =   1125
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Manuf. Full Name:"
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   86
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Cleaned:"
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   85
         Top             =   510
         Width           =   1335
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Web:"
         Height          =   255
         Index           =   13
         Left            =   90
         TabIndex        =   84
         Top             =   810
         Width           =   1335
      End
   End
   Begin VB.Frame contactInfoCompanyFrame 
      Height          =   3495
      Left            =   4950
      TabIndex        =   42
      Top             =   270
      Width           =   3315
      Begin VB.TextBox CompanyAddress2 
         Height          =   255
         Left            =   930
         TabIndex        =   44
         Top             =   420
         Width           =   2295
      End
      Begin VB.ComboBox ContactInfoTypeSelector 
         Height          =   315
         Left            =   450
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   2040
         Width           =   2595
      End
      Begin VB.TextBox ShipsFromZipCode 
         Height          =   255
         Left            =   930
         TabIndex        =   92
         Top             =   1500
         Width           =   2295
      End
      Begin VB.TextBox CompanyEmail 
         Height          =   255
         Left            =   930
         TabIndex        =   52
         Top             =   3180
         Width           =   2295
      End
      Begin VB.TextBox CompanyURL 
         Height          =   255
         Left            =   930
         TabIndex        =   51
         Top             =   2910
         Width           =   2295
      End
      Begin VB.TextBox CompanyFax 
         Height          =   255
         Left            =   930
         TabIndex        =   50
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox CompanyPhone 
         Height          =   255
         Left            =   930
         TabIndex        =   49
         Top             =   2370
         Width           =   2295
      End
      Begin VB.TextBox CompanyCountry 
         Height          =   255
         Left            =   930
         TabIndex        =   48
         Top             =   1770
         Width           =   2295
      End
      Begin VB.TextBox CompanyZipCode 
         Height          =   255
         Left            =   930
         TabIndex        =   47
         Top             =   1230
         Width           =   2295
      End
      Begin VB.TextBox CompanyState 
         Height          =   255
         Left            =   930
         TabIndex        =   46
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox CompanyCity 
         Height          =   255
         Left            =   930
         TabIndex        =   45
         Top             =   690
         Width           =   2295
      End
      Begin VB.TextBox CompanyAddress1 
         Height          =   255
         Left            =   930
         TabIndex        =   43
         Top             =   150
         Width           =   2295
      End
      Begin VB.Label generalLabel 
         Caption         =   "Ships From:"
         Height          =   225
         Index           =   14
         Left            =   60
         TabIndex        =   91
         Top             =   1530
         Width           =   855
      End
      Begin VB.Label generalLabel 
         Caption         =   "Fax:"
         Height          =   225
         Index           =   16
         Left            =   120
         TabIndex        =   59
         Top             =   2670
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Caption         =   "Phone:"
         Height          =   225
         Index           =   18
         Left            =   120
         TabIndex        =   58
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Caption         =   "Country:"
         Height          =   225
         Index           =   19
         Left            =   120
         TabIndex        =   57
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Caption         =   "Zip Code:"
         Height          =   225
         Index           =   20
         Left            =   120
         TabIndex        =   56
         Top             =   1260
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Caption         =   "State:"
         Height          =   225
         Index           =   21
         Left            =   120
         TabIndex        =   55
         Top             =   990
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Caption         =   "City:"
         Height          =   225
         Index           =   22
         Left            =   120
         TabIndex        =   54
         Top             =   720
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Caption         =   "Address:"
         Height          =   225
         Index           =   23
         Left            =   120
         TabIndex        =   53
         Top             =   180
         Width           =   795
      End
   End
   Begin VB.Frame miscInfoFrame 
      Height          =   1425
      Left            =   60
      TabIndex        =   67
      Top             =   6510
      Width           =   11505
      Begin VB.TextBox Repair 
         Height          =   1125
         Left            =   8640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   74
         Top             =   240
         Width           =   2805
      End
      Begin VB.TextBox Warranty 
         Height          =   1125
         Left            =   5790
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   72
         Top             =   240
         Width           =   2805
      End
      Begin VB.TextBox CompanyInfo 
         Height          =   1125
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   69
         Top             =   240
         Width           =   2805
      End
      Begin VB.TextBox ServiceCenters 
         Height          =   1125
         Left            =   2940
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   68
         Top             =   240
         Width           =   2805
      End
      Begin VB.Label ExtGeneralLabel 
         Caption         =   "Repair:"
         Height          =   195
         Index           =   3
         Left            =   8700
         TabIndex        =   75
         Top             =   0
         Width           =   585
      End
      Begin VB.Label ExtGeneralLabel 
         Caption         =   "Warranty:"
         Height          =   195
         Index           =   2
         Left            =   5880
         TabIndex        =   73
         Top             =   0
         Width           =   765
      End
      Begin VB.Label ExtGeneralLabel 
         Caption         =   "Company Info:"
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   71
         Top             =   30
         Width           =   1035
      End
      Begin VB.Label ExtGeneralLabel 
         Caption         =   "Service Centers:"
         Height          =   195
         Index           =   1
         Left            =   2970
         TabIndex        =   70
         Top             =   30
         Width           =   1215
      End
   End
   Begin VB.Frame poInfoFrame 
      Height          =   1155
      Left            =   60
      TabIndex        =   60
      Top             =   4770
      Width           =   3045
      Begin VB.TextBox MinimumAmount 
         Height          =   285
         Left            =   1860
         TabIndex        =   63
         Top             =   150
         Width           =   1125
      End
      Begin VB.TextBox Prepaid 
         Height          =   285
         Left            =   1860
         TabIndex        =   62
         Top             =   480
         Width           =   1125
      End
      Begin VB.TextBox PrepaidSpecial 
         Height          =   285
         Left            =   1860
         TabIndex        =   61
         Top             =   810
         Width           =   1125
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Minimum Order Amount:"
         Height          =   255
         Index           =   15
         Left            =   60
         TabIndex        =   66
         Top             =   180
         Width           =   1755
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Prepaid @:"
         Height          =   255
         Index           =   42
         Left            =   60
         TabIndex        =   65
         Top             =   510
         Width           =   1755
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Prepaid @ (special):"
         Height          =   255
         Index           =   43
         Left            =   60
         TabIndex        =   64
         Top             =   840
         Width           =   1755
      End
   End
   Begin VB.Frame contactInfoRepFrame 
      Height          =   3495
      Left            =   8310
      TabIndex        =   25
      Top             =   270
      Width           =   3315
      Begin VB.CheckBox chkSalesRepPrimary 
         Caption         =   "Primary Contact for line"
         Height          =   225
         Left            =   720
         TabIndex        =   106
         Top             =   3150
         Width           =   2385
      End
      Begin VB.ComboBox SalesRepName 
         Height          =   315
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   104
         Top             =   180
         Width           =   2295
      End
      Begin VB.TextBox SalesRepPosition 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   930
         Locked          =   -1  'True
         TabIndex        =   103
         Top             =   540
         Width           =   2295
      End
      Begin VB.CommandButton btnEmailRep 
         Caption         =   "Email:"
         Height          =   255
         Left            =   60
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   2790
         Width           =   825
      End
      Begin VB.TextBox SalesRepEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   930
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2790
         Width           =   2295
      End
      Begin VB.TextBox SalesRepFax 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   930
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox SalesRepCell 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   930
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2250
         Width           =   2295
      End
      Begin VB.TextBox SalesRepPhone 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   930
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1980
         Width           =   2295
      End
      Begin VB.TextBox SalesRepZipCode 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   930
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1620
         Width           =   2295
      End
      Begin VB.TextBox SalesRepState 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   930
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1350
         Width           =   2295
      End
      Begin VB.TextBox SalesRepCity 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   930
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox SalesRepAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   930
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   810
         Width           =   2295
      End
      Begin VB.Label generalLabel 
         Caption         =   "Position:"
         Height          =   225
         Index           =   29
         Left            =   120
         TabIndex        =   101
         Top             =   570
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Caption         =   "Fax:"
         Height          =   225
         Index           =   27
         Left            =   120
         TabIndex        =   41
         Top             =   2550
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Caption         =   "Cell:"
         Height          =   225
         Index           =   28
         Left            =   120
         TabIndex        =   40
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Caption         =   "Phone:"
         Height          =   225
         Index           =   30
         Left            =   120
         TabIndex        =   39
         Top             =   2010
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Caption         =   "Zip Code:"
         Height          =   225
         Index           =   32
         Left            =   120
         TabIndex        =   34
         Top             =   1650
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Caption         =   "State:"
         Height          =   225
         Index           =   33
         Left            =   120
         TabIndex        =   33
         Top             =   1380
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Caption         =   "City:"
         Height          =   225
         Index           =   34
         Left            =   120
         TabIndex        =   32
         Top             =   1110
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Caption         =   "Address:"
         Height          =   225
         Index           =   35
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Caption         =   "Name:"
         Height          =   225
         Index           =   36
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.ComboBox ProductLine 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   30
      Width           =   765
   End
   Begin VB.CheckBox chkSkipHidden 
      Caption         =   "skip hidden?"
      Height          =   255
      Left            =   2550
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1245
   End
   Begin VB.CommandButton btnNextLC 
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Next Line Code"
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton btnPrevLC 
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Previous Line Code"
      Top             =   8040
      Width           =   345
   End
   Begin VB.Frame wholesalePriceFrame 
      Caption         =   "Wholesale Price Levels:"
      Height          =   885
      Left            =   60
      TabIndex        =   5
      Top             =   3810
      Width           =   4065
      Begin VB.TextBox WholesalePriceLevel 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   3450
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   510
         Width           =   435
      End
      Begin VB.TextBox WholesalePriceLevel 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2940
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   510
         Width           =   435
      End
      Begin VB.TextBox WholesalePriceLevel 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2430
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   510
         Width           =   435
      End
      Begin VB.TextBox WholesalePriceLevel 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   510
         Width           =   435
      End
      Begin VB.TextBox WholesalePriceLevel 
         Height          =   285
         Index           =   1
         Left            =   1410
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   510
         Width           =   435
      End
      Begin VB.TextBox WholesalePriceLevel 
         Height          =   285
         Index           =   0
         Left            =   900
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   510
         Width           =   435
      End
      Begin VB.Label generalLabel 
         Caption         =   "Percent:"
         Height          =   255
         Index           =   12
         Left            =   150
         TabIndex        =   12
         Top             =   540
         Width           =   645
      End
      Begin VB.Label generalLabel 
         Caption         =   "F"
         Height          =   225
         Index           =   11
         Left            =   3600
         TabIndex        =   11
         Top             =   270
         Width           =   195
      End
      Begin VB.Label generalLabel 
         Caption         =   "E"
         Height          =   225
         Index           =   10
         Left            =   3120
         TabIndex        =   10
         Top             =   270
         Width           =   195
      End
      Begin VB.Label generalLabel 
         Caption         =   "D"
         Height          =   225
         Index           =   9
         Left            =   2580
         TabIndex        =   9
         Top             =   270
         Width           =   195
      End
      Begin VB.Label generalLabel 
         Caption         =   "C"
         Height          =   225
         Index           =   8
         Left            =   2070
         TabIndex        =   8
         Top             =   270
         Width           =   195
      End
      Begin VB.Label generalLabel 
         Caption         =   "B"
         Height          =   225
         Index           =   7
         Left            =   1530
         TabIndex        =   7
         Top             =   270
         Width           =   195
      End
      Begin VB.Label generalLabel 
         Caption         =   "A"
         Height          =   225
         Index           =   6
         Left            =   1050
         TabIndex        =   6
         Top             =   270
         Width           =   195
      End
   End
   Begin VB.ComboBox PrimaryVendor 
      Height          =   315
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   390
      Width           =   3405
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   345
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8010
      Width           =   1185
   End
   Begin VB.Label lblAcctNumber 
      Caption         =   "ACCT # GOES HERE"
      Height          =   375
      Left            =   90
      TabIndex        =   98
      Top             =   6000
      Width           =   3495
   End
   Begin VB.Label lblVendorNumber 
      Height          =   255
      Left            =   1110
      TabIndex        =   97
      Top             =   750
      Width           =   1455
   End
   Begin VB.Label lblManufName 
      Height          =   255
      Left            =   1860
      TabIndex        =   24
      Top             =   60
      Width           =   2865
   End
   Begin VB.Label generalLabel 
      Caption         =   "Sales Rep Contact Info:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   37
      Left            =   8400
      TabIndex        =   23
      Top             =   30
      Width           =   2475
   End
   Begin VB.Label generalLabel 
      Caption         =   "Company Contact Info:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   5010
      TabIndex        =   22
      Top             =   30
      Width           =   2475
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11640
      Y1              =   7980
      Y2              =   7980
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Vendor:"
      Height          =   255
      Index           =   1
      Left            =   390
      TabIndex        =   4
      Top             =   420
      Width           =   645
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Product Line:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   975
   End
End
Attribute VB_Name = "AddLineCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fillingForm As Boolean
Private changed As Boolean

Private salesRepInfoCache As Variant
Private salesRepLastSelection As Long

Public Sub LoadProductLine(prodLine As String)
    Me.ProductLine = prodLine
    fillForm
End Sub

Private Sub UpdateNow()
    fillForm
End Sub

Private Sub btnCancelDSCharges_Click()
    fillDSCostFrame
End Sub

Private Sub btnEmailCompany_Click()
    If Me.CompanyEmail <> "" Then
        OpenEmailTo Me.CompanyEmail
    End If
End Sub

Private Sub btnEmailRep_Click()
    If Me.SalesRepEmail <> "" Then
        OpenEmailTo Me.SalesRepEmail
    End If
End Sub

Private Sub btnExit_Click()
    CreateHashes "ManufHash"
    Unload AddLineCode
End Sub

Private Sub btnGoToWebsite_Click()
    If Me.CompanyURL <> "" Then
        OpenDefaultApp Me.CompanyURL
    End If
End Sub

Private Sub btnNewVendor_Click()
    Load AddVendor
    AddVendor.Show MODAL
    
    Dim vendor As String
    vendor = Me.PrimaryVendor
    
    Dim MAS200IE As Mas200ImportExport
    Set MAS200IE = New Mas200ImportExport
    MAS200IE.RefreshVendors
    Set MAS200IE = Nothing
    CreateHashes "VendorHash"
    requeryVendor
    
    If InCombo(vendor, Me.PrimaryVendor, True) Then
        Me.PrimaryVendor = vendor
    End If
End Sub

Private Sub btnNextLC_Click()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT TOP 1 ProductLine FROM ProductLine WHERE ProductLine>'" & Me.ProductLine & "'" & IIf(Me.chkSkipHidden, " AND Hide=0", "") & " ORDER BY ProductLine")
    If Not rst.EOF Then
        Me.ProductLine = rst("ProductLine")
    End If
    rst.Close
    Set rst = Nothing
    fillForm
End Sub

Private Sub btnPrevLC_Click()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT TOP 1 ProductLine FROM ProductLine WHERE ProductLine<'" & Me.ProductLine & "'" & IIf(Me.chkSkipHidden, " AND Hide=0", "") & " ORDER BY ProductLine DESC")
    If Not rst.EOF Then
        Me.ProductLine = rst("ProductLine")
    End If
    rst.Close
    Set rst = Nothing
    fillForm
End Sub

Private Sub btnSaveDSCharges_Click()
    Dim i As Long
    For i = 1 To Me.dscLow.UBound
        If Me.dscLow(i) = "" Then
            i = Me.dscLow.UBound
        ElseIf CInt(Me.dscLow(i)) < CInt(Me.dscLow(i - 1)) Then
            MsgBox "invalid data, from's must go in ascending order"
            Exit Sub
        ElseIf Me.dsCost(i) = "" Then
            MsgBox "invalid data, you are missing a cost"
            Exit Sub
        End If
    Next i
    DB.Execute "DELETE FROM ProductLineDropshipFlatFees WHERE ProductLineID='" & Me.ProductLineID & "'"
    Dim calcType As Long
    Select Case True
        Case Is = Me.opTierCalcType(0)
            calcType = 0
        Case Is = Me.opTierCalcType(1)
            calcType = 1
        Case Else
            Err.Raise 123, "AddLineCode", "missing AmountCalculationType"
    End Select
    Dim cost As String, amt As String, finalDone As Boolean
    For i = 1 To Me.dscLow.UBound
        If Me.dscLow(i) = "" Then
            'done, put the final one on
            cost = 2 ^ 31
            amt = cleanDSTierCost(Me.dsCost(i - 1))
            DB.Execute "INSERT INTO ProductLineDropshipFlatFees ( ProductLineID, ThresholdCost, AmountCalculationType, Amount ) VALUES ( " & Me.ProductLineID & ", " & cost & ", " & calcType & ", " & amt & " )"
            i = Me.dscLow.UBound
            finalDone = True
        Else
            cost = Replace(Replace(Me.dscLow(i), ",", ""), "$", "")
            amt = cleanDSTierCost(Me.dsCost(i - 1))
            DB.Execute "INSERT INTO ProductLineDropshipFlatFees ( ProductLineID, ThresholdCost, AmountCalculationType, Amount ) VALUES ( " & Me.ProductLineID & ", " & cost & ", " & calcType & ", " & amt & " )"
        End If
    Next i
    If Not finalDone Then
        cost = 2 ^ 31
        amt = cleanDSTierCost(Me.dsCost(i - 1))
        DB.Execute "INSERT INTO ProductLineDropshipFlatFees ( ProductLineID, ThresholdCost, AmountCalculationType, Amount ) VALUES ( " & Me.ProductLineID & ", " & cost & ", " & calcType & ", " & amt & " )"
    End If
    
    Logic.CalculateDropshipChargesClearCacheFor Me.ProductLine
End Sub

Private Sub chkDefaultDropshippable_Click()
    If Not fillingForm Then
        updateProdLine "DefaultDropshippable", Me.chkDefaultDropshippable, Me.ProductLine, ""
        Dim rst As ADODB.Recordset
        If Me.chkDefaultDropshippable Then
            If MsgBox("Update all items to be dropshippable?", vbYesNo) = vbYes Then
                Set rst = DB.retrieve("SELECT InventoryMaster.ItemNumber FROM InventoryMaster INNER JOIN vItemStatusBreakdown ON InventoryMaster.ItemNumber=vItemStatusBreakdown.ItemNumber WHERE ProductLine='" & Me.ProductLine & "' AND DCbyMfr=0 AND DSable=0")
                While Not rst.EOF
                    DB.Execute "UPDATE InventoryMaster SET ItemStatus=" & SetItemDSable(rst("ItemNumber")) & " WHERE ItemNumber='" & rst("ItemNumber") & "'"
                    rst.MoveNext
                Wend
                rst.Close
                Set rst = Nothing
            End If
        Else
            If MsgBox("Remove dropship flag from all items?", vbYesNo) = vbYes Then
                Set rst = DB.retrieve("SELECT ItemNumber FROM InventoryMaster WHERE ProductLine='" & Me.ProductLine & "'")
                While Not rst.EOF
                    updateInventoryMaster "ItemStatus", UnSetItemDSable(rst("ItemNumber")), rst("ItemNumber"), ""
                    rst.MoveNext
                Wend
                rst.Close
                Set rst = Nothing
            End If
        End If
    End If
End Sub

Private Sub chkHide_Click()
    If Not fillingForm Then
        updateProdLine "Hide", Me.chkHide, Me.ProductLine, ""
    End If
End Sub

Private Sub chkMappViolationWarning_Click()
    If Not fillingForm Then
        updateProdLine "MappViolationWarning", Me.chkMappViolationWarning, Me.ProductLine, ""
    End If
End Sub

Private Sub chkSalesRepPrimary_Click()
    If Not fillingForm Then
        fillingForm = True
        Me.chkSalesRepPrimary = IIf(Me.chkSalesRepPrimary, 0, 1)
        fillingForm = False
    End If
End Sub

Private Sub ContactInfoTypeSelector_Click()
    If IsNumeric(Me.ProductLineID) Then
        fillCompanyContactInfo
    End If
End Sub

Private Sub dscLow_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = LimitToCurrency(KeyAscii)
End Sub

Private Sub dscLow_LostFocus(Index As Integer)
    If Me.dscLow(Index) = "" Then
        'Me.dscHigh(Index) = ""
        'Me.dsCost(Index) = ""
        If Index > 0 Then 'always true, actually
            Me.dscHigh(Index - 1) = "and up"
        End If
        Dim i As Long
        For i = Index To Me.dscLow.UBound
            Me.dscHigh(Index) = ""
            Me.dsCost(Index) = ""
        Next i
    ElseIf validateCurrency(Me.dscLow(Index)) Then
        Me.dscLow(Index) = Format(Me.dscLow(Index), "Currency")
        Me.dscHigh(Index - 1) = Format(Me.dscLow(Index) - 0.01, "Currency")
        If Me.dscHigh(Index) = "" Then
            Me.dscHigh(Index) = "and up"
        End If
    Else
        MsgBox "Error, must be a valid number"
    End If
End Sub

Private Sub dsCost_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = LimitToCurrency(KeyAscii)
End Sub

Private Sub dsCost_LostFocus(Index As Integer)
    If Me.dsCost(Index) = "" Then
        'do nothing
    ElseIf Me.opTierCalcType(0) And validateCurrency(cleanDSTierCost(Me.dsCost(Index))) Then
        Me.dsCost(Index) = cleanDSTierCost(Me.dsCost(Index))
        Me.dsCost(Index) = CDec(Me.dsCost(Index)) & "%"
        If Me.dsCost(Index) = "0%" Then
            Me.dsCost(Index) = "$0.00 / ppd"
        End If
    ElseIf Me.opTierCalcType(1) And validateCurrency(cleanDSTierCost(Me.dsCost(Index))) Then
        Me.dsCost(Index) = Format(Me.dsCost(Index), "Currency")
        If Me.dsCost(Index) = "$0.00" Then
            Me.dsCost(Index) = "$0.00 / ppd"
        End If
    Else
        MsgBox "Error, must be a valid number"
    End If
End Sub

Private Sub Form_Load()
    requeryContactInfoTypeSelector
    requeryLines
    requeryVendor
'PL gets pushed from inv maint (or anywhere) now
'    Me.ProductLine = Left(InventoryMaintenance.ItemNumber, 3)
'    fillForm
    If Not HasPermissionsTo("InventoryMaintenanceWrite") Then
        Me.PrimaryVendor.Enabled = False
        Me.basicInfoFrame.Enabled = False
        Me.contactInfoCompanyFrame.Enabled = False 'website/email buttons are overlayed on frame so still usable
        Me.poInfoFrame.Enabled = False
        Me.dropshipInfoFrame.Enabled = False
        Me.dropshipChargesFrame.Enabled = False
        Me.wholesalePriceFrame.Enabled = False
        Me.miscInfoFrame.Enabled = False
    End If
End Sub


Private Sub requeryContactInfoTypeSelector()
    'the InfoTypeID is based on the indexes here, so don't reorder these unless you
    'run an UPDATE ProductLineContactInfo SET InfoTypeID=XX WHERE InfoTypeID=XX first
    Me.ContactInfoTypeSelector.Clear
    Me.ContactInfoTypeSelector.AddItem "General Contact"
    Me.ContactInfoTypeSelector.AddItem "Payables"
    Me.ContactInfoTypeSelector.AddItem "Purchasing"
    Me.ContactInfoTypeSelector.AddItem "Service"
    Me.ContactInfoTypeSelector.AddItem "Parts"
    Me.ContactInfoTypeSelector.AddItem "Rebate"
    
    Me.ContactInfoTypeSelector = "General Contact"
End Sub

Private Sub requeryLines()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM vRequeryLineCodes ORDER BY ProductLine")
    Me.ProductLine.Clear
    While Not rst.EOF
        Me.ProductLine.AddItem rst("ProductLine")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryVendor()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM vRequeryVendor ORDER BY VendorName")
    Me.PrimaryVendor.Clear
    While Not rst.EOF
        Me.PrimaryVendor.AddItem (Trim$(rst("VendorName")))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryCoopVendor(VendorNumber As String)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT RealVendor FROM ProductLineCoopVendors WHERE CoopVendor='" & VendorNumber & "'")
    Me.realVendor.Clear
    While Not rst.EOF
        Me.realVendor.AddItem (rst("RealVendor"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub fillForm()
    fillingForm = True
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM ProductLine WHERE ProductLine='" & Me.ProductLine & "'")
    Me.ProductLineID = rst("ID")
    Me.lblManufName.Caption = Nz(rst("ManufFullNameCleaned"))
    Me.PrimaryVendor = VendorHash.item(Nz(rst("PrimaryVendorNumber").value))
    Me.lblVendorNumber.Caption = Nz(rst("PrimaryVendorNumber"))
    requeryCoopVendor Nz(rst("PrimaryVendorNumber"))
    Me.chkCoop = SQLBool(rst("Coop"))
    If Me.chkCoop Then
        Enable Me.realVendor
        Me.realVendor = Nz(rst("RealVendor"))
    Else
        Me.realVendor = ""
        Disable Me.realVendor
    End If
    Me.chkHide = SQLBool(rst("Hide"))
    Me.chkDefaultDropshippable = SQLBool(rst("DefaultDropShippable"))
    Me.chkMappViolationWarning = SQLBool(rst("MappViolationWarning"))
    Me.BOBelow = IIf(rst("BOBelow") = 0, "", Format(rst("BOBelow"), "Currency"))
    Me.AvailLimitDefault = rst("AvailLimitDefault")
    Me.StdLeadTime = rst("StdLeadTime")
    Me.ManufFullName = Nz(rst("ManufFullName"))
    Me.ManufFullNameCleaned = Nz(rst("ManufFullNameCleaned"))
    Me.ManufFullNameWeb = Nz(rst("ManufFullNameWeb"))
    If Me.ManufFullNameWeb = "" Then
        Enable Me.ManufFullNameWeb
    Else
        Disable Me.ManufFullNameWeb
    End If
    fillCompanyInfo rst
    
    fillWholesalePricing
    
    Me.CompanyInfo = Nz(rst("CompanyInfo"))
    Me.ServiceCenters = Nz(rst("ServiceCenters"))
    Me.Warranty = Nz(rst("Warranty"))
    Me.Repair = Nz(rst("Repair"))
    Me.MinimumAmount = IIf(IsNumeric(rst("MinimumOrderAmount")), Format(rst("MinimumOrderAmount"), "Currency"), Nz(rst("MinimumOrderAmount")))
    Me.Prepaid = IIf(IsNumeric(rst("PrepaidAmount")), Format(rst("PrepaidAmount"), "Currency"), Nz(rst("PrepaidAmount")))
    Me.PrepaidSpecial = IIf(IsNumeric(rst("PrepaidSpecialAmount")), Format(rst("PrepaidSpecialAmount"), "Currency"), Nz(rst("PrepaidSpecialAmount")))
    
    If Nz(rst("PrimaryVendorNumber")) <> "" Then
        Dim acct As String
        acct = DLookup("Reference", "AP_Vendor", "VendorNo='" & rst("PrimaryVendorNumber") & "'")
        If acct <> "" Then
            Me.lblAcctNumber.Caption = "Reference: " & acct
        Else
            Me.lblAcctNumber.Caption = ""
        End If
    Else
        Me.lblAcctNumber.Caption = ""
    End If
    loadSalesRepFrameInfo
    If Me.SalesRepName.ListCount > 1 Then
        Dim foundPrimary As Boolean
        foundPrimary = False
        Dim i As Long
        For i = 0 To UBound(salesRepInfoCache)
            If salesRepInfoCache(i).item("IsPrimary") Then
                foundPrimary = True
                Me.SalesRepName.ListIndex = i + 1
                Exit For
            End If
        Next i
        If Not foundPrimary Then
            Me.SalesRepName.ListIndex = 1
        End If
        salesRepLastSelection = Me.SalesRepName.ListIndex
        fillSalesRepFrame
    Else
        fillSalesRepFrame 'this will clear it out
        salesRepLastSelection = Me.SalesRepName.ListIndex
    End If
    fillDSFrame
    
    fillingForm = False
End Sub

Private Sub loadSalesRepFrameInfo()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT DISTINCT ProductLineSalesReps.IsPrimary, SalesReps.ID, SalesReps.FirstName, SalesReps.LastName, SalesReps.PositionTitle, SalesReps.Address, SalesReps.City, SalesReps.State, SalesReps.ZipCode, SalesReps.Phone, SalesReps.Cell, SalesReps.Fax, SalesReps.Email FROM SalesReps LEFT OUTER JOIN ProductLineSalesReps ON SalesReps.ID=ProductLineSalesReps.SalesRepID WHERE ProductLineSalesReps.ProductLineID=" & Me.ProductLineID)
    Me.SalesRepName.Clear
    Me.SalesRepName.AddItem "<Add/Edit Sales Reps...>"
    While Not rst.EOF
        Me.SalesRepName.AddItem IIf(Nz(rst("FirstName")) = "", "", Nz(rst("FirstName")) & " ") & Nz(rst("LastName")) & IIf(Nz(rst("PositionTitle")) = "", "", " (" & Nz(rst("PositionTitle")) & ")")
        rst.MoveNext
    Wend
    salesRepInfoCache = RSToArrayOfHashes(rst)
    rst.Close
    Set rst = Nothing
    
    ExpandDropDownToFit Me.SalesRepName
    
    'Dim primaryId As String
    'primaryId = DLookup("PrimarySalesRepID", "ProductLine", "ID=" & Me.ProductLineID)
    'Dim i As Long
    'For i = 0 To UBound(salesRepInfoCache)
    '    salesRepInfoCache(i).Add "PrimaryContact", IIf(salesRepInfoCache(i).item("ID") = primaryId, True, False)
    'Next i
End Sub

Private Sub opTierCalcType_Click(Index As Integer)
    If Not fillingForm Then
        Dim i As Long
        For i = Me.dsCost.LBound To Me.dsCost.UBound
            If Me.dsCost(i) <> "" Then
                Dim temp As Variant
                temp = CDec(cleanDSTierCost(Me.dsCost(i)))
                If temp = 0 Then
                    Me.dsCost(i) = "$0.00 / ppd"
                Else
                    Select Case True
                        Case Is = Me.opTierCalcType(0)
                            Me.dsCost(i) = temp & "%"
                        Case Is = Me.opTierCalcType(1)
                            Me.dsCost(i) = Format(temp, "Currency")
                        Case Else
                            Err.Raise "123", "AddLineCode", "missing AmountCalculationType"
                    End Select
                End If
            End If
        Next i
    End If
End Sub

Private Sub SalesRepName_Click()
    If Not fillingForm Then
        If Me.SalesRepName = "<Add/Edit Sales Reps...>" Then
            Load SalesRepMaintenance
            SalesRepMaintenance.SetLineFilter Me.ProductLine
            SalesRepMaintenance.Show
            fillingForm = True
            Me.SalesRepName.ListIndex = salesRepLastSelection
            fillingForm = False
        Else
            fillSalesRepFrame
            salesRepLastSelection = Me.SalesRepName.ListIndex
        End If
    End If
End Sub

Private Sub fillCompanyInfo(rst As ADODB.Recordset)
    Me.CompanyAddress1 = Nz(rst("CompanyAddress1"))
    Me.CompanyAddress2 = Nz(rst("CompanyAddress2"))
    Me.CompanyCity = Nz(rst("CompanyCity"))
    Me.CompanyState = Nz(rst("CompanyState"))
    Me.CompanyZipCode = Nz(rst("CompanyZip"))
    Me.ShipsFromZipCode = Nz(rst("ShipsFromZip"))
    Me.CompanyCountry = Nz(rst("CompanyCountry"))
    fillCompanyContactInfo
End Sub

Private Sub fillCompanyContactInfo()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Phone, Fax, URL, Email FROM ProductLineContactInfo WHERE ProductLineID=" & Me.ProductLineID & " AND InfoTypeID=" & Me.ContactInfoTypeSelector.ListIndex)
    If rst.EOF Then
        Me.CompanyPhone = ""
        Me.CompanyFax = ""
        Me.CompanyURL = ""
        Me.CompanyEmail = ""
    Else
        Me.CompanyPhone = Nz(rst("Phone"))
        Me.CompanyFax = Nz(rst("Fax"))
        Me.CompanyURL = Nz(rst("URL"))
        Me.CompanyEmail = Nz(rst("Email"))
    End If
End Sub

Private Sub fillWholesalePricing()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT WholesalePriceLevel, Percentage FROM ProductLineWholesalePricing WHERE ProductLineID=" & Me.ProductLineID & " ORDER BY WholesalePriceLevel")
    Dim nextNum As Long
    nextNum = 0
    While Not rst.EOF
        If rst("WholesalePriceLevel") <> nextNum Then
            DB.Execute "INSERT INTO ProductLineWholesalePricing ( ProductLineID, WholesalePriceLevel, Percentage ) VALUES ( " & Me.ProductLineID & ", " & nextNum & ", 20 )"
            Me.WholesalePriceLevel(nextNum) = 20
        Else
            Me.WholesalePriceLevel(nextNum) = rst("Percentage")
            rst.MoveNext
        End If
        nextNum = nextNum + 1
    Wend
    For nextNum = nextNum To 5
        DB.Execute "INSERT INTO ProductLineWholesalePricing ( ProductLineID, WholesalePriceLevel, Percentage ) VALUES ( " & Me.ProductLineID & ", " & nextNum & ", 20 )"
        Me.WholesalePriceLevel(nextNum) = 20
    Next nextNum
End Sub

Private Sub fillSalesRepFrame()
    If Me.SalesRepName.ListIndex = -1 Then
        Me.SalesRepPosition = ""
        Me.SalesRepAddress = ""
        Me.SalesRepCity = ""
        Me.SalesRepState = ""
        Me.SalesRepZipCode = ""
        Me.SalesRepPhone = ""
        Me.SalesRepCell = ""
        Me.SalesRepFax = ""
        Me.SalesRepEmail = ""
    Else
        Dim dx As Long
        dx = Me.SalesRepName.ListIndex - 1
        Me.SalesRepPosition = salesRepInfoCache(dx).item("PositionTitle")
        Me.SalesRepAddress = salesRepInfoCache(dx).item("Address")
        Me.SalesRepCity = salesRepInfoCache(dx).item("City")
        Me.SalesRepState = salesRepInfoCache(dx).item("State")
        Me.SalesRepZipCode = salesRepInfoCache(dx).item("ZipCode")
        Me.SalesRepPhone = salesRepInfoCache(dx).item("Phone")
        Me.SalesRepCell = salesRepInfoCache(dx).item("Cell")
        Me.SalesRepFax = salesRepInfoCache(dx).item("Fax")
        Me.SalesRepEmail = salesRepInfoCache(dx).item("Email")
        Dim oldFF As Boolean
        oldFF = fillingForm
        fillingForm = True
        Me.chkSalesRepPrimary = SQLBool(salesRepInfoCache(dx).item("IsPrimary"))
        fillingForm = oldFF
    End If
End Sub

Private Sub fillDSFrame()
    Dim i As Long
    For i = 0 To 1
        Me.LiftGateCharge(i) = ""
        Me.LiftGateWeightThreshold(i) = ""
        Me.DropshipAdder(i) = ""
        Me.DropshipAdderThreshold(i) = ""
        Me.DropshipAdderMaximum(i) = ""
        Me.chkDropshipAdderPercent(i) = vbChecked
        Me.chkDropshipAdderDollars(i) = vbUnchecked
    Next i
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT AddressType, LiftGateCharge, LiftGateWeightThreshold, AdderAmount, AdderType, AdderThreshold, AdderMaximum FROM ProductLineDropshipCharges WHERE ProductLineID=" & Me.ProductLineID)
    While Not rst.EOF
        Me.LiftGateCharge(CLng(rst("AddressType"))) = Nz(rst("LiftGateCharge"))
        Me.LiftGateWeightThreshold(CLng(rst("AddressType"))) = Nz(rst("LiftGateWeightThreshold"))
        Me.DropshipAdder(CLng(rst("AddressType"))) = Nz(rst("AdderAmount"))
        Me.DropshipAdderThreshold(CLng(rst("AddressType"))) = Nz(rst("AdderThreshold"))
        Me.DropshipAdderMaximum(CLng(rst("AddressType"))) = Nz(rst("AdderMaximum"))
        If rst("AdderType") = 0 Then
            Me.chkDropshipAdderPercent(CLng(rst("AddressType"))) = vbChecked
            Me.chkDropshipAdderDollars(CLng(rst("AddressType"))) = vbUnchecked
        Else
            Me.chkDropshipAdderPercent(CLng(rst("AddressType"))) = vbUnchecked
            Me.chkDropshipAdderDollars(CLng(rst("AddressType"))) = vbChecked
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    fillDSCostFrame
End Sub

Private Sub fillDSCostFrame()
    Dim i As Long
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ThresholdCost, AmountCalculationType, Amount FROM ProductLineDropshipFlatFees WHERE ProductLineID='" & Me.ProductLineID & "' ORDER BY ThresholdCost")
    If rst.EOF Then
        For i = Me.opTierCalcType.LBound To Me.opTierCalcType.UBound
            Me.opTierCalcType(i) = False
        Next i
        Me.dscHigh(0) = "n/a"
        Me.dsCost(0) = ""
        For i = Me.dscHigh.LBound + 1 To Me.dscHigh.UBound
            Me.dscLow(i) = ""
            Me.dscHigh(i) = ""
            Me.dsCost(i) = ""
        Next i
    Else
        Me.opTierCalcType(rst("AmountCalculationType")) = True
        For i = 1 To rst.RecordCount
            If rst("ThresholdCost") = 2 ^ 31 Then
                If i <> Me.dscLow.UBound + 1 Then
                    Me.dscLow(i) = ""
                    Me.dscHigh(i) = ""
                    Me.dsCost(i) = ""
                End If
                Me.dscHigh(i - 1) = "and up"
            Else
                Me.dscLow(i) = Format(rst("ThresholdCost"), "Currency")
                Me.dscHigh(i - 1) = Format(rst("ThresholdCost") - 0.01, "Currency")
            End If
            If rst("Amount") = 0 Then
                Me.dsCost(i - 1) = "$0.00 / ppd"
            Else
                Select Case CLng(rst("AmountCalculationType"))
                    Case 0
                        Me.dsCost(i - 1) = CDec(rst("Amount")) & "%"
                    Case 1
                        Me.dsCost(i - 1) = Format(rst("Amount"), "Currency")
                    Case Else
                        Err.Raise 123, "AddLineCode", "missing AmountCalculationType"
                End Select
            End If
            rst.MoveNext
        Next i
        For i = i To Me.dscHigh.UBound
            Me.dscLow(i) = ""
            Me.dscHigh(i) = ""
            Me.dsCost(i) = ""
        Next i
    End If
End Sub




Private Sub AvailLimitDefault_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            AvailLimitDefault_KeyPress KeyCode
        Case Is = vbKeyReturn
            AvailLimitDefault_LostFocus
    End Select
End Sub

Private Sub AvailLimitDefault_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
    If KeyAscii <> 0 Then
        changed = True
    End If
End Sub

Private Sub AvailLimitDefault_LostFocus()
    If changed = True Then
        If IsIntegeric(Me.AvailLimitDefault) Then
            updateProdLine "AvailLimitDefault", Me.AvailLimitDefault, Me.ProductLine, ""
        Else
            MsgBox "Invalid entry"
            Me.AvailLimitDefault.SetFocus
        End If
    End If
End Sub

Private Sub BOBelow_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            BOBelow_KeyPress KeyCode
        Case Is = vbKeyReturn
            BOBelow_LostFocus
    End Select
End Sub

Private Sub BOBelow_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        BOBelow_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub BOBelow_LostFocus()
    If changed = True Then
        If Me.BOBelow = "" Then
            updateProdLine "BOBelow", "0", Me.ProductLine, ""
            changed = False
        ElseIf validateCurrency(Me.BOBelow) Then
            Me.BOBelow = Format(Me.BOBelow, "Currency")
            updateProdLine "BOBelow", Me.BOBelow, Me.ProductLine, ""
            changed = False
        Else
            MsgBox "Invalid amount"
            Me.BOBelow.SetFocus
        End If
    End If
End Sub

Private Sub chkCoop_Click()
    If Not fillingForm Then
        updateProdLine "Coop", Me.chkCoop, Me.ProductLine, ""
        If Me.chkCoop Then
            Enable Me.realVendor
        Else
            Me.realVendor = ""
            Disable Me.realVendor
        End If
    End If
End Sub

Private Sub CompanyAddress1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            CompanyAddress1_KeyPress KeyCode
        Case Is = vbKeyReturn
            CompanyAddress1_LostFocus
    End Select
End Sub

Private Sub CompanyAddress1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        CompanyAddress1_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub CompanyAddress1_LostFocus()
    If changed = True Then
        If Len(Me.CompanyAddress1) <= 50 Then
            updateProdLine "CompanyAddress1", Me.CompanyAddress1, Me.ProductLine, "'"
            changed = False
        Else
            MsgBox "Invalid address, must be less than 50 chars."
            Me.CompanyAddress1.SetFocus
        End If
    End If
End Sub

Private Sub CompanyAddress2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            CompanyAddress2_KeyPress KeyCode
        Case Is = vbKeyReturn
            CompanyAddress2_LostFocus
    End Select
End Sub

Private Sub CompanyAddress2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        CompanyAddress2_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub CompanyAddress2_LostFocus()
    If changed = True Then
        If Len(Me.CompanyAddress2) <= 50 Then
            updateProdLine "CompanyAddress2", Me.CompanyAddress2, Me.ProductLine, "'"
            changed = False
        Else
            MsgBox "Invalid address, must be less than 50 chars."
            Me.CompanyAddress2.SetFocus
        End If
    End If
End Sub

Private Sub CompanyCity_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            CompanyCity_KeyPress KeyCode
        Case Is = vbKeyReturn
            CompanyCity_LostFocus
    End Select
End Sub

Private Sub CompanyCity_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        CompanyCity_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub CompanyCity_LostFocus()
    If changed = True Then
        If Len(Me.CompanyCity) <= 50 Then
            updateProdLine "CompanyCity", Me.CompanyCity, Me.ProductLine, "'"
            changed = False
        Else
            MsgBox "Invalid city, must be less than 50 chars."
            Me.CompanyCity.SetFocus
        End If
    End If
End Sub

Private Sub CompanyCountry_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            CompanyCountry_KeyPress KeyCode
        Case Is = vbKeyReturn
            CompanyCountry_LostFocus
    End Select
End Sub

Private Sub CompanyCountry_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        CompanyCountry_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub CompanyCountry_LostFocus()
    If changed = True Then
        If Len(Me.CompanyCountry) <= 50 Then
            updateProdLine "CompanyCountry", Me.CompanyCountry, Me.ProductLine, "'"
            changed = False
        Else
            MsgBox "Invalid country, must be less than 50 chars."
            Me.CompanyCountry.SetFocus
        End If
    End If
End Sub

Private Sub CompanyEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            CompanyEmail_KeyPress KeyCode
        Case Is = vbKeyReturn
            CompanyEmail_LostFocus
    End Select
End Sub

Private Sub CompanyEmail_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub CompanyEmail_LostFocus()
    If changed = True Then
        If Len(Me.CompanyEmail) <= 128 Then
            'updateProdLine "CompanyEmail", Me.CompanyEmail, Me.ProductLine, "'"
            If 0 = DLookup("COUNT(*)", "ProductLineContactInfo", "ProductLineID=" & Me.ProductLineID & " AND InfoTypeID=" & Me.ContactInfoTypeSelector.ListIndex) Then
                DB.Execute "INSERT INTO ProductLineContactInfo ( ProductLineID, InfoTypeID ) VALUES ( " & Me.ProductLineID & ", " & Me.ContactInfoTypeSelector.ListIndex & " )"
            End If
            DB.Execute "UPDATE ProductLineContactInfo SET Email='" & EscapeSQuotes(Me.CompanyEmail) & "' WHERE ProductLineID=" & Me.ProductLineID & " AND InfoTypeID=" & Me.ContactInfoTypeSelector.ListIndex
            changed = False
        Else
            MsgBox "Invalid email, must be less than 128 chars."
            Me.CompanyEmail.SetFocus
        End If
    End If
End Sub

Private Sub CompanyFax_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            CompanyFax_KeyPress KeyCode
        Case Is = vbKeyReturn
            CompanyFax_LostFocus
    End Select
End Sub

Private Sub CompanyFax_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub CompanyFax_LostFocus()
    If changed = True Then
        If Len(Me.CompanyFax) <= 128 Then
            'updateProdLine "CompanyFax", Me.CompanyFax, Me.ProductLine, "'"
            If 0 = DLookup("COUNT(*)", "ProductLineContactInfo", "ProductLineID=" & Me.ProductLineID & " AND InfoTypeID=" & Me.ContactInfoTypeSelector.ListIndex) Then
                DB.Execute "INSERT INTO ProductLineContactInfo ( ProductLineID, InfoTypeID ) VALUES ( " & Me.ProductLineID & ", " & Me.ContactInfoTypeSelector.ListIndex & " )"
            End If
            DB.Execute "UPDATE ProductLineContactInfo SET Fax='" & EscapeSQuotes(Me.CompanyFax) & "' WHERE ProductLineID=" & Me.ProductLineID & " AND InfoTypeID=" & Me.ContactInfoTypeSelector.ListIndex
            changed = False
        Else
            MsgBox "invalid fax, must be less than 128 chars."
            Me.CompanyFax.SetFocus
        End If
    End If
End Sub

Private Sub CompanyPhone_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            CompanyPhone_KeyPress KeyCode
        Case Is = vbKeyReturn
            CompanyPhone_LostFocus
    End Select
End Sub

Private Sub CompanyPhone_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub CompanyPhone_LostFocus()
    If changed = True Then
        If Len(Me.CompanyPhone) <= 128 Then
            'updateProdLine "CompanyPhone", Me.CompanyPhone, Me.ProductLine, "'"
            If 0 = DLookup("COUNT(*)", "ProductLineContactInfo", "ProductLineID=" & Me.ProductLineID & " AND InfoTypeID=" & Me.ContactInfoTypeSelector.ListIndex) Then
                DB.Execute "INSERT INTO ProductLineContactInfo ( ProductLineID, InfoTypeID ) VALUES ( " & Me.ProductLineID & ", " & Me.ContactInfoTypeSelector.ListIndex & " )"
            End If
            DB.Execute "UPDATE ProductLineContactInfo SET Phone='" & EscapeSQuotes(Me.CompanyPhone) & "' WHERE ProductLineID=" & Me.ProductLineID & " AND InfoTypeID=" & Me.ContactInfoTypeSelector.ListIndex
            changed = False
        Else
            MsgBox "invalid phone, must be less than 128 chars."
            Me.CompanyPhone.SetFocus
        End If
    End If
End Sub

Private Sub CompanyState_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            CompanyState_KeyPress KeyCode
        Case Is = vbKeyReturn
            CompanyState_LostFocus
    End Select
End Sub

Private Sub CompanyState_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub CompanyState_LostFocus()
    If changed = True Then
        If Len(Me.CompanyState) <= 16 Then
            updateProdLine "CompanyState", Me.CompanyState, Me.ProductLine, "'"
            changed = False
        Else
            MsgBox "invalid state, must be less than 16 chars."
            Me.CompanyState.SetFocus
        End If
    End If
End Sub

Private Sub CompanyURL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            CompanyURL_KeyPress KeyCode
        Case Is = vbKeyReturn
            CompanyURL_LostFocus
    End Select
End Sub

Private Sub CompanyURL_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub CompanyURL_LostFocus()
    If changed = True Then
        If Len(Me.CompanyURL) <= 128 Then
            'updateProdLine "CompanyURL", Me.CompanyURL, Me.ProductLine, "'"
            If 0 = DLookup("COUNT(*)", "ProductLineContactInfo", "ProductLineID=" & Me.ProductLineID & " AND InfoTypeID=" & Me.ContactInfoTypeSelector.ListIndex) Then
                DB.Execute "INSERT INTO ProductLineContactInfo ( ProductLineID, InfoTypeID ) VALUES ( " & Me.ProductLineID & ", " & Me.ContactInfoTypeSelector.ListIndex & " )"
            End If
            DB.Execute "UPDATE ProductLineContactInfo SET URL='" & EscapeSQuotes(Me.CompanyURL) & "' WHERE ProductLineID=" & Me.ProductLineID & " AND InfoTypeID=" & Me.ContactInfoTypeSelector.ListIndex
            changed = False
        Else
            MsgBox "invalid url, must be less than 128 chars."
            Me.CompanyURL.SetFocus
        End If
    End If
End Sub

Private Sub CompanyZipCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            CompanyZipCode_KeyPress KeyCode
        Case Is = vbKeyReturn
            CompanyZipCode_LostFocus
    End Select
End Sub

Private Sub CompanyZipCode_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub CompanyZipCode_LostFocus()
    If changed = True Then
        If Me.CompanyZipCode = "" Or validateZipCode(Me.CompanyZipCode) Then
            updateProdLine "CompanyZip", Me.CompanyZipCode, Me.ProductLine, "'"
            changed = False
            If Me.CompanyZipCode <> "" Then
                Dim citystate As String
                citystate = DLookup("City", "LookupCityZipCode", "ZipCode='" & Me.CompanyZipCode & "'")
                If citystate <> "" Then
                    Me.CompanyCity = citystate
                    changed = True
                    CompanyCity_LostFocus
                End If
                citystate = DLookup("State", "LookupCityZipCode", "ZipCode='" & Me.CompanyZipCode & "'")
                If citystate <> "" Then
                    Me.CompanyState = citystate
                    changed = True
                    CompanyState_LostFocus
                End If
                If Me.ShipsFromZipCode <> "" Then
                    Me.ShipsFromZipCode = Me.CompanyZipCode
                    changed = True
                    ShipsFromZipCode_LostFocus
                End If
            End If
        Else
            MsgBox "invalid zip code, talk to brian if this is wrong."
            Me.CompanyZipCode.SetFocus
        End If
    End If
End Sub

Private Sub ManufFullName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            ManufFullName_LostFocus
        Case Is = vbKeyDelete
            ManufFullName_KeyPress KeyCode
    End Select
End Sub

Private Sub ManufFullName_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub ManufFullName_LostFocus()
    If changed = True Then
        If validateManufFullName(Me.ManufFullName) Then
            updateProdLine "ManufFullName", Me.ManufFullName, Me.ProductLine, "'"
            changed = False
        Else
            MsgBox "Invalid name, must be less than 32 characters."
            Me.ManufFullName.SetFocus
        End If
    End If
End Sub

Private Sub ManufFullNameCleaned_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            ManufFullNameCleaned_LostFocus
        Case Is = vbKeyDelete
            ManufFullNameCleaned_KeyPress KeyCode
    End Select
End Sub

Private Sub ManufFullNameCleaned_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub ManufFullNameCleaned_LostFocus()
    If changed = True Then
        If validateManufFullName(Me.ManufFullNameCleaned) Then
            If Left(DLookup("ManufFullNameCleaned", "ProductLine", "ProductLine='" & Me.ProductLine & "'"), 1) = "#" And Left(Me.ManufFullNameCleaned, 1) <> "#" Then
                If MsgBox("Un-hide this line code?", vbYesNo) = vbYes Then
                    updateProdLine "Hide", "0", Me.ProductLine, ""
                    fillingForm = True
                    Me.chkHide = 0
                    fillingForm = False
                End If
            ElseIf Left(Me.ManufFullNameCleaned, 1) = "#" Then
                If MsgBox("Hide this line code?", vbYesNo) = vbYes Then
                    updateProdLine "Hide", "1", Me.ProductLine, ""
                    fillingForm = True
                    Me.chkHide = 1
                    fillingForm = False
                End If
            End If
            updateProdLine "ManufFullNameCleaned", Me.ManufFullNameCleaned, Me.ProductLine, "'"
            changed = False
        Else
            MsgBox "Invalid name, must be less than 32 characters."
            Me.ManufFullNameCleaned.SetFocus
        End If
    End If
End Sub

Private Sub ManufFullNameWeb_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            ManufFullNameWeb_KeyPress KeyCode
        Case Is = vbKeyReturn
            ManufFullNameWeb_LostFocus
    End Select
End Sub

Private Sub ManufFullNameWeb_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub ManufFullNameWeb_LostFocus()
    If changed = True Then
        '9/18/08 - finally do this right, only changeable once, then
        'you're stuck with what you have. any changes after that
        'will need to be made in the db directly
        Me.ManufFullNameWeb = Trim(Me.ManufFullNameWeb)
        Dim oldname As String
        oldname = DLookup("ManufFullNameWeb", "ProductLine", "ProductLine='" & Me.ProductLine & "'")
        If oldname = "" Then
            If vbYes = MsgBox("Are you sure you want to set the web name to " & qq(Me.ManufFullNameWeb) & "?" & vbCrLf & vbCrLf & "This is the ONLY time you get to set it!" & vbCrLf & vbCrLf & "Also, this will make a web path for some reason!", vbYesNo) Then
                Dim weburl As String
                weburl = FormatForYahooURL(Me.ManufFullNameWeb)
                Dim oktocontinue As Boolean, WebPathID As String
                oktocontinue = False
                WebPathID = DLookup("ID", "WebPaths", "ParentID IS NULL AND URLIdentifier='" & weburl & "'")
                If WebPathID = "" Then
                    oktocontinue = True
                Else
                    If vbYes = MsgBox("This web name is already in use, items in this line will be merged, is this ok?", vbYesNo) Then
                        oktocontinue = True
                    Else
                        oktocontinue = False
                    End If
                End If
                If oktocontinue Then
                    updateProdLine "ManufFullNameWeb", Me.ManufFullNameWeb, Me.ProductLine, "'"
                    CreateHashes "ManufHash"
                    If WebPathID = "" Then
                        DB.Execute "INSERT INTO WebPaths ( WebsiteID, URLIdentifier, WebPathName, PathLevel, IsManufacturerPage ) VALUES ( " & CURRENT_WEBSITE_ID & ", '" & weburl & "', '" & EscapeSQuotes(Me.ManufFullNameWeb) & "', 1, 1 )"
                        Dim newid As String
                        newid = DLookup("@@IDENTITY", "WebPaths")
                        updateProdLine "WebSectionID", newid, Me.ProductLine, ""
                    End If
                    Me.chkHide.SetFocus
                    Disable Me.ManufFullNameWeb
                Else
                    Me.ManufFullNameWeb = oldname
                End If
            Else
                Me.ManufFullNameWeb = oldname
            End If
        Else
            MsgBox "web name is already set to " & qq(oldname) & ", so it's not changeable!" & vbCrLf & vbCrLf & "This is a bug, talk to Brian."
            Me.ManufFullNameWeb = oldname
        End If
        changed = False
    End If
End Sub

Private Sub MinimumAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            MinimumAmount_KeyPress KeyCode
        Case Is = vbKeyReturn
            MinimumAmount_LostFocus
    End Select
End Sub

Private Sub MinimumAmount_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub MinimumAmount_LostFocus()
    If changed = True Then
        If Me.MinimumAmount = "" Then
            updateProdLine "MinimumOrderAmount", "Null", Me.ProductLine, ""
            changed = False
        Else
            If validateCurrency(Me.MinimumAmount) Then
                Me.MinimumAmount = FormatCurrency(Me.MinimumAmount)
            End If
            updateProdLine "MinimumOrderAmount", Me.MinimumAmount, Me.ProductLine, "'"
        End If
        changed = False
    End If
End Sub

Private Sub WholesalePriceLevel_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        WholesalePriceLevel_LostFocus Index
    Else
        KeyAscii = LimitToNumbers(KeyAscii)
        If KeyAscii <> 0 Then
            changed = True
        End If
    End If
End Sub

Private Sub WholesalePriceLevel_LostFocus(Index As Integer)
    If changed = True Then
        If validateGeneralInteger(Me.WholesalePriceLevel(Index)) Then
            DB.Execute "UPDATE ProductLineWholesalePricing SET Percentage=" & Me.WholesalePriceLevel(Index) & " WHERE ProductLineID=" & Me.ProductLineID & " AND WholesalePriceLevel=" & Index
            changed = False
        Else
            MsgBox "Invalid Wholesale price level, must be an integer."
            Me.WholesalePriceLevel(Index).SetFocus
        End If
    End If
End Sub

Private Sub Prepaid_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            Prepaid_KeyPress KeyCode
        Case Is = vbKeyReturn
            Prepaid_LostFocus
    End Select
End Sub

Private Sub Prepaid_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub Prepaid_LostFocus()
    If changed = True Then
        If Me.Prepaid = "" Then
            updateProdLine "PrepaidAmount", "Null", Me.ProductLine, ""
        Else
            If validateCurrency(Me.Prepaid) Then
                Me.Prepaid = Format(Me.Prepaid, "Currency")
            End If
            updateProdLine "PrepaidAmount", Me.Prepaid, Me.ProductLine, "'"
        End If
        changed = False
    End If
End Sub

Private Sub PrepaidSpecial_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            PrepaidSpecial_KeyPress KeyCode
        Case Is = vbKeyReturn
            PrepaidSpecial_LostFocus
    End Select
End Sub

Private Sub PrepaidSpecial_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub PrepaidSpecial_LostFocus()
    If changed = True Then
        If Me.PrepaidSpecial = "" Then
            updateProdLine "PrepaidSpecialAmount", "Null", Me.ProductLine, ""
        Else
            If validateCurrency(Me.PrepaidSpecial) Then
                Me.PrepaidSpecial = Format(Me.PrepaidSpecial, "Currency")
            End If
            updateProdLine "PrepaidSpecialAmount", Me.PrepaidSpecial, Me.ProductLine, "'"
        End If
        changed = False
    End If
End Sub

Private Sub PrimaryVendor_Click()
    changed = True
    PrimaryVendor_LostFocus
End Sub

Private Sub PrimaryVendor_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.PrimaryVendor, KeyCode, Shift
End Sub

Private Sub PrimaryVendor_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.PrimaryVendor, KeyAscii
    If KeyAscii = vbKeyReturn Then
        PrimaryVendor_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub PrimaryVendor_LostFocus()
    AutoCompleteLostFocus Me.PrimaryVendor
    Dim newvend As String
    newvend = lookupPrimaryVendorNumber(Me.PrimaryVendor)
    If newvend <> DLookup("PrimaryVendorNumber", "ProductLine", "ProductLine='" & Me.ProductLine & "'") Then
        updateProdLine "PrimaryVendorNumber", newvend, Me.ProductLine, "'"
        requeryCoopVendor newvend
        If MsgBox("Change all items in this line code to new vendor?", vbYesNo) = vbYes Then
            DB.Execute "UPDATE InventoryMaster SET PrimaryVendor='" & newvend & "' WHERE ProductLine='" & Me.ProductLine & "'"
        End If
    End If
    changed = False
End Sub

Private Sub ProductLine_Click()
    changed = True
    ProductLine_LostFocus
End Sub

Private Sub ProductLine_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.ProductLine, KeyCode, Shift
End Sub

Private Sub ProductLine_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.ProductLine, KeyAscii
    If KeyAscii = vbKeyReturn Then
        ProductLine_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub ProductLine_LostFocus()
    AutoCompleteLostFocus Me.ProductLine
    If changed = True Then
        changed = False
        fillForm
    End If
End Sub

Private Sub realVendor_Click()
    realVendor_LostFocus
End Sub

Private Sub realVendor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        realVendor_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub realVendor_LostFocus()
    If changed = True Then
        Me.realVendor = UCase(Me.realVendor)
        If DLookup("RealVendor", "ProductLineCoopVendors", "CoopVendor='" & lookupPrimaryVendorNumber(Me.PrimaryVendor) & "' AND RealVendor='" & Me.realVendor & "'") = "" Then
            If MsgBox("Vendor " & Me.realVendor & " is not associated with " & Me.PrimaryVendor & ", add now?", vbYesNo) = vbYes Then
                DB.Execute "INSERT INTO ProductLineCoopVendors ( CoopVendor, RealVendor ) VALUES ( '" & lookupPrimaryVendorNumber(Me.PrimaryVendor) & "', '" & Me.realVendor & "' )"
                updateProdLine "RealVendor", Me.realVendor, Me.ProductLine, "'"
            Else
                Me.realVendor = ""
            End If
        Else
            updateProdLine "RealVendor", Me.realVendor, Me.ProductLine, "'"
        End If
    End If
End Sub

Private Sub ShipsFromZipCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            ShipsFromZipCode_KeyPress KeyCode
        Case Is = vbKeyReturn
            ShipsFromZipCode_LostFocus
    End Select
End Sub

Private Sub ShipsFromZipCode_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub ShipsFromZipCode_LostFocus()
    If changed = True Then
        If Me.ShipsFromZipCode = "" Or validateZipCode(Me.ShipsFromZipCode) Then
            updateProdLine "ShipsFromZip", Me.ShipsFromZipCode, Me.ProductLine, "'"
            changed = False
        Else
            MsgBox "invalid zip code, talk to brian if this is wrong."
            Me.ShipsFromZipCode.SetFocus
        End If
    End If
End Sub

Private Sub StdLeadTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        StdLeadTime_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub StdLeadTime_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.StdLeadTime) Then
            updateProdLine "StdLeadTime", Me.StdLeadTime, Me.ProductLine, ""
            changed = False
        Else
            MsgBox "Invalid Std lead time, must be an integer."
            Me.StdLeadTime.SetFocus
        End If
    End If
End Sub

Private Sub CompanyInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            CompanyInfo_KeyPress KeyCode
        Case Is = vbKeyReturn
            CompanyInfo_LostFocus
    End Select
End Sub

Private Sub CompanyInfo_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub CompanyInfo_LostFocus()
    If changed = True Then
        updateProdLine "CompanyInfo", Me.CompanyInfo, Me.ProductLine, "'"
        changed = False
    End If
End Sub

Private Sub Repair_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            Repair_KeyPress KeyCode
        Case Is = vbKeyReturn
            Repair_LostFocus
    End Select
End Sub

Private Sub Repair_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub Repair_LostFocus()
    If changed = True Then
        updateProdLine "Repair", Me.Repair, Me.ProductLine, "'"
        changed = False
    End If
End Sub

Private Sub ServiceCenters_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            ServiceCenters_KeyPress KeyCode
        Case Is = vbKeyReturn
            ServiceCenters_LostFocus
    End Select
End Sub

Private Sub ServiceCenters_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub ServiceCenters_LostFocus()
    If changed = True Then
        updateProdLine "ServiceCenters", Me.ServiceCenters, Me.ProductLine, "'"
        changed = False
    End If
End Sub

Private Sub Warranty_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            Warranty_KeyPress KeyCode
        Case Is = vbKeyReturn
            Warranty_LostFocus
    End Select
End Sub

Private Sub Warranty_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub Warranty_LostFocus()
    If changed = True Then
        updateProdLine "Warranty", Me.Warranty, Me.ProductLine, "'"
        changed = False
    End If
End Sub

Private Sub LiftGateCharge_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            LiftGateCharge_KeyPress Index, KeyCode
        Case Is = vbKeyReturn
            LiftGateCharge_LostFocus Index
    End Select
End Sub

Private Sub LiftGateCharge_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii, True)
    If KeyAscii <> 0 Then
        changed = True
    End If
End Sub

Private Sub LiftGateCharge_LostFocus(Index As Integer)
    If changed = True Then
        If 0 = DLookup("COUNT(*)", "ProductLineDropshipCharges", "ProductLineID=" & Me.ProductLineID & " AND AddressType=" & Index) Then
            DB.Execute "INSERT INTO ProductLineDropshipCharges ( ProductLineID, AddressType ) VALUES ( " & Me.ProductLineID & ", " & Index & " )"
        End If
        DB.Execute "UPDATE ProductLineDropshipCharges SET LiftGateCharge=" & IIf(Me.LiftGateCharge(Index) = "", "NULL", Me.LiftGateCharge(Index)) & " WHERE ProductLineID=" & Me.ProductLineID & " AND AddressType=" & Index
        changed = False
    End If
End Sub

Private Sub LiftGateWeightThreshold_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            LiftGateWeightThreshold_KeyPress Index, KeyCode
        Case Is = vbKeyReturn
            LiftGateWeightThreshold_LostFocus Index
    End Select
End Sub

Private Sub LiftGateWeightThreshold_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii, True)
    If KeyAscii <> 0 Then
        changed = True
    End If
End Sub

Private Sub LiftGateWeightThreshold_LostFocus(Index As Integer)
    If changed = True Then
        If 0 = DLookup("COUNT(*)", "ProductLineDropshipCharges", "ProductLineID=" & Me.ProductLineID & " AND AddressType=" & Index) Then
            DB.Execute "INSERT INTO ProductLineDropshipCharges ( ProductLineID, AddressType ) VALUES ( " & Me.ProductLineID & ", " & Index & " )"
        End If
        DB.Execute "UPDATE ProductLineDropshipCharges SET LiftGateWeightThreshold=" & IIf(Me.LiftGateWeightThreshold(Index) = "", "NULL", Me.LiftGateWeightThreshold(Index)) & " WHERE ProductLineID=" & Me.ProductLineID & " AND AddressType=" & Index
        changed = False
    End If
End Sub

Private Sub DropshipAdder_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            DropshipAdder_KeyPress Index, KeyCode
        Case Is = vbKeyReturn
            DropshipAdder_LostFocus Index
    End Select
End Sub

Private Sub DropshipAdder_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii, True)
    If KeyAscii <> 0 Then
        changed = True
    End If
End Sub

Private Sub DropshipAdder_LostFocus(Index As Integer)
    If changed = True Then
        If 0 = DLookup("COUNT(*)", "ProductLineDropshipCharges", "ProductLineID=" & Me.ProductLineID & " AND AddressType=" & Index) Then
            DB.Execute "INSERT INTO ProductLineDropshipCharges ( ProductLineID, AddressType ) VALUES ( " & Me.ProductLineID & ", " & Index & " )"
        End If
        DB.Execute "UPDATE ProductLineDropshipCharges SET AdderAmount=" & IIf(Me.DropshipAdder(Index) = "", "NULL", Me.DropshipAdder(Index)) & " WHERE ProductLineID=" & Me.ProductLineID & " AND AddressType=" & Index
        changed = False
    End If
End Sub

Private Sub DropshipAdderThreshold_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            DropshipAdderThreshold_KeyPress Index, KeyCode
        Case Is = vbKeyReturn
            DropshipAdderThreshold_LostFocus Index
    End Select
End Sub

Private Sub DropshipAdderThreshold_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii, True)
    If KeyAscii <> 0 Then
        changed = True
    End If
End Sub

Private Sub DropshipAdderThreshold_LostFocus(Index As Integer)
    If changed = True Then
        If 0 = DLookup("COUNT(*)", "ProductLineDropshipCharges", "ProductLineID=" & Me.ProductLineID & " AND AddressType=" & Index) Then
            DB.Execute "INSERT INTO ProductLineDropshipCharges ( ProductLineID, AddressType ) VALUES ( " & Me.ProductLineID & ", " & Index & " )"
        End If
        DB.Execute "UPDATE ProductLineDropshipCharges SET AdderThreshold=" & IIf(Me.DropshipAdderThreshold(Index) = "", "NULL", Me.DropshipAdderThreshold(Index)) & " WHERE ProductLineID=" & Me.ProductLineID & " AND AddressType=" & Index
        changed = False
    End If
End Sub

Private Sub chkDropshipAdderPercent_Click(Index As Integer)
    If Not fillingForm Then
        If Me.chkDropshipAdderPercent(Index) = vbUnchecked Then
            fillingForm = True
            Me.chkDropshipAdderPercent(Index) = vbChecked
            fillingForm = False
        Else
            If 0 = DLookup("COUNT(*)", "ProductLineDropshipCharges", "ProductLineID=" & Me.ProductLineID & " AND AddressType=" & Index) Then
                DB.Execute "INSERT INTO ProductLineDropshipCharges ( ProductLineID, AddressType ) VALUES ( " & Me.ProductLineID & ", " & Index & " )"
            End If
            DB.Execute "UPDATE ProductLineDropshipCharges SET AdderType=0 WHERE ProductLineID=" & Me.ProductLineID & " AND AddressType=" & Index
            fillingForm = True
            Me.chkDropshipAdderDollars(Index) = vbUnchecked
            fillingForm = False
        End If
    End If
End Sub

Private Sub chkDropshipAdderDollars_Click(Index As Integer)
    If Not fillingForm Then
        If Me.chkDropshipAdderDollars(Index) = vbUnchecked Then
            fillingForm = True
            Me.chkDropshipAdderDollars(Index) = vbChecked
            fillingForm = False
        Else
            If 0 = DLookup("COUNT(*)", "ProductLineDropshipCharges", "ProductLineID=" & Me.ProductLineID & " AND AddressType=" & Index) Then
                DB.Execute "INSERT INTO ProductLineDropshipCharges ( ProductLineID, AddressType ) VALUES ( " & Me.ProductLineID & ", " & Index & " )"
            End If
            DB.Execute "UPDATE ProductLineDropshipCharges SET AdderType=1 WHERE ProductLineID=" & Me.ProductLineID & " AND AddressType=" & Index
            fillingForm = True
            Me.chkDropshipAdderPercent(Index) = vbUnchecked
            fillingForm = False
        End If
    End If
End Sub

Private Sub DropshipAdderMaximum_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            DropshipAdderMaximum_KeyPress Index, KeyCode
        Case Is = vbKeyReturn
            DropshipAdderMaximum_LostFocus Index
    End Select
End Sub

Private Sub DropshipAdderMaximum_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii, True)
    If KeyAscii <> 0 Then
        changed = True
    End If
End Sub

Private Sub DropshipAdderMaximum_LostFocus(Index As Integer)
    If changed = True Then
        If 0 = DLookup("COUNT(*)", "ProductLineDropshipCharges", "ProductLineID=" & Me.ProductLineID & " AND AddressType=" & Index) Then
            DB.Execute "INSERT INTO ProductLineDropshipCharges ( ProductLineID, AddressType ) VALUES ( " & Me.ProductLineID & ", " & Index & " )"
        End If
        DB.Execute "UPDATE ProductLineDropshipCharges SET AdderMaximum=" & IIf(Me.DropshipAdderMaximum(Index) = "", "NULL", Me.DropshipAdderMaximum(Index)) & " WHERE ProductLineID=" & Me.ProductLineID & " AND AddressType=" & Index
        changed = False
    End If
End Sub


Private Function cleanDSTierCost(s As String) As String
    cleanDSTierCost = Replace(Replace(Replace(Replace(s, "$", ""), "%", ""), ",", ""), " / ppd", "")
End Function
