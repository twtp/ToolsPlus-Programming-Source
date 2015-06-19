VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#3.5#0"; "TPControls.ocx"
Begin VB.Form ProductSuggestionUtility 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Suggestions"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   16845
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnWebsiteStoreSelection 
      Caption         =   "Website ==>"
      Height          =   435
      Left            =   3150
      TabIndex        =   208
      Top             =   30
      Width           =   1245
   End
   Begin VB.CommandButton btnWebsiteSelection 
      Caption         =   "Website: Tools Plus"
      Height          =   435
      Left            =   4500
      TabIndex        =   177
      Tag             =   "0"
      Top             =   30
      Width           =   2025
   End
   Begin VB.Frame Frame6 
      Caption         =   "Generate For:"
      Height          =   615
      Left            =   13230
      TabIndex        =   159
      Top             =   510
      Width           =   2985
      Begin VB.ComboBox GenerationSelector 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   161
         Top             =   210
         Width           =   2745
      End
   End
   Begin VB.Frame Frame5 
      Height          =   5955
      Left            =   14160
      TabIndex        =   154
      Top             =   1830
      Width           =   2625
      Begin VB.CommandButton btnSelectedListMoveDown 
         Caption         =   "Move Down"
         Height          =   405
         Left            =   120
         TabIndex        =   158
         Top             =   5430
         Width           =   2385
      End
      Begin VB.CommandButton btnSelectedListMoveUp 
         Caption         =   "Move Up"
         Height          =   405
         Left            =   120
         TabIndex        =   157
         Top             =   5010
         Width           =   2385
      End
      Begin VB.CommandButton btnSelectedListAddSection 
         Caption         =   "Add New Section"
         Height          =   405
         Left            =   120
         TabIndex        =   156
         Top             =   4440
         Width           =   2385
      End
      Begin VB.ListBox SelectedList 
         Height          =   4155
         Left            =   90
         TabIndex        =   155
         Top             =   180
         Width           =   2445
      End
   End
   Begin VB.CommandButton btnSelectedListRemove 
      Caption         =   "á"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   13380
      TabIndex        =   153
      Top             =   4410
      Width           =   735
   End
   Begin VB.CommandButton btnSelectedListAdd 
      Caption         =   "â"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   13380
      TabIndex        =   152
      Top             =   3720
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtering:"
      Height          =   1275
      Left            =   60
      TabIndex        =   14
      Top             =   510
      Width           =   7575
      Begin VB.TextBox filterInputText 
         Height          =   285
         Left            =   1860
         TabIndex        =   176
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton filterBtnApply 
         Caption         =   "Apply Rule"
         Height          =   315
         Left            =   2460
         TabIndex        =   148
         Top             =   870
         Visible         =   0   'False
         Width           =   1815
      End
      Begin TPControls.DateDropdown filterInputDate 
         Height          =   315
         Left            =   1860
         TabIndex        =   147
         Top             =   180
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
      End
      Begin VB.ComboBox filterInputBoolean 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   146
         Top             =   840
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox filterInputInteger 
         Height          =   285
         Left            =   2730
         TabIndex        =   145
         Top             =   510
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.TextBox filterInputDecimal 
         Height          =   285
         Left            =   2730
         TabIndex        =   144
         Top             =   180
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.ComboBox filterInputComparison 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   143
         Top             =   510
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox filterFieldSelector 
         Height          =   315
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   142
         Top             =   510
         Width           =   1755
      End
      Begin VB.ListBox criteriaFilter 
         Height          =   1035
         Left            =   4440
         TabIndex        =   141
         Top             =   180
         Width           =   3045
      End
   End
   Begin VB.Frame Frame4 
      Height          =   5955
      Left            =   60
      TabIndex        =   5
      Top             =   1830
      Width           =   13275
      Begin VB.VScrollBar itemScrollBar 
         Height          =   5205
         Left            =   12930
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   630
         Width           =   285
      End
      Begin VB.Frame ItemFrame 
         BackColor       =   &H006A240A&
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   13
         Left            =   30
         TabIndex        =   132
         Top             =   5400
         Width           =   12915
         Begin VB.TextBox AverageCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   13
            Left            =   11970
            Locked          =   -1  'True
            TabIndex        =   205
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox StandardCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   13
            Left            =   11070
            Locked          =   -1  'True
            TabIndex        =   192
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenInInventory 
            Caption         =   "I"
            Height          =   225
            Index           =   13
            Left            =   900
            TabIndex        =   175
            Top             =   60
            Width           =   255
         End
         Begin VB.TextBox ItemNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   13
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   140
            Top             =   30
            Width           =   1425
         End
         Begin VB.TextBox ItemDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   13
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   139
            Top             =   30
            Width           =   3765
         End
         Begin VB.TextBox QtySold 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   13
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   138
            Top             =   30
            Width           =   825
         End
         Begin VB.TextBox NumTransactions 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   13
            Left            =   7260
            Locked          =   -1  'True
            TabIndex        =   137
            Top             =   30
            Width           =   1095
         End
         Begin VB.TextBox QtyOnHand 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   13
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   136
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox QtyOnOrder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   13
            Left            =   9270
            Locked          =   -1  'True
            TabIndex        =   135
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox InternetPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   13
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   134
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenWebsite 
            Caption         =   "WWW"
            Height          =   225
            Index           =   13
            Left            =   60
            TabIndex        =   133
            Top             =   60
            Width           =   795
         End
      End
      Begin VB.Frame ItemFrame 
         BackColor       =   &H006A240A&
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   12
         Left            =   30
         TabIndex        =   123
         Top             =   5040
         Width           =   12915
         Begin VB.TextBox AverageCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   12
            Left            =   11970
            Locked          =   -1  'True
            TabIndex        =   204
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox StandardCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   12
            Left            =   11070
            Locked          =   -1  'True
            TabIndex        =   191
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenInInventory 
            Caption         =   "I"
            Height          =   225
            Index           =   12
            Left            =   900
            TabIndex        =   174
            Top             =   60
            Width           =   255
         End
         Begin VB.TextBox ItemNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   12
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   131
            Top             =   30
            Width           =   1425
         End
         Begin VB.TextBox ItemDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   12
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   130
            Top             =   30
            Width           =   3765
         End
         Begin VB.TextBox QtySold 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   12
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   129
            Top             =   30
            Width           =   825
         End
         Begin VB.TextBox NumTransactions 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   12
            Left            =   7260
            Locked          =   -1  'True
            TabIndex        =   128
            Top             =   30
            Width           =   1095
         End
         Begin VB.TextBox QtyOnHand 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   12
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   127
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox QtyOnOrder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   12
            Left            =   9270
            Locked          =   -1  'True
            TabIndex        =   126
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox InternetPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   12
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   125
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenWebsite 
            Caption         =   "WWW"
            Height          =   225
            Index           =   12
            Left            =   60
            TabIndex        =   124
            Top             =   60
            Width           =   795
         End
      End
      Begin VB.Frame ItemFrame 
         BackColor       =   &H006A240A&
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   11
         Left            =   30
         TabIndex        =   114
         Top             =   4680
         Width           =   12915
         Begin VB.TextBox AverageCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   11
            Left            =   11970
            Locked          =   -1  'True
            TabIndex        =   203
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox StandardCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   11
            Left            =   11070
            Locked          =   -1  'True
            TabIndex        =   190
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenInInventory 
            Caption         =   "I"
            Height          =   225
            Index           =   11
            Left            =   900
            TabIndex        =   173
            Top             =   60
            Width           =   255
         End
         Begin VB.TextBox ItemNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   11
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   122
            Top             =   30
            Width           =   1425
         End
         Begin VB.TextBox ItemDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   11
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   121
            Top             =   30
            Width           =   3765
         End
         Begin VB.TextBox QtySold 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   11
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   120
            Top             =   30
            Width           =   825
         End
         Begin VB.TextBox NumTransactions 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   11
            Left            =   7260
            Locked          =   -1  'True
            TabIndex        =   119
            Top             =   30
            Width           =   1095
         End
         Begin VB.TextBox QtyOnHand 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   11
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   118
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox QtyOnOrder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   11
            Left            =   9270
            Locked          =   -1  'True
            TabIndex        =   117
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox InternetPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   11
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   116
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenWebsite 
            Caption         =   "WWW"
            Height          =   225
            Index           =   11
            Left            =   60
            TabIndex        =   115
            Top             =   60
            Width           =   795
         End
      End
      Begin VB.Frame ItemFrame 
         BackColor       =   &H006A240A&
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   10
         Left            =   30
         TabIndex        =   105
         Top             =   4320
         Width           =   12915
         Begin VB.TextBox AverageCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   10
            Left            =   11970
            Locked          =   -1  'True
            TabIndex        =   202
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox StandardCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   10
            Left            =   11070
            Locked          =   -1  'True
            TabIndex        =   189
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenInInventory 
            Caption         =   "I"
            Height          =   225
            Index           =   10
            Left            =   900
            TabIndex        =   172
            Top             =   60
            Width           =   255
         End
         Begin VB.TextBox ItemNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   10
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   113
            Top             =   30
            Width           =   1425
         End
         Begin VB.TextBox ItemDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   10
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   112
            Top             =   30
            Width           =   3765
         End
         Begin VB.TextBox QtySold 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   10
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   111
            Top             =   30
            Width           =   825
         End
         Begin VB.TextBox NumTransactions 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   10
            Left            =   7260
            Locked          =   -1  'True
            TabIndex        =   110
            Top             =   30
            Width           =   1095
         End
         Begin VB.TextBox QtyOnHand 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   10
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   109
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox QtyOnOrder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   10
            Left            =   9270
            Locked          =   -1  'True
            TabIndex        =   108
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox InternetPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   10
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   107
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenWebsite 
            Caption         =   "WWW"
            Height          =   225
            Index           =   10
            Left            =   60
            TabIndex        =   106
            Top             =   60
            Width           =   795
         End
      End
      Begin VB.Frame ItemFrame 
         BackColor       =   &H006A240A&
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   9
         Left            =   30
         TabIndex        =   80
         Top             =   3960
         Width           =   12915
         Begin VB.TextBox AverageCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   9
            Left            =   11970
            Locked          =   -1  'True
            TabIndex        =   201
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox StandardCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   9
            Left            =   11070
            Locked          =   -1  'True
            TabIndex        =   188
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenInInventory 
            Caption         =   "I"
            Height          =   225
            Index           =   9
            Left            =   900
            TabIndex        =   171
            Top             =   60
            Width           =   255
         End
         Begin VB.CommandButton btnOpenWebsite 
            Caption         =   "WWW"
            Height          =   225
            Index           =   9
            Left            =   60
            TabIndex        =   104
            Top             =   60
            Width           =   795
         End
         Begin VB.TextBox InternetPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   9
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   94
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox QtyOnOrder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   9
            Left            =   9270
            Locked          =   -1  'True
            TabIndex        =   93
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox QtyOnHand 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   9
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   92
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox NumTransactions 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   9
            Left            =   7260
            Locked          =   -1  'True
            TabIndex        =   91
            Top             =   30
            Width           =   1095
         End
         Begin VB.TextBox QtySold 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   9
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   90
            Top             =   30
            Width           =   825
         End
         Begin VB.TextBox ItemDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   9
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   89
            Top             =   30
            Width           =   3765
         End
         Begin VB.TextBox ItemNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   9
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   88
            Top             =   30
            Width           =   1425
         End
      End
      Begin VB.Frame ItemFrame 
         BackColor       =   &H006A240A&
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   8
         Left            =   30
         TabIndex        =   79
         Top             =   3600
         Width           =   12915
         Begin VB.TextBox AverageCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   8
            Left            =   11970
            Locked          =   -1  'True
            TabIndex        =   200
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox StandardCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   8
            Left            =   11070
            Locked          =   -1  'True
            TabIndex        =   187
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenInInventory 
            Caption         =   "I"
            Height          =   225
            Index           =   8
            Left            =   900
            TabIndex        =   170
            Top             =   60
            Width           =   255
         End
         Begin VB.CommandButton btnOpenWebsite 
            Caption         =   "WWW"
            Height          =   225
            Index           =   8
            Left            =   60
            TabIndex        =   103
            Top             =   60
            Width           =   795
         End
         Begin VB.TextBox InternetPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   8
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   87
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox QtyOnOrder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   8
            Left            =   9270
            Locked          =   -1  'True
            TabIndex        =   86
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox QtyOnHand 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   8
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   85
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox NumTransactions 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   8
            Left            =   7260
            Locked          =   -1  'True
            TabIndex        =   84
            Top             =   30
            Width           =   1095
         End
         Begin VB.TextBox QtySold 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   8
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   83
            Top             =   30
            Width           =   825
         End
         Begin VB.TextBox ItemDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   8
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   82
            Top             =   30
            Width           =   3765
         End
         Begin VB.TextBox ItemNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   8
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   81
            Top             =   30
            Width           =   1425
         End
      End
      Begin VB.Frame ItemFrame 
         BackColor       =   &H006A240A&
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   7
         Left            =   30
         TabIndex        =   50
         Top             =   3240
         Width           =   12915
         Begin VB.TextBox AverageCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   7
            Left            =   11970
            Locked          =   -1  'True
            TabIndex        =   199
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox StandardCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   7
            Left            =   11070
            Locked          =   -1  'True
            TabIndex        =   186
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenInInventory 
            Caption         =   "I"
            Height          =   225
            Index           =   7
            Left            =   900
            TabIndex        =   169
            Top             =   60
            Width           =   255
         End
         Begin VB.CommandButton btnOpenWebsite 
            Caption         =   "WWW"
            Height          =   225
            Index           =   7
            Left            =   60
            TabIndex        =   102
            Top             =   60
            Width           =   795
         End
         Begin VB.TextBox ItemNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   7
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   30
            Width           =   1425
         End
         Begin VB.TextBox ItemDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   7
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   77
            Top             =   30
            Width           =   3765
         End
         Begin VB.TextBox QtySold 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   7
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   76
            Top             =   30
            Width           =   825
         End
         Begin VB.TextBox NumTransactions 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   7
            Left            =   7260
            Locked          =   -1  'True
            TabIndex        =   75
            Top             =   30
            Width           =   1095
         End
         Begin VB.TextBox QtyOnHand 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   7
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   74
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox QtyOnOrder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   7
            Left            =   9270
            Locked          =   -1  'True
            TabIndex        =   73
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox InternetPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   7
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   30
            Width           =   885
         End
      End
      Begin VB.Frame ItemFrame 
         BackColor       =   &H006A240A&
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   6
         Left            =   30
         TabIndex        =   49
         Top             =   2880
         Width           =   12915
         Begin VB.TextBox AverageCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   6
            Left            =   11970
            Locked          =   -1  'True
            TabIndex        =   198
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox StandardCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   6
            Left            =   11070
            Locked          =   -1  'True
            TabIndex        =   185
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenInInventory 
            Caption         =   "I"
            Height          =   225
            Index           =   6
            Left            =   900
            TabIndex        =   168
            Top             =   60
            Width           =   255
         End
         Begin VB.CommandButton btnOpenWebsite 
            Caption         =   "WWW"
            Height          =   225
            Index           =   6
            Left            =   60
            TabIndex        =   101
            Top             =   60
            Width           =   795
         End
         Begin VB.TextBox ItemNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   6
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   71
            Top             =   30
            Width           =   1425
         End
         Begin VB.TextBox ItemDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   6
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   30
            Width           =   3765
         End
         Begin VB.TextBox QtySold 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   6
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   30
            Width           =   825
         End
         Begin VB.TextBox NumTransactions 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   6
            Left            =   7260
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   30
            Width           =   1095
         End
         Begin VB.TextBox QtyOnHand 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   6
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox QtyOnOrder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   6
            Left            =   9270
            Locked          =   -1  'True
            TabIndex        =   66
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox InternetPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   6
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   30
            Width           =   885
         End
      End
      Begin VB.Frame ItemFrame 
         BackColor       =   &H006A240A&
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   5
         Left            =   30
         TabIndex        =   48
         Top             =   2520
         Width           =   12915
         Begin VB.TextBox AverageCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   5
            Left            =   11970
            Locked          =   -1  'True
            TabIndex        =   197
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox StandardCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   5
            Left            =   11070
            Locked          =   -1  'True
            TabIndex        =   184
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenInInventory 
            Caption         =   "I"
            Height          =   225
            Index           =   5
            Left            =   900
            TabIndex        =   167
            Top             =   60
            Width           =   255
         End
         Begin VB.CommandButton btnOpenWebsite 
            Caption         =   "WWW"
            Height          =   225
            Index           =   5
            Left            =   60
            TabIndex        =   100
            Top             =   60
            Width           =   795
         End
         Begin VB.TextBox ItemNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   5
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   30
            Width           =   1425
         End
         Begin VB.TextBox ItemDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   5
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   30
            Width           =   3765
         End
         Begin VB.TextBox QtySold 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   5
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   30
            Width           =   825
         End
         Begin VB.TextBox NumTransactions 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   5
            Left            =   7260
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   30
            Width           =   1095
         End
         Begin VB.TextBox QtyOnHand 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   5
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox QtyOnOrder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   5
            Left            =   9270
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox InternetPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   5
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   30
            Width           =   885
         End
      End
      Begin VB.Frame ItemFrame 
         BackColor       =   &H006A240A&
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   4
         Left            =   30
         TabIndex        =   47
         Top             =   2160
         Width           =   12915
         Begin VB.TextBox AverageCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   4
            Left            =   11970
            Locked          =   -1  'True
            TabIndex        =   196
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox StandardCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   4
            Left            =   11070
            Locked          =   -1  'True
            TabIndex        =   183
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenInInventory 
            Caption         =   "I"
            Height          =   225
            Index           =   4
            Left            =   900
            TabIndex        =   166
            Top             =   60
            Width           =   255
         End
         Begin VB.CommandButton btnOpenWebsite 
            Caption         =   "WWW"
            Height          =   225
            Index           =   4
            Left            =   60
            TabIndex        =   99
            Top             =   60
            Width           =   795
         End
         Begin VB.TextBox ItemNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   4
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   30
            Width           =   1425
         End
         Begin VB.TextBox ItemDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   4
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   30
            Width           =   3765
         End
         Begin VB.TextBox QtySold 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   4
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   30
            Width           =   825
         End
         Begin VB.TextBox NumTransactions 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   4
            Left            =   7260
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   30
            Width           =   1095
         End
         Begin VB.TextBox QtyOnHand 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   4
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox QtyOnOrder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   4
            Left            =   9270
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox InternetPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   4
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   30
            Width           =   885
         End
      End
      Begin VB.Frame ItemFrame 
         BackColor       =   &H006A240A&
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   3
         Left            =   30
         TabIndex        =   39
         Top             =   1800
         Width           =   12915
         Begin VB.TextBox AverageCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   3
            Left            =   11970
            Locked          =   -1  'True
            TabIndex        =   195
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox StandardCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   3
            Left            =   11070
            Locked          =   -1  'True
            TabIndex        =   182
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenInInventory 
            Caption         =   "I"
            Height          =   225
            Index           =   3
            Left            =   900
            TabIndex        =   165
            Top             =   60
            Width           =   255
         End
         Begin VB.CommandButton btnOpenWebsite 
            Caption         =   "WWW"
            Height          =   225
            Index           =   3
            Left            =   60
            TabIndex        =   98
            Top             =   60
            Width           =   795
         End
         Begin VB.TextBox ItemNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   3
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   30
            Width           =   1425
         End
         Begin VB.TextBox ItemDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   3
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   30
            Width           =   3765
         End
         Begin VB.TextBox QtySold 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   3
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   30
            Width           =   825
         End
         Begin VB.TextBox NumTransactions 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   3
            Left            =   7260
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   30
            Width           =   1095
         End
         Begin VB.TextBox QtyOnHand 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   3
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox QtyOnOrder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   3
            Left            =   9270
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox InternetPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   3
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   30
            Width           =   885
         End
      End
      Begin VB.Frame ItemFrame 
         BackColor       =   &H006A240A&
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   2
         Left            =   30
         TabIndex        =   31
         Top             =   1440
         Width           =   12915
         Begin VB.TextBox AverageCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   2
            Left            =   11970
            Locked          =   -1  'True
            TabIndex        =   194
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox StandardCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   2
            Left            =   11070
            Locked          =   -1  'True
            TabIndex        =   181
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenInInventory 
            Caption         =   "I"
            Height          =   225
            Index           =   2
            Left            =   900
            TabIndex        =   164
            Top             =   60
            Width           =   255
         End
         Begin VB.CommandButton btnOpenWebsite 
            Caption         =   "WWW"
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   97
            Top             =   60
            Width           =   795
         End
         Begin VB.TextBox ItemNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   2
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   30
            Width           =   1425
         End
         Begin VB.TextBox ItemDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   2
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   30
            Width           =   3765
         End
         Begin VB.TextBox QtySold 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   2
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   30
            Width           =   825
         End
         Begin VB.TextBox NumTransactions 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   2
            Left            =   7260
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   30
            Width           =   1095
         End
         Begin VB.TextBox QtyOnHand 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   2
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox QtyOnOrder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   2
            Left            =   9270
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox InternetPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   2
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   30
            Width           =   885
         End
      End
      Begin VB.Frame ItemFrame 
         BackColor       =   &H006A240A&
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   1
         Left            =   30
         TabIndex        =   23
         Top             =   1080
         Width           =   12915
         Begin VB.TextBox AverageCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   1
            Left            =   11970
            Locked          =   -1  'True
            TabIndex        =   193
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox StandardCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   1
            Left            =   11070
            Locked          =   -1  'True
            TabIndex        =   180
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenInInventory 
            Caption         =   "I"
            Height          =   225
            Index           =   1
            Left            =   900
            TabIndex        =   163
            Top             =   60
            Width           =   255
         End
         Begin VB.CommandButton btnOpenWebsite 
            Caption         =   "WWW"
            Height          =   225
            Index           =   1
            Left            =   60
            TabIndex        =   96
            Top             =   60
            Width           =   795
         End
         Begin VB.TextBox ItemNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   1
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   30
            Width           =   1425
         End
         Begin VB.TextBox ItemDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   1
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   30
            Width           =   3765
         End
         Begin VB.TextBox QtySold 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   1
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   30
            Width           =   825
         End
         Begin VB.TextBox NumTransactions 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   1
            Left            =   7260
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   30
            Width           =   1095
         End
         Begin VB.TextBox QtyOnHand 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   1
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox QtyOnOrder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   1
            Left            =   9270
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox InternetPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   1
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   30
            Width           =   885
         End
      End
      Begin VB.Frame ItemFrame 
         BackColor       =   &H006A240A&
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   0
         Left            =   30
         TabIndex        =   15
         Top             =   720
         Width           =   12915
         Begin VB.TextBox AverageCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   11970
            Locked          =   -1  'True
            TabIndex        =   179
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox StandardCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   11070
            Locked          =   -1  'True
            TabIndex        =   178
            Top             =   30
            Width           =   885
         End
         Begin VB.CommandButton btnOpenInInventory 
            Caption         =   "I"
            Height          =   225
            Index           =   0
            Left            =   900
            TabIndex        =   162
            Top             =   60
            Width           =   255
         End
         Begin VB.CommandButton btnOpenWebsite 
            Caption         =   "WWW"
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   95
            Top             =   60
            Width           =   795
         End
         Begin VB.TextBox ItemNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   30
            Width           =   1425
         End
         Begin VB.TextBox ItemDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   30
            Width           =   3765
         End
         Begin VB.TextBox QtySold 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   30
            Width           =   825
         End
         Begin VB.TextBox NumTransactions 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   7260
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   30
            Width           =   1095
         End
         Begin VB.TextBox QtyOnHand 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox QtyOnOrder 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   9270
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   30
            Width           =   885
         End
         Begin VB.TextBox InternetPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   10170
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   30
            Width           =   885
         End
      End
      Begin VB.Label headerLabel 
         Alignment       =   2  'Center
         Caption         =   "Average Cost"
         Height          =   405
         Index           =   8
         Left            =   12000
         TabIndex        =   207
         Tag             =   "AverageCost"
         Top             =   180
         Width           =   885
      End
      Begin VB.Label headerLabel 
         Alignment       =   2  'Center
         Caption         =   "Standard Cost"
         Height          =   405
         Index           =   7
         Left            =   11100
         TabIndex        =   206
         Tag             =   "StandardCost"
         Top             =   180
         Width           =   885
      End
      Begin VB.Line Line2 
         X1              =   30
         X2              =   13200
         Y1              =   5820
         Y2              =   5820
      End
      Begin VB.Label headerLabel 
         Alignment       =   2  'Center
         Caption         =   "Internet Price"
         Height          =   405
         Index           =   6
         Left            =   10200
         TabIndex        =   13
         Tag             =   "InternetPrice"
         Top             =   180
         Width           =   885
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   13200
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Label headerLabel 
         Alignment       =   2  'Center
         Caption         =   "Qty on Order"
         Height          =   405
         Index           =   5
         Left            =   9300
         TabIndex        =   11
         Tag             =   "QtyOnOrder"
         Top             =   180
         Width           =   885
      End
      Begin VB.Label headerLabel 
         Alignment       =   2  'Center
         Caption         =   "Qty on Hand"
         Height          =   405
         Index           =   4
         Left            =   8400
         TabIndex        =   10
         Tag             =   "QtyOnHand"
         Top             =   180
         Width           =   885
      End
      Begin VB.Label headerLabel 
         Alignment       =   2  'Center
         Caption         =   "# of Transactions"
         Height          =   405
         Index           =   3
         Left            =   7290
         TabIndex        =   9
         Tag             =   "NumTransactions"
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label headerLabel 
         Alignment       =   2  'Center
         Caption         =   "Quantity Sold"
         Height          =   405
         Index           =   2
         Left            =   6450
         TabIndex        =   8
         Tag             =   "QtySold"
         Top             =   180
         Width           =   885
      End
      Begin VB.Label headerLabel 
         Caption         =   "Description"
         Height          =   255
         Index           =   1
         Left            =   2820
         TabIndex        =   7
         Tag             =   "ItemDescription"
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label headerLabel 
         Alignment       =   2  'Center
         Caption         =   "Item Number"
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
         Index           =   0
         Left            =   1260
         TabIndex        =   6
         Tag             =   "ItemNumber"
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Sales Timeframe:"
      Height          =   615
      Left            =   7740
      TabIndex        =   4
      Top             =   1170
      Width           =   3525
      Begin VB.CommandButton btnChangeSalesTimeFrame 
         Caption         =   "Go"
         Height          =   285
         Left            =   2940
         TabIndex        =   160
         ToolTipText     =   "GO GO GADGET DATERANGE!"
         Top             =   270
         Width           =   525
      End
      Begin TPControls.DateDropdown SalesFromDate 
         Height          =   315
         Left            =   60
         TabIndex        =   149
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
      End
      Begin TPControls.DateDropdown SalesToDate 
         Height          =   315
         Left            =   1560
         TabIndex        =   151
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
      End
      Begin VB.Label generalLabel 
         Caption         =   "to"
         Height          =   225
         Index           =   0
         Left            =   1380
         TabIndex        =   150
         Top             =   300
         Width           =   225
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item Name Filter:"
      Height          =   615
      Left            =   7740
      TabIndex        =   1
      Top             =   510
      Width           =   2595
      Begin VB.TextBox FragmentToSearch 
         Height          =   285
         Left            =   1050
         TabIndex        =   3
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lblFragment 
         Alignment       =   1  'Right Justify
         Caption         =   "Fragment:"
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.Label generalLabel 
      Caption         =   "Product Suggestions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2745
   End
End
Attribute VB_Name = "ProductSuggestionUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NUM_LINES As Long = 14

Private ItemList() As Dictionary
Private FilteredItemList() As Dictionary
Private ForeColorsList(NUM_LINES - 1) As Long

Private SelectionArray() As Variant

Private lastQSearch As String
Private currentsort As String
Private currentSortDesc As Boolean
Private lastTopItemDx As Long
Private lastSelItemDx As Long
Private lastHistoryStart As Date
Private lastHistoryEnd As Date

Private incomingSectionName As String

Public Sub SetIncomingSectionName(newSection As String)
    incomingSectionName = newSection
End Sub


Private Sub requerySelectedList()
    Me.SelectedList.Clear
    
    On Error Resume Next
    Dim foo As Long
    foo = UBound(SelectionArray)
    If Err.Number = 9 Then
        'done
        Err.Clear
        Exit Sub
    End If
    
    On Error GoTo 0
    
    Dim i As Long, j As Long
    For i = 0 To UBound(SelectionArray)
        Me.SelectedList.AddItem SelectionArray(i)(0)
        For j = 1 To UBound(SelectionArray(i))
            Me.SelectedList.AddItem "    " & SelectionArray(i)(j)
        Next j
    Next i
End Sub

Private Sub setSelectedListSelection(SectionName As String, Optional itemName As String = "")
    If Left(itemName, 4) = "    " Then
        itemName = Mid(itemName, 5)
    End If
    Dim i As Long
    Dim foundSection As Boolean
    For i = 0 To Me.SelectedList.ListCount - 1
        If Not foundSection Then
            If Me.SelectedList.list(i) = SectionName Then
                If itemName = "" Then
                    Me.SelectedList.ListIndex = i
                    Exit For
                Else
                    foundSection = True
                End If
            End If
        Else
            If Me.SelectedList.list(i) = "    " & itemName Then
                Me.SelectedList.ListIndex = i
                Exit For
            End If
        End If
    Next i
End Sub

Private Sub btnWebsiteSelection_Click()
    Dim newsite As Long
    newsite = CLng(Me.btnWebsiteSelection.tag) + 1
    If Not WebsiteNameHash.exists(newsite) Then
        newsite = 0
    End If
    Me.btnWebsiteSelection.tag = newsite
    Me.btnWebsiteSelection.Caption = "Website: " & WebsiteNameHash.item(newsite)
    requeryFullItemList
    'TODO: display too?
End Sub

Private Sub btnWebsiteStoreSelection_Click()
    If Me.btnWebsiteStoreSelection.Caption = "Store" Then
        Me.btnWebsiteStoreSelection.Caption = "Website ==>"
        Me.btnWebsiteSelection.Visible = True
    Else
        Me.btnWebsiteStoreSelection.Caption = "Store"
        Me.btnWebsiteSelection.Visible = False
    End If
    requeryFullItemList
End Sub

Private Sub criteriaFilter_Click()
    'TODO: push to filter?
End Sub

Private Sub criteriaFilter_DblClick()
    Me.criteriaFilter.RemoveItem Me.criteriaFilter.ListIndex
    refilterList
End Sub

Private Sub filterInputDecimal_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii, True)
End Sub

Private Sub filterInputInteger_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
End Sub

Private Sub Form_Load()
    Dim tabs(1) As Long
    tabs(0) = 84
    tabs(1) = 108
    SetListTabStops Me.criteriaFilter.hwnd, tabs

    'these all start out invisible
    'left alignments are ok in the designer, so just set the vertical
    Me.filterInputBoolean.Top = Me.filterFieldSelector.Top
    Me.filterInputComparison.Top = Me.filterFieldSelector.Top
    Me.filterInputDate.Top = Me.filterFieldSelector.Top
    Me.filterInputDecimal.Top = Me.filterFieldSelector.Top
    Me.filterInputInteger.Top = Me.filterFieldSelector.Top
    Me.filterInputText.Top = Me.filterFieldSelector.Top
    
    Me.SalesToDate.value = Date
    Me.SalesFromDate.value = Date - 14
    
    'SelectionArray = Split("", "zerolengtharray")
    
    Me.itemScrollBar.LargeChange = NUM_LINES - 1
    lastTopItemDx = 0
    lastSelItemDx = -1
    lastHistoryStart = Me.SalesFromDate.value
    lastHistoryEnd = Me.SalesToDate.value
    setFocusBackground lastSelItemDx
    currentsort = "ItemNumber"
    currentSortDesc = False
    requeryFilterComparisons
    requeryFilterBooleans
    requeryFilterFields
    setDefaultFilters
    requeryGenerationSelector
    requeryFullItemList
End Sub

Private Sub requeryFilterComparisons()
    Me.filterInputComparison.Clear
    Me.filterInputComparison.AddItem "=="
    Me.filterInputComparison.AddItem "!="
    Me.filterInputComparison.AddItem ">"
    Me.filterInputComparison.AddItem "<"
    Me.filterInputComparison.AddItem ">="
    Me.filterInputComparison.AddItem "<="
    ExpandDropDownToFit Me.filterInputComparison
End Sub

Private Sub requeryFilterBooleans()
    Me.filterInputBoolean.Clear
    Me.filterInputBoolean.AddItem "True"
    Me.filterInputBoolean.AddItem "False"
    ExpandDropDownToFit Me.filterInputBoolean
End Sub

Private Sub requeryFilterFields()
    Me.filterFieldSelector.Clear
    Me.filterFieldSelector.AddItem "Internet Price"
    Me.filterFieldSelector.AddItem "Qty on Order"
    Me.filterFieldSelector.AddItem "Qty on Hand"
    Me.filterFieldSelector.AddItem "Number of Transactions"
    Me.filterFieldSelector.AddItem "Qty Sold"
    Me.filterFieldSelector.AddItem "Reconditioned"
    Me.filterFieldSelector.AddItem "Price Drops"
    Me.filterFieldSelector.AddItem "New Items Since"
    Me.filterFieldSelector.AddItem "Rebates/Specials"
    Me.filterFieldSelector.AddItem "Discontinued"
    Me.filterFieldSelector.AddItem "Triad Code"
    Me.filterFieldSelector.AddItem "Standard Cost"
    Me.filterFieldSelector.AddItem "Average Cost"
    ExpandDropDownToFit Me.filterFieldSelector
End Sub

Private Sub setDefaultFilters()
    Me.criteriaFilter.AddItem "Discontinued" & vbTab & "False"
End Sub

Private Sub filterFieldSelector_Click()
    If Me.filterFieldSelector.ListIndex = -1 Then
        Exit Sub
    End If
    Me.filterInputBoolean.ListIndex = -1
    Me.filterInputComparison.ListIndex = -1
    Me.filterInputDate.value = Date
    Me.filterInputDecimal = ""
    Me.filterInputInteger = ""
    Me.filterInputText = ""
    Me.filterInputBoolean.Visible = False
    Me.filterInputComparison.Visible = False
    Me.filterInputDate.Visible = False
    Me.filterInputDecimal.Visible = False
    Me.filterInputInteger.Visible = False
    Me.filterInputText.Visible = False
    Me.filterBtnApply.Visible = True
    Select Case Me.filterFieldSelector
        'decimal fields
        Case Is = "Internet Price", "Price Drops", "Standard Cost", "Average Cost"
            Me.filterInputComparison.Visible = True
            Me.filterInputDecimal.Visible = True
        'integer fields
        Case Is = "Qty on Order", "Qty on Hand", "Number of Transactions", "Qty Sold"
            Me.filterInputComparison.Visible = True
            Me.filterInputInteger.Visible = True
        'boolean fields
        Case Is = "Reconditioned", "Discontinued", "Rebates/Specials"
            Me.filterInputBoolean.Visible = True
        'date fields
        Case Is = "New Items Since"
            Me.filterInputDate.value = Me.SalesFromDate.value
            Me.filterInputDate.Visible = True
        'plain text fields
        Case Is = "Triad Code"
            Me.filterInputText.Visible = True
        'oops
        Case Else
            MsgBox "Tell Brian to add " & qq(Me.filterFieldSelector) & " to the filterFieldSelector_Click() select clause!"
            Me.filterBtnApply.Visible = False
    End Select
End Sub

Private Sub filterBtnApply_Click()
    Dim filterString As String
    filterString = ""
    Select Case Me.filterFieldSelector
        Case Is = "Internet Price", "Price Drops", "Standard Cost", "Average Cost"
            If Me.filterInputComparison.ListIndex = -1 Then
                MsgBox "Select a comparison!"
            ElseIf Me.filterInputDecimal = "" Or Not IsNumeric(Me.filterInputDecimal) Then
                MsgBox "Enter a valid value!"
            Else
                filterString = Me.filterFieldSelector & vbTab & Me.filterInputComparison & vbTab & Me.filterInputDecimal
            End If
        Case Is = "Qty on Order", "Qty on Hand", "Number of Transactions", "Qty Sold"
            If Me.filterInputComparison.ListIndex = -1 Then
                MsgBox "Select a comparison!"
            ElseIf Me.filterInputInteger = "" Or Not IsIntegeric(Me.filterInputInteger) Then
                MsgBox "Enter a valid value!"
            Else
                filterString = Me.filterFieldSelector & vbTab & Me.filterInputComparison & vbTab & Me.filterInputInteger
            End If
        Case Is = "Reconditioned", "Discontinued", "Rebates/Specials"
            If Me.filterInputBoolean.ListIndex = -1 Then
                MsgBox "Select a value!"
            Else
                filterString = Me.filterFieldSelector & vbTab & Me.filterInputBoolean
            End If
        Case Is = "New Items Since"
            If DateDiff("d", Date, Me.filterInputDate.value) > 0 Then
                MsgBox "Date cannot be in the future!"
            Else
                filterString = Me.filterFieldSelector & vbTab & Me.filterInputDate.value
            End If
        Case Is = "Triad Code"
            filterString = Me.filterFieldSelector & vbTab & Me.filterInputText
        Case Else
            MsgBox "Tell Brian to add " & qq(Me.filterFieldSelector) & " to the filterBtnApply_Click() select clause!"
    End Select
    
    If filterString = "" Then
        'bad input, skip
    Else
        Me.filterFieldSelector.ListIndex = -1
        Me.filterInputBoolean.ListIndex = -1
        Me.filterInputComparison.ListIndex = -1
        Me.filterInputDate.value = Date
        Me.filterInputDecimal = ""
        Me.filterInputInteger = ""
        Me.filterInputText = ""
        Me.filterInputBoolean.Visible = False
        Me.filterInputComparison.Visible = False
        Me.filterInputDate.Visible = False
        Me.filterInputDecimal.Visible = False
        Me.filterInputInteger.Visible = False
        Me.filterInputText.Visible = False
        Me.filterBtnApply.Visible = False
        Me.criteriaFilter.AddItem filterString
        'it'd be nice if this refiltered the already filtered list, with
        'just this one additional filter criteria line, rather than
        'starting from the beginning
        refilterList
    End If
End Sub

Private Sub requeryFullItemList()
    If DateDiff("d", lastHistoryStart, lastHistoryEnd) < 0 Then
        MsgBox """From"" date is later than ""To"" date, fix this!"
        Exit Sub
    End If
    Mouse.Hourglass True
    Dim transactionTypeToUse As String, websiteToUse As Long
    If Me.btnWebsiteStoreSelection.Caption = "Store" Then
        transactionTypeToUse = "P"
        websiteToUse = 0 'TODO: make sproc website-independent
    Else
        transactionTypeToUse = "S"
        websiteToUse = Me.btnWebsiteSelection.tag
    End If
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("EXEC spProductsSalesTC '" & lastHistoryStart & "', '" & lastHistoryEnd & "', '', 0, 1, " & websiteToUse & ", '" & transactionTypeToUse & "'")  'all prodlines, all rows implied
    ReDim ItemList(rst.RecordCount - 1) As Dictionary
    If rst.RecordCount > 0 Then
        Dim i As Long
        Dim curr As Dictionary
        For i = 0 To rst.RecordCount - 1
            Set curr = New Dictionary
            Dim j As Long
            For j = 0 To rst.fields.count - 1
                curr.Add CStr(rst.fields(j).name), Trim(Nz(rst.fields(j).value, "0"))
            Next j
            Set ItemList(i) = curr
            rst.MoveNext
        Next i
        
        refilterList
    Else
        MsgBox "NO RECORDS! KIND OF FUCKED! FIX ME!"
    End If
    Mouse.Hourglass False
End Sub

Private Sub refilterList()
    ReDim FilteredItemList(UBound(ItemList)) As Dictionary
    Dim i As Long
    Dim j As Long
    For i = 0 To UBound(ItemList)
        If itemMatchesCriteria(i) Then
            If itemMatchesFragment(i) Then
                'add
                Set FilteredItemList(j) = ItemList(i)
                j = j + 1
            End If
        End If
    Next i
    
    If j = 0 Then
        MsgBox "No items match your filter criteria!"
    Else
        ReDim Preserve FilteredItemList(j - 1) As Dictionary
    
        QuickSortHashes FilteredItemList, Array(Array(currentsort, currentSortDesc))
        
        lastTopItemDx = 0
        lastSelItemDx = -1
        Me.itemScrollBar.Min = 0
        Me.itemScrollBar.max = IIf(UBound(FilteredItemList) - 1 - NUM_LINES < 0, 0, UBound(FilteredItemList) + 1 - NUM_LINES)
        If Me.itemScrollBar.value = 0 Then
            itemScrollBar_Change
        Else
            Me.itemScrollBar.value = 0
        End If
    End If
End Sub

Private Function itemMatchesCriteria(dx As Long) As Boolean
    Dim retval As Boolean
    retval = True
    Dim item As Dictionary
    Set item = ItemList(dx)
    Dim i As Long
    For i = 0 To Me.criteriaFilter.ListCount - 1
        Dim parts() As String
        parts = Split(Me.criteriaFilter.list(i), vbTab)
        Select Case parts(0)
            Case Is = "Internet Price"
                If Not compareCriteria(item.item("InternetPrice"), parts(1), parts(2)) Then
                    retval = False
                    Exit For
                End If
            Case Is = "Qty on Order"
                If Not compareCriteria(item.item("QtyOnOrder"), parts(1), parts(2)) Then
                    retval = False
                    Exit For
                End If
            Case Is = "Qty on Hand"
                If Not compareCriteria(item.item("QtyOnHand"), parts(1), parts(2)) Then
                    retval = False
                    Exit For
                End If
            Case Is = "Number of Transactions"
                If Not compareCriteria(item.item("NumTransactions"), parts(1), parts(2)) Then
                    retval = False
                    Exit For
                End If
            Case Is = "Qty Sold"
                If Not compareCriteria(item.item("QtySold"), parts(1), parts(2)) Then
                    retval = False
                    Exit For
                End If
            Case Is = "Reconditioned"
                If parts(1) = "True" And CBool(item.item("Reconditioned")) = False Then
                    retval = False
                    Exit For
                ElseIf parts(1) = "False" And CBool(item.item("Reconditioned")) = True Then
                    retval = False
                    Exit For
                End If
            Case Is = "Price Drops"
                If Not compareCriteria(-1 * CDbl(item.item("PriceChange")), parts(1), parts(2)) Then
                    retval = False
                    Exit For
                End If
            Case Is = "New Items Since"
                If IsNull(item.item("PublishedDate")) Then
                    retval = False
                    Exit For
                ElseIf DateDiff("d", item.item("PublishedDate"), parts(1)) > 0 Then
                    retval = False
                    Exit For
                End If
            Case Is = "Rebates/Specials"
                If parts(1) = "True" And CLng(item.item("ActiveSpecialCount")) = 0 Then
                    retval = False
                    Exit For
                ElseIf parts(1) = "False" And CLng(item.item("ActiveSpecialCount")) > 0 Then
                    retval = False
                    Exit For
                End If
            Case Is = "Discontinued"
                If parts(1) = "True" And CBool(item.item("Discontinued")) = False Then
                    retval = False
                    Exit For
                ElseIf parts(1) = "False" And CBool(item.item("Discontinued")) = True Then
                    retval = False
                    Exit For
                End If
            Case Is = "Triad Code"
                If parts(1) <> item.item("TriadCode") Then
                    retval = False
                    Exit For
                End If
            Case Is = "Standard Cost"
                If Not compareCriteria(item.item("StandardCost"), parts(1), parts(2)) Then
                    retval = False
                    Exit For
                End If
            Case Is = "Average Cost"
                If Not compareCriteria(item.item("AverageCost"), parts(1), parts(2)) Then
                    retval = False
                    Exit For
                End If
            Case Else
                MsgBox "Tell Brian to add " & qq(Me.filterFieldSelector) & " to the itemMatchesCriteria() select clause!"
        End Select
    Next i
    itemMatchesCriteria = retval
End Function

Private Function compareCriteria(value As String, comparisonString As String, threshold As String) As Boolean
    Dim val1 As Variant
    Dim val2 As Variant
    If IsNumeric(value) And IsNumeric(threshold) Then
        If IsIntegeric(value) And IsIntegeric(threshold) Then
            val1 = CLng(value)
            val2 = CLng(threshold)
        Else
            val1 = CDbl(value)
            val2 = CDbl(threshold)
        End If
    Else
        val1 = CStr(value)
        val2 = CStr(threshold)
    End If
    
    Select Case comparisonString
        Case Is = "=="
            compareCriteria = CBool(val1 = val2)
        Case Is = "!="
            compareCriteria = CBool(val1 <> val2)
        Case Is = ">"
            compareCriteria = CBool(val1 > val2)
        Case Is = "<"
            compareCriteria = CBool(val1 < val2)
        Case Is = ">="
            compareCriteria = CBool(val1 >= val2)
        Case Is = "<="
            compareCriteria = CBool(val1 <= val2)
        Case Else
            MsgBox "Tell Brian to add " & qq(comparisonString) & " to the compareCriteria() select clause!"
            compareCriteria = True
    End Select
End Function

Private Function itemMatchesFragment(dx As Long) As Boolean
    If lastQSearch = "" Then
        itemMatchesFragment = True
    Else
        Dim item As String
        item = LCase(ItemList(dx).item("ItemNumber"))
        Dim frag As String
        frag = LCase(IIf(Left(lastQSearch, 1) = "^", Mid(lastQSearch, 2), lastQSearch))
        If Left(lastQSearch, 1) = "^" Then
            itemMatchesFragment = CBool(Left(item, Len(frag)) = frag)
        Else
            itemMatchesFragment = CBool(InStr(item, frag))
        End If
    End If
End Function

Private Sub FragmentToSearch_Change()
    If Me.FragmentToSearch = lastQSearch Then
        'do nothing
    Else
        lastQSearch = Me.FragmentToSearch
        refilterList
    End If
    
    If Me.FragmentToSearch <> "" And Me.lblFragment.FontBold = False Then
        Me.lblFragment.FontBold = True
    ElseIf Me.FragmentToSearch = "" And Me.lblFragment.FontBold = True Then
        Me.lblFragment.FontBold = False
    End If
End Sub

Private Sub headerLabel_Click(Index As Integer)
    If Me.headerLabel(Index).tag <> "" Then
        If Me.headerLabel(Index).tag = currentsort Then
            'actually, we should just reverse the array
            currentSortDesc = Not currentSortDesc
        Else
            currentsort = Me.headerLabel(Index).tag
            currentSortDesc = False
            Dim i As Long
            For i = 0 To Me.headerLabel.UBound
                Me.headerLabel(i).FontBold = CBool(i = Index)
            Next i
        End If
        
        QuickSortHashes FilteredItemList, Array(Array(currentsort, currentSortDesc))
        
        lastTopItemDx = 0
        lastSelItemDx = -1
        If Me.itemScrollBar.value = 0 Then
            itemScrollBar_Change
        Else
            Me.itemScrollBar.value = 0
        End If
    End If
End Sub

Private Sub itemScrollBar_Change()
    Mouse.Hourglass True
    If UBound(FilteredItemList) >= 0 Then
        fillItemSubform Me.itemScrollBar.value
    End If
    lastTopItemDx = Me.itemScrollBar.value
    setFocusBackground lastSelItemDx
    Mouse.Hourglass False
End Sub

Private Sub addLine(arraydx As Long, guidx As Long)
    Me.ItemNumber(guidx) = FilteredItemList(arraydx).item("ItemNumber")
    Me.ItemDescription(guidx) = FilteredItemList(arraydx).item("ItemDescription")
    Me.QtySold(guidx) = FilteredItemList(arraydx).item("QtySold")
    Me.NumTransactions(guidx) = FilteredItemList(arraydx).item("NumTransactions")
    Me.QtyOnHand(guidx) = FilteredItemList(arraydx).item("QtyOnHand")
    Me.QtyOnOrder(guidx) = FilteredItemList(arraydx).item("QtyOnOrder")
    Me.InternetPrice(guidx) = FormatCurrency(FilteredItemList(arraydx).item("InternetPrice"), 2, vbTrue, vbFalse, vbFalse)
    Me.StandardCost(guidx) = FormatCurrency(FilteredItemList(arraydx).item("StandardCost"), 2, vbTrue, vbFalse, vbFalse)
    Me.AverageCost(guidx) = FormatCurrency(FilteredItemList(arraydx).item("AverageCost"), 2, vbTrue, vbFalse, vbFalse)
    If CBool(FilteredItemList(arraydx).item("Discontinued")) Then
        ForeColorsList(guidx) = RED
    Else
        ForeColorsList(guidx) = BLACK
    End If
End Sub

Private Sub fillItemSubform(Optional topDx As Long = 0)
    Dim i As Long
    For i = 0 To NUM_LINES - 1
        If topDx + i > UBound(FilteredItemList) Then
            Exit For
        End If
        addLine topDx + i, i
        Me.ItemNumber(i).Visible = True
        Me.ItemDescription(i).Visible = True
        Me.QtySold(i).Visible = True
        Me.NumTransactions(i).Visible = True
        Me.QtyOnHand(i).Visible = True
        Me.QtyOnOrder(i).Visible = True
        Me.InternetPrice(i).Visible = True
        Me.StandardCost(i).Visible = True
        Me.AverageCost(i).Visible = True
        Me.btnOpenWebsite(i).Visible = True
        Me.btnOpenInInventory(i).Visible = True
    Next i
    For i = i To NUM_LINES - 1
        Me.ItemNumber(i).Visible = False
        Me.ItemDescription(i).Visible = False
        Me.QtySold(i).Visible = False
        Me.NumTransactions(i).Visible = False
        Me.QtyOnHand(i).Visible = False
        Me.QtyOnOrder(i).Visible = False
        Me.InternetPrice(i).Visible = False
        Me.StandardCost(i).Visible = False
        Me.AverageCost(i).Visible = False
        Me.btnOpenWebsite(i).Visible = False
        Me.btnOpenInInventory(i).Visible = False
    Next i
End Sub

Private Sub setFocusBackground(dx As Long)
    Dim i As Long
    Dim guidx As Long
    guidx = dx - lastTopItemDx
    For i = 0 To Me.ItemFrame.UBound
        Me.ItemFrame(i).BackColor = IIf(guidx = i, RGB(10, 36, 106), &H8000000F)
        
        Me.ItemNumber(i).BackColor = IIf(guidx = i, RGB(10, 36, 106), vbWhite)
        Me.ItemDescription(i).BackColor = IIf(guidx = i, RGB(10, 36, 106), vbWhite)
        Me.QtySold(i).BackColor = IIf(guidx = i, RGB(10, 36, 106), vbWhite)
        Me.NumTransactions(i).BackColor = IIf(guidx = i, RGB(10, 36, 106), vbWhite)
        Me.QtyOnHand(i).BackColor = IIf(guidx = i, RGB(10, 36, 106), vbWhite)
        Me.QtyOnOrder(i).BackColor = IIf(guidx = i, RGB(10, 36, 106), vbWhite)
        Me.InternetPrice(i).BackColor = IIf(guidx = i, RGB(10, 36, 106), vbWhite)
        Me.StandardCost(i).BackColor = IIf(guidx = i, RGB(10, 36, 106), vbWhite)
        Me.AverageCost(i).BackColor = IIf(guidx = i, RGB(10, 36, 106), vbWhite)
        
        Me.ItemNumber(i).ForeColor = IIf(ForeColorsList(i) = RED, RED, IIf(guidx = i, vbWhite, vbBlack))
        Me.ItemDescription(i).ForeColor = IIf(guidx = i, vbWhite, vbBlack)
        Me.QtySold(i).ForeColor = IIf(guidx = i, vbWhite, vbBlack)
        Me.NumTransactions(i).ForeColor = IIf(guidx = i, vbWhite, vbBlack)
        Me.QtyOnHand(i).ForeColor = IIf(guidx = i, vbWhite, vbBlack)
        Me.QtyOnOrder(i).ForeColor = IIf(guidx = i, vbWhite, vbBlack)
        Me.InternetPrice(i).ForeColor = IIf(guidx = i, vbWhite, vbBlack)
        Me.StandardCost(i).ForeColor = IIf(guidx = i, vbWhite, vbBlack)
        Me.AverageCost(i).ForeColor = IIf(guidx = i, vbWhite, vbBlack)
    Next i
End Sub

Private Sub SalesFromDate_LostFocus()
    If Me.SalesFromDate.value <> lastHistoryStart Then
        lastHistoryStart = Me.SalesFromDate.value
        'requeryFullItemList
    End If
End Sub

Private Sub SalesToDate_LostFocus()
    If Me.SalesToDate.value <> lastHistoryEnd Then
        lastHistoryEnd = Me.SalesToDate.value
        'requeryFullItemList
    End If
End Sub

Private Sub btnChangeSalesTimeFrame_Click()
    requeryFullItemList
End Sub



'----------------------
Private Sub btnOpenWebsite_Click(Index As Integer)
    OpenDefaultApp "http://" & WebsiteURLHash.item(CLng(Me.btnWebsiteSelection.tag)) & "/" & LCase(CreateYahooID(Me.ItemNumber(Index))) & ".html"
End Sub

Private Sub btnOpenInInventory_Click(Index As Integer)
    Dim imLoaded As Boolean, smLoaded As Boolean
    imLoaded = IsFormLoaded("InventoryMaintenance")
    smLoaded = IsFormLoaded("SignMaintenance")
    If imLoaded Or smLoaded Then
        If imLoaded Then
            InventoryMaintenance.GoToItem Me.ItemNumber(Index), True
            FocusOnForm "InventoryMaintenance"
        ElseIf smLoaded Then
            SignMaintenance.GoToItem Me.ItemNumber(Index), True
            FocusOnForm "SignMaintenance"
        End If
    Else
        Load InventoryMaintenance
        InventoryMaintenance.GoToItem Me.ItemNumber(Index), True
        InventoryMaintenance.Show
    End If
End Sub

Private Sub ItemNumber_GotFocus(Index As Integer)
    lastSelItemDx = Index + lastTopItemDx
    setFocusBackground lastSelItemDx
End Sub

Private Sub ItemDescription_GotFocus(Index As Integer)
    lastSelItemDx = Index + lastTopItemDx
    setFocusBackground lastSelItemDx
End Sub

Private Sub QtySold_GotFocus(Index As Integer)
    lastSelItemDx = Index + lastTopItemDx
    setFocusBackground lastSelItemDx
End Sub

Private Sub NumTransactions_GotFocus(Index As Integer)
    lastSelItemDx = Index + lastTopItemDx
    setFocusBackground lastSelItemDx
End Sub

Private Sub QtyOnHand_GotFocus(Index As Integer)
    lastSelItemDx = Index + lastTopItemDx
    setFocusBackground lastSelItemDx
End Sub

Private Sub QtyOnOrder_GotFocus(Index As Integer)
    lastSelItemDx = Index + lastTopItemDx
    setFocusBackground lastSelItemDx
End Sub

Private Sub InternetPrice_GotFocus(Index As Integer)
    lastSelItemDx = Index + lastTopItemDx
    setFocusBackground lastSelItemDx
End Sub



'----------------------------

Private Sub btnSelectedListAdd_Click()
    If lastSelItemDx = -1 Then
        Exit Sub
    End If
    
    If Me.SelectedList.ListIndex = -1 Then
        MsgBox "Select a section to add this to first"
        Exit Sub
    End If
    
    Dim sectionDx As Long
    If listIndexIsSection(Me.SelectedList.ListIndex) Then
        sectionDx = Me.SelectedList.ListIndex
    Else
        sectionDx = getItemSectionListIndex(Me.SelectedList.ListIndex)
    End If
    
    Dim SectionName As String
    SectionName = Me.SelectedList.list(sectionDx)
    
    Dim tempA As Variant
    tempA = SelectionArray(getSectionArrayIndex(sectionDx))
    ReDim Preserve tempA(UBound(tempA) + 1) As Variant
    tempA(UBound(tempA)) = FilteredItemList(lastSelItemDx).item("ItemNumber")
    SelectionArray(getSectionArrayIndex(sectionDx)) = tempA
    
    requerySelectedList
    setSelectedListSelection SectionName, FilteredItemList(lastSelItemDx).item("ItemNumber")
End Sub

Private Sub btnSelectedListAddSection_Click()
    Dim newname As String
    'newName = InputBox("Enter new section name:")
    'If newName = "" Then
    '    Exit Sub
    'End If
    Load ProductSuggestionSectionDialog
    ProductSuggestionSectionDialog.Show MODAL
    If incomingSectionName = "" Then
        Exit Sub
    End If
    newname = incomingSectionName
    incomingSectionName = ""
    
    Dim i As Long
    For i = 0 To Me.SelectedList.ListCount - 1
        If LCase(Me.SelectedList.list(i)) = LCase(newname) Then
            MsgBox "There is already a section with this name!"
            Exit Sub
        End If
    Next i
    
    On Error GoTo errh
    ReDim Preserve SelectionArray(UBound(SelectionArray) + 1) As Variant
    SelectionArray(UBound(SelectionArray)) = Array(newname)
    Me.SelectedList.AddItem newname
    setSelectedListSelection newname
    Exit Sub
errh:
    If Err.Number = 9 Then
        Err.Clear
        ReDim SelectionArray(0) As Variant
        Resume Next
    Else
        MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    End If
End Sub

Private Sub btnSelectedListMoveDown_Click()
    Dim sectionDx As Long, nextSectionDx As Long, i As Long
    Dim SectionName As String, itemName As String
    
    If listIndexIsSection(Me.SelectedList.ListIndex) Then
        'move section
        sectionDx = Me.SelectedList.ListIndex
        
        nextSectionDx = -999
        For i = sectionDx + 1 To Me.SelectedList.ListCount - 1
            If listIndexIsSection(i) Then
                nextSectionDx = i
                Exit For
            End If
        Next i
        If nextSectionDx = -999 Then
            'already last
            Exit Sub
        End If
        
        SectionName = Me.SelectedList
        
        swapSections getSectionArrayIndex(sectionDx), getSectionArrayIndex(nextSectionDx)
        requerySelectedList
        setSelectedListSelection SectionName
    Else
        'move item
        If Me.SelectedList.ListIndex = Me.SelectedList.ListCount - 1 Then
            'last, do not move
        ElseIf listIndexIsSection(Me.SelectedList.ListIndex + 1) Then
            'next is section, do not move
        Else
            sectionDx = getItemSectionListIndex(Me.SelectedList.ListIndex)
            SectionName = Me.SelectedList.list(sectionDx)
            itemName = Me.SelectedList
            swapItems getSectionArrayIndex(sectionDx), getItemArrayIndex(Me.SelectedList.ListIndex), getItemArrayIndex(Me.SelectedList.ListIndex + 1)
            requerySelectedList
            setSelectedListSelection SectionName, itemName
        End If
    End If
End Sub

Private Sub btnSelectedListMoveUp_Click()
    If Me.SelectedList.ListIndex <= 0 Then
        Exit Sub
    End If
    
    Dim sectionDx As Long, prevSectionDx As Long, i As Long
    Dim SectionName As String, itemName As String
    
    If listIndexIsSection(Me.SelectedList.ListIndex) Then
        'move entire section
        sectionDx = Me.SelectedList.ListIndex
        
        For i = sectionDx - 1 To 0 Step -1
            If listIndexIsSection(i) Then
                prevSectionDx = i
                Exit For
            End If
        Next i
        
        SectionName = Me.SelectedList
        
        swapSections getSectionArrayIndex(sectionDx), getSectionArrayIndex(prevSectionDx)
        requerySelectedList
        setSelectedListSelection SectionName
    Else
        'move single item
        If listIndexIsSection(Me.SelectedList.ListIndex - 1) Then
            'do not move
        Else
            sectionDx = getItemSectionListIndex(Me.SelectedList.ListIndex)
            SectionName = Me.SelectedList.list(sectionDx)
            itemName = Me.SelectedList
            swapItems getSectionArrayIndex(sectionDx), getItemArrayIndex(Me.SelectedList.ListIndex), getItemArrayIndex(Me.SelectedList.ListIndex - 1)
            requerySelectedList
            setSelectedListSelection SectionName, itemName
        End If
    End If
End Sub

Private Function listIndexIsSection(dx) As Boolean
    listIndexIsSection = Not CBool(Left(Me.SelectedList.list(dx), 4) = "    ")
End Function

Private Function getSectionArrayIndex(listdx As Long) As Long
    Dim i As Long
    Dim counter As Long
    counter = 0
    For i = 0 To listdx
        If listIndexIsSection(i) Then
            counter = counter + 1
        End If
    Next i
    getSectionArrayIndex = counter - 1
End Function

Private Function getItemSectionListIndex(itemListDx As Long) As Long
    Dim i As Long
    Dim dx As Long
    For i = itemListDx To 0 Step -1
        If listIndexIsSection(i) Then
            dx = i
            Exit For
        End If
    Next i
    getItemSectionListIndex = dx
End Function

Private Function getItemArrayIndex(listdx As Long) As Long
    Dim i As Long
    Dim counter As Long
    counter = 0
    For i = listdx To 0 Step -1
        If listIndexIsSection(i) Then
            Exit For
        Else
            counter = counter + 1
        End If
    Next i
    getItemArrayIndex = counter 'this doesn't get a -1, because the section name is 1st
End Function

Private Sub swapSections(arrayDx1 As Long, arrayDx2 As Long)
    Dim temp As Variant
    temp = SelectionArray(arrayDx1)
    SelectionArray(arrayDx1) = SelectionArray(arrayDx2)
    SelectionArray(arrayDx2) = temp
End Sub

Private Sub swapItems(sectionArrayDx As Long, itemArrayDx1 As Long, itemArrayDx2 As Long)
    Dim temp As Variant
    temp = SelectionArray(sectionArrayDx)(itemArrayDx1)
    SelectionArray(sectionArrayDx)(itemArrayDx1) = SelectionArray(sectionArrayDx)(itemArrayDx2)
    SelectionArray(sectionArrayDx)(itemArrayDx2) = temp
End Sub

Private Sub btnSelectedListRemove_Click()
    If Me.SelectedList.ListIndex = -1 Then
        Exit Sub
    End If
    
    Dim i As Long, tempA As Variant, sectionArrayDx As Long, stopNow As Boolean
    
    If listIndexIsSection(Me.SelectedList.ListIndex) Then
        sectionArrayDx = getSectionArrayIndex(Me.SelectedList.ListIndex)
        
        If UBound(SelectionArray(sectionArrayDx)) > 0 Then
            If vbNo = MsgBox("Are you sure you want to remove this section? It will remove " & UBound(SelectionArray(sectionArrayDx)) & " item(s)!", vbYesNo) Then
                stopNow = True
            End If
        End If
        
        If Not stopNow Then
            For i = sectionArrayDx + 1 To UBound(SelectionArray)
                SelectionArray(i - 1) = SelectionArray(i)
            Next i
            ReDim Preserve SelectionArray(UBound(SelectionArray) - 1) As Variant
            i = Me.SelectedList.ListIndex
            Me.SelectedList.RemoveItem i
            While Not listIndexIsSection(i) And i <= Me.SelectedList.ListCount - 1
                Me.SelectedList.RemoveItem i
            Wend
        End If
    Else
        sectionArrayDx = getSectionArrayIndex(getItemSectionListIndex(Me.SelectedList.ListIndex))
        tempA = SelectionArray(sectionArrayDx)
        i = getItemArrayIndex(Me.SelectedList.ListIndex)
        tempA(i) = ""
        For i = i + 1 To UBound(tempA)
            tempA(i - 1) = tempA(i)
        Next i
        ReDim Preserve tempA(UBound(tempA) - 1) As Variant
        SelectionArray(sectionArrayDx) = tempA
        Me.SelectedList.RemoveItem Me.SelectedList.ListIndex
    End If
End Sub


'------------------------
Private Sub requeryGenerationSelector()
    Me.GenerationSelector.Clear
    Me.GenerationSelector.AddItem "Front Page"
    Me.GenerationSelector.AddItem "Newsletter"
    ExpandDropDownToFit Me.GenerationSelector
End Sub

Private Sub GenerationSelector_Click()
    If Me.SelectedList.ListCount = 0 Then
        MsgBox "There's nothing in the list!"
        Exit Sub
    End If
    
    Select Case Me.GenerationSelector
        Case Is = "Front Page"
            If generation_FrontPageSanityCheck Then
                generation_FrontPage
            End If
        Case Is = "Newsletter"
            If generation_NewsletterSanityCheck Then
                generation_Newsletter
            End If
        Case Else
            MsgBox "Brian needs to add " & qq(Me.GenerationSelector) & " to the GenerationSelector click event!"
    End Select
End Sub

Private Function generation_FrontPageSanityCheck() As Boolean
    Dim Ret As String
    
    Ret = generation_checkEmptySections()
    If Ret <> "" Then
        MsgBox "Section " & qq(Ret) & " is empty, you should remove it"
        generation_FrontPageSanityCheck = False
        Exit Function
    End If
    
    Ret = generation_checkItemQuantityModulo(4)
    If Ret <> "" Then
        MsgBox "Section " & qq(Ret) & " should have a multiple of 4 for the item count"
        generation_FrontPageSanityCheck = False
        Exit Function
    End If
    
    If generation_checkTotalNumberOfSections() > 5 Then
        MsgBox "The front page can only have up to 5 sections"
        generation_FrontPageSanityCheck = False
        Exit Function
    End If
    
    Ret = generation_checkItemDuplicates()
    If Ret <> "" Then
        MsgBox "Item " & qq(Ret) & " appears more than once in the list"
        generation_FrontPageSanityCheck = False
        Exit Function
    End If
    
    generation_FrontPageSanityCheck = True
End Function

Private Sub generation_FrontPage()
    Dim i As Long, j As Long
    Dim line As String, everything As String
    
    For i = 0 To UBound(SelectionArray)
        everything = everything & SelectionArray(i)(0) & IIf(IsAlphaNumeric(Right(SelectionArray(i)(0), 1)), ":", "") & vbCrLf
        line = ""
        For j = 1 To UBound(SelectionArray(i))
            line = IIf(line = "", "", line & " ") & CreateYahooID(CStr(SelectionArray(i)(j)))
        Next j
        everything = everything & line & vbCrLf & vbCrLf & vbCrLf
    Next i
    
    Open DESTINATION_DIR & "fp.txt" For Output As #1
    Print #1, everything
    Close #1
    
    OpenDefaultApp DESTINATION_DIR & "fp.txt"
End Sub

Private Function generation_NewsletterSanityCheck() As Boolean
    Dim Ret As String
    
    Ret = generation_checkEmptySections()
    If Ret <> "" Then
        MsgBox "Section " & qq(Ret) & " is empty, you should remove it"
        generation_NewsletterSanityCheck = False
        Exit Function
    End If
    
    Ret = generation_checkItemDuplicates()
    If Ret <> "" Then
        MsgBox "Item " & qq(Ret) & " appears more than once in the list"
        generation_NewsletterSanityCheck = False
        Exit Function
    End If
    
    generation_NewsletterSanityCheck = True
End Function

Private Sub generation_Newsletter()
    Dim i As Long, j As Long
    Dim line As String, everything As String
    
    For i = 0 To UBound(SelectionArray)
        'everything = everything & SelectionArray(i)(0) & IIf(IsAlphaNumeric(Right(SelectionArray(i)(0), 1)), ":", "") & vbCrLf
        line = ""
        For j = 1 To UBound(SelectionArray(i))
            'line = IIf(line = "", "", line & " ") & CreateYahooID(CStr(SelectionArray(i)(j)))
            Dim yid As String
            yid = CreateYahooID(CStr(SelectionArray(i)(j)))
            Dim itemName As String
            itemName = DLookup("ManufFullNameWeb", "ProductLine", "ProductLine='" & Left(CStr(SelectionArray(i)(j)), 3) & "'") & " " & DLookup("Desc2", "PartNumbers", "ItemNumber='" & CStr(SelectionArray(i)(j)) & "'")
            Dim itemprice As String
            itemprice = Format(DLookup("IDiscountMarkupPriceRate1", "InventoryMaster", "ItemNumber='" & CStr(SelectionArray(i)(j)) & "'"), "Currency")
            line = line & "<td style=""text-align: center; border-style:solid; border-width:1px; border-color:#000000;"" width=""140"">" & vbCrLf
            line = line & vbTab & "<a href=""" & yid & ".html""><img src=""" & yid & ".jpg"" width=""130"" height=""130"" border=""0"" alt=""""><br>" & vbCrLf
            line = line & vbTab & itemName & "<br>" & vbCrLf
            line = line & vbTab & DLookup("Desc3", "PartNumbers", "ItemNumber='" & CStr(SelectionArray(i)(j)) & "'") & "<br>" & vbCrLf
            line = line & vbTab & "<b style=""color:blue; font-family: Arial,sans-serif; font-weight:bold; font-size:14px;"">Price:</b> <b style=""color:red; font-family: Arial,sans-serif; font-weight:bold; font-size:14px;"">" & itemprice & "</b>" & vbCrLf
            line = line & "</td>" & vbCrLf
            line = line & vbCrLf & vbCrLf
        Next j
        everything = everything & line & vbCrLf & vbCrLf & vbCrLf
    Next i
    
    Open DESTINATION_DIR & "nl.txt" For Output As #1
    Print #1, everything
    Close #1
    
    OpenDefaultApp DESTINATION_DIR & "nl.txt"
End Sub

Private Function generation_checkEmptySections() As String
    Dim i As Long
    For i = 0 To UBound(SelectionArray)
        If UBound(SelectionArray(i)) = 0 Then
            generation_checkEmptySections = SelectionArray(i)(0)
            Exit Function
        End If
    Next i
    generation_checkEmptySections = ""
End Function

Private Function generation_checkItemQuantityModulo(num As Long) As String
    Dim i As Long
    For i = 0 To UBound(SelectionArray)
        If UBound(SelectionArray(i)) Mod num <> 0 Then
            generation_checkItemQuantityModulo = SelectionArray(i)(0)
            Exit Function
        End If
    Next i
    generation_checkItemQuantityModulo = ""
End Function

Private Function generation_checkTotalNumberOfSections() As Long
    generation_checkTotalNumberOfSections = UBound(SelectionArray) + 1
End Function

Private Function generation_checkItemDuplicates(Optional sectionOnly As Boolean = False) As String
    Dim seen As Dictionary
    Set seen = New Dictionary
    Dim i As Long, j As Long
    For i = 0 To UBound(SelectionArray)
        If sectionOnly Then
            Set seen = New Dictionary
        End If
        For j = 1 To UBound(SelectionArray(i))
            If seen.exists(CStr(SelectionArray(i)(j))) Then
                generation_checkItemDuplicates = CStr(SelectionArray(i)(j))
                Exit Function
            End If
            seen.Add CStr(SelectionArray(i)(j)), 1
        Next j
    Next i
    generation_checkItemDuplicates = ""
End Function
