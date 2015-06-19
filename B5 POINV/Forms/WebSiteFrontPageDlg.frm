VERSION 5.00
Begin VB.Form WebSiteFrontPageDlg 
   Caption         =   "Web Site Front Page Maintenance"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnItems 
      Caption         =   "Col5-items"
      Height          =   315
      Index           =   4
      Left            =   90
      TabIndex        =   22
      Top             =   5670
      Width           =   1275
   End
   Begin VB.TextBox items 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   1410
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   5670
      Width           =   5535
   End
   Begin VB.CommandButton btnHeading 
      Caption         =   "Col5-heading"
      Height          =   315
      Index           =   4
      Left            =   90
      TabIndex        =   20
      Top             =   2280
      Width           =   1275
   End
   Begin VB.TextBox heading 
      Height          =   285
      Index           =   4
      Left            =   1410
      TabIndex        =   19
      Top             =   2280
      Width           =   5535
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save These"
      Height          =   435
      Left            =   1230
      TabIndex        =   18
      Top             =   6570
      Width           =   1725
   End
   Begin VB.TextBox items 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   1410
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   4980
      Width           =   5535
   End
   Begin VB.TextBox items 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1410
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   4290
      Width           =   5535
   End
   Begin VB.TextBox items 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1410
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   3600
      Width           =   5535
   End
   Begin VB.TextBox items 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   1410
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   2910
      Width           =   5535
   End
   Begin VB.CommandButton btnItems 
      Caption         =   "Col4-items"
      Height          =   315
      Index           =   3
      Left            =   90
      TabIndex        =   13
      Top             =   4980
      Width           =   1275
   End
   Begin VB.CommandButton btnItems 
      Caption         =   "Col3-items"
      Height          =   315
      Index           =   2
      Left            =   90
      TabIndex        =   12
      Top             =   4290
      Width           =   1275
   End
   Begin VB.CommandButton btnItems 
      Caption         =   "Col2-items"
      Height          =   315
      Index           =   1
      Left            =   90
      TabIndex        =   11
      Top             =   3600
      Width           =   1275
   End
   Begin VB.CommandButton btnItems 
      Caption         =   "Col1-items"
      Height          =   315
      Index           =   0
      Left            =   90
      TabIndex        =   10
      Top             =   2910
      Width           =   1275
   End
   Begin VB.TextBox heading 
      Height          =   285
      Index           =   3
      Left            =   1410
      TabIndex        =   9
      Top             =   1920
      Width           =   5535
   End
   Begin VB.CommandButton btnHeading 
      Caption         =   "Col4-heading"
      Height          =   315
      Index           =   3
      Left            =   90
      TabIndex        =   8
      Top             =   1920
      Width           =   1275
   End
   Begin VB.TextBox heading 
      Height          =   285
      Index           =   2
      Left            =   1410
      TabIndex        =   7
      Top             =   1560
      Width           =   5535
   End
   Begin VB.CommandButton btnHeading 
      Caption         =   "Col3-heading"
      Height          =   315
      Index           =   2
      Left            =   90
      TabIndex        =   6
      Top             =   1560
      Width           =   1275
   End
   Begin VB.TextBox heading 
      Height          =   285
      Index           =   1
      Left            =   1410
      TabIndex        =   5
      Top             =   1200
      Width           =   5535
   End
   Begin VB.CommandButton btnHeading 
      Caption         =   "Col2-heading"
      Height          =   315
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Top             =   1200
      Width           =   1275
   End
   Begin VB.TextBox heading 
      Height          =   285
      Index           =   0
      Left            =   1410
      TabIndex        =   3
      Top             =   840
      Width           =   5535
   End
   Begin VB.CommandButton btnHeading 
      Caption         =   "Col1-heading"
      Height          =   315
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   840
      Width           =   1275
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "OK"
      Height          =   435
      Left            =   3270
      TabIndex        =   1
      Top             =   6570
      Width           =   1725
   End
   Begin VB.Label generalLabel 
      Caption         =   "Web Site Front Page Maintenance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   4875
   End
End
Attribute VB_Name = "WebSiteFrontPageDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload WebSiteFrontPageDlg
End Sub

Private Sub btnHeading_Click(Index As Integer)
    ClipBoard_SetData Me.heading(Index)
    Me.heading(Index).SelStart = 0
    Me.heading(Index).SelLength = Len(Me.heading(Index))
    Me.heading(Index).SetFocus
End Sub

Private Sub btnItems_Click(Index As Integer)
    ClipBoard_SetData Me.items(Index)
    Me.items(Index).SelStart = 0
    Me.items(Index).SelLength = Len(Me.items(Index))
    Me.items(Index).SetFocus
End Sub

Private Sub btnSave_Click()
    Dim i As Long, j As Long, pcs As Variant
    DB.Execute "UPDATE WebFrontPage SET CurrentPos=-1"
    DB.Execute "TRUNCATE TABLE WebFrontPageItems"
    For i = Me.heading.LBound To Me.heading.UBound
        DB.Execute "UPDATE WebFrontPage SET CurrentPos=" & i & " WHERE ColName='" & EscapeSQuotes(Left(Me.heading(i), Len(Me.heading(i)) - 1)) & "'"
        pcs = Split(Me.items(i).tag, " ")
        For j = LBound(pcs) To UBound(pcs)
            DB.Execute "INSERT INTO WebFrontPageItems ( WhichCol, WhichPos, ItemNumber ) VALUES ( " & i & ", " & j & ", '" & EscapeSQuotes(UCase(pcs(j))) & "' )"
        Next j
    Next i
End Sub

