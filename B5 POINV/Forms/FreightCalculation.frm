VERSION 5.00
Begin VB.Form FreightCalculation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Run Freight Calculation"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   3780
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkDoAll 
      Caption         =   "Refresh Existing"
      Height          =   255
      Left            =   1980
      TabIndex        =   11
      Top             =   2100
      Width           =   1695
   End
   Begin VB.CheckBox chkHideDOSBox 
      Caption         =   "Hide DOS box?"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2100
      Width           =   1635
   End
   Begin VB.TextBox ZipCode 
      Height          =   315
      Left            =   1500
      TabIndex        =   9
      Top             =   1650
      Width           =   1365
   End
   Begin VB.ComboBox LCFrom 
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Top             =   750
      Width           =   1365
   End
   Begin VB.ComboBox LCTo 
      Height          =   315
      Left            =   1500
      TabIndex        =   2
      Top             =   1110
      Width           =   1365
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "Start"
      Height          =   435
      Left            =   480
      TabIndex        =   1
      Top             =   2490
      Width           =   1275
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   435
      Left            =   1890
      TabIndex        =   0
      Top             =   2490
      Width           =   1275
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip Code:"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   8
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label lblStatusBar 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2970
      Width           =   3705
   End
   Begin VB.Label generalLabel 
      Caption         =   "Run Freight Calculation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   3615
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "From Line Code:"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   780
      Width           =   1275
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "To Line Code:"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   1140
      Width           =   1275
   End
End
Attribute VB_Name = "FreightCalculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    If Me.LCFrom <> "" And Me.LCTo <> "" And Me.ZipCode <> "" Then
        Dim tempfile As String
        Randomize (Timer)
        tempfile = CInt(Rnd * 10000) & ".txt"
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("spLineCodesRange '" & Me.LCFrom & "', '" & Me.LCTo & "'")
        'do each line code individually, so it's not completely locked out
        While Not rst.EOF
            Me.lblStatusBar.caption = "Processing " & rst("ProductLine")
            DoEvents
            FreightEstimateMultiItem rst("ProductLine"), rst("ProductLine"), IIf(Me.chkDoAll, "refresh", "skip"), Me.ZipCode, IIf(Me.chkHideDOSBox, True, False)
            rst.MoveNext
        Wend
        Me.lblStatusBar.caption = ""
    Else
        MsgBox "Enter from+to line codes and zip code."
    End If
End Sub

Private Sub Form_Load()
    requeryLines
End Sub

Private Sub requeryLines()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM vRequeryLineCodes ORDER BY ProductLine")
    Me.LCFrom.Clear
    Me.LCTo.Clear
    While Not rst.EOF
        Me.LCFrom.AddItem (rst("ProductLine"))
        Me.LCTo.AddItem (rst("ProductLine"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub LCFrom_Click()
    LCFrom_LostFocus
End Sub

Private Sub LCFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.LCFrom, KeyCode, Shift
End Sub

Private Sub LCFrom_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.LCFrom, KeyAscii
    If KeyAscii = vbKeyReturn Then
        LCFrom_LostFocus
    End If
End Sub

Private Sub LCFrom_LostFocus()
    AutoCompleteLostFocus Me.LCFrom
    Me.LCTo = Me.LCFrom
End Sub

Private Sub LCTo_Click()
    LCTo_LostFocus
End Sub

Private Sub LCTo_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.LCTo, KeyCode, Shift
End Sub

Private Sub LCTo_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.LCTo, KeyAscii
    If KeyAscii = vbKeyReturn Then
        LCTo_LostFocus
    End If
End Sub

Private Sub LCTo_LostFocus()
    AutoCompleteLostFocus Me.LCTo
End Sub

