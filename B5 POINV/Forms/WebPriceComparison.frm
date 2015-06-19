VERSION 5.00
Begin VB.Form WebPriceComparison 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Run Web Price Comparison"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4125
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkHideDOSBox 
      Caption         =   "Hide DOS box?"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2310
      Width           =   1635
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   2280
      TabIndex        =   8
      Top             =   2790
      Width           =   1275
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "Start WPC"
      Height          =   435
      Left            =   870
      TabIndex        =   7
      Top             =   2790
      Width           =   1275
   End
   Begin VB.ComboBox LCTo 
      Height          =   315
      Left            =   1350
      TabIndex        =   3
      Top             =   1860
      Width           =   1365
   End
   Begin VB.ComboBox LCFrom 
      Height          =   315
      Left            =   1350
      TabIndex        =   2
      Top             =   1500
      Width           =   1365
   End
   Begin VB.ComboBox ItemNumber 
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   870
      Width           =   2475
   End
   Begin VB.Label Label1 
      Caption         =   "by path?"
      Height          =   465
      Left            =   3120
      TabIndex        =   10
      Top             =   1410
      Width           =   885
   End
   Begin VB.Label lblStatusBar 
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3300
      Width           =   4065
   End
   Begin VB.Label generalLabel 
      Caption         =   "Run Web Price Comparison"
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
      Left            =   120
      TabIndex        =   6
      Top             =   150
      Width           =   3615
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "To Line Code:"
      Height          =   255
      Index           =   2
      Left            =   30
      TabIndex        =   5
      Top             =   1890
      Width           =   1275
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "From Line Code:"
      Height          =   255
      Index           =   1
      Left            =   30
      TabIndex        =   4
      Top             =   1530
      Width           =   1275
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Single Item:"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   900
      Width           =   1035
   End
End
Attribute VB_Name = "WebPriceComparison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    If Me.ItemNumber <> "" Then
        Me.lblStatusBar.caption = "Processing " & Me.ItemNumber
        PriceCompareSingleItemOrLC Me.ItemNumber, True, IIf(Me.chkHideDOSBox, True, False)
        Me.lblStatusBar.caption = ""
    ElseIf Me.LCFrom <> "" And Me.LCTo <> "" Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT ProductLine FROM ProductLine WHERE Hide=0 AND ProductLine BETWEEN '" & Me.LCFrom & "' AND '" & Me.LCTo & "' ORDER BY ProductLine")
        While Not rst.EOF
            Me.lblStatusBar.caption = "Processing " & rst("ProductLine")
            DoEvents
            PriceCompareSingleItemOrLC rst("ProductLine"), True, IIf(Me.chkHideDOSBox, True, False)
            rst.MoveNext
        Wend
        Me.lblStatusBar.caption = ""
    Else
        MsgBox "Enter an itemnumber, or from+to line codes."
    End If
End Sub

Private Sub Form_Load()
    requeryItems
    requeryLines
End Sub

Private Sub requeryItems()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber FROM InventoryMaster")
    Me.ItemNumber.Clear
    While Not rst.EOF
        Me.ItemNumber.AddItem (rst("ItemNumber"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
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

Private Sub ItemNumber_Click()
    ItemNumber_LostFocus
End Sub

Private Sub ItemNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.ItemNumber, KeyCode, Shift
End Sub

Private Sub ItemNumber_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.ItemNumber, KeyAscii
    If KeyAscii = vbKeyReturn Then
        ItemNumber_LostFocus
    End If
End Sub

Private Sub ItemNumber_LostFocus()
    AutoCompleteLostFocus Me.ItemNumber
    Me.LCFrom = ""
    Me.LCTo = ""
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
    Me.ItemNumber = ""
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
    Me.ItemNumber = ""
End Sub
