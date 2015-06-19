VERSION 5.00
Begin VB.Form ReportsMenu 
   Caption         =   "Select A Report"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnPrint 
      Caption         =   "Print Report"
      Height          =   435
      Left            =   4650
      TabIndex        =   61
      Top             =   4740
      Width           =   1515
   End
   Begin VB.CommandButton btnPreview 
      Caption         =   "Preview Report"
      Height          =   435
      Left            =   3120
      TabIndex        =   60
      Top             =   4740
      Width           =   1515
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   405
      Left            =   6570
      TabIndex        =   16
      Top             =   4740
      Width           =   1185
   End
   Begin VB.Frame frameOptions 
      Caption         =   "Select Options"
      Height          =   3915
      Left            =   2610
      TabIndex        =   1
      Top             =   480
      Width           =   5325
      Begin VB.TextBox gmPct2 
         Height          =   255
         Left            =   4470
         TabIndex        =   55
         Top             =   3450
         Width           =   705
      End
      Begin VB.TextBox gmPct1 
         Height          =   255
         Left            =   3360
         TabIndex        =   53
         Top             =   3450
         Width           =   525
      End
      Begin VB.TextBox pctOverOP 
         Height          =   255
         Left            =   4650
         TabIndex        =   48
         Top             =   2970
         Width           =   495
      End
      Begin VB.TextBox pctDiscount 
         Height          =   255
         Left            =   4650
         TabIndex        =   44
         Top             =   2670
         Width           =   495
      End
      Begin VB.CheckBox chkRankD 
         Height          =   225
         Left            =   2940
         TabIndex        =   42
         Top             =   3000
         Width           =   225
      End
      Begin VB.CheckBox chkRankC 
         Height          =   225
         Left            =   2280
         TabIndex        =   41
         Top             =   3000
         Width           =   225
      End
      Begin VB.CheckBox chkRankB 
         Height          =   225
         Left            =   1650
         TabIndex        =   40
         Top             =   3000
         Width           =   225
      End
      Begin VB.CheckBox chkRankA 
         Height          =   225
         Left            =   990
         TabIndex        =   38
         Top             =   3000
         Width           =   225
      End
      Begin VB.TextBox RankD 
         Height          =   255
         Left            =   2820
         TabIndex        =   33
         Top             =   2730
         Width           =   405
      End
      Begin VB.TextBox RankC 
         Height          =   255
         Left            =   2190
         TabIndex        =   32
         Top             =   2730
         Width           =   405
      End
      Begin VB.TextBox RankB 
         Height          =   255
         Left            =   1530
         TabIndex        =   31
         Top             =   2730
         Width           =   405
      End
      Begin VB.TextBox RankA 
         Height          =   255
         Left            =   870
         TabIndex        =   30
         Top             =   2730
         Width           =   405
      End
      Begin VB.CheckBox useYTD 
         Caption         =   "Use YTD"
         Height          =   255
         Left            =   3900
         TabIndex        =   26
         Top             =   2310
         Width           =   1065
      End
      Begin VB.CheckBox useLastPeriod 
         Caption         =   "Use Last Period"
         Height          =   255
         Left            =   2100
         TabIndex        =   25
         Top             =   2310
         Width           =   1515
      End
      Begin VB.CheckBox useCurrentPeriod 
         Caption         =   "Use Current Period"
         Height          =   255
         Left            =   210
         TabIndex        =   24
         Top             =   2310
         Width           =   1665
      End
      Begin VB.CheckBox onlyNegQty 
         Caption         =   "Only QOH <= 0"
         Height          =   255
         Left            =   2700
         TabIndex        =   23
         Top             =   1980
         Width           =   1515
      End
      Begin VB.CheckBox onlyDropShip 
         Caption         =   "Only DropShip"
         Height          =   255
         Left            =   540
         TabIndex        =   22
         Top             =   3450
         Width           =   1335
      End
      Begin VB.CheckBox notBeingDC 
         Caption         =   "Not Being D/C"
         Height          =   255
         Left            =   2700
         TabIndex        =   21
         Top             =   1710
         Width           =   1965
      End
      Begin VB.CheckBox onlyBeingDC 
         Caption         =   "Only Being D/C"
         Height          =   255
         Left            =   510
         TabIndex        =   20
         Top             =   1710
         Width           =   1485
      End
      Begin VB.CheckBox onlyPosQty 
         Caption         =   "Only QOH > 0"
         Height          =   255
         Left            =   510
         TabIndex        =   19
         Top             =   1980
         Width           =   1335
      End
      Begin VB.CheckBox notDC 
         Caption         =   "Not D/C"
         Height          =   255
         Left            =   2700
         TabIndex        =   18
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox onlyDC 
         Caption         =   "Only D/C"
         Height          =   255
         Left            =   510
         TabIndex        =   17
         Top             =   1440
         Width           =   1065
      End
      Begin VB.CheckBox chkWhseAll 
         Height          =   255
         Left            =   4680
         TabIndex        =   15
         Top             =   900
         Width           =   255
      End
      Begin VB.ComboBox cmbWhse 
         Height          =   315
         Left            =   2100
         TabIndex        =   14
         Top             =   900
         Width           =   2265
      End
      Begin VB.CheckBox chkLineAll 
         Height          =   255
         Left            =   4680
         TabIndex        =   8
         Top             =   540
         Width           =   255
      End
      Begin VB.ComboBox cmbLineEnd 
         Height          =   315
         Left            =   3330
         TabIndex        =   7
         Top             =   510
         Width           =   1035
      End
      Begin VB.ComboBox cmbLineStart 
         Height          =   315
         Left            =   2100
         TabIndex        =   6
         Top             =   510
         Width           =   1035
      End
      Begin VB.Label lblAnd 
         Caption         =   "and %:"
         Height          =   225
         Left            =   3960
         TabIndex        =   54
         Top             =   3480
         Width           =   525
      End
      Begin VB.Label lblBetween 
         Caption         =   "Between %:"
         Height          =   225
         Left            =   2430
         TabIndex        =   52
         Top             =   3480
         Width           =   915
      End
      Begin VB.Label lblPctOverOP 
         Caption         =   "% Over OP:"
         Height          =   225
         Left            =   3780
         TabIndex        =   47
         Top             =   3000
         Width           =   825
      End
      Begin VB.Label lblPctDiscount 
         Caption         =   "% Discount:"
         Height          =   225
         Left            =   3750
         TabIndex        =   45
         Top             =   2700
         Width           =   885
      End
      Begin VB.Label lblOnlyUse 
         Caption         =   "Only Use:"
         Height          =   255
         Left            =   180
         TabIndex        =   39
         Top             =   2970
         Width           =   735
      End
      Begin VB.Label lblD 
         Caption         =   "D:"
         Height          =   225
         Left            =   2640
         TabIndex        =   37
         Top             =   2730
         Width           =   165
      End
      Begin VB.Label lblC 
         Caption         =   "C:"
         Height          =   225
         Left            =   1980
         TabIndex        =   36
         Top             =   2730
         Width           =   165
      End
      Begin VB.Label lblB 
         Caption         =   "B:"
         Height          =   225
         Left            =   1320
         TabIndex        =   35
         Top             =   2730
         Width           =   195
      End
      Begin VB.Label lblA 
         Caption         =   "A:"
         Height          =   225
         Left            =   660
         TabIndex        =   34
         Top             =   2730
         Width           =   195
      End
      Begin VB.Label lblWhse 
         Caption         =   "Warehouse:"
         Height          =   255
         Left            =   1140
         TabIndex        =   13
         Top             =   930
         Width           =   885
      End
      Begin VB.Label lblLC 
         Caption         =   "Line Code:"
         Height          =   255
         Left            =   1140
         TabIndex        =   12
         Top             =   540
         Width           =   885
      End
      Begin VB.Label lblStart 
         Caption         =   "Start"
         Height          =   255
         Left            =   2310
         TabIndex        =   11
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lblEnd 
         Caption         =   "End"
         Height          =   225
         Left            =   3660
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblAll 
         Caption         =   "All"
         Height          =   225
         Left            =   4680
         TabIndex        =   9
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.Frame frameReports 
      Caption         =   "Select Report"
      Height          =   4875
      Left            =   150
      TabIndex        =   0
      Top             =   510
      Width           =   2175
      Begin VB.OptionButton opSOcompPO 
         Caption         =   "SO > QtyOnPO"
         Height          =   255
         Left            =   150
         TabIndex        =   59
         Top             =   4500
         Width           =   1725
      End
      Begin VB.OptionButton opXZSOnot0 
         Caption         =   "X-Z SO Not 0"
         Height          =   225
         Left            =   150
         TabIndex        =   58
         Top             =   4200
         Width           =   1455
      End
      Begin VB.OptionButton opXZQOHnot0 
         Caption         =   "X-Z QOH Not 0"
         Height          =   225
         Left            =   150
         TabIndex        =   57
         Top             =   3900
         Width           =   1725
      End
      Begin VB.OptionButton opXZQOHSOnot0 
         Caption         =   "X-Z QOH SO Not 0"
         Height          =   255
         Left            =   150
         TabIndex        =   56
         Top             =   3600
         Width           =   1695
      End
      Begin VB.OptionButton opGM 
         Caption         =   "Gross Margin"
         Height          =   255
         Left            =   150
         TabIndex        =   51
         Top             =   3300
         Width           =   1545
      End
      Begin VB.OptionButton opProdLines 
         Caption         =   "Product Lines"
         Height          =   255
         Left            =   150
         TabIndex        =   50
         Top             =   3000
         Width           =   1605
      End
      Begin VB.OptionButton opInvWeights 
         Caption         =   "Inventory Weights"
         Height          =   255
         Left            =   150
         TabIndex        =   49
         Top             =   2700
         Width           =   1815
      End
      Begin VB.OptionButton opOverOP 
         Caption         =   "Over Order Point"
         Height          =   255
         Left            =   150
         TabIndex        =   46
         Top             =   2400
         Width           =   1725
      End
      Begin VB.OptionButton opShowInvPrice 
         Caption         =   "Show Inventory/Price"
         Height          =   255
         Left            =   150
         TabIndex        =   43
         Top             =   2100
         Width           =   1875
      End
      Begin VB.OptionButton opSalesRanking 
         Caption         =   "Sales Ranking"
         Height          =   255
         Left            =   150
         TabIndex        =   29
         Top             =   1800
         Width           =   1425
      End
      Begin VB.OptionButton opZeroNegSales 
         Caption         =   "Zero/Negative Sales"
         Height          =   255
         Left            =   150
         TabIndex        =   28
         Top             =   1500
         Width           =   1785
      End
      Begin VB.OptionButton opPriceBreaks 
         Caption         =   "Price Breaks"
         Height          =   285
         Left            =   150
         TabIndex        =   27
         Top             =   1200
         Width           =   1425
      End
      Begin VB.OptionButton opPriceList 
         Caption         =   "Price List"
         Height          =   285
         Left            =   150
         TabIndex        =   5
         Top             =   900
         Width           =   1155
      End
      Begin VB.OptionButton opInventory 
         Caption         =   "Inventory"
         Height          =   285
         Left            =   150
         TabIndex        =   4
         Top             =   600
         Width           =   1395
      End
      Begin VB.OptionButton opBelowOrderPoint 
         Caption         =   "Below Order Point"
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Reports and Forms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3060
      TabIndex        =   3
      Top             =   60
      Width           =   2355
   End
End
Attribute VB_Name = "ReportsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload ReportsMenu
    MainMenu.Show
End Sub

Private Sub btnPreview_Click()
    Dim reportName As String, strsql As String
    strsql = PrepareReport()
    Dim rst As ADODB.Recordset
    Set rst = retrieve(strsql)
    'Load Inventory
    'Set Inventory.DataSource = rst
    'Inventory.Show
    'Inventory.
    'Load reportName
    'Forms!reportName.Show
End Sub

Private Sub btnPrint_Click()
    Dim copies As String
    'Do
        copies = InputBox("Enter the number of copies to be printed:", , "1")
    'While copies <> "" Or Not IsNumeric(copies)
    If copies = "" Then
        Exit Sub
    End If
    Dim reportName As String, strsql As String
    'reportName = PrepareReport(strsql)
    'call print?
End Sub

Private Sub chkLineAll_Click()
    Me.cmbLineStart.Enabled = Not CBool(Me.chkLineAll)
    Me.cmbLineEnd.Enabled = Not CBool(Me.chkLineAll)
End Sub

Private Sub chkWhseAll_Click()
    Me.cmbWhse.Enabled = Not CBool(Me.chkWhseAll)
End Sub

Private Sub Form_Load()
    requeryLC
    requeryWhse
    Me.RankA = "70"
    Me.RankB = "80"
    Me.RankC = "90"
    Me.RankD = "100"
    Me.chkRankA = 1
    Me.chkRankB = 1
    Me.chkRankC = 1
    Me.chkRankD = 1
    Me.useYTD = 1
End Sub




Private Sub requeryLC()
    Dim rst As ADODB.Recordset
    Set rst = retrieve("SELECT ProductLine FROM ProductLine WHERE Hide=0 ORDER BY ProductLine")
    While Not rst.EOF
        Me.cmbLineStart.AddItem (rst("ProductLine"))
        Me.cmbLineEnd.AddItem (rst("ProductLine"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryWhse()
    Dim rst As ADODB.Recordset
    Set rst = retrieve("SELECT WhseCode, WhseDescription FROM IMC_WarehouseCode ORDER BY WhseCode")
    While Not rst.EOF
        Me.cmbWhse.AddItem (Trim(rst("WhseCode")) & "  |  " & Trim(rst("WhseDescription")))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Function PrepareReport() As String
    Dim strsql As String, whereClause As String
    If Me.opBelowOrderPoint Or _
       Me.opInventory Or _
       Me.opPriceList Or _
       Me.opPriceBreaks Or _
       Me.opInvWeights Then
        whereClause = prepareProdLinesWhse & prepareOnlyNots
        strsql = "SELECT * FROM InventoryMaster, PartNumbers, IM2_InventoryItemWhseDetl WHERE InventoryMaster.ItemNumber=PartNumbers.ItemNumber AND InventoryMaster.ItemNumber=IM2_InventoryItemWhseDetl.ItemNumber " & whereClause
        PrepareReport = strsql
    End If

End Function

Private Function prepareOnlyNots() As String
    Dim str As String
    str = ""
    If Me.onlyDC Then
        str = str & " AND DiscontinuedFlag=1"
    ElseIf Me.notDC Then
        str = str & " AND DiscontinuedFlag=0"
    End If
    If Me.onlyBeingDC Then
        str = str & " AND DiscontinuedFlag=1 AND QtyOnHand>AvailabilityLimit"
    ElseIf Me.notBeingDC Then
        str = str & " AND NOT (DiscontinuedFlag=1 AND QtyOnHand>AvailabilityLimit"
    End If
    If Me.onlyPosQty Then
        str = str & " AND QtyOnHand>0"
    ElseIf Me.onlyNegQty Then
        str = str & " AND QtyOnHand<=0"
    End If
    If Me.onlyDropShip Then
        str = str & " AND DropShippable=1"
    End If
    prepareOnlyNots = str
End Function

Private Function prepareProdLinesWhse() As String
    If Me.chkLineAll And Me.chkWhseAll Then
        prepareProdLinesWhse = ""
    ElseIf Me.chkLineAll And Not CBool(Me.chkWhseAll) Then
        prepareProdLinesWhse = " AND WhseCode='" & Trim(Left(Me.cmbWhse, InStr(Me.cmbWhse, " |  "))) & "'"
    ElseIf Me.chkWhseAll And Not CBool(Me.chkLineAll) Then
        prepareProdLinesWhse = " AND ProductLine BETWEEN '" & Me.cmbLineStart & "' AND '" & Me.cmbLineEnd & "'"
    Else
        prepareProdLinesWhse = " AND ProductLine BETWEEN '" & Me.cmbLineStart & "' AND '" & Me.cmbLineEnd & "' AND WhseCode='" & Trim(Left(Me.cmbWhse, InStr(Me.cmbWhse, "|  "))) & "'"""
    End If
End Function




'---------------------------
' all the hiding/unhiding
'---------------------------
Private Sub disableAll()
    handleLCWhse (False)
    Me.onlyDC.Visible = False
    Me.notDC.Visible = False
    Me.onlyBeingDC.Visible = False
    Me.notBeingDC.Visible = False
    Me.onlyPosQty.Visible = False
    Me.onlyNegQty.Visible = False
    Me.useCurrentPeriod.Visible = False
    Me.useLastPeriod.Visible = False
    Me.useYTD.Visible = False
    Me.onlyDropShip.Visible = False
    handleRank (False)
    Me.lblPctDiscount.Visible = False
    Me.pctDiscount.Visible = False
    Me.lblPctOverOP.Visible = False
    Me.pctOverOP.Visible = False
    handleGM False
End Sub

Private Sub handleLCWhse(tf As Boolean)
    Me.cmbLineStart.Visible = tf
    Me.cmbLineEnd.Visible = tf
    Me.chkLineAll.Visible = tf
    Me.cmbWhse.Visible = tf
    Me.chkWhseAll.Visible = tf
    Me.lblStart.Visible = tf
    Me.lblEnd.Visible = tf
    Me.lblAll.Visible = tf
    Me.lblLC.Visible = tf
    Me.lblWhse.Visible = tf
End Sub

Private Sub handleRank(tf As Boolean)
    Me.lblA.Visible = tf
    Me.lblB.Visible = tf
    Me.lblC.Visible = tf
    Me.lblD.Visible = tf
    Me.RankA.Visible = tf
    Me.RankB.Visible = tf
    Me.RankC.Visible = tf
    Me.RankD.Visible = tf
    Me.chkRankA.Visible = tf
    Me.chkRankB.Visible = tf
    Me.chkRankC.Visible = tf
    Me.chkRankD.Visible = tf
    Me.lblOnlyUse.Visible = tf
End Sub

Private Sub handleGM(tf As Boolean)
    Me.lblBetween.Visible = tf
    Me.gmPct1.Visible = tf
    Me.lblAnd.Visible = tf
    Me.gmPct2.Visible = tf
End Sub

Private Sub notBeingDC_Click()
    If Me.notBeingDC Then
        Me.onlyBeingDC = False
    End If
End Sub

Private Sub notDC_Click()
    If Me.notDC Then
        Me.onlyDC = False
    End If
End Sub

Private Sub onlyBeingDC_Click()
    If Me.onlyBeingDC Then
        Me.notBeingDC = False
        Me.onlyDC = False
    End If
End Sub

Private Sub onlyDC_Click()
    If Me.onlyDC Then
        Me.notDC = False
        Me.onlyBeingDC = False
    End If
End Sub

Private Sub onlyNegQty_Click()
    If Me.onlyNegQty Then
        Me.onlyPosQty = False
    End If
End Sub

Private Sub onlyPosQty_Click()
    If Me.onlyPosQty Then
        Me.onlyNegQty = False
    End If
End Sub

Private Sub opBelowOrderPoint_Click()
    disableAll
    handleLCWhse True
    Me.onlyDC.Visible = True
    Me.notDC.Visible = True
    Me.onlyBeingDC.Visible = True
    Me.notBeingDC.Visible = True
    Me.onlyPosQty.Visible = True
    Me.onlyNegQty.Visible = True
    Me.onlyDropShip.Visible = True
End Sub

Private Sub opGM_Click()
    disableAll
    handleLCWhse True
    Me.onlyDC.Visible = True
    Me.notDC.Visible = True
    Me.onlyBeingDC.Visible = True
    Me.notBeingDC.Visible = True
    Me.onlyPosQty.Visible = True
    Me.onlyNegQty.Visible = True
    Me.onlyDropShip.Visible = True
    handleGM True
End Sub

Private Sub opInventory_Click()
    disableAll
    handleLCWhse True
    Me.onlyDC.Visible = True
    Me.notDC.Visible = True
    Me.onlyBeingDC.Visible = True
    Me.notBeingDC.Visible = True
    Me.onlyPosQty.Visible = True
    Me.onlyNegQty.Visible = True
    Me.onlyDropShip.Visible = True
End Sub

Private Sub opInvWeights_Click()
    disableAll
    handleLCWhse True
    Me.onlyDC.Visible = True
    Me.notDC.Visible = True
    Me.onlyBeingDC.Visible = True
    Me.notBeingDC.Visible = True
    Me.onlyPosQty.Visible = True
    Me.onlyNegQty.Visible = True
    Me.onlyDropShip.Visible = True
End Sub

Private Sub opOverOP_Click()
    disableAll
    handleLCWhse True
    Me.onlyDC.Visible = True
    Me.notDC.Visible = True
    Me.onlyBeingDC.Visible = True
    Me.notBeingDC.Visible = True
    Me.onlyPosQty.Visible = True
    Me.onlyNegQty.Visible = True
    Me.onlyDropShip.Visible = True
    Me.lblPctOverOP.Visible = True
    Me.pctOverOP.Visible = True
End Sub

Private Sub opPriceBreaks_Click()
    disableAll
    handleLCWhse True
    Me.onlyDC.Visible = True
    Me.notDC.Visible = True
    Me.onlyBeingDC.Visible = True
    Me.notBeingDC.Visible = True
    Me.onlyPosQty.Visible = True
    Me.onlyNegQty.Visible = True
    Me.onlyDropShip.Visible = True
End Sub

Private Sub opPriceList_Click()
    disableAll
    handleLCWhse True
    Me.onlyDC.Visible = True
    Me.notDC.Visible = True
    Me.onlyBeingDC.Visible = True
    Me.notBeingDC.Visible = True
    Me.onlyPosQty.Visible = True
    Me.onlyNegQty.Visible = True
    Me.onlyDropShip.Visible = True
End Sub

Private Sub opProdLines_Click()
    disableAll
    handleLCWhse True
End Sub

Private Sub opSalesRanking_Click()
    disableAll
    handleLCWhse True
    Me.onlyDC.Visible = True
    Me.notDC.Visible = True
    Me.onlyBeingDC.Visible = True
    Me.notBeingDC.Visible = True
    Me.onlyPosQty.Visible = True
    Me.onlyNegQty.Visible = True
    Me.onlyDropShip.Visible = True
    Me.useCurrentPeriod.Visible = True
    Me.useLastPeriod.Visible = True
    Me.useYTD.Visible = True
    handleRank True
End Sub

Private Sub opShowInvPrice_Click()
    disableAll
    handleLCWhse True
    Me.onlyDC.Visible = True
    Me.notDC.Visible = True
    Me.onlyBeingDC.Visible = True
    Me.notBeingDC.Visible = True
    Me.onlyPosQty.Visible = True
    Me.onlyNegQty.Visible = True
    Me.onlyDropShip.Visible = True
    Me.lblPctDiscount.Visible = True
    Me.pctDiscount.Visible = True
End Sub

Private Sub opSOcompPO_Click()
    disableAll
End Sub

Private Sub opXZQOHnot0_Click()
    disableAll
End Sub

Private Sub opXZQOHSOnot0_Click()
    disableAll
End Sub

Private Sub opXZSOnot0_Click()
    disableAll
End Sub

Private Sub opZeroNegSales_Click()
    disableAll
    handleLCWhse True
    Me.onlyDC.Visible = True
    Me.notDC.Visible = True
    Me.onlyBeingDC.Visible = True
    Me.notBeingDC.Visible = True
    Me.onlyPosQty.Visible = True
    Me.onlyNegQty.Visible = True
    Me.onlyDropShip.Visible = True
    Me.useCurrentPeriod.Visible = True
    Me.useLastPeriod.Visible = True
    Me.useYTD.Visible = True
End Sub

Private Sub useCurrentPeriod_Click()
    If Me.useCurrentPeriod Then
        Me.useLastPeriod = False
        Me.useYTD = False
    End If
End Sub

Private Sub useLastPeriod_Click()
    If Me.useLastPeriod Then
        Me.useCurrentPeriod = False
        Me.useYTD = False
    End If
End Sub

Private Sub useYTD_Click()
    If Me.useYTD Then
        Me.useCurrentPeriod = False
        Me.useLastPeriod = False
    End If
End Sub
