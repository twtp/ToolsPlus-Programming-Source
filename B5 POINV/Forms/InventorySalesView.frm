VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#2.1#0"; "TPControls.ocx"
Begin VB.Form InventorySalesView 
   Caption         =   "Inventory - View Sales"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9630
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox TimePeriod 
      Height          =   315
      Left            =   4110
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   60
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter Options:"
      Height          =   885
      Left            =   90
      TabIndex        =   4
      Top             =   5250
      Width           =   6645
      Begin VB.ComboBox cmbVendorFilter 
         Height          =   315
         Left            =   4260
         TabIndex        =   13
         Top             =   150
         Width           =   2265
      End
      Begin VB.ComboBox cmbProductLineFilter 
         Height          =   315
         Left            =   4260
         TabIndex        =   9
         Top             =   480
         Width           =   2265
      End
      Begin VB.CheckBox chkHideXXX 
         Caption         =   "Hide XXX"
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   270
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CheckBox chkHideZeroQOH 
         Caption         =   "Hide QOH = 0"
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   270
         Width           =   1545
      End
      Begin VB.CheckBox chkHideDropship 
         Caption         =   "Hide Dropship"
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   540
         Width           =   1545
      End
      Begin VB.CheckBox chkHideZZZ 
         Caption         =   "Hide ZZZ"
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   540
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.Label Label4 
         Caption         =   "Vendor:"
         Height          =   225
         Left            =   3180
         TabIndex        =   14
         Top             =   210
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Product Line:"
         Height          =   225
         Left            =   3180
         TabIndex        =   10
         Top             =   540
         Width           =   1035
      End
   End
   Begin VB.CommandButton btnSetFilter 
      Caption         =   "Set PO/INV Filter"
      Height          =   795
      Left            =   6810
      TabIndex        =   3
      Top             =   5280
      Width           =   1755
   End
   Begin TPControls.SimpleListView ItemList 
      Height          =   4635
      Left            =   90
      TabIndex        =   2
      Top             =   510
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   8176
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Close"
      Height          =   405
      Left            =   8340
      TabIndex        =   1
      Top             =   30
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Time Period:"
      Height          =   255
      Left            =   3150
      TabIndex        =   12
      Top             =   90
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Inventory - View Sales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2775
   End
End
Attribute VB_Name = "InventorySalesView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private changed As Boolean
Private fillingform As Boolean

Private sortColumn As Long
Private sortdescending As Boolean

Private Const NUM_COLUMNS As Long = 8

Private Const COL_ITEMNUMBER As Long = 0
Private Const COL_QTY_ON_HAND As Long = 1
Private Const COL_TOTAL_SOLD As Long = 2
Private Const COL_SO_SOLD As Long = 3
Private Const COL_P2_SOLD As Long = 4
Private Const COL_SOP2_RATIO As Long = 5
Private Const COL_UNITS_WEEK As Long = 6
Private Const COL_STOCK_SALES_RATIO As Long = 7

Private Const CLOSE_BUTTON_LEFT_MINIMUM As Long = 7200

Public Function GetItemSelection() As String()
    GetItemSelection = Me.ItemList.GetTextSelected(COL_ITEMNUMBER)
End Function

Private Sub btnSetFilter_Click()
    If Me.ItemList.SelCount = 0 Then
        MsgBox "You don't have any items selected!"
    ElseIf Not IsFormLoaded("InventoryMaintenance") Then
        MsgBox "Inventory Maintenance is not loaded!"
    Else
        InventoryMaintenance.SetFilter "Filter By Sales"
    End If
End Sub

Private Sub chkHideDropship_Click()
    requeryItemList
End Sub

Private Sub chkHideXXX_Click()
    requeryItemList
End Sub

Private Sub chkHideZZZ_Click()
    requeryItemList
End Sub

Private Sub chkHideZeroQOH_Click()
    requeryItemList
End Sub

Private Sub cmbProductLineFilter_Click()
    requeryItemList
End Sub

Private Sub cmbProductLineFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.cmbProductLineFilter, KeyCode, Shift
    changed = True
    If KeyCode = vbKeyReturn Then
        cmbProductLineFilter_Click
    End If
End Sub

Private Sub cmbProductLineFilter_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.cmbProductLineFilter, KeyAscii
    changed = True
End Sub

Private Sub cmbProductLineFilter_LostFocus()
    AutoCompleteLostFocus Me.cmbProductLineFilter
    If Me.cmbProductLineFilter = "" Then
        Me.cmbProductLineFilter = "<All>"
    ElseIf changed Then
        cmbProductLineFilter_Click
        changed = False
    End If
End Sub

Private Sub cmbVendorFilter_Click()
    requeryItemList
End Sub

Private Sub cmbVendorFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.cmbVendorFilter, KeyCode, Shift
    changed = True
    If KeyCode = vbKeyReturn Then
        cmbVendorFilter_Click
    End If
End Sub

Private Sub cmbVendorFilter_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.cmbVendorFilter, KeyAscii
    changed = True
End Sub

Private Sub cmbVendorFilter_LostFocus()
    AutoCompleteLostFocus Me.cmbVendorFilter
    If Me.cmbVendorFilter = "" Then
        Me.cmbVendorFilter = "<All>"
    ElseIf changed Then
        cmbVendorFilter_Click
        changed = False
    End If
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    fillingform = True
    Me.ItemList.MultiSelect = True
    sortColumn = COL_TOTAL_SOLD
    sortdescending = True
    Me.ItemList.SetColumnNames Array("ItemNumber", "Qty On Hand", "Total Sold", "SO Sales", "P2 Sales", "SO/P2 Ratio", "Units/Week", "Stock/Sales Ratio")
    Me.ItemList.SetColumnWidths Array(1500, 1000, 1000, 1000, 1000, 1000, 1000, 1000)
    requeryProductLineFilter
    requeryVendorFilter
    requeryTimePeriod
    fillingform = False
    requeryItemList
End Sub

Private Sub requeryProductLineFilter()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ProductLine FROM ProductLine WHERE Hide=0")
    Me.cmbProductLineFilter.Clear
    Me.cmbProductLineFilter.AddItem "<All>"
    While Not rst.EOF
        Me.cmbProductLineFilter.AddItem rst("ProductLine")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Me.cmbProductLineFilter = "<All>"
End Sub

Private Sub requeryVendorFilter()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT VendorName FROM AP_Vendor ORDER BY VendorName")
    Me.cmbVendorFilter.Clear
    Me.cmbVendorFilter.AddItem "<All>"
    While Not rst.EOF
        Me.cmbVendorFilter.AddItem rst("VendorName")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Me.cmbVendorFilter = "<All>"
End Sub

Private Sub requeryTimePeriod()
    Me.TimePeriod.Clear
    Me.TimePeriod.AddItem "This Month"
    Me.TimePeriod.AddItem "Last Month"
    Me.TimePeriod.AddItem "Current Period"
    Me.TimePeriod.AddItem "Last Period"
    Me.TimePeriod.AddItem "Sales Year-to-Date"
    Me.TimePeriod.AddItem "Sales Last Year"
    Me.TimePeriod.AddItem "Last 30 Days"
    Me.TimePeriod.AddItem "30 to 60 Days Ago"
    Me.TimePeriod.AddItem "Last 90 Days"
    Me.TimePeriod.AddItem "90 to 180 Days Ago"
    Me.TimePeriod.AddItem "Last 365 Days"
    Me.TimePeriod.AddItem "365 to 730 Days Ago"
    Me.TimePeriod = "This Month"
End Sub

Private Sub requeryItemList()
    If fillingform Then
        Exit Sub
    End If
    
    Dim lastItem As Variant
    If Me.ItemList.ListCount <> 0 And Me.ItemList.SelCount <> 0 Then
        lastItem = Me.ItemList.GetTextSelected(COL_ITEMNUMBER)
    Else
        lastItem = Split("", "zerolengtharray")
    End If
    Dim periodType As Long, period As Long
    periodType = getPeriodType()
    period = getPeriod()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve(buildQuery())
    Me.ItemList.Clear
    While Not rst.EOF
        Me.ItemList.Add formatFieldDisplays(rst)
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    If UBound(lastItem) >= 0 Then
        Dim i As Long, j As Long, k As Long
        For i = 0 To Me.ItemList.ListCount - 1
            For j = 0 To UBound(lastItem)
                If lastItem(j) = "" Then
                    Exit For
                End If
                If Me.ItemList.GetText(i, COL_ITEMNUMBER) = lastItem(j) Then
                    Me.ItemList.SelectRow i, True, True, True
                    For k = j + 1 To UBound(lastItem)
                        lastItem(k - 1) = lastItem(k)
                    Next k
                    lastItem(UBound(lastItem)) = ""
                    Exit For
                End If
            Next j
            If lastItem(0) = "" Then
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub ItemList_ColumnClicked(i As Long)
    If i = sortColumn Then
        sortdescending = Not sortdescending
    Else
        sortdescending = True
        sortColumn = i
    End If
    requeryItemList
End Sub

Private Sub ItemList_DblClick(i As Long, j As Long, txt As String)
    If IsFormLoaded("InventoryMaintenance") Then
        InventoryMaintenance.GoToItem CStr(Me.ItemList.GetTextSelected(0)(0)), True
        FocusOnForm "InventoryMaintenance"
    End If
End Sub

Private Function getPeriodType() As Long
    '0 is static
    '1 is rolling
    Select Case Me.TimePeriod
        Case Is = "This Month"
            getPeriodType = 0
        Case Is = "Last Month"
            getPeriodType = 0
        Case Is = "Current Period"
            getPeriodType = 0
        Case Is = "Last Period"
            getPeriodType = 0
        Case Is = "Sales Year-to-Date"
            getPeriodType = 0
        Case Is = "Sales Last Year"
            getPeriodType = 0
        Case Is = "Last 30 Days"
            getPeriodType = 1
        Case Is = "30 to 60 Days Ago"
            getPeriodType = 1
        Case Is = "Last 90 Days"
            getPeriodType = 1
        Case Is = "90 to 180 Days Ago"
            getPeriodType = 1
        Case Is = "Last 365 Days"
            getPeriodType = 1
        Case Is = "365 to 730 Days Ago"
            getPeriodType = 1
    End Select
End Function

Private Function getPeriod() As Long
    Select Case Me.TimePeriod
        Case Is = "This Month"
            getPeriod = 0
        Case Is = "Last Month"
            getPeriod = 1
        Case Is = "Current Period"
            getPeriod = 2
        Case Is = "Last Period"
            getPeriod = 3
        Case Is = "Sales Year-to-Date"
            getPeriod = 4
        Case Is = "Sales Last Year"
            getPeriod = 5
        Case Is = "Last 30 Days"
            getPeriod = 0
        Case Is = "30 to 60 Days Ago"
            getPeriod = 1
        Case Is = "Last 90 Days"
            getPeriod = 2
        Case Is = "90 to 180 Days Ago"
            getPeriod = 3
        Case Is = "Last 365 Days"
            getPeriod = 4
        Case Is = "365 to 730 Days Ago"
            getPeriod = 5
    End Select
End Function

Private Function getSortColumn() As String
    Select Case sortColumn
        Case Is = COL_ITEMNUMBER
            getSortColumn = "S.ItemNumber"
        Case Is = COL_QTY_ON_HAND
            getSortColumn = "QtyOnHand"
        Case Is = COL_TOTAL_SOLD
            getSortColumn = "TotalSales"
        Case Is = COL_SO_SOLD
            getSortColumn = "SOSales"
        Case Is = COL_P2_SOLD
            getSortColumn = "P2Sales"
        Case Is = COL_SOP2_RATIO
            getSortColumn = "SOP2Ratio"
        Case Is = COL_UNITS_WEEK
            getSortColumn = "UPW"
        Case Is = COL_STOCK_SALES_RATIO
            getSortColumn = "StockSalesRatio"
    End Select
End Function

Private Function getSortOrder() As String
    getSortOrder = IIf(sortdescending, "DESC", "ASC")
End Function

Private Sub TimePeriod_Click()
    requeryItemList
End Sub

Private Function buildQuery() As String
    Dim periodType As Long, period As Long
    periodType = getPeriodType()
    period = getPeriod()
    
    Dim sofield As String
    sofield = "S.Per" & period & "SO"
    Dim p2field As String
    p2field = "S.Per" & period & "P2"
    Dim totalfield As String
    totalfield = "(" & sofield & " + " & p2field & ")"
    
    Dim fields(NUM_COLUMNS - 1) As String
    
    fields(COL_ITEMNUMBER) = "S.ItemNumber"
    fields(COL_QTY_ON_HAND) = "vActualWhseQty.QtyOnHand"
    fields(COL_TOTAL_SOLD) = totalfield & " AS TotalSales"
    fields(COL_SO_SOLD) = sofield & " AS SOSales"
    fields(COL_P2_SOLD) = p2field & " AS P2Sales"
    fields(COL_SOP2_RATIO) = "CAST(" & sofield & " as float) / NULLIF(" & p2field & ",0) AS SOP2Ratio"
    fields(COL_UNITS_WEEK) = totalfield & " / (dbo.CalcDaysInPeriod(" & periodType & "," & period & ",GETDATE()) / 7) AS UPW"
    fields(COL_STOCK_SALES_RATIO) = "CAST(vActualWhseQty.QtyOnHand as float) / NULLIF(" & totalfield & ",0) AS StockSalesRatio"

    buildQuery = "SELECT " & Join(fields, ", ") & " " & _
                 "FROM vInventoryStatisticsDenormalizedAllWhses AS S" & _
                 "  INNER JOIN vItemStatusBreakdown ON S.ItemNumber=vItemStatusBreakdown.ItemNumber " & _
                 "  INNER JOIN vActualWhseQty ON S.ItemNumber=vActualWhseQty.ItemNumber " & _
                 "  INNER JOIN InventoryMaster ON S.ItemNumber=InventoryMaster.ItemNumber " & _
                 "WHERE vItemStatusBreakdown.DC=0 " & _
                 "  AND S.PeriodType=" & periodType & " " & _
                 IIf(Me.chkHideXXX, "AND S.ItemNumber NOT LIKE 'XX%' ", "") & _
                 IIf(Me.chkHideZZZ, "AND S.ItemNumber NOT LIKE 'ZZ%' ", "") & _
                 IIf(Me.chkHideZeroQOH, "AND vActualWhseQty.QtyOnHand>0 ", "") & _
                 IIf(Me.chkHideDropship, "AND vItemStatusBreakdown.DSable=0 ", "") & _
                 addProductLineFilterClause() & _
                 addVendorFilterClause() & _
                 "ORDER BY " & getSortColumn() & " " & getSortOrder()
End Function

Private Function formatFieldDisplays(rst As ADODB.Recordset) As Variant
    Dim fields(NUM_COLUMNS - 1) As String
    
    fields(COL_ITEMNUMBER) = rst(COL_ITEMNUMBER)
    fields(COL_QTY_ON_HAND) = rst(COL_QTY_ON_HAND)
    fields(COL_TOTAL_SOLD) = rst(COL_TOTAL_SOLD)
    fields(COL_SO_SOLD) = rst(COL_SO_SOLD)
    fields(COL_P2_SOLD) = rst(COL_P2_SOLD)
    
    If fields(COL_SO_SOLD) = 0 And fields(COL_P2_SOLD) = 0 Then
        'no sales for either, so no ratio at all
        fields(COL_SOP2_RATIO) = ""
    ElseIf fields(COL_SO_SOLD) = 0 And fields(COL_P2_SOLD) <> 0 Then
        'no SO sales, so the ratio can't be calculated
        fields(COL_SOP2_RATIO) = "0"
    ElseIf fields(COL_SO_SOLD) <> 0 And fields(COL_P2_SOLD) = 0 Then
        'no P2 sales, infinite ratio
        fields(COL_SOP2_RATIO) = "Inf"
    ElseIf CLng(fields(COL_SO_SOLD)) > CLng(fields(COL_P2_SOLD)) Then
        'something to 1 ratio
        fields(COL_SOP2_RATIO) = Round(rst(COL_SOP2_RATIO), 1) & " : 1"
    ElseIf CLng(fields(COL_SO_SOLD)) < CLng(fields(COL_P2_SOLD)) Then
        '1 to something ratio
        fields(COL_SOP2_RATIO) = "1 : " & Round(fields(COL_P2_SOLD) / fields(COL_SO_SOLD), 1)
    Else
        fields(COL_SOP2_RATIO) = "1 : 1"
    End If
    
    fields(COL_UNITS_WEEK) = Round(rst(COL_UNITS_WEEK), 1)
    
    If Nz(rst(COL_STOCK_SALES_RATIO)) = "" Then
        fields(COL_STOCK_SALES_RATIO) = "Inf"
    Else
        fields(COL_STOCK_SALES_RATIO) = Round(rst(COL_STOCK_SALES_RATIO), 1)
    End If
    
    formatFieldDisplays = fields
End Function

Private Function addProductLineFilterClause() As String
    If Me.cmbProductLineFilter = "<All>" Then
        addProductLineFilterClause = ""
    Else
        Dim clause As String
        clause = "AND ("
        clause = clause & "LEFT(S.ItemNumber,3)='" & Me.cmbProductLineFilter & "'"
        clause = clause & " OR "
        clause = clause & "(LEFT(S.ItemNumber,3)='XXX' AND SUBSTRING(S.ItemNumber,4,3)='" & Me.cmbProductLineFilter & "')"
        clause = clause & " OR "
        clause = clause & "(LEFT(S.ItemNumber,3)='ZZZ' AND SUBSTRING(S.ItemNumber,4,3)='" & Me.cmbProductLineFilter & "')"
        clause = clause & ") "
        addProductLineFilterClause = clause
    End If
End Function

Private Function addVendorFilterClause() As String
    If Me.cmbVendorFilter = "<All>" Then
        addVendorFilterClause = ""
    Else
        addVendorFilterClause = "AND InventoryMaster.PrimaryVendor='" & lookupPrimaryVendorNumber(Me.cmbVendorFilter) & "' "
    End If
End Function

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        'do nothing
    Else
        resizeControls
    End If
End Sub

Private Function resizeControls() As Boolean
    Me.ItemList.width = Me.width - 4 * Me.ItemList.Left
    
    If Me.width - Me.Exit.width - 180 <= CLOSE_BUTTON_LEFT_MINIMUM Then
        Me.Exit.Left = CLOSE_BUTTON_LEFT_MINIMUM
    Else
        Me.Exit.Left = Me.width - Me.Exit.width - 120
    End If
    
    Dim newListHeight As Long
    newListHeight = Me.Height - Me.ItemList.Top - Me.Frame1.Height - 480
    If newListHeight < 30 Then
        newListHeight = 30
    End If
    Me.ItemList.Height = newListHeight
    Me.Frame1.Top = Me.ItemList.Top + Me.ItemList.Height + 60
    Me.btnSetFilter.Top = Me.Frame1.Top + 30
End Function
