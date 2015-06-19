VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#3.5#0"; "TPControls.ocx"
Begin VB.Form EBayKitPriceManager 
   Caption         =   "EBay Kit Price Manager"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   12165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnInventoryMaintenance 
      Caption         =   "To PO/INV"
      Height          =   465
      Left            =   9960
      TabIndex        =   8
      Top             =   600
      Width           =   1425
   End
   Begin VB.CommandButton btnExcel 
      Caption         =   "To Excel"
      Height          =   465
      Left            =   8370
      TabIndex        =   7
      Top             =   600
      Width           =   1425
   End
   Begin VB.CommandButton btnCacheClear 
      Caption         =   "Clear Cache"
      Height          =   435
      Left            =   5880
      TabIndex        =   6
      Top             =   540
      Width           =   1575
   End
   Begin TPControls.NumericUpDown NumericUpDown 
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   1020
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Value           =   "Text1"
      Minimum         =   -99999
      Maximum         =   99999
      Increment       =   1
      DisplayFormat   =   "##"
      AllowDecimals   =   0   'False
      TimeBeforeHoldDown=   200
      TimeBetweenRepeats=   50
      Enabled         =   -1  'True
      BackColor       =   -2147483643
      DefaultValue    =   "30"
   End
   Begin VB.CheckBox chkFilterMargin 
      Caption         =   "Profit Margin Under"
      Height          =   255
      Left            =   750
      TabIndex        =   4
      Top             =   1020
      Width           =   1725
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      Height          =   645
      Left            =   3780
      TabIndex        =   3
      Top             =   480
      Width           =   1365
   End
   Begin VB.CheckBox chkFilterUp 
      Caption         =   "Only Up"
      Height          =   285
      Left            =   750
      TabIndex        =   2
      Top             =   690
      Value           =   1  'Checked
      Width           =   1725
   End
   Begin VB.CheckBox chkFilterPublished 
      Caption         =   "Only Published"
      Height          =   285
      Left            =   750
      TabIndex        =   1
      Top             =   360
      Value           =   1  'Checked
      Width           =   1725
   End
   Begin TPControls.SimpleListView ItemListView 
      Height          =   6135
      Left            =   30
      TabIndex        =   0
      Top             =   1320
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   10821
      MultiSelect     =   0   'False
      Sorted          =   0   'False
      Enabled         =   -1  'True
   End
End
Attribute VB_Name = "EBayKitPriceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cache As Scripting.Dictionary
Private stopFlag As Boolean
Private sortColumn As Long
Private sortdescending As Boolean
Private lastItemList As PerlArray






Private Sub btnInventoryMaintenance_Click()
    ReDim items(Me.ItemListView.ListCount - 1) As String
    Dim i As Long
    For i = 0 To Me.ItemListView.ListCount - 1
        items(i) = Me.ItemListView.GetText(i, 0)
    Next i
    QuickSort items
    If Not IsFormLoaded("InventoryMaintenance") Then
        Load InventoryMaintenance
        InventoryMaintenance.Show
    End If
    InventoryMaintenance.SetFilterByItemList items
End Sub

Private Sub Form_Load()
    If cache Is Nothing Then
        Set cache = New Scripting.Dictionary
    End If
    stopFlag = False
    sortColumn = 0
    sortdescending = False
    
    Me.NumericUpDown.value = Me.NumericUpDown.DefaultValue
    Call chkFilterMargin_Click
    
    Me.ItemListView.SetColumnNames Array("Kit Item", "Published", "Up", "QOH", "Cost of Goods", "Shipping", "Shipping (est)", "EBay Fees", "Total Cost", "EBay Price", "Profit", "Profit Margin")
    Me.ItemListView.SetColumnWidths Array("1450", "900", "900", "900", "1200", "1200", "1200", "1200", "1200", "1200", "1200", "1200")
End Sub

Private Sub Form_Resize()
    If Me.width > 175 Then
        Me.ItemListView.width = Me.width - 150
    End If
    If Me.Height > 1750 Then
        Me.ItemListView.Height = Me.Height - 1725
    End If
End Sub

Private Sub btnLoad_Click()
    Mouse.Hourglass True
    Call displayItems(sortItems(queryItems()))
    Mouse.Hourglass False
End Sub

Private Sub btnCacheClear_Click()
    Set cache = Nothing
    Set cache = New Scripting.Dictionary
End Sub

Private Sub chkFilterMargin_Click()
    If Me.chkFilterMargin Then
        Enable Me.NumericUpDown
    Else
        Disable Me.NumericUpDown
    End If
End Sub

Private Sub btnExcel_Click()
    Mouse.Hourglass True
    Dim excelapp As Excel.Application
    Set excelapp = New Excel.Application
    excelapp.Workbooks.Add
    excelapp.ActiveWorkbook.ActiveSheet.Range("A1").value = Me.ItemListView.GetColumnName(0)
    excelapp.ActiveWorkbook.ActiveSheet.Range("B1").value = Me.ItemListView.GetColumnName(1)
    excelapp.ActiveWorkbook.ActiveSheet.Range("C1").value = Me.ItemListView.GetColumnName(2)
    excelapp.ActiveWorkbook.ActiveSheet.Range("D1").value = Me.ItemListView.GetColumnName(3)
    excelapp.ActiveWorkbook.ActiveSheet.Range("E1").value = Me.ItemListView.GetColumnName(4)
    excelapp.ActiveWorkbook.ActiveSheet.Range("F1").value = Me.ItemListView.GetColumnName(5)
    excelapp.ActiveWorkbook.ActiveSheet.Range("G1").value = Me.ItemListView.GetColumnName(6)
    excelapp.ActiveWorkbook.ActiveSheet.Range("H1").value = Me.ItemListView.GetColumnName(7)
    excelapp.ActiveWorkbook.ActiveSheet.Range("I1").value = Me.ItemListView.GetColumnName(8)
    excelapp.ActiveWorkbook.ActiveSheet.Range("J1").value = Me.ItemListView.GetColumnName(9)
    excelapp.ActiveWorkbook.ActiveSheet.Range("K1").value = Me.ItemListView.GetColumnName(10)
    excelapp.ActiveWorkbook.ActiveSheet.Range("L1").value = Me.ItemListView.GetColumnName(11)
    Dim i As Long
    For i = 0 To Me.ItemListView.ListCount - 1
        excelapp.ActiveWorkbook.ActiveSheet.Range("A" & (i + 2) & ":L" & (i + 2)).value = Me.ItemListView.GetRow(i)
    Next i
    Mouse.Hourglass False
    excelapp.Visible = True
End Sub



Private Sub ItemListView_ColumnClicked(i As Long)
    If sortColumn = i Then
        sortdescending = Not sortdescending
        Call displayItems(reverseItems(lastItemList))
    Else
        sortdescending = False
        sortColumn = i
        Call displayItems(sortItems(lastItemList))
    End If
End Sub

Private Sub ItemListView_DblClick(i As Long, j As Long, txt As String)
    If IsFormLoaded("InventoryMaintenance") Then
        InventoryMaintenance.GoToItem Me.ItemListView.GetTextSelected
    End If
End Sub




Private Function getWhereClause() As String
    If Me.chkFilterMargin Then
        Enable Me.NumericUpDown
    Else
        Disable Me.NumericUpDown
    End If
    
    Dim wc As PerlArray
    Set wc = New PerlArray
    wc.Push "InventoryMaster.IsMASKit=1"
    If Me.chkFilterPublished Then
        wc.Push "EBayPublished=1"
    End If
    If Me.chkFilterUp Then
        wc.Push "EBayStatusID=1"
    End If
    
    getWhereClause = wc.Join(" AND ")
End Function

Private Function queryItems() As PerlArray
    Dim retval As PerlArray
    Set retval = New PerlArray
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT InventoryMaster.ItemNumber, EBayPrice, EBayPublished, EBayStatusID FROM InventoryMaster INNER JOIN PartNumbers ON InventoryMaster.ItemNumber=PartNumbers.ItemNumber WHERE " & getWhereClause())
    Dim rst2 As ADODB.Recordset
    While Not rst.EOF
        Dim item As String
        item = CStr(rst("ItemNumber"))
        
        Dim skipFlag As Boolean
        skipFlag = False
        
        If Not cache.exists(item) Then
            cache.Add item, New Scripting.Dictionary
            cache.item(item).Add "item", CStr(rst("ItemNumber"))
            cache.item(item).Add "pub", CStr(IIf(rst("EBayPublished") = True, "Y", ""))
            cache.item(item).Add "up", CStr(IIf(rst("EBayStatusID") = 1, "Y", ""))
            cache.item(item).Add "price", CDec(rst("EBayPrice"))
        End If
        
        If Not cache.item(item).exists("qoh") Then
            cache.item(item).Add "qoh", CLng(EBayGetAvailableQuantityForKit(rst("ItemNumber")))
        End If
        
        If Not cache.item(item).exists("cog") Then
            Set rst2 = DB.retrieve("SELECT COALESCE(SUM(IM_SalesKitDetail.QuantityPerAssembly * IM2.StdCost),0) FROM InventoryMaster AS IM1 INNER JOIN IM_SalesKitDetail ON IM1.ItemNumber=IM_SalesKitDetail.SalesKitNo INNER JOIN InventoryMaster IM2 ON IM_SalesKitDetail.ComponentItemCode=IM2.ItemNumber WHERE IM1.ItemNumber='" & rst("ItemNumber") & "'")
            cache.item(item).Add "cog", CDec(rst2(0))
            rst2.Close
        End If
        
        If Not cache.item(item).exists("shipping-combined") Then
            Set rst2 = DB.retrieve("SELECT AVG(Cost) FROM FreightActual WHERE ItemNumber='" & rst("ItemNumber") & "'")
            If Not IsNull(rst2(0)) Then
                cache.item(item).Add "shipping-actual", CDec(rst2(0))
                cache.item(item).Add "shipping-estimated", CDec(0)
                cache.item(item).Add "shipping-combined", CDec(rst2(0))
            Else
                rst2.Close
                Set rst2 = DB.retrieve("SELECT COALESCE(AVG(Cost),0) AS AvgCost, Service FROM FreightEstimates WHERE ItemNumber='" & rst("ItemNumber") & "' GROUP BY Service ORDER BY AvgCost")
                If rst2.RecordCount > 0 Then
                    cache.item(item).Add "shipping-actual", CDec(0)
                    cache.item(item).Add "shipping-estimated", CDec(rst2(0))
                    cache.item(item).Add "shipping-combined", CDec(rst2(0))
                Else
                    cache.item(item).Add "shipping-actual", CDec(0)
                    cache.item(item).Add "shipping-estimated", CDec(0)
                    cache.item(item).Add "shipping-combined", CDec(0)
                End If
            End If
            rst2.Close
        End If
        
        If Not cache.item(item).exists("fees") Then
            cache.item(item).Add "fees", CDec(EBay.EBayCalculateTotalSellingFees(CStr(rst("EBayPrice"))) - EBay.EBayCalculateFinalValueDiscount(EBay.EBayCalculateFinalValueFee(CStr(rst("EBayPrice")))))
        End If
        
        If Not cache.item(item).exists("total") Then
            cache.item(item).Add "total", CDec(cache.item(item).item("cog") + cache.item(item).item("shipping-combined") + cache.item(item).item("fees"))
        End If
        
        If Not cache.item(item).exists("margin") Then
            Dim margin As Variant, marginpct As Variant
            If cache.item(item).item("total") = 0 Or rst("EBayPrice") = 0 Then
                margin = CStr("Undefined")
                marginpct = CStr("Undefined")
            Else
                margin = rst("EBayPrice") - cache.item(item).item("total")
                marginpct = CDec(rst("EBayPrice") / cache.item(item).item("total") * 100 - 100)
            End If
            cache.item(item).Add "margin", margin
            cache.item(item).Add "marginpct", marginpct
        End If
        If cache.item(item).item("margin") <> "Undefined" Then
            If Me.chkFilterMargin Then
                If CDec(cache.item(item).item("marginpct")) > CDec(Me.NumericUpDown.value) Then
                    skipFlag = True
                End If
            End If
        End If
        
        If Not skipFlag Then
            retval.Push cache.item(item)
        End If
        
        If stopFlag Then
            rst.MoveLast
            stopFlag = False
        End If
        
        rst.MoveNext
        
        DoEvents
    Wend
    rst.Close
    
    Set queryItems = retval
End Function

Private Function reverseItems(items As PerlArray) As PerlArray
    Set reverseItems = items.Reverse()
End Function

Private Function sortItems(items As PerlArray) As PerlArray
    ReDim temp(items.Scalar - 1) As Scripting.Dictionary
    Dim i As Long
    For i = 0 To items.Scalar - 1
        Set temp(i) = items.Elem(i)
    Next i
    
    QuickSortHashes temp, Array(Array(indexToKey(sortColumn), sortdescending), Array("item", False))
    
    Dim retval As PerlArray
    Set retval = New PerlArray
    For i = 0 To UBound(temp)
        retval.Push temp(i)
    Next i
    Set sortItems = retval
End Function

Private Sub displayItems(items As PerlArray)
    Set lastItemList = items
    
    Me.ItemListView.Clear
    Dim i As Long
    For i = 0 To items.Scalar - 1
        Dim margin As String, marginpct As String
        margin = items.Elem(i).item("margin")
        marginpct = items.Elem(i).item("marginpct")
        If margin = "Undefined" Then
            'skip
        Else
            margin = Format(margin, "Currency")
            marginpct = Round(CDec(marginpct), 1) & "%"
        End If
        Me.ItemListView.Add Array( _
            items.Elem(i).item("item"), _
            items.Elem(i).item("pub"), _
            items.Elem(i).item("up"), _
            items.Elem(i).item("qoh"), _
            Format(items.Elem(i).item("cog"), "Currency"), _
            IIf(items.Elem(i).item("shipping-actual") > 0, Format(items.Elem(i).item("shipping-actual"), "Currency"), ""), _
            IIf(items.Elem(i).item("shipping-actual") = 0, IIf(items.Elem(i).item("shipping-estimated") > 0, Format(items.Elem(i).item("shipping-estimated"), "Currency"), ""), ""), _
            Format(items.Elem(i).item("fees"), "Currency"), _
            Format(items.Elem(i).item("total"), "Currency"), _
            Format(items.Elem(i).item("price"), "Currency"), _
            margin, _
            marginpct _
          )
    Next i
End Sub

Private Function indexToKey(dx As Long) As String
    '("Kit Item", "Published", "Up", "QOH", "Cost of Goods", "Shipping", "Shipping (est)", "EBay Fees", "Total Cost", "EBay Price", "Profit", "Profit Margin")
    Select Case dx
        Case Is = 0
            indexToKey = "item"
        Case Is = 1
            indexToKey = "pub"
        Case Is = 2
            indexToKey = "up"
        Case Is = 3
            indexToKey = "qoh"
        Case Is = 4
            indexToKey = "cog"
        Case Is = 5
            indexToKey = "shipping-combined"
        Case Is = 6
            indexToKey = "shipping-combined"
        Case Is = 7
            indexToKey = "fees"
        Case Is = 8
            indexToKey = "total"
        Case Is = 9
            indexToKey = "price"
        Case Is = 10
            indexToKey = "margin"
        Case Is = 11
            indexToKey = "marginpct"
        Case Else
            indexToKey = "item"
    End Select
End Function
