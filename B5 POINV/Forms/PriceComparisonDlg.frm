VERSION 5.00
Begin VB.Form PriceComparisonDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Price Comparison for ITEM"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnChangeWebPriceForLine 
      Caption         =   "Change Web Price For LINE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3150
      TabIndex        =   17
      Top             =   720
      Width           =   2145
   End
   Begin VB.TextBox StoreName 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   30
      Width           =   1755
   End
   Begin VB.CommandButton btnIgnoreListRemove 
      Caption         =   "Remove from List"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5490
      TabIndex        =   14
      Top             =   1800
      Width           =   1545
   End
   Begin VB.TextBox PCID 
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Text            =   "ID"
      Top             =   1590
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton btnChangeStoreAndWebPrice 
      Caption         =   "Change Store+Web Price"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3150
      TabIndex        =   12
      Top             =   1020
      Width           =   2145
   End
   Begin VB.CommandButton btnChangeWebPrice 
      Caption         =   "Change Web Price"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3150
      TabIndex        =   11
      Top             =   420
      Width           =   2145
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   2940
      TabIndex        =   10
      Top             =   1530
      Width           =   1245
   End
   Begin VB.ListBox IgnoreList 
      Height          =   1425
      Left            =   5460
      TabIndex        =   9
      Top             =   60
      Width           =   1545
   End
   Begin VB.CommandButton btnIgnoreListAdd 
      Caption         =   "Add to Ignore List"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5490
      TabIndex        =   8
      Top             =   1530
      Width           =   1545
   End
   Begin VB.TextBox FreeShipping 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1020
      Width           =   1755
   End
   Begin VB.TextBox DateSpidered 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   690
      Width           =   1755
   End
   Begin VB.TextBox Source 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   1755
   End
   Begin VB.CommandButton btnGoToWebsite 
      Caption         =   "Go To Website"
      Height          =   285
      Left            =   3150
      TabIndex        =   4
      Top             =   60
      Width           =   2145
   End
   Begin VB.TextBox itemNumber 
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Text            =   "item number"
      Top             =   1590
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label generalLabel 
      Caption         =   "Store Name:"
      Height          =   255
      Index           =   3
      Left            =   90
      TabIndex        =   16
      Top             =   60
      Width           =   1065
   End
   Begin VB.Label generalLabel 
      Caption         =   "Free Shipping:"
      Height          =   255
      Index           =   2
      Left            =   90
      TabIndex        =   3
      Top             =   1050
      Width           =   1065
   End
   Begin VB.Label generalLabel 
      Caption         =   "Date Spidered:"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label generalLabel 
      Caption         =   "Source:"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   390
      Width           =   1065
   End
End
Attribute VB_Name = "PriceComparisonDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnChangeStoreAndWebPrice_Click()
    Dim newprice As String
    newprice = Format(DLookup("Price", "PriceComparisons", "ID=" & Me.PCID), "Currency")
    If MsgBox("Change the STORE and WEB prices for " & Me.ItemNumber & " to " & newprice & "?") Then
        LogItemWebPriceChanged Me.ItemNumber, newprice
        LogItemStorePriceChanged Me.ItemNumber, newprice
        UpdateNotificationEmail "store price", Me.ItemNumber
        UpdateNotificationEmail "web price", Me.ItemNumber
        updateInventoryMaster "IDiscountMarkupPriceRate1", newprice, Me.ItemNumber, ""
        'updateInventoryMaster "OldPrice", "StdPrice", Me.ItemNumber, ""
        updateInventoryMaster "DiscountMarkupPriceRate1", newprice, Me.ItemNumber, ""
        updateInventoryMaster "StdPrice", newprice, Me.ItemNumber, ""
        InventoryMaintenance.IDiscountMarkupPriceRate(1) = newprice
        InventoryMaintenance.DiscountMarkupPriceRate(1) = newprice
        InventoryMaintenance.StdPrice = newprice
        updateDropshipItemWebPrice Me.ItemNumber, newprice
    End If
End Sub

Private Sub btnChangeWebPrice_Click()
    Dim newprice As String
    newprice = Format(DLookup("Price", "PriceComparisons", "ID=" & Me.PCID), "Currency")
    If MsgBox("Change the WEB price for " & Me.ItemNumber & " to " & newprice & "?") Then
        LogItemWebPriceChanged Me.ItemNumber, newprice
        UpdateNotificationEmail "web price", Me.ItemNumber
        updateInventoryMaster "IDiscountMarkupPriceRate1", newprice, Me.ItemNumber, ""
        InventoryMaintenance.IDiscountMarkupPriceRate(1) = newprice
        updateDropshipItemWebPrice Me.ItemNumber, newprice
    End If
End Sub

Private Sub btnChangeWebPriceForLine_Click()
    If MsgBox("Change the WEB price for ALL items in line code to this store's comparison?", vbYesNo) = vbYes Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT * FROM PriceComparisons WHERE ItemNumber LIKE '" & InventoryMaintenance.prodLine & "%' AND StoreName='" & Me.StoreName & "'")
        While Not rst.EOF
            LogItemWebPriceChanged Me.ItemNumber, rst("Price")
            UpdateNotificationEmail "web price", Me.ItemNumber
            DB.Execute "UPDATE InventoryMaster SET IDiscountMarkupPriceRate1=" & rst("Price") & " WHERE ItemNumber='" & rst("ItemNumber") & "'"
            updateDropshipItemWebPrice Me.ItemNumber, rst("Price")
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
    End If
End Sub

Private Sub btnExit_Click()
    Unload PriceComparisonDlg
End Sub

Private Sub btnGoToWebsite_Click()
    OpenDefaultApp DLookup("URL", "PriceComparisons", "ID=" & Me.PCID)
End Sub

Private Sub btnIgnoreListAdd_Click()
    If MsgBox("Add """ & Me.StoreName & """ to the ignore list?", vbYesNo) = vbYes Then
        DB.Execute "INSERT INTO PriceCompIgnore ( StoreName ) VALUES ( '" & EscapeSQuotes(Me.StoreName) & "' )"
        DB.Execute "DELETE FROM PriceComparisons WHERE StoreName='" & EscapeSQuotes(Me.StoreName) & "'"
        Me.IgnoreList.AddItem Me.StoreName
    End If
End Sub

Private Sub btnIgnoreListRemove_Click()
    If MsgBox("Remove " & Me.IgnoreList & " from the ignore list?", vbYesNo) = vbYes Then
        DB.Execute "DELETE FROM PriceCompIgnore WHERE StoreName='" & EscapeSQuotes(Me.IgnoreList) & "'"
    End If
End Sub

Private Sub Form_Load()
    Dim rst As ADODB.Recordset
    If HasPermissionsTo("PriceComparison") Then
        Me.btnChangeStoreAndWebPrice.Enabled = True
        Me.btnChangeWebPrice.Enabled = True
        Me.btnChangeWebPriceForLine.Enabled = True
        Me.btnIgnoreListAdd.Enabled = True
        Me.btnIgnoreListRemove.Enabled = True
    End If
    
    requeryIgnoreList
    
    Me.ItemNumber = InventoryMaintenance.ItemNumber
    If InventoryMaintenance.priceComparison.Selected(0) = False Then
    Me.PCID = Mid(InventoryMaintenance.priceComparison, InStrRev(InventoryMaintenance.priceComparison, vbTab) + 1)
    End If
    If InventoryMaintenance.InternetComparisonLst.Selected(0) = False Then
    Me.PCID = Mid(InventoryMaintenance.InternetComparisonLst, InStrRev(InventoryMaintenance.InternetComparisonLst, vbTab) + 1)
    End If
    If InventoryMaintenance.EBayComparisonsLst.Selected(0) = False Then
    Me.PCID = Mid(InventoryMaintenance.EBayComparisonsLst, InStrRev(InventoryMaintenance.EBayComparisonsLst, vbTab) + 1)
    End If
    
    Me.Caption = Replace(Me.Caption, "ITEM", Me.ItemNumber)
    Set rst = DB.retrieve("SELECT StoreName, Source, LastUpdated, FreeShipping, URL FROM PriceComparisons WHERE ID=" & Me.PCID)
    Me.StoreName = rst("StoreName")
    Me.source = rst("Source")
    Me.DateSpidered = rst("LastUpdated")
    Me.FreeShipping = IIf(rst("FreeShipping"), "Yes", "No")
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryIgnoreList()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT StoreName FROM PriceCompIgnore")
    Me.IgnoreList.Clear
    While Not rst.EOF
        Me.IgnoreList.AddItem rst("StoreName")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub
