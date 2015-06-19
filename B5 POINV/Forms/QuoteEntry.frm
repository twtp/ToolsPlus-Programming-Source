VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#3.5#0"; "TPControls.ocx"
Begin VB.Form QuoteEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quote Entry"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10515
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton FillInfoFromMasOrderButton 
      Caption         =   "From MAS"
      Height          =   285
      Left            =   2880
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox CustomerCityField 
      Height          =   285
      Left            =   1410
      TabIndex        =   8
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox CustomerAddressField 
      Height          =   285
      Left            =   1410
      TabIndex        =   6
      Top             =   1020
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   4350
      TabIndex        =   32
      Top             =   1680
      Width           =   5295
      Begin VB.ComboBox QuoteFromComboBox 
         Height          =   315
         ItemData        =   "QuoteEntry.frx":0000
         Left            =   2430
         List            =   "QuoteEntry.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   150
         Width           =   2745
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "How did quote come in to us?"
         Height          =   255
         Index           =   21
         Left            =   90
         TabIndex        =   33
         Top             =   180
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Height          =   525
      Left            =   4350
      TabIndex        =   29
      Top             =   1170
      Width           =   5295
      Begin VB.ComboBox PaymentMethodComboBox 
         Height          =   315
         ItemData        =   "QuoteEntry.frx":003C
         Left            =   2430
         List            =   "QuoteEntry.frx":004C
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   150
         Width           =   2745
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Expected method of payment?"
         Height          =   255
         Index           =   20
         Left            =   90
         TabIndex        =   30
         Top             =   180
         Width           =   2175
      End
   End
   Begin VB.TextBox SelectedCompetitorsURL 
      Height          =   285
      Left            =   1770
      TabIndex        =   54
      Top             =   6480
      Width           =   3495
   End
   Begin VB.TextBox CustomerCompanyField 
      Height          =   285
      Left            =   1410
      TabIndex        =   4
      Top             =   720
      Width           =   2655
   End
   Begin VB.Frame Frame 
      Height          =   1725
      Left            =   6090
      TabIndex        =   55
      Top             =   5280
      Width           =   4395
      Begin VB.TextBox CommentsField 
         Height          =   1215
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         Top             =   390
         Width           =   4185
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Any other thoughts/questions/comments:"
         Height          =   255
         Index           =   18
         Left            =   60
         TabIndex        =   56
         Top             =   180
         Width           =   2955
      End
   End
   Begin VB.CheckBox SelectedPriceBestCheckbox 
      Alignment       =   1  'Right Justify
      Caption         =   "Best"
      Height          =   255
      Left            =   3480
      TabIndex        =   50
      Top             =   5790
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox CustomerEmailField 
      Height          =   285
      Left            =   1410
      TabIndex        =   16
      Top             =   2220
      Width           =   2655
   End
   Begin VB.CommandButton ItemRemoveButton 
      Caption         =   "Del =>"
      Height          =   495
      Left            =   6300
      TabIndex        =   42
      Top             =   4020
      Width           =   435
   End
   Begin VB.CommandButton SubmitButton 
      Caption         =   "Submit"
      Height          =   435
      Left            =   2010
      TabIndex        =   58
      Top             =   7050
      Width           =   1755
   End
   Begin VB.Frame Frame3 
      Height          =   525
      Left            =   4350
      TabIndex        =   23
      Top             =   660
      Width           =   5295
      Begin VB.OptionButton OrderExistsInMasOption 
         Caption         =   "Yes"
         Height          =   225
         Index           =   0
         Left            =   1740
         TabIndex        =   25
         Top             =   180
         Width           =   585
      End
      Begin VB.OptionButton OrderExistsInMasOption 
         Caption         =   "No"
         Height          =   225
         Index           =   1
         Left            =   2400
         TabIndex        =   26
         Top             =   180
         Value           =   -1  'True
         Width           =   585
      End
      Begin VB.TextBox OrderNumberField 
         Height          =   285
         Left            =   3960
         MaxLength       =   7
         TabIndex        =   28
         Top             =   150
         Width           =   1245
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Order already in Mas?"
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   24
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Order #:"
         Height          =   255
         Index           =   5
         Left            =   3270
         TabIndex        =   27
         Top             =   180
         Width           =   645
      End
   End
   Begin VB.CommandButton AllItemsFilterApplyButton 
      Caption         =   "Go"
      Height          =   255
      Left            =   10050
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2400
      Width           =   345
   End
   Begin TPControls.NumericUpDown SelectedPriceUpDown 
      Height          =   285
      Left            =   1770
      TabIndex        =   49
      Top             =   5760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      Value           =   "0.00"
      Minimum         =   0
      Maximum         =   99999
      Increment       =   1
      DisplayFormat   =   "0.00"
      AllowDecimals   =   -1  'True
      TimeBeforeHoldDown=   500
      TimeBetweenRepeats=   50
      Enabled         =   -1  'True
      BackColor       =   -2147483643
      DefaultValue    =   0
   End
   Begin TPControls.NumericUpDown SelectedQuantityUpDown 
      Height          =   285
      Left            =   1770
      TabIndex        =   47
      Top             =   5430
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      Value           =   "0"
      Minimum         =   0
      Maximum         =   99999
      Increment       =   1
      DisplayFormat   =   "0"
      AllowDecimals   =   0   'False
      TimeBeforeHoldDown=   500
      TimeBetweenRepeats=   50
      Enabled         =   -1  'True
      BackColor       =   -2147483643
      DefaultValue    =   0
   End
   Begin VB.TextBox SelectedItemField 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1770
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   5100
      Width           =   1575
   End
   Begin VB.TextBox AllItemsFilterField 
      Height          =   255
      Left            =   8550
      TabIndex        =   39
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton ItemAddButton 
      Caption         =   "Add <="
      Height          =   495
      Left            =   6300
      TabIndex        =   41
      Top             =   3240
      Width           =   435
   End
   Begin VB.ListBox AllItemsList 
      Height          =   2400
      Left            =   6750
      TabIndex        =   43
      Top             =   2640
      Width           =   3735
   End
   Begin TPControls.SimpleListView ItemsListView 
      Height          =   2205
      Left            =   60
      TabIndex        =   36
      Top             =   2820
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   3889
      MultiSelect     =   0   'False
      Sorted          =   -1  'True
      Enabled         =   -1  'True
   End
   Begin VB.TextBox CustomerPhoneNumberField 
      Height          =   285
      Left            =   1410
      TabIndex        =   14
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   525
      Left            =   4350
      TabIndex        =   17
      Top             =   150
      Width           =   5295
      Begin VB.OptionButton ReturningCustomerOption 
         Caption         =   "Yes"
         Height          =   225
         Index           =   0
         Left            =   1530
         TabIndex        =   19
         Top             =   180
         Width           =   585
      End
      Begin VB.OptionButton ReturningCustomerOption 
         Caption         =   "No"
         Height          =   225
         Index           =   1
         Left            =   2190
         TabIndex        =   20
         Top             =   180
         Value           =   -1  'True
         Width           =   585
      End
      Begin VB.TextBox CustomerNumberField 
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   22
         Top             =   150
         Width           =   1245
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Repeat customer?"
         Height          =   255
         Index           =   16
         Left            =   90
         TabIndex        =   18
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer #:"
         Height          =   255
         Index           =   15
         Left            =   3090
         TabIndex        =   21
         Top             =   180
         Width           =   855
      End
   End
   Begin VB.TextBox CustomerZipCodeField 
      Height          =   285
      Left            =   2910
      MaxLength       =   10
      TabIndex        =   12
      Top             =   1620
      Width           =   1155
   End
   Begin VB.TextBox CustomerStateField 
      Height          =   285
      Left            =   1410
      MaxLength       =   2
      TabIndex        =   10
      Top             =   1620
      Width           =   555
   End
   Begin VB.TextBox CustomerNameField 
      Height          =   285
      Left            =   1410
      TabIndex        =   2
      Top             =   420
      Width           =   2655
   End
   Begin TPControls.NumericUpDown SelectedCompetitorsPriceUpDown 
      Height          =   285
      Left            =   1770
      TabIndex        =   52
      Top             =   6150
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      Value           =   "0.00"
      Minimum         =   0
      Maximum         =   99999
      Increment       =   1
      DisplayFormat   =   "0.00"
      AllowDecimals   =   -1  'True
      TimeBeforeHoldDown=   500
      TimeBetweenRepeats=   50
      Enabled         =   -1  'True
      BackColor       =   -2147483643
      DefaultValue    =   0
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "City:"
      Height          =   255
      Index           =   23
      Left            =   90
      TabIndex        =   7
      Top             =   1350
      Width           =   1275
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Address:"
      Height          =   255
      Index           =   22
      Left            =   90
      TabIndex        =   5
      Top             =   1050
      Width           =   1275
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Competitor's URL:"
      Height          =   255
      Index           =   7
      Left            =   420
      TabIndex        =   53
      Top             =   6510
      Width           =   1305
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Competitor's Price:"
      Height          =   255
      Index           =   6
      Left            =   420
      TabIndex        =   51
      Top             =   6180
      Width           =   1305
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Company:"
      Height          =   255
      Index           =   19
      Left            =   90
      TabIndex        =   3
      Top             =   750
      Width           =   1275
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Email:"
      Height          =   255
      Index           =   17
      Left            =   90
      TabIndex        =   15
      Top             =   2250
      Width           =   1275
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Target Price:"
      Height          =   255
      Index           =   14
      Left            =   420
      TabIndex        =   48
      Top             =   5790
      Width           =   1305
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Wanted Quantity:"
      Height          =   255
      Index           =   13
      Left            =   420
      TabIndex        =   46
      Top             =   5460
      Width           =   1305
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Selected Item:"
      Height          =   255
      Index           =   12
      Left            =   420
      TabIndex        =   44
      Top             =   5130
      Width           =   1305
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Filter:"
      Height          =   255
      Index           =   11
      Left            =   8100
      TabIndex        =   38
      Top             =   2400
      Width           =   435
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "All Items:"
      Height          =   255
      Index           =   10
      Left            =   6810
      TabIndex        =   37
      Top             =   2400
      Width           =   705
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Items:"
      Height          =   255
      Index           =   9
      Left            =   150
      TabIndex        =   35
      Top             =   2580
      Width           =   465
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Phone Number:"
      Height          =   255
      Index           =   8
      Left            =   90
      TabIndex        =   13
      Top             =   1950
      Width           =   1275
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip Code:"
      Height          =   255
      Index           =   3
      Left            =   2010
      TabIndex        =   11
      Top             =   1650
      Width           =   855
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "State:"
      Height          =   255
      Index           =   2
      Left            =   90
      TabIndex        =   9
      Top             =   1650
      Width           =   1275
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer Name:"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Width           =   1275
   End
   Begin VB.Label generalLabel 
      Caption         =   "Email a New Quote"
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
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   2325
   End
End
Attribute VB_Name = "QuoteEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private itemDB As Dictionary
Private lastFilter As String
Private currentItemLVIndex As Long
Private stopChangeEvent As Boolean





Private Sub FillInfoFromMasOrderButton_Click()
On Error GoTo errh
    Dim salesOrderNumber As String
    salesOrderNumber = InputBox("Enter sales order number:")
    If salesOrderNumber = "" Then
        Exit Sub
    End If
    
    salesOrderNumber = UCase(salesOrderNumber)
    While Len(salesOrderNumber) < 7
        salesOrderNumber = "0" & salesOrderNumber
    Wend
    
    Dim sql As String
    'no phone number?
    sql = "SELECT ShipToName, ShipToAddress1, ShipToCity, ShipToState, ShipToZipCode, EmailAddress FROM SO_SalesOrderHeader WHERE SalesOrderNo='" & EscapeSQuotes(salesOrderNumber) & "'"
    Dim rst As ADODB.Recordset
    Set rst = DB.MASRetrieveViaJSON(sql)
    If rst.EOF Then
        MsgBox "Error finding sales order!"
        Exit Sub
    End If
    Me.OrderExistsInMasOption(0).value = True
    Me.OrderNumberField.Text = salesOrderNumber
    Me.CustomerNameField.Text = Nz(rst("ShipToName"))
    Me.CustomerAddressField.Text = Nz(rst("ShipToAddress1"))
    Me.CustomerCityField.Text = Nz(rst("ShipToCity"))
    Me.CustomerStateField.Text = Nz(rst("ShipToState"))
    Me.CustomerZipCodeField.Text = Nz(rst("ShipToZipCode"))
    Me.CustomerEmailField.Text = Nz(rst("EmailAddress"))
    rst.Close
    Set rst = Nothing
    
    sql = "SELECT ItemCode, QuantityOrdered, UnitPrice FROM SO_SalesOrderDetail WHERE SalesOrderNo='" & salesOrderNumber & "' AND (ItemType='1' OR ItemType='2')"
    Set rst = DB.MASRetrieveViaJSON(sql)
    While Not rst.EOF
        Dim item As String
        item = Trim(rst("ItemCode"))
        If Not itemDB.exists(item) Then
            MsgBox "Error finding item '" & item & "' in lookup table!"
        Else
            Me.ItemsListView.Add Array(item, CStr(itemDB.item(item).item("Description")), rst("QuantityOrdered"), Format(rst("UnitPrice"), "$0.00"), "", "")
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    Exit Sub
errh:
    MsgBox "Something went wrong, tell Brian!" & vbCrLf & vbCrLf & "Error message: " & Err.Description
End Sub

Private Sub Form_Load()
    Disable Me.CustomerNumberField
    Disable Me.OrderNumberField
    
    Disable Me.SelectedQuantityUpDown
    Disable Me.SelectedPriceUpDown
    Disable Me.SelectedPriceBestCheckbox
    Disable Me.SelectedCompetitorsPriceUpDown
    Disable Me.SelectedCompetitorsURL
    
    Me.PaymentMethodComboBox.ListIndex = 0
    Me.QuoteFromComboBox.ListIndex = 0
    
    lastFilter = ""
    stopChangeEvent = False
    
    Me.ItemsListView.SetColumnNames Array("Item", "Description", "Qty", "Price", "Comp Price", "Comp URL")
    Me.ItemsListView.SetColumnWidths Array("1000", "1650", "600", "900", "1000", "1000")
    
    Dim tabs(0) As Long
    tabs(0) = 60
    SetListTabStops Me.AllItemsList.hwnd, tabs
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT InventoryMaster.ItemNumber, ItemDescription, IDiscountMarkupPriceRate1 FROM InventoryMaster INNER JOIN vItemStatusBreakdown ON InventoryMaster.ItemNumber=vItemStatusBreakdown.ItemNumber WHERE vItemStatusBreakdown.DC=0 AND InventoryMaster.ItemNumber<'Z%' AND InventoryMaster.Hide=0")
    Set itemDB = New Dictionary
    While Not rst.EOF
        Dim temp As Dictionary
        Set temp = New Dictionary
        temp.Add "ItemNumber", CStr(rst("ItemNumber"))
        temp.Add "Description", CStr(rst("ItemDescription"))
        temp.Add "DefaultPrice", CStr(rst("IDiscountMarkupPriceRate1"))
        itemDB.Add CStr(rst("ItemNumber")), temp
        'Me.AllItemsList.AddItem CStr(rst("ItemNumber")) & vbTab & CStr(rst("ItemDescription"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub ReturningCustomerOption_Click(Index As Integer)
    If Index = 0 Then
        Enable Me.CustomerNumberField
    Else
        Disable Me.CustomerNumberField
        Me.CustomerNumberField.Text = ""
    End If
End Sub

Private Sub OrderExistsInMasOption_Click(Index As Integer)
    If Index = 0 Then
        Enable Me.OrderNumberField
    Else
        Disable Me.OrderNumberField
        Me.OrderNumberField.Text = ""
    End If
End Sub

Private Sub AllItemsFilterField_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        AllItemsFilterApplyButton_Click
    End If
End Sub

Private Sub AllItemsFilterApplyButton_Click()
    If Me.AllItemsFilterField.Text = lastFilter Then
        'no change
        Exit Sub
    End If
    
    lastFilter = UCase(Me.AllItemsFilterField.Text)
    Me.AllItemsList.Clear
    Dim iter As Variant
    For Each iter In itemDB.keys
        Dim temp As Dictionary
        Set temp = itemDB.item(CStr(iter))
        If InStr(temp.item("ItemNumber"), lastFilter) >= 1 Then
            Me.AllItemsList.AddItem CStr(temp.item("ItemNumber")) & vbTab & CStr(temp.item("Description"))
        End If
    Next iter
End Sub

Private Sub AllItemsList_DblClick()
    ItemAddButton_Click
End Sub

Private Sub ItemAddButton_Click()
    If Me.AllItemsList.ListIndex = -1 Then
        Exit Sub
    End If
    
    Dim item As String
    item = ListBoxLineColumn(Me.AllItemsList)
    
    Dim i As Long
    For i = 0 To Me.ItemsListView.ListCount - 1
        If Me.ItemsListView.GetText(i, 0) = item Then
            Me.ItemsListView.SelectRow (i)
            Exit Sub
        End If
    Next i
    
    Me.ItemsListView.Add Array(item, CStr(itemDB.item(item).item("Description")), "1", Format(itemDB.item(item).item("DefaultPrice"), "$0.00"), "", "")
End Sub

Private Sub ItemRemoveButton_Click()
    If Me.ItemsListView.SelIndex = -1 Then
        Exit Sub
    End If
    
    Me.ItemsListView.RemoveSelected
End Sub

Private Sub ItemsListView_SelectionChanged()
    stopChangeEvent = True
    If Me.ItemsListView.SelIndex = -1 Then
        Me.SelectedItemField.Text = ""
        Me.SelectedQuantityUpDown.value = 0
        Disable Me.SelectedQuantityUpDown
        Me.SelectedPriceUpDown.value = 0
        Disable Me.SelectedPriceUpDown
        Me.SelectedPriceBestCheckbox.value = vbUnchecked
        Disable Me.SelectedPriceBestCheckbox
        Me.SelectedCompetitorsPriceUpDown.value = 0
        Disable Me.SelectedCompetitorsPriceUpDown
        Me.SelectedCompetitorsURL.Text = ""
        Disable Me.SelectedCompetitorsURL
        currentItemLVIndex = -1
    Else
        Me.SelectedItemField.Text = Me.ItemsListView.GetTextSelected(0)
        Me.SelectedQuantityUpDown.value = Me.ItemsListView.GetTextSelected(2)
        Me.SelectedQuantityUpDown.DefaultValue = Me.SelectedQuantityUpDown.value
        Enable Me.SelectedQuantityUpDown
        Dim ourPrice As String
        ourPrice = Me.ItemsListView.GetTextSelected(3)
        If ourPrice <> "<best>" Then
            Me.SelectedPriceUpDown.value = Replace(ourPrice, "$", "")
            Me.SelectedPriceUpDown.DefaultValue = Me.SelectedPriceUpDown.value
            Enable Me.SelectedPriceUpDown
            Me.SelectedPriceBestCheckbox.value = vbUnchecked
        Else
            Me.SelectedPriceUpDown.value = 0
            Me.SelectedPriceUpDown.DefaultValue = 0
            Disable Me.SelectedPriceUpDown
            Me.SelectedPriceBestCheckbox.value = vbChecked
        End If
        Enable Me.SelectedPriceBestCheckbox
        Dim compPrice As String
        compPrice = Me.ItemsListView.GetTextSelected(4)
        If compPrice <> "" Then
            Me.SelectedCompetitorsPriceUpDown.value = Replace(compPrice, "$", "")
            Me.SelectedCompetitorsPriceUpDown.DefaultValue = Me.SelectedCompetitorsPriceUpDown.value
        Else
            Me.SelectedCompetitorsPriceUpDown.value = 0
            Me.SelectedCompetitorsPriceUpDown.DefaultValue = 0
        End If
        Enable Me.SelectedCompetitorsPriceUpDown
        Me.SelectedCompetitorsURL.Text = Me.ItemsListView.GetTextSelected(5)
        Enable Me.SelectedCompetitorsURL
        currentItemLVIndex = CLng(Me.ItemsListView.SelIndex)
    End If
    stopChangeEvent = False
End Sub

Private Sub SelectedQuantityUpDown_ValueChanged()
    If stopChangeEvent = False Then
        Me.ItemsListView.Edit Me.SelectedQuantityUpDown.value, currentItemLVIndex, 2
    End If
End Sub

Private Sub SelectedPriceUpDown_ValueChanged()
    If stopChangeEvent = False Then
        Me.ItemsListView.Edit Format(Me.SelectedPriceUpDown.value, "$0.00"), currentItemLVIndex, 3
    End If
End Sub

Private Sub SelectedPriceBestCheckbox_Click()
    If stopChangeEvent = False Then
        stopChangeEvent = True
        If Me.SelectedPriceBestCheckbox Then
            Me.SelectedPriceUpDown.value = 0
            Disable Me.SelectedPriceUpDown
            Me.ItemsListView.Edit "<best>", currentItemLVIndex, 3
        Else
            Me.SelectedPriceUpDown.value = Format(itemDB.item(CStr(Me.SelectedItemField.Text)).item("DefaultPrice"), "0.00")
            Enable Me.SelectedPriceUpDown
            Me.ItemsListView.Edit Format(Me.SelectedPriceUpDown.value, "$0.00"), currentItemLVIndex, 3
        End If
        stopChangeEvent = False
    End If
End Sub

Private Sub SelectedCompetitorsPriceUpDown_ValueChanged()
    If stopChangeEvent = False Then
        If Me.SelectedCompetitorsPriceUpDown.value = 0 Then
            Me.ItemsListView.Edit "", currentItemLVIndex, 4
        Else
            Me.ItemsListView.Edit Format(Me.SelectedCompetitorsPriceUpDown.value, "$0.00"), currentItemLVIndex, 4
        End If
    End If
End Sub

Private Sub SelectedCompetitorsURL_Change()
    If stopChangeEvent = False Then
        Me.ItemsListView.Edit Me.SelectedCompetitorsURL, currentItemLVIndex, 5
    End If
End Sub

Private Sub SubmitButton_Click()
    Dim warnings As String
    warnings = completenessCheck()
    If warnings <> "" Then
        MsgBox warnings
        Exit Sub
    End If
    
    Mouse.Hourglass True
    
    Dim excelapp As Excel.Application
    Set excelapp = New Excel.Application
    
    Dim i As Long
    
    excelapp.Workbooks.Add
    Dim currRowIndex As Long
    currRowIndex = 1
    If Me.CustomerNameField.Text <> "" Then
        excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = Me.CustomerNameField.Text
        currRowIndex = currRowIndex + 1
    End If
    If Me.CustomerCompanyField.Text <> "" Then
        excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = Me.CustomerCompanyField.Text
        currRowIndex = currRowIndex + 1
    End If
    If Me.CustomerAddressField.Text <> "" Then
        excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = Me.CustomerAddressField.Text
        currRowIndex = currRowIndex + 1
    End If
    excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = IIf(Me.CustomerCityField.Text = "", "", Me.CustomerCityField.Text & ", ") & Me.CustomerStateField.Text & " " & Me.CustomerZipCodeField.Text
    currRowIndex = currRowIndex + 1
    If Me.CustomerPhoneNumberField.Text <> "" Then
        excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = Me.CustomerPhoneNumberField.Text
        currRowIndex = currRowIndex + 1
    End If
    If Me.CustomerEmailField.Text <> "" Then
        excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = Me.CustomerEmailField.Text
        currRowIndex = currRowIndex + 1
    End If
    
    currRowIndex = currRowIndex + 1
    
    Select Case True
        Case Is = Me.ReturningCustomerOption(0).value
            excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = "Returning customer: Yes"
            excelapp.ActiveWorkbook.ActiveSheet.Range("B" & currRowIndex).value = Me.CustomerNumberField.Text
            currRowIndex = currRowIndex + 1
        Case Is = Me.ReturningCustomerOption(1).value
            excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = "Returning customer: No"
            currRowIndex = currRowIndex + 1
    End Select
    
    Select Case True
        Case Is = Me.OrderExistsInMasOption(0).value
            excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = "Order already in Mas: Yes"
            excelapp.ActiveWorkbook.ActiveSheet.Range("B" & currRowIndex).value = Me.OrderNumberField.Text
            currRowIndex = currRowIndex + 1
        Case Is = Me.OrderExistsInMasOption(1).value
            excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = "Order already in Mas: No"
            currRowIndex = currRowIndex + 1
    End Select
    
    excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = "Expected payment method: " & Me.PaymentMethodComboBox.Text
    currRowIndex = currRowIndex + 1
    
    excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = "Quote came from: " & Me.QuoteFromComboBox.Text
    currRowIndex = currRowIndex + 1
    
    If Me.CommentsField.Text <> "" Then
        currRowIndex = currRowIndex + 1
        Dim comm As String
        comm = Me.CommentsField.Text
        While CBool(InStr(comm, vbCrLf & vbCrLf))
            comm = Replace(comm, vbCrLf & vbCrLf, vbCrLf)
        Wend
        Dim commA As Variant
        commA = Split(comm, vbCrLf)
        For i = 0 To UBound(commA)
            excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = commA(i)
            currRowIndex = currRowIndex + 1
        Next i
    End If
    
    currRowIndex = currRowIndex + 1
    
    Dim columnIndex As Long
    Dim columnIndexForQty As Long
    Dim columnIndexForCostCalc As Long
    Dim columnIndicesForTotals As PerlArray
    Set columnIndicesForTotals = New PerlArray
    columnIndex = 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Item"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Description"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Dropship"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "QOH"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Req Qty"
    columnIndexForQty = columnIndex
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Std Cost"
    columnIndexForCostCalc = columnIndex
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Std Cost Extension"
    columnIndicesForTotals.Push columnIndex
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Avg Cost"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Avg Cost Extension"
    columnIndicesForTotals.Push columnIndex
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "New Cost"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "New Cost Extension"
    columnIndicesForTotals.Push columnIndex
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "PO Comment"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Store Price"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Store Price Extension"
    columnIndicesForTotals.Push columnIndex
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Store Price Profit Pct"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Web Price"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Web Price Extension"
    columnIndicesForTotals.Push columnIndex
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Web Price Profit Pct"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Target Price"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Target Price Extension"
    columnIndicesForTotals.Push columnIndex
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Target Price Profit Pct"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Competitors Price"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Competitors URL"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Sales This Year"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Sales Last Year"
    columnIndex = columnIndex + 1
    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "Notes"
    columnIndex = columnIndex + 1
    
    currRowIndex = currRowIndex + 1
    
    Dim itemHash As Dictionary
    Set itemHash = New Dictionary
    Dim firstItemRowIndex As Long, lastItemRowIndex As Long
    firstItemRowIndex = currRowIndex
    For i = 0 To Me.ItemsListView.ListCount - 1
        columnIndex = 1
        
        Dim item As String
        item = Me.ItemsListView.GetText(i, 0)
    
        itemHash.Add CStr(item), CLng(Me.ItemsListView.GetText(i, 2))
        
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT vActualWhseQty.QtyOnHand, InventoryMaster.StdCost, InventoryMaster.AvgCost, InventoryMaster.TNC, InventoryMaster.DiscountMarkupPriceRate1, InventoryMaster.IDiscountMarkupPriceRate1, InventoryMaster.POComment, vItemStatusBreakdown.DSable FROM InventoryMaster INNER JOIN vActualWhseQty ON InventoryMaster.ItemNumber=vActualWhseQty.ItemNumber INNER JOIN vItemStatusBreakdown ON InventoryMaster.ItemNumber=vItemStatusBreakdown.ItemNumber WHERE InventoryMaster.ItemNumber='" & item & "'")
        
        Dim reqPrice As Variant
        reqPrice = Me.ItemsListView.GetText(i, 3)
        
        ' item
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = item
        columnIndex = columnIndex + 1
        ' description
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = Me.ItemsListView.GetText(i, 1)
        columnIndex = columnIndex + 1
        ' dropship
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = IIf(rst("DSable"), "Y", "N")
        columnIndex = columnIndex + 1
        ' quantity on hand
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = rst("QtyOnHand")
        columnIndex = columnIndex + 1
        ' requested quantity
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = Me.ItemsListView.GetText(i, 2)
        columnIndex = columnIndex + 1
        ' standard cost
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = rst("StdCost")
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "$#,##0.00"
        columnIndex = columnIndex + 1
        ' standard cost extension
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).Formula = "=" & ColumnIndexToLetter(columnIndexForQty) & currRowIndex & "*" & ColumnIndexToLetter(columnIndex - 1) & currRowIndex
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "$#,##0.00"
        columnIndex = columnIndex + 1
        ' average cost
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = rst("AvgCost")
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "$#,##0.00"
        columnIndex = columnIndex + 1
        ' average cost extension
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).Formula = "=" & ColumnIndexToLetter(columnIndexForQty) & currRowIndex & "*" & ColumnIndexToLetter(columnIndex - 1) & currRowIndex
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "$#,##0.00"
        columnIndex = columnIndex + 1
        ' new cost
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = rst("TNC")
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "$#,##0.00"
        columnIndex = columnIndex + 1
        ' new cost extension
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).Formula = "=" & ColumnIndexToLetter(columnIndexForQty) & currRowIndex & "*" & ColumnIndexToLetter(columnIndex - 1) & currRowIndex
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "$#,##0.00"
        columnIndex = columnIndex + 1
        ' po comment
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = Nz(rst("POComment"))
        columnIndex = columnIndex + 1
        ' store price
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = rst("DiscountMarkupPriceRate1")
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "$#,##0.00"
        columnIndex = columnIndex + 1
        ' store price extension
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).Formula = "=" & ColumnIndexToLetter(columnIndexForQty) & currRowIndex & "*" & ColumnIndexToLetter(columnIndex - 1) & currRowIndex
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "$#,##0.00"
        columnIndex = columnIndex + 1
        ' store price profit percentage
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).Formula = "=" & ColumnIndexToLetter(columnIndex - 2) & currRowIndex & "/" & ColumnIndexToLetter(columnIndexForCostCalc) & currRowIndex & "-1"
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "0.0%"
        columnIndex = columnIndex + 1
        ' web price
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = rst("IDiscountMarkupPriceRate1")
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "$#,##0.00"
        columnIndex = columnIndex + 1
        ' web price extension
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).Formula = "=" & ColumnIndexToLetter(columnIndexForQty) & currRowIndex & "*" & ColumnIndexToLetter(columnIndex - 1) & currRowIndex
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "$#,##0.00"
        columnIndex = columnIndex + 1
        ' web price profit percentage
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).Formula = "=" & ColumnIndexToLetter(columnIndex - 2) & currRowIndex & "/" & ColumnIndexToLetter(columnIndexForCostCalc) & currRowIndex & "-1"
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "0.0%"
        columnIndex = columnIndex + 1
        ' target price
        If reqPrice = "<best>" Then
            excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = ""
        Else
            excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = Replace(reqPrice, "$", "")
            excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "$#,##0.00"
        End If
        columnIndex = columnIndex + 1
        ' target price extension
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).Formula = "=" & ColumnIndexToLetter(columnIndexForQty) & currRowIndex & "*" & ColumnIndexToLetter(columnIndex - 1) & currRowIndex
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "$#,##0.00"
        columnIndex = columnIndex + 1
        ' target price profit percentage
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).Formula = "=" & ColumnIndexToLetter(columnIndex - 2) & currRowIndex & "/" & ColumnIndexToLetter(columnIndexForCostCalc) & currRowIndex & "-1"
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "0.0%"
        columnIndex = columnIndex + 1
        ' competitor's price
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = Replace(Me.ItemsListView.GetText(i, 4), "$", "")
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "$#,##0.00"
        columnIndex = columnIndex + 1
        ' competitor's url
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = Me.ItemsListView.GetText(i, 5)
        columnIndex = columnIndex + 1
        
        rst.Close
        Set rst = DB.retrieve("SELECT Period, P2SalesAmount + SOSalesAmount AS TotalAmount FROM InventoryStatistics WHERE PeriodType=0 AND WhseCode='000' AND ItemNumber='" & item & "' AND Period IN (4,5) ORDER BY Period")
        ' sales this year
        ' sales last year
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = "0"
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex + 1) & currRowIndex).value = "0"
        While Not rst.EOF
            If rst("Period") = 4 Then
                excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = rst("TotalAmount")
            ElseIf rst("Period") = 5 Then
                excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex + 1) & currRowIndex).value = rst("TotalAmount")
            End If
            rst.MoveNext
        Wend
        rst.Close
        columnIndex = columnIndex + 2
        
        ' notes
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).value = ""
        columnIndex = columnIndex + 1
        
        lastItemRowIndex = currRowIndex
        
        currRowIndex = currRowIndex + 1
    Next i
    
    currRowIndex = currRowIndex + 1
    
    While columnIndicesForTotals.Scalar <> 0
        columnIndex = CLng(columnIndicesForTotals.Shift)
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).Formula = "=SUM(" & ColumnIndexToLetter(columnIndex) & firstItemRowIndex & ":" & ColumnIndexToLetter(columnIndex) & lastItemRowIndex & ")"
        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(columnIndex) & currRowIndex).NumberFormat = "$#,##0.00"
    Wend
    
    Dim shipCost As Dictionary
    Set shipCost = ShippingDB.EstimateFreightForShipment(itemHash, Me.CustomerZipCodeField.Text, Me.CustomerStateField.Text)
    If shipCost Is Nothing Then
        excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = "Error trying to estimate shipping costs!"
    Else
        If shipCost.exists("UPS Ground") Then
            excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = "Estimated UPS ground shipping cost: $" & shipCost.item("UPS Ground")
        ElseIf shipCost.exists("USPS Priority Mail") Then
            excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = "Estimated USPS priority shipping cost: $" & shipCost.item("USPS Priority Mail")
        Else
            excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = "Error trying to estimate shipping costs!"
        End If
    End If
    currRowIndex = currRowIndex + 1
    
    excelapp.Visible = True
    
    Mouse.Hourglass False
    
    'clearForm
End Sub




Private Function completenessCheck() As String
    Dim warnings As PerlArray
    Set warnings = New PerlArray
    
    If Me.CustomerNameField.Text = "" And Me.CustomerCompanyField.Text = "" Then
        warnings.Push "Please fill out either the Customer Name or Company fields."
    End If
    If Me.CustomerPhoneNumberField.Text = "" And Me.CustomerEmailField.Text = "" Then
        warnings.Push "Please fill out either the Customer Phone Number or Email fields."
    End If
    If Me.CustomerStateField.Text = "" Then
        warnings.Push "Please fill out the Customer State field."
    End If
    If Me.CustomerZipCodeField.Text = "" Then
        warnings.Push "Please fill out the Customer Zip Code field."
    End If
    If Me.ItemsListView.ListCount = 0 Then
        warnings.Push "You must have at least one item in the quote."
    End If
    
    If warnings.Scalar = 0 Then
        completenessCheck = ""
    Else
        completenessCheck = warnings.Join(vbCrLf)
    End If
End Function

Private Function clearForm() As String
    Me.CustomerNameField.Text = ""
    Me.CustomerCompanyField.Text = ""
    Me.CustomerAddressField.Text = ""
    Me.CustomerCityField.Text = ""
    Me.CustomerPhoneNumberField.Text = ""
    Me.CustomerEmailField.Text = ""
    Me.CustomerStateField.Text = ""
    Me.CustomerZipCodeField.Text = ""
    Me.ReturningCustomerOption(1).value = True
    Me.OrderExistsInMasOption(1).value = True
    Me.CommentsField.Text = ""
    Me.AllItemsFilterField.Text = ""
    lastFilter = ""
    Me.AllItemsList.Clear
    Me.ItemsListView.Clear
End Function
    
    
    
    
    
    'Dim emailBody As String
    'emailBody = Me.CustomerNameField.Text & vbCrLf & Me.CustomerStateField.Text & ", " & Me.CustomerZipCodeField.Text & vbCrLf
    'If Me.CustomerPhoneNumberField.Text <> "" Then
    '    emailBody = emailBody & Me.CustomerPhoneNumberField.Text & vbCrLf
    'End If
    'If Me.CustomerEmailField.Text <> "" Then
    '    emailBody = emailBody & Me.CustomerEmailField.Text & vbCrLf
    'End If
    'emailBody = emailBody & vbCrLf
    
    'emailBody = emailBody & "Returning customer: "
    'Select Case True
    '    Case Is = Me.ReturningCustomerOption(0).value
    '        emailBody = emailBody & "Yes" & vbCrLf & "    Customer No: " & UCase(Me.CustomerNumberField.Text)
    '    Case Is = Me.ReturningCustomerOption(1).value
    '        emailBody = emailBody & "No"
    'End Select
    'emailBody = emailBody & vbCrLf
    '
    'emailBody = emailBody & "Order already in Mas: "
    'Select Case True
    '    Case Is = Me.OrderExistsInMasOption(0).value
    '        emailBody = emailBody & "Yes" & vbCrLf & "    Order No: " & UCase(Me.OrderNumberField.Text)
    '    Case Is = Me.OrderExistsInMasOption(1).value
    '        emailBody = emailBody & "No"
    'End Select
    'emailBody = emailBody & vbCrLf
    '
    'If Me.CommentsField.Text <> "" Then
    '    emailBody = emailBody & "Additional comments:" & vbCrLf & IndentTextParagraph(Me.CommentsField.Text)
    'End If
    'emailBody = emailBody & vbCrLf
    
    'emailBody = emailBody & vbCrLf
    'Dim i As Long
    'Dim totalCOGS As Variant
    'totalCOGS = CDec(0)
    'Dim totalPOGS As Variant
    'totalPOGS = CDec(0)
    'Dim itemHash As Dictionary
    'Set itemHash = New Dictionary
    'Dim foundBestPriceFlag As Boolean
    'foundBestPriceFlag = False
    'For i = 0 To Me.ItemsListView.ListCount - 1
    '    Dim item As String
    '    item = Me.ItemsListView.GetText(i, 0)
    '
    '    itemHash.Add CStr(item), CLng(Me.ItemsListView.GetText(i, 2))
    '
    '    Dim itemStanza As String
    '    itemStanza = item & " - " & Me.ItemsListView.GetText(i, 1)
    '
    '    Dim rst As ADODB.Recordset
    '    Set rst = DB.retrieve("SELECT vActualWhseQty.QtyOnHand, InventoryMaster.StdCost, InventoryMaster.AvgCost, InventoryMaster.TNC, InventoryMaster.IDiscountMarkupPriceRate1, vItemStatusBreakdown.DSable FROM InventoryMaster INNER JOIN vActualWhseQty ON InventoryMaster.ItemNumber=vActualWhseQty.ItemNumber INNER JOIN vItemStatusBreakdown ON InventoryMaster.ItemNumber=vItemStatusBreakdown.ItemNumber WHERE InventoryMaster.ItemNumber='" & item & "'")
    '    If rst("DSable") Then
    '        itemStanza = itemStanza & " - D/S-able" & vbCrLf
    '    Else
    '        itemStanza = itemStanza & " - NOT D/S-able!" & vbCrLf
    '    End If
    '
    '    itemStanza = itemStanza & "   Current Qty:   " & rst("QtyOnHand") & vbCrLf
    '    itemStanza = itemStanza & "   Requested Qty: " & Me.ItemsListView.GetText(i, 2) & vbCrLf
    '    itemStanza = itemStanza & vbCrLf
    '
    '    itemStanza = itemStanza & "   Std Cost: " & Format(rst("StdCost"), "$0.00") & vbCrLf
    '    If rst("TNC") <> 0 Then
    '        itemStanza = itemStanza & "   New Cost: " & Format(rst("TNC"), "$0.00") & vbCrLf
    '    End If
    '    If Nz(rst("AvgCost")) <> 0 Then
    '        itemStanza = itemStanza & "   Avg Cost: " & Format(rst("AvgCost"), "$0.00") & vbCrLf
    '    End If
    '    totalCOGS = CDec(totalCOGS + CDec(rst("StdCost") * CLng(Me.ItemsListView.GetText(i, 2))))
    '    itemStanza = itemStanza & vbCrLf
    '
    '    itemStanza = itemStanza & "   Web Price:       " & Format(rst("IDiscountMarkupPriceRate1"), "$0.00") & vbCrLf
    '    Dim reqPrice As Variant
    '    reqPrice = Me.ItemsListView.GetText(i, 3)
    '    If reqPrice = "<best>" Then
    '        foundBestPriceFlag = True
    '        itemStanza = itemStanza & "   Requested Price: best price we can" & vbCrLf
    '    Else
    '        reqPrice = CDec(Replace(reqPrice, "$", ""))
    '        itemStanza = itemStanza & "   Requested Price: " & Format(reqPrice, "$0.00") & " (" & Format(rst("IDiscountMarkupPriceRate1") - reqPrice, "$0.00") & " discount)" & vbCrLf
    '        totalPOGS = CDec(totalPOGS + reqPrice * CLng(Me.ItemsListView.GetText(i, 2)))
    '    End If
    '    itemStanza = itemStanza & vbCrLf
    '    rst.Close
    '
    '    itemStanza = itemStanza & "   Sales             P2     SO     D/S    Returns" & vbCrLf
    '    Set rst = DB.retrieve("SELECT Period, P2SalesAmount, SOSalesAmount, ReturnsAmount, DropshipsAmount FROM InventoryStatistics WHERE PeriodType=0 AND WhseCode='000' AND ItemNumber='" & item & "' ORDER BY Period")
    '    Dim salesArray(5, 3) As Long
    '    Dim j As Long, k As Long
    '    For j = 0 To UBound(salesArray, 1)
    '        For k = 0 To UBound(salesArray, 2)
    '            salesArray(j, k) = 0
    '        Next k
    '    Next j
    '    While Not rst.EOF
    '        salesArray(CLng(rst("Period")), 0) = CLng(rst("P2SalesAmount"))
    '        salesArray(CLng(rst("Period")), 1) = CLng(rst("SOSalesAmount"))
    '        salesArray(CLng(rst("Period")), 2) = CLng(rst("DropshipsAmount"))
    '        salesArray(CLng(rst("Period")), 3) = CLng(rst("ReturnsAmount"))
    '        rst.MoveNext
    '    Wend
    '    rst.Close
    '    For j = 0 To UBound(salesArray, 1)
    '        Dim line As String
    '        line = "      "
    '        Select Case j
    '            Case Is = 0
    '                line = line & "This Month:"
    '            Case Is = 1
    '                line = line & "Last Month:"
    '            Case Is = 2
    '                line = line & "Current Per:"
    '            Case Is = 3
    '                line = line & "Last Per:"
    '            Case Is = 4
    '                line = line & "YTD:"
    '            Case Is = 5
    '                line = line & "Last Year:"
    '        End Select
    '
    '        line = line & Space(23 - Len(line) - Len(CStr(salesArray(j, 0)))) & salesArray(j, 0)
    '        line = line & Space(30 - Len(line) - Len(CStr(salesArray(j, 1)))) & salesArray(j, 1)
    '        line = line & Space(38 - Len(line) - Len(CStr(salesArray(j, 2)))) & salesArray(j, 2)
    '        line = line & Space(45 - Len(line) - Len(CStr(salesArray(j, 3)))) & salesArray(j, 3)
    '
    '        itemStanza = itemStanza & line & vbCrLf
    '    Next j
    '    itemStanza = itemStanza & vbCrLf
    '
    '    emailBody = emailBody & itemStanza
    'Next i
    
    'emailBody = emailBody & "Total Cost of Goods (using std cost):    " & Format(totalCOGS, "$0.00") & vbCrLf
    'If foundBestPriceFlag Then
    '    emailBody = emailBody & "Total Price of Goods (using req. price): can't calculate, customer asked for best price" & vbCrLf
    'Else
    '    emailBody = emailBody & "Total Price of Goods (using req. price): " & Format(totalPOGS, "$0.00") & vbCrLf
    'End If
    '
    'Dim shipCost As Dictionary
    'Set shipCost = ShippingDB.EstimateFreightForShipment(itemHash, Me.CustomerZipCodeField.Text, Me.CustomerStateField.Text)
    'If shipCost.exists("UPS Ground") Then
    '    emailBody = emailBody & "Estimated UPS ground shipping cost:      " & Format(shipCost.item("UPS Ground"), "$0.00") & vbCrLf
    'ElseIf shipCost.exists("USPS Priority Mail") Then
    '    emailBody = emailBody & "Estimated USPS priority shipping cost:   " & Format(shipCost.item("USPS Priority Mail"), "$0.00") & vbCrLf
    'Else
    '    emailBody = emailBody & "Error trying to estimate shipping costs!" & vbCrLf
    'End If
    
    'excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = "Total Cost of Goods"
    'excelapp.ActiveWorkbook.ActiveSheet.Range("B" & currRowIndex).value = "=SUMPRODUCT(E" & firstItemRowIndex & ":E" & lastItemRowIndex & ",F" & firstItemRowIndex & ":F" & lastItemRowIndex & ")"
    'currRowIndex = currRowIndex + 1
    '
    'excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = "Total Normal Price of Goods"
    'excelapp.ActiveWorkbook.ActiveSheet.Range("B" & currRowIndex).value = "=SUMPRODUCT(E" & firstItemRowIndex & ":E" & lastItemRowIndex & ",J" & firstItemRowIndex & ":J" & lastItemRowIndex & ")"
    'currRowIndex = currRowIndex + 1
    '
    'excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRowIndex).value = "Total Target Price of Goods"
    'excelapp.ActiveWorkbook.ActiveSheet.Range("B" & currRowIndex).value = "=SUMPRODUCT(E" & firstItemRowIndex & ":E" & lastItemRowIndex & ",K" & firstItemRowIndex & ":K" & lastItemRowIndex & ")"
    'currRowIndex = currRowIndex + 1
    
    'Debug.Print emailBody
    'MsgBox emailBody
    
    'Dim emailSent As Boolean
    'emailSent = OutlookIntegration.OpenEmailTo("brian@tools-plus.com", "Quote for " & Me.CustomerNameField.Text, emailBody)
    'If emailSent = False Then
    '    MsgBox "Couldn't open Outlook to send email! Trying to send immediately!"
    '    emailSent = Utilities.SendEmailTo("brian@tools-plus.com", "Quote for " & Me.CustomerNameField.Text, emailBody)
    'End If
    'If emailSent = False Then
    '    MsgBox "Email send failed! I have no way to send emails from your computer! Tell Brian!"
    'End If
    '
    'If emailSent Then
    '    If vbYes = MsgBox("Clear form?", vbYesNo) Then
    '        clearForm
    '    End If
    'End If
    
    'we don't really have a way of verifying that the sendemailto absolutely works, and i
    'don't want to lose a quote, so it's outlook or nothing. if outlook fails to open,
    'we an open a notepad window that they can copy from and paste into outlook.
    'Dim emailSent As Boolean
    'emailSent = OutlookIntegration.OpenEmailTo("eric@tools-plus.com", "Quote for " & Me.CustomerNameField.Text, emailBody)
    'If emailSent = False Then
    '    MsgBox "Couldn't open Outlook to send email! I'm going to open a notepad window that you can just copy from and paste into an email."
    '    Open DESTINATION_DIR & "quote-email.txt" For Output As #1
    '    Print #1, emailBody
    '    Close #1
    '    OpenDefaultApp DESTINATION_DIR & "quote-email.txt"
    'End If
