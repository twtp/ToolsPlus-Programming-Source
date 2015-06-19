VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form DynamicPricingFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dynamic Pricing"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   1575
      Left            =   0
      TabIndex        =   18
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2778
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox reserveMinTxt 
      Height          =   285
      Left            =   1560
      TabIndex        =   17
      Top             =   3100
      Width           =   855
   End
   Begin VB.ListBox marketLst 
      Height          =   1035
      Left            =   20
      TabIndex        =   16
      Top             =   2040
      Width           =   2655
   End
   Begin VB.ListBox sitesLst 
      Height          =   1035
      Left            =   2760
      TabIndex        =   15
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "P"
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton applyBtn 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   445
      Left            =   4200
      TabIndex        =   9
      Top             =   3075
      Width           =   1215
   End
   Begin VB.CheckBox dynamicJetChk 
      Caption         =   "Dynamic Jet.com Price"
      Height          =   255
      Left            =   7080
      TabIndex        =   7
      Top             =   3255
      Width           =   2655
   End
   Begin VB.CheckBox dynamicSiteChk 
      Caption         =   "Dynamic TP-Site Price"
      Height          =   255
      Left            =   7080
      TabIndex        =   5
      Top             =   2655
      Width           =   2655
   End
   Begin VB.CheckBox dynamicStoreChk 
      Caption         =   "Dynamic Store Price"
      Height          =   255
      Left            =   7080
      TabIndex        =   3
      Top             =   2100
      Width           =   2655
   End
   Begin VB.CheckBox dynamicEBayChk 
      Caption         =   "Dynamic EBay Price"
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   1500
      Width           =   2655
   End
   Begin VB.Label marketplacesLbl 
      Caption         =   "Marketplaces:"
      Height          =   255
      Left            =   15
      TabIndex        =   14
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label sitesLbl 
      Caption         =   "Sites:"
      Height          =   255
      Left            =   2775
      TabIndex        =   13
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lastCheckedLbl 
      Caption         =   "Checked:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3370
      Width           =   3255
   End
   Begin VB.Label reserveMinLbl 
      Caption         =   "Reserve Min: $"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3100
      Width           =   2655
   End
   Begin VB.Label jetDotComPriceLbl 
      Caption         =   "Current Jet.com Price: $"
      Height          =   255
      Left            =   6960
      TabIndex        =   8
      Top             =   3015
      Width           =   2655
   End
   Begin VB.Label tpSitePriceLbl 
      Caption         =   "Current TP-Site Price: $"
      Height          =   255
      Left            =   6960
      TabIndex        =   6
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label storePriceLbl 
      Caption         =   "Current Store Price: $"
      Height          =   255
      Left            =   6960
      TabIndex        =   4
      Top             =   1845
      Width           =   2655
   End
   Begin VB.Label dynamicPriceLbl 
      Caption         =   "Dynamic Price: $"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label itemNumberLbl 
      Alignment       =   2  'Center
      Caption         =   "ItemNumber"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "DynamicPricingFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public currentMinPrice As String
Public CurrentDynamicPrice As String

Private Sub applyBtn_Click()
SetDynamicPricingDetails
'now reload the maintenance form...
InventoryMaintenance.RequeryPoInvInPlace
End Sub

Private Sub Command1_Click()
 ClipBoard_SetImage Utilities.CaptureVB6Form(Me)
    MsgBox "Screenshot captured!"
End Sub

Private Sub dynamicEBayChk_Click()
    applyBtn.Enabled = True
End Sub

Private Sub dynamicJetChk_Click()
    applyBtn.Enabled = True
End Sub

Private Sub dynamicSiteChk_Click()
    applyBtn.Enabled = True
End Sub

Private Sub dynamicStoreChk_Click()
    applyBtn.Enabled = True
End Sub

Private Sub SetDynamicPricingDetails()
For Each item In ListView1.ListItems
    Select Case item.SubItems(1)
    Case "TP Site"
        If item.Checked = True Then
        
        Else
        
        End If
    Case "TP Store"
        If item.Checked = True Then
        
        Else
        
        End If
    Case "EBay"
        If item.Checked = True Then
        
        Else
        
        End If
    Case "JetDotCom"
        If item.Checked = True Then
             DB.Execute "UPDATE PriceWiserItems SET JetDotCom=1 WHERE ItemNumber='" & Me.itemNumberLbl.Caption & "'"
             DB.Execute "UPDATE JetOffload SET OffloadRequired=1, Price=" & item.SubItems(3) & " WHERE ItemNumber='" & Me.itemNumberLbl.Caption & "'"
        Else
             DB.Execute "UPDATE PriceWiserItems SET JetDotCom=0 WHERE ItemNumber='" & Me.itemNumberLbl.Caption & "'"
             DB.Execute "UPDATE JetOffload SET OffloadRequired=1 WHERE ItemNumber='" & Me.itemNumberLbl.Caption & "'"
        End If
        
       
    End Select
Next

Dim updateMinPrice As ADODB.Recordset
Set updateMinPrice = DB.retrieve("SELECT MinPrice,MaxPrice FROM PriceWiserItems WHERE ItemNumber='" & Me.itemNumberLbl.Caption & "'")
If updateMinPrice.RecordCount > 0 Then
    If Format(updateMinPrice("MinPrice"), "Currency") <> Format(Me.reserveMinTxt.Text, "Currency") Then DB.Execute ("UPDATE PriceWiserItems SET MinPrice=" & Me.reserveMinTxt.Text & " WHERE ItemNumber='" & Me.itemNumberLbl.Caption & "'")
End If

'If Format(Me.reserveMinTxt.Text, "Currency") > Format(Split(Me.dynamicPriceLbl.Caption, "$")(1), "Currency") Then
 'MsgBox "this item's reservemin is > than its dynamic price! should be raising dynamic price to reservemin..."
 
'End If
End Sub



Private Sub Form_Load()

'refreshForm

SetupColumns
Me.ListView1.LabelEdit = lvwManual
End Sub
Public Sub SetupColumns()
DynamicPricingFrm.ListView1.ListItems.Clear
DynamicPricingFrm.ListView1.Refresh

'ListView1.ColumnHeaders.Add 0, , "Dynamic", ListView1.width / 5, vbCenter
ListView1.ColumnHeaders.Clear
ListView1.ListItems.Clear
ListView1.Refresh
ListView1.ColumnHeaders.Add , , "Dynamic", ListView1.width / 7
ListView1.ColumnHeaders.Add , , "Store", ListView1.width / 5
ListView1.ColumnHeaders.Add , , "Curr. Price", ListView1.width / 6
ListView1.ColumnHeaders.Add , , "New Price", ListView1.width / 6

ListView1.HideColumnHeaders = False
ListView1.View = lvwReport

ListView1.CheckBoxes = True


'Dim EBayRow As ListItem
'Set EBayRow = ListView1.ListItems.Add()
'EBayRow.Text = ""
'EBayRow.SubItems(1) = "EBay"
'EBayRow.SubItems(2) = InventoryMaintenance.EBayPrice.Text

'Dim TPSite As ListItem
'Set TPSite = ListView1.ListItems.Add()
'TPSite.Text = ""
'TPSite.SubItems(1) = "TP Site"
'TPSite.SubItems(2) = InventoryMaintenance.IDiscountMarkupPriceRate(1)

'Dim TPStore As ListItem
'Set TPStore = ListView1.ListItems.Add()
'TPStore.Text = ""
'TPStore.SubItems(1) = "TP Store"
'TPStore.SubItems(2) = InventoryMaintenance.DiscountMarkupPriceRate(1)

'Dim JetRow As ListItem
'Set JetRow = ListView1.ListItems.Add()
'JetRow.Text = ""
'JetRow.SubItems(1) = "Jet.com"
'JetRow.SubItems(2) = InventoryMaintenance.JetPrice.Text
End Sub
Public Sub DisableFormControls()
Me.reserveMinTxt.Enabled = False
Me.ListView1.Enabled = False
Me.dynamicEBayChk.Enabled = False
Me.dynamicJetChk.Enabled = False
Me.dynamicSiteChk.Enabled = False
Me.dynamicStoreChk.Enabled = False
End Sub
Public Sub EnableFormControls()
Me.reserveMinTxt.Enabled = True
Me.ListView1.Enabled = True
Me.dynamicEBayChk.Enabled = True
Me.dynamicSiteChk.Enabled = True
Me.dynamicJetChk.Enabled = True
Me.dynamicStoreChk.Enabled = True
End Sub
Public Sub SetCurrentDynamicPrice(price As Currency)
CurrentDynamicPrice = price
End Sub
Public Sub refreshForm()

'ListView1.ListItems.Add , , "test1"
'ListView1.ListItems.Add , , "test2"



'Dim EBayRow As ListItem
'Set EBayRow = ListView1.ListItems.Add()
'EBayRow.Text = ""
'EBayRow.SubItems(1) = "EBay"
'EBayRow.SubItems(2) = InventoryMaintenance.EBayPrice.Text

'Dim TPSite As ListItem
'Set TPSite = ListView1.ListItems.Add()
'TPSite.Text = ""
'TPSite.SubItems(1) = "TP Site"
'TPSite.SubItems(2) = InventoryMaintenance.IDiscountMarkupPriceRate(1)

'Dim TPStore As ListItem
'Set TPStore = ListView1.ListItems.Add()
'TPStore.Text = ""
'TPStore.SubItems(1) = "TP Store"
'TPStore.SubItems(2) = InventoryMaintenance.DiscountMarkupPriceRate(1)

'Dim JetRow As ListItem
'Set JetRow = ListView1.ListItems.Add()
'JetRow.Text = ""
'JetRow.SubItems(1) = "Jet.com"
'JetRow.SubItems(2) = InventoryMaintenance.JetPrice.Text

End Sub


Private Sub ListView1_ItemCheck(ByVal item As MSComctlLib.ListItem)
applyBtn.Enabled = True
If item.Checked = True Then
    item.SubItems(3) = "$" & Split(Me.dynamicPriceLbl.Caption, "$")(1)
Else
    'now need to grab regular price of item....
    If item.SubItems(1) = "JetDotCom" Then
        Dim jetRes As ADODB.Recordset
        Set jetRes = DB.retrieve("SELECT Price FROM JetOffload WHERE ItemNumber='" & Me.itemNumberLbl.Caption & "'")
        If jetRes.RecordCount > 0 Then item.SubItems(3) = Format(jetRes("Price"), "Currency")
    End If
    If item.SubItems(1) = "TPWeb" Then
    
    End If
    If item.SubItems(1) = "TPStore" Then
    
    End If
    If item.SubItems(1) = "EBay" Then
    
    End If
End If
End Sub

Private Sub reserveMinTxt_Change()
Me.applyBtn.Enabled = True

End Sub

Private Sub reserveMinTxt_LostFocus()
Dim itm As ListItem
If Format(reserveMinTxt.Text, "Currency") > Format(Split(Me.dynamicPriceLbl.Caption, "$")(1), "Currency") Then
    
    For Each itm In ListView1.ListItems
        If itm.Checked = True Then itm.SubItems(3) = Format(reserveMinTxt.Text, "Currency")
    Next
Else
    For Each itm2 In ListView1.ListItems
    If itm2.Checked = True Then
        If Format(Split(Me.dynamicPriceLbl.Caption, "$")(1)) <> Format(itm2.SubItems(3), "Currency") Then
            itm2.SubItems(3) = Format(Split(Me.dynamicPriceLbl.Caption, "$")(1), "Currency")
        End If
    End If
    Next
End If

'If reserveMinTxt.Text <> Me.currentMinPrice Then
'    'save it to db
'    If IsNumeric(reserveMinTxt.Text) = True Then
'        DB.Execute "UPDATE PriceWiserItems SET MinPrice=" & reserveMinTxt.Text & " WHERE ItemNumber='" & Me.itemNumberLbl.Caption & "'"
'        'Now we need to check to see if the new minprice is higher than the max price and/or new price. if so, raise them to the minprice...
'        Dim itemInfo As ADODB.Recordset
'        Set itemInfo = DB.retrieve("SELECT ID,MaxPrice,NewPrice,Enabled,EBay,TPStore,TPWeb,JetDotCom FROM PriceWiserItems WHERE ItemNumber='" & Me.itemNumberLbl.Caption & "'")
'        If itemInfo.RecordCount > 0 Then
'            If CDec(itemInfo("MaxPrice")) < CDec(reserveMinTxt.Text) Then
'                DB.Execute "UPDATE PriceWiserItems SET MaxPrice=" & reserveMinTxt.Text & " WHERE ID=" & itemInfo("ID")
'            End If
'            If CDec(itemInfo("NewPrice")) < CDec(reserveMinTxt.Text) Then
'                DB.Execute "UPDATE PriceWiserItems SET NewPrice=" & reserveMinTxt.Text & " WHERE ID=" & itemInfo("ID")
'            End If
'            If CBool(itemInfo("Enabled")) = True Then
'                If CBool(itemInfo("EBay")) = True Then
'                    'get ebay price... is it lower than min? if so, change it...
'                End If
'                If CBool(itemInfo("TPStore")) = True Then
'                    'get tpstore price.. is it lower than min? if so, change it...
'                End If
'                If CBool(itemInfo("TPWeb")) = True Then
'                    'get tpweb price.. is it lower than min? if so, change it...
'                End If
'                If CBool(itemInfo("JetDotCom")) = True Then
'                    Dim jetinfo As ADODB.Recordset
'                    Set jetinfo = DB.retrieve("SELECT ID,Price FROM JetOffload WHERE ItemNumber='" & Me.itemNumberLbl.Caption & "'")
'                    If jetinfo.RecordCount > 0 Then
'                        If CDec(jetinfo("Price")) < CDec(reserveMinTxt.Text) Then
'                            'ok, change jet.com price and flag it for offload.
'                            DB.Execute "UPDATE JetOffload SET OffloadRequired=1,Price=" & reserveMinTxt.Text & " WHERE ID=" & jetinfo("ID")
'                            'great, now update all the other fields that now contain outdated data in poinv!
'
'                        End If
'                    End If
'                End If
'            End If
'        End If
'    Else
'        MsgBox "The Reserve Price Minimum is not a valid number. Try typing it in again."
'        reserveMinTxt.Text = ""
'        reserveMinTxt.SetFocus
'    End If
'
'End If


End Sub
