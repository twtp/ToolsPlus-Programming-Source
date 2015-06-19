VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#3.5#0"; "TPControls.ocx"
Begin VB.Form KitCreator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kit Creator"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   13125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton submitQtyBtn 
      Caption         =   "Submit Qty"
      Height          =   405
      Left            =   3900
      TabIndex        =   43
      Top             =   6510
      Width           =   1035
   End
   Begin VB.TextBox customQty 
      Height          =   375
      Left            =   2520
      TabIndex        =   42
      Top             =   6510
      Width           =   1365
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "1"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   41
      Top             =   3630
      Width           =   435
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "2"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   6900
      TabIndex        =   40
      Top             =   3630
      Width           =   435
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "3"
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   7320
      TabIndex        =   39
      Top             =   3630
      Width           =   435
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "4"
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   6480
      TabIndex        =   38
      Top             =   3990
      Width           =   435
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "7"
      Enabled         =   0   'False
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   37
      Top             =   4350
      Width           =   435
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "8"
      Enabled         =   0   'False
      Height          =   345
      Index           =   8
      Left            =   6900
      TabIndex        =   36
      Top             =   4380
      Width           =   435
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "9"
      Enabled         =   0   'False
      Height          =   345
      Index           =   9
      Left            =   7350
      TabIndex        =   35
      Top             =   4380
      Width           =   405
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "18"
      Enabled         =   0   'False
      Height          =   375
      Index           =   18
      Left            =   6480
      TabIndex        =   33
      Top             =   5130
      Width           =   435
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "20"
      Enabled         =   0   'False
      Height          =   375
      Index           =   20
      Left            =   6900
      TabIndex        =   32
      Top             =   5130
      Width           =   435
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "24"
      Enabled         =   0   'False
      Height          =   375
      Index           =   24
      Left            =   7350
      TabIndex        =   31
      Top             =   5130
      Width           =   405
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "25"
      Enabled         =   0   'False
      Height          =   375
      Index           =   25
      Left            =   6480
      TabIndex        =   30
      Top             =   5490
      Width           =   435
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "36"
      Enabled         =   0   'False
      Height          =   375
      Index           =   36
      Left            =   6900
      TabIndex        =   29
      Top             =   5490
      Width           =   435
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "48"
      Enabled         =   0   'False
      Height          =   375
      Index           =   48
      Left            =   7350
      TabIndex        =   28
      Top             =   5490
      Width           =   405
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "125"
      Enabled         =   0   'False
      Height          =   375
      Index           =   125
      Left            =   5580
      TabIndex        =   26
      Top             =   6420
      Width           =   615
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "144"
      Enabled         =   0   'False
      Height          =   375
      Index           =   144
      Left            =   6210
      TabIndex        =   25
      Top             =   6420
      Width           =   615
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "256"
      Enabled         =   0   'False
      Height          =   375
      Index           =   256
      Left            =   6840
      TabIndex        =   24
      Top             =   6420
      Width           =   615
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "500"
      Enabled         =   0   'False
      Height          =   375
      Index           =   500
      Left            =   7440
      TabIndex        =   23
      Top             =   6420
      Width           =   615
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "1000"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1000
      Left            =   8070
      TabIndex        =   22
      Top             =   6420
      Width           =   615
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "100"
      Enabled         =   0   'False
      Height          =   375
      Index           =   100
      Left            =   7350
      TabIndex        =   21
      Top             =   5850
      Width           =   405
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "50"
      Enabled         =   0   'False
      Height          =   375
      Index           =   50
      Left            =   6480
      TabIndex        =   20
      Top             =   5850
      Width           =   435
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "10"
      Enabled         =   0   'False
      Height          =   375
      Index           =   10
      Left            =   6480
      TabIndex        =   18
      Top             =   4740
      Width           =   435
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "6"
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   7350
      TabIndex        =   17
      Top             =   3990
      Width           =   405
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "5"
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   6900
      TabIndex        =   16
      Top             =   3990
      Width           =   435
   End
   Begin VB.CommandButton btnMoveDown 
      Caption         =   "\/"
      Enabled         =   0   'False
      Height          =   345
      Left            =   12840
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4230
      Width           =   285
   End
   Begin VB.CommandButton btnMoveUp 
      Caption         =   "/\"
      Enabled         =   0   'False
      Height          =   345
      Left            =   12840
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3840
      Width           =   285
   End
   Begin VB.CommandButton btnSaveKit 
      Caption         =   "Save New Kit"
      Height          =   465
      Left            =   10650
      TabIndex        =   13
      Top             =   6420
      Width           =   2175
   End
   Begin TPControls.SimpleListView NewKitList 
      Height          =   3195
      Left            =   7800
      TabIndex        =   12
      Top             =   3150
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   5636
      MultiSelect     =   0   'False
      Sorted          =   0   'False
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton btnMoveLeft 
      Caption         =   "<="
      Enabled         =   0   'False
      Height          =   465
      Left            =   6450
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3150
      Width           =   525
   End
   Begin VB.CommandButton btnMoveRight 
      Caption         =   "=>"
      Enabled         =   0   'False
      Height          =   465
      Left            =   7230
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3150
      Width           =   525
   End
   Begin VB.Frame Frame 
      Caption         =   "Selected Item"
      Height          =   1785
      Left            =   7230
      TabIndex        =   4
      Top             =   1080
      Width           =   5715
      Begin VB.TextBox SelectedItemSpec3 
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1350
         Width           =   4005
      End
      Begin VB.TextBox SelectedItemSpec2 
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1020
         Width           =   4005
      End
      Begin VB.TextBox SelectedItemSpec1 
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   690
         Width           =   4005
      End
      Begin VB.TextBox SelectedItemNumber 
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   300
         Width           =   4005
      End
      Begin TPControls.Picture SelectedItemPicture 
         Height          =   1575
         Left            =   4200
         TabIndex        =   5
         Top             =   180
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   2778
      End
   End
   Begin TPControls.SimpleListView FullList 
      Height          =   5265
      Left            =   90
      TabIndex        =   3
      Top             =   1080
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   9287
      MultiSelect     =   0   'False
      Sorted          =   -1  'True
      Enabled         =   -1  'True
   End
   Begin VB.ComboBox LineCodeSelector 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   690
      Width           =   1635
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "15"
      Enabled         =   0   'False
      Height          =   375
      Index           =   15
      Left            =   7350
      TabIndex        =   34
      Top             =   4740
      Width           =   405
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "12"
      Enabled         =   0   'False
      Height          =   375
      Index           =   12
      Left            =   6900
      TabIndex        =   19
      Top             =   4740
      Width           =   435
   End
   Begin VB.CommandButton btnQuantity 
      Caption         =   "75"
      Enabled         =   0   'False
      Height          =   375
      Index           =   75
      Left            =   6900
      TabIndex        =   27
      Top             =   5850
      Width           =   435
   End
   Begin VB.Label generalLabel 
      Caption         =   "All Items"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   720
      Width           =   825
   End
   Begin VB.Label generalLabel 
      Caption         =   "Kit Creator"
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
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4125
   End
End
Attribute VB_Name = "KitCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' line code => item number => { name, qoh, spec1, spec2, spec3 }
Private allItemsDB As Dictionary

Public Sub SelectItem(Item As String)
    Dim dx As Long
    dx = InComboPosition(Left(Item, 3), Me.LineCodeSelector)
    If dx = -1 Then
        Exit Sub
    End If
    
    Me.LineCodeSelector.ListIndex = dx
    
    Dim i As Long
    For i = 0 To Me.FullList.ListCount - 1
        If Me.FullList.GetText(i) = Item Then
            Me.FullList.SelectRow i
            Exit For
        End If
    Next i
End Sub

Private Sub Form_Load()
    Me.FullList.SetColumnNames Array("Item", "QOH", "Description")
    Me.FullList.SetColumnWidths Array("1450", "600", "4400")
    Me.NewKitList.SetColumnNames Array("Item", "Quantity", "Description")
    Me.NewKitList.SetColumnWidths Array("1450", "500", "2500")
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ProductLine FROM ProductLine WHERE Hide=0 ORDER BY ProductLine")
    Me.LineCodeSelector.AddItem "<Select>"
    While Not rst.EOF
        Me.LineCodeSelector.AddItem CStr(rst("ProductLine"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    Me.LineCodeSelector.ListIndex = 0
    
    Set allItemsDB = New Dictionary
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set allItemsDB = Nothing
End Sub

Private Sub LineCodeSelector_Click()
    Me.FullList.Clear
    If Me.LineCodeSelector.ListIndex = 0 Then
        'do nothing
    Else
        Dim lc As String
        lc = Me.LineCodeSelector.Text
        If Not allItemsDB.exists(lc) Then
            loadCache lc
        End If
        
        Dim iter As Variant
        For Each iter In allItemsDB.Item(lc).keys
            Me.FullList.Add Array(CStr(iter), allItemsDB.Item(lc).Item(CStr(iter)).Item("qoh"), allItemsDB.Item(lc).Item(CStr(iter)).Item("name"))
        Next iter
    End If
End Sub

Private Sub FullList_SelectionChanged()
    If Me.FullList.SelCount = 0 Then
        clearPreview
        Me.btnMoveRight.Enabled = False
    Else
        fillPreview Me.FullList.GetTextSelected
        Me.btnMoveRight.Enabled = True
    End If
End Sub

Private Sub FullList_DblClick(i As Long, j As Long, txt As String)
    btnMoveRight_Click
End Sub

Private Sub NewKitList_SelectionChanged()
    Dim onoff As Boolean
    onoff = CBool(Me.NewKitList.SelCount = 1)
    
    Me.btnMoveDown.Enabled = onoff
    Me.btnMoveUp.Enabled = onoff
    Me.btnMoveLeft.Enabled = onoff
    Dim iter As Variant
    For Each iter In Me.btnQuantity
        iter.Enabled = onoff
    Next iter
End Sub

Private Sub NewKitList_DblClick(i As Long, j As Long, txt As String)
    btnMoveLeft_Click
End Sub



Private Sub btnMoveLeft_Click()
    If Me.NewKitList.SelCount = 0 Then
        Exit Sub
    End If
    
    Dim qty As String
    qty = Me.NewKitList.GetTextSelected(1)
    If qty = "1" Then
        Me.NewKitList.RemoveSelected
    Else
        Me.NewKitList.Edit CStr(CLng(qty) - 1), Me.NewKitList.SelIndex, 1
    End If
End Sub

Private Sub btnMoveRight_Click()
    If Me.FullList.SelCount = 0 Then
        Exit Sub
    End If
    
    Dim Item As String
    Item = Me.FullList.GetTextSelected
    
    Dim dx As Long, i As Long
    dx = -1
    For i = 0 To Me.NewKitList.ListCount - 1
        If Me.NewKitList.GetText(i) = Item Then
            dx = i
            Exit For
        End If
    Next i
    
    If dx = -1 Then
        Me.NewKitList.Add Array(Item, "1", CStr(allItemsDB.Item(Left(Item, 3)).Item(Item).Item("name")))
        Me.NewKitList.SelectRow (Me.NewKitList.ListCount - 1)
    Else
        Dim qty As String
        qty = Me.NewKitList.GetText(dx, 1)
        Me.NewKitList.Edit CStr(CLng(qty) + 1), dx, 1
        Me.NewKitList.SelectRow i
    End If
End Sub

Private Sub btnMoveUp_Click()
    If Me.NewKitList.SelCount = 0 Then
        Exit Sub
    End If
    If Me.NewKitList.ListCount < 2 Then
        Exit Sub
    End If
    Dim dx As Long
    dx = Me.NewKitList.SelIndex
    If dx = 0 Then
        Exit Sub
    End If
    
    Dim temp As Variant
    temp = Me.NewKitList.GetRow(dx)
    Me.NewKitList.Edit Me.NewKitList.GetText(dx - 1, 0), dx, 0
    Me.NewKitList.Edit Me.NewKitList.GetText(dx - 1, 1), dx, 1
    Me.NewKitList.Edit CStr(temp(0)), dx - 1, 0
    Me.NewKitList.Edit CStr(temp(1)), dx - 1, 1
    Me.NewKitList.SelectRow dx - 1
End Sub

Private Sub btnMoveDown_Click()
    If Me.NewKitList.SelCount = 0 Then
        Exit Sub
    End If
    If Me.NewKitList.ListCount < 2 Then
        Exit Sub
    End If
    Dim dx As Long
    dx = Me.NewKitList.SelIndex
    If dx = Me.NewKitList.ListCount - 1 Then
        Exit Sub
    End If
    
    Dim temp As Variant
    temp = Me.NewKitList.GetRow(dx)
    Me.NewKitList.Edit Me.NewKitList.GetText(dx + 1, 0), dx, 0
    Me.NewKitList.Edit Me.NewKitList.GetText(dx + 1, 1), dx, 1
    Me.NewKitList.Edit CStr(temp(0)), dx + 1, 0
    Me.NewKitList.Edit CStr(temp(1)), dx + 1, 1
    Me.NewKitList.SelectRow dx + 1
End Sub

Private Sub btnQuantity_Click(Index As Integer)
    If Me.NewKitList.SelCount = 1 Then
        Me.NewKitList.Edit CStr(Index), Me.NewKitList.SelIndex, 1
    End If
End Sub



Private Sub fillPreview(Item As String)
    Me.SelectedItemNumber = Item & " - " & allItemsDB.Item(Left(Item, 3)).Item(Item).Item("name")
    Me.SelectedItemSpec1 = allItemsDB.Item(Left(Item, 3)).Item(Item).Item("spec1")
    Me.SelectedItemSpec2 = allItemsDB.Item(Left(Item, 3)).Item(Item).Item("spec2")
    Me.SelectedItemSpec3 = allItemsDB.Item(Left(Item, 3)).Item(Item).Item("spec3")
    Me.SelectedItemPicture.DisplayImage PictureDBFunctions.GenerateItemPicturePathAny(Item)
End Sub

Private Sub clearPreview()
    Me.SelectedItemNumber = ""
    Me.SelectedItemSpec1 = ""
    Me.SelectedItemSpec2 = ""
    Me.SelectedItemSpec3 = ""
    Me.SelectedItemPicture.DisplayImage ""
End Sub

Private Sub loadCache(lc As String)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT InventoryMaster.ItemNumber, QtyOnHand, Desc1, Desc2, Desc3, Desc4, Spec1HTML, Spec2HTML, Spec3HTML " & _
                          "FROM InventoryMaster " & _
                          "         INNER JOIN PartNumbers ON InventoryMaster.ItemNumber=PartNumbers.ItemNumber " & _
                          "         INNER JOIN vItemStatusBreakdown ON InventoryMaster.ItemNumber=vItemStatusBreakdown.ItemNumber " & _
                          "         INNER JOIN vActualWhseQty ON InventoryMaster.ItemNumber=vActualWhseQty.ItemNumber " & _
                          "WHERE /*InventoryMaster.WebsiteID=" & CURRENT_WEBSITE_ID & _
                          "  AND*/ InventoryMaster.IsMASKit=0 " & _
                          "  AND vItemStatusBreakdown.DC=0 " & _
                          "  AND PartNumbers.WebPointered=0 " & _
                          "  AND LEFT(InventoryMaster.ItemNumber,3)='" & lc & "'")
    allItemsDB.Add lc, New Dictionary
    While Not rst.EOF
        Dim Item As String
        Item = rst("ItemNumber")
        
        If Not allItemsDB.Item(lc).exists(Item) Then
            allItemsDB.Item(lc).Add Item, New Dictionary
        End If
        
        allItemsDB.Item(lc).Item(Item).Add "name", TrimWhitespace(ManufWebHashPLtoName.Item(lc) & " " & Nz(rst("Desc2")) & " " & Nz(rst("Desc1")) & " " & Nz(rst("Desc3")) & " " & Nz(rst("Desc4")))
        allItemsDB.Item(lc).Item(Item).Add "qoh", Nz(rst("QtyOnHand"))
        allItemsDB.Item(lc).Item(Item).Add "spec1", Nz(rst("Spec1HTML"))
        allItemsDB.Item(lc).Item(Item).Add "spec2", Nz(rst("Spec2HTML"))
        allItemsDB.Item(lc).Item(Item).Add "spec3", Nz(rst("Spec3HTML"))
        
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Public Sub btnSaveKit_Click()
    If Me.NewKitList.ListCount = 0 Then
        MsgBox "Load some items into the kit list on the right side!"
        Exit Sub
    End If
    If Me.NewKitList.ListCount = 1 And Me.NewKitList.GetText(0, 1) = "1" Then
        MsgBox "Can't create a kit with only 1 item!"
        Exit Sub
    End If
    
    Dim defaultItemNumber As String
    If Me.NewKitList.ListCount = 1 Then
        defaultItemNumber = Me.NewKitList.GetText(0, 0) & "-" & Me.NewKitList.GetText(0, 1)
    Else
        defaultItemNumber = Me.NewKitList.GetText(0, 0) & "-"
    End If
    
    Dim newitem As String
    newitem = InputBox("Enter item number for new kit:", "Prompt", defaultItemNumber)
    If newitem = "" Then
        Exit Sub
    End If
    
    Dim msg As String, i As Long
    msg = "Kit name: " & newitem & vbCrLf & vbCrLf & "Comprised of:"
    Dim itemArray As PerlArray
    Set itemArray = New PerlArray
    For i = 0 To Me.NewKitList.ListCount - 1
        Dim qty As String
        qty = Me.NewKitList.GetText(i, 1)
        msg = msg & vbCrLf & IIf(qty = "1", "", qty & " x ") & Me.NewKitList.GetText(i, 0)
        itemArray.Push Array(CStr(Me.NewKitList.GetText(i, 0)), CLng(qty))
    Next i
    
    msg = msg & vbCrLf & vbCrLf & "Kit will be added to the '" & WebsiteNameHash.Item(CURRENT_WEBSITE_ID) & "' website." & vbCrLf & vbCrLf & "Create new kit with these parameters?"
    
    If vbNo = MsgBox(msg, vbYesNo) Then
        Exit Sub
    End If
    
    If Logic.CreateItemAsKit(newitem, CURRENT_WEBSITE_ID, itemArray) Then
        'MsgBox "Created kit! Remember that you'll need to run a Mas export before it's 100% usable."
        Me.NewKitList.Clear
        
        If IsFormLoaded("InventoryMaintenance") Then
            InventoryMaintenance.GoToItem newitem, True
            FocusOnForm "InventoryMaintenance"
        End If
    End If
End Sub



Private Sub submitQtyBtn_Click()
    If Me.NewKitList.SelCount = 1 Then
        Me.NewKitList.Edit CStr(customQty.Text), Me.NewKitList.SelIndex, 1
    End If
    customQty.Text = ""
End Sub
