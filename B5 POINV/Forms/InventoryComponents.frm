VERSION 5.00
Begin VB.Form InventoryComponents 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Components"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   8700
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      Caption         =   "Item Attributes:"
      Enabled         =   0   'False
      Height          =   555
      Left            =   60
      TabIndex        =   40
      Top             =   480
      Width           =   2775
      Begin VB.CheckBox chkIsMASKit 
         Caption         =   "Mas Kit"
         Height          =   225
         Left            =   180
         TabIndex        =   41
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.TextBox LoadedItemID 
      Height          =   225
      Left            =   6570
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "item"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame4 
      Caption         =   "Add Existing Component:"
      Height          =   3015
      Left            =   60
      TabIndex        =   27
      Top             =   3660
      Width           =   3165
      Begin VB.CommandButton btnExistingComponentAdd 
         Caption         =   "Add Selected Component"
         Height          =   405
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2895
      End
      Begin VB.ListBox ExistingComponentList 
         Height          =   1620
         ItemData        =   "InventoryComponents.frx":0000
         Left            =   90
         List            =   "InventoryComponents.frx":0002
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   870
         Width           =   2955
      End
      Begin VB.TextBox ExistingComponentSearch 
         Height          =   285
         Left            =   690
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   570
         Width           =   2385
      End
      Begin VB.ComboBox ExistingComponentFilter 
         Height          =   315
         Left            =   690
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   2385
      End
      Begin VB.Label Label2 
         Caption         =   "Filter:"
         Height          =   225
         Left            =   90
         TabIndex        =   29
         Top             =   300
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Search:"
         Height          =   225
         Left            =   90
         TabIndex        =   28
         Top             =   600
         Width           =   585
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Add New Component:"
      Height          =   765
      Left            =   60
      TabIndex        =   26
      Top             =   2850
      Width           =   3165
      Begin VB.CommandButton btnComponentAddNew 
         Caption         =   "Add New Component"
         Height          =   405
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   465
      Left            =   7290
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   1395
   End
   Begin VB.TextBox LoadedComponentID 
      Height          =   225
      Left            =   6570
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "ID"
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Frame Frame2 
      Height          =   5565
      Left            =   3330
      TabIndex        =   25
      Top             =   1110
      Width           =   5265
      Begin VB.CheckBox ComponentGiftWarning 
         Caption         =   "Gift Not Concealable Warning"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   3150
         Width           =   2685
      End
      Begin VB.TextBox ComponentQuantity 
         Height          =   285
         Left            =   1260
         TabIndex        =   45
         Top             =   210
         Width           =   555
      End
      Begin VB.Frame Frame7 
         Caption         =   "Barcodes:"
         Height          =   1155
         Left            =   1140
         TabIndex        =   42
         Top             =   4410
         Width           =   1965
         Begin VB.ListBox ComponentBarcodeList 
            Height          =   840
            Left            =   90
            TabIndex        =   43
            Top             =   240
            Width           =   1785
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Currently a component of:"
         Height          =   1155
         Left            =   3090
         TabIndex        =   39
         Top             =   4410
         Width           =   2175
         Begin VB.ListBox ComponentItemList 
            Height          =   840
            Left            =   90
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   240
            Width           =   1995
         End
      End
      Begin VB.CommandButton btnPickLocationGuess 
         Caption         =   "Guess"
         Height          =   315
         Left            =   4170
         TabIndex        =   17
         Top             =   3450
         Width           =   795
      End
      Begin VB.ComboBox PreferredPickLocation 
         Height          =   315
         ItemData        =   "InventoryComponents.frx":0004
         Left            =   1980
         List            =   "InventoryComponents.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3450
         Width           =   2175
      End
      Begin VB.TextBox ComponentWeight 
         Height          =   285
         Left            =   810
         TabIndex        =   9
         Top             =   1050
         Width           =   945
      End
      Begin VB.TextBox ComponentNMFCClassification 
         Height          =   285
         Left            =   810
         TabIndex        =   19
         Top             =   4080
         Width           =   1845
      End
      Begin VB.CommandButton btnPackingBoxGuess 
         Caption         =   "Guess"
         Height          =   315
         Left            =   2850
         TabIndex        =   15
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox PackingBoxSize 
         Height          =   315
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2760
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.ComboBox ComponentPackaging 
         Height          =   315
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2400
         Width           =   2865
      End
      Begin VB.CheckBox ComponentORMD 
         Caption         =   "HazMat / ORM-D"
         Height          =   255
         Left            =   150
         TabIndex        =   18
         Top             =   3780
         Width           =   1635
      End
      Begin VB.TextBox ComponentDimension 
         Height          =   285
         Index           =   2
         Left            =   810
         TabIndex        =   12
         Top             =   2040
         Width           =   945
      End
      Begin VB.TextBox ComponentDimension 
         Height          =   285
         Index           =   1
         Left            =   810
         TabIndex        =   11
         Top             =   1740
         Width           =   945
      End
      Begin VB.TextBox ComponentDimension 
         Height          =   285
         Index           =   0
         Left            =   810
         TabIndex        =   10
         Top             =   1440
         Width           =   945
      End
      Begin VB.TextBox ComponentName 
         Height          =   285
         Left            =   2010
         TabIndex        =   8
         Top             =   660
         Width           =   3135
      End
      Begin VB.Label Label11 
         Caption         =   "Qty in this item:"
         Height          =   225
         Left            =   150
         TabIndex        =   44
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Preferred Whse Pick Loc:"
         Height          =   225
         Left            =   90
         TabIndex        =   38
         Top             =   3480
         Width           =   1845
      End
      Begin VB.Label Label7 
         Caption         =   "Weight:"
         Height          =   225
         Left            =   150
         TabIndex        =   37
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label9 
         Caption         =   "NMFC #:"
         Height          =   225
         Left            =   120
         TabIndex        =   36
         Top             =   4110
         Width           =   705
      End
      Begin VB.Label PackingBoxLabel 
         Caption         =   "Box Size:"
         Height          =   225
         Left            =   450
         TabIndex        =   35
         Top             =   2820
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label8 
         Caption         =   "Packing Type:"
         Height          =   225
         Left            =   90
         TabIndex        =   34
         Top             =   2460
         Width           =   1125
      End
      Begin VB.Label Label6 
         Caption         =   "Height:"
         Height          =   225
         Left            =   150
         TabIndex        =   33
         Top             =   2070
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Width:"
         Height          =   225
         Left            =   150
         TabIndex        =   32
         Top             =   1770
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Length:"
         Height          =   225
         Left            =   150
         TabIndex        =   31
         Top             =   1470
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Part Number / Description:"
         Height          =   225
         Left            =   90
         TabIndex        =   30
         Top             =   690
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Components:"
      Height          =   1575
      Left            =   60
      TabIndex        =   24
      Top             =   1110
      Width           =   3165
      Begin VB.CommandButton btnComponentRemove 
         Caption         =   "Remove Selected"
         Height          =   225
         Left            =   1470
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1515
      End
      Begin VB.ListBox ComponentList 
         Height          =   1035
         ItemData        =   "InventoryComponents.frx":0008
         Left            =   90
         List            =   "InventoryComponents.frx":000A
         TabIndex        =   1
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Label lblTitle 
      Caption         =   "Components for %ITEMNAME%"
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
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4905
   End
End
Attribute VB_Name = "InventoryComponents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private boxIDtoGUIIndexMap As Dictionary

Private allComponents() As String
Private lastQSearch As String

Private fillingForm As Boolean
Private changed As Boolean
Private whichCtl As String

Private permissionsLockDown As Boolean

'=============== PUBLIC =========================

Public Sub LoadItem(itemID As String)
    Mouse.Hourglass True
    
    Me.LoadedItemID = itemID
    'Me.lblTitle.Caption = Replace(Me.lblTitle.Caption, "%ITEMNAME%", item)
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber, IsMASKit FROM InventoryMaster WHERE ItemID=" & itemID)
    Dim item As String
    If rst.EOF Then
        MsgBox "WTF"
    Else
        fillingForm = True
        item = rst("ItemNumber")
        Me.lblTitle.Caption = Replace(Me.lblTitle.Caption, "%ITEMNAME%", rst("ItemNumber"))
        Me.chkIsMASKit = SQLBool(rst("IsMASKit"))
        fillingForm = False
    End If
    rst.Close
    
    Set rst = DB.retrieve("SELECT InventoryComponents.ID, InventoryComponents.Component, InventoryComponentMap.Quantity FROM InventoryComponentMap INNER JOIN InventoryComponents ON InventoryComponentMap.ComponentID=InventoryComponents.ID WHERE InventoryComponentMap.ItemID=" & itemID)
    Me.ComponentList.Clear
    Me.btnComponentRemove.Enabled = False
    While Not rst.EOF
        Me.ComponentList.AddItem rst("Quantity") & vbTab & rst("Component") & vbTab & rst("ID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    Dim pos As Long
    pos = InComboPosition(Left(item, 3), Me.ExistingComponentFilter, False)
    If pos <> -1 Then
        Me.ExistingComponentFilter.ListIndex = pos
    Else
        Me.ExistingComponentFilter.ListIndex = 0
    End If
    
    If Me.ComponentList.ListCount = 1 Then
        Me.ComponentList.ListIndex = 0
    End If
    
    Mouse.Hourglass False
End Sub

'============= INTERNAL ====================
Private Sub selectComponent(id As String)
    Dim i As Long
    For i = 0 To Me.ComponentList.ListCount - 1
        'If getComponentID(Me.ComponentList, i) = "id" Then
        If ListBoxLineColumn(Me.ComponentList, i, 2) = id Then
            If Me.ComponentList.ListIndex = i Then
                'explicitly call?
                ComponentList_Click
            Else
                Me.ComponentList.ListIndex = i
            End If
            Exit For
        End If
    Next i
End Sub

Private Function getComponentID(ctl As ListBox, Optional dx As Long = -2) As String
    If dx = -2 And ctl.ListIndex = -1 Then
        'uh oh
        getComponentID = -1
        Exit Function
    Else
        If dx = -2 Then
            dx = ctl.ListIndex
        End If
        getComponentID = Mid(ctl.List(dx), InStr(ctl.List(dx), vbTab) + 1)
    End If
End Function

Private Function getComponentName(ctl As ListBox, Optional dx As Long = -2) As String
    If dx = -2 And ctl.ListIndex = -1 Then
        'uh oh
        getComponentName = "ERROR"
        Exit Function
    Else
        If dx = -2 Then
            dx = ctl.ListIndex
        End If
        getComponentName = Left(ctl.List(dx), InStr(ctl.List(dx), vbTab) - 1)
    End If
End Function

Private Sub LoadComponent(id As String)
    Mouse.Hourglass True
    
    If changed Then
        handleChange
    End If
    
    Me.LoadedComponentID = id
    
    fillingForm = True
    
    Me.ComponentQuantity = ListBoxLineColumn(Me.ComponentList, Me.ComponentList.ListIndex, 0)
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM InventoryComponents WHERE ID=" & id)
    Me.ComponentName = rst("Component")
    Me.ComponentWeight = rst("Weight")
    Me.ComponentDimension(0) = rst("Length")
    Me.ComponentDimension(1) = rst("Width")
    Me.ComponentDimension(2) = rst("Height")
    Me.ComponentPackaging.ListIndex = rst("PackingType")
    If Nz(rst("ShippingBoxID")) = "" Then
        Me.PackingBoxSize.ListIndex = -1
    Else
        Me.PackingBoxSize.ListIndex = boxIDtoGUIIndexMap.item(CStr(rst("ShippingBoxID")))
    End If
    If rst("PackingType") = 0 Then
        Me.PackingBoxLabel.Visible = True
        Me.PackingBoxSize.Visible = True
        Me.btnPackingBoxGuess.Visible = True
    Else
        Me.PackingBoxLabel.Visible = False
        Me.PackingBoxSize.Visible = False
        Me.btnPackingBoxGuess.Visible = False
    End If
    Me.ComponentGiftWarning = SQLBool(rst("GiftNotConcealableWarning"))
    Me.PreferredPickLocation.ListIndex = Nz(rst("PreferredPickLocation"), "-1")
    Me.ComponentORMD = SQLBool(rst("HazMat"))
    Me.ComponentNMFCClassification = Nz(rst("NMFC"))
    If permissionsLockDown Then
        Disable Me.ComponentName
        Disable Me.ComponentWeight
        Disable Me.ComponentDimension(0)
        Disable Me.ComponentDimension(1)
        Disable Me.ComponentDimension(2)
        Disable Me.ComponentPackaging
        Disable Me.PackingBoxSize
        Disable Me.ComponentGiftWarning
        Disable Me.PreferredPickLocation
        Disable Me.ComponentORMD
        Disable Me.ComponentNMFCClassification
    Else
        Enable Me.ComponentName
        Enable Me.ComponentWeight
        Enable Me.ComponentDimension(0)
        Enable Me.ComponentDimension(1)
        Enable Me.ComponentDimension(2)
        Enable Me.ComponentPackaging
        Enable Me.PackingBoxSize
        Enable Me.ComponentGiftWarning
        Enable Me.PreferredPickLocation
        Enable Me.ComponentORMD
        Enable Me.ComponentNMFCClassification
        Me.btnPackingBoxGuess.Enabled = True
        Me.btnPickLocationGuess.Enabled = True
    End If
    
    rst.Close
    Set rst = Nothing
    
    Set rst = DB.retrieve("SELECT InventoryMaster.ItemNumber FROM InventoryComponentMap INNER JOIN InventoryMaster ON InventoryComponentMap.ItemID=InventoryMaster.ItemID WHERE ComponentID=" & Me.LoadedComponentID)
    Me.ComponentItemList.Clear
    While Not rst.EOF
        Me.ComponentItemList.AddItem rst("ItemNumber")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    Set rst = DB.retrieve("SELECT Barcode FROM InventoryComponentBarcodes WHERE ComponentID=" & Me.LoadedComponentID)
    Me.ComponentBarcodeList.Clear
    While Not rst.EOF
        Me.ComponentBarcodeList.AddItem rst("Barcode")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    fillingForm = False
    
    Mouse.Hourglass False
End Sub

Private Sub handleChange()
    If changed Then
        Select Case whichCtl
            Case Is = "ComponentName"
                ComponentName_LostFocus
            Case Is = "ComponentWeight"
                ComponentWeight_LostFocus
            Case Is = "ComponentDimension0"
                ComponentDimension_LostFocus 0
            Case Is = "ComponentDimension1"
                ComponentDimension_LostFocus 1
            Case Is = "ComponentDimension2"
                ComponentDimension_LostFocus 2
            Case Is = ComponentNMFCClassification
                ComponentNMFCClassification_LostFocus
            Case Else
                MsgBox "Control not in handleChange() switch, losing changes!"
        End Select
    End If
End Sub

Private Sub requeryExistingComponentList()
    Dim rst As ADODB.Recordset
    If Me.ExistingComponentFilter = "<All>" Then
        Set rst = DB.retrieve("SELECT ID, Component FROM InventoryComponents ORDER BY Component")
    Else
        Set rst = DB.retrieve("SELECT ID, Component FROM InventoryComponents WHERE ID IN (SELECT DISTINCT InventoryComponentMap.ComponentID FROM InventoryComponentMap INNER JOIN InventoryMaster ON InventoryComponentMap.ItemID=InventoryMaster.ItemID WHERE LEFT(InventoryMaster.ItemNumber,3)='" & Me.ExistingComponentFilter & "') ORDER BY InventoryComponents.Component")
    End If
    If rst.RecordCount = 0 Then
        allComponents = Split("", "zerolengtharray")
    Else
        ReDim allComponents(rst.RecordCount - 1) As String
        Dim i As Long
        For i = 0 To rst.RecordCount - 1
            allComponents(i) = rst("Component") & vbTab & rst("ID")
            rst.MoveNext
        Next i
    End If
    fillExistingComponentList
End Sub

Private Sub fillExistingComponentList()
    Dim cursel As String
    If Me.ExistingComponentList.ListIndex = -1 Then
        cursel = ""
    Else
        cursel = Me.ExistingComponentList
    End If
    Dim foundsel As Boolean
    foundsel = False
    Me.ExistingComponentList.Clear
    Dim i As Long
    For i = 0 To UBound(allComponents)
        If InList(allComponents(i), Me.ComponentList, False, 1) Then
            'do not display
        ElseIf CBool(InStr(1, CStr(allComponents(i)), lastQSearch, vbTextCompare)) Then
            Me.ExistingComponentList.AddItem allComponents(i)
            If allComponents(i) = cursel Then
                fillingForm = True
                Me.ExistingComponentList = allComponents(i)
                fillingForm = False
                foundsel = True
            End If
        End If
    Next i
    If Not foundsel Then
        clearForm
        Me.btnExistingComponentAdd.Enabled = False
    Else
        'TODO: make sure selection is in visible part of list
    End If
End Sub

Private Sub clearForm()
    fillingForm = True
    Me.LoadedComponentID = ""
    Me.ComponentName = ""
    Me.ComponentWeight = ""
    Me.ComponentDimension(0) = ""
    Me.ComponentDimension(1) = ""
    Me.ComponentDimension(2) = ""
    Me.ComponentPackaging.ListIndex = -1
    Me.PackingBoxSize.ListIndex = -1
    Me.ComponentGiftWarning = 0
    Me.PreferredPickLocation.ListIndex = -1
    Me.ComponentORMD = 0
    Me.ComponentNMFCClassification = ""
    Me.ComponentItemList.Clear
    Disable Me.ComponentName
    Disable Me.ComponentWeight
    Disable Me.ComponentDimension(0)
    Disable Me.ComponentDimension(1)
    Disable Me.ComponentDimension(2)
    Disable Me.ComponentPackaging
    Disable Me.ComponentGiftWarning
    Disable Me.PackingBoxSize
    Disable Me.ComponentORMD
    Disable Me.ComponentNMFCClassification
    Disable Me.PreferredPickLocation
    Me.btnPackingBoxGuess.Enabled = False
    Me.btnPickLocationGuess.Enabled = False
    fillingForm = False
End Sub

Private Sub requeryComponentPackaging()
    Me.ComponentPackaging.Clear
    Me.ComponentPackaging.AddItem "In Brown Box"
    Me.ComponentPackaging.AddItem "Own Box w/ Additional Padding"
    Me.ComponentPackaging.AddItem "Own Box no Additional Padding"
End Sub

Private Sub requeryBoxSizes()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, Length, Width, Height FROM BoxDimensions WHERE ID>0 ORDER BY Length, Width, Height")
    Set boxIDtoGUIIndexMap = New Dictionary
    boxIDtoGUIIndexMap.Add "-1", "-1"
    Me.PackingBoxSize.Clear
    While Not rst.EOF
        boxIDtoGUIIndexMap.Add CStr(rst("ID")), CStr(Me.PackingBoxSize.ListCount)
        Me.PackingBoxSize.AddItem rst("Length") & " x " & rst("Width") & " x " & rst("Height")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryFilterBox()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ProductLine FROM ProductLine WHERE Hide=0 ORDER BY ProductLine")
    Me.ExistingComponentFilter.Clear
    Me.ExistingComponentFilter.AddItem "<All>"
    While Not rst.EOF
        Me.ExistingComponentFilter.AddItem rst("ProductLine")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    'initial value will be set in LoadItem()
End Sub

Private Sub requeryPickLocations()
    Me.PreferredPickLocation.Clear
    Me.PreferredPickLocation.AddItem "Pallet / Bulk"
    Me.PreferredPickLocation.AddItem "Flow Rack"
    Me.PreferredPickLocation.AddItem "Carousel (Horizontal)"
    Me.PreferredPickLocation.AddItem "Carousel (Vertical)"
    Me.PreferredPickLocation.AddItem "Overflow"
End Sub

'============== GUI =====================
Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub ComponentBarcodeList_DblClick()
    If HasPermissionsTo("Barcode") Then
        If IsFormLoaded("Barcode") Then
            'Barcode.LoadItemComponent Me.ComponentItemList.list(0), Me.LoadedComponentID
            'FocusOnForm "Barcode"
            MsgBox "hmm, you already have the barcode window open, and I can't open another one. Sorry!"
        Else
            Load Barcode
            Barcode.LoadItemComponent Me.ComponentItemList.List(0), Me.LoadedComponentID
            Barcode.Show MODAL
        End If
    End If
End Sub

Private Sub ComponentGiftWarning_Click()
    If Not fillingForm Then
        DB.Execute "UPDATE InventoryComponents SET GiftNotConcealableWarning=" & SQLBool(Me.ComponentGiftWarning) & " WHERE ID=" & Me.LoadedComponentID
    End If
End Sub

Private Sub ComponentWeight_GotFocus()
    Me.ComponentWeight.selStart = 0
    Me.ComponentWeight.SelLength = Len(Me.ComponentWeight)
End Sub

Private Sub ComponentDimension_GotFocus(Index As Integer)
    Me.ComponentDimension(Index).selStart = 0
    Me.ComponentDimension(Index).SelLength = Len(Me.ComponentDimension(Index))
End Sub

Private Sub Form_Load()
    SetListTabs Me.ComponentList, Array(12, 288)
    SetListTabs Me.ExistingComponentList, Array(300)
    'calling procedure should call LoadItem() after making form
    requeryComponentPackaging
    requeryBoxSizes
    requeryFilterBox
    requeryPickLocations
    clearForm
    If HasPermissionsTo("WeightDims") Then
        'ok
        permissionsLockDown = False
    Else
        'read only
        Me.Frame3.Visible = False 'add new
        Me.Frame4.Visible = False 'add existing
        Me.Frame6.Visible = False 'attributes
        Me.btnPackingBoxGuess.Visible = False
        Me.btnPickLocationGuess.Visible = False
        Me.btnComponentRemove.Visible = False
        permissionsLockDown = True
    End If
    lastQSearch = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If changed Then
        handleChange
    End If
End Sub

'============= ATTRIBUTES ====================
Private Sub chkIsMASKit_Click()
    If Not fillingForm Then
        'DB.Execute "UPDATE InventoryMaster SET IsMASKit=" & Me.chkIsMASKit & " WHERE ItemID=" & Me.LoadedItemID
        'MsgBox "NOTE: you still need to " & IIf(Me.chkIsMASKit, "create", "delete") & " the kit manually in MAS!"
        MsgBox "This checkbox no longer does anything, you need to create kits through the Kit Creator!"
        fillingForm = True
        Me.chkIsMASKit = IIf(Me.chkIsMASKit, "0", "1")
        fillingForm = False
    End If
End Sub

'============= LEFT SIDE =====================

Private Sub ComponentList_Click()
    Me.btnComponentRemove.Enabled = True
    Dim cid As String
    'cid = getComponentID(Me.ComponentList)
    cid = ListBoxLineColumn(Me.ComponentList, Me.ComponentList.ListIndex, 2)
    If cid = Me.LoadedComponentID Then
        'do nothing
    Else
        LoadComponent cid
    End If
End Sub

Private Sub btnComponentRemove_Click()
    If Me.ComponentList.ListIndex = -1 Then
        MsgBox "Error, this button should be disabled!"
    ElseIf Me.ComponentList.ListCount = 1 Then
        MsgBox "You can't remove the last component!"
    Else
        Dim cid As String
        'cid = getComponentID(Me.ComponentList) 'should be the same as loaded
        cid = ListBoxLineColumn(Me.ComponentList, Me.ComponentList.ListIndex, 2)
        Dim cantDeassoc As Boolean
        cantDeassoc = False
        If Me.chkIsMASKit = 0 Then
            If DLookup("COALESCE(SUM(Quantity),0)", "LocationContents", "ComponentID=" & cid) > 0 Then
                cantDeassoc = True
            End If
        End If
        If cantDeassoc Then
            MsgBox "Don't know how to de-associate a component that has warehouse locations! Talk to Brian!"
        Else
            Dim doDel As Boolean, doDeAssoc As Boolean, delBarcodes As Boolean
            doDel = False
            doDeAssoc = False
            If "1" = DLookup("COUNT(*)", "InventoryComponentMap", "ComponentID=" & cid) Then
                If vbYes = MsgBox("De-associating this will delete the component! Are you sure you want to delete it?", vbYesNo) Then
                    doDel = True
                    doDeAssoc = True
                End If
            Else
                If vbYes = MsgBox("Really de-associate this component?", vbYesNo) Then
                    doDeAssoc = True
                End If
            End If
            If doDeAssoc Then
                If doDel Then
                    DB.Execute "DELETE FROM InventoryComponents WHERE ID=" & cid
                    DB.Execute "DELETE FROM InventoryComponentBarcodes WHERE ComponentID=" & cid
                End If
                DB.Execute "DELETE FROM InventoryComponentMap WHERE ItemID=" & Me.LoadedItemID & " AND ComponentID=" & cid
                ExternalComponentDBRemove Me.LoadedItemID, cid
                Me.ComponentList.RemoveItem Me.ComponentList.ListIndex
                clearForm
            End If
        End If
    End If
End Sub

Private Sub btnComponentAddNew_Click()
    Dim newname As String
    newname = ""
retry:
    newname = InputBox("Enter the part number and/or description for this component:")
    If newname = "" Then
        'do nothing
    ElseIf Len(newname) > 32 Then
        MsgBox "Must be under 32 characters! Truncating!"
        newname = Left(newname, 32)
        GoTo retry
    Else
        DB.Execute "INSERT INTO InventoryComponents ( Component ) VALUES ( '" & EscapeSQuotes(newname) & "' )"
        Dim newid As String
        newid = DLookup("@@IDENTITY", "InventoryComponents")
        Dim item As String
        item = DLookup("ItemNumber", "InventoryMaster", "ItemID=" & Me.LoadedItemID)
        DB.Execute "INSERT INTO InventoryComponentMap ( ItemID, ItemNumber, ComponentID ) VALUES ( " & Me.LoadedItemID & ", '" & item & "', " & newid & " )"
        Me.ComponentList.AddItem "1" & vbTab & newname & vbTab & newid
        selectComponent newid
    End If
End Sub

Private Sub ExistingComponentFilter_Click()
    Mouse.Hourglass True
    requeryExistingComponentList
    Mouse.Hourglass False
End Sub

Private Sub ExistingComponentList_Click()
    Me.btnExistingComponentAdd.Enabled = True
End Sub

Private Sub ExistingComponentSearch_Change()
    If Me.ExistingComponentSearch = lastQSearch Then
        'do nothing
    Else
        lastQSearch = Me.ExistingComponentSearch
        Mouse.Hourglass True
        fillExistingComponentList
        Mouse.Hourglass False
    End If
End Sub

Private Sub btnExistingComponentAdd_Click()
    If Me.ExistingComponentList.SelCount <> 1 Then
        MsgBox "No selection! This should have been disabled..."
    Else
        Dim item As String
        item = DLookup("ItemNumber", "InventoryMaster", "ItemID=" & Me.LoadedItemID)
        If vbYes = MsgBox("Add " & getComponentName(Me.ExistingComponentList) & " to the " & item & "?", vbYesNo) Then
            DB.Execute "INSERT INTO InventoryComponentMap ( ItemID, ItemNumber, ComponentID ) VALUES ( " & Me.LoadedItemID & ", '" & item & "', " & getComponentID(Me.ExistingComponentList) & " )"
            Me.ComponentList.AddItem "1" & vbTab & Me.ExistingComponentList
            Me.ExistingComponentList.RemoveItem Me.ExistingComponentList.ListIndex
            'loadComponent getComponentID(Me.ExistingComponentList)
        End If
    End If
End Sub

'============ RIGHT SIDE ===============
Private Sub ComponentQuantity_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            ComponentQuantity_LostFocus
        Case Is = vbKeyDelete
            ComponentQuantity_KeyPress KeyCode
    End Select
End Sub

Private Sub ComponentQuantity_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
    If KeyAscii <> 0 Then
        changed = True
        whichCtl = "ComponentQuantity"
    End If
End Sub

Private Sub ComponentQuantity_LostFocus()
    If changed Then
        If IsIntegeric(Me.ComponentQuantity) Then
            DB.Execute "UPDATE InventoryComponentMap SET Quantity=" & Me.ComponentQuantity & " WHERE ItemID=" & Me.LoadedItemID & " AND ComponentID=" & Me.LoadedComponentID
            changed = False
            
            Dim i As Long
            For i = 0 To Me.ComponentList.ListCount - 1
                If ListBoxLineColumn(Me.ComponentList, i, 2) = Me.LoadedComponentID Then
                    Me.ComponentList.List(i) = Me.ComponentQuantity & vbTab & ListBoxLineColumn(Me.ComponentList, i, 1) & vbTab & ListBoxLineColumn(Me.ComponentList, i, 2)
                    Exit For
                End If
            Next i
        Else
            MsgBox "Error, must be a number"
        End If
    End If
End Sub

Private Sub ComponentName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            ComponentName_LostFocus
        Case Is = vbKeyDelete
            ComponentName_KeyPress KeyCode
    End Select
End Sub

Private Sub ComponentName_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "ComponentName"
End Sub

Private Sub ComponentName_LostFocus()
    If changed Then
        If Len(Me.ComponentName) > 32 Then
            MsgBox "More than 32 chars! Truncating!"
            Me.ComponentName = Left(Me.ComponentName, 32)
        End If
        If Me.ComponentName = "" Then
            MsgBox "Component name cannot be blank!"
            Me.ComponentName = DLookup("Component", "InventoryComponents", "ID=" & Me.LoadedComponentID)
        Else
            DB.Execute "UPDATE InventoryComponents SET Component='" & EscapeSQuotes(Me.ComponentName) & "' WHERE ID=" & Me.LoadedComponentID
        End If
        changed = False
        
        Dim i As Long
        For i = 0 To Me.ComponentList.ListCount - 1
            If ListBoxLineColumn(Me.ComponentList, i, 2) = Me.LoadedComponentID Then
                Me.ComponentList.List(i) = ListBoxLineColumn(Me.ComponentList, i, 0) & vbTab & Me.ComponentName & vbTab & ListBoxLineColumn(Me.ComponentList, i, 2)
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub ComponentWeight_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            ComponentWeight_LostFocus
        Case Is = vbKeyDelete
            ComponentWeight_KeyPress KeyCode
    End Select
End Sub

Private Sub ComponentWeight_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii, True, True)
    If KeyAscii <> 0 Then
        changed = True
        whichCtl = "ComponentWeight"
    End If
End Sub

Private Sub ComponentWeight_LostFocus()
    If changed Then
        Me.ComponentWeight = ConvertFracToDec(Me.ComponentWeight)
        Me.ComponentWeight = Round(Me.ComponentWeight, 3)
        DB.Execute "UPDATE InventoryComponents SET Weight=" & Me.ComponentWeight & " WHERE ID=" & Me.LoadedComponentID
        'DB.Execute "UPDATE InventoryComponents SET LastModifiedBy=" & GetCurrentUserID() & " WHERE ID=" & Me.LoadedComponentID
        changed = False
        'If CDbl(Me.ComponentWeight) > 150 Then
        '    Dim setTF As Boolean
        '    setTF = False
        '    Dim rst As ADODB.Recordset
        '    'Set rst = DB.retrieve("SELECT COUNT(*) FROM InventoryMaster WHERE TruckFreight=0 AND ItemID IN (SELECT ItemID FROM InventoryComponentMap WHERE ComponentID=" & Me.LoadedComponentID & ")")
        '    Set rst = DB.retrieve("SELECT COUNT(*) FROM InventoryMaster WHERE ShippingType=0 AND ItemID IN (SELECT ItemID FROM InventoryComponentMap WHERE ComponentID=" & Me.LoadedComponentID & ")")
        '    If rst(0) = 1 Then
        '        If vbYes = MsgBox("Over 150 lbs, do you want to set this item to truck freight?", vbYesNo) Then
        '            setTF = True
        '        End If
        '    ElseIf rst(0) > 1 Then
        '        If vbYes = MsgBox("Over 150 lbs, do you want to set all items that have this as a component to truck freight?", vbYesNo) Then
        '            setTF = True
        '        End If
        '    End If
        '    rst.Close
        '    Set rst = Nothing
        '    If setTF Then
        '        'DB.Execute "UPDATE InventoryMaster SET TruckFreight=1 WHERE TruckFreight=0 AND ItemID IN (SELECT ItemID FROM InventoryComponentMap WHERE ComponentID=" & Me.LoadedComponentID & ")"
        '        DB.Execute "UPDATE InventoryMaster SET ShippingType=1 WHERE ShippingType=0 AND ItemID IN (SELECT ItemID FROM InventoryComponentMap WHERE ComponentID=" & Me.LoadedComponentID & ")"
        '    End If
        'End If
    End If
End Sub

Private Sub ComponentDimension_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            ComponentDimension_LostFocus Index
        Case Is = vbKeyDelete
            ComponentDimension_KeyPress Index, KeyCode
    End Select
End Sub

Private Sub ComponentDimension_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii, True, True)
    If KeyAscii <> 0 Then
        changed = True
        whichCtl = "ComponentDimension" & Index
    End If
End Sub

Private Sub ComponentDimension_LostFocus(Index As Integer)
    If changed Then
        If Me.ComponentDimension(Index) = "" Then
            Me.ComponentDimension(Index) = "0"
        End If
        Me.ComponentDimension(Index) = ConvertFracToDec(Me.ComponentDimension(Index))
        If IsNumeric(Me.ComponentDimension(Index)) Then
            Me.ComponentDimension(Index) = Round(Me.ComponentDimension(Index), 3)
            Dim fieldName As String
            Select Case Index
                Case Is = 0
                    fieldName = "Length"
                Case Is = 1
                    fieldName = "Width"
                Case Is = 2
                    fieldName = "Height"
            End Select
            DB.Execute "UPDATE InventoryComponents SET " & fieldName & "=" & Me.ComponentDimension(Index) & " WHERE ID=" & Me.LoadedComponentID
            'DB.Execute "UPDATE InventoryComponents SET LastModifiedBy=" & GetCurrentUserID() & " WHERE ID=" & Me.LoadedComponentID
        Else
            MsgBox "Dimensions must be numeric!"
        End If
        changed = False
    End If
End Sub

Private Sub ComponentPackaging_Click()
    If Not fillingForm Then
        DB.Execute "UPDATE InventoryComponents SET PackingType=" & Me.ComponentPackaging.ListIndex & " WHERE ID=" & Me.LoadedComponentID
        If Me.ComponentPackaging.ListIndex = 0 Then
            Me.PackingBoxLabel.Visible = True
            Me.PackingBoxSize.Visible = True
            Me.btnPackingBoxGuess.Visible = True
            btnPackingBoxGuess_Click
        Else
            Me.PackingBoxLabel.Visible = False
            Me.PackingBoxSize.Visible = False
            Me.btnPackingBoxGuess.Visible = False
            DB.Execute "UPDATE InventoryComponents SET ShippingBoxID=NULL WHERE ID=" & Me.LoadedComponentID
        End If
    End If
End Sub

Private Sub PackingBoxSize_Click()
    If Not fillingForm Then
        Dim boxID As String, iter As Variant
        boxID = "-1"
        For Each iter In boxIDtoGUIIndexMap.keys
            If boxIDtoGUIIndexMap.item(CStr(iter)) = Me.PackingBoxSize.ListIndex Then
                boxID = iter
                Exit For
            End If
        Next iter
        DB.Execute "UPDATE InventoryComponents SET ShippingBoxID=" & boxID & " WHERE ID=" & Me.LoadedComponentID
    End If
End Sub

Private Sub btnPackingBoxGuess_Click()
    Dim boxID As String
    boxID = DetermineBoxSize2(Me.ComponentDimension(0), Me.ComponentDimension(1), Me.ComponentDimension(2), Me.ComponentWeight)
    Me.PackingBoxSize.ListIndex = boxIDtoGUIIndexMap.item(boxID)
    PackingBoxSize_Click
    If boxID = "-1" Then
        MsgBox "Unable to find a box for this!"
    End If
End Sub

Private Sub PreferredPickLocation_Click()
    If Not fillingForm Then
        DB.Execute "UPDATE InventoryComponents SET PreferredPickLocation=" & Me.PreferredPickLocation.ListIndex & " WHERE ID=" & Me.LoadedComponentID
    End If
End Sub

Private Sub btnPickLocationGuess_Click()
    'TODO: refactor
    Dim debugmsg As String
    debugmsg = ""
    
    Dim tooBigForCarousel As Boolean, tooBigForFlow As Boolean, i As Long
    tooBigForCarousel = False
    tooBigForFlow = False
    
    Dim cubicInches As Double
    cubicInches = CDbl(Me.ComponentDimension(0)) * CDbl(Me.ComponentDimension(1)) * CDbl(Me.ComponentDimension(2))
    
    For i = 0 To 2
        If CDbl(Me.ComponentDimension(i)) > 16 Then
            tooBigForCarousel = True
            debugmsg = debugmsg & "dimension " & i & " is too big for carousel" & vbCrLf
        End If
        If CDbl(Me.ComponentDimension(i)) > 60 Then
            tooBigForFlow = True
            debugmsg = debugmsg & "dimension " & i & " is too big for flow rack" & vbCrLf
        End If
    Next i
    If (cubicInches) ^ (1 / 2) > 20 Then
        tooBigForCarousel = True
        debugmsg = debugmsg & "volume is too big for carousel" & vbCrLf
    End If
    If (cubicInches) ^ (1 / 2) > 235 Then
        tooBigForFlow = True
        debugmsg = debugmsg & "volume is too big for flow rack" & vbCrLf
    End If
    
    If CDbl(Me.ComponentWeight) > 12 Then
        tooBigForCarousel = True
        debugmsg = debugmsg & "weight is too big for carousel" & vbCrLf
    End If
    If CDbl(Me.ComponentWeight) > 60 Then
        tooBigForFlow = True
        debugmsg = debugmsg & "weight is too big for flow rack" & vbCrLf
    End If
    
    If tooBigForFlow Then
        Me.PreferredPickLocation.ListIndex = 0
    ElseIf tooBigForCarousel Then
        Me.PreferredPickLocation.ListIndex = 2
    Else
        Me.PreferredPickLocation.ListIndex = 1
    End If
    PreferredPickLocation_Click
    
    If True Then
        MsgBox debugmsg
    End If
End Sub

Private Sub ComponentORMD_Click()
    If Not fillingForm Then
        DB.Execute "UPDATE InventoryComponents SET HazMat=" & SQLBool(Me.ComponentORMD) & " WHERE ID=" & Me.LoadedComponentID
    End If
End Sub

Private Sub ComponentNMFCClassification_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            ComponentNMFCClassification_LostFocus
        Case Is = vbKeyDelete
            ComponentNMFCClassification_KeyPress KeyCode
    End Select
End Sub

Private Sub ComponentNMFCClassification_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "ComponentNMFCClassification"
End Sub

Private Sub ComponentNMFCClassification_LostFocus()
    If changed Then
        If Len(Me.ComponentNMFCClassification) > 16 Then
            MsgBox "More than 16 chars! Truncating!"
            Me.ComponentNMFCClassification = Left(Me.ComponentNMFCClassification, 16)
        End If
        DB.Execute "UPDATE InventoryComponents SET NMFC='" & EscapeSQuotes(Me.ComponentNMFCClassification) & "' WHERE ID=" & Me.LoadedComponentID
        changed = False
    End If
End Sub
