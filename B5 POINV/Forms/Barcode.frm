VERSION 5.00
Begin VB.Form Barcode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Barcode Lookup Maintenance"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9285
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton MasterPackBtn 
      Caption         =   "Create Master Pack Barcode"
      Height          =   615
      Left            =   7680
      TabIndex        =   19
      Top             =   3720
      Width           =   1575
   End
   Begin VB.ListBox MasterPackBarcodeList 
      Height          =   1425
      Left            =   5820
      TabIndex        =   18
      Top             =   2700
      Width           =   1815
   End
   Begin VB.ListBox ComponentList 
      Height          =   3180
      Left            =   3180
      TabIndex        =   16
      Top             =   960
      Width           =   2595
   End
   Begin VB.CommandButton OpenBarcodePrinting 
      Caption         =   "Barcode Printing"
      Height          =   525
      Left            =   7650
      TabIndex        =   15
      Top             =   4770
      Width           =   1545
   End
   Begin VB.CommandButton AddGeneratedBarcode 
      Caption         =   "Create TP Barcode"
      Enabled         =   0   'False
      Height          =   525
      Left            =   7710
      TabIndex        =   14
      Top             =   3090
      Width           =   1545
   End
   Begin VB.CheckBox ShowDiscontinuedItems 
      Caption         =   "Show Discontinued Items"
      Height          =   405
      Left            =   1650
      TabIndex        =   13
      Top             =   4170
      Width           =   2235
   End
   Begin VB.CheckBox ShowOldLineCodes 
      Caption         =   "Show Old Line Codes"
      Height          =   435
      Left            =   210
      TabIndex        =   12
      Top             =   4170
      Width           =   1155
   End
   Begin VB.CommandButton DeleteBarcode 
      Caption         =   "Delete Barcode"
      Enabled         =   0   'False
      Height          =   465
      Left            =   7710
      TabIndex        =   11
      Top             =   2250
      Width           =   1545
   End
   Begin VB.CommandButton EditBarcode 
      Caption         =   "Edit Barcode"
      Enabled         =   0   'False
      Height          =   465
      Left            =   7710
      TabIndex        =   10
      Top             =   1740
      Width           =   1545
   End
   Begin VB.CommandButton AddBarcode 
      Caption         =   "Add New Barcode"
      Enabled         =   0   'False
      Height          =   465
      Left            =   7710
      TabIndex        =   9
      Top             =   1230
      Width           =   1545
   End
   Begin VB.ListBox BarcodeList 
      Height          =   1620
      Left            =   5820
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   525
      Left            =   3150
      TabIndex        =   6
      Top             =   4800
      Width           =   1755
   End
   Begin VB.ListBox PartsList 
      Height          =   3180
      Left            =   1470
      TabIndex        =   1
      Top             =   960
      Width           =   1665
   End
   Begin VB.ListBox LineCodeList 
      Height          =   3180
      Left            =   210
      TabIndex        =   0
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label generalLabel 
      Caption         =   "Component:"
      Height          =   255
      Index           =   5
      Left            =   3900
      TabIndex        =   17
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label generalLabel 
      Caption         =   "Bar Code:"
      Height          =   255
      Index           =   2
      Left            =   5880
      TabIndex        =   8
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label generalLabel 
      Caption         =   "Bar Code Lookup Maintenance"
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
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3945
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "# BCs:"
      Height          =   255
      Index           =   4
      Left            =   3210
      TabIndex        =   4
      Top             =   720
      Width           =   525
   End
   Begin VB.Label generalLabel 
      Caption         =   "Item Number:"
      Height          =   255
      Index           =   3
      Left            =   1530
      TabIndex        =   3
      Top             =   720
      Width           =   1185
   End
   Begin VB.Label generalLabel 
      Caption         =   "Line Code:"
      Height          =   255
      Index           =   1
      Left            =   270
      TabIndex        =   2
      Top             =   720
      Width           =   885
   End
End
Attribute VB_Name = "Barcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim masterBarcodeListLastClicked As Boolean
Private lockDown As Boolean

Public Sub LoadItem(item As String)
    Dim pos As Long
    pos = InListPosition(Left(item, 3), Me.LineCodeList)
    If pos = -1 Then
        Me.ShowOldLineCodes = 1
        pos = InListPosition(Left(item, 3), Me.LineCodeList)
    End If
    Me.LineCodeList.Selected(pos) = True
    Dim i As Long
itemloop:
    If Me.PartsList.ListCount > 0 Then
        For i = 0 To Me.PartsList.ListCount - 1
            'If Left(Me.PartsList.list(i), InStr(Me.PartsList.list(i), vbTab) - 1) = item Then
            If Me.PartsList.list(i) = item Then
                Me.PartsList.Selected(i) = True
                Exit For
            End If
        Next i
        If Me.PartsList.SelCount = 0 Then
            Me.ShowDiscontinuedItems = 1
            GoTo itemloop
        End If
    End If
End Sub

Public Sub LoadItemComponent(item As String, componentID As String)
    LoadItem item
    If Me.PartsList.SelCount <> 0 Then
        Dim i As Long
        For i = 0 To Me.ComponentList.ListCount - 1
            Dim cid As String
            cid = ListBoxLineColumn(Me.ComponentList, i, 2)
            If cid = componentID Then
                Me.ComponentList.Selected(i) = True
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    SetListTabs Me.ComponentList, Array(42, 258)
    requeryLines
    
    lockDown = Not HasPermissionsTo("Barcode")
    If lockDown Then
        Me.OpenBarcodePrinting.Enabled = False
    End If
End Sub

Private Sub LineCodeList_Click()
    requeryParts
End Sub

Private Sub MasterPackBarcodeList_Click()
masterBarcodeListLastClicked = True
    If Not lockDown Then
        Me.EditBarcode.Enabled = True
        Me.DeleteBarcode.Enabled = True
    End If
End Sub

Private Sub MasterPackBtn_Click()
    Dim bc As String
    Dim qty As String
    bc = InputBox("Enter Master Pack Barcode:", "Master Pack Creation")
    qty = InputBox("Enter Quantity:", "Master Pack Creation")
    If IsNumeric(qty) = False Then
        MsgBox "Quantity needs to be a number, not text!"
        Exit Sub
    End If
    
            Dim cid As String
            cid = ListBoxLineColumn(Me.ComponentList, Me.ComponentList.ListIndex, 2)
            'Dim bc As String
            'bc = getNextGeneratedBarcode()
            'DB.Execute "INSERT INTO BarcodeLookup ( Barcode, ItemNumber, Internal ) VALUES ( '" & bc & "', '" & Left(Me.PartsList, InStr(Me.PartsList, vbTab) - 1) & "', 1 )"
            DB.Execute "INSERT INTO InventoryComponentBarcodes ( Barcode, ComponentID, Internal ) VALUES ( '" & bc & "', " & cid & ", 1 )"
            DB.Execute "UPDATE InventoryComponentBarcodes SET Quantity=" & qty & " WHERE Barcode='" & bc & "'"
            Me.MasterPackBarcodeList.AddItem "(" & qty & ") " & bc
            modifyCount 1

End Sub

Private Sub OpenBarcodePrinting_Click()
    Load BarcodePrinting
    BarcodePrinting.Show
End Sub

Private Sub PartsList_Click()
    'requeryBarcodes
    requeryComponents
End Sub

Private Sub ComponentList_Click()
    requeryBarcodes
End Sub

Private Sub BarcodeList_Click()
masterBarcodeListLastClicked = False
    If Not lockDown Then
        Me.EditBarcode.Enabled = True
        Me.DeleteBarcode.Enabled = True
    End If
End Sub

Private Sub BarcodeList_DblClick()
    EditBarcode_Click
End Sub

Private Sub requeryLines()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ProductLine FROM ProductLine" & IIf(Me.ShowOldLineCodes, "", " WHERE Hide=0") & " ORDER BY ProductLine")
    Me.LineCodeList.Clear
    Me.PartsList.Clear
    Me.ComponentList.Clear
    Me.BarcodeList.Clear
    While Not rst.EOF
        Me.LineCodeList.AddItem rst("ProductLine")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Me.AddBarcode.Enabled = False
    Me.EditBarcode.Enabled = False
    Me.DeleteBarcode.Enabled = False
    Me.AddGeneratedBarcode.Enabled = False
End Sub

Private Sub requeryParts()
    If Me.LineCodeList.SelCount > 0 Then
        Dim rst As ADODB.Recordset
        'Set rst = DB.retrieve("SELECT ItemNumber, COUNT(Barcode) AS NumBC FROM vBarcodes WHERE LEFT(ItemNumber,3)='" & Me.LineCodeList & "'" & IIf(Me.ShowDiscontinuedItems, "", " AND Discontinued=0") & " GROUP BY ItemNumber ORDER BY ItemNumber")
        Set rst = DB.retrieve("SELECT InventoryMaster.ItemNumber " & _
                              "FROM InventoryMaster " & _
                              "       INNER JOIN vItemStatusBreakdown ON InventoryMaster.ItemNumber=vItemStatusBreakdown.ItemNumber " & _
                              "WHERE Hide=0 " & _
                              "  AND InventoryMaster.ItemNumber<'X%' " & _
                              "  AND LEFT(InventoryMaster.ItemNumber,3)='" & Me.LineCodeList & "' " & _
                              "  " & IIf(Me.ShowDiscontinuedItems, "", " AND DC=0") & " " & _
                              "ORDER BY InventoryMaster.ItemNumber")
        Me.PartsList.Clear
        Me.ComponentList.Clear
        Me.BarcodeList.Clear
        While Not rst.EOF
            'Me.PartsList.AddItem rst("ItemNumber") & vbTab & rst("NumBC")
            Me.PartsList.AddItem rst("ItemNumber")
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
    End If
    Me.AddBarcode.Enabled = False
    Me.EditBarcode.Enabled = False
    Me.DeleteBarcode.Enabled = False
    Me.AddGeneratedBarcode.Enabled = False
End Sub

Private Sub requeryComponents()
    If Me.PartsList.SelCount > 0 Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT COUNT(InventoryComponentBarcodes.Barcode) AS NumBC, InventoryComponents.Component, InventoryComponentMap.ComponentID " & _
                              "FROM InventoryComponents " & _
                              "        INNER JOIN InventoryComponentMap ON InventoryComponents.ID=InventoryComponentMap.ComponentID " & _
                              "        LEFT OUTER JOIN InventoryComponentBarcodes ON InventoryComponents.ID=InventoryComponentBarcodes.ComponentID " & _
                              "WHERE InventoryComponentMap.ItemID='" & Utilities.GetItemIDByItemCode(Me.PartsList) & "' " & _
                              "GROUP BY InventoryComponents.Component, InventoryComponentMap.ComponentID")
        Me.ComponentList.Clear
        Me.BarcodeList.Clear
        While Not rst.EOF
            Me.ComponentList.AddItem rst("NumBC") & vbTab & rst("Component") & vbTab & rst("ComponentID")
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
    End If
    If Me.ComponentList.ListCount > 0 Then
        Me.ComponentList.Selected(0) = True
    Else
        Me.AddBarcode.Enabled = False
        Me.EditBarcode.Enabled = False
        Me.DeleteBarcode.Enabled = False
        Me.AddGeneratedBarcode.Enabled = False
    End If
End Sub

Private Sub requeryBarcodes()
    'Dim item As String
    'item = Left(Me.PartsList, InStr(Me.PartsList, vbTab) - 1)
    Dim cid As String
    cid = ListBoxLineColumn(Me.ComponentList, Me.ComponentList.ListIndex, 2)
    Dim rst As ADODB.Recordset
    'Set rst = DB.retrieve("SELECT Barcode FROM vBarcodes WHERE ItemNumber='" & item & "'")
    Set rst = DB.retrieve("SELECT Barcode, Quantity FROM vBarcodesComponents WHERE ComponentID=" & cid)
    Me.BarcodeList.Clear
    Me.MasterPackBarcodeList.Clear
    While Not rst.EOF
        If Nz(rst("Barcode")) <> "" Then
            If rst("Quantity") = 1 Then
                Me.BarcodeList.AddItem rst("Barcode")
            Else
                Me.MasterPackBarcodeList.AddItem "(" & rst("Quantity") & ") " & rst("Barcode")
            End If
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    If Not lockDown Then
        Me.AddBarcode.Enabled = True
        Me.AddGeneratedBarcode.Enabled = True
    End If
    Me.EditBarcode.Enabled = False
    Me.DeleteBarcode.Enabled = False
End Sub

Private Sub AddBarcode_Click()
    'Dim item As String, bc As String
    'item = Left(Me.PartsList, InStr(Me.PartsList, vbTab) - 1)
    Dim cid As String, bc As String
    cid = ListBoxLineColumn(Me.ComponentList, Me.ComponentList.ListIndex, 2)
    'bc = AddNewBarcodeFor(item)
    bc = AddNewBarcodeFor(CLng(cid))
    If bc <> "" Then
        Me.BarcodeList.AddItem bc
        modifyCount 1
    End If
End Sub

Private Sub AddGeneratedBarcode_Click()
    If vbYes = MsgBox("Generate a new barcode for this item?", vbYesNo) Then
        Dim cid As String
        cid = ListBoxLineColumn(Me.ComponentList, Me.ComponentList.ListIndex, 2)
        Dim bc As String
        bc = getNextGeneratedBarcode()
        'DB.Execute "INSERT INTO BarcodeLookup ( Barcode, ItemNumber, Internal ) VALUES ( '" & bc & "', '" & Left(Me.PartsList, InStr(Me.PartsList, vbTab) - 1) & "', 1 )"
        DB.Execute "INSERT INTO InventoryComponentBarcodes ( Barcode, ComponentID, Internal ) VALUES ( '" & bc & "', " & cid & ", 1 )"
        Me.BarcodeList.AddItem bc
        modifyCount 1
    End If
End Sub

Private Sub EditBarcode_Click()
    Dim newbc As String
    If masterBarcodeListLastClicked = False Then
        newbc = EditExistingBarcode(Me.BarcodeList, masterBarcodeListLastClicked)
        If newbc <> "" Then
            Me.BarcodeList.list(Me.BarcodeList.ListIndex) = newbc
        End If
    Else
        newbc = EditExistingBarcode(Me.MasterPackBarcodeList, masterBarcodeListLastClicked)
        If newbc <> "" Then
            Me.MasterPackBarcodeList.list(Me.MasterPackBarcodeList.ListIndex) = newbc
        End If
        
    End If
End Sub

Private Sub DeleteBarcode_Click()
If masterBarcodeListLastClicked = False Then
    If DeleteExistingBarcode(Me.BarcodeList) Then
        Me.BarcodeList.RemoveItem Me.BarcodeList.ListIndex
        Me.EditBarcode.Enabled = False
        Me.DeleteBarcode.Enabled = False
        modifyCount -1
    End If
Else
    If DeleteExistingBarcode(CStr(Split(Me.MasterPackBarcodeList, " ")(1))) Then
        Me.MasterPackBarcodeList.RemoveItem Me.MasterPackBarcodeList.ListIndex
        Me.EditBarcode.Enabled = False
        Me.DeleteBarcode.Enabled = False
        modifyCount -1
    End If
End If
End Sub

Private Sub ShowDiscontinuedItems_Click()
    requeryParts
End Sub

Private Sub ShowOldLineCodes_Click()
    requeryLines
End Sub

Private Sub modifyCount(amount As Long)
    'Dim item As String, count As Long
    'item = Left(Me.PartsList, InStr(Me.PartsList, vbTab) - 1)
    'count = CLng(Mid(Me.PartsList, InStr(Me.PartsList, vbTab) + 1))
    'count = count + amount
    'Me.PartsList.list(Me.PartsList.ListIndex) = item & vbTab & count
    Dim count As Long, cName As String, cid As Long
    count = ListBoxLineColumn(Me.ComponentList, Me.ComponentList.ListIndex, 0)
    cName = ListBoxLineColumn(Me.ComponentList, Me.ComponentList.ListIndex, 1)
    cid = ListBoxLineColumn(Me.ComponentList, Me.ComponentList.ListIndex, 2)
    Me.ComponentList.list(Me.ComponentList.ListIndex) = (count + amount) & vbTab & cName & vbTab & cid
End Sub

Private Function getNextGeneratedBarcode() As String
    Dim Last As String
    Last = DLookup("MAX(CAST(SUBSTRING(Barcode,3,10) as bigint))", "InventoryComponentBarcodes", "Internal=1")
    If Last = "" Then
        Last = "1"
    Else
        Last = CDbl(Last) + 1
    End If
    getNextGeneratedBarcode = "TP" & Format(Last, "0000000000")
End Function
