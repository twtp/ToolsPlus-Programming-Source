VERSION 5.00
Begin VB.Form PurchaseOrderSelectionDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order Selection"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   3210
      TabIndex        =   7
      Top             =   3540
      Width           =   1395
   End
   Begin VB.CommandButton btnAddToPurchaseOrder 
      Caption         =   "Add to PO"
      Height          =   405
      Left            =   1530
      TabIndex        =   6
      Top             =   3540
      Width           =   1395
   End
   Begin VB.ListBox PurchaseOrderList 
      Height          =   1620
      ItemData        =   "PurchaseOrderSelectionDialog.frx":0000
      Left            =   60
      List            =   "PurchaseOrderSelectionDialog.frx":0002
      TabIndex        =   5
      Top             =   1320
      Width           =   6585
   End
   Begin VB.TextBox QtyToOrder 
      Height          =   315
      Left            =   1530
      TabIndex        =   4
      Top             =   2970
      Width           =   1125
   End
   Begin VB.TextBox ItemNumber 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1260
      TabIndex        =   2
      Top             =   720
      Width           =   1725
   End
   Begin VB.Label generalLabel 
      Caption         =   "Ship To"
      Height          =   225
      Index           =   5
      Left            =   1770
      TabIndex        =   10
      Top             =   1110
      Width           =   555
   End
   Begin VB.Label generalLabel 
      Caption         =   "S/O #"
      Height          =   225
      Index           =   4
      Left            =   690
      TabIndex        =   9
      Top             =   1110
      Width           =   555
   End
   Begin VB.Label generalLabel 
      Caption         =   "Qty"
      Height          =   225
      Index           =   3
      Left            =   90
      TabIndex        =   8
      Top             =   1110
      Width           =   555
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Qty to Order:"
      Height          =   225
      Index           =   2
      Left            =   450
      TabIndex        =   3
      Top             =   3030
      Width           =   1035
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Item:"
      Height          =   225
      Index           =   1
      Left            =   780
      TabIndex        =   1
      Top             =   750
      Width           =   435
   End
   Begin VB.Label generalLabel 
      Caption         =   "Purchase Order Selection"
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
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3105
   End
End
Attribute VB_Name = "PurchaseOrderSelectionDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pokeys As Variant

Private Sub Form_Load()
    Me.ItemNumber = InventoryMaintenance.ItemNumber
    If Left(Me.ItemNumber, 3) <> "ZZZ" Then
        MsgBox "Error, should only have ZZZ items here, talk to Brian!"
    End If
    
    Me.QtyToOrder = "0"
    
    ReDim tabs(2) As Long
    tabs(0) = 24
    tabs(1) = tabs(0) + 54
    tabs(2) = tabs(1) + 300
    SetListTabStops Me.PurchaseOrderList.hwnd, tabs
    
    Me.PurchaseOrderList.Clear
    'Me.PurchaseOrderList.AddItem "Qty" & vbTab & "S/O #" & vbTab & "Bill To" & vbTab & "POID"
    
    'form-level, we'll need this later too
    pokeys = GetPurchaseOrderKeys(Me.ItemNumber)
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT PurchaseOrderLines.Quantity, PurchaseOrders.MiscNotes, PurchaseOrders.ShipToName, PurchaseOrders.ID FROM PurchaseOrders LEFT OUTER JOIN PurchaseOrderLines ON PurchaseOrders.ID=PurchaseOrderLines.HeaderID AND PurchaseOrderLines.ItemNumber='" & Me.ItemNumber & "' WHERE Exported=0 AND VendorNumber='" & pokeys(0) & "' AND CoopVendor='" & pokeys(1) & "' AND ProductLineCode='" & pokeys(2) & "'")
    While Not rst.EOF
        Dim ln As String
        ln = Nz(rst("Quantity"), "0") & vbTab
        
        Dim noteLines As Variant
        noteLines = Split(Nz(rst("MiscNotes")), vbCrLf)
        Dim i As Long, custOrderNum As String
        custOrderNum = ""
        For i = 0 To UBound(noteLines)
            If CBool(InStr(noteLines(i), "ORDER NUMBER")) Then
                custOrderNum = Right(noteLines(i), 7)
                Exit For
            End If
        Next i
        ln = ln & custOrderNum & vbTab
        
        ln = ln & Nz(rst("ShipToName")) & vbTab
        
        ln = ln & rst("ID")
        
        Me.PurchaseOrderList.AddItem ln
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    Me.PurchaseOrderList.AddItem "Create New Purchase Order..."
End Sub

Private Sub PurchaseOrderList_Click()
    If Me.PurchaseOrderList.ListIndex = Me.PurchaseOrderList.ListCount - 1 Then
        Me.QtyToOrder = 0
    Else
        Me.QtyToOrder = ListBoxLineColumn(Me.PurchaseOrderList, Me.PurchaseOrderList.ListIndex, 0)
    End If
End Sub

Private Sub QtyToOrder_GotFocus()
    Me.QtyToOrder.selStart = 0
    Me.QtyToOrder.SelLength = Len(Me.QtyToOrder)
End Sub

Private Sub QtyToOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        btnAddToPurchaseOrder_Click
    End If
End Sub

Private Sub QtyToOrder_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
End Sub

Private Sub btnAddToPurchaseOrder_Click()
    If Me.PurchaseOrderList.SelCount = 0 Then
        MsgBox "Select a PO!"
        Exit Sub
    End If
    
    'create a new PO option, make sure qty > 0
    If Me.PurchaseOrderList.ListIndex = Me.PurchaseOrderList.ListCount - 1 Then
        If Me.QtyToOrder = 0 Then
            MsgBox "New PO needs to have a quantity!"
            Exit Sub
        End If
    End If
    
    'done with validation, now run
    Dim id As Long
    'create a new PO, if required, otherwise just grab the id
    If Me.PurchaseOrderList.ListIndex = Me.PurchaseOrderList.ListCount - 1 Then
        'dupe of CreatePurchaseOrder()
        Dim Terms As String
        Terms = DLookup("TermsCode", "AP_Vendor", "VendorNo='" & pokeys(0) & "'")
        If Terms = "" Then
            Terms = "00"
        End If
        DB.Execute "INSERT INTO PurchaseOrders ( VendorNumber, CoopVendor, ProductLineCode, POTerms ) VALUES ( '" & pokeys(0) & "', '" & pokeys(1) & "', '" & pokeys(2) & "', '" & Terms & "' )"
        id = CLng(DLookup("@@IDENTITY", "PurchaseOrders"))
        
        'if creating a new d/s po, ask for order to fill address, etc
        If Left(Me.ItemNumber, 3) = "ZZZ" Then
            ConvertToDropshipPO id
        End If
    Else
        'Dim parts As Variant
        'parts = Split(Me.PurchaseOrderList.Text, vbTab)
        'id = CLng(parts(3))
        id = ListBoxLineColumn(Me.PurchaseOrderList, Me.PurchaseOrderList.ListIndex, 3)
    End If
    
    'po with id now exists, remove/add/edit item
    PurchaseOrderLineEdit CLng(id), Me.ItemNumber, CLng(Me.QtyToOrder)
    
    'done?
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
