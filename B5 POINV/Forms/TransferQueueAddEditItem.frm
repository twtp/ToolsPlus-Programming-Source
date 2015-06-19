VERSION 5.00
Object = "{39C208C4-2615-4D49-911A-50F903B45C86}#2.0#0"; "TPControls.ocx"
Begin VB.Form TransferQueueAddEditItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox MaxQuantity 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1080
      Width           =   645
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete"
      Height          =   405
      Left            =   3210
      TabIndex        =   13
      Top             =   2670
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox id 
      Height          =   285
      Left            =   4080
      TabIndex        =   12
      Top             =   60
      Width           =   495
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   1770
      TabIndex        =   10
      Top             =   2670
      Width           =   1305
   End
   Begin VB.CommandButton Save 
      Caption         =   "Save"
      Height          =   405
      Left            =   390
      TabIndex        =   9
      Top             =   2670
      Width           =   1245
   End
   Begin VB.ComboBox ReasonNeeded 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1830
      Width           =   1785
   End
   Begin VB.ComboBox NeededByDatePart 
      Height          =   315
      ItemData        =   "TransferQueueAddEditItem.frx":0000
      Left            =   3360
      List            =   "TransferQueueAddEditItem.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1440
      Width           =   915
   End
   Begin TPControls.DateDropdown NeededByDate 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   1440
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
   End
   Begin VB.TextBox Quantity 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Text            =   "1"
      Top             =   1050
      Width           =   795
   End
   Begin VB.ComboBox ItemNumber 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   660
      Width           =   1875
   End
   Begin VB.Label Label5 
      Caption         =   "max:"
      Height          =   255
      Left            =   2220
      TabIndex        =   14
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label headerLabel 
      Caption         =   "Add Item"
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
      Left            =   150
      TabIndex        =   11
      Top             =   60
      Width           =   2115
   End
   Begin VB.Label Label4 
      Caption         =   "Reason:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1860
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "Needed By:"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1470
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Quantity:"
      Height          =   285
      Left            =   330
      TabIndex        =   2
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Item Number:"
      Height          =   255
      Left            =   270
      TabIndex        =   0
      Top             =   690
      Width           =   1005
   End
End
Attribute VB_Name = "TransferQueueAddEditItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub LoadLine(id As Long)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber, Quantity, NeededByTime, Reason FROM InventoryTransferQueue WHERE ID=" & id)
    If Not rst.EOF Then
        Me.id = id
        Me.headerLabel.Caption = "Edit Item"
        Me.Delete.Visible = True
        Me.ItemNumber = rst("ItemNumber")
        Me.Quantity = rst("Quantity")
        Me.NeededByDate.value = Left(rst("NeededByTime"), InStr(rst("NeededByTime"), " ") - 1)
        If InStr(rst("NeededByTime"), "AM") Then
            Me.NeededByDatePart = "AM"
        Else
            Me.NeededByDatePart = "PM"
        End If
        Me.ReasonNeeded = rst("Reason")
        setMaxQty rst("ItemNumber")
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.ReasonNeeded.Clear
    Me.ReasonNeeded.AddItem "Y Order"
    Me.ReasonNeeded.AddItem "Cust. Pickup"
    Me.ReasonNeeded.AddItem "Stock"
    Me.ReasonNeeded.AddItem "Other"
    
    Me.NeededByDatePart.Clear
    Me.NeededByDatePart.AddItem "AM"
    Me.NeededByDatePart.AddItem "PM"
    
    Dim i As Long
    For i = 0 To UBound(ITEMS_CACHE)
        Me.ItemNumber.AddItem ITEMS_CACHE(i)
    Next i
    ExpandDropDownToFit Me.ItemNumber
    
    Me.NeededByDate.value = Date
    Me.NeededByDatePart = Me.NeededByDatePart.List(0)
    
    Me.ReasonNeeded = Me.ReasonNeeded.List(0)
End Sub

Private Sub ItemNumber_Click()
    setMaxQty Me.ItemNumber
End Sub

Private Sub ItemNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.ItemNumber, KeyCode, Shift
End Sub

Private Sub ItemNumber_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.ItemNumber, KeyAscii
End Sub

Private Sub ItemNumber_LostFocus()
    AutoCompleteLostFocus Me.ItemNumber
    setMaxQty Me.ItemNumber
End Sub

Private Sub Quantity_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
        'ok
    ElseIf KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Then
        'ok
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Save_Click()
    If Me.ItemNumber = "" Or Me.Quantity = "" Then
        MsgBox "You didn't fill out the form"
    ElseIf Me.Quantity > Me.MaxQuantity Then
        MsgBox "The quantity you requested is more than the available quantity!"
        Me.Quantity.SetFocus
    Else
        If Me.id = "" Then
            DB.execute "INSERT INTO InventoryTransferQueue ( ItemNumber, Quantity, NeededByTime, Reason, Requester, RequestTime ) VALUES ( " & _
                         "'" & Me.ItemNumber & "', " & Me.Quantity & ", '" & Me.NeededByDate.value & " 12:00:01 " & Me.NeededByDatePart & "', '" & Me.ReasonNeeded & "', '" & Environ("UserName") & "', GETDATE() )"
        Else
            DB.execute "UPDATE InventoryTransferQueue SET ItemNumber='" & Me.ItemNumber & "', " & _
                                                          "Quantity=" & Me.Quantity & ", " & _
                                                          "NeededByTime='" & Me.NeededByDate.value & " 12:00:01 " & Me.NeededByDatePart & "', " & _
                                                          "Reason='" & Me.ReasonNeeded & "', " & _
                                                          "RequestTime=GETDATE()" & _
                                                    " WHERE ID=" & Me.id
        End If
        Unload Me
    End If
End Sub

Private Sub setMaxQty(item As String)
    Dim masrst As ADODB.Recordset
    Set masrst = MASDB.retrieve("SELECT QtyOnHand FROM IM2_InventoryItemWhseDetl WHERE ItemNumber='" & item & "' AND WhseCode='000'")
    If Not masrst.EOF Then
        Me.MaxQuantity = masrst("QtyOnHand")
    Else
        Me.MaxQuantity = "?"
    End If
    masrst.Close
    Set masrst = Nothing
End Sub
