VERSION 5.00
Begin VB.Form ShippingExtrasDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shipping Extras"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton AddNewButton 
      Caption         =   "Add New"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3540
      Width           =   2535
   End
   Begin VB.TextBox LoadedExtraID 
      Height          =   315
      Left            =   4050
      TabIndex        =   12
      Text            =   "LoadedExtraID"
      Top             =   150
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   495
      Left            =   5970
      TabIndex        =   11
      Top             =   0
      Width           =   1365
   End
   Begin VB.Frame Frame 
      Enabled         =   0   'False
      Height          =   3045
      Left            =   2730
      TabIndex        =   4
      Top             =   780
      Width           =   4575
      Begin VB.TextBox PopupExtension 
         Height          =   1815
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1170
         Width           =   3525
      End
      Begin VB.TextBox PITExtension 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   870
         Width           =   3525
      End
      Begin VB.TextBox ShortDescription 
         Height          =   285
         Left            =   930
         TabIndex        =   6
         Top             =   300
         Width           =   3555
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Add to Popup (HTML):"
         Height          =   735
         Index           =   4
         Left            =   180
         TabIndex        =   9
         Top             =   1230
         Width           =   735
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Add to PIT:"
         Height          =   225
         Index           =   3
         Left            =   60
         TabIndex        =   7
         Top             =   900
         Width           =   855
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Description:"
         Height          =   225
         Index           =   2
         Left            =   60
         TabIndex        =   5
         Top             =   330
         Width           =   855
      End
   End
   Begin VB.ListBox ShippingExtrasList 
      Height          =   2400
      Left            =   90
      TabIndex        =   3
      Top             =   1110
      Width           =   2565
   End
   Begin VB.TextBox VendorNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   750
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   780
      Width           =   1875
   End
   Begin VB.Label generalLabel 
      Caption         =   "Vendor:"
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   810
      Width           =   585
   End
   Begin VB.Label generalLabel 
      Caption         =   "Add / Edit Shipping Extras"
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
      Width           =   3435
   End
End
Attribute VB_Name = "ShippingExtrasDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private changed As Boolean



Private Sub Form_Load()
    changed = False
    Dim tabs(0) As Long
    tabs(0) = 300
    SetListTabStops Me.ShippingExtrasList.hwnd, tabs

    Dim item As String
    item = SignMaintenance.itemNumber
    Me.VendorNumber = DLookup("PrimaryVendor", "InventoryMaster", "ItemNumber='" & item & "'")
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, ShortDescription FROM ShippingExtras WHERE VendorNumber='" & EscapeSQuotes(Me.VendorNumber) & "' ORDER BY ShortDescription")
    While Not rst.EOF
        Me.ShippingExtrasList.AddItem rst("ShortDescription") & vbTab & rst("ID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub ShippingExtrasList_Click()
    Dim id As String
    id = ListBoxLineColumn(Me.ShippingExtrasList, Me.ShippingExtrasList.ListIndex, 1)
    Me.LoadedExtraID = id
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ShortDescription, PITExtension, PopupExtension FROM ShippingExtras WHERE ID=" & id)
    If rst.EOF Then
        MsgBox "Error retrieving!"
    Else
        Me.ShortDescription = rst("ShortDescription")
        Me.PITExtension = Nz(rst("PITExtension"))
        Me.PopupExtension = Nz(rst("PopupExtension"))
    End If
    rst.Close
    Set rst = Nothing
    Me.Frame.Enabled = True
End Sub

Private Sub AddNewButton_Click()
    Dim Desc As String
    Desc = MsgBox("Enter description for new shipping policy to add to vendor " & qq(Me.VendorNumber) & ":")
    If Desc <> "" Then
        DB.Execute "INSERT INTO ShippingExtras ( VendorNumber, ShortDescription ) VALUES ( '" & EscapeSQuotes(Me.VendorNumber) & "', '" & EscapeSQuotes(Desc) & "' )"
        Me.ShippingExtrasList.AddItem Desc & vbTab & DLookup("@@IDENTITY", "ShippingExtras")
        Me.ShippingExtrasList.ListIndex = Me.ShippingExtrasList.ListCount - 1
    End If
End Sub

Private Sub ShortDescription_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            ShortDescription_LostFocus
        Case Is = vbKeyDelete
            ShortDescription_KeyPress KeyCode
    End Select
End Sub

Private Sub ShortDescription_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub ShortDescription_LostFocus()
    If changed = True Then
        If Len(Me.ShortDescription) > 20 Then
            Me.ShortDescription = Left(Me.ShortDescription, 20)
        End If
        DB.Execute "UPDATE ShippingExtras SET ShortDescription='" & EscapeSQuotes(Me.ShortDescription) & "' WHERE ID=" & Me.LoadedExtraID
        changed = False
        
        Dim i As Long
        For i = 0 To Me.ShippingExtrasList.ListCount - 1
            Dim id As String
            id = ListBoxLineColumn(Me.ShippingExtrasList, i, 1)
            If id = Me.LoadedExtraID Then
                Me.ShippingExtrasList.list(i) = Me.ShortDescription & vbTab & id
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub PITExtension_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            PITExtension_LostFocus
        Case Is = vbKeyDelete
            PITExtension_KeyPress KeyCode
    End Select
End Sub

Private Sub PITExtension_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub PITExtension_LostFocus()
    If changed = True Then
        If Len(Me.PITExtension) > 128 Then
            Me.PITExtension = Left(Me.PITExtension, 128)
        End If
        DB.Execute "UPDATE ShippingExtras SET PITExtension='" & EscapeSQuotes(Me.PITExtension) & "' WHERE ID=" & Me.LoadedExtraID
        changed = False
    End If
End Sub

Private Sub PopupExtension_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            PopupExtension_LostFocus
        Case Is = vbKeyDelete
            PopupExtension_KeyPress KeyCode
    End Select
End Sub

Private Sub PopupExtension_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub PopupExtension_LostFocus()
    If changed = True Then
        DB.Execute "UPDATE ShippingExtras SET PopupExtension='" & EscapeSQuotes(Me.PopupExtension) & "' WHERE ID=" & Me.LoadedExtraID
        changed = False
    End If
End Sub
