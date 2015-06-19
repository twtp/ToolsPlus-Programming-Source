VERSION 5.00
Begin VB.Form RestorePO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restore Exported Purchase Orders"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   405
      Left            =   3600
      TabIndex        =   7
      Top             =   3090
      Width           =   1215
   End
   Begin VB.CommandButton btnRestore 
      Caption         =   "Restore This PO"
      Height          =   375
      Left            =   330
      TabIndex        =   6
      Top             =   3090
      Width           =   1665
   End
   Begin VB.ListBox poList 
      Height          =   1815
      Left            =   60
      TabIndex        =   5
      Top             =   1050
      Width           =   4785
   End
   Begin VB.Label generalLabel 
      Caption         =   "Line Code"
      Height          =   255
      Index           =   4
      Left            =   3810
      TabIndex        =   4
      Top             =   780
      Width           =   885
   End
   Begin VB.Label generalLabel 
      Caption         =   "Vendor"
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   3
      Top             =   780
      Width           =   1155
   End
   Begin VB.Label generalLabel 
      Caption         =   "Date Exported"
      Height          =   255
      Index           =   2
      Left            =   870
      TabIndex        =   2
      Top             =   780
      Width           =   1155
   End
   Begin VB.Label generalLabel 
      Caption         =   "PO"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   780
      Width           =   495
   End
   Begin VB.Label generalLabel 
      Caption         =   "Restore Exported Purchase Orders"
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
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4425
   End
End
Attribute VB_Name = "RestorePO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload RestorePO
End Sub

Private Sub btnRestore_Click()
    Dim ponum As String
    ponum = Mid(Me.poList, InStrRev(Me.poList, vbTab) + 1)
    If MsgBox("Are you sure you want to restore PO " & Left(Me.poList, InStr(Me.poList, vbTab) - 1) & "?", vbYesNo) = vbYes Then
        DB.Execute "UPDATE PurchaseOrders SET Exported=0, DateExported=Null WHERE ID=" & ponum
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT ItemNumber, QtyOrdered FROM PurchaseOrdersExported WHERE ID=" & ponum)
        While Not rst.EOF
            DB.Execute "UPDATE InventoryMaster SET POID=" & ponum & ", QtyOrdered=" & rst("QtyOrdered") & " WHERE ItemNumber='" & rst("ItemNumber") & "'"
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
        DB.Execute "DELETE FROM PurchaseOrdersExported WHERE ID=" & ponum
        MsgBox "Restore successful."
    End If
End Sub

Private Sub Form_Load()
    ReDim tabs(3) As Long
    tabs(0) = 36
    tabs(1) = 90
    tabs(2) = 168
    tabs(3) = 360 'offscreen id
    SetListTabStops Me.poList.hwnd, tabs
    requeryPOs
End Sub

Private Sub requeryPOs()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT PONumber, DateExported, VendorNumber, CoopVendor, ProductLineCode, ID FROM PurchaseOrders WHERE Exported=1 AND PONumber='A24065' ORDER BY DateExported DESC")
    Me.poList.Clear
    While Not rst.EOF
        Me.poList.AddItem rst("PONumber") & vbTab & Format(rst("DateExported"), "Short Date") & vbTab & rst("VendorNumber") & IIf(rst("CoopVendor") <> rst("VendorNumber"), " / " & rst("CoopVendor"), "") & vbTab & rst("ProductLineCode") & vbTab & rst("ID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub poList_DblClick()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT TOP 11 ItemNumber, QtyOrdered FROM PurchaseOrdersExported WHERE ID=" & Mid(Me.poList, InStrRev(Me.poList, vbTab) + 1) & " ORDER BY ItemNumber")
    Dim msg As String
    Dim i As Long
    For i = 1 To 10
        If rst.EOF Then
            i = 10
        Else
            msg = msg & vbCrLf & rst("ItemNumber") & " ordered " & rst("QtyOrdered")
            rst.MoveNext
        End If
    Next i
    If Not rst.EOF Then
        msg = msg & vbCrLf & "...list continues."
    End If
    MsgBox "Items on purchase order: " & vbCrLf & msg
End Sub
