VERSION 5.00
Begin VB.Form TransferReconciliationExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer Reconciliation - Export to MAS 200"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Exit 
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   30
      Width           =   1035
   End
   Begin VB.CommandButton ExportButton 
      Caption         =   "EXPORT!!!1!!"
      Height          =   1245
      Left            =   4020
      TabIndex        =   6
      Top             =   1950
      Width           =   1515
   End
   Begin VB.TextBox ToWhse 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   900
      Width           =   945
   End
   Begin VB.TextBox FromWhse 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   570
      Width           =   945
   End
   Begin VB.ListBox ItemList 
      Height          =   2985
      Left            =   90
      TabIndex        =   0
      Top             =   1260
      Width           =   3825
   End
   Begin VB.Label Label3 
      Caption         =   "To:"
      Height          =   225
      Left            =   390
      TabIndex        =   4
      Top             =   930
      Width           =   525
   End
   Begin VB.Label Label2 
      Caption         =   "From:"
      Height          =   225
      Left            =   390
      TabIndex        =   2
      Top             =   600
      Width           =   525
   End
   Begin VB.Label Label1 
      Caption         =   "Transfer Reconciliation - Export"
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
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   3885
   End
End
Attribute VB_Name = "TransferReconciliationExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub LoadLocations(fromWhseCode As String, toWhseCode As String)
    Me.FromWhse = fromWhseCode
    Me.ToWhse = toWhseCode
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber, Quantity FROM ExportInventoryTransfers WHERE Exported=0 AND FromWhse='" & fromWhseCode & "' AND ToWhse='" & toWhseCode & "'")
    Me.ItemList.Clear
    While Not rst.EOF
        Me.ItemList.AddItem rst("ItemNumber") & vbTab & rst("Quantity")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub ExportButton_Click()
    Dim refno As String
    refno = DLookup("MAX(ReferenceNumber)", "ExportInventoryTransfers")
    If refno = "" Then
        refno = "0"
    End If
    refno = Format(CLng(refno) + 1, "00000000")
    DB.Execute "UPDATE ExportInventoryTransfers SET ReferenceNumber='" & refno & "', DateTimeExported=GETDATE() WHERE Exported=0 AND FromWhse='" & Me.FromWhse & "' AND ToWhse='" & Me.ToWhse & "'"
    Shell "s:\mastest\mas90-signs\mas200_import_transfer.bat"
    MsgBox "Wait until MAS 200 window is finished before hitting OK"
    If vbYes = MsgBox("Did the import finish successfully?", vbYesNo) Then
        DB.Execute "UPDATE ExportInventoryTransfers SET Exported=1 WHERE ReferenceNumber=" & CLng(refno)
    Else
        MsgBox "That sucks. You should probably talk to Brian."
    End If
End Sub
