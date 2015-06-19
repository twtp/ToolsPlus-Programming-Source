VERSION 5.00
Begin VB.Form TransferRequestHist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer Request - History Drilldown"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8565
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox ItemNumber 
      Height          =   315
      Left            =   390
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   600
      Width           =   2445
   End
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   6840
      TabIndex        =   9
      Top             =   60
      Width           =   1635
      Begin VB.OptionButton ViewOption 
         Caption         =   "View History"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   11
         Top             =   510
         Width           =   1275
      End
      Begin VB.OptionButton ViewOption 
         Caption         =   "View Current"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   10
         Top             =   210
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   495
      Left            =   3780
      TabIndex        =   5
      Top             =   3960
      Width           =   1305
   End
   Begin VB.ListBox HistoryList 
      Height          =   2595
      Left            =   90
      TabIndex        =   1
      Top             =   1290
      Width           =   8385
   End
   Begin VB.Label Label7 
      Caption         =   "Notes"
      Height          =   225
      Left            =   5820
      TabIndex        =   8
      Top             =   1050
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Action"
      Height          =   225
      Left            =   3270
      TabIndex        =   7
      Top             =   1050
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Quantity"
      Height          =   225
      Left            =   4440
      TabIndex        =   6
      Top             =   1050
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Priority"
      Height          =   225
      Left            =   5130
      TabIndex        =   4
      Top             =   1050
      Width           =   555
   End
   Begin VB.Label Label3 
      Caption         =   "Person"
      Height          =   225
      Left            =   2160
      TabIndex        =   3
      Top             =   1050
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Date / Time"
      Height          =   225
      Left            =   150
      TabIndex        =   2
      Top             =   1050
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "Transfer Request - History Drilldown"
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
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4455
   End
End
Attribute VB_Name = "TransferRequestHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FromWhse As Long
Private ToWhse As Long

Private Sub Form_Load()
    Dim tabs(4) As Long
    tabs(0) = 90            'date/time
    tabs(1) = tabs(0) + 48  'person
    tabs(2) = tabs(1) + 54  'action
    tabs(3) = tabs(2) + 30  'quantity
    tabs(4) = tabs(3) + 30  'priority
    SetListTabStops Me.HistoryList.hwnd, tabs
    requeryItems
End Sub

Public Sub SetRequestHistoryParams(item As String, fromWhseID As Long, toWhseID As Long)
    FromWhse = fromWhseID
    ToWhse = toWhseID
    
    Me.ItemNumber = item
End Sub

Private Sub requeryList()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT DateTimeEntered, ShortName, Priority, Quantity, Action, Notes FROM vTransferRequestHistory WHERE ItemNumber='" & Me.ItemNumber & "' AND FromWhseID=" & FromWhse & " AND ToWhseID=" & ToWhse & " AND ArchiveFlag=" & IIf(Me.ViewOption(0) = True, "0", "1") & " ORDER BY DateTimeEntered DESC")
    Me.HistoryList.Clear
    While Not rst.EOF
        Me.HistoryList.AddItem rst("DateTimeEntered") & vbTab & rst("ShortName") & vbTab & rst("Action") & vbTab & rst("Quantity") & vbTab & rst("Priority") & vbTab & Nz(rst("Notes"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub ItemNumber_Change()
    requeryList
End Sub

Private Sub ItemNumber_Click()
    requeryList
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub

Private Sub ViewOption_Click(Index As Integer)
    requeryList
End Sub

Private Sub requeryItems()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber FROM InventoryMaster ORDER BY ItemNumber")
    Me.ItemNumber.Clear
    While Not rst.EOF
        Me.ItemNumber.AddItem rst("ItemNumber")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub
