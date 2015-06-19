VERSION 5.00
Object = "{39C208C4-2615-4D49-911A-50F903B45C86}#2.0#0"; "TPControls.ocx"
Begin VB.Form TransferQueueMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory Transfer List"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TimerMinuteCount 
      Height          =   285
      Left            =   6510
      TabIndex        =   6
      Top             =   90
      Width           =   525
   End
   Begin VB.TextBox LastUpdatedTime 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5670
      TabIndex        =   5
      Text            =   "TIME"
      Top             =   4020
      Width           =   825
   End
   Begin VB.Timer Timer1 
      Left            =   5970
      Top             =   30
   End
   Begin VB.CommandButton RefreshList 
      Caption         =   "Refresh List"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   3900
      Width           =   2055
   End
   Begin VB.CommandButton AddNewItem 
      Caption         =   "Add New Item"
      Height          =   495
      Left            =   270
      TabIndex        =   2
      Top             =   3900
      Width           =   2055
   End
   Begin TPControls.SimpleListView TransferList 
      Height          =   3255
      Left            =   90
      TabIndex        =   0
      Top             =   510
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   5741
   End
   Begin VB.Label Label2 
      Caption         =   "Last updated:"
      Height          =   255
      Left            =   4620
      TabIndex        =   4
      Top             =   4020
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Inventory Transfer List"
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
      TabIndex        =   1
      Top             =   90
      Width           =   2895
   End
End
Attribute VB_Name = "TransferQueueMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DELAY_MINS = 15

Private Sub AddNewItem_Click()
    Load TransferQueueAddEditItem
    TransferQueueAddEditItem.Show 1
    doRefresh
End Sub

Private Sub Form_Load()
    Set DB = New DBConn.DatabaseSingleton
    DB.CurrentDB = DBENG_MSSQL
    Set MASDB = New DBConn.DatabaseSingleton
    MASDB.CurrentDB = DBENG_MAS90
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber FROM InventoryMaster WHERE ItemNumber<'XX%' ORDER BY ItemNumber")
    ReDim ITEMS_CACHE(rst.RecordCount - 1) As String
    Dim i As Long
    i = 0
    While Not rst.EOF
        ITEMS_CACHE(i) = CStr(rst("ItemNumber"))
        i = i + 1
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing

    Me.TimerMinuteCount = DELAY_MINS
    Me.Timer1.Interval = 60000 '1 min
    Me.Timer1.Enabled = True

    Me.TransferList.SetColumnNames Array("ID", "ItemNumber", "Quantity", "Needed By", "Reason", "Requester") ', "Request Time")
    Me.TransferList.SetColumnWidths Array("0", "1600", "900", "1500", "1500", "1100") ', "1500")
    
    doRefresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set DB = Nothing
End Sub

Private Sub RefreshList_Click()
    doRefresh
End Sub



Private Sub doRefresh()
    Me.Timer1.Enabled = False
    
    'Me.TransferList.Add Array("DEL36-L31X-BC50", "1", "12/31/2007 AM", "Cust. Pickup", "Bob") ', "3/5/2007 12:00:00 AM")
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, ItemNumber, Quantity, Status, NeededByTime, Reason, Requester FROM InventoryTransferQueue WHERE Received=0 ORDER BY NeededByTime")
    Me.TransferList.Clear
    Dim nt As String
    While Not rst.EOF
        nt = Left(rst("NeededByTime"), InStr(rst("NeededByTime"), " "))
        nt = Format(nt, "ddd, m/d")
        If InStr(rst("NeededByTime"), "AM") Then
            nt = nt & " AM"
        Else
            nt = nt & " PM"
        End If
        Me.TransferList.Add Array(rst("ID"), rst("ItemNumber"), rst("Quantity"), nt, rst("Reason"), rst("Requester"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    Me.LastUpdatedTime = Time()
    Me.TimerMinuteCount = DELAY_MINS
    Me.Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Me.TimerMinuteCount = Me.TimerMinuteCount - 1
    If Me.TimerMinuteCount = 0 Then
        doRefresh
    End If
End Sub

Private Sub TransferList_DblClick(i As Long, j As Long, txt As String)
    Dim user As String
    user = Me.TransferList.GetText(i, 5)
    If LCase(user) = LCase(Environ("UserName")) Then
        Load TransferQueueAddEditItem
        TransferQueueAddEditItem.LoadLine Me.TransferList.GetText(i, 0)
        TransferQueueAddEditItem.Show 1
        doRefresh
    End If
End Sub
