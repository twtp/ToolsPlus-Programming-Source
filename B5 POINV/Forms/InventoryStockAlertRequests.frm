VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#2.2#0"; "TPControls.ocx"
Begin VB.Form InventoryStockAlertRequests 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Alert Requests"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7755
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox HideItemsOnPO 
      Caption         =   "Hide items currently on PO or being D/C"
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   3285
   End
   Begin TPControls.SimpleListView RequestList 
      Height          =   3495
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   6165
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Close"
      Height          =   405
      Left            =   6480
      TabIndex        =   1
      Top             =   30
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Stock Alert Requests"
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
      TabIndex        =   0
      Top             =   60
      Width           =   2775
   End
End
Attribute VB_Name = "InventoryStockAlertRequests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private orderby As String
Private orderdesc As Boolean

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    orderby = "ItemNumber"
    orderdesc = False
    Me.RequestList.SetColumnNames Array("ItemNumber", "# Requests", "Qty On BO", "Initial Days Ago", "Qty On PO", "Being D/C")
    Me.RequestList.SetColumnWidths Array(1500, 1000, 1000, 1500, 1000, 1000)
    requeryRequestList
End Sub

Private Sub HideItemsOnPO_Click()
    requeryRequestList
End Sub

Private Sub requeryRequestList()
    Dim rst As ADODB.Recordset
    Dim whereClause As String
    If Me.HideItemsOnPO Then
        whereClause = "WHERE QtyOnPO<=NumRequests+QtyOnBO AND Discontinued=0"
    End If
    Set rst = DB.retrieve("SELECT ItemNumber, NumRequests, QtyOnBO, InitialDaysAgo, QtyOnPO, CASE WHEN Discontinued=1 THEN 'Y' ELSE 'N' END AS BeingDC FROM vCustomerAlertsStockItemInfo " & whereClause & "ORDER BY " & orderby & IIf(orderdesc, " DESC", " ASC"))
    Me.RequestList.Clear
    Me.RequestList.Add rst, , True
    rst.Close
    Set rst = Nothing
End Sub

Private Sub RequestList_ColumnClicked(i As Long)
    Dim field As String
    Select Case i
        Case Is = 0
            field = "ItemNumber"
        Case Is = 1
            field = "NumRequests"
        Case Is = 2
            field = "QtyOnBO"
        Case Is = 3
            field = "InitialDaysAgo"
        Case Is = 4
            field = "QtyOnPO"
        Case Is = 5
            field = "BeingDC"
        Case Else
            field = "ItemNumber"
    End Select
    If field = orderby Then
        orderdesc = Not orderdesc
    Else
        orderdesc = False
    End If
    orderby = field
    requeryRequestList
End Sub

Private Sub RequestList_DblClick(i As Long, j As Long, txt As String)
    If IsFormLoaded("InventoryMaintenance") Then
        InventoryMaintenance.GoToItem Me.RequestList.GetTextSelected(0), True
        FocusOnForm "InventoryMaintenance"
    End If
End Sub

