VERSION 5.00
Object = "{39C208C4-2615-4D49-911A-50F903B45C86}#3.2#0"; "TPControls.ocx"
Begin VB.Form Receiving 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receiving"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton LoadJob 
      Caption         =   "Load Job"
      Height          =   645
      Left            =   4680
      TabIndex        =   10
      Top             =   1110
      Width           =   735
   End
   Begin VB.TextBox JobID 
      Height          =   285
      Left            =   3210
      TabIndex        =   9
      Text            =   "JobID"
      Top             =   60
      Width           =   1065
   End
   Begin TPControls.SimpleListView ItemsList 
      Height          =   4155
      Left            =   240
      TabIndex        =   8
      Top             =   2370
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7329
   End
   Begin VB.CommandButton ScanForMatchingPO 
      Caption         =   "Scan For Matching PO(s)"
      Height          =   315
      Left            =   5850
      TabIndex        =   7
      Top             =   1050
      Width           =   2955
   End
   Begin VB.TextBox PONumber 
      Height          =   315
      Left            =   7560
      TabIndex        =   6
      Top             =   1470
      Width           =   1215
   End
   Begin VB.ListBox JobList 
      Height          =   1035
      Left            =   150
      TabIndex        =   4
      Top             =   900
      Width           =   4485
   End
   Begin VB.Label generalLabel 
      Caption         =   "Load by PO Number:"
      Height          =   255
      Index           =   4
      Left            =   5880
      TabIndex        =   5
      Top             =   1530
      Width           =   1635
   End
   Begin VB.Label generalLabel 
      Caption         =   "Warehouse"
      Height          =   225
      Index           =   3
      Left            =   3690
      TabIndex        =   3
      Top             =   660
      Width           =   945
   End
   Begin VB.Label generalLabel 
      Caption         =   "Started By"
      Height          =   225
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      Top             =   660
      Width           =   855
   End
   Begin VB.Label generalLabel 
      Caption         =   "Name"
      Height          =   225
      Index           =   1
      Left            =   210
      TabIndex        =   1
      Top             =   660
      Width           =   645
   End
   Begin VB.Label generalLabel 
      Caption         =   "Receiving"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2415
   End
End
Attribute VB_Name = "Receiving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.ItemsList.SetColumnNames Array("ItemNumber", "Qty Scanned", "PO #", "Qty on PO")
    Me.ItemsList.SetColumnWidths Array("1500", "1000", "1000", "1000")
    
    requeryJobList
End Sub

Private Sub LoadJob_Click()
    Me.JobID = Mid(Me.JobList, InStrRev(Me.JobList, vbTab) + 1)
    requeryItemsList
End Sub


Private Sub requeryJobList()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, Name, ShortName, WarehouseName FROM vReceivingMgmtJobs")
    Me.JobList.Clear
    While Not rst.EOF
        Me.JobList.AddItem rst("Name") & vbTab & rst("ShortName") & vbTab & rst("WarehouseName") & vbTab & rst("ID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryItemsList()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber, Quantity FROM HandheldReceivingJobDetailLines WHERE JobID=" & Me.JobID)
    Me.ItemsList.Clear
    While Not rst.EOF
    
    Wend
    rst.Close
    Set rst = Nothing
End Sub
