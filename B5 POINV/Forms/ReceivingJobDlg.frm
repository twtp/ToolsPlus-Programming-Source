VERSION 5.00
Begin VB.Form ReceivingJobDlg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receiving - Select Job"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox JobList 
      Height          =   1035
      Left            =   90
      TabIndex        =   0
      Top             =   960
      Width           =   4305
   End
End
Attribute VB_Name = "ReceivingJobDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub requeryJobList()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, Name, ShortName, WarehouseName, QtyScanned FROM vReceivingMgmtJobs")
    Me.JobList.Clear
    While Not rst.EOF
        Me.JobList.AddItem rst("Name") & vbTab & rst("ShortName") & vbTab & rst("WarehouseName") & vbTab & rst("QtyScanned") & vbTab & rst("ID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub Form_Load()
    requeryJobList
End Sub
