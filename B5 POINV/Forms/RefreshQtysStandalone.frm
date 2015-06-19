VERSION 5.00
Begin VB.Form RefreshQtysStandalone 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Refresh Quantities"
   ClientHeight    =   690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   1110
      TabIndex        =   1
      Top             =   390
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Caption         =   "STATUS"
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3555
   End
End
Attribute VB_Name = "RefreshQtysStandalone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mas As Mas200ImportExport
Attribute mas.VB_VarHelpID = -1
Private running As Boolean

Private Sub btnCancel_Click()
    ClearProcessFlag "RefreshQtys"
    Unload Me
    End
End Sub

Public Sub doIt()
    running = True
    Mouse.Hourglass True
    Set mas = New Mas200ImportExport
    mas.RefreshQuantities
    Set mas = Nothing
    Mouse.Hourglass False
    running = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If running Then
        If MsgBox("Are you sure you want to cancel?", vbYesNo) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub mas_StatusChanged(status As String)
    Me.lblStatus.Caption = status
    DoEvents
End Sub
