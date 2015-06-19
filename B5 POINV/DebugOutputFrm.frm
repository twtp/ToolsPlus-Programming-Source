VERSION 5.00
Begin VB.Form DebugOutputFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Output"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton DebugFrmOKBtn 
      Caption         =   "OK"
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox DebugOutputTxt 
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   8175
   End
End
Attribute VB_Name = "DebugOutputFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub WriteOutput(OutputStr As String)

DebugOutputTxt.Text = OutputStr

End Sub

Private Sub DebugFrmOKBtn_Click()
Unload Me
End Sub

