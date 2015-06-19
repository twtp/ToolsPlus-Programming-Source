VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl DateDropdown 
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2625
   ScaleHeight     =   735
   ScaleWidth      =   2625
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16187393
      CurrentDate     =   39093
   End
End
Attribute VB_Name = "DateDropdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Click()

Public Property Get Value() As String
Attribute Value.VB_MemberFlags = "200"
    Value = UserControl.DTPicker1.Value
End Property

Public Property Let Value(newDate As String)
    UserControl.DTPicker1.Value = newDate
    PropertyChanged "Value"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.DTPicker1.Enabled
End Property

Public Property Let Enabled(state As Boolean)
    UserControl.DTPicker1.Enabled = state
    PropertyChanged "Enabled"
End Property

Private Sub DTPicker1_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Resize()
    UserControl.DTPicker1.Width = UserControl.Width
    UserControl.DTPicker1.Height = UserControl.Height
End Sub
