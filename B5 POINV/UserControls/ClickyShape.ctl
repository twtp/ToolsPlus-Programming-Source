VERSION 5.00
Begin VB.UserControl ClickyShape 
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2955
   ScaleHeight     =   2955
   ScaleWidth      =   2955
   Begin VB.Shape Shape1 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   2625
      Left            =   0
      Shape           =   5  'Rounded Square
      Top             =   0
      Width           =   2625
   End
End
Attribute VB_Name = "ClickyShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Property Get FillColor() As Long
    FillColor = UserControl.Shape1.FillColor
End Property

Public Property Let FillColor(newColor As Long)
    UserControl.Shape1.FillColor = newColor
    PropertyChanged "FillColor"
End Property

Public Property Get Shape() As Long
    Shape = UserControl.Shape1.Shape
End Property

Public Property Let Shape(newShape As Long)
    UserControl.Shape1.Shape = newShape
    PropertyChanged "Shape"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
    UserControl.Shape1.Width = UserControl.Width
    UserControl.Shape1.Height = UserControl.Height
End Sub
