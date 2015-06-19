VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl DateMonthview 
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   ScaleHeight     =   2700
   ScaleWidth      =   3495
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   15990785
      CurrentDate     =   39097
   End
End
Attribute VB_Name = "DateMonthview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Click(selectedDate As String)

Public Property Get Value() As String
Attribute Value.VB_MemberFlags = "200"
    Value = UserControl.MonthView1.Value
End Property

Public Property Let Value(newDate As String)
    UserControl.MonthView1.Value = newDate
    PropertyChanged "Value"
End Property

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    RaiseEvent Click(CStr(DateClicked))
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 2700
    UserControl.Height = 2370
End Sub
