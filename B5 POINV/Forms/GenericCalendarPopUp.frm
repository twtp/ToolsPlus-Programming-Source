VERSION 5.00
Object = "{39C208C4-2615-4D49-911A-50F903B45C86}#1.3#0"; "TPControls.ocx"
Begin VB.Form GenericCalendarPopUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Date"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   3165
   StartUpPosition =   1  'CenterOwner
   Begin TPControls.DateMonthview DateMonthview1 
      Height          =   2370
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
   End
   Begin VB.TextBox ctlName 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Text            =   "ctl"
      Top             =   1260
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox formName 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Text            =   "form"
      Top             =   960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1620
      TabIndex        =   1
      Top             =   2610
      Width           =   1425
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   90
      TabIndex        =   0
      Top             =   2610
      Width           =   1425
   End
End
Attribute VB_Name = "GenericCalendarPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SetUpCalendar(defaultDate As String, formName As String, controlname As String)
    Me.DateMonthview1.value = defaultDate
    Me.formName = formName
    Me.ctlName = controlname
End Sub

Private Sub btnCancel_Click()
    Unload GenericCalendarPopUp
End Sub

Private Sub btnOK_Click()
    SetControlValue Me.formName, Me.ctlName, Me.DateMonthview1.value
    Unload GenericCalendarPopUp
End Sub

Private Sub Form_Load()
    Me.DateMonthview1.value = Date
End Sub
