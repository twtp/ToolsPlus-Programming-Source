VERSION 5.00
Begin VB.Form DiscontinueDSableDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Discontinue Item"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox ItemNumber 
      Height          =   285
      Left            =   4410
      TabIndex        =   4
      Text            =   "ItemNumber"
      Top             =   390
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton btnSkip 
      Caption         =   "Skip For Now"
      Height          =   465
      Left            =   3930
      TabIndex        =   2
      Top             =   720
      Width           =   1485
   End
   Begin VB.CommandButton btnDC 
      Caption         =   "Discontinue Completely"
      Height          =   465
      Left            =   2033
      TabIndex        =   1
      Top             =   720
      Width           =   1725
   End
   Begin VB.CommandButton btnKeep 
      Caption         =   "Keep At Web"
      Height          =   465
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "The ITEMNUMBER needs to be discontinued in the store, but is dropshippable. What should happen at the web?"
      Height          =   585
      Left            =   210
      TabIndex        =   3
      Top             =   60
      Width           =   5085
   End
End
Attribute VB_Name = "DiscontinueDSableDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ReturnValue As String

Private Sub btnDC_Click()
    SetNotDropShippable Me.ItemNumber
    HandleDiscontinuation Me.ItemNumber
    ReturnValue = "discontinue"
    Unload DiscontinueDSableDlg
End Sub

Private Sub btnKeep_Click()
    HandleDiscontinuation Me.ItemNumber
    ReturnValue = "dropship"
    Unload DiscontinueDSableDlg
End Sub

Private Sub btnSkip_Click()
    ReturnValue = "skip"
    Unload DiscontinueDSableDlg
End Sub

Private Sub ItemNumber_Change()
    Me.Label1 = Replace(Me.Label1, "ITEMNUMBER", Me.ItemNumber)
End Sub
