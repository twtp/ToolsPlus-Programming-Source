VERSION 5.00
Begin VB.Form ProductSuggestionSectionDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Section"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   465
      Left            =   2430
      TabIndex        =   6
      Top             =   3870
      Width           =   1365
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   465
      Left            =   840
      TabIndex        =   5
      Top             =   3870
      Width           =   1365
   End
   Begin VB.TextBox SectionName 
      Height          =   315
      Left            =   540
      TabIndex        =   4
      Top             =   3270
      Width           =   3375
   End
   Begin VB.ListBox SectionNameList 
      Height          =   2010
      Left            =   540
      TabIndex        =   3
      Top             =   900
      Width           =   3375
   End
   Begin VB.OptionButton opSectionType 
      Caption         =   "New Section Name"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   2
      Top             =   3000
      Width           =   2445
   End
   Begin VB.OptionButton opSectionType 
      Caption         =   "Preset Section Name"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   630
      Width           =   2445
   End
   Begin VB.Label Label1 
      Caption         =   "Product Suggestions - Add Section"
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
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4245
   End
End
Attribute VB_Name = "ProductSuggestionSectionDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOK_Click()
    If Me.opSectionType(0) Then
        If Me.SectionNameList.ListIndex = -1 Then
            MsgBox "Select a section!"
        Else
            ProductSuggestionUtility.SetIncomingSectionName Me.SectionNameList
            Unload Me
        End If
    Else
        If Me.SectionName = "" Then
            MsgBox "Enter a section name!"
        Else
            ProductSuggestionUtility.SetIncomingSectionName Me.SectionName
            Unload Me
        End If
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.SectionNameList.AddItem "Under $10"
    Me.SectionNameList.AddItem "Under $50"
    Me.SectionNameList.AddItem "Over $50"
    Me.SectionNameList.AddItem "Factory Reconditioned Deals"
    Me.SectionNameList.AddItem "Christmas Gift Ideas"
    Me.SectionNameList.AddItem "Father's Day"
    Me.SectionNameList.AddItem "Our Top Sellers"
    Me.SectionNameList.AddItem "New at Tools Plus"
    Me.SectionNameList.AddItem "Rebates and Bonuses"
    Me.SectionNameList.AddItem "Price Drops!"
    Me.opSectionType(0) = 1
End Sub

Private Sub opSectionType_Click(Index As Integer)
    If Index = 0 Then
        Me.SectionNameList.Enabled = True
        Disable Me.SectionName
    Else
        Me.SectionNameList.Enabled = False
        Enable Me.SectionName
        Me.SectionName.SetFocus
    End If
End Sub
