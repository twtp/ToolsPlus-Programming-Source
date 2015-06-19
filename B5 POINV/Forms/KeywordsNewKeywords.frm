VERSION 5.00
Begin VB.Form KeywordsNewKeywords 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Keywords - Create New Keywords"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnExit 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1290
      TabIndex        =   3
      Top             =   1980
      Width           =   1635
   End
   Begin VB.CommandButton btnListTool 
      Caption         =   "Use the keyword association/list tool"
      Height          =   525
      Left            =   90
      TabIndex        =   2
      Top             =   1200
      Width           =   4035
   End
   Begin VB.CommandButton btnPhraseKeywords 
      Caption         =   "Scan for new PHRASE keywords"
      Height          =   525
      Left            =   90
      TabIndex        =   1
      Top             =   630
      Width           =   4035
   End
   Begin VB.CommandButton btnItemKeywords 
      Caption         =   "Scan for new ITEM keywords"
      Height          =   525
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   4035
   End
End
Attribute VB_Name = "KeywordsNewKeywords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload KeywordsNewKeywords
End Sub

Private Sub btnItemKeywords_Click()
    Load KeywordsNewItems
    Unload KeywordsNewKeywords
    KeywordsNewItems.Show MODAL
End Sub

Private Sub btnListTool_Click()
    Load KeywordsAssociation
    Unload KeywordsNewKeywords
    KeywordsAssociation.Show MODAL
End Sub

Private Sub btnPhraseKeywords_Click()
    Load KeywordsNewPhrases
    Unload KeywordsNewKeywords
    KeywordsNewPhrases.Show MODAL
End Sub
