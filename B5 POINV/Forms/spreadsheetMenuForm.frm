VERSION 5.00
Begin VB.Form NewProdMenuForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu - New Products Spreadsheet Import"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   4875
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnExit 
      Caption         =   "Back To Utilities Menu"
      Height          =   435
      Left            =   1230
      TabIndex        =   4
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton btnOpenExcel 
      Caption         =   "Open Excel Spreadsheet"
      Height          =   435
      Left            =   1230
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton btnView 
      Caption         =   "View Current Items"
      Height          =   435
      Left            =   1230
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton btnImport 
      Caption         =   "Import New Spreadsheet"
      Height          =   435
      Left            =   1230
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblStatusBar 
      Height          =   225
      Left            =   30
      TabIndex        =   5
      Top             =   5430
      Width           =   4785
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "New Products Spreadsheet Import"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   4185
   End
End
Attribute VB_Name = "NewProdMenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload NewProdMenuForm
End Sub

Private Sub btnImport_Click()
    ImportNewProductsSpreadsheet
End Sub

Private Sub btnOpenExcel_Click()
    OpenNewProductsSpreadsheet
End Sub

Private Sub btnView_Click()
    Load NewProdDetailForm
    NewProdDetailForm.Show
    'NewProdMenuForm.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UtilitiesMenu.Show
End Sub
