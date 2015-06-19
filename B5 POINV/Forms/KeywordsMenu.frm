VERSION 5.00
Begin VB.Form KeywordsMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu - Keywords"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   4875
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnExportToAdWords 
      Caption         =   "Export to AdWords"
      Height          =   435
      Left            =   1230
      TabIndex        =   9
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton btnViewToBeDeleted 
      Caption         =   "To Be Deleted"
      Height          =   435
      Left            =   2970
      TabIndex        =   8
      Top             =   3000
      Width           =   675
   End
   Begin VB.CommandButton btnExportClearFlags 
      Caption         =   "Clear Flags"
      Height          =   435
      Left            =   2280
      TabIndex        =   7
      Top             =   3000
      Width           =   675
   End
   Begin VB.CommandButton btnExport 
      Caption         =   "Export Changes"
      Height          =   435
      Left            =   1230
      TabIndex        =   6
      Top             =   3000
      Width           =   1035
   End
   Begin VB.CommandButton btnCostRev 
      Caption         =   "Cost / Revenue Analysis"
      Height          =   435
      Left            =   1230
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton btnImport 
      Caption         =   "Import New Data"
      Height          =   435
      Left            =   1230
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Back to Main Menu"
      Height          =   435
      Left            =   1230
      TabIndex        =   4
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton btnPhraseList 
      Caption         =   "Keyword Maintenance"
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
      Width           =   4545
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Keywords"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   345
      TabIndex        =   0
      Top             =   0
      Width           =   4185
   End
End
Attribute VB_Name = "KeywordsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private downloadedOK As Boolean

Public Sub SetDownloadedOK(tf As Boolean)
    downloadedOK = tf
End Sub

Private Sub btnCostRev_Click()
    Load KeywordsCostRevenue
    KeywordsCostRevenue.Show
End Sub

Private Sub btnExit_Click()
    Unload KeywordsMenu
End Sub

Private Sub btnExport_Click()
    Shell PERL & " " & KWD_Y_OFFLOAD, vbNormalFocus
End Sub

Private Sub btnExportClearFlags_Click()
    'sign in to overture, and download the files
    Load KeywordsClearOvertureFlags
    KeywordsClearOvertureFlags.Show MODAL
    If downloadedOK Then
        Shell PERL & " " & KWD_O_CLEAR_FLAGS, vbNormalFocus
    End If
End Sub

Private Sub btnExportToAdWords_Click()
    ExportDataToAdwords
End Sub

Private Sub btnImport_Click()
    'Load KeywordsImportTool
    'KeywordsImportTool.Show
    Load KeywordsImportData
    KeywordsImportData.Show
End Sub

Private Sub btnPhraseList_Click()
    Load KeywordsPhraseList
    KeywordsPhraseList.Show
End Sub

Private Sub btnViewToBeDeleted_Click()
    Load KeywordsToBeDeleted
    KeywordsToBeDeleted.Show
End Sub

Private Sub Form_Load()
    'do some cleanup before we do anything
    DB.Execute "EXEC spKeywordsChildrenCleanup"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MainMenu.Show
End Sub
