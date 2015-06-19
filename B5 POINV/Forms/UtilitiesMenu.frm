VERSION 5.00
Begin VB.Form UtilitiesMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu - Utilities"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   4875
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnFAQBuilder 
      Caption         =   "FAQ Builder"
      Height          =   435
      Left            =   1695
      TabIndex        =   14
      Top             =   3000
      Width           =   1485
   End
   Begin VB.CommandButton btnAudioRunnerStandalone 
      Caption         =   "Audio Runner"
      Height          =   435
      Left            =   3330
      TabIndex        =   13
      Top             =   1800
      Width           =   1485
   End
   Begin VB.CommandButton btnBlogs 
      Caption         =   "Blogs"
      Enabled         =   0   'False
      Height          =   435
      Left            =   60
      TabIndex        =   12
      Top             =   3600
      Width           =   1485
   End
   Begin VB.CommandButton btnQuotes 
      Caption         =   "Quotes"
      Enabled         =   0   'False
      Height          =   435
      Left            =   60
      TabIndex        =   11
      Top             =   3000
      Width           =   1485
   End
   Begin VB.CommandButton btnMoreUtils 
      Caption         =   "More Utilities... (Rarely Used)"
      Height          =   435
      Left            =   3330
      TabIndex        =   10
      Top             =   3600
      Width           =   1485
   End
   Begin VB.CommandButton btnFreightQuoteStandalone 
      Caption         =   "Freight Calculator"
      Height          =   435
      Left            =   3330
      TabIndex        =   9
      Top             =   1200
      Width           =   1485
   End
   Begin VB.CommandButton btnGoToImportProductsMenu 
      Caption         =   "Import New Products"
      Height          =   435
      Left            =   1695
      TabIndex        =   8
      Top             =   2400
      Width           =   1485
   End
   Begin VB.CommandButton btnWebSiteFrontPage 
      Caption         =   "Website Front Page"
      Height          =   435
      Left            =   1695
      TabIndex        =   7
      Top             =   1800
      Width           =   1485
   End
   Begin VB.CommandButton btnBarcode 
      Caption         =   "Bar Code"
      Height          =   435
      Left            =   60
      TabIndex        =   6
      Top             =   1800
      Width           =   1485
   End
   Begin VB.CommandButton btnCustomerQuotes 
      Caption         =   "Customer Quotes"
      Height          =   435
      Left            =   60
      TabIndex        =   5
      Top             =   2400
      Width           =   1485
   End
   Begin VB.CommandButton btnPrintSigns 
      Caption         =   "Print Signs"
      Height          =   435
      Left            =   60
      TabIndex        =   1
      Top             =   1200
      Width           =   1485
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Back To Main Menu"
      Height          =   435
      Left            =   1230
      TabIndex        =   3
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton btnWebOffload 
      Caption         =   "Web Offload"
      Height          =   435
      Left            =   1695
      TabIndex        =   2
      Top             =   1200
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Utilities"
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
      Left            =   345
      TabIndex        =   0
      Top             =   0
      Width           =   4185
   End
   Begin VB.Label lblStatusBar 
      Height          =   225
      Left            =   30
      TabIndex        =   4
      Top             =   5430
      Width           =   4845
   End
End
Attribute VB_Name = "UtilitiesMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lockWebSiteFP As Boolean

Private Sub btnAudioRunnerStandalone_Click()
    'Shell AR_STANDALONE_APP, vbNormalFocus
    Shell AR_PLAYLIST_MAKER, vbNormalFocus
End Sub

Private Sub btnBarcode_Click()
    Load Barcode
    Barcode.Show
End Sub

Private Sub btnBlogs_Click()
    Load BlogManager
    BlogManager.Show
End Sub

Private Sub btnCustomerQuotes_Click()
    Load CustomerQuotes
    CustomerQuotes.Show
End Sub

Private Sub btnExit_Click()
    Unload UtilitiesMenu
End Sub

Private Sub btnFAQBuilder_Click()
    Load FAQBuilder
    FAQBuilder.Show
End Sub

Private Sub btnFreightQuoteStandalone_Click()
    Shell FQ_STANDALONE_APP, vbNormalFocus
End Sub

Private Sub btnGoToImportProductsMenu_Click()
    Load NewProdMenuForm
    NewProdMenuForm.Show , MainMenu
    UtilitiesMenu.Hide
End Sub

Private Sub btnMoreUtils_Click()
    Load UtilitiesMenuMore
    UtilitiesMenu.Hide
    UtilitiesMenuMore.Show
End Sub

Private Sub btnPrintSigns_Click()
    Load PrintSigns
    PrintSigns.Show
End Sub

Private Sub btnQuotes_Click()
    Load QuoteInquiry
    QuoteInquiry.Show
End Sub

'Private Sub btnReports_Click()
'    OpenReportsDB
'End Sub


Private Sub btnWebOffload_Click()
    Load WebOffload
    WebOffload.Show
End Sub

Private Sub btnWebSiteFrontPage_Click()
    If lockWebSiteFP Then
        MsgBox "You need to install ImageMagick (talk to Brian) before using this."
    Else
        Load WebSiteFrontPage
        WebSiteFrontPage.Show
    End If
End Sub

Private Sub Form_Load()
    If Not HasPermissionsTo("PrintSigns") Then
        Me.btnPrintSigns.Enabled = False
    End If
    If Not HasPermissionsTo("WebOffload") Then
        Me.btnWebOffload.Enabled = False
        Me.btnWebSiteFrontPage.Enabled = False
        Me.btnCustomerQuotes.Enabled = False
        Me.btnFAQBuilder.Enabled = False
    End If
    If Not HasPermissionsTo("BarCode") Then
        Me.btnBarcode.Enabled = False
    End If
    If Not HasPermissionsTo("NewProducts") Then
        Me.btnGoToImportProductsMenu.Enabled = False
    End If
    If Not HasPermissionsTo("RecordAnnouncements") Then
        Me.btnAudioRunnerStandalone.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MainMenu.Show
End Sub
