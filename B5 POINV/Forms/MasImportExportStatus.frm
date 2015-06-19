VERSION 5.00
Begin VB.Form MasImportExportStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mas200 Import/Export Status"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6450
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblStatusBar 
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   2490
      Width           =   6015
   End
   Begin VB.Label lblImport 
      Caption         =   "Updating Sales History..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   3150
      TabIndex        =   10
      Top             =   1410
      Width           =   3000
   End
   Begin VB.Label lblImport 
      Caption         =   "Importing Warehouses..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   90
      TabIndex        =   9
      Top             =   1740
      Width           =   3000
   End
   Begin VB.Label lblImport 
      Caption         =   "Importing Vendors..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   90
      TabIndex        =   8
      Top             =   1410
      Width           =   3000
   End
   Begin VB.Label lblImport 
      Caption         =   "Importing UDFs..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   3150
      TabIndex        =   7
      Top             =   1080
      Width           =   3000
   End
   Begin VB.Label lblImport 
      Caption         =   "Importing Prices..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   3150
      TabIndex        =   6
      Top             =   750
      Width           =   3000
   End
   Begin VB.Label lblImport 
      Caption         =   "Importing Quantities..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   3150
      TabIndex        =   5
      Top             =   420
      Width           =   3000
   End
   Begin VB.Label lblImport 
      Caption         =   "Importing Inventory..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   3150
      TabIndex        =   4
      Top             =   90
      Width           =   3000
   End
   Begin VB.Label lblImport 
      Caption         =   "Importing Addresses..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   90
      TabIndex        =   3
      Top             =   750
      Width           =   3000
   End
   Begin VB.Label lblImport 
      Caption         =   "Importing Product Lines..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   90
      TabIndex        =   2
      Top             =   1080
      Width           =   3000
   End
   Begin VB.Label lblImport 
      Caption         =   "Exporting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   420
      Width           =   3000
   End
   Begin VB.Label lblImport 
      Caption         =   "Preparing To Export..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3000
   End
End
Attribute VB_Name = "MasImportExportStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LBL_EXP_PREP = 0
Private Const LBL_EXP_GO = 1
Private Const LBL_ADDRESSES = 2
Private Const LBL_PRODLINES = 3
Private Const LBL_VENDORS = 4
Private Const LBL_WAREHOUSES = 5
Private Const LBL_INVENTORY = 6
Private Const LBL_INV_QTYS = 7
Private Const LBL_INV_PRICES = 8
Private Const LBL_INV_UDFS = 9
Private Const LBL_SALES_HIST = 10

Private WithEvents MAS200IE As Mas200ImportExport
Attribute MAS200IE.VB_VarHelpID = -1

Public Sub onlyExport()
    Dim ctl As Label
    For Each ctl In Me.lblImport
        ctl.Enabled = False
    Next ctl
    Me.lblImport(LBL_EXP_PREP).Enabled = True
    Me.lblImport(LBL_EXP_GO).Enabled = True
End Sub

Public Sub doImport()
    Set MAS200IE = New Mas200ImportExport
    MAS200IE.StartImport
    Set MAS200IE = Nothing
End Sub

Public Sub doExport()
    Me.onlyExport
    Set MAS200IE = New Mas200ImportExport
    MAS200IE.StartExport False
    Set MAS200IE = Nothing
End Sub

Public Sub doRetryExport()
    Me.onlyExport
    Set MAS200IE = New Mas200ImportExport
    MAS200IE.StartExport False, True
    Set MAS200IE = Nothing
End Sub

Private Sub MAS200IE_StatusChanged(status As String)
    Me.lblStatusBar.Caption = status
    DoEvents
End Sub

Private Sub MAS200IE_StepChanged(step As String)
    Dim lbl As Label
    For Each lbl In Me.lblImport
        lbl.FontBold = False
    Next lbl
    Select Case step
        Case Is = "ExportPrep"
            Me.lblImport(LBL_EXP_PREP).FontBold = True
        Case Is = "Export"
            Me.lblImport(LBL_EXP_GO).FontBold = True
        Case Is = "Addresses"
            Me.lblImport(LBL_ADDRESSES).FontBold = True
        Case Is = "PL"
            Me.lblImport(LBL_PRODLINES).FontBold = True
        Case Is = "Vendors"
            Me.lblImport(LBL_VENDORS).FontBold = True
        Case Is = "Whse"
            Me.lblImport(LBL_WAREHOUSES).FontBold = True
        Case Is = "Inv"
            Me.lblImport(LBL_INVENTORY).FontBold = True
        Case Is = "InvQ"
            Me.lblImport(LBL_INV_QTYS).FontBold = True
        Case Is = "InvP"
            Me.lblImport(LBL_INV_PRICES).FontBold = True
        Case Is = "InvUDF"
            Me.lblImport(LBL_INV_UDFS).FontBold = True
        Case Is = "SalesHist"
            Me.lblImport(LBL_SALES_HIST).FontBold = True
    End Select
    DoEvents
End Sub
