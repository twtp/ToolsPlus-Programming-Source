VERSION 5.00
Begin VB.Form Mas200AutoImportStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mas200 Automatic Import/Export"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblImport 
      Caption         =   "Updating Sales Rankings..."
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
      Index           =   12
      Left            =   3150
      TabIndex        =   13
      Top             =   1650
      Width           =   3000
   End
   Begin VB.Label lblImport 
      Caption         =   "Updating Freight..."
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
      Index           =   11
      Left            =   3150
      TabIndex        =   12
      Top             =   1320
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
      TabIndex        =   11
      Top             =   0
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
      TabIndex        =   10
      Top             =   330
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
      TabIndex        =   9
      Top             =   990
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
      TabIndex        =   8
      Top             =   660
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
      Left            =   90
      TabIndex        =   7
      Top             =   1980
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
      TabIndex        =   6
      Top             =   0
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
      TabIndex        =   5
      Top             =   330
      Width           =   3000
   End
   Begin VB.Label lblImport 
      Caption         =   "Enabling Triggers..."
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
      TabIndex        =   4
      Top             =   660
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
      TabIndex        =   3
      Top             =   1320
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
      TabIndex        =   2
      Top             =   1650
      Width           =   3000
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
      TabIndex        =   1
      Top             =   990
      Width           =   3000
   End
   Begin VB.Label lblStatusBar 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   6015
   End
End
Attribute VB_Name = "Mas200AutoImportStatus"
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
Private Const LBL_INV_DONE = 9
Private Const LBL_SALES_HIST = 10
Private Const LBL_FREIGHT = 11
Private Const LBL_SALES_RANK = 12

Private WithEvents import As Mas200ImportExport
Attribute import.VB_VarHelpID = -1

Public Sub doImport(Optional updateFreight As Boolean = True, _
                    Optional updateSalesRank As Boolean = True, _
                    Optional threshold As Dictionary = Nothing, _
                    Optional triadCodeTime As String = "", _
                    Optional triadCodeMetric As String = "", _
                    Optional forceOPCalc As Boolean = False)
    Set import = New Mas200ImportExport
    
    import.StartAutomaticImport updateFreight, updateSalesRank, threshold, triadCodeTime, triadCodeMetric, forceOPCalc
    Set import = Nothing
End Sub

Private Sub import_StatusChanged(status As String)
    Me.lblStatusBar.caption = status
    DoEvents
End Sub

Private Sub import_StepChanged(step As String)
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
        Case Is = "InvDone"
            Me.lblImport(LBL_INV_DONE).FontBold = True
        Case Is = "SalesHist"
            Me.lblImport(LBL_SALES_HIST).FontBold = True
        Case Is = "FreightActual"
            Me.lblImport(LBL_FREIGHT).FontBold = True
        Case Is = "SalesRank"
            Me.lblImport(LBL_SALES_RANK).FontBold = True
        Case Else
            'unknown step..probably should add it.
    End Select
    DoEvents
End Sub
