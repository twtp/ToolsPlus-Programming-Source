VERSION 5.00
Begin VB.Form UtilitiesMenuMore 
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
   Begin VB.CommandButton btnWarehouseLocations 
      Caption         =   "Whse / Location Maintenance"
      Height          =   435
      Left            =   1695
      TabIndex        =   10
      Top             =   1800
      Width           =   1485
   End
   Begin VB.CommandButton btnBDMAPPImport 
      Caption         =   "B&&D MAPP Import"
      Height          =   435
      Left            =   60
      TabIndex        =   9
      Top             =   3000
      Width           =   1485
   End
   Begin VB.CommandButton btnGeneralImport 
      Caption         =   "General XSheet Import"
      Height          =   435
      Left            =   60
      TabIndex        =   8
      Top             =   2400
      Width           =   1485
   End
   Begin VB.CommandButton btnProfitMargin 
      Caption         =   "Profit Margin Analysis"
      Enabled         =   0   'False
      Height          =   435
      Left            =   60
      TabIndex        =   7
      Top             =   1800
      Width           =   1485
   End
   Begin VB.CommandButton btnRecalcOP 
      Caption         =   "Recalculate Order Points"
      Height          =   435
      Left            =   3330
      TabIndex        =   6
      Top             =   1200
      Width           =   1485
   End
   Begin VB.CommandButton btnShowSaleCalcs 
      Caption         =   "Show/Sale Calculations"
      Height          =   435
      Left            =   60
      TabIndex        =   5
      Top             =   1200
      Width           =   1485
   End
   Begin VB.CommandButton btnRunPriceComp 
      Caption         =   "Run Price Comparison"
      Height          =   435
      Left            =   3330
      TabIndex        =   4
      Top             =   2430
      Width           =   1485
   End
   Begin VB.CommandButton btnRunFreightEst 
      Caption         =   "Run Freight Estimation"
      Height          =   435
      Left            =   3330
      TabIndex        =   3
      Top             =   1830
      Width           =   1485
   End
   Begin VB.CommandButton btnUserMaint 
      Caption         =   "User Maintenance"
      Height          =   435
      Left            =   1695
      TabIndex        =   2
      Top             =   1200
      Width           =   1485
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Back To Previous Utilities"
      Height          =   435
      Left            =   1230
      TabIndex        =   1
      Top             =   4800
      Width           =   2415
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
End
Attribute VB_Name = "UtilitiesMenuMore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBDMAPPImport_Click()
    Load MappImportBD
    MappImportBD.Show
End Sub

Private Sub btnExit_Click()
    Unload UtilitiesMenuMore
End Sub

Private Sub btnGeneralImport_Click()
    Load GeneralSpreadsheetImport
    GeneralSpreadsheetImport.Show
End Sub

Private Sub btnRecalcOP_Click()
    Load ReorderPointsForm
    ReorderPointsForm.Show
End Sub

Private Sub btnRunFreightEst_Click()
    Load FreightCalculation
    FreightCalculation.Show
End Sub

Private Sub btnRunPriceComp_Click()
    Load WebPriceComparison
    WebPriceComparison.Show
End Sub

Private Sub btnShowSaleCalcs_Click()
    Load ShowSaleCalculations
    ShowSaleCalculations.Show
End Sub

Private Sub btnUserMaint_Click()
    Load UserPermissions
    UserPermissions.Show
End Sub

Private Sub btnWarehouseLocations_Click()
    Load WarehouseLocationMaintenance
    WarehouseLocationMaintenance.Show
End Sub

Private Sub Form_Load()
    If Not HasPermissionsTo("PriceComparison") Then
        Me.btnRunPriceComp.Enabled = False
    End If
    If Not HasPermissionsTo("FreightEstimation") Then
        Me.btnRunFreightEst.Enabled = False
    End If
    If Not HasPermissionsTo("UserMaintenance") Then
        Me.btnUserMaint.Enabled = False
    End If
    If Not HasPermissionsTo("InventoryMaintenanceWrite") Then
        Me.btnRecalcOP.Enabled = False
        Me.btnBDMAPPImport.Enabled = False
    End If
    If Not HasPermissionsTo("ProfitMargin") Then
        Me.btnProfitMargin.Enabled = False
    End If
    If Not HasPermissionsTo("NewProducts") Then
        Me.btnGeneralImport.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UtilitiesMenu.Show
End Sub
