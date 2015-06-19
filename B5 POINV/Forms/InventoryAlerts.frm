VERSION 5.00
Begin VB.Form InventoryAlerts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventory - Alerts"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnDateAlertNonTriggered 
      Caption         =   "Date Non-Triggered Alerts (Filter)"
      Enabled         =   0   'False
      Height          =   465
      Left            =   450
      TabIndex        =   8
      Top             =   2850
      Width           =   2775
   End
   Begin VB.CommandButton btnQtyAlertNonTriggered 
      Caption         =   "Quantity Non-Triggered Alerts (Filter)"
      Enabled         =   0   'False
      Height          =   465
      Left            =   450
      TabIndex        =   7
      Top             =   2310
      Width           =   2775
   End
   Begin VB.CommandButton btnQtyDateAlertFilter 
      Caption         =   "Quantity and Date Alert (Filter)"
      Enabled         =   0   'False
      Height          =   465
      Left            =   450
      TabIndex        =   6
      Top             =   690
      Width           =   2775
   End
   Begin VB.CommandButton btnDateAlertFilter 
      Caption         =   "Date Alert (Filter)"
      Enabled         =   0   'False
      Height          =   465
      Left            =   450
      TabIndex        =   5
      Top             =   1770
      Width           =   2775
   End
   Begin VB.CommandButton btnCustomerInStockRequests 
      Caption         =   "In-Stock Requests > 60 days old"
      Enabled         =   0   'False
      Height          =   465
      Left            =   450
      TabIndex        =   4
      Top             =   3930
      Width           =   2775
   End
   Begin VB.CommandButton btnDiscontinuePopup 
      Caption         =   "Discontinue Items (popup)"
      Enabled         =   0   'False
      Height          =   465
      Left            =   450
      TabIndex        =   3
      Top             =   3390
      Width           =   2775
   End
   Begin VB.CommandButton btnQtyAlertFilter 
      Caption         =   "Quantity Alert (Filter)"
      Enabled         =   0   'False
      Height          =   465
      Left            =   450
      TabIndex        =   2
      Top             =   1230
      Width           =   2775
   End
   Begin VB.CommandButton btnPriceDifferent 
      Caption         =   "Price Different (Filter)"
      Enabled         =   0   'False
      Height          =   465
      Left            =   450
      TabIndex        =   1
      Top             =   150
      Width           =   2775
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   285
      Left            =   1275
      TabIndex        =   0
      Top             =   4650
      Width           =   1125
   End
End
Attribute VB_Name = "InventoryAlerts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCustomerInStockRequests_Click()
    Load InventoryStockAlertRequests
    Unload Me 'need to unload to show ISAR non-modally
    InventoryStockAlertRequests.Show
End Sub



Private Sub btnDiscontinuePopup_Click()
    CheckForBeingDCItems
    btnExit_Click
End Sub

Private Sub btnExit_Click()
    Unload InventoryAlerts
End Sub

Private Sub btnPriceDifferent_Click()
    InventoryMaintenance.SetFilter "Price Different"
    btnExit_Click
End Sub

Private Sub btnQtyDateAlertFilter_Click()
    InventoryMaintenance.SetFilter "Alerts - All"
    btnExit_Click
End Sub

Private Sub btnQtyAlertFilter_Click()
    InventoryMaintenance.SetFilter "Alerts - Qty"
    btnExit_Click
End Sub

Private Sub btnDateAlertFilter_Click()
    InventoryMaintenance.SetFilter "Alerts - Date"
    btnExit_Click
End Sub

Private Sub btnQtyAlertNonTriggered_Click()
    InventoryMaintenance.SetFilter "Alerts - Qty Non-Triggered"
    btnExit_Click
End Sub

Private Sub btnDateAlertNonTriggered_Click()
    InventoryMaintenance.SetFilter "Alerts - Date Non-Triggered"
    btnExit_Click
End Sub



Private Sub Form_Load()
    Me.btnQtyDateAlertFilter.Enabled = CBool(CountInventoryQuantityTriggers(1, "") <> 0)
    If Me.btnQtyDateAlertFilter.Enabled Then
        Me.btnQtyAlertFilter.Enabled = CBool(CountInventoryQuantityTriggers(1, "qty") <> 0)
        Me.btnDateAlertFilter.Enabled = CBool(CountInventoryQuantityTriggers(1, "date") <> 0)
        Me.btnQtyAlertNonTriggered.Enabled = CBool(CountInventoryQuantityTriggers(0, "qty") <> 0)
        Me.btnDateAlertNonTriggered.Enabled = CBool(CountInventoryQuantityTriggers(0, "date") <> 0)
    End If
    Me.btnDiscontinuePopup.Enabled = CBool(CountBeingDCItems() <> 0)
    Me.btnPriceDifferent.Enabled = CBool(CountPriceDifferent() <> 0)
    Me.btnCustomerInStockRequests.Enabled = CBool(CountCustomerInStockRequestsAged() <> 0)
End Sub

