VERSION 5.00
Begin VB.Form SearchByPath 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search By Path"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8745
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnGoToSigns 
      Caption         =   "Specs Inquiry for item"
      Enabled         =   0   'False
      Height          =   435
      Left            =   5040
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
   Begin VB.CommandButton btnGoToPoInv 
      Caption         =   "Inventory Inquiry for item"
      Enabled         =   0   'False
      Height          =   435
      Left            =   5040
      TabIndex        =   1
      Top             =   750
      Width           =   3615
   End
   Begin VB.CommandButton btnFilterByItems 
      Caption         =   "Filter for just these items"
      Height          =   435
      Left            =   5010
      TabIndex        =   0
      Top             =   150
      Width           =   3645
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close Search Form"
      Height          =   1245
      Left            =   7380
      TabIndex        =   3
      Top             =   2580
      Width           =   1125
   End
   Begin PoInv.ItemFinder ItemFinder1 
      Height          =   4545
      Left            =   60
      TabIndex        =   4
      Top             =   30
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   8017
   End
End
Attribute VB_Name = "SearchByPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload SearchByPath
End Sub

Private Sub btnFilterByItems_Click()
    saveSelection
    Dim strsql As String, i As Long
    For i = LBound(FBS_PATHS) To UBound(FBS_PATHS)
        If FBS_PATHS(i) = "" Then
            i = UBound(FBS_PATHS)
        Else
            strsql = strsql & "InventoryMaster.ItemNumber IN (SELECT ItemNumber FROM vPNWebPaths WHERE WebPathName='" & EscapeSQuotes(FBS_PATHS(i)) & "' AND PathLevel=" & i & ") AND "
        End If
    Next i
    strsql = Left(strsql, Len(strsql) - 5)
    strsql = "SELECT DISTINCT InventoryMaster.ItemNumber FROM InventoryMaster INNER JOIN PartNumbers ON InventoryMaster.ItemNumber=PartNumbers.ItemNumber WHERE " & strsql & " ORDER BY InventoryMaster.ItemNumber"

    If IsFormLoaded("SignMaintenance") Then
        SignMaintenance.FilterBySearch strsql
    End If
    If IsFormLoaded("InventoryMaintenance") Then
        InventoryMaintenance.FilterBySearch strsql
    End If
    Unload SearchByPath
End Sub

Private Sub btnGoToPoInv_Click()
    If IsFormLoaded("InventoryMaintenance") Then
        InventoryMaintenance.GoToItem Me.ItemFinder1.GetSelectedItem, True
    Else
        Load InventoryMaintenance
        InventoryMaintenance.GoToItem Me.ItemFinder1.GetSelectedItem, True
        InventoryMaintenance.Show
    End If
    saveSelection
End Sub

Private Sub btnGoToSigns_Click()
    If IsFormLoaded("SignMaintenance") Then
        SignMaintenance.GoToItem Me.ItemFinder1.GetSelectedItem, True
    Else
        Load SignMaintenance
        SignMaintenance.GoToItem Me.ItemFinder1.GetSelectedItem, True
        SignMaintenance.Show
    End If
    saveSelection
End Sub

Private Sub Form_Load()
    loadSelection
    Me.ItemFinder1.Initialize
End Sub

Private Sub saveSelection()
    'why does vb have to suck so much?
    ReDim Paths(1 To 5) As String
    Paths = Me.ItemFinder1.GetSettings
    Dim i As Long
    For i = LBound(Paths) To UBound(Paths)
        FBS_PATHS(i) = Paths(i)
    Next i
End Sub

Private Sub loadSelection()
    Me.ItemFinder1.ApplySettings FBS_PATHS
End Sub

Private Sub ItemFinder1_HasSelection()
    Me.btnGoToPoInv.Enabled = True
    Me.btnGoToSigns.Enabled = True
End Sub

Private Sub ItemFinder1_LostSelection()
    Me.btnGoToPoInv.Enabled = False
    Me.btnGoToSigns.Enabled = False
End Sub
