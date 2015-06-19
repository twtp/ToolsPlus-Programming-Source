VERSION 5.00
Begin VB.Form ReorderPointsForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate Reorder Points"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   8115
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Query By:"
      Height          =   1515
      Left            =   60
      TabIndex        =   23
      Top             =   1500
      Width           =   5685
      Begin VB.CheckBox chkVendorAll 
         Height          =   225
         Left            =   5190
         TabIndex        =   33
         Top             =   930
         Width           =   225
      End
      Begin VB.CheckBox chkLineAll 
         Height          =   225
         Left            =   5190
         TabIndex        =   32
         Top             =   570
         Width           =   225
      End
      Begin VB.ComboBox cmbVendorEnd 
         Height          =   315
         Left            =   3090
         TabIndex        =   31
         Top             =   930
         Width           =   1875
      End
      Begin VB.ComboBox cmbVendorStart 
         Height          =   315
         Left            =   1110
         TabIndex        =   30
         Top             =   930
         Width           =   1875
      End
      Begin VB.ComboBox cmbLineEnd 
         Height          =   315
         Left            =   3090
         TabIndex        =   29
         Top             =   510
         Width           =   1875
      End
      Begin VB.ComboBox cmbLineStart 
         Height          =   315
         Left            =   1110
         TabIndex        =   28
         Top             =   510
         Width           =   1875
      End
      Begin VB.Label generalLabel 
         Caption         =   "All"
         Height          =   255
         Index           =   13
         Left            =   5190
         TabIndex        =   34
         Top             =   210
         Width           =   315
      End
      Begin VB.Label generalLabel 
         Caption         =   "End"
         Height          =   255
         Index           =   12
         Left            =   3690
         TabIndex        =   27
         Top             =   210
         Width           =   465
      End
      Begin VB.Label generalLabel 
         Caption         =   "Start"
         Height          =   255
         Index           =   11
         Left            =   1800
         TabIndex        =   26
         Top             =   210
         Width           =   525
      End
      Begin VB.Label generalLabel 
         Caption         =   "Vendor:"
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   25
         Top             =   960
         Width           =   705
      End
      Begin VB.Label generalLabel 
         Caption         =   "Line:"
         Height          =   255
         Index           =   9
         Left            =   390
         TabIndex        =   24
         Top             =   540
         Width           =   705
      End
   End
   Begin VB.TextBox LeadTime 
      Height          =   285
      Left            =   2130
      TabIndex        =   22
      Top             =   630
      Width           =   945
   End
   Begin VB.TextBox WeightLYR 
      Height          =   285
      Left            =   7080
      TabIndex        =   12
      Top             =   900
      Width           =   945
   End
   Begin VB.TextBox WeightYTD 
      Height          =   285
      Left            =   6090
      TabIndex        =   11
      Top             =   900
      Width           =   945
   End
   Begin VB.TextBox WeightLastPeriod 
      Height          =   285
      Left            =   5100
      TabIndex        =   10
      Top             =   900
      Width           =   945
   End
   Begin VB.TextBox WeightCurrentPeriod 
      Height          =   285
      Left            =   4110
      TabIndex        =   9
      Top             =   900
      Width           =   945
   End
   Begin VB.TextBox WeeksLYR 
      Height          =   285
      Left            =   7080
      TabIndex        =   8
      Top             =   540
      Width           =   945
   End
   Begin VB.TextBox WeeksYTD 
      Height          =   285
      Left            =   6090
      TabIndex        =   7
      Top             =   540
      Width           =   945
   End
   Begin VB.TextBox WeeksLastPeriod 
      Height          =   285
      Left            =   5100
      TabIndex        =   6
      Top             =   540
      Width           =   945
   End
   Begin VB.TextBox WeeksCurrentPeriod 
      Height          =   285
      Left            =   4110
      TabIndex        =   5
      Top             =   540
      Width           =   945
   End
   Begin VB.TextBox SafetySupply 
      Height          =   285
      Left            =   2130
      TabIndex        =   4
      Top             =   1020
      Width           =   945
   End
   Begin VB.TextBox ShelfSupply 
      Height          =   285
      Left            =   2130
      TabIndex        =   3
      Top             =   240
      Width           =   945
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   465
      Left            =   7050
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton btnRecalculate 
      Caption         =   "Recalculate"
      Height          =   525
      Left            =   6390
      TabIndex        =   1
      Top             =   1920
      Width           =   1665
   End
   Begin VB.CommandButton btnResetNewOPs 
      Caption         =   "Reset All New OPs"
      Height          =   345
      Left            =   6390
      TabIndex        =   0
      Top             =   1530
      Width           =   1665
   End
   Begin VB.Label lblStatusBar 
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   3120
      Width           =   5775
   End
   Begin VB.Label generalLabel 
      Caption         =   "Weight %:"
      Height          =   255
      Index           =   8
      Left            =   3330
      TabIndex        =   21
      Top             =   900
      Width           =   765
   End
   Begin VB.Label generalLabel 
      Caption         =   "Weeks:"
      Height          =   255
      Index           =   7
      Left            =   3330
      TabIndex        =   20
      Top             =   570
      Width           =   765
   End
   Begin VB.Label generalLabel 
      Caption         =   "Last Year"
      Height          =   255
      Index           =   6
      Left            =   7080
      TabIndex        =   19
      Top             =   180
      Width           =   945
   End
   Begin VB.Label generalLabel 
      Caption         =   "YTD"
      Height          =   255
      Index           =   5
      Left            =   6090
      TabIndex        =   18
      Top             =   180
      Width           =   945
   End
   Begin VB.Label generalLabel 
      Caption         =   "Last Period"
      Height          =   255
      Index           =   4
      Left            =   5100
      TabIndex        =   17
      Top             =   180
      Width           =   945
   End
   Begin VB.Label generalLabel 
      Caption         =   "Curr Period"
      Height          =   255
      Index           =   3
      Left            =   4110
      TabIndex        =   16
      Top             =   180
      Width           =   945
   End
   Begin VB.Label generalLabel 
      Caption         =   "Safety Stock Weeks:"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   15
      Top             =   1050
      Width           =   2055
   End
   Begin VB.Label generalLabel 
      Caption         =   "Lead Time Weeks:"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   14
      Top             =   660
      Width           =   2055
   End
   Begin VB.Label generalLabel 
      Caption         =   "Weeks Of Supply On Shelf:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   13
      Top             =   270
      Width           =   2055
   End
End
Attribute VB_Name = "ReorderPointsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEF_LINE_START = "AAA"
Private Const DEF_LINE_END = "XXA"
Private Const DEF_VEN_START = "0"
Private Const DEF_VEN_END = "zzzzzzz"

Private WithEvents MAS200IE As Mas200ImportExport
Attribute MAS200IE.VB_VarHelpID = -1
Private changed As Boolean
Private fillingform As Boolean

Private Sub btnExit_Click()
    Unload ReorderPointsForm
End Sub

Private Sub btnRecalculate_Click()
    Dim PLstart As String, PLend As String
    Dim Vstart As String, Vend As String
    If Me.chkLineAll Then
        PLstart = DEF_LINE_START
        PLend = DEF_LINE_END
    Else
        PLstart = Me.cmbLineStart
        PLend = Me.cmbLineEnd
    End If
    If Me.chkVendorAll Then
        Vstart = DEF_VEN_START
        Vend = DEF_VEN_END
    Else
        Vstart = Me.cmbVendorStart
        Vend = Me.cmbVendorEnd
    End If

    Set MAS200IE = New Mas200ImportExport
    MAS200IE.CalculateReorderPoints PLstart, PLend, Vstart, Vend
    Set MAS200IE = Nothing
End Sub

Private Sub MAS200IE_StatusChanged(status As String)
    Me.lblStatusBar.caption = status
End Sub

Private Sub btnResetNewOPs_Click()
    If MsgBox("Are you sure you want to delete recent calculations of new order points?", vbYesNo) = vbYes Then
        DB.Execute "UPDATE InventoryMaster SET NewReorderPoint=-1 WHERE ProductLine BETWEEN '" & DEF_LINE_START & "' AND '" & DEF_LINE_END & "'"
        DB.Execute "UPDATE InventoryMaster SET UpdateOrderPoint=0"
    End If
End Sub

Private Sub chkLineAll_Click()
    If Not fillingform Then
        If Me.chkLineAll Then
            Me.cmbLineStart = ""
            Disable Me.cmbLineStart
            Me.cmbLineEnd = ""
            Disable Me.cmbLineEnd
            DB.Execute "UPDATE ReorderPointConfig SET StartLine='" & DEF_LINE_START & "', EndLine='" & DEF_LINE_END & "'"
        Else
            Enable Me.cmbLineStart
            Me.cmbLineStart = DEF_LINE_START
            Enable Me.cmbLineEnd
            Me.cmbLineEnd = DEF_LINE_END
            DB.Execute "UPDATE ReorderPointConfig SET StartLine='" & DEF_LINE_START & "', EndLine='" & DEF_LINE_END & "'"
        End If
    End If
End Sub

Private Sub chkVendorAll_Click()
    If Not fillingform Then
        If Me.chkVendorAll Then
            Me.cmbVendorStart = ""
            Disable Me.cmbVendorStart
            Me.cmbVendorEnd = ""
            Disable Me.cmbVendorEnd
            DB.Execute "UPDATE ReorderPointConfig SET StartVen='" & DEF_VEN_START & "', EndVen='" & DEF_VEN_END & "'"
        Else
            Enable Me.cmbVendorStart
            Me.cmbVendorStart = DEF_VEN_START
            Enable Me.cmbVendorEnd
            Me.cmbVendorEnd = DEF_VEN_END
            DB.Execute "UPDATE ReorderPointConfig SET StartVen='" & DEF_VEN_START & "', EndVen='" & DEF_VEN_END & "'"
        End If
    End If
End Sub

Private Sub cmbVendorEnd_Click()
    cmbVendorEnd_LostFocus
End Sub

Private Sub cmbVendorEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.cmbVendorEnd, KeyCode, Shift
End Sub

Private Sub cmbVendorEnd_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.cmbVendorStart, KeyAscii
    If KeyAscii = vbKeyReturn Then
        cmbVendorEnd_LostFocus
    End If
End Sub

Private Sub cmbVendorEnd_LostFocus()
    DB.Execute "UPDATE ReorderPointConfig SET EndVen='" & Me.cmbVendorEnd & "'"
End Sub

Private Sub Form_Load()

    MsgBox "Don't use this, it's broken!!!!!"
    GoTo done

    DB.Execute "UPDATE ReorderPointConfig SET CurrWeek=" & DatePart("ww", Date)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM ReorderPointStoredWeeks WHERE WeekNo=" & DatePart("ww", Date))
    DB.Execute "spReorderPointConfigUpdate " & rst("WeeksCurPer") & ", " & _
                                               rst("WeeksLastPer") & ", " & _
                                               rst("WeeksYTD") & ", " & _
                                               rst("WeeksLastYear") & ", " & _
                                               rst("WeightCurPer") & ", " & _
                                               rst("WeightLastPer") & ", " & _
                                               rst("WeightYTD") & ", " & _
                                               rst("WeightLastYear")
    rst.Close
    Set rst = Nothing
    requeryLines
    requeryVendors
    fillForm
    
done:

End Sub





Private Sub requeryLines()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM  vRequeryLineCodes ORDER BY ProductLine")
    Me.cmbLineStart.Clear
    Me.cmbLineEnd.Clear
    While Not rst.EOF
        Me.cmbLineStart.AddItem (rst("ProductLine"))
        Me.cmbLineEnd.AddItem (rst("ProductLine"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryVendors()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM vRequeryVendor ORDER BY VendorName")
    Me.cmbVendorStart.Clear
    Me.cmbVendorEnd.Clear
    While Not rst.EOF
        Me.cmbVendorStart.AddItem (rst("VendorName"))
        Me.cmbVendorEnd.AddItem (rst("VendorName"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub fillForm()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM ReorderPointConfig WHERE ID=1")
    If rst.EOF Then
        MsgBox "Could not find default config."
        Unload ReorderPointsForm
        UtilitiesMenu.Show
    Else
        Me.ShelfSupply = rst("ShelfSupplyWks")
        Me.LeadTime = rst("LeadTimeWks")
        Me.SafetySupply = rst("SafetySupplyWks")
        Me.WeeksCurrentPeriod = rst("WeeksCurPer")
        Me.WeeksLastPeriod = rst("WeeksLastPer")
        Me.WeeksYTD = rst("WeeksYTD")
        Me.WeeksLYR = rst("WeeksLastYear")
        Me.WeightCurrentPeriod = rst("WeightCurPer")
        Me.WeightLastPeriod = rst("WeightLastPer")
        Me.WeightYTD = rst("WeightYTD")
        Me.WeightLYR = rst("WeightLastYear")
        fillingform = True
        If rst("StartLine") = DEF_LINE_START And rst("EndLine") = DEF_LINE_END Then
            Disable Me.cmbLineStart
            Disable Me.cmbLineEnd
            Me.chkLineAll = 1
        Else
            Me.cmbLineStart = rst("StartLine")
            Me.cmbLineEnd = rst("EndLine")
        End If
        If rst("StartVen") = DEF_VEN_START And rst("EndVen") = DEF_VEN_END Then
            Disable Me.cmbVendorStart
            Disable Me.cmbVendorEnd
            Me.chkVendorAll = 1
        Else
            Me.cmbVendorStart = rst("StartVen")
            Me.cmbVendorEnd = rst("EndVen")
        End If
        fillingform = False
    End If
    rst.Close
    Set rst = Nothing
End Sub





Private Sub cmbLineStart_Click()
    cmbLineStart_LostFocus
End Sub

Private Sub cmbLineStart_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.cmbLineStart, KeyCode, Shift
End Sub

Private Sub cmbLineStart_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.cmbLineStart, KeyAscii
    If KeyAscii = vbKeyReturn Then
        cmbLineStart_LostFocus
    End If
End Sub

Private Sub cmbLineStart_LostFocus()
    DB.Execute "UPDATE ReorderPointConfig SET StartLine='" & Me.cmbLineStart & "'"
End Sub

Private Sub cmbLineEnd_Click()
    cmbLineEnd_LostFocus
End Sub

Private Sub cmbLineEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.cmbLineEnd, KeyCode, Shift
End Sub

Private Sub cmbLineEnd_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.cmbLineEnd, KeyAscii
    If KeyAscii = vbKeyReturn Then
        cmbLineEnd_LostFocus
    End If
End Sub

Private Sub cmbLineEnd_LostFocus()
    DB.Execute "UPDATE ReorderPointConfig SET EndLine='" & Me.cmbLineEnd & "'"
End Sub

Private Sub cmbVendorStart_Click()
    cmbVendorStart_LostFocus
End Sub

Private Sub cmbVendorStart_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.cmbVendorStart, KeyCode, Shift
End Sub

Private Sub cmbVendorStart_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.cmbVendorStart, KeyAscii
    If KeyAscii = vbKeyReturn Then
        cmbVendorStart_LostFocus
    End If
End Sub

Private Sub cmbVendorStart_LostFocus()
    DB.Execute "UPDATE ReorderPointConfig SET StartVen='" & lookupPrimaryVendorNumber(Me.cmbVendorStart) & "'"
End Sub

Private Sub LeadTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        LeadTime_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub LeadTime_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.LeadTime) Then
            DB.Execute "UPDATE ReorderPointConfig SET LeadTimeWks=" & Me.LeadTime
            changed = False
        End If
    End If
End Sub

Private Sub SafetySupply_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SafetySupply_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub SafetySupply_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.SafetySupply) Then
            DB.Execute "UPDATE ReorderPointConfig SET SafetySupplyWks=" & Me.SafetySupply
            changed = False
        End If
    End If
End Sub

Private Sub ShelfSupply_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        ShelfSupply_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub ShelfSupply_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.ShelfSupply) Then
            DB.Execute "UPDATE ReorderPointConfig SET ShelfSupplyWks=" & Me.ShelfSupply
            changed = False
        End If
    End If
End Sub

Private Sub WeeksCurrentPeriod_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        WeeksCurrentPeriod_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub WeeksCurrentPeriod_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.WeeksCurrentPeriod) Then
            DB.Execute "UPDATE ReorderPointConfig SET WeeksCurPer=" & Me.WeeksCurrentPeriod
            changed = False
        End If
    End If
End Sub

Private Sub WeeksLastPeriod_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        WeeksLastPeriod_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub WeeksLastPeriod_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.WeeksLastPeriod) Then
            DB.Execute "UPDATE ReorderPointConfig SET WeeksLastPer=" & Me.WeeksLastPeriod
            changed = False
        End If
    End If
End Sub

Private Sub WeeksLYR_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        WeeksLYR_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub WeeksLYR_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.WeeksLYR) Then
            DB.Execute "UPDATE ReorderPointConfig SET WeeksLastYear=" & Me.WeeksLYR
            changed = False
        End If
    End If
End Sub

Private Sub WeeksYTD_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        WeeksYTD_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub WeeksYTD_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.WeeksYTD) Then
            DB.Execute "UPDATE ReorderPointConfig SET WeeksYTD=" & Me.WeeksYTD
            changed = False
        End If
    End If
End Sub

Private Sub WeightCurrentPeriod_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        WeightCurrentPeriod_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub WeightCurrentPeriod_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.WeightCurrentPeriod) Then
            DB.Execute "UPDATE ReorderPointConfig SET WeightCurPer=" & Me.WeightCurrentPeriod
            changed = False
        End If
    End If
End Sub

Private Sub WeightLastPeriod_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        WeightLastPeriod_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub WeightLastPeriod_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.WeeksLastPeriod) Then
            DB.Execute "UPDATE ReorderPointConfig SET WeightLastPer=" & Me.WeightLastPeriod
            changed = False
        End If
    End If
End Sub

Private Sub WeightLYR_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        WeightLYR_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub WeightLYR_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.WeightLYR) Then
            DB.Execute "UPDATE ReorderPointConfig SET WeightLastYear=" & Me.WeightLYR
            changed = False
        End If
    End If
End Sub

Private Sub WeightYTD_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        WeightYTD_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub WeightYTD_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.WeightYTD) Then
            DB.Execute "UPDATE ReorderPointConfig SET WeightYTD=" & Me.WeightYTD
            changed = False
        End If
    End If
End Sub
