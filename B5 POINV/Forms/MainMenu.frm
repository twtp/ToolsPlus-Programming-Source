VERSION 5.00
Begin VB.Form MainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tools Plus - Main Menu"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   465
      Left            =   3780
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   3750
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton btnSigns 
      Caption         =   "Product Specs Inquiry"
      Height          =   435
      Left            =   1230
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton btnGoToImportExportMenu 
      Caption         =   "MAS 200 Import/Export"
      Height          =   435
      Left            =   1230
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton btnGoToUtilitiesMenu 
      Caption         =   "Utilities"
      Height          =   435
      Left            =   1230
      TabIndex        =   5
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   435
      Left            =   1230
      TabIndex        =   6
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton btnGoToKeywordsMenu 
      Caption         =   "Keywords"
      Height          =   435
      Left            =   1230
      TabIndex        =   4
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CommandButton btnInvMaint 
      Caption         =   "Inventory Inquiry"
      Height          =   435
      Left            =   1230
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblUserOverride 
      Alignment       =   1  'Right Justify
      Height          =   225
      Left            =   3420
      TabIndex        =   8
      Top             =   5430
      Width           =   1425
   End
   Begin VB.Label lblStatusBar 
      Height          =   225
      Left            =   30
      TabIndex        =   7
      Top             =   5430
      Width           =   3195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Main Menu"
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
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ImportTimer As Timer
Attribute ImportTimer.VB_VarHelpID = -1




Private Sub btnExit_Click()
    Unload MainMenu
End Sub

Private Sub btnGoToImportExportMenu_Click()
    Load MasImportExportMenu
    MasImportExportMenu.Show , MainMenu
    MainMenu.Hide
End Sub

Private Sub btnGoToKeywordsMenu_Click()
    Load KeywordsMenu
    KeywordsMenu.Show , MainMenu
    MainMenu.Hide
End Sub

Private Sub btnGoToUtilitiesMenu_Click()
    Load UtilitiesMenu
    UtilitiesMenu.Show , MainMenu
    MainMenu.Hide
End Sub

Private Sub btnInvMaint_Click()
    Mouse.Hourglass True
    If IsFormLoaded("InventoryMaintenance") Then
        If IsFormMinimized("InventoryMaintenance") Then
            UnMinimizeForm "InventoryMaintenance"
        Else
            FocusOnForm "InventoryMaintenance"
        End If
    Else
        Load InventoryMaintenance
        InventoryMaintenance.Show
    End If
    Mouse.Hourglass False
End Sub

Private Sub btnSigns_Click()
    Mouse.Hourglass True
    If IsFormLoaded("SignMaintenance") Then
        If IsFormMinimized("SignMaintenance") Then
            UnMinimizeForm "SignMaintenance"
        Else
            FocusOnForm "SignMaintenance"
        End If
    Else
        Load SignMaintenance
        SignMaintenance.Show
    End If
    Mouse.Hourglass False
End Sub

Private Sub Command1_Click()
    Load CompanyInfo
    CompanyInfo.Show
End Sub

Private Sub Command2_Click()
    Load InventoryCountMgmt
    InventoryCountMgmt.Show
End Sub

Private Sub Form_Load()
    Me.Label1.ToolTipText = "Version: " & app.Major & "." & app.Minor & "." & app.Revision

    If Not HasPermissionsTo("InventoryMaintenanceRead") Then
        Me.btnInvMaint.Enabled = False
    End If
    If Not HasPermissionsTo("SignMaintenanceRead") Then
        Me.btnSigns.Enabled = False
    End If
    If Not HasPermissionsTo("MasImport") Then
        Me.btnGoToImportExportMenu.Enabled = False
    End If
    If Not HasPermissionsTo("Keywords") Then
        Me.btnGoToKeywordsMenu.Enabled = False
    End If
    
    
    Select Case LCase(Environ("UserName"))
        Case Is = "briandonorfio"
            Me.Command1.Visible = True
            Me.Command2.Visible = True
        Case Is = "davidlegere"
            Me.Command1.Visible = True
        Case Is = "eric"
            Me.Command2.Visible = True
        Case Is = "kirkmehlin"
            Me.Command2.Visible = True
        Case Is = "lorenmehlin"
            Me.Command2.Visible = True
    End Select
    
    
    setAutoImportLabel
    Me.lblUserOverride.caption = Environ("UserName")

    Set ImportTimer = Controls.Add("VB.Timer", "ImportTimer")
    If WARN_ON_IMPORT Then
        ImportTimer_Timer 'start off the event right now
        ImportTimer.Interval = 60000
        ImportTimer.Enabled = True
    Else
        ImportTimer.Enabled = False
    End If
    
    'ImportTimer.Enabled = False
    
'    If RunningInIDE Then
'        Me.Command1.Visible = True
'        Me.Command2.Visible = True
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    For Each frm In Forms
        If frm.name <> "MainMenu" Then
            If MsgBox(frm.caption & " is still open, close anyway?", vbYesNo) = vbNo Then
                Cancel = 1
                Exit Sub
            End If
        End If
    Next frm
    For Each frm In Forms
        If frm.name <> "MainMenu" Then
            Unload frm
            Set frm = Nothing
        End If
    Next frm
    
    If Not RunningInIDE Then
        LogOutPoInv POINV_SESSION_ID
    End If
    
    cleanup
    
    End
End Sub

Private Sub ImportTimer_Timer()
    'some long-running jobs may want to turn off the timer temporarily, so check that first
    If WARN_ON_IMPORT Then
        ImportTimer.Enabled = False
        ReDim flags(1) As Boolean
        flags = ImportExportProcessFlags
        If IMPORT_GOING And Not flags(0) Then
            'import finished
            IMPORT_GOING = False
            MsgBox "Import has finished!"
        ElseIf EXPORT_GOING And Not flags(1) Then
            'export finished
            EXPORT_GOING = False
            If IsFormLoaded("InventoryMaintenance") Then
                InventoryMaintenance.LockForExport False
            End If
            If IsFormLoaded("PurchaseOrdersForm") Then
                PurchaseOrdersForm.LockForExport False
            End If
        ElseIf Not IMPORT_GOING And flags(0) Then
            'import started
            IMPORT_GOING = True
            MsgBox "An import has been started! Any changes you make will not flag the item to export to MAS 200!" & vbCrLf & vbCrLf & "it'd be nice if we could lock everything read-only..."
        ElseIf Not EXPORT_GOING And flags(1) Then
            'export started
            EXPORT_GOING = True
            If IsFormLoaded("InventoryMaintenance") Then
                InventoryMaintenance.LockForExport True
            End If
            If IsFormLoaded("PurchaseOrdersForm") Then
                PurchaseOrdersForm.LockForExport True
            End If
        End If
        ImportTimer.Enabled = True
    End If
End Sub

Private Sub setAutoImportLabel()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT TimeStarted, TimeFinished, Success FROM EventLog WHERE Event='autoimport' AND TimeStarted>'" & Format(Now(), "Short Date") & "'")
    If rst.EOF Then
        Me.lblStatusBar.caption = "Today's Automatic Import: FAILURE!"
        Me.lblStatusBar.FontBold = True
        Me.lblStatusBar.ToolTipText = "Autoimport was not found in the Event Log! No import done this morning?"
    Else
        If rst("Success") Then
            Me.lblStatusBar.caption = "Today's Automatic Import: Success!"
            Me.lblStatusBar.ToolTipText = "Started: " & Format(rst("TimeStarted"), "Long Time") & "    Finished: " & Format(rst("TimeFinished"), "Long Time")
        Else
            Me.lblStatusBar.caption = "Today's Automatic Import: FAILURE!"
            Me.lblStatusBar.FontBold = True
            Me.lblStatusBar.ToolTipText = "Autoimport was not successful! Either crashed, or just not finished yet?"
        End If
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub lblUserOverride_DblClick()
    Dim passwd As String
    passwd = InputBox("Enter administrative password:")
    If passwd = "" Then
        'do nothing
    ElseIf passwd = "foobar" Then
        USING_SECURITY = False
        MsgBox "Lockout functions disabled, close and reopen all windows. I hope you know what you're doing."
    Else
        MsgBox "Invalid password, stop playing with this."
    End If
End Sub
