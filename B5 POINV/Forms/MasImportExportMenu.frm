VERSION 5.00
Begin VB.Form MasImportExportMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu - MAS 200 Import/Export"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   4875
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnUnlock 
      Caption         =   "Unlock"
      Height          =   225
      Left            =   1830
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton btnImport 
      Caption         =   "Import Data From MAS 200"
      Height          =   435
      Left            =   1230
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton btnRetryExport 
      Caption         =   "Retry Last Export To MAS 200"
      Height          =   435
      Left            =   1230
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Back To Main Menu"
      Height          =   435
      Left            =   1230
      TabIndex        =   4
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton btnExport 
      Caption         =   "Export Changes To MAS 200"
      Height          =   435
      Left            =   1230
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblUserDate 
      Alignment       =   2  'Center
      Caption         =   "Started by USERNAME at DATETIME"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3630
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label lblCurrentlyRunning 
      Alignment       =   2  'Center
      Caption         =   "An IMP/EXP is currently running."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   690
      TabIndex        =   6
      Top             =   3330
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MAS 200 Import/Export"
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
      TabIndex        =   5
      Top             =   5430
      Width           =   4815
   End
End
Attribute VB_Name = "MasImportExportMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload MasImportExportMenu
End Sub

Private Sub btnExport_Click()
    If MsgBox("Are you sure you want to EXPORT changes to MAS 200?", vbYesNo) = vbYes Then
        EXPORT_GOING = True                 'we don't need to warn outselves about the export, just set it manually
        Load MasImportExportStatus          'mutex and logs are handled in-class, so we just need to set these bools
        MasImportExportStatus.Show
        MasImportExportMenu.Hide
        MasImportExportStatus.doExport
        MasImportExportMenu.Show
        Unload MasImportExportStatus
        EXPORT_GOING = False                'clear the warning flag
    End If
End Sub

Private Sub btnImport_Click()
    If MsgBox("Are you sure you want to EXPORT and then IMPORT new data?", vbYesNo) = vbYes Then
        MsgBox "THE IMPORT IS BROKEN." & vbCrLf & vbCrLf & "until it gets fixed, run an autoimport manually:" & vbCrLf & qq("s:\mastest\mas90-signs\a_dist\automaticimport.exe"): Exit Sub
        IMPORT_GOING = True
        Load MasImportExportStatus
        MasImportExportStatus.Show
        MasImportExportMenu.Hide
        MasImportExportStatus.doImport
        MasImportExportMenu.Show
        Unload MasImportExportStatus
        IMPORT_GOING = False
    End If
End Sub

Private Sub btnRetryExport_Click()
    If MsgBox("Are you sure you want to RETRY the previous EXPORT?", vbYesNo) = vbYes Then
        EXPORT_GOING = True
        Load MasImportExportStatus
        MasImportExportStatus.Show
        MasImportExportMenu.Hide
        MasImportExportStatus.doRetryExport
        MasImportExportMenu.Show
        Unload MasImportExportStatus
        EXPORT_GOING = False
    End If
End Sub

Private Sub btnUnlock_Click()
    If MsgBox("If you're sure that the import isn't running (crashed on last run, etc.), click ok, otherwise check with the last person to run it.", vbOKCancel) = vbOK Then
        IMPORT_GOING = False
        EXPORT_GOING = False
        ClearProcessFlag "import"
        ClearProcessFlag "export"
        If InStr(Me.lblCurrentlyRunning.caption, "import") Then
            DB.Execute "EXEC spImportEnableTriggers"
        End If
        Me.btnImport.Enabled = True
        Me.btnExport.Enabled = True
        Me.btnRetryExport.Enabled = True
        Me.lblCurrentlyRunning.Visible = False
        Me.lblUserDate.Visible = False
        Me.btnUnlock.Visible = False
    End If
End Sub

Private Sub Form_Load()
    'checks the process flags, not the global booleans, just in case they're out of sync
    ReDim pflags(1) As Boolean
    pflags = ImportExportProcessFlags
    If pflags(0) Or pflags(1) Then
        Me.btnImport.Enabled = False
        Me.btnExport.Enabled = False
        Me.btnRetryExport.Enabled = False
        Me.lblCurrentlyRunning.Visible = True
        Me.lblUserDate.Visible = True
        Me.btnUnlock.Visible = True
        Me.lblUserDate.caption = Replace(Me.lblUserDate.caption, "USERNAME", DLookup("UserName", "ProcessFlags", "Task='" & IIf(pflags(0), "import", "export") & "'"))
        Me.lblUserDate.caption = Replace(Me.lblUserDate.caption, "DATETIME", DLookup("DateStamp", "ProcessFlags", "Task='" & IIf(pflags(0), "import", "export") & "'"))
        Me.lblCurrentlyRunning.caption = Replace(Me.lblCurrentlyRunning.caption, "IMP/EXP", IIf(pflags(0), "import", "export"))
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MainMenu.Show
End Sub
