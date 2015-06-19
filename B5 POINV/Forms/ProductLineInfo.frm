VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#2.2#0"; "TPControls.ocx"
Begin VB.Form ProductLineInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnReloadLogos 
      Caption         =   "Reload Logos"
      Height          =   345
      Left            =   7920
      TabIndex        =   16
      Top             =   4260
      Width           =   1545
   End
   Begin VB.TextBox ProductLineAndName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3330
      TabIndex        =   15
      Top             =   90
      Width           =   3885
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   465
      Left            =   8070
      TabIndex        =   14
      Top             =   0
      Width           =   1425
   End
   Begin VB.TextBox ReconditionedLogoPath 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   225
      Left            =   6990
      TabIndex        =   13
      Top             =   2730
      Width           =   2385
   End
   Begin VB.TextBox ManufLogoPath 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   225
      Left            =   6990
      TabIndex        =   10
      Top             =   900
      Width           =   2385
   End
   Begin TPControls.Picture ManufLogo 
      Height          =   1215
      Left            =   7020
      TabIndex        =   9
      Top             =   1140
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2143
   End
   Begin VB.TextBox ProductLineID 
      Height          =   315
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "PL_ID"
      Top             =   360
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox ReconditionedURL 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   5940
      Width           =   6615
   End
   Begin VB.TextBox Writeup2HTML 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3390
      Width           =   6675
   End
   Begin VB.TextBox Writeup1HTML 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   870
      Width           =   6675
   End
   Begin TPControls.Picture ReconditionedLogo 
      Height          =   1215
      Left            =   7020
      TabIndex        =   12
      Top             =   2970
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2143
   End
   Begin VB.Label lblOffloadWarning 
      Caption         =   "** THIS LINE WILL NOT GET OFFLOADED **"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1470
      TabIndex        =   17
      ToolTipText     =   "This is either because no items are published, it's LC-SUB mapped to another line, or it's marked hidden"
      Top             =   510
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label6 
      Caption         =   "Reconditioned Logo:"
      Height          =   225
      Left            =   6930
      TabIndex        =   11
      Top             =   2490
      Width           =   1545
   End
   Begin VB.Label Label5 
      Caption         =   "Manufacturer Logo:"
      Height          =   225
      Left            =   6930
      TabIndex        =   8
      Top             =   660
      Width           =   1545
   End
   Begin VB.Label Label4 
      Caption         =   "Reconditioned Info URL:"
      Height          =   225
      Left            =   60
      TabIndex        =   7
      Top             =   5730
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Bottom Caption:"
      Height          =   225
      Left            =   60
      TabIndex        =   6
      Top             =   3180
      Width           =   2025
   End
   Begin VB.Label Label2 
      Caption         =   "Product Line Web Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      TabIndex        =   5
      Top             =   60
      Width           =   2805
   End
   Begin VB.Label Label1 
      Caption         =   "Write-up:"
      Height          =   225
      Left            =   60
      TabIndex        =   4
      Top             =   660
      Width           =   735
   End
End
Attribute VB_Name = "ProductLineInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private changed As Boolean

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnReloadLogos_Click()
    Me.ManufLogo.DisplayImage Me.ManufLogoPath
    Me.ReconditionedLogo.DisplayImage Me.ReconditionedLogoPath
End Sub

Private Sub Form_Load()
    Dim pl As String
    'this is the only link to SignMaintenance, if we want to detach, change this
    pl = Left(SignMaintenance.ItemNumber, 3)
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, ProductLine, ManufFullNameWeb, Writeup1HTML, Writeup2HTML, ReconditionedURL FROM ProductLine WHERE ProductLine='" & pl & "'")
    If rst.EOF Then
        MsgBox "wtf?"
    Else
        Me.ProductLineID = rst("ID")
        Me.ProductLineAndName = rst("ProductLine") & " - " & Nz(rst("ManufFullNameWeb"))
        Me.Writeup1HTML = Nz(rst("Writeup1HTML"))
        Me.Writeup2HTML = Nz(rst("Writeup2HTML"))
        Me.ReconditionedURL = Nz(rst("ReconditionedURL"))
        Me.ManufLogoPath = GenerateLogoPicPath(rst("ProductLine"))
        Me.ReconditionedLogoPath = GenerateReconLogoPicPath(rst("ProductLine"))
        btnReloadLogos_Click
        If 0 = DLookup("COUNT(*)", "vWebOffloadLines", "ID=" & rst("ID")) Then
            Me.lblOffloadWarning.Visible = True
        Else
            Me.lblOffloadWarning.Visible = False
        End If
    End If
End Sub

Private Sub Writeup1HTML_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            Writeup1HTML_LostFocus
        Case Is = vbKeyDelete
            Writeup1HTML_KeyPress KeyCode
    End Select
End Sub

Private Sub Writeup1HTML_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub Writeup1HTML_LostFocus()
    If changed = True Then
        DB.Execute "UPDATE ProductLine SET Writeup1HTML='" & EscapeSQuotes(Me.Writeup1HTML) & "' WHERE ID=" & Me.ProductLineID
        changed = False
    End If
End Sub

Private Sub Writeup2HTML_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            Writeup2HTML_LostFocus
        Case Is = vbKeyDelete
            Writeup2HTML_KeyPress KeyCode
    End Select
End Sub

Private Sub Writeup2HTML_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub Writeup2HTML_LostFocus()
    If changed = True Then
        DB.Execute "UPDATE ProductLine SET Writeup2HTML='" & EscapeSQuotes(Me.Writeup2HTML) & "' WHERE ID=" & Me.ProductLineID
        changed = False
    End If
End Sub

Private Sub ReconditionedURL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            ReconditionedURL_LostFocus
        Case Is = vbKeyDelete
            ReconditionedURL_KeyPress KeyCode
    End Select
End Sub

Private Sub ReconditionedURL_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub ReconditionedURL_LostFocus()
    If changed = True Then
        If Len(Me.ReconditionedURL) <= 256 Then
            DB.Execute "UPDATE ProductLine SET ReconditionedURL='" & EscapeSQuotes(Me.ReconditionedURL) & "' WHERE ID=" & Me.ProductLineID
            changed = False
        Else
            MsgBox "URL must be under 256 chars!"
            Me.ReconditionedURL.SetFocus
        End If
    End If
End Sub
