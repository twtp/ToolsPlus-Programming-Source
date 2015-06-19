VERSION 5.00
Begin VB.Form KeywordsEngineCostsImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keywords - Cost Import"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5880
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Filename 
      Height          =   285
      Left            =   1260
      TabIndex        =   13
      Top             =   720
      Width           =   3105
   End
   Begin VB.CommandButton btnBrowseForFile 
      Caption         =   "Browse..."
      Height          =   285
      Left            =   4410
      TabIndex        =   12
      Top             =   720
      Width           =   885
   End
   Begin VB.Frame Frame2 
      Caption         =   "General"
      Height          =   2595
      Left            =   2250
      TabIndex        =   4
      Top             =   1230
      Width           =   3405
      Begin VB.ComboBox cmbService 
         Height          =   315
         Left            =   210
         TabIndex        =   15
         Top             =   270
         Width           =   2445
      End
      Begin VB.CheckBox chkHasHeaderRow 
         Caption         =   "Has header row"
         Height          =   285
         Left            =   210
         TabIndex        =   8
         Top             =   750
         Width           =   1425
      End
      Begin VB.TextBox phraseColumn 
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Text            =   "A"
         Top             =   1140
         Width           =   645
      End
      Begin VB.TextBox costColumn 
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Text            =   "B"
         Top             =   1500
         Width           =   645
      End
      Begin VB.CommandButton btnImport 
         Caption         =   "Import"
         Height          =   405
         Left            =   660
         TabIndex        =   5
         Top             =   2100
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "phrase/itemnumber column"
         Height          =   255
         Left            =   900
         TabIndex        =   10
         Top             =   1170
         Width           =   2115
      End
      Begin VB.Label Label3 
         Caption         =   "cost column"
         Height          =   225
         Left            =   900
         TabIndex        =   9
         Top             =   1530
         Width           =   2085
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Specific:"
      Height          =   2595
      Left            =   90
      TabIndex        =   3
      Top             =   1230
      Width           =   1875
      Begin VB.CommandButton btnImportNextag 
         Caption         =   "NexTag"
         Height          =   285
         Left            =   90
         TabIndex        =   16
         Top             =   270
         Width           =   1635
      End
      Begin VB.CommandButton btnImportShopzillaBizrate 
         Caption         =   "ShopzillaBizrate"
         Height          =   285
         Left            =   90
         TabIndex        =   11
         Top             =   600
         Width           =   1635
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   4800
      TabIndex        =   2
      Top             =   60
      Width           =   1005
   End
   Begin VB.TextBox engineid 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Text            =   "engineid"
      Top             =   120
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "File:"
      Height          =   255
      Left            =   900
      TabIndex        =   14
      Top             =   750
      Width           =   405
   End
   Begin VB.Label lblEngineName 
      Caption         =   "ENGINE_NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   2625
   End
End
Attribute VB_Name = "KeywordsEngineCostsImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private currentService As String

Public Sub LoadEngine(engine As String)
    Me.engineid = DLookup("ID", "KeywordsEngines", "Engine='" & EscapeSQuotes(engine) & "'")
    Me.lblEngineName.caption = engine
End Sub



Private Sub btnBrowseForFile_Click()
    Dim dlg As FilePicker
    Set dlg = New FilePicker
    dlg.SetParent Me.hwnd
    dlg.AddFilter "Comma-Separated Values Files", "*.csv"
    Me.Filename = dlg.ShowDialogOpen
    Set dlg = Nothing
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnImport_Click()
    If Not filledOut Then
        Exit Sub
    End If
    
    Dim excelapp As Excel.Application
    Set excelapp = New Excel.Application
    excelapp.Workbooks.Open Me.Filename
    
    Dim i As Long
    i = IIf(Me.chkHasHeaderRow, 1, 2)
    
    Dim existsList As Dictionary
    Set existsList = New Dictionary
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT SearchPhrase, CurrentBid FROM KeywordsAttributes WHERE EngineID=" & Me.engineid)
    While Not rst.EOF
        existsList.Add CStr(rst("SearchPhrase")), CStr(rst("CurrentBid"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    Dim done As Boolean, phrase As String, cost As String
    While Not done
        phrase = excelapp.ActiveWorkbook.ActiveSheet.Range(Me.phraseColumn & i).value
        cost = excelapp.ActiveWorkbook.ActiveSheet.Range(Me.costColumn & i).value
        If phrase = "" Then
            done = True
        Else
            If existsList.exists(CStr(phrase)) Then
                If existsList.item(CStr(phrase)) = Format(cost, "0.00") Then
                    'don't overwrite an existing cost?
                Else
                    DB.Execute "UPDATE KeywordsAttributes SET CurrentBid=" & cost & " WHERE SearchPhrase='" & EscapeSQuotes(phrase) & "' AND EngineID=" & Me.engineid
                End If
            Else
                DB.Execute "INSERT INTO KeywordsAttributes ( SearchPhrase, EngineID, CurrentBid ) VALUES ( '" & EscapeSQuotes(phrase) & "', " & Me.engineid & ", " & cost & " )"
            End If
        End If
    Wend
    excelapp.Quit
    Set excelapp = Nothing
    Set existsList = Nothing
End Sub

Private Sub btnImportNextag_Click()
    If checkFile() Then
        ShellWait PERL & " s:\mastest\mas90-signs\keywords_scripts\import_cost_nextag.pl " & qq(Me.Filename)
    End If
End Sub

Private Sub btnImportShopzillaBizrate_Click()
    If checkFile() Then
        ShellWait PERL & " s:\mastest\mas90-signs\keywords_scripts\import_cost_shopzilla.pl " & qq(Me.Filename)
    End If
End Sub

Private Sub costColumn_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToLetters(KeyAscii)
End Sub

Private Sub Form_Load()
    requeryServices
End Sub

Private Sub phraseColumn_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToLetters(KeyAscii)
End Sub

Private Function filledOut() As Boolean
    filledOut = False
    If Me.Filename = "" Then
        MsgBox "need filename!"
    ElseIf Me.phraseColumn = "" Then
        MsgBox "need phrase column #!"
    ElseIf Me.costColumn = "" Then
        MsgBox "need cost column #!"
    Else
        filledOut = True
    End If
End Function

Private Sub requeryServices()
    Me.cmbService.Clear
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, Engine FROM KeywordsEngines ORDER BY ID")
    While Not rst.EOF
        Me.cmbService.AddItem rst("Engine")
        rst.MoveNext
    Wend
    ExpandDropDownToFit Me.cmbService
    Me.cmbService = Me.cmbService.list(0) 'this breaks if we don't have any services...
    currentService = Me.cmbService.list(0)
End Sub

Private Function checkFile() As Boolean
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    checkFile = False
    If Me.Filename = "" Then
        MsgBox "You didn't enter a filename!"
    ElseIf Not fso.FileExists(Me.Filename) Then
        MsgBox "Unable to find file " & qq(Me.Filename)
    Else
        checkFile = True
    End If
    Set fso = Nothing
End Function
