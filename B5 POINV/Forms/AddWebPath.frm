VERSION 5.00
Begin VB.Form AddWebPath 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Web Paths"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancelChanges 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   7050
      TabIndex        =   14
      Top             =   5580
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton btnAcceptChanges 
      Caption         =   "Accept"
      Height          =   285
      Left            =   7050
      TabIndex        =   13
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox SearchableName 
      Height          =   285
      Left            =   1110
      TabIndex        =   10
      Top             =   5610
      Visible         =   0   'False
      Width           =   5865
   End
   Begin VB.TextBox PathName 
      Height          =   285
      Left            =   1110
      TabIndex        =   9
      Top             =   5280
      Visible         =   0   'False
      Width           =   5865
   End
   Begin VB.CommandButton btnDeleteSelPath 
      Caption         =   "Delete Path"
      Height          =   315
      Left            =   3900
      TabIndex        =   8
      Top             =   4830
      Width           =   1485
   End
   Begin VB.CommandButton btnAddNewPath 
      Caption         =   "Add New Path"
      Height          =   315
      Left            =   2190
      TabIndex        =   7
      Top             =   4830
      Width           =   1485
   End
   Begin VB.CommandButton btnEditPath 
      Caption         =   "Edit Path"
      Height          =   315
      Left            =   480
      TabIndex        =   6
      Top             =   4830
      Width           =   1485
   End
   Begin VB.TextBox currentPath 
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Text            =   "hidden current path"
      Top             =   750
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ListBox PathList 
      Height          =   3570
      Left            =   150
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1170
      Width           =   8655
   End
   Begin VB.ComboBox cmbPaths 
      Height          =   315
      Left            =   1890
      TabIndex        =   3
      Top             =   720
      Width           =   675
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   495
      Left            =   5790
      TabIndex        =   0
      Top             =   60
      Width           =   1425
   End
   Begin VB.Label kwdLabel 
      Caption         =   "Keywords:"
      Height          =   225
      Left            =   180
      TabIndex        =   12
      Top             =   5640
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label pathLabel 
      Caption         =   "Path Name:"
      Height          =   225
      Left            =   180
      TabIndex        =   11
      Top             =   5310
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label generalLabel 
      Caption         =   "Web Path Level:"
      Height          =   255
      Index           =   6
      Left            =   570
      TabIndex        =   2
      Top             =   780
      Width           =   1275
   End
   Begin VB.Label generalLabel 
      Caption         =   "Web Paths"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2970
      TabIndex        =   1
      Top             =   120
      Width           =   1425
   End
End
Attribute VB_Name = "AddWebPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private changed As Boolean
Private action As String

Private Sub btnAcceptChanges_Click()
    If Me.SearchableName = "" Then
        Me.SearchableName = Me.PathName
    End If
    If action = "add" Then
        If DLookup("ID", "WebPaths", "PathLevel=" & Me.currentPath & " AND WebPathName='" & EscapeSQuotes(Me.PathName) & "'") <> "" Then
            MsgBox "Path already exists!"
        Else
            DB.Execute "INSERT INTO WebPaths ( WebPathName, PathLevel, WebPathSearchableName ) VALUES ( '" & EscapeSQuotes(Me.PathName) & "', " & Me.currentPath & ", '" & EscapeSQuotes(IIf(Me.SearchableName = "", Me.PathName, Me.SearchableName)) & "' )"
            Me.PathList.AddItem Me.PathName & vbTab & IIf(Me.SearchableName = "" Or Me.PathName = Me.SearchableName, " ", "*") & vbTab & IIf(Me.SearchableName = "", Me.PathName, Me.SearchableName)
        End If
    ElseIf action = "edit" Then
        DB.Execute "UPDATE WebPaths SET WebPathName='" & EscapeSQuotes(Me.PathName) & "', WebPathSearchableName='" & Me.SearchableName & "' WHERE WebPathName='" & EscapeSQuotes(Left(Me.PathList, InStr(Me.PathList, vbTab) - 1)) & "' AND PathLevel=" & Me.currentPath
        Me.PathList.RemoveItem Me.PathList.ListIndex
        Me.PathList.AddItem Me.PathName & vbTab & IIf(Me.SearchableName = "" Or Me.PathName = Me.SearchableName, " ", "*") & vbTab & IIf(Me.SearchableName = "", Me.PathName, Me.SearchableName)
    Else
        MsgBox "unknown action, that sucks."
    End If
    btnCancelChanges_Click
    'requeryPathsList
End Sub

Private Sub btnAddNewPath_Click()
    Me.PathName.Visible = True
    Me.SearchableName.Visible = True
    Me.pathLabel.Visible = True
    Me.kwdLabel.Visible = True
    Me.btnAcceptChanges.Visible = True
    Me.btnCancelChanges.Visible = True
    Me.cmbPaths.Enabled = False
    Me.btnEditPath.Enabled = False
    Me.btnAddNewPath.Enabled = False
    Me.btnDeleteSelPath.Enabled = False
    Me.PathList.Enabled = False
    action = "add"
End Sub

Private Sub btnCancelChanges_Click()
    Me.PathName = ""
    Me.SearchableName = ""
    Me.PathName.Visible = False
    Me.SearchableName.Visible = False
    Me.pathLabel.Visible = False
    Me.kwdLabel.Visible = False
    Me.btnAcceptChanges.Visible = False
    Me.btnCancelChanges.Visible = False
    Me.btnEditPath.Enabled = CBool(Me.PathList.SelCount)
    Me.btnAddNewPath.Enabled = True
    Me.btnDeleteSelPath.Enabled = CBool(Me.PathList.SelCount)
    Me.PathList.Enabled = True
    Me.cmbPaths.Enabled = True
End Sub

Private Sub btnDeleteSelPath_Click()
    If MsgBox("Are you sure you want to delete this path (deletes associations w/ items too)?", vbYesNo) Then
        Dim pathid As String
        pathid = DLookup("ID", "WebPaths", "WebPathName='" & EscapeSQuotes(Left(Me.PathList, InStr(Me.PathList, vbTab) - 1)) & "' AND PathLevel=" & Me.currentPath)
        DB.Execute "DELETE FROM PartNumberWebPaths WHERE WebPathID=" & pathid
        DB.Execute "DELETE FROM WebPaths WHERE ID=" & pathid
        Me.PathList.RemoveItem Me.PathList.ListIndex
    End If
End Sub

Private Sub btnEditPath_Click()
    Me.PathName.Visible = True
    Me.PathName = Left(Me.PathList, InStr(Me.PathList, vbTab) - 1)
    Me.SearchableName.Visible = True
    Me.SearchableName = Mid(Me.PathList, InStrRev(Me.PathList, vbTab) + 1)
    Me.pathLabel.Visible = True
    Me.kwdLabel.Visible = True
    Me.btnAcceptChanges.Visible = True
    Me.btnCancelChanges.Visible = True
    Me.cmbPaths.Enabled = False
    Me.btnEditPath.Enabled = False
    Me.btnAddNewPath.Enabled = False
    Me.btnDeleteSelPath.Enabled = False
    Me.PathList.Enabled = False
    action = "edit"
End Sub

Private Sub btnExit_Click()
    Unload AddWebPath
End Sub

Private Sub cmbPaths_Click()
    changed = True
    cmbPaths_LostFocus
End Sub

Private Sub cmbPaths_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.cmbPaths, KeyCode, Shift
    Select Case KeyCode
        Case Is = vbKeyReturn
            cmbPaths_LostFocus
        Case Is = vbKeyDelete
            cmbPaths_KeyPress KeyCode
    End Select
End Sub

Private Sub cmbPaths_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub cmbPaths_LostFocus()
    If changed = True Then
        If Me.cmbPaths = "" Then
            Me.cmbPaths = Me.currentPath
        ElseIf Me.cmbPaths <> Me.currentPath Then
            Me.currentPath = Me.cmbPaths
            requeryPathsList
        End If
        changed = False
    End If
End Sub

Private Sub Form_Load()
    Dim tabs(1) As Long
    tabs(0) = 144
    tabs(1) = tabs(0) + 6
    SetListTabStops Me.PathList.hwnd, tabs
    requeryPathLevels
    Me.currentPath = "1"
    requeryPathsList
End Sub

Private Sub PathList_Click()
    Me.btnEditPath.Enabled = True
    Me.btnDeleteSelPath.Enabled = True
End Sub

Private Sub PathName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = Asc(":")
            KeyCode = 0     'colons not allowed
    End Select
End Sub

Private Sub PathName_LostFocus()
    If action = "add" And Me.SearchableName = "" Then
        Me.SearchableName = Me.PathName
    ElseIf action <> "add" And Me.PathName = "" Then
        Me.PathName = Left(Me.PathList, InStr(Me.PathList, vbTab) - 1)
    End If
End Sub

Private Sub requeryPathLevels()
    Me.cmbPaths.Clear
    Me.cmbPaths.AddItem "1"
    Me.cmbPaths.AddItem "2"
    Me.cmbPaths.AddItem "3"
    Me.cmbPaths.AddItem "4"
    Me.cmbPaths.AddItem "5"
End Sub

Private Sub requeryPathsList()
    Dim rst As ADODB.Recordset
    If Me.cmbPaths = "" Then
        Me.cmbPaths = Me.currentPath
    End If
    Set rst = DB.retrieve("SELECT WebPathName, WebPathSearchableName FROM WebPaths WHERE PathLevel=" & Me.cmbPaths & " ORDER BY WebPathName")
    Me.PathList.Clear
    While Not rst.EOF
        Me.PathList.AddItem rst("WebPathName") & vbTab & IIf(Nz(rst("WebPathSearchableName")) = "" Or Nz(rst("WebpathSearchableName")) = rst("WebPathName"), " ", "*") & vbTab & Nz(rst("WebPathSearchableName"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Me.btnEditPath.Enabled = False
    Me.btnDeleteSelPath.Enabled = False
End Sub

Private Sub SearchableName_LostFocus()
    If Me.SearchableName = "" Then
        Me.SearchableName = Me.PathName
    End If
End Sub
