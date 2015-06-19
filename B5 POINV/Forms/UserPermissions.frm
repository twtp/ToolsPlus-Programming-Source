VERSION 5.00
Begin VB.Form UserPermissions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Maintenance"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   9570
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   6375
      Left            =   1710
      TabIndex        =   6
      Top             =   840
      Width           =   7815
      Begin VB.CheckBox chkPermission 
         Caption         =   "permissions go here"
         Height          =   225
         Index           =   0
         Left            =   2520
         TabIndex        =   13
         Top             =   450
         Width           =   2475
      End
      Begin VB.TextBox NTUsername 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   540
         Width           =   1995
      End
      Begin VB.TextBox ShortName 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1290
         Width           =   1995
      End
      Begin VB.ComboBox Location 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
         Width           =   1995
      End
      Begin VB.Label generalLabel 
         Caption         =   "Modules Allowed:"
         Height          =   225
         Index           =   2
         Left            =   2520
         TabIndex        =   14
         Top             =   210
         Width           =   1335
      End
      Begin VB.Label generalLabel 
         Caption         =   "NT Username:"
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   12
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label generalLabel 
         Caption         =   "Short Name:"
         Height          =   225
         Index           =   4
         Left            =   150
         TabIndex        =   11
         Top             =   1050
         Width           =   1335
      End
      Begin VB.Label generalLabel 
         Caption         =   "Location:"
         Height          =   225
         Index           =   5
         Left            =   150
         TabIndex        =   10
         Top             =   1800
         Width           =   1335
      End
   End
   Begin VB.CommandButton btnSaveChanges 
      Caption         =   "Save Changes"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   7260
      Width           =   1485
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   375
      Left            =   5550
      TabIndex        =   4
      Top             =   30
      Width           =   1515
   End
   Begin VB.CommandButton btnAddUser 
      Caption         =   "Add New User"
      Height          =   375
      Left            =   150
      TabIndex        =   3
      Top             =   7260
      Width           =   1485
   End
   Begin VB.ListBox UsersList 
      Height          =   6300
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   930
      Width           =   1545
   End
   Begin VB.Label generalLabel 
      Caption         =   "Select User:"
      Height          =   225
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   690
      Width           =   1335
   End
   Begin VB.Label generalLabel 
      Caption         =   "User Maintenance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   2505
   End
End
Attribute VB_Name = "UserPermissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tt As FormTooltips

Private locIDMap As Dictionary

Private changed As Boolean
Private stopevents As Boolean

Private Const CHECK_LEFT_OFFSET As Long = 2520
Private Const CHECK_TOP_OFFSET As Long = 450
Private Const CHECK_WIDTH As Long = 2475
Private Const CHECK_HEIGHT As Long = 225

Private Const CHECK_COLUMN_WIDTH As Long = 2535
Private Const CHECK_ROW_WIDTH As Long = 225

Private Const CHECKS_PER_COLUMN As Long = 26



Private Sub btnAddUser_Click()
    If changed Then
        If vbYes = MsgBox("You have unsaved changes, save now?", vbYesNo) Then
            If Not saveChanges Then
                Exit Sub
            End If
        End If
        changed = False
    End If
    
    Dim username As String
    username = InputBox("Enter username (must the exactly the same as NT authentication name):")
    If username <> "" Then
        username = LCase(username)
        username = Replace(username, "'", "")
        If DLookup("NTUsername", "Users", "NTUsername='" & username & "'") <> "" Then
            MsgBox "User already exists."
        Else
            Dim ShortName As String
            ShortName = InputBox("Enter this person's name / nickname:")
            If ShortName <> "" Then
                DB.Execute "INSERT INTO Users ( NTUsername, ShortName ) VALUES ( '" & EscapeSQuotes(username) & "', '" & EscapeSQuotes(ShortName) & "' )"
                requeryUsers
                Me.UsersList = username & vbTab & DLookup("@@IDENTITY", "Users")
                UsersList_Click
            End If
        End If
    End If
End Sub

Private Sub btnExit_Click()
    If changed Then
        If vbYes = MsgBox("You have unsaved changes, are you sure you want to exit?", vbYesNo) Then
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub btnSaveChanges_Click()
    If changed Then
        saveChanges
        If Me.NTUsername = Environ("UserName") Then
            SecurityPermissions.ReloadPermissionsHash
        End If
    End If
End Sub

Private Sub chkPermission_Click(Index As Integer)
    If Not stopevents Then
        changed = True
    End If
End Sub

Private Sub Form_Load()
    Set tt = New FormTooltips
    ReDim tabs(0) As Long
    tabs(0) = 300
    SetListTabStops Me.UsersList.hWnd, tabs
    requeryUsers
    requeryLocations
    setupPermissionsChecks
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set tt = Nothing
    Set locIDMap = Nothing
    SecurityPermissions.ReloadPermissionsHash
End Sub

Private Sub Location_Click()
    If Not stopevents Then
        changed = True
    End If
End Sub

Private Sub ShortName_KeyDown(KeyCode As Integer, Shift As Integer)
    changed = True
End Sub

Private Sub ShortName_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub UsersList_Click()
    If changed Then
        If vbYes = MsgBox("You have unsaved changes, save now?", vbYesNo) Then
            If Not saveChanges Then
                Exit Sub
            End If
        End If
        changed = False
    End If
    
    If Me.UsersList.SelCount = 0 Then
        Me.Frame1.Enabled = False
    Else
        Me.Frame1.Enabled = True
    End If
    
    stopevents = True
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT NTUsername, ShortName, DefaultWhse FROM Users WHERE ID=" & getPID())
    If rst.EOF Then
        Me.NTUsername = ""
        Me.ShortName = ""
        Me.Location.ListIndex = -1
    Else
        Me.NTUsername = rst("NTUsername")
        Me.ShortName = rst("ShortName")
        If Nz(rst("DefaultWhse")) = "" Then
            Me.Location = "<none>"
        Else
            Me.Location = DLookup("Name", "Warehouses", "ID=" & rst("DefaultWhse"))
        End If
    End If
    rst.Close
    
    Set rst = DB.retrieve("SELECT ModuleID FROM UserModulePermissions WHERE PersonID=" & getPID())
    Dim i As Long
    For i = 0 To Me.chkPermission.UBound
        Me.chkPermission(i) = 0
    Next i
    While Not rst.EOF
        For i = 0 To Me.chkPermission.UBound
            If Me.chkPermission(i).tag = rst("ModuleID") Then
                Me.chkPermission(i) = 1
                Exit For
            End If
        Next i
        rst.MoveNext
    Wend
    stopevents = False
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryUsers()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT NTUsername, ID FROM Users ORDER BY NTUsername")
    Me.UsersList.Clear
    While Not rst.EOF
        Me.UsersList.AddItem rst("NTUsername") & vbTab & rst("ID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub setupPermissionsChecks()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ModuleName, Explanation, ID FROM UserModules ORDER BY ModuleName")
    Dim xpos As Long, ypos As Long
    xpos = 0
    ypos = 0
    While Not rst.EOF
        Dim checknum As Long
        checknum = (xpos * CHECKS_PER_COLUMN) + ypos
        If checknum <> 0 Then
            Load Me.chkPermission(checknum)
        End If
        Me.chkPermission(checknum).Left = CHECK_LEFT_OFFSET + xpos * CHECK_COLUMN_WIDTH
        Me.chkPermission(checknum).Top = CHECK_TOP_OFFSET + ypos * CHECK_ROW_WIDTH
        Me.chkPermission(checknum).width = CHECK_WIDTH
        Me.chkPermission(checknum).Height = CHECK_HEIGHT
        Me.chkPermission(checknum).Caption = spaceModuleNameNicely(rst("ModuleName"))
        Me.chkPermission(checknum).tag = rst("ID")
        Me.chkPermission(checknum).Visible = True
        tt.AddToolTipToCtl Me.chkPermission(checknum), CStr(rst("Explanation"))
        ypos = ypos + 1
        If ypos = CHECKS_PER_COLUMN Then
            xpos = xpos + 1
            ypos = 0
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Function spaceModuleNameNicely(nonSpacedName As String) As String
    Dim modname As String
    Dim i As Long
    modname = Left(nonSpacedName, 1)
    For i = 2 To Len(nonSpacedName)
        If Mid(nonSpacedName, i, 1) = UCase(Mid(nonSpacedName, i, 1)) Then
            modname = modname & " "
        End If
        modname = modname & Mid(nonSpacedName, i, 1)
    Next i
    spaceModuleNameNicely = modname
End Function

Private Sub requeryLocations()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, Name FROM Warehouses ORDER BY Name")
    Me.Location.Clear
    Me.Location.AddItem "<none>"
    Set locIDMap = New Dictionary
    locIDMap.Add 0, "<none>"
    Dim i  As Long
    i = 1
    While Not rst.EOF
        Me.Location.AddItem rst("Name")
        locIDMap.Add i, CStr(rst("ID"))
        i = i + 1
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Function saveChanges() As Boolean
    If Me.ShortName = "" Then
        MsgBox "You must enter a short name!"
        saveChanges = False
        Exit Function
    ElseIf Me.Location.ListIndex = -1 Then
        Me.Location = "<none>"
    End If

    Dim pid As String
    pid = getPID(Me.NTUsername)
    
    DB.Execute "UPDATE Users SET ShortName='" & EscapeSQuotes(Me.ShortName) & "', DefaultWhse=" & IIf(Me.Location = "<none>", "NULL", locIDMap.item(Me.Location.ListIndex)) & " WHERE ID=" & pid
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ModuleID FROM UserModulePermissions WHERE PersonID=" & pid)
    Dim cache As Dictionary
    Set cache = New Dictionary
    While Not rst.EOF
        cache.Add CStr(rst("ModuleID").value), 1
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    Dim i As Long
    For i = 0 To Me.chkPermission.UBound
        If Me.chkPermission(i) = 1 Then
            If cache.exists(Me.chkPermission(i).tag) Then
                cache.Remove Me.chkPermission(i).tag
            Else
                DB.Execute "INSERT INTO UserModulePermissions ( PersonID, ModuleID ) VALUES ( " & pid & ", " & Me.chkPermission(i).tag & " )"
            End If
        End If
    Next i
    
    For i = 0 To UBound(cache.Keys)
        DB.Execute "DELETE FROM UserModulePermissions WHERE PersonID=" & pid & " AND ModuleID=" & cache.Keys(i)
    Next i
    
    Set cache = Nothing
    
    changed = False
    
    saveChanges = True
End Function

Private Function getPID(Optional username As String = "") As String
    If username = "" Then
        getPID = Mid(Me.UsersList, InStr(Me.UsersList, vbTab) + 1)
    Else
        getPID = DLookup("ID", "Users", "NTUsername='" & username & "'")
    End If
End Function
