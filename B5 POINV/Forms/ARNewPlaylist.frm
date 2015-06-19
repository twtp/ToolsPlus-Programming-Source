VERSION 5.00
Begin VB.Form ARNewPlaylist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audio Runner Playlist Maker"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Or Select An Existing Playlist"
      Height          =   2055
      Left            =   390
      TabIndex        =   7
      Top             =   2280
      Width           =   3765
      Begin VB.ListBox Playlists 
         Height          =   1620
         Left            =   270
         TabIndex        =   8
         Top             =   300
         Width           =   3255
      End
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Quit"
      Height          =   435
      Left            =   2610
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   435
      Left            =   1050
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Enter A Playlist Name"
      Height          =   765
      Left            =   390
      TabIndex        =   3
      Top             =   1350
      Width           =   3765
      Begin VB.TextBox PlaylistName 
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   270
         Width           =   3345
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Playlist Type"
      Height          =   945
      Left            =   390
      TabIndex        =   0
      Top             =   300
      Width           =   3795
      Begin VB.OptionButton PlaylistType 
         Caption         =   "Announcements"
         Height          =   465
         Index           =   1
         Left            =   1860
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   1485
      End
      Begin VB.OptionButton PlaylistType 
         Caption         =   "Music"
         Height          =   465
         Index           =   0
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   300
         Value           =   -1  'True
         Width           =   1485
      End
   End
End
Attribute VB_Name = "ARNewPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    requeryPlaylists
End Sub

Private Sub OK_Click()
    If Me.PlaylistType(MUSIC_PLS) = False And Me.PlaylistType(ANNOUNCEMENT_PLS) = False Then
        MsgBox "You didn't select a playlist type!"
        Exit Sub
    End If
    
    If Me.PlaylistName = "" Or Me.Playlists.SelCount = 0 Then
        MsgBox "Enter a name for the playlist, or select from the list"
        Exit Sub
    End If
    
    If Me.PlaylistName <> "" Then
        If nameTaken(PLAYLIST_DIR & getSubDir(Me.PlaylistType) & Me.PlaylistName & FILE_EXT) Then
            'just open for editing
        Else
            Open PLAYLIST_DIR & getSubDir(Me.PlaylistType) & Me.PlaylistName & FILE_EXT For Output As #1
            Print #1, "path" & vbTab & "context" & vbTab & "criteria_type" & vbTab & "criteria"
            Close #1
        End If
    Else
        'this is really just a backup, if the _Click doesn't fire to fill the name?
        Me.PlaylistName = Me.Playlists
    End If
    Load ARPlaylistEdit
    ARPlaylistEdit.PlaylistName = Me.PlaylistName
    ARPlaylistEdit.PlaylistType = getSubDir(Me.PlaylistType)
    ARPlaylistEdit.fname = PLAYLIST_DIR & getSubDir(Me.PlaylistType) & Me.PlaylistName & FILE_EXT
    ARPlaylistEdit.Show
    Unload Me
End Sub

Private Sub PlaylistName_KeyPress(KeyAscii As Integer)
    If Not isOkFileNameChar(KeyAscii) Then
        KeyAscii = 0
    End If
End Sub

Private Function isOkFileNameChar(KeyAscii As Integer) As Boolean
    If KeyAscii = Asc("/") Or _
       KeyAscii = Asc("\") Or _
       KeyAscii = Asc("?") Or _
       KeyAscii = Asc("*") Or _
       KeyAscii = Asc(":") Or _
       KeyAscii = Asc("""") Or _
       KeyAscii = Asc("<") Or _
       KeyAscii = Asc(">") Or _
       KeyAscii = Asc("|") Then
        isOkFileNameChar = False
    Else
        isOkFileNameChar = True
    End If
End Function

Private Function nameTaken(fname As String) As Boolean
    Dim foo As String
    foo = Dir$(fname)
    If foo = "" Then
        nameTaken = False
    Else
        nameTaken = True
    End If
End Function

Private Sub requeryPlaylists()
    Me.Playlists.Clear
    Dim Path As String
    Path = Dir$(PLAYLIST_DIR & getSubDir(Me.PlaylistType) & "*" & FILE_EXT)
    While Path <> ""
        Path = Replace(Path, FILE_EXT, "")
        Me.Playlists.AddItem Path
        Path = Dir$()
    Wend
End Sub

Private Sub Playlists_Click()
    Me.PlaylistName = Me.Playlists
End Sub

Private Sub Playlists_DblClick()
    Me.PlaylistName = Me.Playlists
    OK_Click
End Sub

Private Sub PlaylistType_Click(Index As Integer)
    requeryPlaylists
End Sub
