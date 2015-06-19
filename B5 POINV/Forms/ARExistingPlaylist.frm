VERSION 5.00
Begin VB.Form ARExistingPlaylist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audio Runner Playlist Maker"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Select Playlist"
      Height          =   2055
      Left            =   450
      TabIndex        =   5
      Top             =   1410
      Width           =   3795
      Begin VB.ListBox Playlists 
         Height          =   1620
         Left            =   270
         TabIndex        =   6
         Top             =   300
         Width           =   3255
      End
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2460
      TabIndex        =   4
      Top             =   3690
      Width           =   1215
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   435
      Left            =   1020
      TabIndex        =   3
      Top             =   3690
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Playlist Type"
      Height          =   945
      Left            =   450
      TabIndex        =   0
      Top             =   360
      Width           =   3795
      Begin VB.OptionButton PlaylistType 
         Caption         =   "Music"
         Height          =   465
         Index           =   0
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton PlaylistType 
         Caption         =   "Announcements"
         Height          =   465
         Index           =   1
         Left            =   1860
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   300
         Width           =   1485
      End
   End
End
Attribute VB_Name = "ARExistingPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
    Load ARSelectMenu
    ARSelectMenu.Show
    Unload Me
End Sub

Private Sub Form_Load()
    requeryPlaylists
End Sub

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

Private Sub OK_Click()
    If Me.Playlists.SelCount = 0 Then
        MsgBox "No playlist selected!"
        Exit Sub
    End If
    
    Load ARPlaylistEdit
    ARPlaylistEdit.PlaylistName = Me.Playlists
    ARPlaylistEdit.PlaylistType = getSubDir(Me.PlaylistType)
    ARPlaylistEdit.fname = PLAYLIST_DIR & getSubDir(Me.PlaylistType) & Me.Playlists & FILE_EXT
    ARPlaylistEdit.Show
    Unload Me
End Sub

Private Sub PlaylistType_Click(index As Integer)
    requeryPlaylists
End Sub
