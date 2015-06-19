VERSION 5.00
Begin VB.Form ARSelectMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audio Runner Playlist Maker"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Exit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   1013
      TabIndex        =   2
      Top             =   1740
      Width           =   2655
   End
   Begin VB.CommandButton EditPlaylist 
      Caption         =   "Edit Existing Playlist"
      Height          =   495
      Left            =   1013
      TabIndex        =   1
      Top             =   870
      Width           =   2655
   End
   Begin VB.CommandButton NewPlaylist 
      Caption         =   "Create New Playlist"
      Height          =   495
      Left            =   1013
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "ARSelectMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EditPlaylist_Click()
    Load ARExistingPlaylist
    ARExistingPlaylist.Show
    Unload Me
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub NewPlaylist_Click()
    Load ARNewPlaylist
    ARNewPlaylist.Show
    Unload Me
End Sub
