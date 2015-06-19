VERSION 5.00
Begin VB.Form ConvertFrontEnd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Thumbnailer"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ConvertToThumbnail 
      Caption         =   "Convert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1410
      TabIndex        =   2
      Top             =   720
      Width           =   1755
   End
   Begin VB.TextBox itemNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   210
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "ItemNumber:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1275
   End
End
Attribute VB_Name = "ConvertFrontEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CONVERT_EXE = "s:\mastest\mas90-signs\utilities\convert.exe"
Private Const THUMB_WIDTH = "100>"
Private Const WEBIMG_PATH = "h:\final web graphics\"

Private Sub ConvertToThumbnail_Click()
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    Dim fname As String, dest As String
    fname = WEBIMG_PATH & Left(Me.itemNumber, 3) & "\" & Me.itemNumber & ".jpg"
    If fso.FileExists(fname) Then
        dest = Environ("UserProfile") & "\Desktop\" & UCase(Me.itemNumber) & ".jpg"
        ShellWait CONVERT_EXE & " " & """" & fname & """" & " -resize """ & THUMB_WIDTH & """ """ & dest & """"
        MsgBox "Thumbnail complete, saved as """ & dest & """"
    Else
        MsgBox "Unable to find file """ & fname & """!"
    End If
    
    Set fso = Nothing
End Sub
