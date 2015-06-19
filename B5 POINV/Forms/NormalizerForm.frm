VERSION 5.00
Begin VB.Form NormalizerForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Normalizer"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   525
      Left            =   6960
      TabIndex        =   5
      Top             =   4290
      Width           =   1785
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1620
      TabIndex        =   3
      Text            =   "180"
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "big GO button"
      Height          =   585
      Left            =   3060
      TabIndex        =   2
      Top             =   4080
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   3180
      ItemData        =   "NormalizerForm.frx":0000
      Left            =   270
      List            =   "NormalizerForm.frx":0002
      OLEDropMode     =   1  'Manual
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   690
      Width           =   8505
   End
   Begin VB.Label Label2 
      Caption         =   "normalize to %:"
      Height          =   225
      Left            =   390
      TabIndex        =   4
      Top             =   4110
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "this will OVERWRITE EXISTING FILES!!! so make sure you're not working off the live copy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   930
      TabIndex        =   1
      Top             =   90
      Width           =   5685
   End
End
Attribute VB_Name = "NormalizerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If MsgBox("Normalize these files?", vbYesNo = vbYes) Then
        While Me.List1.ListCount <> 0
            ShellWait "h:\audio runner\programs\normalize.exe -m " & Me.Text1 & " """ & Me.List1.List(0) & """", vbNormalFocus
            Me.List1.RemoveItem 0
        Wend
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    For i = 1 To Data.Files.Count
        If fso.FolderExists(Data.Files(i)) Then
            addFolderRecursive fso.GetFolder(Data.Files(i))
        Else
            addFile fso.GetFile(Data.Files(i))
        End If
    Next i
End Sub

Private Sub addFolderRecursive(fname As Folder)
    Dim thisfolder As Folder, thisfile As File
    For Each thisfolder In fname.SubFolders
        addFolderRecursive thisfolder
    Next thisfolder
    For Each thisfile In fname.Files
        addFile thisfile
        DoEvents
    Next thisfile
End Sub

Private Sub addFile(fname As File)
    If LCase(Right(fname.Name, 4)) = ".wav" Then
        If Not InList(fname.Path) Then
            Me.List1.AddItem fname.Path
        End If
    End If
End Sub

Private Function InList(key As String) As Boolean
    Dim found As Boolean
    Dim low As Long, high As Long, try As Long
    low = 0
    high = Me.List1.ListCount - 1
    While low <= high And found = False
        try = (low + high) / 2
        If UCase$(key) = UCase$(Me.List1.List(try)) Then
            found = True
        ElseIf UCase$(key) > UCase$(Me.List1.List(try)) Then
            low = try + 1
        Else
            high = try - 1
        End If
    Wend
    InList = found
End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = vbKeyBack Then
        'ok
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Text1_LostFocus()
    If Me.Text1 = "" Then
        Me.Text1 = "180"
    End If
End Sub
