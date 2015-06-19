VERSION 5.00
Begin VB.UserControl Picture 
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   915
   ScaleHeight     =   915
   ScaleWidth      =   915
   Begin VB.Image Image1 
      Height          =   915
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   915
   End
End
Attribute VB_Name = "Picture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Click()
Public Event DblClick()

Private cachePointer As Scripting.Dictionary

Public Sub DisplayImage(path As String)
On Error GoTo errh
    If path = "" Then
        UserControl.Image1 = LoadPicture("")
        UserControl.Image1.Tag = ""
    ElseIf Dir$(path) = "" Then
        Err.Raise 53
        GoTo errh
    Else
        Dim pic As IPictureDisp
        If cachePointer Is Nothing Then
            Set pic = LoadPicture(path)
            loadPicFrom pic, LCase(path)
        Else
            If Not cachePointer.Exists(LCase(path)) Then
                Set pic = LoadPicture(path)
                cachePointer.Add LCase(path), pic
            End If
            loadPicFrom cachePointer.item(LCase(path)), LCase(path)
        End If
    End If
    Exit Sub
errh:
    UserControl.Image1 = LoadPicture("")
    UserControl.Image1.Tag = ""
    Select Case Err.Number
        Case Is = 52
            'bad file name or number, probably a missing or invalid directory (CON, etc), ignore
        Case Is = 53
            'file not found, ignore
        Case Is = 481
            'invalid picture file, probably CMYK format?
            MsgBox "Picture in CMYK format!" & vbCrLf & vbCrLf & "Open it in Corel, click Image->Convert To, select RGB 24-bit color, and resave it."
        Case Else
            MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    End Select
    Err.Clear
End Sub

Public Sub loadPicFrom(pic As IPictureDisp, path As String)
    If pic.Height / UserControl.Height > pic.Width / UserControl.Width Then
        UserControl.Image1.Height = UserControl.Height
        UserControl.Image1.Width = pic.Width / (pic.Height / UserControl.Height)
    Else
        UserControl.Image1.Width = UserControl.Width
        UserControl.Image1.Height = pic.Height / (pic.Width / UserControl.Width)
    End If
    UserControl.Image1.Picture = pic
    UserControl.Image1.Tag = path
End Sub

Public Sub SetCachePointer(cache As Scripting.Dictionary)
    Set cachePointer = cache
End Sub

Public Function GetImagePath() As String
    GetImagePath = UserControl.Image1.Tag
End Function

Private Sub Image1_Click()
    RaiseEvent Click
End Sub

Private Sub Image1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Resize()
    UserControl.Image1.Width = UserControl.Width
    UserControl.Image1.Height = UserControl.Height
End Sub

