VERSION 5.00
Begin VB.Form SignsIncludesAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Signs - Add Includes"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5130
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnExit 
      Caption         =   "Close w/o Adding"
      Height          =   405
      Left            =   2670
      TabIndex        =   4
      Top             =   2850
      Width           =   1845
   End
   Begin VB.CommandButton btnAddThese 
      Caption         =   "Add These"
      Height          =   435
      Left            =   660
      TabIndex        =   3
      Top             =   2820
      Width           =   1755
   End
   Begin VB.CommandButton btnFormat 
      Caption         =   "Format Nicely"
      Height          =   345
      Left            =   1230
      TabIndex        =   2
      Top             =   2340
      Width           =   2265
   End
   Begin VB.TextBox IncludesPaste 
      Height          =   1725
      Left            =   450
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   540
      Width           =   3585
   End
   Begin VB.Label Label1 
      Caption         =   "Paste Here:"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   270
      Width           =   1215
   End
End
Attribute VB_Name = "SignsIncludesAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim highestSortOrder As Long

Private Sub btnAddThese_Click()
    Dim Lines As Variant
    Lines = Split(Me.IncludesPaste, vbCrLf)
    Dim i As Long
    For i = 0 To UBound(Lines)
        highestSortOrder = highestSortOrder + 1
        If Lines(i) <> "" Then
            Lines(i) = UnFrontpage(CStr(Lines(i)))
            DB.Execute "INSERT INTO PartNumberIncludes ( ItemNumber, IncludeText, SortOrder ) VALUES ( '" & SignMaintenance.ItemNumber & "', '" & EscapeSQuotes(CStr(Lines(i))) & "', " & CStr(highestSortOrder) & " )", True
        End If
    Next i
    Unload SignsIncludesAdd
End Sub

Private Sub btnExit_Click()
    Unload SignsIncludesAdd
End Sub

Private Sub btnFormat_Click()
    Dim Lines As Variant
    Lines = Split(Me.IncludesPaste, vbCrLf)
    Dim i As Long
    If UBound(Lines) = 0 Then
        Lines = Split(Replace(Me.IncludesPaste, ", ", vbCrLf), vbCrLf)
    End If
    For i = 0 To UBound(Lines)
        Lines(i) = Trim(Lines(i))
        If Left(Lines(i), 1) = "*" Or Left(Lines(i), 1) = "-" Then
            Lines(i) = Mid(Lines(i), 2)
            Lines(i) = Trim(Lines(i))
        End If
        Lines(i) = UCase(Left(Lines(i), 1)) & Mid(Lines(i), 2)
    Next i
    Me.IncludesPaste = ""
    For i = 0 To UBound(Lines)
        If Lines(i) <> "" Then
            Me.IncludesPaste = Me.IncludesPaste & Lines(i) & vbCrLf
        End If
    Next i
    Me.IncludesPaste = Chomp(Me.IncludesPaste)
End Sub

Private Sub Form_Load()
    Dim temp As String
    temp = DLookup("MAX(SortOrder)", "PartNumberIncludes", "ItemNumber='" & SignMaintenance.ItemNumber & "'")
    If temp = "" Then
        highestSortOrder = -1
    Else
        highestSortOrder = temp
    End If
End Sub
