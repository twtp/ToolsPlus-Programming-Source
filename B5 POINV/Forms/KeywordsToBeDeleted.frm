VERSION 5.00
Begin VB.Form KeywordsToBeDeleted 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keywords - To Be Deleted"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4545
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnUndo 
      Caption         =   "Undo"
      Height          =   285
      Left            =   3480
      TabIndex        =   7
      Top             =   2160
      Width           =   675
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2190
      TabIndex        =   6
      Top             =   4980
      Width           =   1455
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   450
      TabIndex        =   5
      Top             =   4980
      Width           =   1455
   End
   Begin VB.ListBox KwdList 
      Height          =   2400
      Left            =   210
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2490
      Width           =   3945
   End
   Begin VB.Label Label4 
      Caption         =   "Remove an item from the list by double clicking. When you hit ok, changes are saved."
      Height          =   465
      Left            =   120
      TabIndex        =   3
      Top             =   1890
      Width           =   4245
   End
   Begin VB.Label Label3 
      Caption         =   $"KeywordsToBeDeleted.frx":0000
      Height          =   825
      Left            =   120
      TabIndex        =   2
      Top             =   1020
      Width           =   4245
   End
   Begin VB.Label Label2 
      Caption         =   "The following keywords need to be manually deleted at Overture, they are already marked deleted in the database."
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Overture - Keyword Deletion"
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
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   4125
   End
End
Attribute VB_Name = "KeywordsToBeDeleted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private deleted As PerlArray

Private Sub btnCancel_Click()
    Set deleted = Nothing
    Unload KeywordsToBeDeleted
End Sub

Private Sub btnOK_Click()
    Dim phrNum As Long
    Dim phr As Variant
    'For Each phr In deleted.DataArray
    For phrNum = 0 To deleted.Scalar - 1
        phr = deleted.Elem(phrNum)
        DB.Execute "UPDATE KeywordsAttributes SET DoNotUse=1 WHERE SearchPhrase='" & EscapeSQuotes(CStr(phr)) & "' AND Service='Overture'"
    'Next phr
    Next phrNum
    Set deleted = Nothing
    Unload KeywordsToBeDeleted
End Sub

Private Sub btnUndo_Click()
    Dim foo As String
    foo = Nz(deleted.Pop)
    If foo <> "" Then
        Me.KwdList.AddItem foo
    End If
End Sub

Private Sub Form_Load()
    Set deleted = New PerlArray
    requeryKwdList
End Sub

Private Sub requeryKwdList()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT SearchPhrase FROM vKeywords WHERE Service='Overture' AND DoNotUse=2")
    Me.KwdList.Clear
    While Not rst.EOF
        Me.KwdList.AddItem rst("SearchPhrase")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub KwdList_DblClick()
    deleted.Push Me.KwdList
    Me.KwdList.RemoveItem Me.KwdList.ListIndex
End Sub
