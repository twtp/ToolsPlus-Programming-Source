VERSION 5.00
Begin VB.Form KeywordsAssociation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keywords - Association Tool"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8535
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox HideDeleted 
      Caption         =   "Hide Deleted"
      Height          =   255
      Left            =   4620
      TabIndex        =   9
      Top             =   60
      Width           =   1725
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   7230
      TabIndex        =   8
      Top             =   30
      Width           =   1305
   End
   Begin VB.CommandButton btnAddThis 
      Caption         =   "Add This Keyword"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4500
      TabIndex        =   7
      Top             =   840
      Width           =   1785
   End
   Begin VB.CommandButton btnGo 
      Caption         =   "Search"
      Height          =   255
      Left            =   3510
      TabIndex        =   6
      Top             =   840
      Width           =   885
   End
   Begin VB.ListBox Assoc 
      Height          =   4545
      Index           =   2
      Left            =   5730
      TabIndex        =   5
      Top             =   1260
      Width           =   2685
   End
   Begin VB.ListBox Assoc 
      Height          =   4545
      Index           =   1
      Left            =   2940
      TabIndex        =   4
      Top             =   1260
      Width           =   2685
   End
   Begin VB.ListBox Assoc 
      Height          =   4545
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   1260
      Width           =   2685
   End
   Begin VB.TextBox SearchPhrase 
      Height          =   285
      Left            =   150
      TabIndex        =   1
      Top             =   810
      Width           =   3315
   End
   Begin VB.Label Label2 
      Caption         =   "Enter new or existing keyword:"
      Height          =   225
      Left            =   180
      TabIndex        =   2
      Top             =   570
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "Keywords - Association Tool"
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
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4125
   End
End
Attribute VB_Name = "KeywordsAssociation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lastPhrase As String

Private Sub Assoc_DblClick(Index As Integer)
    Me.SearchPhrase = Me.Assoc(Index)
    btnGo_Click
    Me.btnAddThis.Enabled = False
End Sub

Private Sub btnAddThis_Click()
    Load KeywordsAddKeyword
    KeywordsAddKeyword.source = "ListTool"
    KeywordsAddKeyword.SearchPhrase = Me.SearchPhrase
    KeywordsAddKeyword.Show MODAL
    If DLookup("SearchPhrase", "KeywordsPhrases", "SearchPhrase='" & EscapeSQuotes(Me.SearchPhrase) & "'") = "" Then
        Me.btnAddThis.Enabled = True
    Else
        Me.btnAddThis.Enabled = False
    End If
End Sub

Private Sub btnExit_Click()
    Unload KeywordsAssociation
End Sub

Private Sub btnGo_Click()
    If Me.SearchPhrase = "" Then
        Me.SearchPhrase = lastPhrase
    ElseIf Me.SearchPhrase <> lastPhrase Then
        Dim words As Variant
        words = Split(Me.SearchPhrase, " ")
        Dim rst As ADODB.Recordset
        Dim i As Long
        For i = Me.Assoc.LBound To Me.Assoc.UBound
            Me.Assoc(i).Clear
        Next i
        For i = Me.Assoc.LBound To IIf(UBound(words) > Me.Assoc.UBound, Me.Assoc.UBound, UBound(words))
            Set rst = db.retrieve("SELECT SearchPhrase FROM " & IIf(Me.HideDeleted, "vKeywords", "KeywordsPhrases") & " WHERE SearchPhrase LIKE '%" & EscapeSQuotes(CStr(words(i))) & "%'" & getWhereClause)
            While Not rst.EOF
                Me.Assoc(i).AddItem rst("SearchPhrase")
                rst.MoveNext
            Wend
            rst.Close
        Next i
        Set rst = Nothing
        lastPhrase = Me.SearchPhrase
    End If
End Sub

Private Sub SearchPhrase_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SearchPhrase_LostFocus
    End If
End Sub

Private Sub SearchPhrase_LostFocus()
    btnGo_Click
    If Me.SearchPhrase <> "" Then
        If DLookup("SearchPhrase", "KeywordsPhrases", "SearchPhrase='" & EscapeSQuotes(Me.SearchPhrase) & "'") = "" Then
            Me.btnAddThis.Enabled = True
        Else
            Me.btnAddThis.Enabled = False
        End If
    Else
        Me.btnAddThis.Enabled = False
    End If
End Sub

Private Function getWhereClause() As String
    Dim clause As String
    If Me.HideDeleted Then
        clause = " AND DoNotUse=0"
    End If
    'more here
    getWhereClause = clause
End Function
