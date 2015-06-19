VERSION 5.00
Begin VB.Form KeywordsSuggestDlg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keywords - Phrase Suggestions"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox searchPhrase 
      Height          =   285
      Left            =   2970
      TabIndex        =   34
      Text            =   "hidden search phrase"
      Top             =   4230
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   375
      Left            =   1710
      TabIndex        =   33
      Top             =   4320
      Width           =   1125
   End
   Begin VB.CommandButton btnAddThis 
      Caption         =   "Add This"
      Height          =   285
      Index           =   9
      Left            =   3270
      TabIndex        =   32
      Top             =   3810
      Width           =   975
   End
   Begin VB.TextBox NumSearches 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   9
      Left            =   2430
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   3810
      Width           =   735
   End
   Begin VB.TextBox Suggestion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   9
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   3810
      Width           =   2205
   End
   Begin VB.CommandButton btnAddThis 
      Caption         =   "Add This"
      Height          =   285
      Index           =   8
      Left            =   3270
      TabIndex        =   29
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox NumSearches 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   8
      Left            =   2430
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Suggestion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   8
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3480
      Width           =   2205
   End
   Begin VB.CommandButton btnAddThis 
      Caption         =   "Add This"
      Height          =   285
      Index           =   7
      Left            =   3270
      TabIndex        =   26
      Top             =   3150
      Width           =   975
   End
   Begin VB.TextBox NumSearches 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   7
      Left            =   2430
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3150
      Width           =   735
   End
   Begin VB.TextBox Suggestion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   7
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3150
      Width           =   2205
   End
   Begin VB.CommandButton btnAddThis 
      Caption         =   "Add This"
      Height          =   285
      Index           =   6
      Left            =   3270
      TabIndex        =   23
      Top             =   2820
      Width           =   975
   End
   Begin VB.TextBox NumSearches 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   6
      Left            =   2430
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2820
      Width           =   735
   End
   Begin VB.TextBox Suggestion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   6
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2820
      Width           =   2205
   End
   Begin VB.CommandButton btnAddThis 
      Caption         =   "Add This"
      Height          =   285
      Index           =   5
      Left            =   3270
      TabIndex        =   20
      Top             =   2490
      Width           =   975
   End
   Begin VB.TextBox NumSearches 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   5
      Left            =   2430
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2490
      Width           =   735
   End
   Begin VB.TextBox Suggestion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   5
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2490
      Width           =   2205
   End
   Begin VB.CommandButton btnAddThis 
      Caption         =   "Add This"
      Height          =   285
      Index           =   4
      Left            =   3270
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox NumSearches 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   4
      Left            =   2430
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Suggestion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   4
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2160
      Width           =   2205
   End
   Begin VB.CommandButton btnAddThis 
      Caption         =   "Add This"
      Height          =   285
      Index           =   3
      Left            =   3270
      TabIndex        =   14
      Top             =   1830
      Width           =   975
   End
   Begin VB.TextBox NumSearches 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   3
      Left            =   2430
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1830
      Width           =   735
   End
   Begin VB.TextBox Suggestion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   3
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1830
      Width           =   2205
   End
   Begin VB.CommandButton btnAddThis 
      Caption         =   "Add This"
      Height          =   285
      Index           =   2
      Left            =   3270
      TabIndex        =   11
      Top             =   1500
      Width           =   975
   End
   Begin VB.TextBox NumSearches 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   2
      Left            =   2430
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1500
      Width           =   735
   End
   Begin VB.TextBox Suggestion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   2
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1500
      Width           =   2205
   End
   Begin VB.CommandButton btnAddThis 
      Caption         =   "Add This"
      Height          =   285
      Index           =   1
      Left            =   3270
      TabIndex        =   8
      Top             =   1170
      Width           =   975
   End
   Begin VB.TextBox NumSearches 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   1
      Left            =   2430
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1170
      Width           =   735
   End
   Begin VB.TextBox Suggestion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   1
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1170
      Width           =   2205
   End
   Begin VB.VScrollBar subformScrollBar 
      Height          =   3405
      Left            =   4320
      TabIndex        =   3
      Top             =   750
      Width           =   255
   End
   Begin VB.CommandButton btnAddThis 
      Caption         =   "Add This"
      Height          =   285
      Index           =   0
      Left            =   3270
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox NumSearches 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   0
      Left            =   2430
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Suggestion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   0
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   2205
   End
   Begin VB.Label lblTitle 
      Caption         =   "Suggestions for PHRASE"
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
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   4635
   End
   Begin VB.Label generalLabel 
      Caption         =   "# last month"
      Height          =   255
      Index           =   1
      Left            =   2430
      TabIndex        =   5
      Top             =   480
      Width           =   915
   End
   Begin VB.Label generalLabel 
      Caption         =   "Suggestion"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   4530
      X2              =   0
      Y1              =   4140
      Y2              =   4140
   End
   Begin VB.Line Line1 
      X1              =   4530
      X2              =   0
      Y1              =   750
      Y2              =   750
   End
End
Attribute VB_Name = "KeywordsSuggestDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NUM_BOXES As Long = 10

Dim currentPos As Long
Dim suggList() As String

Private Sub btnAddThis_Click(Index As Integer)
    Load KeywordsAddKeyword
    KeywordsAddKeyword.source = "SuggestDlg"
    KeywordsAddKeyword.searchPhrase = Me.Suggestion(Index)
    KeywordsAddKeyword.Show MODAL
    doHighlighting Index
End Sub

Private Sub btnExit_Click()
    Unload KeywordsSuggestDlg
End Sub

Private Sub SearchPhrase_Change()
    Me.lblTitle.caption = Replace(Me.lblTitle.caption, "PHRASE", Me.searchPhrase)
    requeryForm
    fillLines
End Sub

Private Sub requeryForm()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = retrieve("SELECT Suggestion FROM KeywordsSuggestions WHERE SearchPhrase='" & EscapeSQuotes(Me.searchPhrase) & "' ORDER BY Suggestion")
    If rst.RecordCount > 0 Then
        ReDim suggList(rst.RecordCount - 1) As String
        Dim i As Long
        For i = 0 To rst.RecordCount - 1
            suggList(i) = rst("Suggestion")
            rst.MoveNext
        Next i
        Me.subformScrollBar.Min = 0
        Me.subformScrollBar.max = IIf(rst.RecordCount - NUM_BOXES < 0, 0, rst.RecordCount - NUM_BOXES)
    Else
        ReDim suggList(0) As String
        suggList(0) = "NORECORDS"
        Me.subformScrollBar.Min = 0
        Me.subformScrollBar.max = 0
    End If
    currentPos = 0
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub fillLines(Optional offset As Long = 0)
    Dim rst As ADODB.Recordset
    Dim i As Long
    For i = 0 To NUM_BOXES - 1
        If UBound(suggList) >= i Then
            If suggList(offset + i) <> "NORECORDS" Then
                Set rst = retrieve("SELECT * FROM KeywordsSuggestions WHERE SearchPhrase='" & EscapeSQuotes(Me.searchPhrase) & "' AND Suggestion='" & EscapeSQuotes(suggList(offset + i)) & "'")
                Me.Suggestion(i) = rst("Suggestion")
                Me.NumSearches(i) = rst("NumSearches")
                SetSubformControlsVisible i, True
                rst.Close
            Else
                While i < NUM_BOXES
                    SetSubformControlsVisible i, False
                    i = i + 1
                Wend
            End If
        Else
            While i < NUM_BOXES
                SetSubformControlsVisible i, False
                i = i + 1
            Wend
        End If
    Next i
    doHighlighting
End Sub

Private Sub SetSubformControlsVisible(Index As Long, onoff As Boolean)
    Me.Suggestion(Index).Visible = onoff
    Me.NumSearches(Index).Visible = onoff
    Me.btnAddThis(Index).Visible = onoff
End Sub

Private Sub subformScrollBar_Change()
    Mouse.Hourglass True
    If suggList(0) <> "NORECORDS" Then
        If Abs(Me.subformScrollBar.value - currentPos) > 1 Then
            fillLines Me.subformScrollBar.value
        Else
            If Me.subformScrollBar.value > currentPos Then
                shuffleUp
                addLineAfter Me.Suggestion(NUM_BOXES - 2)
            Else
                shuffleDown
                addLineBefore Me.Suggestion(1)
            End If
        End If
        currentPos = Me.subformScrollBar.value
    End If
    Mouse.Hourglass False
End Sub

Private Sub shuffleUp()
    Dim i As Long
    For i = 0 To NUM_BOXES - 2
        shuffleLines i + 1, i
    Next i
End Sub

Private Sub shuffleDown()
    Dim i As Long
    For i = NUM_BOXES - 1 To 1 Step -1
        shuffleLines i - 1, i
    Next i
End Sub

Private Sub shuffleLines(fromIndex As Long, toIndex As Long)
    Me.Suggestion(toIndex) = Me.Suggestion(fromIndex)
    Me.Suggestion(toIndex).FontBold = Me.Suggestion(fromIndex).FontBold
    Me.NumSearches(toIndex) = Me.NumSearches(fromIndex)
    Me.NumSearches(toIndex).FontBold = Me.NumSearches(fromIndex).FontBold
    Me.btnAddThis(toIndex).Enabled = Me.btnAddThis(fromIndex).Enabled
End Sub

Private Sub addLineAfter(line As String)
    Dim i As Long
    i = bsearch(line, suggList) + 1
    Dim rst As ADODB.Recordset
    Set rst = retrieve("SELECT * FROM KeywordsSuggestions WHERE SearchPhrase='" & EscapeSQuotes(Me.searchPhrase) & "' AND Suggestion='" & EscapeSQuotes(suggList(i)) & "'")
    Me.Suggestion(NUM_BOXES - 1) = rst("Suggestion")
    Me.NumSearches(NUM_BOXES - 1) = rst("NumSearches")
    rst.Close
    doHighlighting NUM_BOXES - 1
    Set rst = Nothing
End Sub

Private Sub addLineBefore(line As String)
    Dim i As Long
    i = bsearch(line, suggList) - 1
    Dim rst As ADODB.Recordset
    Set rst = retrieve("SELECT * FROM KeywordsSuggestions WHERE SearchPhrase='" & EscapeSQuotes(Me.searchPhrase) & "' AND Suggestion='" & EscapeSQuotes(suggList(i)) & "'")
    Me.Suggestion(0) = rst("Suggestion")
    Me.NumSearches(0) = rst("NumSearches")
    rst.Close
    doHighlighting 0
    Set rst = Nothing
End Sub

Private Sub doHighlighting(Optional Index As Integer = -1)
    Dim max As Long, i As Long
    max = IIf(Index = -1, NUM_BOXES - 1, Index)
    i = IIf(Index = -1, 0, Index)
    For i = i To max
        If DLookup("SearchPhrase", "KeywordsPhrases", "SearchPhrase='" & EscapeSQuotes(Me.Suggestion(i)) & "'") <> "" Then
            Me.Suggestion(i).FontBold = True
            Me.NumSearches(i).FontBold = True
            Me.btnAddThis(i).Enabled = False
        Else
            Me.Suggestion(i).FontBold = False
            Me.NumSearches(i).FontBold = False
            Me.btnAddThis(i).Enabled = True
        End If
    Next i
End Sub
