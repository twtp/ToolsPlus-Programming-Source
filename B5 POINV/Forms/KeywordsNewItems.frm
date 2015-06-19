VERSION 5.00
Begin VB.Form KeywordsNewItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keywords - New Items"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   3480
      TabIndex        =   7
      Top             =   3030
      Width           =   1095
   End
   Begin VB.CommandButton btnAddSelected 
      Caption         =   "Add Selected"
      Height          =   435
      Left            =   2370
      TabIndex        =   6
      Top             =   3030
      Width           =   1095
   End
   Begin VB.CommandButton btnAddAll 
      Caption         =   "Add All"
      Height          =   435
      Left            =   1260
      TabIndex        =   5
      Top             =   3030
      Width           =   1095
   End
   Begin VB.CommandButton btnQuery 
      Caption         =   "Query"
      Height          =   435
      Left            =   150
      TabIndex        =   4
      Top             =   3030
      Width           =   1095
   End
   Begin VB.ListBox itemList 
      Height          =   2010
      Left            =   90
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   810
      Width           =   4575
   End
   Begin VB.Label lblStatusBar 
      AutoSize        =   -1  'True
      Height          =   225
      Left            =   60
      TabIndex        =   8
      Top             =   3480
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Search Phrase"
      Height          =   225
      Left            =   1650
      TabIndex        =   3
      Top             =   570
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "ItemNumber"
      Height          =   225
      Left            =   150
      TabIndex        =   2
      Top             =   570
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Keywords - New Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   4065
   End
End
Attribute VB_Name = "KeywordsNewItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAddAll_Click()
    Mouse.Hourglass True
    lockForm True
    ReDim todo(Me.itemList.ListCount - 1) As String
    todo = ListBoxAsArray(Me.itemList, False)
    doList todo
    btnQuery_Click
    lockForm False
    Mouse.Hourglass False
End Sub

Private Sub btnAddSelected_Click()
    Mouse.Hourglass True
    lockForm True
    ReDim todo(Me.itemList.SelCount - 1) As String
    todo = ListBoxAsArray(Me.itemList, True)
    doList todo
    btnQuery_Click
    lockForm False
    Mouse.Hourglass False
End Sub

Private Sub btnExit_Click()
    Unload KeywordsNewItems
End Sub

Private Sub btnQuery_Click()
    Dim rst As ADODB.Recordset
    'vItemsNeedKeywords handles trimming and filtering, so we're good with just this recordset
    Set rst = DB.retrieve("SELECT ItemNumber, FullPhrase FROM vItemsNeedKeywords ORDER BY ItemNumber")
    Me.itemList.Clear
    While Not rst.EOF
        DoEvents
        Me.itemList.AddItem rst("ItemNumber") & vbTab & rst("FullPhrase")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub Form_Load()
    Dim tabs(0) As Long
    tabs(0) = 72
    SetListTabStops Me.itemList.hwnd, tabs
End Sub

Private Sub doList(items() As String)
    Dim i As Long, item As String, phrase As String
    For i = 0 To UBound(items)
        item = Left(items(i), InStr(items(i), vbTab) - 1)
        'phrase = Mid(Items(i), InStr(Items(i), vbTab) + 1)
        Me.lblStatusBar.caption = "Working on " & item
        DoEvents
        AddItemKeyword item, True
    Next i
    Me.lblStatusBar.caption = ""
End Sub

Private Sub lockForm(yesno As Boolean)
    Me.btnQuery.Enabled = Not yesno
    Me.btnAddAll.Enabled = Not yesno
    Me.btnAddSelected.Enabled = Not yesno
    Me.btnExit.Enabled = Not yesno
End Sub

