VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form BlogManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Blog Articles"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   12450
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton opView 
      Caption         =   "Preview"
      Height          =   255
      Index           =   2
      Left            =   5310
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   960
      Width           =   1035
   End
   Begin VB.OptionButton opView 
      Caption         =   "HTML"
      Height          =   255
      Index           =   1
      Left            =   4260
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   960
      Width           =   1035
   End
   Begin VB.OptionButton opView 
      Caption         =   "Text"
      Height          =   255
      Index           =   0
      Left            =   3210
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   960
      Value           =   -1  'True
      Width           =   1035
   End
   Begin VB.TextBox newImageID 
      Height          =   255
      Left            =   10860
      TabIndex        =   24
      Text            =   "hidden new image"
      Top             =   660
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Frame PicFrame 
      Caption         =   "Article Images:"
      Height          =   2355
      Left            =   10350
      TabIndex        =   21
      Top             =   1020
      Width           =   2025
      Begin VB.ListBox ImageList 
         Height          =   1620
         Left            =   90
         TabIndex        =   23
         Top             =   600
         Width           =   1845
      End
      Begin VB.CommandButton btnAddImage 
         Caption         =   "Add New Picture"
         Height          =   285
         Left            =   150
         TabIndex        =   22
         Top             =   270
         Width           =   1695
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   11280
      TabIndex        =   20
      Top             =   0
      Width           =   1155
   End
   Begin VB.CommandButton btnPublish 
      Caption         =   "TODO: Publish"
      Enabled         =   0   'False
      Height          =   435
      Left            =   7410
      TabIndex        =   19
      Top             =   6360
      Width           =   1125
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save Draft"
      Enabled         =   0   'False
      Height          =   435
      Left            =   6150
      TabIndex        =   18
      Top             =   6360
      Width           =   1125
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   435
      Left            =   4920
      TabIndex        =   17
      Top             =   6360
      Width           =   1125
   End
   Begin VB.TextBox articleID 
      Height          =   255
      Left            =   3180
      TabIndex        =   16
      Text            =   "current article ID"
      Top             =   30
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame SelectFrame 
      Caption         =   "Select article to view/edit:"
      Height          =   3315
      Left            =   60
      TabIndex        =   7
      Top             =   630
      Width           =   2955
      Begin VB.CommandButton btnNewArticle 
         Caption         =   "New Article"
         Height          =   555
         Left            =   2070
         TabIndex        =   15
         Top             =   840
         Width           =   765
      End
      Begin VB.OptionButton opShow 
         Caption         =   "Show Drafts Only"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   14
         Top             =   1170
         Width           =   1935
      End
      Begin VB.OptionButton opShow 
         Caption         =   "Show Published Only"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   13
         Top             =   930
         Width           =   2025
      End
      Begin VB.OptionButton opShow 
         Caption         =   "Show All"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   690
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.ListBox titleSelect 
         Height          =   1620
         Left            =   120
         TabIndex        =   11
         Top             =   1620
         Width           =   2715
      End
      Begin VB.ComboBox authorName 
         Height          =   315
         Left            =   1260
         TabIndex        =   9
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label generalLabel 
         Caption         =   "Title:"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   10
         Top             =   1410
         Width           =   975
      End
      Begin VB.Label generalLabel 
         Caption         =   "Author Name:"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   330
         Width           =   975
      End
   End
   Begin VB.TextBox articleText 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   3150
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   7065
   End
   Begin VB.TextBox articleTitle 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3750
      TabIndex        =   4
      Top             =   540
      Width           =   4755
   End
   Begin VB.Frame ItemsFrame 
      Caption         =   "Items to associate:"
      Enabled         =   0   'False
      Height          =   2295
      Left            =   10350
      TabIndex        =   0
      Top             =   3540
      Width           =   2025
      Begin VB.ListBox ItemsAssociated 
         Height          =   1620
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   1785
      End
      Begin VB.ComboBox ItemPicker 
         Height          =   315
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   1785
      End
   End
   Begin SHDocVwCtl.WebBrowser articlePreview 
      Height          =   4905
      Left            =   3480
      TabIndex        =   29
      Top             =   1560
      Visible         =   0   'False
      Width           =   7125
      ExtentX         =   12568
      ExtentY         =   8652
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox articleHTML 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   3270
      MultiLine       =   -1  'True
      TabIndex        =   28
      Top             =   1320
      Visible         =   0   'False
      Width           =   7065
   End
   Begin VB.Label generalLabel 
      Caption         =   "Manage Blog Articles"
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
      Index           =   4
      Left            =   90
      TabIndex        =   6
      Top             =   60
      Width           =   2955
   End
   Begin VB.Label generalLabel 
      Caption         =   "Title:"
      Height          =   255
      Index           =   2
      Left            =   3180
      TabIndex        =   3
      Top             =   570
      Width           =   555
   End
End
Attribute VB_Name = "BlogManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SHOW_ALL = 0
Private Const SHOW_PUB = 1
Private Const SHOW_NEW = 2

Private Const VIEW_TEXT = 0
Private Const VIEW_HTML = 1
Private Const VIEW_PAGE = 2

Private Const ALL_AUTHORS = "[see all]"
Private Const NO_TITLE = "[untitled]"

Private currentAuthor As String



Private Sub btnAddImage_Click()
    If Me.articleID = -1 Then
        btnSave_Click
    End If
    Load BlogAddImage
    BlogAddImage.articleID = Me.articleID
    BlogAddImage.Show MODAL
    requeryImageList Me.newImageID
    Me.articleText.SetFocus
End Sub

Private Sub btnCancel_Click()
    If MsgBox("Are you sure you want to discard any changes you may have made?", vbYesNo) = vbYes Then
        unloadArticle
    End If
End Sub

Private Sub btnExit_Click()
    If Me.btnSave.Enabled Then
        If MsgBox("You will lose any changes you have made, exit anyway?", vbYesNo) = vbYes Then
            Unload BlogManager
        End If
    Else
        Unload BlogManager
    End If
End Sub

Private Sub btnSave_Click()
    Mouse.Hourglass True
    If Me.articleID = -1 Then
        Dim authid As String
        authid = DLookup("ID", "BlogAuthors", "UserName='" & EscapeSQuotes(Me.authorName) & "'")
        DB.Execute "INSERT INTO BlogArticles ( Title, Post, AuthorID ) VALUES ( '" & EscapeSQuotes(Me.articleTitle) & "', '" & EscapeSQuotes(Me.articleText) & "', " & authid & " )"
        Me.articleID = DLookup("@@IDENTITY", "BlogArticles")
    Else
        DB.Execute "UPDATE BlogArticles SET Title='" & EscapeSQuotes(Me.articleTitle) & "', Post='" & EscapeSQuotes(Me.articleText) & "' WHERE ID=" & Me.articleID
    End If
    'images are saved when they are added, just items here
    Dim itemsHash As Dictionary
    Set itemsHash = New Dictionary
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber FROM BlogItemAssociations WHERE ArticleID=" & Me.articleID)
    While Not rst.EOF
        itemsHash.Add rst("ItemNumber"), "1"
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Dim i As Long
    For i = 0 To Me.ItemsAssociated.ListCount - 1
        If itemsHash.exists(Me.ItemsAssociated.list(i)) Then
            itemsHash.Remove Me.ItemsAssociated.list(i)
        Else
            DB.Execute "INSERT INTO BlogItemAssociations ( ArticleID, ItemNumber ) VALUES ( " & Me.articleID & ", '" & Me.ItemsAssociated.list(i) & "' )"
        End If
    Next i
    Dim thisitem As Variant
    For Each thisitem In itemsHash.Keys
        DB.Execute "DELETE FROM BlogItemAssociations WHERE ArticleID=" & Me.articleID & " AND ItemNumber='" & thisitem & "'"
    Next thisitem
    Mouse.Hourglass False
End Sub

Private Sub Form_Load()
    DB.Execute "DELETE FROM BlogImages WHERE ArticleID NOT IN (SELECT ID FROM BlogArticles)"
    requeryItemsList
    requeryAuthors
    If InCombo(Environ("UserName"), Me.authorName) Then
        Me.authorName = Environ("UserName")
        currentAuthor = Me.authorName
        requeryArticles
    Else
        MsgBox "Your username isn't registered yet, either add a new author, or select someone else."
    End If
End Sub


Private Sub authorName_Click()
    authorName_LostFocus
End Sub

Private Sub authorName_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.authorName, KeyCode, Shift
    If KeyCode = vbKeyReturn Then
        authorName_LostFocus
    End If
End Sub

Private Sub authorName_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.authorName, KeyAscii
End Sub

Private Sub authorName_LostFocus()
    AutoCompleteLostFocus Me.authorName
    If Me.authorName = "" Then
        Me.authorName = ALL_AUTHORS
    End If
    If Me.authorName <> currentAuthor Then
        currentAuthor = Me.authorName
        requeryArticles
    End If
End Sub

Private Sub ItemPicker_Click()
    ItemPicker_LostFocus
End Sub

Private Sub ItemPicker_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.ItemPicker, KeyCode, Shift
    If KeyCode = vbKeyReturn Then
        ItemPicker_LostFocus
    End If
End Sub

Private Sub ItemPicker_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.ItemPicker, KeyAscii
End Sub

Private Sub ItemPicker_LostFocus()
    AutoCompleteLostFocus Me.ItemPicker
    If Me.ItemPicker <> "" Then
        If Not InList(Me.ItemPicker, Me.ItemsAssociated, True) Then
            Me.ItemsAssociated.AddItem Me.ItemPicker
        End If
    End If
End Sub

Private Sub ItemsAssociated_DblClick()
    If MsgBox("Remove " & Me.ItemsAssociated & "?", vbYesNo) = vbYes Then
        Me.ItemsAssociated.RemoveItem Me.ItemsAssociated.ListIndex
    End If
End Sub

Private Sub requeryItemsList()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber FROM vNonDiscontinuedItems ORDER BY ItemNumber")
    Me.ItemPicker.Clear
    While Not rst.EOF
        Me.ItemPicker.AddItem rst("ItemNumber")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    ExpandDropDownToFit Me.ItemPicker
End Sub

Private Sub requeryAuthors()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT UserName FROM BlogAuthors ORDER BY BlogName")
    Me.authorName.Clear
    Me.authorName.AddItem ALL_AUTHORS
    While Not rst.EOF
        Me.authorName.AddItem rst("UserName")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    ExpandDropDownToFit Me.authorName
End Sub

Private Sub requeryArticles()
    Dim rst As ADODB.Recordset
    Dim whereclause As String
    whereclause = IIf(Me.authorName = ALL_AUTHORS, "", "WHERE BlogAuthors.UserName='" & EscapeSQuotes(Me.authorName) & "'")
    Select Case True
        Case Is = Me.opShow(SHOW_ALL)
            'nothing
        Case Is = Me.opShow(SHOW_PUB)
            whereclause = IIf(whereclause = "", "", whereclause & " AND ") & " PostDate IS NOT NULL"
        Case Is = Me.opShow(SHOW_NEW)
            whereclause = IIf(whereclause = "", "", whereclause & " AND ") & " PostDate IS NULL"
    End Select
    Set rst = DB.retrieve("SELECT BlogArticles.ID, BlogArticles.Title FROM BlogArticles INNER JOIN BlogAuthors ON BlogArticles.AuthorID=BlogAuthors.ID " & whereclause & " ORDER BY BlogArticles.CreateDate DESC")
    Me.titleSelect.Clear
    While Not rst.EOF
        Me.titleSelect.AddItem IIf(Nz(rst("Title")) = "", NO_TITLE, rst("Title")) & vbTab & rst("ID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryImageList(Optional addNewNow As String = "-1")
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, Description, UNCAddress FROM BlogImages WHERE ArticleID=" & Me.articleID & " ORDER BY ID")
    Me.ImageList.Clear
    While Not rst.EOF
        Me.ImageList.AddItem rst("ID") & vbTab & IIf(Nz(rst("Description")) <> "", Nz(rst("Description")), filenameparse(Nz(rst("UNCAddress"))))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    If addNewNow > 0 Then
        Me.articleText = Left(Me.articleText, Me.articleText.SelStart) & "{pic" & addNewNow & "}" & Mid(Me.articleText, Me.articleText.SelStart + 1)
    End If
End Sub

Private Sub requeryItemAssoc()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber FROM BlogItemAssociations WHERE ArticleID=" & Me.articleID & " ORDER BY ItemNumber")
    Me.ItemsAssociated.Clear
    While Not rst.EOF
        Me.ItemsAssociated.AddItem rst("ItemNumber")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub opShow_Click(Index As Integer)
    requeryArticles
End Sub

Private Sub opView_Click(Index As Integer)
    Select Case Index
        Case Is = VIEW_TEXT
            Me.articleHTML.Visible = False
            Me.articlePreview.Visible = False
            Me.articleText.Visible = True
        Case Is = VIEW_HTML
            Me.articleHTML = text2html(Me.articleText, "{pic(left|right|center)(7)}")
            Me.articleHTML.Visible = True
            Me.articlePreview.Visible = False
            Me.articleText.Visible = False
        Case Is = VIEW_PAGE
            Me.articleHTML = text2html(Me.articleText, "{pic(left|right|center)(7)}")
            Open DESTINATION_DIR & "blogpreview.html" For Output As #1
            Print #1, Me.articleHTML
            Close #1
            Me.articlePreview.Navigate2 DESTINATION_DIR & "blogpreview.html"
            Me.articleHTML.Visible = False
            Me.articlePreview.Visible = True
            Me.articleText.Visible = False
    End Select
End Sub

Private Sub titleSelect_DblClick()
    loadArticle Mid(Me.titleSelect, InStr(Me.titleSelect, vbTab) + 1)
End Sub

Private Sub btnNewArticle_Click()
    loadArticle
End Sub

Private Sub loadArticle(Optional aid As Long = -1)
    Me.articleID = aid
    If aid <> -1 Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT Title, Post, CASE WHEN PostDate IS NULL THEN 0 ELSE 1 END AS Published FROM BlogArticles WHERE ID=" & aid)
        If rst.EOF Then
            MsgBox "Error, couldn't find article!"
        Else
            lockSelectors True
            lockEditor False, CBool(rst("Published"))
            Me.articleTitle = rst("Title")
            Me.articleText = rst("Post")
        End If
        rst.Close
        Set rst = Nothing
    Else
        lockSelectors True
        lockEditor False, False
    End If
    requeryImageList
    requeryItemAssoc
End Sub

Private Sub unloadArticle()
    lockSelectors False
    lockEditor True, True
End Sub

Private Sub lockSelectors(onoff As Boolean)
    If onoff Then
        Disable Me.authorName
        Disable Me.btnNewArticle
        Disable Me.opShow(SHOW_ALL)
        Disable Me.opShow(SHOW_PUB)
        Disable Me.opShow(SHOW_NEW)
        Disable Me.titleSelect
    Else
        Enable Me.authorName
        Enable Me.btnNewArticle
        Enable Me.opShow(SHOW_ALL)
        Enable Me.opShow(SHOW_PUB)
        Enable Me.opShow(SHOW_NEW)
        Enable Me.titleSelect
    End If
    Me.SelectFrame.Enabled = Not onoff
End Sub

Private Sub lockEditor(onoff As Boolean, readonly As Boolean)
    If onoff Then
        Disable Me.articleTitle
        Disable Me.articleText
        Disable Me.ImageList
        Disable Me.btnAddImage
        Disable Me.ItemPicker
        Disable Me.ItemsAssociated
    Else
        Enable Me.articleTitle
        Enable Me.articleText
        Enable Me.ImageList
        Enable Me.btnAddImage
        Enable Me.ItemPicker
        Enable Me.ItemsAssociated
    End If
    Me.ItemsFrame.Enabled = Not onoff
    Me.PicFrame.Enabled = Not onoff
    Me.articleTitle.Enabled = Not onoff
    Me.articleText.Enabled = Not onoff
    If onoff Then 'lock buttons too, readonly doesn't matter
        Me.btnCancel.Enabled = False
        Me.btnSave.Enabled = False
        Me.btnPublish.Enabled = False
        Me.articleText = "" 'and clear things too
        Me.articleTitle = ""
        Me.articleID = ""
    ElseIf readonly Then 'published?, lock buttons
        Me.btnCancel.Enabled = True
        Me.btnSave.Enabled = False
        Me.btnPublish.Enabled = False
    Else 'not published, so editable, keep buttons unlocked
        Me.btnCancel.Enabled = True
        Me.btnSave.Enabled = True
        Me.btnPublish.Enabled = True
    End If
End Sub

Private Sub lockSave(onoff As Boolean)
    Me.btnSave.Enabled = Not onoff
    Me.btnCancel.Enabled = Not onoff
    Me.btnPublish.Enabled = Not onoff
End Sub
