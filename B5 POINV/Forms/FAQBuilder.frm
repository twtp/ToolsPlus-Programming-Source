VERSION 5.00
Begin VB.Form FAQBuilder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FAQ Builder"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkActive 
      Caption         =   "Active"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7440
      TabIndex        =   15
      Top             =   960
      Width           =   1905
   End
   Begin VB.TextBox QID 
      Height          =   285
      Left            =   5880
      TabIndex        =   14
      Text            =   "QID"
      Top             =   210
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton btnEmailKirk 
      Caption         =   "Email HTML to Kirk"
      Height          =   585
      Left            =   6870
      TabIndex        =   13
      Top             =   4740
      Width           =   1725
   End
   Begin VB.CommandButton btnGeneratePage 
      Caption         =   "Preview HTML"
      Height          =   585
      Left            =   5130
      TabIndex        =   12
      Top             =   4740
      Width           =   1725
   End
   Begin VB.CommandButton btnMoveDown 
      Caption         =   "q"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   300
      TabIndex        =   11
      Top             =   5040
      Width           =   285
   End
   Begin VB.CommandButton btnMoveUp 
      Caption         =   "p"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   300
      TabIndex        =   10
      Top             =   4740
      Width           =   285
   End
   Begin VB.CommandButton btnDeleteThis 
      Caption         =   "Delete Selected"
      Height          =   585
      Left            =   3090
      TabIndex        =   9
      Top             =   4740
      Width           =   1185
   End
   Begin VB.CommandButton btnAddNewQuestion 
      Caption         =   "Add New Question"
      Height          =   585
      Left            =   1860
      TabIndex        =   8
      Top             =   4740
      Width           =   1185
   End
   Begin VB.CommandButton btnAddNewCategory 
      Caption         =   "Add New Category"
      Height          =   585
      Left            =   630
      TabIndex        =   7
      Top             =   4740
      Width           =   1185
   End
   Begin VB.TextBox Answer 
      Height          =   2355
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2310
      Width           =   4755
   End
   Begin VB.TextBox Question 
      Height          =   645
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1260
      Width           =   4755
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   495
      Left            =   8340
      TabIndex        =   2
      Top             =   0
      Width           =   1155
   End
   Begin VB.ListBox FAQTree 
      Height          =   4155
      Left            =   90
      TabIndex        =   0
      Top             =   510
      Width           =   4485
   End
   Begin VB.Label Label3 
      Caption         =   "Answer:"
      Height          =   255
      Left            =   4710
      TabIndex        =   5
      Top             =   2070
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Heading:"
      Height          =   255
      Left            =   4710
      TabIndex        =   4
      Top             =   1020
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "FAQ Page Builder"
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
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4905
   End
End
Attribute VB_Name = "FAQBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private changed As Boolean
Private whichCtl As String
Private fillingForm As Boolean

Private Sub btnAddNewCategory_Click()
    Dim cat As String
    cat = InputBox("Enter category name:")
    If cat <> "" Then
        Dim sorder As String
        sorder = DLookup("MAX(SortOrder)", "FAQCategories")
        If sorder = "" Then
            sorder = 0
        Else
            sorder = sorder + 1
        End If
        DB.Execute "INSERT INTO FAQCategories ( CategoryName, SortOrder ) VALUES ( '" & EscapeSQuotes(cat) & "', " & sorder & " )"
        Me.FAQTree.AddItem cat & vbTab & DLookup("@@IDENTITY", "FAQCategories")
        Me.FAQTree.Selected(Me.FAQTree.ListCount - 1) = True
        Me.Question.SetFocus
    End If
End Sub

Private Sub btnAddNewQuestion_Click()
    If Me.FAQTree.SelCount = 0 Then
        MsgBox "Select a category to add this question into first."
    Else
        Dim q As String
        q = InputBox("Enter question:")
        If q <> "" Then
            Dim pid As String
            pid = getParentID()
            Dim sorder As String
            sorder = DLookup("MAX(SortOrder)", "FAQQuestions", "CategoryID=" & pid)
            If sorder = "" Then
                sorder = 0
            Else
                sorder = sorder + 1
            End If
            DB.Execute "INSERT INTO FAQQuestions ( CategoryID, Question, Answer, SortOrder ) VALUES ( " & pid & ", '" & EscapeSQuotes(q) & "', 'Enter answer text here...', " & sorder & " )"
            Dim i As Long
            i = Me.FAQTree.ListIndex + 1
            While i <> Me.FAQTree.ListCount And isQuestion(i)
                i = i + 1
            Wend
            Me.FAQTree.AddItem "      " & q & vbTab & DLookup("@@IDENTITY", "FAQQuestions"), i
            Me.FAQTree.Selected(i) = True
            Me.Answer.SetFocus
            Me.Answer.selStart = 0
            Me.Answer.SelLength = Len(Me.Answer)
        End If
    End If
End Sub

Private Sub btnDeleteThis_Click()
    Dim sorder As String, id As String, pid As String
    If Me.FAQTree.SelCount <> 0 Then
        If isQuestion() Then
            If vbYes = MsgBox("Delete " & qq(getSelectedText()) & "?", vbYesNo) Then
                id = getSelectedID()
                sorder = DLookup("SortOrder", "FAQQuestions", "ID=" & id)
                pid = DLookup("CategoryID", "FAQQuestions", "ID=" & id)
                DB.Execute "DELETE FROM FAQQuestions WHERE ID=?"
                DB.Execute "UPDATE FAQQuestions SET SortOrder=SortOrder-1 WHERE CategoryID=" & pid & " AND SortOrder>" & sorder
                Me.FAQTree.RemoveItem Me.FAQTree.ListIndex
            End If
        Else
            If vbYes = MsgBox("Delete category " & qq(getSelectedText()) & "?" & vbCrLf & vbCrLf & "This will delete any questions in this category!", vbYesNo) Then
                id = getSelectedID()
                sorder = DLookup("SortOrder", "FAQCategories", "ID=" & id)
                DB.Execute "DELETE FROM FAQCategories WHERE ID=" & id
                DB.Execute "UPDATE FAQCategories SET SortOrder=SortOrder-1 WHERE SortOrder>" & sorder
                DB.Execute "DELETE FROM FAQQuestions WHERE CategoryID=" & id
                Dim i As Long
                i = Me.FAQTree.ListIndex
                Me.FAQTree.RemoveItem i
                While i <> Me.FAQTree.ListCount And isQuestion(i)
                    Me.FAQTree.RemoveItem i
                Wend
            End If
        End If
    End If
End Sub

Private Sub btnEmailKirk_Click()
    OpenEmailTo "kirk@tools-plus.com", "FAQ Update", generateHTML()
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnGeneratePage_Click()
    MsgBox generateHTML()
End Sub

Private Sub btnMoveDown_Click()
    If Me.FAQTree.SelCount <> 0 And Me.FAQTree.ListIndex <> Me.FAQTree.ListCount Then
        If isQuestion() Then
            If isQuestion(Me.FAQTree.ListIndex + 1) Then
                'another question, ok to move
                DB.Execute "UPDATE FAQQuestions SET SortOrder=SortOrder+1 WHERE ID=" & getSelectedID()
                DB.Execute "UPDATE FAQQuestions SET SortOrder=SortOrder-1 WHERE ID=" & getNextID()
                Dim temp As String
                temp = getSelectedAll()
                Me.FAQTree.List(Me.FAQTree.ListIndex) = getNextAll()
                Me.FAQTree.List(Me.FAQTree.ListIndex + 1) = temp
                Me.FAQTree.ListIndex = Me.FAQTree.ListIndex + 1
            End If
        Else
            Dim i As Long
            i = Me.FAQTree.ListIndex + 1
            While i < Me.FAQTree.ListCount And isQuestion(i)
                i = i + 1
            Wend
            'ok to move?
            If Not isQuestion(i) And i <> Me.FAQTree.ListCount Then
                Dim txt As String
                txt = getSelectedAll()
                DB.Execute "UPDATE FAQCategories SET SortOrder=SortOrder+1 WHERE ID=" & getSelectedID()
                i = Me.FAQTree.ListIndex + 1
                While isQuestion(i)
                    i = i + 1
                Wend
                DB.Execute "UPDATE FAQCategories SET SortOrder=SortOrder-1 WHERE ID=" & getParentID(i)
                requeryFAQTree
                For i = 0 To Me.FAQTree.ListCount - 1
                    If Me.FAQTree.List(i) = txt Then
                        Me.FAQTree.Selected(i) = True
                        Exit For
                    End If
                Next i
            End If
        End If
    End If
End Sub

Private Sub btnMoveUp_Click()
    If Me.FAQTree.SelCount <> 0 And Me.FAQTree.ListIndex <> 0 Then
        If isQuestion() Then
            If isQuestion(Me.FAQTree.ListIndex - 1) Then
                'another question, ok to move
                DB.Execute "UPDATE FAQQuestions SET SortOrder=SortOrder-1 WHERE ID=" & getSelectedID()
                DB.Execute "UPDATE FAQQuestions SET SortOrder=SortOrder+1 WHERE ID=" & getPrevID()
                Dim temp As String
                temp = getSelectedAll()
                Me.FAQTree.List(Me.FAQTree.ListIndex) = getPrevAll()
                Me.FAQTree.List(Me.FAQTree.ListIndex - 1) = temp
                Me.FAQTree.ListIndex = Me.FAQTree.ListIndex - 1
            End If
        Else
            'category, not first, ok to move
            Dim txt As String
            txt = getSelectedAll()
            DB.Execute "UPDATE FAQCategories SET SortOrder=SortOrder-1 WHERE ID=" & getSelectedID()
            If isQuestion(Me.FAQTree.ListIndex - 1) Then
                DB.Execute "UPDATE FAQCategories SET SortOrder=SortOrder+1 WHERE ID=(SELECT CategoryID FROM FAQQuestions WHERE ID=" & getPrevID() & ")"
            Else
                DB.Execute "UPDATE FAQCategories SET SortOrder=SortOrder+1 WHERE ID=" & getPrevID()
            End If
            requeryFAQTree
            Dim i As Long
            For i = 0 To Me.FAQTree.ListCount - 1
                If Me.FAQTree.List(i) = txt Then
                    Me.FAQTree.Selected(i) = True
                    Exit For
                End If
            Next i
        End If
    End If
End Sub

Private Sub chkActive_Click()
    If Not fillingForm Then
        DB.Execute "UPDATE FAQQuestions SET Active=" & Me.chkActive & " WHERE ID=" & getSelectedID()
    End If
End Sub

Private Sub FAQTree_Click()
    'Debug.Print "FAQTree_Click"
    fillingForm = True
    Me.QID = getSelectedID()
    Me.Question = getSelectedText()
    Enable Me.Question
    If isQuestion() Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT Answer, Active FROM FAQQuestions WHERE ID=" & getSelectedID())
        Me.Answer = rst("Answer")
        Me.chkActive = SQLBool(rst("Active"))
        Enable Me.Answer
        Enable Me.chkActive
    Else
        Me.Answer = ""
        Me.chkActive = 0
        Disable Me.Answer
        Disable Me.chkActive
    End If
    fillingForm = False
End Sub

Private Sub Form_Load()
    requeryFAQTree
    Disable Me.Question
    Disable Me.Answer
    SetListTabs Me.FAQTree, Array(300)
End Sub

Private Sub Question_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            Question_KeyPress KeyCode
        Case Is = vbKeyReturn
            Question_LostFocus
    End Select
End Sub

Private Sub Question_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "Question"
End Sub

Private Sub Question_LostFocus()
    If changed = True Then
        Me.Question = Replace(Me.Question, vbCrLf, " ")
        Me.Question = TrimWhitespace(StripHTML(Me.Question), True, True, True, True)
        If isQuestion() Then
            DB.Execute "UPDATE FAQQuestions SET Question='" & EscapeSQuotes(Me.Question) & "' WHERE ID=" & Me.QID
        Else
            DB.Execute "UPDATE FAQCategories SET CategoryName='" & EscapeSQuotes(Me.Question) & "' WHERE ID=" & Me.QID
        End If
        changed = False
    End If
End Sub

Private Sub Answer_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyDelete
            Answer_KeyPress KeyCode
        Case Is = vbKeyReturn
            Answer_LostFocus
    End Select
End Sub

Private Sub Answer_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "Answer"
End Sub

Private Sub Answer_LostFocus()
    'Debug.Print "Answer_LostFocus"
    If changed = True Then
    '    Debug.Print "  changed - updating"
        Me.Answer = Trim(Me.Answer)
        DB.Execute "UPDATE FAQQuestions SET Answer='" & EscapeSQuotes(Me.Answer) & "' WHERE ID=" & Me.QID
        changed = False
    End If
End Sub

Private Sub requeryFAQTree()
    Me.FAQTree.Clear
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT FAQCategories.ID AS CategoryID, FAQQuestions.ID AS QuestionID, FAQQuestions.Question, FAQQuestions.Answer, FAQCategories.CategoryName FROM FAQCategories LEFT OUTER JOIN FAQQuestions ON FAQCategories.ID=FAQQuestions.CategoryID ORDER BY FAQCategories.SortOrder, FAQQuestions.SortOrder")
    Dim lastCatID As Long
    lastCatID = -1
    While Not rst.EOF
        If rst("CategoryID") <> lastCatID Then
            lastCatID = rst("CategoryID")
            Me.FAQTree.AddItem rst("CategoryName") & vbTab & rst("CategoryID")
        End If
        If Nz(rst("QuestionID")) <> "" Then
            Me.FAQTree.AddItem "      " & rst("Question") & vbTab & rst("QuestionID")
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Function generateHTML() As String
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT FAQQuestions.CategoryID, FAQQuestions.ID AS QuestionID, FAQQuestions.Question, FAQQuestions.Answer, FAQCategories.CategoryName FROM FAQQuestions INNER JOIN FAQCategories ON FAQQuestions.CategoryID=FAQCategories.ID WHERE FAQQuestions.Active=1 ORDER BY FAQCategories.SortOrder, FAQQuestions.SortOrder")
    Dim lastCatID As Long
    lastCatID = -1
    Dim html As String
    html = "<!-- DO NOT EDIT THIS PAGE DIRECTLY, USE THE FAQ BUILDER PROGRAM (ON THE UTILITIES MENU) -->" & vbCrLf
    html = html & "<ul id=""faqbody"">" & vbCrLf
    While Not rst.EOF
        If rst("CategoryID") <> lastCatID Then
            If lastCatID <> -1 Then
                html = html & "  </dl>" & vbCrLf
                html = html & " </li>" & vbCrLf
            End If
            lastCatID = rst("CategoryID")
            html = html & " <li><h3>" & rst("CategoryName") & "</h3>" & vbCrLf
            html = html & "  <dl>" & vbCrLf
        End If
        html = html & "   <dt><a href=""#"" onClick=""flip(this); return false;"">" & rst("Question") & "</a></dt>" & vbCrLf
        html = html & "   <dd class=""answer""><p>" & Replace(Replace(rst("Answer"), vbCrLf & vbCrLf, "</p><p>"), vbCrLf, "<br />") & "</p></dd>" & vbCrLf
        rst.MoveNext
    Wend
    html = html & "  </dl>" & vbCrLf
    html = html & " </li>" & vbCrLf
    html = html & "</ul>" & vbCrLf
    rst.Close
    Set rst = Nothing
    
    html = html & vbCrLf
    html = html & "<script type=""text/javascript"">" & vbCrLf
    html = html & "  function reset() {" & vbCrLf
    html = html & "    var dds = document.getElementById(""faqbody"").getElementsByTagName(""dd"");" & vbCrLf
    html = html & "    for (var i=0; i<dds.length; i++) {" & vbCrLf
    html = html & "      dds[i].style.display = ""none"";" & vbCrLf
    html = html & "    }" & vbCrLf
    html = html & "  }" & vbCrLf
    html = html & vbCrLf
    html = html & "  function flip(thelink) {" & vbCrLf
    html = html & "    var thedd = thelink.parentNode.nextSibling;" & vbCrLf
    html = html & "    while (thedd.nodeType != 1 /*Node.ELEMENT_NODE*/) {" & vbCrLf
    html = html & "      thedd = thedd.nextSibling;" & vbCrLf
    html = html & "    }" & vbCrLf
    html = html & "    thedd.style.display = thedd.style.display==""none"" ? ""block"" : ""none"";" & vbCrLf
    html = html & "  }" & vbCrLf
    html = html & vbCrLf
    html = html & "  reset();" & vbCrLf
    html = html & "</script>" & vbCrLf
    
    generateHTML = html
End Function

Private Function isQuestion(Optional dx As Long = -1) As Boolean
    If dx = -1 Then
        isQuestion = CBool(Left(Me.FAQTree, 6) = "      ")
    Else
        isQuestion = CBool(Left(Me.FAQTree.List(dx), 6) = "      ")
    End If
End Function

Private Function getSelectedAll() As String
    getSelectedAll = Me.FAQTree
End Function

Private Function getSelectedText() As String
    getSelectedText = Trim(Left(Me.FAQTree, InStr(Me.FAQTree, vbTab) - 1))
End Function

Private Function getSelectedID() As String
    getSelectedID = Mid(Me.FAQTree, InStr(Me.FAQTree, vbTab) + 1)
End Function

Private Function getNextAll() As String
    getNextAll = Me.FAQTree.List(Me.FAQTree.ListIndex + 1)
End Function

Private Function getNextText() As String
    getNextText = Trim(Left(Me.FAQTree.List(Me.FAQTree.ListIndex + 1), InStr(Me.FAQTree.List(Me.FAQTree.ListIndex + 1), vbTab) + 1))
End Function

Private Function getNextID() As String
    getNextID = Mid(Me.FAQTree.List(Me.FAQTree.ListIndex + 1), InStr(Me.FAQTree.List(Me.FAQTree.ListIndex + 1), vbTab) + 1)
End Function

Private Function getPrevAll() As String
    getPrevAll = Me.FAQTree.List(Me.FAQTree.ListIndex - 1)
End Function

Private Function getPrevText() As String
    getPrevText = Trim(Left(Me.FAQTree.List(Me.FAQTree.ListIndex - 1), InStr(Me.FAQTree.List(Me.FAQTree.ListIndex - 1), vbTab) + 1))
End Function

Private Function getPrevID() As String
    getPrevID = Mid(Me.FAQTree.List(Me.FAQTree.ListIndex - 1), InStr(Me.FAQTree.List(Me.FAQTree.ListIndex - 1), vbTab) + 1)
End Function

Private Function getParentText(Optional dx As Long = -1) As String
    Dim i As Long
    If dx = -1 Then
        i = Me.FAQTree.ListIndex
    Else
        i = dx
    End If
    While isQuestion(i)
        i = i - 1
    Wend
    getParentText = Left(Me.FAQTree.List(i), InStr(Me.FAQTree.List(i), vbTab) - 1)
End Function

Private Function getParentID(Optional dx As Long = -1) As String
    Dim i As Long
    If dx = -1 Then
        i = Me.FAQTree.ListIndex
    Else
        i = dx
    End If
    While isQuestion(i)
        i = i - 1
    Wend
    getParentID = Mid(Me.FAQTree.List(i), InStr(Me.FAQTree.List(i), vbTab) + 1)
End Function
