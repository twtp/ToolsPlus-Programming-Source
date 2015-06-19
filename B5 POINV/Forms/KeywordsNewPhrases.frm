VERSION 5.00
Begin VB.Form KeywordsNewPhrases 
   Caption         =   "Keywords - New Phrases"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox UsePhoneticMatching 
      Caption         =   "Use Phonetic Matching"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2250
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox UseRelevanceFilter 
      Caption         =   "Use Relevance Filter"
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1845
   End
   Begin VB.CommandButton btnStopNow 
      Caption         =   "Stop"
      Height          =   435
      Left            =   330
      TabIndex        =   11
      Top             =   3780
      Width           =   1215
   End
   Begin VB.CommandButton btnAddNever 
      Caption         =   "Never Show Selected Again"
      Height          =   465
      Left            =   8010
      TabIndex        =   9
      Top             =   4320
      Width           =   1725
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   1350
      TabIndex        =   8
      Top             =   4470
      Width           =   1545
   End
   Begin VB.CommandButton btnAddAll 
      Caption         =   "Add All"
      Height          =   405
      Left            =   5970
      TabIndex        =   7
      Top             =   4350
      Width           =   1275
   End
   Begin VB.CommandButton btnAddSelected 
      Caption         =   "Add Selected"
      Height          =   405
      Left            =   4530
      TabIndex        =   6
      Top             =   4350
      Width           =   1335
   End
   Begin VB.ListBox PhraseList 
      Height          =   4155
      Left            =   3960
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   60
      Width           =   5835
   End
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   3765
      Begin VB.OptionButton opType 
         Caption         =   "By Desc1+3"
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1140
         Width           =   1785
      End
      Begin VB.OptionButton opType 
         Caption         =   "By Path 2's"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1785
      End
      Begin VB.OptionButton opType 
         Caption         =   "By Manuf + Path 2's"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   1785
      End
   End
   Begin VB.CommandButton btnQuery 
      Caption         =   "Query"
      Height          =   525
      Left            =   330
      TabIndex        =   0
      Top             =   3210
      Width           =   1395
   End
   Begin VB.Label lblStatusBar 
      Height          =   225
      Left            =   30
      TabIndex        =   10
      Top             =   5160
      Width           =   4305
   End
End
Attribute VB_Name = "KeywordsNewPhrases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BY_PATH2 As Long = 0
Private Const BY_MANUF_PATH2 As Long = 1
Private Const BY_DESC13 As Long = 3

Private useless As Dictionary
Private wpKwds As Dictionary    'hash of variant arrays of words relevant to the web path
Private pnWPs As Dictionary     'hash of web paths for each item
Private stopNow As Boolean

Private Sub btnAddAll_Click()
    Mouse.Hourglass True
    lockForm True
    ReDim todo(Me.PhraseList.ListCount - 1) As String
    todo = ListBoxAsArray(Me.PhraseList, False)
    doList todo
    Me.PhraseList.Clear
    lockForm False
    Mouse.Hourglass False
End Sub

Private Sub btnAddNever_Click()
    Mouse.Hourglass True
    lockForm True
    ReDim todo(Me.PhraseList.SelCount - 1) As String
    todo = ListBoxAsArray(Me.PhraseList, True)
    doList todo, True
    RemoveSelectedFrom Me.PhraseList
    lockForm False
    Mouse.Hourglass False
End Sub

Private Sub btnAddSelected_Click()
    Mouse.Hourglass True
    lockForm True
    ReDim todo(Me.PhraseList.SelCount - 1) As String
    todo = ListBoxAsArray(Me.PhraseList, True)
    doList todo
    RemoveSelectedFrom Me.PhraseList
    lockForm False
    Mouse.Hourglass False
End Sub

Private Sub btnExit_Click()
    Unload KeywordsNewPhrases
End Sub

Private Sub btnQuery_Click()
    lockForm True
    Select Case True
        Case Is = Me.opType(BY_PATH2)
            createPhrasesByPath2
        Case Is = Me.opType(BY_MANUF_PATH2)
            createPhrasesByManufPath2
        Case Is = Me.opType(BY_DESC13)
            createPhrasesByDescription
        Case Else
            MsgBox "Select a creation type!"
    End Select
    lockForm False
End Sub

Private Sub btnStopNow_Click()
    stopNow = True
End Sub

Private Sub Form_Load()
    Dim tabs(0) As Long
    tabs(0) = 180
    SetListTabStops Me.PhraseList.hwnd, tabs
End Sub

Private Sub createPhrasesByPath2()

End Sub

Private Sub createPhrasesByManufPath2()
    stopNow = False
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ProductLine, WebPathName, ManufFullNameWeb, WebPathSearchableName FROM vKeywordsNewPhrasesByManufPath2 ORDER BY ProductLine, WebPathName")
    Dim phrases As Variant, lp As String, i As Long
    Me.PhraseList.Clear
    Dim listHash As Dictionary
    Set listHash = New Dictionary
    While Not rst.EOF
        Me.lblStatusBar.Caption = "working on " & rst("ProductLine") & " " & rst("WebPathName")
        DoEvents
        If stopNow Then
            GoTo cleanup
        End If
        If rst("WebPathName") <> "1no-path" Then
            lp = LCase(URLify(rst("ProductLine") & "-" & rst("WebPathName")) & ".html")
            phrases = Split(rst("WebPathSearchableName"), ",")
            For i = LBound(phrases) To UBound(phrases)
                phrases(i) = TrimWhitespace(CStr(phrases(i)))
                If Not listHash.exists(CStr(phrases(i))) Then
                    listHash.Add CStr(phrases(i)), "1"
                    If DLookup("SearchPhrase", "KeywordsPhrases", "SearchPhrase='" & EscapeSQuotes(CStr(phrases(i))) & "'") = "" Then
                        Me.PhraseList.AddItem LCase(rst("ManufFullNameWeb") & " " & TrimWhitespace(CStr(phrases(i)))) & vbTab & lp
                    End If
                End If
            Next i
        End If
        rst.MoveNext
        DoEvents
    Wend
cleanup:
    rst.Close
    Set rst = Nothing
End Sub

Private Sub createPhrasesByDescription()
    stopNow = False
    Mouse.Hourglass True
    cacheUseless
    If Me.UseRelevanceFilter Then
        cacheRelevance
    End If
    Dim rst As ADODB.Recordset, i As Long, j As Long
    Set rst = DB.retrieve("SELECT ItemNumber, Desc1, Desc3 FROM PartNumbers WHERE WebPublished=1")
    Me.PhraseList.Clear
    Dim phrase As String, words As Variant, phrases As Variant
    Dim listHash As Dictionary
    Set listHash = New Dictionary
    While Not rst.EOF
        Me.lblStatusBar.Caption = "working on " & rst("ItemNumber")
        DoEvents
        If stopNow Then
            GoTo cleanup
        End If
        phrase = cleanup(Nz(rst("Desc1")))
        phrase = IIf(phrase = "", "", phrase & " ") & rst("Desc3")
        phrase = FootMarksToAbbrev(InchMarksToAbbrev(phrase))
        phrase = Replace(phrase, ",", "")
        phrase = Replace(phrase, "(", " ")
        phrase = Replace(phrase, ")", " ")
        phrase = Replace(phrase, Chr(169), " ") '(c)
        phrase = Replace(phrase, Chr(174), " ") 'reg
        phrase = Replace(phrase, Chr(153), " ") 'tm
        phrase = Replace(phrase, "'", "")
        phrase = Replace(phrase, """", "")
        'phrase = Replace(phrase, "/", " ")
        'phrase = Replace(phrase, "\", " ")
        phrase = TrimWhitespace(phrase, True, True, True)
        If phrase <> "" Then
            words = Split(phrase, " ")
            For i = LBound(words) To UBound(words)
                words(i) = LCase(TrimWhitespace(CStr(words(i))))
            Next i
            words = removeUseless(words)
            words = dropRepeated(words)
            For i = LBound(words) To IIf(UBound(words) < 3, UBound(words), 3)
                phrases = Combine(words, i + 1)
                For j = LBound(phrases) To UBound(phrases)
                    phrases(j) = addSpacesToUnits(CStr(phrases(j)))
                    If Me.UseRelevanceFilter Then
                        If isRelevant(CStr(phrases(j)), rst("ItemNumber")) Then
                            If Not listHash.exists(CStr(phrases(j))) Then
                                listHash.Add CStr(phrases(j)), "1"
                                Me.PhraseList.AddItem phrases(j)
                            End If
                        End If
                    Else
                        If Not listHash.exists(CStr(phrases(j))) Then
                            listHash.Add CStr(phrases(j)), "1"
                            Me.PhraseList.AddItem phrases(j)
                        End If
                    End If
                Next j
            Next i
        End If
        rst.MoveNext
    Wend
cleanup:
    rst.Close
    Set rst = Nothing
    MsgBox listHash.count
    Set listHash = Nothing
    Me.lblStatusBar.Caption = ""
    Mouse.Hourglass False
End Sub

Private Sub doList(phrases() As String, Optional delete As Boolean = False)
    Dim i As Long, phrase As String, lp As String
    For i = 0 To UBound(phrases)
        phrase = Left(phrases(i), InStr(phrases(i), vbTab) - 1)
        lp = Mid(phrases(i), InStr(phrases(i), vbTab) + 1)
        Me.lblStatusBar.Caption = "Working on " & phrase
        DoEvents
        DB.Execute "INSERT INTO KeywordsPhrases ( SearchPhrase, LandingPage" & IIf(delete, ", Pending, DoNotUse", "") & " ) VALUES ( '" & EscapeSQuotes(phrase) & "', '" & lp & "'" & IIf(delete, ", 0, 1", "") & " )"
    Next i
    Me.lblStatusBar.Caption = ""
End Sub

Private Sub lockForm(yesno As Boolean)
    Me.btnAddAll.Enabled = Not yesno
    Me.btnAddSelected.Enabled = Not yesno
    Me.btnQuery.Enabled = Not yesno
    Me.btnExit.Enabled = Not yesno
    Me.btnAddNever.Enabled = Not yesno
End Sub

Private Function cleanup(phrase As String) As String
    Dim retval As String
    retval = phrase
    retval = Replace(retval, "FREE SHIPPING TO THE 48 STATES", "", , , vbTextCompare)
    retval = Replace(retval, "REBATES", "", , , vbTextCompare)
    retval = Replace(retval, "ONLY $6.50 FREIGHT TO THE 48", "", , , vbTextCompare)
    retval = Replace(retval, "FREE SHIPPING within 100 miles of Waterbury CT", "", , , vbTextCompare)
    cleanup = retval
End Function

Private Sub cacheUseless()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT word FROM KeywordsGenerationUselessWords")
    Set useless = New Dictionary
    While Not rst.EOF
        useless.Add CStr(LCase(rst("word").value)), "1"
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub cacheRelevance()
    Set wpKwds = New Dictionary
    Set pnWPs = New Dictionary
    Dim words As Variant, i As Long
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT WebPathSearchableName, ID FROM WebPaths WHERE PathLevel=2 OR PathLevel=3")
    While Not rst.EOF
        If Not wpKwds.exists(CStr(rst("ID"))) Then
            'split along , to create relevant phrases, not words
            'words = Split(Replace(rst("WebPathSearchableName"), ",", ""), " ")
            words = Split(rst("WebPathSearchableName"), ",")
            For i = LBound(words) To UBound(words)
                words(i) = TrimWhitespace(CStr(words(i)), True, True, True)
            Next i
            words = dropRepeated(words)
            wpKwds.Add CStr(rst("ID")), words
        End If
        rst.MoveNext
    Wend
    rst.Close
    ReDim paths(0) As String
    Set rst = DB.retrieve("SELECT ItemNumber, WebPathID FROM PartNumberWebPaths INNER JOIN WebPaths ON WebPathID=ID WHERE PathLevel=2 OR PathLevel=3")
    While Not rst.EOF
        If pnWPs.exists(CStr(rst("ItemNumber"))) Then
            paths = pnWPs.item(CStr(rst("ItemNumber")))
            ReDim Preserve paths(UBound(paths) + 1) As String
            paths(UBound(paths)) = CStr(rst("WebPathID"))
            pnWPs.item(CStr(rst("ItemNumber"))) = paths
        Else
            If UBound(paths) <> 0 Then
                ReDim paths(0) As String
            End If
            paths(0) = CStr(rst("WebPathID"))
            pnWPs.Add CStr(rst("ItemNumber")), paths
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Function removeUseless(words As Variant) As Variant
    If useless Is Nothing Then
        MsgBox "cacheUseless() was not called, so nothing's going to get filtered!"
        Set useless = New Dictionary
    End If
    Dim i As Long, count As Long
    ReDim retval(UBound(words)) As Variant
    For i = LBound(words) To UBound(words)
        If Not useless.exists(CStr(LCase(words(i)))) Then
            retval(count) = words(i)
            count = count + 1
        End If
    Next i
    If count - 1 <> UBound(words) Then
        ReDim Preserve retval(count - 1) As Variant
    End If
    removeUseless = retval
End Function

Private Function dropRepeated(words As Variant) As Variant
    Dim i As Long, j As Long, count As Long, found As Boolean
    ReDim retval(UBound(words)) As Variant
    For i = LBound(words) To UBound(words)
        found = False
        For j = 0 To i - 1
            If words(j) = words(i) Then
                found = True
                j = i
            End If
        Next j
        If Not found Then
            retval(count) = words(i)
            count = count + 1
        End If
    Next i
    If count - 1 <> UBound(words) Then
        ReDim Preserve retval(count - 1) As Variant
    End If
    dropRepeated = retval
End Function

Private Function addSpacesToUnits(txt As String) As String
    If InStr(txt, "inch") Or InStr(txt, "ft") Then
        Dim retval As String, i As Long
        For i = 1 To Len(txt)
            If IsNumeric(Mid(txt, i, 1)) Then
                If Mid(txt, i + 1, 4) = "inch" Then
                    retval = retval & Mid(txt, i, 1) & " inch"
                    i = i + 4
                ElseIf Mid(txt, i + 1, 2) = "ft" Then
                    retval = retval & Mid(txt, i, 1) & " ft"
                    i = i + 2
                Else
                    retval = retval & Mid(txt, i, 1)
                End If
            Else
                retval = retval & Mid(txt, i, 1)
            End If
        Next i
        addSpacesToUnits = retval
    Else
        addSpacesToUnits = txt
    End If
End Function

Private Function isRelevant(phrase As String, item As String) As Boolean
    isRelevant = False
    Dim WPs As Variant, words As Variant, i As Long, j As Long
    If pnWPs.exists(item) Then
        WPs = pnWPs.item(item)
        For i = LBound(WPs) To UBound(WPs)
            If wpKwds.exists(CStr(WPs(i))) Or wpKwds.exists(CStr(Pluralize(CStr(WPs(i))))) Then
                words = wpKwds.item(CStr(WPs(i)))
                For j = LBound(words) To UBound(words)
                    If InStr(phrase, words(j)) Then
                        isRelevant = True
                        Debug.Print "PASSED (" & item & "): " & phrase
                        Exit Function
                    End If
                Next j
            End If
        Next i
    Else
        isRelevant = True
        Debug.Print "SKIPPED (" & item & "): " & phrase
    End If
    Debug.Print "FAILED: (" & item & "): " & phrase
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set useless = Nothing
    Set pnWPs = Nothing
    Set wpKwds = Nothing
End Sub
