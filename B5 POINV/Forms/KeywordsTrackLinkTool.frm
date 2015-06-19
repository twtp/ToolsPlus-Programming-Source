VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form KeywordsTrackLinkTool 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keywords - Create/Delete Tracklinks"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   12345
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnStop 
      Caption         =   "STOP!"
      Height          =   495
      Left            =   8700
      TabIndex        =   9
      Top             =   5520
      Width           =   885
   End
   Begin VB.CommandButton btnRefreshTL 
      Caption         =   "Clear + Refresh Tracklinks! (long time)"
      Height          =   495
      Left            =   6600
      TabIndex        =   8
      Top             =   5520
      Width           =   2025
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   495
      Left            =   10320
      TabIndex        =   7
      Top             =   5520
      Width           =   1545
   End
   Begin VB.CommandButton btnDeleteTL 
      Caption         =   "Delete TrackLinks!"
      Height          =   495
      Left            =   4500
      TabIndex        =   6
      Top             =   5520
      Width           =   2025
   End
   Begin VB.CommandButton btnCreateLinks 
      Caption         =   "Create Tracklinks!"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   5520
      Width           =   2025
   End
   Begin VB.ListBox deleteTL 
      Height          =   2595
      ItemData        =   "KeywordsTrackLinkTool.frx":0000
      Left            =   60
      List            =   "KeywordsTrackLinkTool.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   3600
      Width           =   2115
   End
   Begin VB.ListBox needTL 
      Height          =   2985
      ItemData        =   "KeywordsTrackLinkTool.frx":0004
      Left            =   60
      List            =   "KeywordsTrackLinkTool.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2115
   End
   Begin SHDocVwCtl.WebBrowser browser 
      Height          =   5205
      Left            =   2250
      TabIndex        =   4
      Top             =   60
      Width           =   10095
      ExtentX         =   17806
      ExtentY         =   9181
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
      Location        =   "http:///"
   End
   Begin VB.Label Label2 
      Caption         =   "to be deleted:"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   3390
      Width           =   1725
   End
   Begin VB.Label Label1 
      Caption         =   "still need tracklinks:"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   30
      Width           =   1725
   End
End
Attribute VB_Name = "KeywordsTrackLinkTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const START_PAGE As String = "http://edit.store.yahoo.com/RT/MGR.toolsplus/"

'everything from now on should probably be 4 letters to make it easier
Private Const YAHOO_PHRASE_PREFIX As String = "KWD"
Private Const YAHOO_ITEM_PREFIX As String = "ITEM"
Private Const GOOGLE_PHRASE_PREFIX As String = "GKWD"
Private Const GOOGLE_ITEM_PREFIX As String = "GITM"

Private Const CREATE_PAGE_TIMEOUT As Long = 180

Private complete As Boolean
Private stopNow As Boolean

Private Sub browser_DocumentComplete(ByVal pDisp As Object, url As Variant)
    complete = True
End Sub

Private Sub btnCreateLinks_Click()
    Mouse.Hourglass True
    complete = False
    stopNow = False
    Me.browser.Navigate2 START_PAGE
    waitForPage
    Me.browser.Navigate2 findLink("Create Links")
    waitForPage CREATE_PAGE_TIMEOUT
    
    Dim i As Long
    Dim href_for As Dictionary
    Set href_for = New Dictionary
    Dim linkText As String, linkHref As String
    Dim to_delete As PerlArray
    Set to_delete = New PerlArray
    'for each link
    For i = 0 To Me.browser.Document.Links.Length - 1
        linkText = Me.browser.Document.Links(i).innerText
        'if it's a keyword link
        If hasPrefix(linkText) Then
            'and we need it
            If InList(linkText, Me.needTL) Then
                linkHref = Me.browser.Document.Links(i).href
                If Not href_for.exists(linkText) Then
                    'add it to the hash if we can
                    href_for.Add linkText, linkHref
                Else
                    'otherwise kill it
                    to_delete.Push linkHref
                End If
            End If
        End If
    Next i
    
    'now, we refresh any null TLs we found
    For i = 0 To href_for.count - 1
        If stopNow Then
            GoTo done
        End If
        linkText = href_for.Keys(i)
        linkHref = href_for(linkText)
        Me.browser.Navigate2 linkHref
        waitForPage
        SetTrackLink cleanKwd(linkText), getSvc(linkText), findTrackLink(Me.browser.Document.body.innerText)
    Next i
    
    'and then delete any duplicates found
    For i = 0 To to_delete.Scalar - 1
        If stopNow Then
            GoTo done
        End If
        Me.browser.Navigate2 to_delete.Elem(CInt(i))
        waitForPage
        Me.browser.Navigate2 findLink("Delete this link")
        waitForPage CREATE_PAGE_TIMEOUT
    Next i
    
    'now we're ready to start creating new keywords, back to the create links page if necessary
    If href_for.count <> 0 And to_delete.Scalar <> 0 Then
        requeryNeedTL
        Me.browser.Navigate2 START_PAGE
        waitForPage
        Me.browser.Navigate2 findLink("Create Links")
        waitForPage CREATE_PAGE_TIMEOUT
    End If
    
    'start at the selected/beginning
    Dim entered As Boolean
    Dim j As Long
    For i = IIf(Me.needTL.SelCount, Me.needTL.ListIndex, 0) To Me.needTL.ListCount - 1
        entered = False
        linkText = Me.needTL.list(i)
        
        With Me.browser.Document.Forms(0)
            'for each element in the form
            For j = 0 To .ALL.Length
                'if it's an input field
                If UCase(.ALL(j).nodeName) = "INPUT" Then
                    'of type text
                    If UCase(.ALL(j).Type) = "TEXT" Then
                        'put the text in there, and bail
                        .ALL(j).value = linkText
                        entered = True
                        Exit For
                    End If
                End If
            Next j
        End With
        
        'can't find the text field, just skip ahead to the end
        If Not entered Then
            GoTo done
        End If
        
        'make the tracklink
        clickSubmitButton
        SetTrackLink cleanKwd(linkText), getSvc(linkText), findTrackLink(Me.browser.Document.body.innerText)
        
        'and reset
        If stopNow Then
            GoTo done
        End If
        Me.browser.GoBack
        waitForPage CREATE_PAGE_TIMEOUT
    Next i
done:
    requeryNeedTL
    Mouse.Hourglass False
End Sub

Private Sub btnDeleteTL_Click()
    Mouse.Hourglass True
    stopNow = False
    complete = False
    
    Me.browser.Navigate2 START_PAGE
    waitForPage
    Me.browser.Navigate2 findLink("Create Links")
    waitForPage CREATE_PAGE_TIMEOUT
    
    Me.deleteTL.ListIndex = -1
    Dim i As Long
    Dim kwd As String, link As String
    For i = 0 To Me.deleteTL.ListCount - 1
        If stopNow Then
            Exit For
        End If
        
        kwd = Me.deleteTL.list(i)
        link = findLink(kwd)
        
        If link = "" Then
            ClearTrackLink cleanKwd(kwd), getSvc(kwd)
        Else
            Me.browser.Navigate2 link
            waitForPage
            Me.browser.Navigate2 findLink("Delete this link")
            waitForPage CREATE_PAGE_TIMEOUT
            ClearTrackLink cleanKwd(kwd), getSvc(kwd)
        End If
    Next i
    Mouse.Hourglass False
    requeryDeleteTL
End Sub

Private Sub btnExit_Click()
    Unload KeywordsTrackLinkTool
End Sub

Private Sub btnRefreshTL_Click()
    If MsgBox("Are you sure you want to do this? It will delete all tracklinks in the db, and replace them with what's up at yahoo.", vbYesNo) = vbNo Then
        Exit Sub
    End If
    Mouse.Hourglass True
    complete = False
    stopNow = False
    
    Me.browser.Navigate2 START_PAGE
    waitForPage
    Me.browser.Navigate2 findLink("Create Links")
    waitForPage CREATE_PAGE_TIMEOUT
    
    Dim i As Long
    Dim href_for As Dictionary
    Set href_for = New Dictionary
    Dim linkText As String, linkHref As String
    Dim to_delete As PerlArray
    Set to_delete = New PerlArray
    'add each keyword link in the page to a hash, or an array to be deleted
    For i = 0 To Me.browser.Document.Links.Length - 1
        linkText = Me.browser.Document.Links(i).innerText
        If hasPrefix(linkText) Then
            linkHref = Me.browser.Document.Links(i).href
            If Not href_for.exists(linkText) Then
                href_for.Add linkText, linkHref
            Else
                to_delete.Push linkHref
            End If
        End If
    Next i
    
    dbExecute "UPDATE KeywordsAttributes SET TrackLink=NULL"
    
    'process each item in the hash: go to page, find tl, save
    For i = 0 To href_for.count - 1
        If stopNow Then
            Exit For
        End If
        linkText = href_for.Keys(i)
        linkHref = href_for(linkText)
        Me.browser.Navigate2 linkHref
        waitForPage
        SetTrackLink cleanKwd(linkText), getSvc(linkText), findTrackLink(Me.browser.Document.body.innerText)
    Next i
    
    'process array: go to page, click delete, wait for reload
    For i = 0 To to_delete.Scalar - 1
        If stopNow Then
            Exit For
        End If
        Me.browser.Navigate2 to_delete.Elem(i)
        waitForPage
        Me.browser.Navigate2 findLink("Delete this link")
        waitForPage CREATE_PAGE_TIMEOUT
    Next i

    Mouse.Hourglass False
End Sub

Private Sub btnStop_Click()
    stopNow = True
End Sub

Private Sub Form_Load()
    Me.browser.Navigate2 START_PAGE
    requeryNeedTL
    requeryDeleteTL
End Sub

Private Sub requeryNeedTL()
    Dim rst As ADODB.Recordset
    Set rst = retrieve("SELECT SearchPhrase, Service, IsItemKeyword FROM vKeywordsTrackLinks WHERE Deleted=0 AND (TrackLink IS NULL OR TrackLink='') ORDER BY SearchPhrase, Service")
    Me.needTL.Clear
    While Not rst.EOF
        Me.needTL.AddItem prefix(rst("Service"), rst("IsItemKeyword")) & " " & rst("SearchPhrase")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryDeleteTL()
    Dim rst As ADODB.Recordset
    Set rst = retrieve("SELECT SearchPhrase, Service, IsItemKeyword FROM vKeywordsTrackLinks WHERE Deleted=1 AND TrackLink IS NOT NULL ORDER BY SearchPhrase, Service")
    Me.deleteTL.Clear
    While Not rst.EOF
        Me.deleteTL.AddItem prefix(rst("Service"), rst("IsItemKeyword")) & " " & rst("SearchPhrase")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Function waitForPage(Optional timeout As Long = 60) As Boolean
    Mouse.Hourglass True
    Dim starttime As Double
    starttime = Timer
    While Not complete
        DoEvents
        If Timer - starttime > timeout Then
            Me.browser.Stop
            complete = False
            waitForPage = False
            Mouse.Hourglass False
            Exit Function
        End If
    Wend
    complete = False
    waitForPage = True
    Mouse.Hourglass False
End Function

Private Function findLink(linkText As String) As String
    findLink = ""
    Dim i As Long
    For i = 0 To Me.browser.Document.Links.Length - 1
        If LCase(Me.browser.Document.Links(i).innerText) = LCase(linkText) Then
            findLink = Me.browser.Document.Links(i).href
            Exit For
        End If
    Next i
End Function

Private Function clickSubmitButton()
    Dim i As Long
    With Me.browser.Document.Forms(0)
        For i = 0 To .ALL.Length
            If UCase(.ALL(i).nodeName) = "INPUT" Then
                If UCase(.ALL(i).Type) = "SUBMIT" Then
                    Debug.Print complete
                    .ALL(i).Click
                    waitForPage
                    Debug.Print complete
                    Exit For
                End If
            End If
        Next i
    End With
End Function

Private Function prefix(svc As String, itemkwd As Boolean) As String
    If itemkwd Then
        Select Case svc
            Case Is = "Overture"
                prefix = YAHOO_ITEM_PREFIX
            Case Is = "AdWords"
                prefix = GOOGLE_ITEM_PREFIX
        End Select
    Else
        Select Case svc
            Case Is = "Overture"
                prefix = YAHOO_PHRASE_PREFIX
            Case Is = "AdWords"
                prefix = GOOGLE_PHRASE_PREFIX
        End Select
    End If
End Function

Private Function hasPrefix(linkText As String) As Boolean
    Dim test As String
    test = Left(linkText, 5)
    If test = YAHOO_ITEM_PREFIX & " " Or test = GOOGLE_ITEM_PREFIX & " " Or test = GOOGLE_PHRASE_PREFIX & " " Then
        hasPrefix = True
    Else
        test = Left(linkText, 4)
        If test = YAHOO_PHRASE_PREFIX & " " Then
            hasPrefix = True
        Else
            hasPrefix = False
        End If
    End If
End Function

Private Function cleanKwd(txt) As String
    Dim test As String
    test = Left(txt, 5)
    If test = GOOGLE_PHRASE_PREFIX & " " Or test = GOOGLE_ITEM_PREFIX & " " Or test = YAHOO_ITEM_PREFIX & " " Then
        cleanKwd = Mid(txt, 6)
    Else
        test = Left(txt, 4)
        If test = YAHOO_PHRASE_PREFIX & " " Then
            cleanKwd = Mid(txt, 5)
        Else
            cleanKwd = ""
        End If
    End If
End Function

Private Function getSvc(txt) As String
    Dim test As String
    test = Left(txt, 5)
    If test = GOOGLE_PHRASE_PREFIX & " " Or test = GOOGLE_ITEM_PREFIX & " " Then
        getSvc = "AdWords"
    ElseIf test = YAHOO_ITEM_PREFIX & " " Then
        getSvc = "Overture"
    Else
        test = Left(txt, 4)
        If test = YAHOO_PHRASE_PREFIX & " " Then
            getSvc = "Overture"
        Else
            getSvc = ""
        End If
    End If
End Function

Private Function findTrackLink(txt As String) As String
    findTrackLink = Mid(txt, InStr(txt, "?toolsplus+") + 11, 6)
End Function
