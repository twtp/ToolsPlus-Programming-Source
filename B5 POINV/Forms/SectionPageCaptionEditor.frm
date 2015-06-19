VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form SectionPageCaptionEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Section Page Editor"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   12405
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnGenerateCSV 
      Caption         =   "CSV"
      Height          =   495
      Left            =   10110
      TabIndex        =   20
      Top             =   30
      Width           =   705
   End
   Begin VB.CommandButton btnHTMLBold 
      Caption         =   "Bold"
      Height          =   435
      Left            =   11670
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2430
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton btnHTMLLink 
      Caption         =   "Link"
      Height          =   435
      Left            =   11670
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1950
      Width           =   675
   End
   Begin VB.CheckBox FlipToPreview 
      Caption         =   "Flip to Preview"
      Enabled         =   0   'False
      Height          =   315
      Left            =   10860
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1110
      Width           =   1455
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   495
      Left            =   11220
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   30
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      Caption         =   "Set Status:"
      Height          =   1395
      Left            =   7380
      TabIndex        =   10
      Top             =   30
      Width           =   2685
      Begin VB.OptionButton opStatus 
         Caption         =   "Mark as Complete"
         Height          =   225
         Index           =   3
         Left            =   180
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   2115
      End
      Begin VB.OptionButton opStatus 
         Caption         =   "Mark as Not Started"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1050
         Width           =   2115
      End
      Begin VB.OptionButton opStatus 
         Caption         =   "Mark as In-Progress"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   780
         Width           =   2115
      End
      Begin VB.OptionButton opStatus 
         Caption         =   "Mark as Needs Approval"
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   510
         Width           =   2115
      End
   End
   Begin SHDocVwCtl.WebBrowser PageBrowser 
      Height          =   5385
      Left            =   120
      TabIndex        =   8
      Top             =   4770
      Width           =   12165
      ExtentX         =   21458
      ExtentY         =   9499
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
   Begin VB.TextBox CaptionText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3285
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1440
      Width           =   11475
   End
   Begin VB.ComboBox PageSelector 
      Height          =   315
      Left            =   150
      TabIndex        =   5
      Top             =   870
      Width           =   3675
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter:"
      Height          =   1395
      Left            =   4080
      TabIndex        =   1
      Top             =   30
      Width           =   2865
      Begin VB.CheckBox chkFilter 
         Caption         =   "Show Complete"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox chkFilter 
         Caption         =   "Show Not Started"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1050
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkFilter 
         Caption         =   "Show In-Progress"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   780
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkFilter 
         Caption         =   "Show Needs Approval"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   510
         Width           =   2175
      End
   End
   Begin SHDocVwCtl.WebBrowser PreviewBrowser 
      Height          =   3285
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   11475
      ExtentX         =   20241
      ExtentY         =   5794
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
   Begin VB.Label Label2 
      Caption         =   "Select Page:"
      Height          =   225
      Left            =   270
      TabIndex        =   9
      Top             =   630
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Section Page Editor"
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
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2685
   End
End
Attribute VB_Name = "SectionPageCaptionEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'NOTE: if the status number for complete changes, then
'you also need to change it in the perl script!

Private changed As Boolean
Private fillingform As Boolean

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnGenerateCSV_Click()
    Dim dlg As FilePicker
    Set dlg = New FilePicker
    dlg.SetParent Me.hwnd
    dlg.AddFilter "Comma Separated Values Files", "*.csv"
    Dim fname As String
    fname = dlg.ShowDialogSave
    Set dlg = Nothing
    If fname <> "" Then
        If LCase(Right(fname, 4)) <> ".csv" Then
            fname = fname & ".csv"
        End If
        Shell PERL & " s:\mastest\mas90-signs\generate_sectionpage_contents_csv.pl " & qq(fname)
    End If
End Sub

Private Sub btnHTMLBold_Click()
    addTag "<strong>"
End Sub

Private Sub btnHTMLLink_Click()
    Dim url As String
    url = Me.PageBrowser.LocationURL
    url = Mid(url, InStrRev(url, "/") + 1)
    url = InputBox("Paste the URL of the page you want to link to:", "Link To", url)
    If url = "" Then
        'do nothing
    Else
        url = Mid(url, InStrRev(url, "/") + 1)
        Dim req As MSXML2.XMLHTTP
        Set req = New MSXML2.XMLHTTP
        req.Open "HEAD", "http://www.tools-plus.com/" & url
        req.Send
        While req.ReadyState <> 4
            'Debug.Print req.ReadyState
            Sleep 100
            DoEvents
        Wend
        If req.status <> 200 Then
            MsgBox "Couldn't find ""http://www.tools-plus.com/" & url & """ (error " & req.status & ")"
        Else
            addTag "<a href=""" & url & """>", "</a>"
        End If
    End If
End Sub

Private Sub CaptionText_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.CaptionText.selStart = 0
                Me.CaptionText.selLength = Len(Me.CaptionText)
                KeyCode = 0
                Shift = 0
            Else
                CaptionText_KeyPress KeyCode
            End If
        Case Is = vbKeyDelete
            CaptionText_KeyPress KeyCode
    End Select
End Sub

Private Sub CaptionText_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub CaptionText_LostFocus()
    If changed = True Then
        Me.CaptionText = UnFrontpage(Me.CaptionText)
        Me.CaptionText = TrimWhitespace(Me.CaptionText, True, True, True, False)
        Me.CaptionText = TrimWhitespace(Me.CaptionText, True, True, False, True)
        DB.Execute "UPDATE SectionPageText SET CaptionText='" & EscapeSQuotes(Me.CaptionText) & "' WHERE URL='" & Me.PageSelector & "'"
        If Me.opStatus(0) Then
            If Len(Me.CaptionText) > 0 Then
                Me.opStatus(1) = True
            End If
        End If
        changed = False
    End If
End Sub

Private Sub chkFilter_Click(Index As Integer)
    requeryPageList
End Sub

Private Sub FlipToPreview_Click()
    If Me.FlipToPreview Then
        Open DESTINATION_DIR & "secpagepreview.html" For Output As #1
        Print #1, "<html>" & vbCrLf & _
                  " <head>" & vbCrLf & _
                  "  <base href=""http://www.tools-plus.com/"" target=""_blank"">" & vbCrLf & _
                  "  <style type=""text/css"">" & vbCrLf & _
                  "   body { font-family: Arial,Helvetica,sans-serif; font-size: 12px; }" & vbCrLf & _
                  "  </style>" & vbCrLf & _
                  " </head>" & vbCrLf & _
                  " <body>"
        Dim paras As Variant
        paras = Split(Me.CaptionText, vbCrLf & vbCrLf)
        Dim i As Long
        For i = 0 To UBound(paras)
            Print #1, "  <p>" & paras(i) & "</p>"
        Next i
        Print #1, " </body>" & vbCrLf & "</html>"
        Close #1
        Me.PreviewBrowser.Navigate2 DESTINATION_DIR & "secpagepreview.html"
        Me.CaptionText.Visible = False
        Me.PreviewBrowser.Visible = True
        enableHTMLButtons False
    Else
        Me.CaptionText.Visible = True
        Me.PreviewBrowser.Visible = False
        enableHTMLButtons True
    End If
End Sub

Private Sub Form_Load()
    requeryPageList
End Sub

Private Sub requeryPageList()
    Dim rst As ADODB.Recordset
    Dim clauses As PerlArray
    Set clauses = New PerlArray
    Dim i As Long
    For i = Me.chkFilter.LBound To Me.chkFilter.UBound
        If Me.chkFilter(i) Then
            clauses.Push "Status=" & i
        End If
    Next i
    Dim whereClause As String
    If clauses.Scalar <> 0 Then
        whereClause = "WHERE " & clauses.Join(" OR ")
    Else
        whereClause = "WHERE 1=2" 'hmm
    End If
    Set rst = DB.retrieve("SELECT URL FROM SectionPageText " & whereClause & " ORDER BY Priority, URL")
    Me.PageSelector.Clear
    While Not rst.EOF
        Me.PageSelector.AddItem rst("URL")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    ExpandDropDownToFit Me.PageSelector
    
    Me.CaptionText = ""
    Me.FlipToPreview = 0
    Me.PageBrowser.Navigate2 "about:blank"
    Me.CaptionText.Enabled = False
    Me.FlipToPreview.Enabled = False
    enableHTMLButtons False
End Sub

Private Sub opStatus_Click(Index As Integer)
    If Not fillingform Then
        DB.Execute "UPDATE SectionPageText SET Status=" & Index & " WHERE URL='" & Me.PageSelector & "'"
    End If
End Sub

Private Sub PageSelector_Click()
    changed = True
    PageSelector_LostFocus
End Sub

Private Sub PageSelector_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.PageSelector, KeyCode, Shift
End Sub

Private Sub PageSelector_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.PageSelector, KeyAscii
    If KeyAscii = vbKeyReturn Then
        PageSelector_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub PageSelector_LostFocus()
    AutoCompleteLostFocus Me.PageSelector
    If changed Then
        fillingform = True
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT Status, CaptionText FROM SectionPageText WHERE URL='" & Me.PageSelector & "'")
        Me.CaptionText.Enabled = True
        Me.FlipToPreview.Enabled = True
        enableHTMLButtons True
        Me.CaptionText = Nz(rst("CaptionText"))
        Me.opStatus(rst("Status")) = 1
        Me.FlipToPreview = 0
        Me.PreviewBrowser.Navigate2 "about:blank"
        Me.PageBrowser.Navigate2 "http://www.tools-plus.com/" & Me.PageSelector & ".html"
        fillingform = False
        changed = False
    End If
End Sub

Private Sub enableHTMLButtons(val As Boolean)
    Me.btnHTMLLink.Enabled = val
    Me.btnHTMLBold.Enabled = val
End Sub

Private Sub addTag(openTag As String, Optional closeTag As String = "")
    If closeTag = "" Then
        closeTag = Replace(openTag, "<", "</")
    End If
    
    Dim selStart As Long, selLength As Long, selEnd As Long
    selStart = Me.CaptionText.selStart
    selLength = Me.CaptionText.selLength
    selEnd = selStart + selLength
    
    Me.CaptionText = Left(Me.CaptionText, selStart) & _
                     openTag & _
                     Mid(Me.CaptionText, selStart + 1, selLength) & _
                     closeTag & _
                     Mid(Me.CaptionText, selEnd + 1)
    
    Me.CaptionText.selStart = selStart + Len(openTag)
    If selLength = 0 Then
        Me.CaptionText.selLength = 0
    Else
        Me.CaptionText.selLength = selLength
    End If
    Me.CaptionText.SetFocus
End Sub
