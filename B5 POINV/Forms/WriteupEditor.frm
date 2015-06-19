VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form WriteupEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Write-Up Editor"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPreview 
      Caption         =   "Flip to Preview"
      Height          =   345
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1230
      Width           =   1485
   End
   Begin VB.CommandButton btnShowBrowser 
      Caption         =   "Browser <<<"
      Height          =   315
      Left            =   10170
      TabIndex        =   19
      Top             =   4920
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser PageBrowser 
      Height          =   4005
      Left            =   90
      TabIndex        =   18
      Top             =   5310
      Width           =   11235
      ExtentX         =   19817
      ExtentY         =   7064
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
   Begin VB.CommandButton btnCloseWriteup 
      Caption         =   "Close"
      Height          =   315
      Left            =   1890
      TabIndex        =   17
      Top             =   1230
      Width           =   1395
   End
   Begin VB.CommandButton btnApproveWriteup 
      Caption         =   "Approve"
      Height          =   465
      Left            =   10170
      TabIndex        =   16
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CheckBox chkNeedsApproval 
      Caption         =   "Ready For Approval"
      Height          =   465
      Left            =   10170
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton btnSaveWriteup 
      Caption         =   "Save Changes"
      Height          =   465
      Left            =   10170
      TabIndex        =   14
      Top             =   2670
      Width           =   1215
   End
   Begin VB.CommandButton btnEditWriteup 
      Caption         =   "Edit This Write-up"
      Height          =   465
      Left            =   10170
      TabIndex        =   13
      Top             =   1110
      Width           =   1215
   End
   Begin VB.TextBox WriteupText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1590
      Width           =   10035
   End
   Begin SHDocVwCtl.WebBrowser WriteupPreview 
      Height          =   3615
      Left            =   60
      TabIndex        =   12
      Top             =   1590
      Width           =   10035
      ExtentX         =   17701
      ExtentY         =   6376
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
   Begin VB.CommandButton btnAddHTMLLink 
      Caption         =   "Add Link"
      Height          =   465
      Left            =   10170
      TabIndex        =   11
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   495
      Left            =   10170
      TabIndex        =   10
      Top             =   30
      Width           =   1215
   End
   Begin VB.Frame FilterFrame 
      Caption         =   "Filter:"
      Height          =   1515
      Left            =   4710
      TabIndex        =   3
      Top             =   30
      Width           =   3165
      Begin VB.CheckBox chkFilterItemsNeedApproval 
         Caption         =   "Show Items That Need Approval"
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2955
      End
      Begin VB.CheckBox chkFilterItemsMissingWriteups 
         Caption         =   "Show Items w/o Write-ups"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2955
      End
      Begin VB.CheckBox chkFilterItemsBeingEditedByOtherUser 
         Caption         =   "Show Items Being Edited By Others"
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2955
      End
      Begin VB.CheckBox chkFilterItemsBeingEditedByCurrentUser 
         Caption         =   "Show Items Being Edited By Me"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Value           =   1  'Checked
         Width           =   2955
      End
      Begin VB.CheckBox chkFilterItemsWithWriteups 
         Caption         =   "Show Items w/ Write-ups"
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.ComboBox ItemNumber 
      Height          =   315
      Left            =   570
      TabIndex        =   2
      Top             =   840
      Width           =   4005
   End
   Begin VB.Label Label2 
      Caption         =   "Item:"
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   870
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Write-Up Editor"
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
      Top             =   90
      Width           =   2205
   End
End
Attribute VB_Name = "WriteupEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fillingForm As Boolean
Private originalText As String

Private Sub btnAddHTMLLink_Click()
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

Private Sub btnApproveWriteup_Click()
    'TODO
End Sub

Private Sub btnCloseWriteup_Click()
    If originalText <> Me.WriteupText Then
        If vbNo = MsgBox("Cancel changes", vbYesNo + vbDefaultButton2) Then
            Exit Sub
        End If
    End If
    
    fillingForm = True
    enableBody False
    enableHeader True
    fillingForm = False
End Sub

Private Sub btnEditWriteup_Click()
    'this button should only be enabled when there is an existing writeup in signs
    If vbYes = MsgBox("Start a new writeup for this item?", vbYesNo + vbDefaultButton2) Then
        fillingForm = True
        originalText = ""
        Me.btnEditWriteup.Visible = False
        Me.btnEditWriteup.Enabled = False
        Me.WriteupText = ""
        Me.WriteupText.Enabled = True
        Me.btnAddHTMLLink.Enabled = True
        Me.btnSaveWriteup.Enabled = True
        Me.chkNeedsApproval = 0
        Me.chkNeedsApproval.Enabled = True
        Me.btnApproveWriteup.Enabled = HasPermissionsTo("WriteupSupervisor")
        fillingForm = False
    End If
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnSaveWriteup_Click()
    If originalText = Me.WriteupText Then
        'do nothing
    Else
        saveWriteup
    End If
End Sub

Private Sub btnShowBrowser_Click()
    If Me.PageBrowser.Visible Then
        Me.PageBrowser.Visible = False
        Me.Height = 5655
        Me.btnShowBrowser.caption = "Browser >>>"
    Else
        Me.PageBrowser.Visible = True
        Me.Height = 9780
        Me.btnShowBrowser.caption = "Browser <<<"
    End If
End Sub

Private Sub chkNeedsApproval_Click()
    If Not fillingForm Then
        Dim continue As Boolean
        continue = False
        If Me.chkNeedsApproval Then
            If originalText <> Me.WriteupText Then
                If saveWriteup() Then
                    continue = True
                Else
                    continue = False
                End If
            Else
                continue = True
            End If
        Else
            continue = True
        End If
        
        If continue Then
            DB.Execute "UPDATE WriteupsInProgress SET NeedsApproval=" & SQLBool(Me.chkNeedsApproval) & " WHERE ItemNumber='" & getCurrentItem() & "'"
        End If
        Me.WriteupText.Enabled = CBool(Me.chkNeedsApproval)
        Me.btnAddHTMLLink.Enabled = CBool(Me.chkNeedsApproval)
        Me.btnSaveWriteup.Enabled = False
    End If
End Sub

Private Sub chkPreview_Click()
    If Me.chkPreview Then
        Open DESTINATION_DIR & "writeuppreview.html" For Output As #1
        Print #1, "<html>" & vbCrLf & _
                  " <head>" & vbCrLf & _
                  "  <base href=""http://www.tools-plus.com/"" target=""_blank"">" & vbCrLf & _
                  "  <style type=""text/css"">" & vbCrLf & _
                  "   body { font-family: Arial,Helvetica,sans-serif; font-size: 12px; }" & vbCrLf & _
                  "  </style>" & vbCrLf & _
                  " </head>" & vbCrLf & _
                  " <body>"
        Dim paras As Variant
        paras = Split(Me.WriteupText, vbCrLf & vbCrLf)
        Dim i As Long
        For i = 0 To UBound(paras)
            Print #1, "  <p>" & paras(i) & "</p>"
        Next i
        Print #1, " </body>" & vbCrLf & "</html>"
        Close #1
        Me.WriteupPreview.Navigate2 DESTINATION_DIR & "writeuppreview.html"
        Me.WriteupText.Visible = False
        Me.WriteupPreview.Visible = True
        Me.btnAddHTMLLink.Enabled = False
    Else
        Me.WriteupText.Visible = True
        Me.WriteupPreview.Visible = False
        If Me.WriteupText.Enabled Then
            Me.btnAddHTMLLink.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.btnApproveWriteup.Visible = HasPermissionsTo("WriteupSupervisor")
    enableHeader True
    enableBody False
    requeryItemList
End Sub

Private Sub ItemNumber_Click()
    ItemNumber_LostFocus
End Sub

Private Sub ItemNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.ItemNumber, KeyCode, Shift
End Sub

Private Sub ItemNumber_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.ItemNumber, KeyAscii
    Select Case KeyAscii
        Case Is = vbKeyReturn
            ItemNumber_LostFocus
    End Select
End Sub

Private Sub ItemNumber_LostFocus()
    AutoCompleteLostFocus Me.ItemNumber
    If Me.ItemNumber = "" Then
        enableHeader True
        enableBody False
    Else
        enableHeader False
        loadWriteup getCurrentItem()
    End If
End Sub

Private Sub chkFilterItemsBeingEditedByCurrentUser_Click()
    requeryItemList
End Sub

Private Sub chkFilterItemsBeingEditedByOtherUser_Click()
    requeryItemList
End Sub

Private Sub chkFilterItemsMissingWriteups_Click()
    requeryItemList
End Sub

Private Sub chkFilterItemsNeedApproval_Click()
    requeryItemList
End Sub

Private Sub chkFilterItemsWithWriteups_Click()
    requeryItemList
End Sub

Private Sub requeryItemList()
    Dim conditions As PerlArray
    Set conditions = New PerlArray
    If Me.chkFilterItemsWithWriteups Then
        conditions.Push "(HasWriteup=1)"
    End If
    If Me.chkFilterItemsNeedApproval Then
        conditions.Push "(HasEdit=1 AND NeedsApproval=1)"
    End If
    If Me.chkFilterItemsBeingEditedByCurrentUser Then
        conditions.Push "(HasEdit=1 AND EditorID=" & GetCurrentUserID() & ")"
    End If
    If Me.chkFilterItemsBeingEditedByOtherUser Then
        conditions.Push "(HasEdit=1 AND EditorID<>" & GetCurrentUserID() & ")"
    End If
    If Me.chkFilterItemsMissingWriteups Then
        conditions.Push "(HasWriteup=0 AND HasEdit=0)"
    End If
    
    If conditions.Scalar = 0 Then
        MsgBox "Select something in the filter box!"
    Else
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT ItemNumber, HasWriteup, HasEdit, EditorName FROM vWriteups WHERE " & conditions.Join(" OR "))
        Me.ItemNumber.Clear
        While Not rst.EOF
            Dim status As String
            If rst("HasWriteup") And rst("HasEdit") Then
                status = "has write-up, being edited by " & rst("EditorName")
            ElseIf rst("HasWriteup") Then
                status = "has write-up"
            ElseIf rst("HasEdit") Then
                status = "being edited by " & rst("EditorName")
            Else
                status = "no write-up"
            End If
            Me.ItemNumber.AddItem rst("ItemNumber") & " - " & status
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
        ExpandDropDownToFit Me.ItemNumber
    End If
End Sub

Private Sub loadWriteup(item As String)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT HasEdit, CurrentWriteup, NewWriteup, EditorID, NeedsApproval, YahooID FROM vWriteups WHERE ItemNumber='" & item & "'")
    Me.PageBrowser.Navigate2 "http://www.tools-plus.com/" & Nz(rst("YahooID")) & ".html"
    Me.chkPreview = 0
    Me.chkPreview.Enabled = True
    fillingForm = True
    If rst("HasEdit") And (rst("EditorID") = GetCurrentUserID Or HasPermissionsTo("WriteupSupervisor")) Then
        'in-progress writeup, and editable
        originalText = Nz(rst("NewWriteup"))
        Me.btnEditWriteup.Visible = False
        Me.btnEditWriteup.Enabled = False
        Me.WriteupText = Nz(rst("NewWriteup"))
        Me.WriteupText.Enabled = Not rst("NeedsApproval")
        Me.btnAddHTMLLink.Enabled = Not rst("NeedsApproval")
        Me.btnSaveWriteup.Enabled = Not rst("NeedsApproval")
        Me.chkNeedsApproval = SQLBool(rst("NeedsApproval"))
        Me.chkNeedsApproval.Enabled = True
        Me.btnApproveWriteup.Enabled = HasPermissionsTo("WriteupSupervisor")
    ElseIf rst("HasEdit") Then
        'in-progress writeup, but not editable
        originalText = Nz(rst("NewWriteup"))
        Me.btnEditWriteup.Visible = False
        Me.btnEditWriteup.Enabled = False
        Me.WriteupText = Nz(rst("NewWriteup"))
        Me.WriteupText.Enabled = False
        Me.btnAddHTMLLink.Enabled = False
        Me.btnSaveWriteup.Enabled = False
        Me.chkNeedsApproval = SQLBool(rst("NeedsApproval"))
        Me.chkNeedsApproval.Enabled = False
        Me.btnApproveWriteup.Enabled = False
    ElseIf Nz(rst("CurrentWriteup")) = "" Then
        'no writeup
        originalText = ""
        Me.btnEditWriteup.Visible = False
        Me.btnEditWriteup.Enabled = False
        Me.WriteupText = ""
        Me.WriteupText.Enabled = True
        Me.btnAddHTMLLink.Enabled = True
        Me.btnSaveWriteup.Enabled = True
        Me.chkNeedsApproval = 0
        Me.chkNeedsApproval.Enabled = True
        Me.btnApproveWriteup.Enabled = HasPermissionsTo("WriteupSupervisor")
    Else
        'published writeup, not editable
        originalText = Nz(rst("CurrentWriteup"))
        Me.btnEditWriteup.Visible = True
        Me.btnEditWriteup.Enabled = True
        Me.WriteupText = Nz(rst("CurrentWriteup"))
        Me.WriteupText.Enabled = False
        Me.btnAddHTMLLink.Enabled = False
        Me.btnSaveWriteup.Enabled = False
        Me.chkNeedsApproval = 0
        Me.chkNeedsApproval.Enabled = False
        Me.btnApproveWriteup.Enabled = False
    End If
    Me.btnCloseWriteup.Enabled = True
    fillingForm = False
End Sub

Private Sub enableHeader(onoff As Boolean)
    Me.ItemNumber.Enabled = onoff
    Me.FilterFrame.Enabled = onoff
End Sub

Private Sub enableBody(onoff As Boolean)
    If onoff Then
        Debug.Assert "on is not supported!"
    Else
        If Me.chkPreview Then
            Me.chkPreview = 0
        End If
        If Not onoff Then
            Me.WriteupText = ""
        End If
        Me.WriteupText.Enabled = False
        Me.chkPreview.Enabled = False
        Me.btnEditWriteup.Enabled = False
        Me.btnAddHTMLLink.Enabled = False
        Me.btnSaveWriteup.Enabled = False
        Me.btnCloseWriteup.Enabled = False
        Me.chkNeedsApproval.Enabled = False
        Me.btnApproveWriteup.Enabled = False
    End If
End Sub

Private Sub addTag(openTag As String, Optional closeTag As String = "")
    If closeTag = "" Then
        closeTag = Replace(openTag, "<", "</")
    End If
    
    Dim selStart As Long, selLength As Long, selEnd As Long
    selStart = Me.WriteupText.selStart
    selLength = Me.WriteupText.selLength
    selEnd = selStart + selLength
    
    Me.WriteupText = Left(Me.WriteupText, selStart) & _
                     openTag & _
                     Mid(Me.WriteupText, selStart + 1, selLength) & _
                     closeTag & _
                     Mid(Me.WriteupText, selEnd + 1)
    
    Me.WriteupText.selStart = selStart + Len(openTag)
    If selLength = 0 Then
        Me.WriteupText.selLength = 0
    Else
        Me.WriteupText.selLength = selLength
    End If
    Me.WriteupText.SetFocus
End Sub

Private Function getCurrentItem() As String
    getCurrentItem = Left(Me.ItemNumber, InStr(Me.ItemNumber, " ") - 1)
End Function

Private Function saveWriteup() As Boolean
    If vbYes = MsgBox("Save changes?", vbYesNo) Then
        Dim item As String
        item = getCurrentItem()
        If 0 = DLookup("COUNT(*)", "WriteupsInProgress", "ItemNumber='" & item & "'") Then
            DB.Execute "INSERT INTO WriteupsInProgress ( ItemNumber, WriteupHTML, UserID ) VALUES ( '" & item & "', '" & EscapeSQuotes(Me.WriteupText) & "', " & GetCurrentUserID & " )"
            'TODO: change combobox to new status?
            Dim newstatus As String
            Dim rst As ADODB.Recordset
            Set rst = DB.retrieve("SELECT CASE WHEN WriteupHTML IS NULL OR DATALENGTH(WriteupHTML)=0 THEN 0 ELSE 1 END FROM PartNumbers WHERE ItemNumber='" & item & "'")
            If rst(0) Then
                newstatus = "has write-up, being edited by " & GetCurrentUserNickname()
            Else
                newstatus = "being edited by " & GetCurrentUserNickname()
            End If
            rst.Close
            Set rst = Nothing
            newstatus = getCurrentItem & " - " & newstatus
            Me.ItemNumber.list(Me.ItemNumber.ListIndex) = newstatus
            'Me.ItemNumber = newstatus
        Else
            DB.Execute "UPDATE WriteupsInProgress SET WriteupHTML='" & EscapeSQuotes(Me.WriteupText) & "' WHERE ItemNumber='" & item & "'"
        End If
        originalText = Me.WriteupText
        saveWriteup = True
    Else
        saveWriteup = False
    End If
End Function
