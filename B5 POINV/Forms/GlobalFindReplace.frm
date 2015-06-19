VERSION 5.00
Begin VB.Form GlobalFindReplace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find & Replace"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7935
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkCaseSensitive 
      Caption         =   "Case sensitive?"
      Height          =   255
      Left            =   270
      TabIndex        =   17
      Top             =   1350
      Width           =   1575
   End
   Begin VB.TextBox FilterText 
      Height          =   285
      Left            =   240
      TabIndex        =   16
      Top             =   2910
      Width           =   3105
   End
   Begin VB.CheckBox chkPromptFirst 
      Caption         =   "Prompt before replace?"
      Height          =   255
      Left            =   270
      TabIndex        =   13
      Top             =   2280
      Value           =   1  'Checked
      Width           =   2085
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   3420
      Width           =   1245
   End
   Begin VB.CommandButton btnGo 
      Caption         =   "GO!!!"
      Height          =   495
      Left            =   450
      TabIndex        =   11
      Top             =   3390
      Width           =   1275
   End
   Begin VB.TextBox ReplaceText 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   1950
      Width           =   3105
   End
   Begin VB.CheckBox chkIsRegex 
      Caption         =   "Regex search?"
      Height          =   255
      Left            =   270
      TabIndex        =   7
      Top             =   1080
      Width           =   1995
   End
   Begin VB.TextBox SearchText 
      Height          =   285
      Left            =   210
      TabIndex        =   6
      Top             =   750
      Width           =   3165
   End
   Begin VB.CommandButton btnMoveRight 
      Caption         =   "-->"
      Height          =   405
      Left            =   5820
      TabIndex        =   5
      Top             =   1500
      Width           =   435
   End
   Begin VB.CommandButton btnMoveLeft 
      Caption         =   "<--"
      Height          =   405
      Left            =   5820
      TabIndex        =   4
      Top             =   1050
      Width           =   435
   End
   Begin VB.ListBox AllFields 
      Height          =   2790
      Left            =   6270
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   330
      Width           =   1635
   End
   Begin VB.ListBox FieldList 
      Height          =   2790
      Left            =   4140
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   330
      Width           =   1635
   End
   Begin VB.Label lblStatusBar 
      Height          =   255
      Left            =   4110
      TabIndex        =   18
      Top             =   3780
      Width           =   3705
   End
   Begin VB.Label Label6 
      Caption         =   "Filter:"
      Height          =   225
      Left            =   270
      TabIndex        =   15
      Top             =   2670
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Find && Replace"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   210
      TabIndex        =   14
      Top             =   90
      Width           =   2865
   End
   Begin VB.Label Label4 
      Caption         =   "Replace With:"
      Height          =   225
      Left            =   270
      TabIndex        =   9
      Top             =   1710
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Search For:"
      Height          =   225
      Left            =   270
      TabIndex        =   8
      Top             =   510
      Width           =   1485
   End
   Begin VB.Label Label2 
      Caption         =   "Available fields:"
      Height          =   225
      Left            =   6330
      TabIndex        =   3
      Top             =   90
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Fields to search:"
      Height          =   225
      Left            =   4200
      TabIndex        =   2
      Top             =   90
      Width           =   1485
   End
End
Attribute VB_Name = "GlobalFindReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private stopNow As Boolean

Private Sub btnExit_Click()
    Unload GlobalFindReplace
End Sub

Private Sub btnGo_Click()
    If Me.FieldList.ListCount = 0 Then
        MsgBox "No fields selected!"
        Exit Sub
    End If
    Mouse.Hourglass True
    stopNow = False
    Dim rst As ADODB.Recordset
    Me.lblStatusBar.caption = "Loading records..."
    On Error GoTo errh
    Set rst = DB.retrieve("SELECT PartNumbers.ItemNumber, " & Join(ListBoxAsArray(Me.FieldList), ", ") & " FROM PartNumbers" & IIf(Me.FilterText <> "", " WHERE " & Me.FilterText, ""))
    On Error GoTo 0
    
    Dim i As Long, found As Boolean
    Dim totalchecks As Long, totalchanges As Long
    While Not rst.EOF
        Me.lblStatusBar.caption = "Working on " & rst("ItemNumber")
        For i = 1 To rst.fields.count - 1
            found = False
            Dim re As RegExp
            If Me.chkIsRegex Then
                Set re = New RegExp
                re.Pattern = Me.SearchText
                re.IgnoreCase = Not CBool(Me.chkCaseSensitive)
                re.Global = True
                If re.Test(Nz(rst(i))) Then
                    found = True
                End If
            Else
                If InStr(1, Nz(rst(i)), Me.SearchText, IIf(Me.chkCaseSensitive, vbBinaryCompare, vbTextCompare)) > 0 Then
                    found = True
                End If
            End If
            
            If found Then
                'If Me.chkPromptFirst Then
                '    MsgBox "Item " & rst("ItemNumber") & ": replacing field " & rst.fields(i).name
                'End If
                Dim newText As String
                If Me.chkIsRegex Then
                    Dim m As MatchCollection
                    Set m = re.Execute(rst(i))
                    Dim j As Long, realReplText As String
                    newText = rst(i)
                    Dim k As Long
                    For j = 0 To m.count - 1
                        realReplText = Me.ReplaceText
                        For k = 0 To m.item(j).SubMatches.count - 1
                            realReplText = Replace(realReplText, "\" & k + 1, m.item(j).SubMatches(k))
                        Next k
                        newText = Replace(newText, m.item(j).value, realReplText, 1, 1)
                        'replaceThis rst.fields(i).name, newText, rst("ItemNumber")
                    Next j
                Else
                    newText = Replace(rst(i), Me.SearchText, Me.ReplaceText, 1, -1, IIf(Me.chkCaseSensitive, vbBinaryCompare, vbTextCompare))
                End If
                If rst(i).value <> newText Then
                    If replaceThis(rst.fields(i).name, newText, rst("ItemNumber"), rst(i)) Then
                        totalchanges = totalchanges + 1
                    End If
                End If
            End If
            Set re = Nothing
            totalchecks = totalchecks + 1
            If stopNow Then
                i = rst.fields.count
            End If
        Next i
        If stopNow Then
            rst.MoveLast
        End If
        rst.MoveNext
        DoEvents
    Wend
    rst.Close
    Set rst = Nothing
    Me.lblStatusBar.caption = ""
    Mouse.Hourglass False
    If stopNow Then
        MsgBox "Search cancelled!"
    Else
        MsgBox "Finished search, made " & totalchanges & " replacements out of " & totalchecks & " fields searched."
    End If
    Exit Sub
errh:
    If Me.FilterText = "" Then
        MsgBox "Invalid SQL, this is probably a bug" & vbCrLf & vbCrLf & "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    Else
        MsgBox "Invalid SQL, this is either a bug, or a problem in your filter string" & vbCrLf & vbCrLf & "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    End If
    Mouse.Hourglass False
    Me.lblStatusBar.caption = ""
End Sub

Private Function replaceThis(fieldName As String, newText As String, item As String, Optional origText As String) As Boolean
    Dim doIt As Boolean
    doIt = True
    stopNow = False
    If Me.chkPromptFirst Then
        If origText = "" Then
            origText = DLookup(fieldName, "PartNumber" & IIf(Right(fieldName, 4) = "HTML", "s", "CleanedSpecs"), "ItemNumber='" & item & "'")
        End If
        Select Case MsgBox("Replace this? (" & item & ", " & fieldName & ")" & vbCrLf & vbCrLf & "old: " & origText & vbCrLf & vbCrLf & "new: " & newText, vbYesNoCancel)
            Case Is = vbYes
                doIt = True
            Case Is = vbNo
                doIt = False
            Case Is = vbCancel
                stopNow = True
        End Select
    End If
    
    If doIt And Not stopNow Then
        DB.Execute "UPDATE PartNumbers SET " & fieldName & "='" & EscapeSQuotes(newText) & "' WHERE ItemNumber='" & item & "'"
        replaceThis = True
    Else
        replaceThis = False
    End If
End Function

Private Sub btnMoveLeft_Click()
    If Me.AllFields.SelCount <> 0 Then
        Dim i As Long
        For i = Me.AllFields.ListCount - 1 To 0 Step -1
            If Me.AllFields.Selected(i) Then
                Me.FieldList.AddItem Me.AllFields.list(i)
                Me.AllFields.RemoveItem i
            End If
        Next i
    End If
End Sub

Private Sub btnMoveRight_Click()
    If Me.FieldList.SelCount <> 0 Then
        Dim i As Long
        For i = Me.FieldList.ListCount - 1 To 0 Step -1
            If Me.FieldList.Selected(i) Then
                Me.AllFields.AddItem Me.FieldList.list(i)
                Me.FieldList.RemoveItem i
            End If
        Next i
    End If
End Sub

Private Sub AllFields_DblClick()
    btnMoveLeft_Click
End Sub

Private Sub FieldList_DblClick()
    btnMoveRight_Click
End Sub

Private Sub Form_Load()
    requeryAllFields
End Sub

Private Sub requeryAllFields()
    Me.AllFields.AddItem "Spec1HTML"
    Me.AllFields.AddItem "Spec2HTML"
    Me.AllFields.AddItem "Spec3HTML"
    Me.AllFields.AddItem "Spec4HTML"
    Me.AllFields.AddItem "Spec5HTML"
    Me.AllFields.AddItem "Spec6HTML"
    Me.AllFields.AddItem "Spec7HTML"
    Me.AllFields.AddItem "Spec8HTML"
    Me.AllFields.AddItem "WriteupHTML"
    Me.AllFields.AddItem "NotesHTML"
    Me.AllFields.AddItem "FeaturesHTML"
    Me.AllFields.AddItem "TechSpecsHTML"
    Me.AllFields.AddItem "MediaHTML"
    Me.AllFields.AddItem "ExtendedSpecsHTML"
    Me.AllFields.AddItem "Desc1"
    Me.AllFields.AddItem "Desc3"
    Me.AllFields.AddItem "Desc4"
End Sub
