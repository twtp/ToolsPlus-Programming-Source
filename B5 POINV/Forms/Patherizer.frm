VERSION 5.00
Begin VB.Form Patherizer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patherizer!"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnRenamePath 
      Caption         =   "Rename Path"
      Height          =   435
      Left            =   12000
      TabIndex        =   36
      Top             =   1500
      Width           =   1275
   End
   Begin VB.CommandButton btnDeletePath 
      Caption         =   "Delete Path"
      Height          =   435
      Left            =   12000
      TabIndex        =   35
      Top             =   870
      Width           =   1245
   End
   Begin VB.CommandButton btnUnpath 
      Caption         =   "<="
      Height          =   735
      Left            =   6360
      TabIndex        =   27
      Top             =   5160
      Width           =   645
   End
   Begin VB.CheckBox chkAllPaths 
      Caption         =   "All Paths"
      Height          =   285
      Left            =   4170
      TabIndex        =   26
      ToolTipText     =   "If checked, this will let you pick path2s directly, if unchecked, you are limited to valid paths w/ items"
      Top             =   1110
      Width           =   1365
   End
   Begin VB.CommandButton btnAddPath 
      Caption         =   "Add New Path"
      Height          =   495
      Left            =   11940
      TabIndex        =   25
      Top             =   180
      Width           =   1395
   End
   Begin VB.ListBox NewItemList 
      Height          =   4155
      Left            =   7110
      TabIndex        =   24
      Top             =   2430
      Width           =   6165
   End
   Begin VB.ComboBox NewWebPath 
      Height          =   315
      Index           =   6
      Left            =   7380
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   1890
      Width           =   3135
   End
   Begin VB.ComboBox NewWebPath 
      Height          =   315
      Index           =   5
      Left            =   7380
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   1530
      Width           =   3135
   End
   Begin VB.ComboBox NewWebPath 
      Height          =   315
      Index           =   4
      Left            =   7380
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   1170
      Width           =   3135
   End
   Begin VB.ComboBox NewWebPath 
      Height          =   315
      Index           =   3
      Left            =   7380
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   810
      Width           =   3135
   End
   Begin VB.ComboBox NewWebPath 
      Height          =   315
      Index           =   2
      Left            =   7380
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   450
      Width           =   3135
   End
   Begin VB.ComboBox NewWebPath 
      Height          =   315
      Index           =   1
      Left            =   7380
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   90
      Width           =   3135
   End
   Begin VB.CommandButton btnPathify 
      Caption         =   "=>"
      Height          =   735
      Left            =   6360
      TabIndex        =   11
      Top             =   3360
      Width           =   645
   End
   Begin VB.CheckBox chkShowMoved 
      Caption         =   "Show Moved"
      Height          =   285
      Left            =   4140
      TabIndex        =   10
      ToolTipText     =   "if checked, all items are displayed, otherwise, items that have been assigned a new path are not shown"
      Top             =   1650
      Value           =   1  'Checked
      Width           =   1305
   End
   Begin VB.ListBox ItemList 
      Height          =   4545
      Left            =   60
      MultiSelect     =   2  'Extended
      TabIndex        =   9
      Top             =   2040
      Width           =   6165
   End
   Begin VB.ComboBox WebPath 
      Height          =   315
      Index           =   4
      Left            =   660
      TabIndex        =   4
      Top             =   1620
      Width           =   3135
   End
   Begin VB.ComboBox WebPath 
      Height          =   315
      Index           =   3
      Left            =   660
      TabIndex        =   3
      Top             =   1260
      Width           =   3135
   End
   Begin VB.ComboBox WebPath 
      Height          =   315
      Index           =   2
      Left            =   660
      TabIndex        =   2
      Top             =   900
      Width           =   3135
   End
   Begin VB.ComboBox WebPath 
      Height          =   315
      Index           =   1
      Left            =   660
      TabIndex        =   1
      Top             =   540
      Width           =   3135
   End
   Begin VB.Label lblItemCount 
      Height          =   315
      Left            =   1470
      TabIndex        =   34
      Top             =   6690
      Width           =   3105
   End
   Begin VB.Label lblNewItemCount 
      Height          =   315
      Left            =   8280
      TabIndex        =   33
      Top             =   6690
      Width           =   3105
   End
   Begin VB.Label lblSubpathCount 
      Height          =   225
      Index           =   5
      Left            =   10560
      TabIndex        =   32
      Top             =   1590
      Width           =   1185
   End
   Begin VB.Label lblSubpathCount 
      Height          =   225
      Index           =   4
      Left            =   10560
      TabIndex        =   31
      Top             =   1230
      Width           =   1185
   End
   Begin VB.Label lblSubpathCount 
      Height          =   225
      Index           =   3
      Left            =   10560
      TabIndex        =   30
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label lblSubpathCount 
      Height          =   225
      Index           =   2
      Left            =   10560
      TabIndex        =   29
      Top             =   510
      Width           =   1185
   End
   Begin VB.Label lblSubpathCount 
      Height          =   225
      Index           =   1
      Left            =   10560
      TabIndex        =   28
      Top             =   150
      Width           =   1185
   End
   Begin VB.Label generalLabel 
      Caption         =   "Path 6"
      Height          =   255
      Index           =   10
      Left            =   6750
      TabIndex        =   23
      Top             =   1950
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Path 5"
      Height          =   255
      Index           =   9
      Left            =   6750
      TabIndex        =   22
      Top             =   1590
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Path 4"
      Height          =   255
      Index           =   8
      Left            =   6750
      TabIndex        =   21
      Top             =   1230
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Path 3"
      Height          =   255
      Index           =   7
      Left            =   6750
      TabIndex        =   20
      Top             =   870
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Path 2"
      Height          =   255
      Index           =   6
      Left            =   6750
      TabIndex        =   19
      Top             =   510
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Path 1"
      Height          =   255
      Index           =   5
      Left            =   6750
      TabIndex        =   18
      Top             =   150
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Path 4"
      Height          =   255
      Index           =   4
      Left            =   30
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Path 3"
      Height          =   255
      Index           =   3
      Left            =   30
      TabIndex        =   7
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Path 2"
      Height          =   255
      Index           =   2
      Left            =   30
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Path 1"
      Height          =   255
      Index           =   1
      Left            =   30
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Patherizer!"
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
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   1605
   End
End
Attribute VB_Name = "Patherizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private changed As Boolean
Private wpSel(1 To 4) As String
Private newWpSel(1 To 6) As String
Private newWpMapper(1 To 6) As Dictionary

Private Sub btnAddPath_Click()
    Dim upperpathstr As String, newpathlevel As Long, newpathname As String, parentid As String
    Dim i As Long
    For i = Me.NewWebPath.LBound To Me.NewWebPath.UBound
        If Me.NewWebPath(i) = "" Then
            newpathlevel = i
            Exit For
        End If
        upperpathstr = IIf(upperpathstr = "", "", upperpathstr & ": ") & Me.NewWebPath(i)
    Next i
    If newpathlevel = 1 Then
        parentid = "NULL"
    Else
        parentid = newWpMapper(newpathlevel - 1).item(Me.NewWebPath(newpathlevel - 1).ListIndex)
    End If
    newpathname = InputBox(upperpathstr, "Add new path " & newpathlevel)
    If newpathname <> "" Then
        DB.Execute "INSERT INTO WebPaths (WebPathName, PathLevel, ParentID) VALUES ( '" & EscapeSQuotes(newpathname) & "', " & newpathlevel & ", " & parentid & " )"
        requeryPathsNew newpathlevel
        'Dim newid As Long
        'newid = DLookup("@@IDENTITY", "WebPaths")
        'Me.NewWebPath(newpathlevel).AddItem newpathname
        'For i = 0 To Me.NewWebPath(newpathlevel).ListCount - 1
        '    If Me.NewWebPath(newpathlevel).list(i) = newpathname Then
        '        Dim j As Long
        '        For j = Me.NewWebPath(newpathlevel).ListCount - 1 To i + 1 Step -1
        '            newWpMapper(newpathlevel).item(j) = newWpMapper(newpathlevel).item(j - 1)
        '        Next j
        '        newWpMapper(newpathlevel).item(i) = CStr(newid)
        '        Me.NewWebPath(newpathlevel).ListIndex = i
        '        Exit For
        '    End If
        'Next i
        If newpathlevel <> 1 Then
            Me.NewWebPath(newpathlevel - 1).SetFocus
        End If
    End If
End Sub

Private Sub btnDeletePath_Click()
    If Me.NewWebPath(1) = "" Then
        Exit Sub
    End If
    
    Dim i As Long, pathLevel As Long
    For i = Me.NewWebPath.LBound To Me.NewWebPath.UBound
        If Me.NewWebPath(i) = "" Then
            pathLevel = i - 1
            Exit For
        End If
    Next i
    
    Dim pathid As String
    pathid = newWpMapper(pathLevel).item(Me.NewWebPath(pathLevel).ListIndex)
    
    If 0 <> DLookup("COUNT(*)", "WebPaths", "ParentID=" & pathid) Then
        MsgBox "Can't delete a path with subpaths!"
    Else
        If vbYes = MsgBox("Are you sure you want to delete " & qq(Me.NewWebPath(pathLevel)) & " and all associated item paths?", vbYesNo) Then
            DB.Execute "DELETE FROM PartNumberWebPaths WHERE WebPathID=" & pathid
            DB.Execute "DELETE FROM WebPaths WHERE ID=" & pathid
            requeryPathsNew pathLevel
        End If
    End If
End Sub

Private Sub btnRenamePath_Click()
    If Me.NewWebPath(1) = "" Then
        Exit Sub
    End If
    
    Dim i As Long, pathLevel As Long
    For i = Me.NewWebPath.LBound To Me.NewWebPath.UBound
        If Me.NewWebPath(i) = "" Then
            pathLevel = i - 1
            Exit For
        End If
    Next i
    
    Dim newname As String
    newname = InputBox("Rename " & qq(Me.NewWebPath(pathLevel)) & " to:", "Rename path")
    If newname <> "" Then
        DB.Execute "UPDATE WebPaths SET WebPathName='" & EscapeSQuotes(newname) & "' WHERE ID=" & newWpMapper(pathLevel).item(Me.NewWebPath(pathLevel).ListIndex)
        Me.NewWebPath(pathLevel).list(Me.NewWebPath(pathLevel).ListIndex) = newname
    End If
End Sub

Private Sub btnPathify_Click()
    If Me.ItemList.SelCount = 0 Then
        MsgBox "No items selected!"
    Else
        Mouse.Hourglass True
        Dim items As Variant
        items = ListBoxAsArray(Me.ItemList, True, 0)
        Dim i As Long
        Dim lastPathDx As Long
        For i = Me.NewWebPath.LBound To Me.NewWebPath.UBound
            If Me.NewWebPath(i) = "" Then
                lastPathDx = i - 1
                Exit For
            End If
            'Dim wpid As Long
            'wpid = newWpMapper(i).item(Me.NewWebPath(i).ListIndex)
            'Dim thisitem As Variant
            'For Each thisitem In items
            '    DB.Execute "INSERT INTO PartNumberWebPaths ( ItemNumber, WebPathID ) VALUES ( '" & CStr(thisitem) & "', " & wpid & " )", True
            'Next thisitem
        Next i
        Dim wpid As Long
        wpid = newWpMapper(lastPathDx).item(Me.NewWebPath(lastPathDx).ListIndex)
        Dim thisitem As Variant
        For Each thisitem In items
            DB.Execute "INSERT INTO PartNumberWebPaths ( ItemNumber, WebPathID ) VALUES ( '" & CStr(thisitem) & "', " & wpid & " )", True
        Next thisitem
        fillItemListNew
        If Me.chkShowMoved = vbUnchecked Then
            fillItemList
        End If
        Mouse.Hourglass False
    End If
End Sub

Private Sub btnUnpath_Click()
    If Me.NewItemList.SelCount <> 0 Then
        If MsgBox("Are you sure you want to unpath this item?", vbYesNo) = vbYes Then
            DB.Execute "DELETE FROM PartNumberWebPaths WHERE ItemNumber='" & Left(Me.NewItemList, InStr(Me.NewItemList, vbTab) - 1) & "'"
            Me.NewItemList.RemoveItem Me.NewItemList.ListIndex
        End If
    End If
End Sub

Private Sub chkAllPaths_Click()
    Dim i As Long
    For i = Me.WebPath.LBound To Me.WebPath.UBound
        Me.WebPath(i).Clear
    Next i
    If Me.chkAllPaths = vbChecked Then
        For i = Me.WebPath.LBound To Me.WebPath.UBound
            requeryPaths i
        Next i
    Else
        requeryPaths 1
    End If
End Sub

Private Sub chkShowMoved_Click()
    fillItemList
End Sub

Private Sub Form_Load()
    requeryPaths 1
    requeryPathsNew 1
    
    Dim tabs(0) As Long
    tabs(0) = 78
    SetListTabStops Me.ItemList.hwnd, tabs
    SetListTabStops Me.NewItemList.hwnd, tabs
End Sub

Private Sub WebPath_Click(Index As Integer)
    changed = True
    WebPath_LostFocus Index
End Sub

Private Sub WebPath_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.WebPath(Index), KeyCode, Shift
End Sub

Private Sub WebPath_KeyPress(Index As Integer, KeyAscii As Integer)
    AutoCompleteKeyPress Me.WebPath(Index), KeyAscii
    changed = True
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        WebPath_LostFocus Index
    End If
End Sub

Private Sub WebPath_LostFocus(Index As Integer)
    AutoCompleteLostFocus Me.WebPath(Index)
    If changed = True Then
        Dim i As Long
        If Me.WebPath(Index) <> wpSel(Index) Then
            wpSel(Index) = Me.WebPath(Index)
            If Me.chkAllPaths = vbUnchecked Then
                If Me.WebPath(Index) = "" Then
                    For i = Index + 1 To Me.WebPath.UBound
                        Me.WebPath(i).Clear
                        wpSel(i) = ""
                    Next i
                Else
                    If Index < Me.WebPath.UBound Then
                        requeryPaths Index + 1
                    End If
                End If
            End If
            fillItemList
        End If
        changed = False
    End If
End Sub

Private Sub requeryPaths(level As Long)
    Dim whereClause As String, i As Long
    Dim rst As ADODB.Recordset

    If Me.chkAllPaths Then
        Set rst = DB.retrieve("SELECT WebPathName FROM WebPaths WHERE PathLevel=" & level & " ORDER BY WebPathName")
    Else
        For i = 1 To level - 1
            whereClause = whereClause & "PartNumbers.ItemNumber IN (SELECT PartNumberWebPaths.ItemNumber FROM PartNumberWebPaths INNER JOIN WebPaths ON PartNumberWebPaths.WebPathID=WebPaths.ID WHERE WebPaths.WebPathName='" & EscapeSQuotes(WebPath(i)) & "' AND WebPaths.PathLevel=" & i & ") AND "
        Next i
        If Me.chkShowMoved = vbUnchecked Then
            whereClause = whereClause & "PartNumbers.ItemNumber NOT IN (SELECT DISTINCT ItemNumber FROM PartNumberWebPaths) AND "
        End If
        Set rst = DB.retrieve("SELECT DISTINCT WebPaths.WebPathName FROM PartNumbers INNER JOIN PartNumberWebPaths ON PartNumbers.ItemNumber=PartNumberWebPaths.ItemNumber INNER JOIN WebPaths ON PartNumberWebPaths.WebPathID=WebPaths.ID WHERE " & whereClause & "WebPaths.PathLevel=" & level & " AND (PartNumbers.WebPublished=1 OR PartNumbers.WebToBePublished=1) ORDER BY WebPaths.WebPathName")
    End If
    Me.WebPath(level).Clear
    Me.WebPath(level).AddItem ""
    wpSel(level) = ""
    While Not rst.EOF
        Me.WebPath(level).AddItem rst("WebPathName")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub fillItemList()
    Mouse.Hourglass True
    Dim whereClause As String, i As Long
    whereClause = " WHERE PartNumbers.BaseItemID=0 AND (PartNumbers.WebPublished=1 OR PartNumbers.WebToBePublished=1)"
    For i = Me.WebPath.LBound To Me.WebPath.UBound
        If Me.WebPath(i) = "" Then
            'skip, but don't exit, because we might be me.chkAllPaths==1
        Else
            whereClause = whereClause & " AND ItemNumber IN (SELECT ItemNumber FROM PartNumberWebPaths INNER JOIN WebPaths ON PartNumberWebPaths.WebPathID=WebPaths.ID WHERE WebPaths.WebPathName='" & EscapeSQuotes(Me.WebPath(i)) & "' AND PathLevel=" & i & ")"
        End If
    Next i
    If Me.chkShowMoved = vbUnchecked Then
        whereClause = whereClause & " AND ItemNumber NOT IN (SELECT DISTINCT ItemNumber FROM PartNumberWebPaths)"
    End If
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT PartNumbers.ItemNumber, COALESCE(PartNumbers.Desc1,'') + ' ' + COALESCE(PartNumbers.Desc3,'') + ' ' + COALESCE(PartNumbers.Desc4,'') AS ItemDesc FROM PartNumbers" & whereClause & " ORDER BY PartNumbers.ItemNumber")
    Me.ItemList.Clear
    While Not rst.EOF
        Me.ItemList.AddItem rst("ItemNumber") & vbTab & TrimWhitespace(CStr(rst("ItemDesc")))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
    
    Me.lblItemCount.caption = Me.ItemList.ListCount & " items in section"
End Sub


Private Sub NewWebPath_Click(Index As Integer)
    changed = True
    NewWebPath_LostFocus Index
End Sub

Private Sub NewWebPath_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.NewWebPath(Index), KeyCode, Shift
End Sub

Private Sub NewWebPath_KeyPress(Index As Integer, KeyAscii As Integer)
    AutoCompleteKeyPress Me.NewWebPath(Index), KeyAscii
    changed = True
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        NewWebPath_LostFocus Index
    End If
End Sub

Private Sub NewWebPath_LostFocus(Index As Integer)
    AutoCompleteLostFocus Me.NewWebPath(Index)
    If changed = True Then
        Dim i As Long
        If Me.NewWebPath(Index) <> newWpSel(Index) Then
            newWpSel(Index) = Me.NewWebPath(Index)
            If Me.NewWebPath(Index) = "" Then
                For i = Index + 1 To Me.NewWebPath.UBound
                    Me.NewWebPath(i).Clear
                    newWpSel(i) = ""
                    Set newWpMapper(i) = New Dictionary
                Next i
            Else
                If Index < Me.NewWebPath.UBound Then
                    requeryPathsNew Index + 1
                    For i = Index + 2 To Me.NewWebPath.UBound
                        Me.NewWebPath(i).Clear
                        newWpSel(i) = ""
                        Set newWpMapper(i) = New Dictionary
                    Next i
                End If
            End If
            fillItemListNew
            If Index <> Me.NewWebPath.UBound Then
                Me.lblSubpathCount(Index).caption = Me.NewWebPath(Index + 1).ListCount - 1 & " subpaths"
                For i = Index + 1 To Me.NewWebPath.UBound - 1
                    Me.lblSubpathCount(i).caption = ""
                Next i
            End If
        End If
        changed = False
    End If
End Sub

Private Sub requeryPathsNew(level As Long)
    Dim whereClause As String, i As Long
    whereClause = "PathLevel=" & level
    If level > 1 Then
        whereClause = whereClause & " AND ParentID=" & newWpMapper(level - 1).item(Me.NewWebPath(level - 1).ListIndex)
    End If
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT WebPaths.WebPathName, WebPaths.ID FROM WebPaths WHERE " & whereClause)
    Me.NewWebPath(level).Clear
    Me.NewWebPath(level).AddItem ""
    newWpSel(level) = ""
    'i = 1
    Dim foo As Dictionary
    Set foo = New Dictionary
    While Not rst.EOF
        Me.NewWebPath(level).AddItem rst("WebPathName")
        'newWpMapper(level).Add i, CStr(rst("ID"))
        foo.Add CStr(rst("WebPathName")), CStr(rst("ID"))
        rst.MoveNext
        'i = i + 1
    Wend
    rst.Close
    Set rst = Nothing
    
    Set newWpMapper(level) = New Dictionary
    newWpMapper(level).Add CLng(0), Null
    For i = 1 To Me.NewWebPath(level).ListCount - 1
        newWpMapper(level).Add i, foo.item(Me.NewWebPath(level).list(i))
    Next i
    Set foo = Nothing
End Sub

Private Sub fillItemListNew()
    Mouse.Hourglass True
    Dim whereClause As String, i As Long
    whereClause = " WHERE (PartNumbers.WebPublished=1 OR PartNumbers.WebToBePublished=1)"
    Dim lastWebPathDx As Long
    For i = Me.NewWebPath.LBound To Me.NewWebPath.UBound
        If Me.NewWebPath(i) = "" Then
            Exit For
        End If
        lastWebPathDx = i
        'whereClause = whereClause & " AND PartNumbers.ItemNumber IN (SELECT ItemNumber FROM PartNumberWebPaths WHERE WebPathID=" & newWpMapper(i).item(Me.NewWebPath(i).ListIndex) & ")"
    Next i
    Me.NewItemList.Clear
    If Me.NewWebPath(1) <> "" Then
        whereClause = whereClause & " AND PartNumbers.ItemNumber IN (SELECT ItemNumber FROM PartNumberWebPaths WHERE WebPathID=" & newWpMapper(lastWebPathDx).item(Me.NewWebPath(lastWebPathDx).ListIndex) & ")"
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT PartNumbers.ItemNumber, COALESCE(PartNumbers.Desc1,'') + ' ' + COALESCE(PartNumbers.Desc3,'') + ' ' + COALESCE(PartNumbers.Desc4,'') AS ItemDesc FROM PartNumbers" & whereClause & " ORDER BY PartNumbers.ItemNumber")
        While Not rst.EOF
            Me.NewItemList.AddItem rst("ItemNumber") & vbTab & TrimWhitespace(CStr(rst("ItemDesc")))
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
    End If
    Mouse.Hourglass False
    
    Me.lblNewItemCount.caption = Me.NewItemList.ListCount & " items in section"
End Sub
