VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#3.5#0"; "TPControls.ocx"
Begin VB.Form GroupItemEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group Item Editor"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   14535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   120
      TabIndex        =   24
      Top             =   600
      Width           =   3855
      Begin VB.OptionButton opFilter 
         Caption         =   "Exclude Published"
         Height          =   255
         Index           =   2
         Left            =   2100
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   180
         Width           =   1605
      End
      Begin VB.OptionButton opFilter 
         Caption         =   "Only Published"
         Height          =   255
         Index           =   1
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   180
         Width           =   1365
      End
      Begin VB.OptionButton opFilter 
         Caption         =   "All"
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   180
         Value           =   -1  'True
         Width           =   585
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   465
      Left            =   9570
      TabIndex        =   7
      Top             =   0
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Height          =   8115
      Left            =   3360
      TabIndex        =   3
      Top             =   1110
      Width           =   11055
      Begin VB.CheckBox chkAvailabilityColumn 
         Caption         =   "Show Availability Column"
         Height          =   225
         Left            =   5100
         TabIndex        =   28
         Top             =   630
         Width           =   2325
      End
      Begin VB.CheckBox chkThumbnailsPerItem 
         Caption         =   "Show Item Thumbnails"
         Height          =   225
         Left            =   5100
         TabIndex        =   23
         Top             =   360
         Width           =   2325
      End
      Begin VB.CommandButton btnColumnHeadingsRemove 
         Caption         =   "Remove Column"
         Height          =   345
         Left            =   1950
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   7260
         Width           =   1335
      End
      Begin VB.CommandButton btnColumnHeadingsAdd 
         Caption         =   "Add Column"
         Height          =   345
         Left            =   1950
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton btnColumnHeadingsReorder 
         Caption         =   "Move Column"
         Height          =   345
         Left            =   510
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   7260
         Width           =   1335
      End
      Begin VB.CommandButton btnColumnHeadingsRename 
         Caption         =   "Rename Column"
         Height          =   345
         Left            =   510
         TabIndex        =   19
         Top             =   7680
         Width           =   1335
      End
      Begin VB.CommandButton btnItemMappingRemove 
         Caption         =   "Remove Item"
         Height          =   345
         Left            =   5820
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton btnColumnHeadingsEdit 
         Caption         =   "Edit Columns"
         Height          =   345
         Left            =   510
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton btnItemMappingAdd 
         Caption         =   "Add Items"
         Height          =   345
         Left            =   4380
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton btnGoToWebsite 
         Caption         =   "W"
         Height          =   255
         Left            =   4230
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   300
         Width           =   405
      End
      Begin VB.CommandButton btnGoToSigns 
         Caption         =   "GO"
         Height          =   255
         Left            =   4230
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   600
         Width           =   405
      End
      Begin VB.TextBox PublishStatus 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   870
         Width           =   2955
      End
      Begin VB.TextBox TemplateItem 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   570
         Width           =   2955
      End
      Begin VB.TextBox PageURL 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   270
         Width           =   2955
      End
      Begin VB.CommandButton btnMappedItemMoveDown 
         Caption         =   "\/"
         Height          =   285
         Left            =   10620
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2340
         Width           =   315
      End
      Begin VB.CommandButton btnMappedItemMoveUp 
         Caption         =   "/\"
         Height          =   285
         Left            =   10620
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2040
         Width           =   315
      End
      Begin TPControls.SimpleListView MappedItemListView 
         Height          =   5415
         Left            =   120
         TabIndex        =   4
         Top             =   1350
         Width           =   10425
         _ExtentX        =   18389
         _ExtentY        =   9551
         MultiSelect     =   0   'False
         Sorted          =   0   'False
         Enabled         =   -1  'True
      End
      Begin VB.Label generalLabel 
         Caption         =   "Status:"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   10
         Top             =   930
         Width           =   1065
      End
      Begin VB.Label generalLabel 
         Caption         =   "Template Item:"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   9
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label generalLabel 
         Caption         =   "Page URL:"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   330
         Width           =   1065
      End
   End
   Begin VB.TextBox LoadedPageID 
      Height          =   315
      Left            =   2970
      TabIndex        =   2
      Text            =   "ID"
      Top             =   90
      Visible         =   0   'False
      Width           =   525
   End
   Begin TPControls.SimpleListView GroupPageListView 
      Height          =   8055
      Left            =   90
      TabIndex        =   1
      Top             =   1200
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   14208
      MultiSelect     =   0   'False
      Sorted          =   0   'False
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Group Item Editor"
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
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2475
   End
End
Attribute VB_Name = "GroupItemEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS As Long = 10

Private changed As Boolean
Private fillingForm As Boolean

Public Sub GoToItem(item As String)
    Dim secondAttempt As Boolean
    secondAttempt = False
retry:
    Dim rst As ADODB.Recordset
    'lookup by pointer is ok because it's an index
    Set rst = DB.retrieve("SELECT GroupPageName FROM PartNumberGroupMaster WHERE ItemNumberPointer='" & item & "'")
    If rst.EOF Then
        MsgBox "Unable to find page with template item " & item & " in database!"
    Else
        Dim i As Long, found As Boolean
        found = False
        For i = 0 To Me.GroupPageListView.ListCount - 1
            If rst("GroupPageName") = Me.GroupPageListView.GetText(i, 0) Then
                Me.GroupPageListView.SelectRow i, True, True, False
                found = True
                Exit For
            End If
        Next i
        If found = False Then
            If Me.opFilter(0) = False Then
                Me.opFilter(0) = True
                GoTo retry
            Else
                If secondAttempt = True Then
                    MsgBox "Unable to find page with template item " & item & " in list!"
                Else
                    rst.Close
                    requeryGroupPageList
                    secondAttempt = True
                    GoTo retry
                End If
            End If
        End If
    End If
    rst.Close
    Set rst = Nothing
End Sub

Public Sub NotifyGroupItemEditorSubitemsModified(pageID As String)
    If Me.LoadedPageID = pageID Then
        fillMappedItemList
    End If
End Sub











Private Sub opFilter_Click(Index As Integer)
    requeryGroupPageList
End Sub

Private Sub btnColumnHeadingsReorder_Click()
'    MsgBox "TODO"
'    Exit Sub
    
    Dim col As Variant
    col = getColumnNumberFromPrompt("Enter number of column to move:")
    If VarType(col) = vbEmpty Then
        'cancel
        Exit Sub
    End If
    
    Dim col2 As Variant
    col2 = getColumnNumberFromPrompt("Enter number of column it should be placed after:", True)
    If VarType(col2) = vbEmpty Then
        'cancel
        Exit Sub
    End If
    
    If col(0) = col2(0) - 1 Or col(0) = col2(0) Then
        'offset for itemnumber column, they picked the same thing
        ' or
        'picked the previous column
        MsgBox "that doesn't make sense, dicknose"
        Exit Sub
    End If
    
    Dim col1DBdx As Long, col2DBdx As Long
    col1DBdx = col(0) - 1
    col2DBdx = col2(0) - 2
    If col1DBdx < col2DBdx Then
        DB.Execute "UPDATE PartNumberGroupLines SET ColumnIndex=99 WHERE GroupID=" & Me.LoadedPageID & " AND ColumnIndex=" & col1DBdx
        DB.Execute "UPDATE PartNumberGroupLines SET ColumnIndex=ColumnIndex-1 WHERE GroupID=" & Me.LoadedPageID & " AND ColumnIndex>" & col1DBdx & " AND ColumnIndex<=" & col2DBdx
        DB.Execute "UPDATE PartNumberGroupLines SET ColumnIndex=" & col2DBdx & " WHERE GroupID=" & Me.LoadedPageID & " AND ColumnIndex=99"
    Else
        DB.Execute "UPDATE PartNumberGroupLines SET ColumnIndex=99 WHERE GroupID=" & Me.LoadedPageID & " AND ColumnIndex=" & col1DBdx
        DB.Execute "UPDATE PartNumberGroupLines SET ColumnIndex=ColumnIndex+1 WHERE GroupID=" & Me.LoadedPageID & " AND ColumnIndex<" & col1DBdx & " AND ColumnIndex>" & col2DBdx
        DB.Execute "UPDATE PartNumberGroupLines SET ColumnIndex=" & col2DBdx + 1 & " WHERE GroupID=" & Me.LoadedPageID & " AND ColumnIndex=99"
    End If
    
    fillMappedItemList
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnGoToWebsite_Click()
    OpenDefaultApp "http://www.tools-plus.com/" & Me.PageURL & ".html"
End Sub

Private Sub btnGoToSigns_Click()
    If Not IsFormLoaded("SignMaintenance") Then
        Load SignMaintenance
    End If
    SignMaintenance.GoToItem Me.TemplateItem
    If IsFormMinimized("SignMaintenance") Then
        UnMinimizeForm "SignMaintenance"
    End If
    If SignMaintenance.Visible = False Then
        SignMaintenance.Show
    End If
    FocusOnForm "SignMaintenance"
End Sub

Private Sub btnColumnHeadingsAdd_Click()
    Dim colName As String
    colName = InputBox("Enter new column name:")
    If colName = "" Then
        Exit Sub
    End If
    
    Dim colnum As String
    colnum = DLookup("MAX(ColumnIndex)", "PartNumberGroupLines", "GroupID=" & Me.LoadedPageID)
    If colnum = "" Then
        colnum = "0"
    Else
        colnum = CStr(CLng(colnum) + 1)
    End If
    
    DB.Execute "INSERT INTO PartNumberGroupLines ( GroupID, ItemNumber, ColumnIndex, ColumnValue, SortOrder ) VALUES ( " & Me.LoadedPageID & ", '%COLUMN%', " & colnum & ", '" & EscapeSQuotes(colName) & "', -1 )"

    DB.Execute "INSERT INTO PartNumberGroupLines (GroupID, ItemNumber, ColumnIndex, ColumnValue, SortOrder) " & _
               "SELECT GroupID, ItemNumber, " & colnum & ", '', SortOrder " & _
               "FROM PartNumberGroupLines " & _
               "WHERE GroupID=" & Me.LoadedPageID & _
               "  AND ColumnIndex=0" & _
               "  AND ItemNumber<>'%COLUMN%'"
    
    fillMappedItemList
End Sub

Private Sub btnColumnHeadingsEdit_Click()
    Dim columnsA As PerlArray
    Set columnsA = New PerlArray
    
    Dim i As Long
    For i = 1 To MAX_COLS
        Dim thisColName As String
        thisColName = InputBox("Enter column " & i & " name (type ""END"" to finish, or cancel to cancel all entries so far):", "Enter column name", "")
        If thisColName = "" Then
            'cancelled
            Set columnsA = Nothing
            Exit Sub
        ElseIf thisColName = "END" Then
            'end of inputs, break and continue
            Exit For
        Else
            'column added
            columnsA.Push thisColName
        End If
    Next i
    
    If columnsA.Scalar = 0 Then
        MsgBox "No columns entered!"
        Set columnsA = Nothing
        Exit Sub
    End If
    
    If vbNo = MsgBox("Proceed with the following columns?" & vbCrLf & vbCrLf & columnsA.Join(vbCrLf), vbYesNo) Then
        Set columnsA = Nothing
        Exit Sub
    End If
    
    DB.Execute "DELETE FROM PartNumberGroupLines WHERE GroupID=" & Me.LoadedPageID
    For i = 0 To columnsA.Scalar - 1
        DB.Execute "INSERT INTO PartNumberGroupLines ( GroupID, ItemNumber, ColumnIndex, ColumnValue, SortOrder ) VALUES ( " & Me.LoadedPageID & ", '%COLUMN%', " & i & ", '" & EscapeSQuotes(columnsA.Elem(i)) & "', -1 )"
    Next i
    
    Set columnsA = Nothing
    
    fillMappedItemList
End Sub

Private Sub btnColumnHeadingsRename_Click()
    Dim col As Variant
    col = getColumnNumberFromPrompt("Enter number of column to rename:")
    If VarType(col) = vbEmpty Then
        'cancel
    Else
        Dim colnum As String
        Dim colName As String
        colnum = col(0)
        colName = col(1)
        
        Dim newval As String
        newval = InputBox("Enter new name for old column", "New Name", colName)
        If newval = "" Then
            'cancel
        Else
            DB.Execute "UPDATE PartNumberGroupLines SET ColumnValue='" & EscapeSQuotes(newval) & "' WHERE GroupID=" & Me.LoadedPageID & " AND ItemNumber='%COLUMN%' AND ColumnIndex=" & (CLng(colnum) - 1)
            fillMappedItemList
        End If
    End If
End Sub

Private Sub btnColumnHeadingsRemove_Click()
    Dim col As Variant
    col = getColumnNumberFromPrompt("Enter number of column to remove:")
    If VarType(col) = vbEmpty Then
        'cancel
    Else
        Dim colnum As String
        Dim colName As String
        colnum = col(0)
        colName = col(1)
        
        If vbYes = MsgBox("Are you sure you want to REMOVE column " & qq(colName) & "?", vbYesNo) Then
            colnum = CStr(CLng(colnum) - 1)
            DB.Execute "DELETE FROM PartNumberGroupLines WHERE GroupID=" & Me.LoadedPageID & " AND ColumnIndex=" & colnum
            DB.Execute "UPDATE PartNumberGroupLines SET ColumnIndex=ColumnIndex-1 WHERE GroupID=" & Me.LoadedPageID & " AND ColumnIndex>" & colnum
            
            fillMappedItemList
        End If
    End If
End Sub

Private Sub btnItemMappingAdd_Click()
    Load GroupItemEditorAddDialog
    GroupItemEditorAddDialog.Show 'MODAL
    
    'fillMappedItemList
End Sub

Private Sub btnItemMappingRemove_Click()
    If Me.MappedItemListView.SelCount = 0 Then
        Exit Sub
    End If
    
    If vbNo = MsgBox("Remove " & Me.MappedItemListView.GetTextSelected(0) & "?", vbYesNo) Then
        Exit Sub
    End If
    
    DB.Execute "DELETE FROM PartNumberGroupLines WHERE GroupID=" & Me.LoadedPageID & " AND ItemNumber='" & Me.MappedItemListView.GetTextSelected(0) & "'"
    DB.Execute "UPDATE PartNumberGroupLines SET SortOrder=SortOrder-1 WHERE GroupID=" & Me.LoadedPageID & " AND SortOrder>" & Me.MappedItemListView.SelIndex
    
    Me.MappedItemListView.RemoveSelected
    
    If Me.MappedItemListView.ListCount = 0 Then
        Me.btnColumnHeadingsEdit.Enabled = True
    End If
End Sub

Private Sub btnMappedItemMoveDown_Click()
    If Me.MappedItemListView.SelCount = 0 Then
        Exit Sub
    End If
    If Me.MappedItemListView.SelIndex = Me.MappedItemListView.ListCount - 1 Then
        Exit Sub
    End If
    
    Dim thisrow As Variant, nextrow As Variant
    thisrow = Me.MappedItemListView.GetRowSelected
    nextrow = Me.MappedItemListView.GetRow(Me.MappedItemListView.SelIndex + 1)
    
    DB.Execute "UPDATE PartNumberGroupLines SET SortOrder=SortOrder+1 WHERE GroupID=" & Me.LoadedPageID & " AND ItemNumber='" & thisrow(0) & "'"
    DB.Execute "UPDATE PartNumberGroupLines SET SortOrder=SortOrder-1 WHERE GroupID=" & Me.LoadedPageID & " AND ItemNumber='" & nextrow(0) & "'"
    
    Dim i As Long
    For i = 0 To UBound(thisrow)
        Me.MappedItemListView.Edit CStr(nextrow(i)), Me.MappedItemListView.SelIndex, i
        Me.MappedItemListView.Edit CStr(thisrow(i)), Me.MappedItemListView.SelIndex + 1, i
    Next i
    
    Me.MappedItemListView.SelectRow Me.MappedItemListView.SelIndex + 1
End Sub

Private Sub btnMappedItemMoveUp_Click()
    If Me.MappedItemListView.SelCount = 0 Then
        Exit Sub
    End If
    If Me.MappedItemListView.SelIndex = 0 Then
        Exit Sub
    End If
    
    Dim thisrow As Variant, prevrow As Variant
    thisrow = Me.MappedItemListView.GetRowSelected
    prevrow = Me.MappedItemListView.GetRow(Me.MappedItemListView.SelIndex - 1)
    
    DB.Execute "UPDATE PartNumberGroupLines SET SortOrder=SortOrder-1 WHERE GroupID=" & Me.LoadedPageID & " AND ItemNumber='" & thisrow(0) & "'"
    DB.Execute "UPDATE PartNumberGroupLines SET SortOrder=SortOrder+1 WHERE GroupID=" & Me.LoadedPageID & " AND ItemNumber='" & prevrow(0) & "'"
    
    Dim i As Long
    For i = 0 To UBound(thisrow)
        Me.MappedItemListView.Edit CStr(prevrow(i)), Me.MappedItemListView.SelIndex, i
        Me.MappedItemListView.Edit CStr(thisrow(i)), Me.MappedItemListView.SelIndex - 1, i
    Next i
    
    Me.MappedItemListView.SelectRow Me.MappedItemListView.SelIndex - 1
End Sub

Private Sub chkThumbnailsPerItem_Click()
    If Not fillingForm Then
        DB.Execute "UPDATE PartNumberGroupMaster SET ShowItemThumbnails=" & IIf(Me.chkThumbnailsPerItem, "1", "0") & " WHERE ID=" & Me.LoadedPageID
    End If
End Sub

Private Sub chkAvailabilityColumn_Click()
    If Not fillingForm Then
        DB.Execute "UPDATE PartNumberGroupMaster SET ShowAvailabilityColumn=" & IIf(Me.chkAvailabilityColumn, "1", "0") & " WHERE ID=" & Me.LoadedPageID
    End If
End Sub

Private Sub Form_Load()
    Me.Frame1.Enabled = False
    Me.MappedItemListView.Enabled = False
    Me.GroupPageListView.sorted = True
    Me.GroupPageListView.SetColumnNames Array("Page Name", "Pub")
    Me.GroupPageListView.SetColumnWidths Array(2300, 500)
    
    requeryGroupPageList
End Sub

Private Sub GroupPageListView_Click(i As Long, j As Long, txt As String)
    Dim pn As String
    pn = Me.GroupPageListView.GetText(i)
    Dim rst As ADODB.Recordset
    'lookup by page name is ok because it's a pdx
    Set rst = DB.retrieve("SELECT ID, GroupPageName, ItemNumberPointer, ShowItemThumbnails, ShowAvailabilityColumn FROM PartNumberGroupMaster WHERE GroupPageName='" & EscapeSQuotes(pn) & "'")
    If rst.EOF Then
        MsgBox "Can't find the page? That's really funny."
        Me.Frame1.Enabled = False
        Me.MappedItemListView.Enabled = False
    ElseIf Me.LoadedPageID = rst("ID") Then
        'skip
    Else
        Me.Frame1.Enabled = True
        Me.MappedItemListView.Enabled = True
        fillingForm = True
        Me.LoadedPageID = rst("ID")
        Me.PageURL = rst("GroupPageName")
        Me.TemplateItem = rst("ItemNumberPointer")
        Me.PublishStatus = getTemplateStatus()
        Me.chkThumbnailsPerItem = SQLBool(rst("ShowItemThumbnails"))
        Me.chkAvailabilityColumn = SQLBool(rst("ShowAvailabilityColumn"))
        fillMappedItemList
        fillingForm = False
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub MappedItemListView_DblClick(i As Long, j As Long, txt As String)
    Select Case j
        Case Is = 0
            btnItemMappingRemove_Click
        Case Else
            Dim item As String
            item = Me.MappedItemListView.GetText(i, 0)
            Dim val As String
            val = InputBox("Edit Value for " & item & " col " & qq(Me.MappedItemListView.GetColumnName(j)) & ":", "Edit Value:", txt)
            If val = "" Then
                'cancel
            ElseIf val = txt Then
                'no change
            ElseIf CBool(InStr(val, "~")) Then
                'tilde is used as separator and can't be literal
                MsgBox "Can't use ""~"" characters in columns!"
            Else
                DB.Execute "UPDATE PartNumberGroupLines SET ColumnValue='" & EscapeSQuotes(val) & "' WHERE GroupID=" & Me.LoadedPageID & " AND ItemNumber='" & EscapeSQuotes(item) & "' AND ColumnIndex=" & (j - 1)
                Me.MappedItemListView.Edit val, i, j
            End If
    End Select
End Sub

Private Sub requeryGroupPageList()
    Dim whereClause As String
    Select Case True
        Case Is = Me.opFilter(0)
            whereClause = ""
        Case Is = Me.opFilter(1)
            whereClause = " WHERE WebPublished=1"
        Case Is = Me.opFilter(2)
            whereClause = " WHERE WebPublished IS NULL OR WebPublished=0"
        Case Else
            MsgBox "Uh oh, Brian forgot to put this filter option in the switch!"
            whereClause = ""
    End Select
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT GroupPageName, CASE WHEN WebPublished=1 THEN 'Y' WHEN WebDeleted=1 THEN 'D' ELSE 'N' END FROM PartNumberGroupMaster LEFT OUTER JOIN PartNumbers ON PartNumberGroupMaster.ItemNumberPointer=PartNumbers.ItemNumber" & whereClause & " ORDER BY GroupPageName")
    Me.GroupPageListView.Clear
    Me.GroupPageListView.Add rst, -1, True
    rst.Close
    Set rst = Nothing
End Sub

Private Sub fillMappedItemList()
    Me.MappedItemListView.Clear
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ColumnValue FROM PartNumberGroupLines WHERE GroupID=" & Me.LoadedPageID & " AND ItemNumber='%COLUMN%' ORDER BY SortOrder, ColumnIndex")
    If rst.EOF Then
        'columns not yet init'ed
        Me.btnColumnHeadingsEdit.Enabled = True
        Me.btnColumnHeadingsRename.Enabled = False
        Me.btnColumnHeadingsAdd.Enabled = True
        Me.btnColumnHeadingsRemove.Enabled = False
        Me.btnColumnHeadingsReorder.Enabled = False
        Me.btnItemMappingAdd.Enabled = False
        Me.btnItemMappingRemove.Enabled = False
        Me.MappedItemListView.ClearColumns
    Else
        ReDim cols(0 To MAX_COLS - 1) As String
        cols(0) = "ItemNumber"
        Dim i As Long
        i = 1
        While Not rst.EOF
            cols(i) = rst("ColumnValue")
            i = i + 1
            rst.MoveNext
        Wend
        rst.Close
        
        Dim numCols As Long
        numCols = i - 1
        ReDim Preserve cols(0 To numCols) As String
        Me.MappedItemListView.ClearColumns
        Me.MappedItemListView.SetColumnNames cols
        Me.MappedItemListView.SetColumnWidths 1500 'rest are default
        
        Set rst = DB.retrieve("SELECT ItemNumber, ColumnValue FROM PartNumberGroupLines WHERE GroupID=" & Me.LoadedPageID & " AND ItemNumber<>'%COLUMN%' ORDER BY SortOrder, ColumnIndex")
        While Not rst.EOF
            cols(0) = rst("ItemNumber")
            For i = 1 To numCols
                cols(i) = rst("ColumnValue")
                rst.MoveNext
            Next i
            Me.MappedItemListView.Add cols
        Wend
        
        Me.btnColumnHeadingsEdit.Enabled = CBool(Me.MappedItemListView.ListCount = 0)
        Me.btnColumnHeadingsRename.Enabled = True
        Me.btnColumnHeadingsAdd.Enabled = True
        Me.btnColumnHeadingsRemove.Enabled = True
        Me.btnColumnHeadingsReorder.Enabled = True
        Me.btnItemMappingAdd.Enabled = True
        Me.btnItemMappingRemove.Enabled = True
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Function getTemplateStatus() As String
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT WebToBePublished, WebPublished, WebDelete, WebDeleted FROM PartNumbers WHERE ItemNumber='" & Me.TemplateItem & "'")
    Dim statusA As PerlArray
    Set statusA = New PerlArray
    If rst.EOF Then
        statusA.Push "ERROR - EOF"
    Else
        If rst("WebToBePublished") Then
            statusA.Push "T/B Pub"
        End If
        If rst("WebPublished") Then
            statusA.Push "Published"
        End If
        If rst("WebDelete") Then
            statusA.Push "Delete"
        End If
        If rst("WebDeleted") Then
            statusA.Push "Deleted"
        End If
    End If
    rst.Close
    Set rst = Nothing
    
    If statusA.Scalar = 0 Then
        getTemplateStatus = ""
    ElseIf statusA.Scalar = 1 Then
        getTemplateStatus = statusA.Elem(0)
    Else
        getTemplateStatus = statusA.Join(" & ")
    End If
    Set statusA = Nothing
End Function

Private Function getColumnNumberFromPrompt(msg As String, Optional includeItemCol As Boolean = False) As Variant
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ColumnIndex, ColumnValue FROM PartNumberGroupLines WHERE GroupID=" & Me.LoadedPageID & " AND ItemNumber='%COLUMN%' ORDER BY ColumnIndex")
    Dim fullMsg As String
    If Right(msg, 2) = vbCrLf Then
        fullMsg = msg
    Else
        fullMsg = msg & vbCrLf
    End If
    Dim colCount As Long, colHash As Dictionary, colOffset As Long
    Set colHash = New Dictionary
    If includeItemCol Then
        colCount = rst.RecordCount + 1
        fullMsg = fullMsg & vbCrLf & "1: ItemNumber"
        colHash.Add "1", "ItemNumber"
        colOffset = 2
    Else
        colCount = rst.RecordCount
        colOffset = 1
    End If
    While Not rst.EOF
        fullMsg = fullMsg & vbCrLf & (rst("ColumnIndex") + colOffset) & ": " & rst("ColumnValue")
        colHash.Add CStr(rst("ColumnIndex") + colOffset), CStr(rst("ColumnValue"))
        rst.MoveNext
    Wend
    Dim col As String, colnum As Long
    col = InputBox(fullMsg)
    If col = "" Then
        'cancel
        getColumnNumberFromPrompt = Empty
    ElseIf IsIntegeric(col) = False Then
        'bad
        MsgBox "Must be a number!"
        getColumnNumberFromPrompt = Empty
    ElseIf CLng(col) = 0 Then
        'bad
        MsgBox "Must be greater than 0"
        getColumnNumberFromPrompt = Empty
    ElseIf CLng(col) > colCount Then
        'also bad
        MsgBox "Must be less than " & colCount
        getColumnNumberFromPrompt = Empty
    Else
        'good
        getColumnNumberFromPrompt = Array(CLng(col), CStr(colHash.item(CStr(col))))
    End If
End Function
