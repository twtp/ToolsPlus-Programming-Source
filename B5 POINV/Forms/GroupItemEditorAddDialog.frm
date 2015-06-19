VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#3.5#0"; "TPControls.ocx"
Begin VB.Form GroupItemEditorAddDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group Item Editor - Add Items"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCacheClear 
      Caption         =   "Clear Specs Cache"
      Height          =   315
      Left            =   5340
      TabIndex        =   23
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame frameColMapSub 
      Enabled         =   0   'False
      Height          =   1335
      Left            =   5310
      TabIndex        =   15
      Top             =   4260
      Width           =   4335
      Begin VB.CheckBox chkFixIsRegex 
         Caption         =   "Regex"
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   900
         Width           =   915
      End
      Begin VB.TextBox ColumnBox 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   180
         Width           =   1815
      End
      Begin VB.ComboBox HighlightComboBox 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   510
         Width           =   1845
      End
      Begin VB.TextBox ColumnFixRegex 
         Height          =   285
         Left            =   960
         TabIndex        =   16
         Top             =   870
         Width           =   1785
      End
      Begin VB.Label generalLabel 
         Caption         =   "Column:"
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   21
         Top             =   210
         Width           =   855
      End
      Begin VB.Label generalLabel 
         Caption         =   "Highlight:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   540
         Width           =   855
      End
      Begin VB.Label generalLabel 
         Caption         =   "Fix:"
         Height          =   255
         Index           =   4
         Left            =   510
         TabIndex        =   19
         Top             =   900
         Width           =   345
      End
   End
   Begin VB.CheckBox chkHideDiscontinued 
      Caption         =   "Hide D/C"
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   900
      Value           =   1  'Checked
      Width           =   1185
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Add These"
      Height          =   405
      Left            =   8160
      TabIndex        =   13
      Top             =   3630
      Width           =   1515
   End
   Begin VB.TextBox LoadedPageID 
      Height          =   285
      Left            =   3360
      TabIndex        =   12
      Text            =   "LoadedPageID"
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox LoadedItem 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1050
      TabIndex        =   11
      Text            =   "LoadedItem"
      Top             =   150
      Width           =   2175
   End
   Begin VB.TextBox ItemsToLoadFilter 
      Height          =   285
      Left            =   600
      TabIndex        =   10
      Top             =   870
      Width           =   1665
   End
   Begin VB.ListBox ColumnMapListBox 
      Height          =   1230
      ItemData        =   "GroupItemEditorAddDialog.frx":0000
      Left            =   210
      List            =   "GroupItemEditorAddDialog.frx":0002
      TabIndex        =   6
      Top             =   4350
      Width           =   5055
   End
   Begin VB.CommandButton btnItemsRemove 
      Caption         =   "<="
      Height          =   435
      Left            =   4260
      TabIndex        =   5
      Top             =   2160
      Width           =   465
   End
   Begin VB.CommandButton btnItemsAdd 
      Caption         =   "=>"
      Height          =   435
      Left            =   4260
      TabIndex        =   4
      Top             =   1650
      Width           =   465
   End
   Begin TPControls.SimpleListView ItemsListView 
      Height          =   2415
      Left            =   4800
      TabIndex        =   3
      Top             =   1170
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   4260
      MultiSelect     =   0   'False
      Sorted          =   0   'False
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton btnItemsLoad 
      Caption         =   "Load"
      Height          =   255
      Left            =   2310
      TabIndex        =   2
      Top             =   900
      Width           =   615
   End
   Begin VB.ListBox AllItemsListBox 
      Height          =   2400
      ItemData        =   "GroupItemEditorAddDialog.frx":0004
      Left            =   120
      List            =   "GroupItemEditorAddDialog.frx":0006
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1170
      Width           =   4065
   End
   Begin VB.Label generalLabel 
      Caption         =   "Filter:"
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   9
      Top             =   900
      Width           =   405
   End
   Begin VB.Label generalLabel 
      Caption         =   "Column Mapping"
      Height          =   225
      Index           =   1
      Left            =   270
      TabIndex        =   8
      Top             =   4140
      Width           =   1215
   End
   Begin VB.Label generalLabel 
      Caption         =   "Items to Add"
      Height          =   225
      Index           =   0
      Left            =   4890
      TabIndex        =   7
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label generalLabel 
      Caption         =   "Editing:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   855
   End
End
Attribute VB_Name = "GroupItemEditorAddDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS As Long = 10

Private fillingForm As Boolean
Private addedItems As Dictionary
Private columnMap As PerlArray
Private specCache As Dictionary

Private Sub btnCacheClear_Click()
    Set specCache = New Dictionary
    requeryItemListView
End Sub

Private Sub btnItemsLoad_Click()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber, ItemDescription FROM vGroupItemPossibleAdditions WHERE ItemNumber LIKE '" & EscapeSQuotes(Me.ItemsToLoadFilter) & "%' AND WebsiteID=(SELECT WebsiteID FROM InventoryMaster WHERE ItemNumber='" & Me.LoadedItem & "')" & IIf(Me.chkHideDiscontinued, " AND Discontinued=0", ""))
    Me.AllItemsListBox.Clear
    While Not rst.EOF
        If addedItems.exists(CStr(rst("ItemNumber"))) Then
            'already to be added, skip
        Else
            Me.AllItemsListBox.AddItem rst("ItemNumber") & vbTab & rst("ItemDescription")
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub AllItemsListBox_DblClick()
    btnItemsAdd_Click
End Sub

Private Sub btnItemsAdd_Click()
    Dim i As Long
    Dim changed As Boolean
    changed = False
    For i = Me.AllItemsListBox.ListCount - 1 To 0 Step -1
        If Me.AllItemsListBox.Selected(i) Then
            Dim item As String
            item = ListBoxLineColumn(Me.AllItemsListBox, i, 0)
            If addedItems.exists(item) Then
                'do nothing
            Else
                changed = True
                addedItems.Add item, True
                Me.AllItemsListBox.RemoveItem i
            End If
        End If
    Next i
    
    If changed Then
        requeryItemListView
    End If
End Sub

Private Sub btnItemsRemove_Click()
    If Me.ItemsListView.SelCount = 0 Then
        Exit Sub
    End If
    
    Dim item As String
    item = Me.ItemsListView.GetTextSelected(0)
    addedItems.Remove item
    Me.ItemsListView.RemoveSelected
    Me.AllItemsListBox.AddItem item & vbTab & TrimWhitespace(specCache.item(item).item("Desc 1") & " " & specCache.item(item).item("Desc 3"), True, True, True)
End Sub

Private Sub ItemsListView_Click(i As Long, j As Long, txt As String)
    If j = 0 Then
        Exit Sub
    End If
    
    Me.ColumnMapListBox.Selected(j - 1) = True
End Sub

Private Sub ItemsListView_DblClick(i As Long, j As Long, txt As String)
    If j <> 0 Then
        Exit Sub
    End If
    
    btnItemsRemove_Click
End Sub

Private Sub ColumnMapListBox_Click()
    fillingForm = True
    If Me.ColumnMapListBox.ListIndex = -1 Then
        Me.ColumnBox = ""
        Me.HighlightComboBox.ListIndex = -1
        Me.ColumnFixRegex = ""
        Me.frameColMapSub.Enabled = False
    Else
        Dim temp As Variant
        temp = columnMap.Elem(Me.ColumnMapListBox.ListIndex)
        Me.ColumnBox = temp(0)
        If temp(1) = "" Then
            Me.HighlightComboBox.ListIndex = -1
        Else
            Dim i As Long
            For i = 0 To Me.HighlightComboBox.ListCount - 1
                If Me.HighlightComboBox.List(i) = temp(1) Then
                    Me.HighlightComboBox.ListIndex = i
                    Exit For
                End If
            Next i
        End If
        Me.ColumnFixRegex = temp(2)
        Me.chkFixIsRegex = temp(3)
        Me.frameColMapSub.Enabled = True
    End If
    fillingForm = False
End Sub

Private Sub HighlightComboBox_Click()
    If Not fillingForm Then
        Dim temp As Variant
        temp = columnMap.Elem(Me.ColumnMapListBox.ListIndex)
        temp(1) = Me.HighlightComboBox
        columnMap.Elem(Me.ColumnMapListBox.ListIndex) = temp
        
        Me.ColumnMapListBox.List(Me.ColumnMapListBox.ListIndex) = writeColumnLine(Me.ColumnMapListBox.ListIndex)
        
        requeryItemListView
    End If
End Sub

Private Sub ColumnFixRegex_Change()
    If Not fillingForm Then
        Dim temp As Variant
        temp = columnMap.Elem(Me.ColumnMapListBox.ListIndex)
        temp(2) = CStr(Me.ColumnFixRegex)
        columnMap.Elem(Me.ColumnMapListBox.ListIndex) = temp
        
        Me.ColumnMapListBox.List(Me.ColumnMapListBox.ListIndex) = writeColumnLine(Me.ColumnMapListBox.ListIndex)
        
        requeryItemListView
    End If
End Sub

Private Sub chkFixIsRegex_Click()
    If Not fillingForm Then
        Dim temp As Variant
        temp = columnMap.Elem(Me.ColumnMapListBox.ListIndex)
        temp(3) = CLng(IIf(Me.chkFixIsRegex, 1, 0))
        columnMap.Elem(Me.ColumnMapListBox.ListIndex) = temp
        
        requeryItemListView
    End If
End Sub

Private Sub btnSave_Click()
    If Me.ItemsListView.ListCount = 0 Then
        Exit Sub
    End If
    
    Dim sortNum As Long
    sortNum = DLookup("MAX(SortOrder)", "PartNumberGroupLines", "GroupID=" & Me.LoadedPageID)
    sortNum = sortNum + 1
    
    Dim i As Long
    For i = 0 To Me.ItemsListView.ListCount - 1
        Dim item As String
        item = Me.ItemsListView.GetText(i, 0)
        
        ReDim cols(0 To Me.ColumnMapListBox.ListCount - 1) As String
        cols = fillColumnsFor(item)
        
        Dim j As Long
        For j = 1 To UBound(cols)
            Dim colnum As Long
            colnum = j - 1
            DB.Execute "INSERT INTO PartNumberGroupLines ( GroupID, ItemNumber, ColumnIndex, ColumnValue, SortOrder ) VALUES ( " & Me.LoadedPageID & ", '" & item & "', " & colnum & ", '" & EscapeSQuotes(cols(j)) & "', " & sortNum & " )"
        Next j
        sortNum = sortNum + 1
    Next i
    
    If IsFormLoaded("GroupItemEditor") Then
        GroupItemEditor.NotifyGroupItemEditorSubitemsModified Me.LoadedPageID
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    fillingForm = False
    Set addedItems = New Dictionary
    Set columnMap = New PerlArray
    Set specCache = New Dictionary
    
    Me.ItemsListView.sorted = True
    
    Me.LoadedPageID = GroupItemEditor.LoadedPageID
    Me.LoadedItem = GroupItemEditor.TemplateItem
    Dim i As Long
    
    SetListTabs Me.AllItemsListBox, Array(72)
    SetListTabs Me.ColumnMapListBox, Array(72, 24)
    
    Me.ItemsToLoadFilter = Left(Me.LoadedItem, 3)
    
    fillItemsListViewColumns
    
    Me.HighlightComboBox.AddItem "Desc 1"
    Me.HighlightComboBox.AddItem "Desc 3"
    For i = 1 To 8
        Me.HighlightComboBox.AddItem "Spec " & i
    Next i
End Sub

Private Sub fillItemsListViewColumns()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ColumnValue FROM PartNumberGroupLines WHERE GroupID=" & Me.LoadedPageID & " AND ItemNumber='%COLUMN%' ORDER BY SortOrder, ColumnIndex")
    If rst.EOF Then
        'columns not yet init'ed
        MsgBox "Whoa, no columns! This is bad!"
    Else
        ReDim cols(0 To MAX_COLS - 1) As String
        cols(0) = "ItemNumber"
        Dim i As Long
        i = 1
        While Not rst.EOF
            cols(i) = rst("ColumnValue")
            i = i + 1
            columnMap.Push Array(CStr(rst("ColumnValue")), "", "", 0)
            Me.ColumnMapListBox.AddItem writeColumnLine(columnMap.Scalar - 1)
            rst.MoveNext
        Wend
        rst.Close
        
        Dim numCols As Long
        numCols = i - 1
        ReDim Preserve cols(0 To numCols) As String
        Me.ItemsListView.SetColumnNames cols
        Me.ItemsListView.SetColumnWidths 1500 'rest are default
    End If
End Sub

Private Sub requeryItemListView()
    Me.ItemsListView.Clear
    Dim item As Variant
    For Each item In addedItems.keys
        Me.ItemsListView.Add fillColumnsFor(CStr(item))
    Next item
End Sub

Private Function writeColumnLine(dx As Long) As String
    writeColumnLine = columnMap.Elem(dx)(0) & vbTab & _
                      IIf(columnMap.Elem(dx)(1) = "", "None", columnMap.Elem(dx)(1)) & vbTab & _
                      IIf(columnMap.Elem(dx)(2) = "", "---", columnMap.Elem(dx)(2))
End Function

Private Function fillColumnsFor(item As String) As String()
    Dim i As Long
    ReDim cols(0 To Me.ColumnMapListBox.ListCount) As String
    cols(0) = CStr(item)
    For i = 1 To UBound(cols)
        cols(i) = ""
    Next i
    
    If columnMap.Scalar > 0 Then
        If Not specCache.exists(item) Then
            Dim rst As ADODB.Recordset
            Set rst = DB.retrieve("SELECT Desc1, Desc3, Spec1HTML, Spec2HTML, Spec3HTML, Spec4HTML, Spec5HTML, Spec6HTML, Spec7HTML, Spec8HTML FROM PartNumbers WHERE ItemNumber='" & item & "'")
            specCache.Add item, New Dictionary
            specCache.item(item).Add "Desc 1", Nz(rst("Desc1"))
            specCache.item(item).Add "Desc 3", Nz(rst("Desc3"))
            specCache.item(item).Add "Spec 1", Nz(rst("Spec1HTML"))
            specCache.item(item).Add "Spec 2", Nz(rst("Spec2HTML"))
            specCache.item(item).Add "Spec 3", Nz(rst("Spec3HTML"))
            specCache.item(item).Add "Spec 4", Nz(rst("Spec4HTML"))
            specCache.item(item).Add "Spec 5", Nz(rst("Spec5HTML"))
            specCache.item(item).Add "Spec 6", Nz(rst("Spec6HTML"))
            specCache.item(item).Add "Spec 7", Nz(rst("Spec7HTML"))
            specCache.item(item).Add "Spec 8", Nz(rst("Spec8HTML"))
        End If
        For i = 1 To columnMap.Scalar
            Dim temp As Variant
            temp = columnMap.Elem(i - 1)
            
            Dim colName As String, colMappedTo As String, colFix As String, colRE As Long
            colName = CStr(temp(0))
            colMappedTo = CStr(temp(1))
            colFix = CStr(temp(2))
            colRE = CLng(temp(3))
            
            If colMappedTo <> "" Then
                cols(i) = specCache.item(item).item(colMappedTo)
                If colFix <> "" Then
                    If colRE = 0 Then
                        cols(i) = Replace(cols(i), colFix, "", 1, -1, vbTextCompare)
                    Else
                        Dim re As RegExp
                        Set re = New RegExp
                        re.Pattern = colFix
                        re.IgnoreCase = True
                        On Error GoTo badregex
                        Dim mc As MatchCollection
                        Set mc = re.Execute(cols(i))
                        If Not mc Is Nothing Then
                            If mc.count > 0 Then
                                If mc.item(0).SubMatches.count > 0 Then
                                    cols(i) = mc.item(0).SubMatches.item(0)
                                Else
                                    cols(i) = ""
                                End If
                            Else
                                cols(i) = ""
                            End If
                        Else
                            cols(i) = ""
                        End If
                        Set mc = Nothing
                        Set re = Nothing
                        On Error GoTo 0
                    End If
                End If
                cols(i) = Trim(cols(i))
            End If
        Next i
    End If
    
    fillColumnsFor = cols
    Exit Function

badregex:
    Set mc = Nothing
    Set re = Nothing
    cols(i) = ""
    Resume Next
End Function
