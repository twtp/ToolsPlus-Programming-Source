VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl SimpleListView 
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   ScaleHeight     =   3285
   ScaleWidth      =   4605
   Begin MSComctlLib.ListView ListView1 
      Height          =   3075
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   5424
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "SimpleListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : SimpleListView
' DateTime  : 9/26/2006 13:22
' Author    : briandonorfio
' Purpose   : Wrapper to make the ListView easier to work with. Most functions DWIM, all
'             integer indices start from 0.
'---------------------------------------------------------------------------------------

Option Explicit

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_HITTEST As Long = (LVM_FIRST + 18)
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVM_SUBITEMHITTEST As Long = (LVM_FIRST + 57)

Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Private Const LVHT_ABOVE = &H8
Private Const LVHT_BELOW = &H10
Private Const LVHT_TORIGHT = &H20
Private Const LVHT_TOLEFT = &H40
Private Const LVHT_NOWHERE As Long = &H1

Private Const LVHT_ONITEMICON As Long = &H2
Private Const LVHT_ONITEMLABEL As Long = &H4
Private Const LVHT_ONITEMSTATEICON As Long = &H8
Private Const LVHT_ONITEM As Long = (LVHT_ONITEMICON Or _
                                     LVHT_ONITEMLABEL Or _
                                     LVHT_ONITEMSTATEICON)
                                     
Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type HITTESTINFO
   pt As POINTAPI
   flags As Long
   iItem As Long
   iSubItem  As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Xpos As Single, Ypos As Single

Private lastSelectionList As Variant

Public Event DblClick(i As Long, j As Long, txt As String)
Public Event Click(i As Long, j As Long, txt As String)
Public Event ColumnClicked(i As Long)
Public Event SelectionChanged()

Public Property Get MultiSelect() As Boolean
    MultiSelect = UserControl.ListView1.MultiSelect
End Property

Public Property Let MultiSelect(newval As Boolean)
    If CanPropertyChange("MultiSelect") Then
        UserControl.ListView1.MultiSelect = newval
        PropertyChanged "MultiSelect"
    End If
End Property

Public Property Get Sorted() As Boolean
    Sorted = UserControl.ListView1.Sorted
End Property

Public Property Let Sorted(yesno As Boolean)
    If CanPropertyChange("Sorted") Then
        UserControl.ListView1.Sorted = yesno
        PropertyChanged "Sorted"
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.ListView1.Enabled
End Property

Public Property Let Enabled(newval As Boolean)
    If CanPropertyChange("Enabled") Then
        UserControl.ListView1.Enabled = newval
        PropertyChanged "Enabled"
    End If
End Property

'---------------------------------------------------------------------------------------
' Procedure : SetColumnNames
' DateTime  : 9/26/2006 13:14
' Author    : briandonorfio
' Purpose   : Sets the column headings to whatever is given. If it's a string, that's
'             the only column, if it's an array, those are the columns.
'---------------------------------------------------------------------------------------
'
Public Sub SetColumnNames(colNames As Variant)
    Dim i As Long
    If VarType(colNames) And vbArray Then
        For i = LBound(colNames) To UBound(colNames)
            UserControl.ListView1.ColumnHeaders.Add , , colNames(i)
        Next i
    ElseIf (VarType(colNames) And vbObject) And TypeName(colNames) = "Recordset" Then
        For i = 0 To colNames.Fields.Count - 1
            UserControl.ListView1.ColumnHeaders.Add , , colNames(i)
        Next i
    Else
        UserControl.ListView1.ColumnHeaders.Add , , colNames
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetColumnName
' Author    : briandonorfio
' Date      : 9/13/2011
' Purpose   : Returns column name for a given zero-based column index
'---------------------------------------------------------------------------------------
'
Public Function GetColumnName(colIndex As Long) As String
    GetColumnName = UserControl.ListView1.ColumnHeaders(colIndex + 1).Text
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetColumnWidths
' DateTime  : 10/10/2006 10:29
' Author    : briandonorfio
' Purpose   : Sets the widths of the columns. colWidths can either be a variant array,
'             or just a single value for the first column. Values should be an integer
'             number of twips. If colWidths is not supplied, autosizes to widest.
'---------------------------------------------------------------------------------------
'
Public Sub SetColumnWidths(Optional colWidths As Variant)
    Dim i As Long
    If IsMissing(colWidths) Then
        For i = 1 To UserControl.ListView1.ColumnHeaders.Count
            SendMessage UserControl.ListView1.hwnd, LVM_SETCOLUMNWIDTH, i, ByVal LVSCW_AUTOSIZE
        Next i
    ElseIf VarType(colWidths) And vbArray Then
        For i = LBound(colWidths) To UBound(colWidths)
            UserControl.ListView1.ColumnHeaders(i + 1).Width = colWidths(i)
        Next i
    Else
        UserControl.ListView1.ColumnHeaders(1).Width = colWidths
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ClearColumns
' Author    : briandonorfio
' Date      : 9/13/2011
' Purpose   : Removes all columns.
'---------------------------------------------------------------------------------------
'
Public Sub ClearColumns()
    UserControl.ListView1.ColumnHeaders.Clear
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Add
' DateTime  : 9/26/2006 13:15
' Author    : briandonorfio
' Purpose   : Adds a new row (or rows) to the list. Can pass a string, array, or
'             recordset as the item, it will figure out what to do, as long as the
'             columns are ok. If index is given, then it will place the new row at that
'             position, otherwise at the end of the list. If item is a recordset and
'             allRecs is true, then it will add all records in the set. If false, then
'             it will just add the current one and not advance the pointer.
'---------------------------------------------------------------------------------------
'
Public Sub Add(item As Variant, Optional Index As Long = -1, Optional allRecs As Boolean = False)
    Dim i As Long, thisrow As MSComctlLib.ListItem
    If Index = -1 Then
        Index = UserControl.ListView1.ListItems.Count
    Else
        Index = Index + 1
    End If
    If VarType(item) And vbArray Then
        Set thisrow = UserControl.ListView1.ListItems.Add(Index + 1, , item(0))
        thisrow.Selected = False 'why is this randomly true?
        For i = 1 To UBound(item)
            thisrow.ListSubItems.Add , , item(i)
        Next i
        Set thisrow = Nothing
    ElseIf (VarType(item) And vbObject) And TypeName(item) = "Recordset" Then
        If Not item.EOF Then
            Do
                Set thisrow = UserControl.ListView1.ListItems.Add(Index + 1, , item(0))
                thisrow.Selected = False
                For i = 1 To item.Fields.Count - 1
                    thisrow.ListSubItems.Add , , item(i)
                Next i
                Set thisrow = Nothing
                If allRecs Then
                    item.MoveNext
                    Index = Index + 1
                End If
            Loop While Not item.EOF And allRecs
        End If
    Else
        UserControl.ListView1.ListItems.Add Index, , item
    End If
    Set UserControl.ListView1.SelectedItem = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetText
' DateTime  : 9/26/2006 13:18
' Author    : briandonorfio
' Purpose   : Returns the value of the cell at the given row/column coordinates
'---------------------------------------------------------------------------------------
'
Public Function GetText(i As Long, Optional j As Long = 0) As String
    If j = 0 Then
        GetText = UserControl.ListView1.ListItems(i + 1)
    Else
        GetText = UserControl.ListView1.ListItems(i + 1).SubItems(j)
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetRow
' DateTime  : 9/26/2006 13:18
' Author    : briandonorfio
' Purpose   : Returns the row specified by the given index as a variant array
'---------------------------------------------------------------------------------------
'
Public Function GetRow(i As Long) As Variant
    ReDim retval(UserControl.ListView1.ColumnHeaders.Count - 1) As Variant
    retval(0) = UserControl.ListView1.ListItems(i + 1)
    Dim cnt As Long
    For cnt = 1 To UBound(retval)
        retval(cnt) = UserControl.ListView1.ListItems(i + 1).SubItems(cnt)
    Next cnt
    GetRow = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetTextSelected
' DateTime  : 9/26/2006 13:19
' Author    : briandonorfio
' Purpose   : Returns the value of the selected row, at the given column. If the control
'             is multiselect, returns an array of strings in order, otherwise just a
'             regular string.
'---------------------------------------------------------------------------------------
'
Public Function GetTextSelected(Optional j As Long = 0) As Variant
    Dim i As Long
    If UserControl.ListView1.MultiSelect Then
        'ReDim retval(-1) As Variant
        Dim retval As Variant
        retval = Split("", "zerolengtharray")
        For i = 1 To UserControl.ListView1.ListItems.Count
            Dim val As String
            If UserControl.ListView1.ListItems(i).Selected Then
                If j = 0 Then
                    val = UserControl.ListView1.ListItems(i).Text
                Else
                    val = UserControl.ListView1.ListItems(i).SubItems(j)
                End If
                ReDim Preserve retval(UBound(retval) + 1) As String
                retval(UBound(retval)) = val
            End If
        Next i
        GetTextSelected = retval
    Else
        If j = 0 Then
            GetTextSelected = CStr(UserControl.ListView1.SelectedItem)
        Else
            GetTextSelected = CStr(UserControl.ListView1.SelectedItem.SubItems(j))
        End If
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetRowSelected
' DateTime  : 9/26/2006 13:20
' Author    : briandonorfio
' Purpose   : If multiselect, returns the selected rows as a two-dimensional variant
'             array. If singleselect, returns the selected row as a variant array.
'---------------------------------------------------------------------------------------
'
Public Function GetRowSelected() As Variant
    Dim i As Long, j As Long, cnt As Long
    If UserControl.ListView1.MultiSelect Then
        'ReDim retvalMS(-1) As Variant
        ReDim retvalMS(SelCount - 1) As Variant
        'retvalMS = Split("", "zerolengtharray")
        cnt = 0
        For i = 1 To UserControl.ListView1.ListItems.Count
            If UserControl.ListView1.ListItems(i).Selected Then
                ReDim subarray(UserControl.ListView1.ColumnHeaders.Count - 1) As Variant
                subarray(0) = UserControl.ListView1.ListItems(i).Text
                For j = 1 To UBound(subarray)
                    subarray(j) = UserControl.ListView1.ListItems(i).SubItems(j)
                Next j
                'ReDim Preserve retvalMS(UBound(retvalMS) + 1) As Variant
                retvalMS(cnt) = subarray
                cnt = cnt + 1
            End If
        Next i
        GetRowSelected = retvalMS
    Else
        ReDim retvalSS(UserControl.ListView1.ColumnHeaders.Count - 1) As Variant
        retvalSS(0) = UserControl.ListView1.SelectedItem
        For i = 1 To UBound(retvalSS)
            retvalSS(i) = UserControl.ListView1.SelectedItem.SubItems(i)
        Next i
        GetRowSelected = retvalSS
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : Edit
' DateTime  : 9/26/2006 13:20
' Author    : briandonorfio
' Purpose   : Replaces the value of the given cell with newtext
'---------------------------------------------------------------------------------------
'
Public Sub Edit(newtext As String, i As Long, Optional j As Long = 0)
    If j = 0 Then
        UserControl.ListView1.ListItems(i + 1).Text = newtext
    Else
        UserControl.ListView1.ListItems(i + 1).SubItems(j) = newtext
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Remove
' DateTime  : 9/26/2006 13:21
' Author    : briandonorfio
' Purpose   : Removes the specified row
'---------------------------------------------------------------------------------------
'
Public Sub Remove(i As Long)
    UserControl.ListView1.ListItems.Remove i + 1
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RemoveSelected
' DateTime  : 9/26/2006 13:21
' Author    : briandonorfio
' Purpose   : Removes the currently selected row(s)
'---------------------------------------------------------------------------------------
'
Public Sub RemoveSelected()
    If UserControl.ListView1.MultiSelect Then
        Dim i As Long
        For i = UserControl.ListView1.ListItems.Count To 1 Step -1
            If UserControl.ListView1.ListItems(i).Selected Then
                UserControl.ListView1.ListItems.Remove i
            End If
        Next i
    Else
        UserControl.ListView1.ListItems.Remove UserControl.ListView1.SelectedItem.Index
        Set UserControl.ListView1.SelectedItem = Nothing
    End If
    If hasSelectionChanged() Then
        RaiseEvent SelectionChanged
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Clear
' DateTime  : 9/26/2006 13:21
' Author    : briandonorfio
' Purpose   : Removes all rows
'---------------------------------------------------------------------------------------
'
Public Sub Clear()
    UserControl.ListView1.ListItems.Clear
    Set UserControl.ListView1.SelectedItem = Nothing
    If hasSelectionChanged() Then
        RaiseEvent SelectionChanged
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InList
' DateTime  : 9/26/2006 13:21
' Author    : briandonorfio
' Purpose   : Searches for the given key in the first column of the list, returns true
'             if found, false otherwise
'---------------------------------------------------------------------------------------
'
Public Function InList(key As String) As Boolean
    Dim found As MSComctlLib.ListItem
    Set found = UserControl.ListView1.FindItem(key)
    InList = Not CBool(found Is Nothing)
End Function

'---------------------------------------------------------------------------------------
' Procedure : SelIndex
' DateTime  : 9/26/2006 13:22
' Author    : briandonorfio
' Purpose   : Returns -1 if nothing is selected, the index of the selected row as a
'             variant/long when multiselect is off, and a variant array when multiselect
'             is on
'---------------------------------------------------------------------------------------
'
Public Function SelIndex() As Variant
    If UserControl.ListView1.SelectedItem Is Nothing Then
        SelIndex = -1
    ElseIf UserControl.ListView1.MultiSelect = False Then
        SelIndex = UserControl.ListView1.SelectedItem.Index - 1
    Else
        Dim retval As String
        Dim i As Long
        For i = 1 To UserControl.ListView1.ListItems.Count
            If UserControl.ListView1.ListItems(i).Selected Then
                retval = retval & i - 1 & ","
            End If
        Next i
        If retval <> "" Then
            retval = Left(retval, Len(retval) - 1)
        End If
        SelIndex = Split(retval, ",")
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : ListCount
' DateTime  : 9/26/2006 13:22
' Author    : briandonorfio
' Purpose   : Returns the current number of items in the list
'---------------------------------------------------------------------------------------
'
Public Function ListCount() As Long
    ListCount = UserControl.ListView1.ListItems.Count
End Function

'---------------------------------------------------------------------------------------
' Procedure : SelCount
' DateTime  : 8/15/2008 15:52
' Author    : briandonorfio
' Purpose   : Returns the current number of selected items in the list
'---------------------------------------------------------------------------------------
'
Public Function SelCount() As Long
    Dim i As Long, retval As Long
    retval = 0
    For i = 1 To UserControl.ListView1.ListItems.Count
        If UserControl.ListView1.ListItems(i).Selected Then
            retval = retval + 1
        End If
    Next i
    SelCount = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : SelectRow
' DateTime  : 10/10/2006 10:19
' Author    : briandonorfio
' Purpose   : Selects/deselects the given row.
'---------------------------------------------------------------------------------------
'
Public Sub SelectRow(i As Long, Optional val As Boolean = True, Optional scrollToLine As Boolean = True, Optional suppressClickEvent As Boolean = False)
    If UserControl.ListView1.ListItems(i + 1).Selected = True Then
        If val Then
            'selected and want to select, do nothing
        Else
            'selected and want to deselect, do it..any event to fire?
            UserControl.ListView1.ListItems(i + 1).Selected = False
            If hasSelectionChanged() Then
                RaiseEvent SelectionChanged
            End If
        End If
    Else
        If val Then
            'not selected, want to select, ok, maybe raise an event
            UserControl.ListView1.ListItems(i + 1).Selected = True
            If scrollToLine Then
                UserControl.ListView1.ListItems(i + 1).EnsureVisible
            End If
            'If Not suppressClickEvent Then
            '    RaiseEvent Click(i, 0, UserControl.ListView1.ListItems(i + 1).Text)
            'End If
            If hasSelectionChanged() And Not suppressClickEvent Then
                RaiseEvent SelectionChanged
            End If
        Else
            'not selected, want to deselect, do nothing
        End If
    End If
End Sub



'-----------
' INTERNAL
'-----------


'---------------------------------------------------------------------------------------
' Procedure : ListView1_ColumnClick
' DateTime  : 8/15/2008 15:53
' Author    : briandonorfio
' Purpose   : Fires off the ColumnClick event, giving the column index
'---------------------------------------------------------------------------------------
'
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    RaiseEvent ColumnClicked(ColumnHeader.Index - 1)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ListView1_DblClick
' DateTime  : 9/26/2006 13:11
' Author    : briandonorfio
' Purpose   : This fires off the DblClick event, giving the column and row indices and
'             the text of that cell.
'---------------------------------------------------------------------------------------
'
Private Sub ListView1_DblClick()
    Dim hti As HITTESTINFO
    hti.pt.X = Xpos \ Screen.TwipsPerPixelX
    hti.pt.Y = Ypos \ Screen.TwipsPerPixelY
    hti.flags = LVHT_ONITEM
    SendMessage UserControl.ListView1.hwnd, LVM_SUBITEMHITTEST, 0&, hti
    Dim thisitem As MSComctlLib.ListItem
    If hti.iItem > -1 Then
        If hti.iSubItem = 0 Then
            RaiseEvent DblClick(hti.iItem, hti.iSubItem, UserControl.ListView1.ListItems(hti.iItem + 1))
        Else
            RaiseEvent DblClick(hti.iItem, hti.iSubItem, UserControl.ListView1.ListItems(hti.iItem + 1).SubItems(hti.iSubItem))
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ListView1_Click
' DateTime  : 4/10/2007 13:02
' Author    : BrianDonorfio
' Purpose   : Same as above, for single clicks
'---------------------------------------------------------------------------------------
'
Private Sub ListView1_Click()
    Dim hti As HITTESTINFO
    hti.pt.X = Xpos \ Screen.TwipsPerPixelX
    hti.pt.Y = Ypos \ Screen.TwipsPerPixelY
    hti.flags = LVHT_ONITEM
    SendMessage UserControl.ListView1.hwnd, LVM_SUBITEMHITTEST, 0&, hti
    Dim thisitem As MSComctlLib.ListItem
    If hti.iItem > -1 Then
        If hti.iSubItem = 0 Then
            RaiseEvent Click(hti.iItem, hti.iSubItem, UserControl.ListView1.ListItems(hti.iItem + 1))
        Else
            RaiseEvent Click(hti.iItem, hti.iSubItem, UserControl.ListView1.ListItems(hti.iItem + 1).SubItems(hti.iSubItem))
        End If
    End If
    If hasSelectionChanged() Then
        RaiseEvent SelectionChanged
    End If
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or _
       KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or _
       KeyCode = vbKeyPageUp Or vbKeyPageDown Then
        If hasSelectionChanged() Then
            RaiseEvent SelectionChanged
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ListView1_MouseDown
' DateTime  : 9/26/2006 13:12
' Author    : briandonorfio
' Purpose   : Hack to get it so things aren't always selected
'---------------------------------------------------------------------------------------
'
Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim hti As HITTESTINFO
    Dim i As Long
    hti.pt.X = X \ Screen.TwipsPerPixelX
    hti.pt.Y = Y \ Screen.TwipsPerPixelY
    hti.flags = LVHT_ABOVE Or LVHT_BELOW Or LVHT_TOLEFT Or LVHT_TORIGHT Or LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_NOWHERE
    i = SendMessage(UserControl.ListView1.hwnd, LVM_SUBITEMHITTEST, 0, hti)
    If i = -1 And (hti.iSubItem = -1 Or hti.iSubItem = 0) Then
        Set UserControl.ListView1.SelectedItem = Nothing
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ListView1_MouseMove
' DateTime  : 9/26/2006 13:12
' Author    : briandonorfio
' Purpose   : Updates the module level position variables, needed for _DblClick
'---------------------------------------------------------------------------------------
'
Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Xpos = X
    Ypos = Y
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UserControl_Resize
' DateTime  : 9/26/2006 13:13
' Author    : briandonorfio
' Purpose   : Sets the size of the control during placement + resizing
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_Resize()
    UserControl.ListView1.Width = UserControl.Width
    UserControl.ListView1.Height = UserControl.Height
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UserControl_GotFocus
' DateTime  : 9/26/2006 13:13
' Author    : briandonorfio
' Purpose   : Focuses on the ListView
'---------------------------------------------------------------------------------------
'
Private Sub UserControl_GotFocus()
    UserControl.ListView1.SetFocus
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.ListView1.MultiSelect = PropBag.ReadProperty("MultiSelect", False)
    UserControl.ListView1.Sorted = PropBag.ReadProperty("Sorted", False)
    UserControl.ListView1.Enabled = PropBag.ReadProperty("Enabled", True)
    
    If UserControl.ListView1.MultiSelect = False Then
        lastSelectionList = -1
    Else
        lastSelectionList = Empty
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "MultiSelect", UserControl.ListView1.MultiSelect
    PropBag.WriteProperty "Sorted", UserControl.ListView1.Sorted
    PropBag.WriteProperty "Enabled", UserControl.ListView1.Enabled
End Sub

Private Function hasSelectionChanged() As Boolean
    Dim currentSelection As Variant
    currentSelection = SelIndex()
    
    If UserControl.ListView1.MultiSelect = False Then
        'currentSelection is the zero-based index of the selected line or -1 for no selection
        'lastSelectionList is the zero-based index of the selected line or -1 for no selection
        hasSelectionChanged = CBool(CLng(lastSelectionList) <> CLng(currentSelection))
    Else
        'currentSelection is a variant array of strings with the the zero-based index, or -1 for no selection
        'lastSelectionList is a variant array of strings with the zero-based index or Empty
        If UserControl.ListView1.SelectedItem Is Nothing Then
            'check for no selection, check if last sel is an array (that's a bitwise and)
            hasSelectionChanged = CBool(VarType(lastSelectionList) And vbArray)
        ElseIf Not CBool(VarType(lastSelectionList) And vbArray) Then
            'check for no selection going to selection
            hasSelectionChanged = True
        ElseIf UBound(lastSelectionList) <> UBound(currentSelection) Then
            'check that we have the same number of selections before actually comparing
            hasSelectionChanged = True
        Else
            'long check index by index
            Dim i As Long
            For i = 0 To UBound(currentSelection)
                If CLng(currentSelection(i)) <> CLng(lastSelectionList(i)) Then
                    hasSelectionChanged = True
                    Exit For
                End If
            Next i
        End If
    End If
    'Debug.Print "changed: " & hasSelectionChanged
    
    'now set the last selection for next time
    setLastSelection currentSelection
End Function

Private Sub setLastSelection(currentSelection As Variant)
    If UserControl.ListView1.MultiSelect = False Then
        If UserControl.ListView1.SelectedItem Is Nothing Then
            lastSelectionList = -1
        Else
            lastSelectionList = CLng(currentSelection)
        End If
    Else
        If UserControl.ListView1.SelectedItem Is Nothing Then
            lastSelectionList = Empty
        Else
            lastSelectionList = currentSelection
        End If
    End If
End Sub
