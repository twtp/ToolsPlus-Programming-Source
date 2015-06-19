VERSION 5.00
Begin VB.UserControl ItemFinder 
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6900
   ScaleHeight     =   3855
   ScaleWidth      =   6900
   Begin VB.ListBox ItemsList 
      Height          =   1620
      ItemData        =   "ItemFinder.ctx":0000
      Left            =   30
      List            =   "ItemFinder.ctx":0002
      TabIndex        =   10
      Top             =   2130
      Width           =   6825
   End
   Begin VB.ComboBox WebPath 
      Height          =   315
      Index           =   5
      Left            =   990
      TabIndex        =   4
      Top             =   1470
      Width           =   3675
   End
   Begin VB.ComboBox WebPath 
      Height          =   315
      Index           =   4
      Left            =   990
      TabIndex        =   3
      Top             =   1110
      Width           =   3675
   End
   Begin VB.ComboBox WebPath 
      Height          =   315
      Index           =   3
      Left            =   990
      TabIndex        =   2
      Top             =   750
      Width           =   3675
   End
   Begin VB.ComboBox WebPath 
      Height          =   315
      Index           =   2
      Left            =   990
      TabIndex        =   1
      Top             =   390
      Width           =   3675
   End
   Begin VB.ComboBox WebPath 
      Height          =   315
      Index           =   1
      Left            =   990
      TabIndex        =   0
      Top             =   30
      Width           =   3675
   End
   Begin VB.Label generalLabel 
      Caption         =   "Description"
      Height          =   225
      Index           =   7
      Left            =   2550
      TabIndex        =   13
      Top             =   1920
      Width           =   1425
   End
   Begin VB.Label generalLabel 
      Caption         =   "Item Number"
      Height          =   225
      Index           =   6
      Left            =   510
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label generalLabel 
      Caption         =   "D/C?"
      Height          =   225
      Index           =   5
      Left            =   30
      TabIndex        =   11
      Top             =   1920
      Width           =   435
   End
   Begin VB.Label generalLabel 
      Caption         =   "Web Path 5"
      Height          =   225
      Index           =   4
      Left            =   30
      TabIndex        =   9
      Top             =   1530
      Width           =   975
   End
   Begin VB.Label generalLabel 
      Caption         =   "Web Path 4"
      Height          =   225
      Index           =   3
      Left            =   30
      TabIndex        =   8
      Top             =   1170
      Width           =   975
   End
   Begin VB.Label generalLabel 
      Caption         =   "Web Path 3"
      Height          =   225
      Index           =   2
      Left            =   30
      TabIndex        =   7
      Top             =   810
      Width           =   975
   End
   Begin VB.Label generalLabel 
      Caption         =   "Web Path 2"
      Height          =   225
      Index           =   1
      Left            =   30
      TabIndex        =   6
      Top             =   450
      Width           =   975
   End
   Begin VB.Label generalLabel 
      Caption         =   "Web Path 1"
      Height          =   225
      Index           =   0
      Left            =   30
      TabIndex        =   5
      Top             =   90
      Width           =   975
   End
End
Attribute VB_Name = "ItemFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event HasSelection()
Public Event LostSelection()
Public Event ListAction()

Private Const DEF_SEL_LIST = "Discontinued, ProductLine, RIGHT(ItemNumber,LEN(ItemNumber)-3) AS PartNumber, ItemDesc"
Private Const DEF_ORDER_BY = "Discontinued, ItemNumber"
'vb sucks
'Private Const DEF_ITEM_FLD = Array("1", "2")
Private DEF_ITEM_FLD As Variant

Private selList As String
Private orderby As String
Private itemfield As Variant
Private changed As Boolean
Private currWP(1 To 5) As String
Private fillAtStart As Boolean
Dim tabs() As Long

'====================
' PROPERTIES
'====================
Public Property Get FillOnNoSel() As Boolean
    FillOnNoSel = fillAtStart
End Property

Public Property Let FillOnNoSel(val As Boolean)
    fillAtStart = val
    'If fillAtStart Then
    '    Dim i As Long, skip As Boolean
    '    For i = UserControl.WebPath.LBound To UserControl.WebPath.UBound
    '        If UserControl.WebPath(i) <> "" Then
    '            skip = True
    '            i = UserControl.WebPath.UBound
    '        End If
    '    Next i
    '    If Not skip Then
    '        requeryItems
    '    End If
    'End If
End Property

Public Property Get ColumnList() As String
    ColumnList = selList
End Property

Public Property Let ColumnList(newColList As String)
    selList = newColList
End Property

Public Property Get OrderByList() As String
    OrderByList = orderby
End Property

Public Property Let OrderByList(newOrderList As String)
    orderby = newOrderList
End Property

Public Property Get CurrentTabStops() As Long()
    CurrentTabStops = tabs
End Property

Public Property Let CurrentTabStops(newTabs() As Long)
    Dim i As Long
    ReDim tabs(UBound(newTabs))
    For i = 0 To UBound(newTabs)
        tabs(i) = newTabs(i)
    Next i
    SetListTabStops UserControl.ItemsList.hwnd, tabs
End Property

Public Property Get ItemNumberField() As Variant
    ItemNumberField = itemfield
End Property

Public Property Let ItemNumberField(newField As Variant)
    If VarType(newField) = vbString Then
        If IsNumeric(newField) Then
            itemfield = newField
        Else
            MsgBox "ItemFinder: Error changing item number field, must be numeric"
        End If
    ElseIf VarType(newField) And vbArray Then
        Dim i As Long
        For i = LBound(newField) To UBound(newField)
            If Not IsNumeric(newField(i)) Then
                MsgBox "ItemFinder: Error changing item number field, must be numeric"
                Exit Property
            End If
        Next i
        itemfield = newField
    Else
        MsgBox "ItemFinder: Error changing item number field, must be string or array"
    End If
End Property


'======================
' METHODS
'====================
Public Sub Initialize()
    requeryWebPath "1"
    If fillAtStart Then
        requeryItems
    End If
End Sub
Public Sub ApplySettings(settings() As String)
    Dim i As Long
    For i = WebPath.UBound To WebPath.LBound
        WebPath(i) = settings(i)
        currWP(i) = settings(i)
        If i <> WebPath.UBound And WebPath(i) <> "" Then
            WebPath(i + 1).Enabled = True
            requeryWebPath CStr(i + 1)
            WebPath(i + 1).Enabled = True
        End If
    Next i
    requeryItems
End Sub

Public Function GetSettings() As String()
    Dim retval(1 To 5) As String
    Dim i As Long
    For i = WebPath.LBound To WebPath.UBound
        retval(i) = WebPath(i)
    Next i
    GetSettings = retval
End Function

Public Function GetSelectedItem() As String
    If ItemsList = "" Then
        GetSelectedItem = ""
    Else
        GetSelectedItem = getItemNumber(Split(UserControl.ItemsList, vbTab))
    End If
End Function

Public Function GetItemList(Optional excDisc As Boolean = False) As String()
    Dim i As Long, max As Long
    max = GetItemListCount(excDisc)
    ReDim retval(max - 1) As String
    'Dim pieces As Variant
    For i = 0 To max - 1
        'pieces = Split(ItemsList.list(i), vbTab)
        'retval(i) = pieces(1) & pieces(2)
        retval(i) = getItemNumber(Split(UserControl.ItemsList.list(i), vbTab))
    Next i
    GetItemList = retval
End Function

Public Function GetItemListCount(Optional excDisc As Boolean = False) As Long
    If excDisc Then
        Dim i As Long
        For i = UserControl.ItemsList.ListCount - 1 To 0 Step -1
            If Left(UserControl.ItemsList.list(i), 1) <> "D" Then
                GetItemListCount = i + 1
                i = 0
            End If
        Next i
    Else
        GetItemListCount = UserControl.ItemsList.ListCount
    End If
End Function



'===================
' INTERNAL
'===================

Private Sub ItemsList_Click()
    RaiseEvent HasSelection
End Sub

Private Sub ItemsList_DblClick()
    RaiseEvent ListAction
End Sub

Private Sub UserControl_Initialize()
    DEF_ITEM_FLD = Array("1", "2")
    selList = DEF_SEL_LIST
    orderby = DEF_ORDER_BY
    itemfield = DEF_ITEM_FLD
    'If Mouse Is Nothing Then
    '    Set Mouse = New MouseHourglass
    'End If
    ReDim tabs(2) As Long
    tabs(0) = 18
    tabs(1) = tabs(0) + 24
    tabs(2) = tabs(1) + 72
    SetListTabStops UserControl.ItemsList.hwnd, tabs
    'requeryWebPath "1"
    'If fillAtStart Then
    '    requeryItems
    'End If
End Sub

Private Sub UserControl_Resize()
    UserControl.ItemsList.width = UserControl.width - 60
    UserControl.ItemsList.Height = IIf(UserControl.Height - 2400 <= 0, 255, UserControl.Height - 1990)
End Sub

Private Sub WebPath_Click(Index As Integer)
    changed = True
    WebPath_LostFocus Index
End Sub

Private Sub WebPath_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown WebPath(Index), KeyCode, Shift
End Sub

Private Sub WebPath_KeyPress(Index As Integer, KeyAscii As Integer)
    AutoCompleteKeyPress WebPath(Index), KeyAscii
    changed = True
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        WebPath_LostFocus Index
    End If
End Sub

Private Sub WebPath_LostFocus(Index As Integer)
    AutoCompleteLostFocus WebPath(Index)
    If changed = True And WebPath(Index) <> currWP(Index) Then
        Dim i As Long
        For i = Index + 1 To WebPath.UBound
            WebPath(i).Clear
            WebPath(i).Enabled = False
            currWP(i) = ""
        Next i
        If WebPath(Index) <> "" Then
            WebPath(Index + 1).Enabled = True
            requeryWebPath CStr(Index + 1)
        End If
        requeryItems
        currWP(Index) = WebPath(Index)
        If UserControl.WebPath(Index) = "" And Index <> 1 Then
            UserControl.WebPath(Index - 1).SetFocus
        ElseIf UserControl.WebPath(Index) <> "" And Index <> 5 Then
            UserControl.WebPath(Index + 1).SetFocus
        End If
    End If
    changed = False
End Sub

Private Sub requeryWebPath(level As String)
    Dim whereclause As String, i As Long
    For i = 1 To level - 1
        whereclause = whereclause & "ItemNumber IN (SELECT ItemNumber FROM vPNWebPaths WHERE WebPathName='" & EscapeSQuotes(WebPath(i)) & "' AND PathLevel=" & i & ") AND "
    Next i
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT DISTINCT WebPathName FROM vPNWebPaths WHERE " & whereclause & " PathLevel=" & level & " ORDER BY WebPathName")
    UserControl.WebPath(level).Clear
    While Not rst.EOF
        UserControl.WebPath(level).AddItem rst("WebPathName")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    ExpandDropDownToFit UserControl.WebPath(level)
End Sub

Private Sub requeryItems()
    Dim whereclause As String, i As Long
    If WebPath(1) = "" Then
        whereclause = ""
    Else
        For i = WebPath.LBound To WebPath.UBound
            If WebPath(i) = "" Then
                i = WebPath.UBound
            Else
                whereclause = whereclause & "ItemNumber IN (SELECT ItemNumber FROM vPNWebPaths WHERE WebPathName='" & EscapeSQuotes(WebPath(i)) & "' AND PathLevel=" & i & ") AND "
            End If
        Next i
        whereclause = Left(whereclause, Len(whereclause) - 5)
    End If
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT " & selList & " FROM vInventoryInquiryItems " & IIf(whereclause = "", "", "WHERE " & whereclause) & " ORDER BY " & orderby)
    UserControl.ItemsList.Clear
    While Not rst.EOF
        UserControl.ItemsList.AddItem Join(RSFieldsAsArray(rst.fields), vbTab)
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    RaiseEvent LostSelection
End Sub

Private Function getItemNumber(pieces As Variant) As String
    If VarType(itemfield) = vbString Then
        getItemNumber = pieces(itemfield)
    Else
        Dim retval As String, i As Long
        For i = 0 To UBound(itemfield)
            retval = retval & pieces(itemfield(i))
        Next i
        getItemNumber = retval
    End If
End Function

