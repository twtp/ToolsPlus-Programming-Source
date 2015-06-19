VERSION 5.00
Begin VB.Form WebSiteFrontPage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Front Page Maintenance"
   ClientHeight    =   11415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11415
   ScaleWidth      =   8925
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame generalFrame 
      Height          =   1875
      Index           =   6
      Left            =   120
      TabIndex        =   60
      Top             =   9480
      Width           =   8655
      Begin VB.CommandButton btnRemoveItem 
         Caption         =   "Delete"
         Height          =   315
         Index           =   4
         Left            =   7860
         TabIndex        =   68
         Top             =   1470
         Width           =   705
      End
      Begin VB.CommandButton btnAddItem 
         Caption         =   "Add"
         Height          =   315
         Index           =   4
         Left            =   5610
         TabIndex        =   67
         Top             =   1470
         Width           =   585
      End
      Begin VB.ComboBox cmbAddItem 
         Height          =   315
         Index           =   4
         Left            =   3300
         TabIndex        =   66
         Top             =   1470
         Width           =   2295
      End
      Begin VB.ListBox fpColumn 
         Height          =   1035
         Index           =   4
         ItemData        =   "WebSiteFrontPage.frx":0000
         Left            =   90
         List            =   "WebSiteFrontPage.frx":0002
         TabIndex        =   65
         Top             =   360
         Width           =   8475
      End
      Begin VB.CommandButton btnMoveDown 
         Caption         =   "Down"
         Height          =   315
         Index           =   4
         Left            =   7260
         TabIndex        =   64
         Top             =   1470
         Width           =   555
      End
      Begin VB.CommandButton btnMoveUp 
         Caption         =   "Up"
         Height          =   315
         Index           =   4
         Left            =   6780
         TabIndex        =   63
         Top             =   1470
         Width           =   465
      End
      Begin VB.ComboBox cmbColName 
         Height          =   315
         Index           =   4
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   30
         Width           =   3075
      End
      Begin VB.CommandButton btnMoveTop 
         Caption         =   "Top"
         Height          =   315
         Index           =   4
         Left            =   6300
         TabIndex        =   61
         Top             =   1470
         Width           =   465
      End
   End
   Begin VB.Frame generalFrame 
      Height          =   855
      Index           =   5
      Left            =   6030
      TabIndex        =   47
      Top             =   540
      Width           =   2865
      Begin VB.CommandButton btnGenerate 
         Caption         =   "Generate"
         Height          =   315
         Left            =   120
         TabIndex        =   51
         Top             =   150
         Width           =   2595
      End
      Begin VB.TextBox numInCategory 
         Height          =   255
         Left            =   690
         TabIndex        =   48
         Top             =   510
         Width           =   585
      End
      Begin VB.Label generalLabel 
         Caption         =   "use top"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   50
         Top             =   510
         Width           =   615
      End
      Begin VB.Label generalLabel 
         Caption         =   "from each category"
         Height          =   255
         Index           =   9
         Left            =   1320
         TabIndex        =   49
         Top             =   510
         Width           =   1395
      End
   End
   Begin VB.Frame generalFrame 
      Height          =   855
      Index           =   4
      Left            =   30
      TabIndex        =   36
      Top             =   540
      Width           =   5955
      Begin VB.CommandButton btnSuggest 
         Caption         =   "Suggest"
         Height          =   555
         Left            =   120
         TabIndex        =   40
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox numDays 
         Height          =   255
         Left            =   2130
         TabIndex        =   39
         Top             =   330
         Width           =   585
      End
      Begin VB.TextBox numToPopulate 
         Height          =   255
         Left            =   3960
         TabIndex        =   38
         Top             =   180
         Width           =   585
      End
      Begin VB.TextBox numNewProd 
         Height          =   255
         Left            =   3960
         TabIndex        =   37
         Top             =   510
         Width           =   585
      End
      Begin VB.Label generalLabel 
         Caption         =   "based on last"
         Height          =   255
         Index           =   2
         Left            =   1140
         TabIndex        =   46
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label generalLabel 
         Caption         =   "days"
         Height          =   255
         Index           =   3
         Left            =   2790
         TabIndex        =   45
         Top             =   330
         Width           =   525
      End
      Begin VB.Label generalLabel 
         Caption         =   "see top"
         Height          =   255
         Index           =   11
         Left            =   3330
         TabIndex        =   44
         Top             =   180
         Width           =   705
      End
      Begin VB.Label generalLabel 
         Caption         =   "for each category"
         Height          =   255
         Index           =   12
         Left            =   4620
         TabIndex        =   43
         Top             =   210
         Width           =   1275
      End
      Begin VB.Label generalLabel 
         Caption         =   "see last"
         Height          =   255
         Index           =   14
         Left            =   3330
         TabIndex        =   42
         Top             =   510
         Width           =   585
      End
      Begin VB.Label generalLabel 
         Caption         =   "new products"
         Height          =   255
         Index           =   15
         Left            =   4620
         TabIndex        =   41
         Top             =   540
         Width           =   1185
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   7140
      TabIndex        =   34
      Top             =   0
      Width           =   1785
   End
   Begin VB.Frame generalFrame 
      Height          =   1875
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   7560
      Width           =   8655
      Begin VB.CommandButton btnMoveTop 
         Caption         =   "Top"
         Height          =   315
         Index           =   3
         Left            =   6300
         TabIndex        =   59
         Top             =   1470
         Width           =   465
      End
      Begin VB.ComboBox cmbColName 
         Height          =   315
         Index           =   3
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   30
         Width           =   3075
      End
      Begin VB.CommandButton btnMoveUp 
         Caption         =   "Up"
         Height          =   315
         Index           =   3
         Left            =   6780
         TabIndex        =   33
         Top             =   1470
         Width           =   465
      End
      Begin VB.CommandButton btnMoveDown 
         Caption         =   "Down"
         Height          =   315
         Index           =   3
         Left            =   7260
         TabIndex        =   32
         Top             =   1470
         Width           =   555
      End
      Begin VB.ListBox fpColumn 
         Height          =   1035
         Index           =   3
         Left            =   90
         TabIndex        =   20
         Top             =   360
         Width           =   8475
      End
      Begin VB.ComboBox cmbAddItem 
         Height          =   315
         Index           =   3
         Left            =   3300
         TabIndex        =   19
         Top             =   1470
         Width           =   2295
      End
      Begin VB.CommandButton btnAddItem 
         Caption         =   "Add"
         Height          =   315
         Index           =   3
         Left            =   5610
         TabIndex        =   18
         Top             =   1470
         Width           =   585
      End
      Begin VB.CommandButton btnRemoveItem 
         Caption         =   "Delete"
         Height          =   315
         Index           =   3
         Left            =   7860
         TabIndex        =   17
         Top             =   1470
         Width           =   705
      End
   End
   Begin VB.Frame generalFrame 
      Height          =   1875
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   5640
      Width           =   8655
      Begin VB.CommandButton btnMoveTop 
         Caption         =   "Top"
         Height          =   315
         Index           =   2
         Left            =   6300
         TabIndex        =   58
         Top             =   1470
         Width           =   465
      End
      Begin VB.ComboBox cmbColName 
         Height          =   315
         Index           =   2
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   30
         Width           =   3075
      End
      Begin VB.CommandButton btnMoveUp 
         Caption         =   "Up"
         Height          =   315
         Index           =   2
         Left            =   6780
         TabIndex        =   31
         Top             =   1470
         Width           =   465
      End
      Begin VB.CommandButton btnMoveDown 
         Caption         =   "Down"
         Height          =   315
         Index           =   2
         Left            =   7260
         TabIndex        =   30
         Top             =   1470
         Width           =   555
      End
      Begin VB.ListBox fpColumn 
         Height          =   1035
         Index           =   2
         Left            =   90
         TabIndex        =   15
         Top             =   360
         Width           =   8475
      End
      Begin VB.ComboBox cmbAddItem 
         Height          =   315
         Index           =   2
         Left            =   3300
         TabIndex        =   14
         Top             =   1470
         Width           =   2295
      End
      Begin VB.CommandButton btnAddItem 
         Caption         =   "Add"
         Height          =   315
         Index           =   2
         Left            =   5610
         TabIndex        =   13
         Top             =   1470
         Width           =   585
      End
      Begin VB.CommandButton btnRemoveItem 
         Caption         =   "Delete"
         Height          =   315
         Index           =   2
         Left            =   7860
         TabIndex        =   12
         Top             =   1470
         Width           =   705
      End
   End
   Begin VB.Frame generalFrame 
      Height          =   1875
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   8655
      Begin VB.CommandButton btnMoveTop 
         Caption         =   "Top"
         Height          =   315
         Index           =   1
         Left            =   6300
         TabIndex        =   57
         Top             =   1470
         Width           =   465
      End
      Begin VB.ComboBox cmbColName 
         Height          =   315
         Index           =   1
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   30
         Width           =   3075
      End
      Begin VB.CommandButton btnMoveUp 
         Caption         =   "Up"
         Height          =   315
         Index           =   1
         Left            =   6780
         TabIndex        =   29
         Top             =   1470
         Width           =   465
      End
      Begin VB.CommandButton btnMoveDown 
         Caption         =   "Down"
         Height          =   315
         Index           =   1
         Left            =   7260
         TabIndex        =   28
         Top             =   1470
         Width           =   555
      End
      Begin VB.ListBox fpColumn 
         Height          =   1035
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   360
         Width           =   8475
      End
      Begin VB.ComboBox cmbAddItem 
         Height          =   315
         Index           =   1
         Left            =   3300
         TabIndex        =   9
         Top             =   1470
         Width           =   2295
      End
      Begin VB.CommandButton btnAddItem 
         Caption         =   "Add"
         Height          =   315
         Index           =   1
         Left            =   5610
         TabIndex        =   8
         Top             =   1470
         Width           =   585
      End
      Begin VB.CommandButton btnRemoveItem 
         Caption         =   "Delete"
         Height          =   315
         Index           =   1
         Left            =   7860
         TabIndex        =   7
         Top             =   1470
         Width           =   705
      End
   End
   Begin VB.Frame generalFrame 
      Height          =   1875
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   8655
      Begin VB.CommandButton btnMoveTop 
         Caption         =   "Top"
         Height          =   315
         Index           =   0
         Left            =   6300
         TabIndex        =   56
         Top             =   1470
         Width           =   465
      End
      Begin VB.ComboBox cmbColName 
         Height          =   315
         Index           =   0
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   30
         Width           =   3075
      End
      Begin VB.CommandButton btnMoveDown 
         Caption         =   "Down"
         Height          =   315
         Index           =   0
         Left            =   7260
         TabIndex        =   27
         Top             =   1470
         Width           =   555
      End
      Begin VB.CommandButton btnMoveUp 
         Caption         =   "Up"
         Height          =   315
         Index           =   0
         Left            =   6780
         TabIndex        =   26
         Top             =   1470
         Width           =   465
      End
      Begin VB.CommandButton btnRemoveItem 
         Caption         =   "Delete"
         Height          =   315
         Index           =   0
         Left            =   7860
         TabIndex        =   5
         Top             =   1470
         Width           =   705
      End
      Begin VB.CommandButton btnAddItem 
         Caption         =   "Add"
         Height          =   315
         Index           =   0
         Left            =   5610
         TabIndex        =   4
         Top             =   1470
         Width           =   585
      End
      Begin VB.ComboBox cmbAddItem 
         Height          =   315
         Index           =   0
         Left            =   3300
         TabIndex        =   3
         Top             =   1470
         Width           =   2295
      End
      Begin VB.ListBox fpColumn 
         Height          =   1035
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   360
         Width           =   8475
      End
   End
   Begin VB.Label generalLabel 
      Caption         =   "QOH"
      Height          =   195
      Index           =   10
      Left            =   8070
      TabIndex        =   35
      Top             =   1560
      Width           =   585
   End
   Begin VB.Label generalLabel 
      Caption         =   "# Orders"
      Height          =   195
      Index           =   7
      Left            =   7110
      TabIndex        =   25
      Top             =   1560
      Width           =   705
   End
   Begin VB.Label generalLabel 
      Caption         =   "Total Qty Sold"
      Height          =   195
      Index           =   6
      Left            =   5880
      TabIndex        =   24
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label generalLabel 
      Caption         =   "Price"
      Height          =   195
      Index           =   5
      Left            =   5160
      TabIndex        =   23
      Top             =   1560
      Width           =   585
   End
   Begin VB.Label generalLabel 
      Caption         =   "Description"
      Height          =   195
      Index           =   4
      Left            =   1860
      TabIndex        =   22
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label generalLabel 
      Caption         =   "Item"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   21
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label generalLabel 
      Caption         =   "Web Site Front Page Maintenance"
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
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4395
   End
End
Attribute VB_Name = "WebSiteFrontPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DEFAULT_SUGG_DAYS As String = "14"
Private Const DEFAULT_NUM_TO_POP As String = "10"
Private Const DEFAULT_NEW_PRODS As String = "50"
Private Const DEFAULT_ITEMS_PER_CAT As String = "4"

Private fillingform As Boolean

Private Sub btnAddItem_Click(Index As Integer)
    If Me.cmbAddItem(Index) <> "" Then
        AddToList Me.cmbAddItem(Index), Me.fpColumn(Index)
        Me.cmbAddItem(Index) = ""
    End If
End Sub

Private Sub btnGenerate_Click()
    If IsFormLoaded("WebSiteFrontPageDlg") Then
        Unload WebSiteFrontPageDlg
    End If
    Load WebSiteFrontPageDlg
    Dim i As Long, j As Long, thisln As String
    ReDim thiscol(0) As Variant
    For i = Me.fpColumn.LBound To Me.fpColumn.UBound
        WebSiteFrontPageDlg.heading(i) = Me.cmbColName(i) & ":"
        WebSiteFrontPageDlg.items(i) = ""
        WebSiteFrontPageDlg.items(i).tag = ""
        For j = 0 To Me.numInCategory - 1
            If j >= Me.fpColumn(i).ListCount Then
                j = Me.numInCategory - 1
            Else
                thisln = Me.fpColumn(i).list(j)
                thisln = Left(thisln, InStr(thisln, vbTab) - 1)
                WebSiteFrontPageDlg.items(i) = WebSiteFrontPageDlg.items(i) & LCase(CreateYahooID(thisln)) & " "
                WebSiteFrontPageDlg.items(i).tag = WebSiteFrontPageDlg.items(i).tag & LCase(thisln) & " "
            End If
        Next j
        If WebSiteFrontPageDlg.items(i) <> "" Then
            WebSiteFrontPageDlg.items(i) = Trim(WebSiteFrontPageDlg.items(i))
            WebSiteFrontPageDlg.items(i).tag = Trim(WebSiteFrontPageDlg.items(i).tag)
        End If
    Next i
    WebSiteFrontPageDlg.Show
End Sub

Private Sub btnMoveDown_Click(Index As Integer)
    ListBoxMoveDown Me.fpColumn(Index)
End Sub

Private Sub btnMoveTop_Click(Index As Integer)
    While ListBoxMoveUp(Me.fpColumn(Index))
    Wend
End Sub

Private Sub btnMoveUp_Click(Index As Integer)
    ListBoxMoveUp Me.fpColumn(Index)
End Sub

Private Sub btnRemoveItem_Click(Index As Integer)
    If Me.fpColumn(Index).SelCount = 1 Then
        If vbYes = MsgBox("Remove " & Left(Me.fpColumn(Index), InStr(Me.fpColumn(Index), vbTab) - 1) & " from the front page?", vbYesNo) Then
            RemoveSelectedFrom Me.fpColumn(Index)
        End If
    End If
End Sub

Private Sub cmbAddItem_Click(Index As Integer)
    cmbAddItem_LostFocus Index
End Sub

Private Sub cmbAddItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.cmbAddItem(Index), KeyCode, Shift
    If KeyCode = vbKeyReturn Then
        btnAddItem_Click Index
    End If
End Sub

Private Sub cmbAddItem_KeyPress(Index As Integer, KeyAscii As Integer)
    AutoCompleteKeyPress Me.cmbAddItem(Index), KeyAscii
End Sub

Private Sub cmbAddItem_LostFocus(Index As Integer)
    AutoCompleteLostFocus Me.cmbAddItem(Index)
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnSuggest_Click()
    Dim i As Long
    Dim rst As ADODB.Recordset, rst2 As ADODB.Recordset
    For i = Me.fpColumn.LBound To Me.fpColumn.UBound
        Me.fpColumn(i).Clear
        If Me.cmbColName(i) <> "" Then
            Set rst = DB.retrieve("SELECT SuggProcedure FROM WebFrontPage WHERE ColName='" & EscapeSQuotes(Me.cmbColName(i)) & "'")
            If rst.EOF Then
                MsgBox "Unable to find column " & qq(Me.cmbColName(i)) & " in the table, how did that happen?"
            Else
                If Nz(rst("SuggProcedure")) <> "" Then
                    'ok, we have a suggest procedure, what now
                    'need a consistent interface (params, etc)
                    ' - days history
                    ' - # of rows
                    Set rst2 = DB.retrieve("EXEC spWebFrontPage" & rst("SuggProcedure") & " " & Me.numDays & ", " & Me.numToPopulate)
                    While Not rst2.EOF
                        AddToList rst2("ItemNumber"), Me.fpColumn(i), rst2
                        rst2.MoveNext
                    Wend
                    rst2.Close
                End If
            End If
            rst.Close
        End If
    Next i
    Set rst = Nothing
    Set rst2 = Nothing
End Sub

Private Sub cmbColName_Click(Index As Integer)
    If Not fillingform Then
        'dbExecute "UPDATE WebFrontPage SET CurrentPos=-1 WHERE CurrentPos=" & Index
        'dbExecute "UPDATE WebFrontPage SET CurrentPos=" & Index & " WHERE ColName='" & EscapeSQuotes(Me.cmbColName(Index)) & "'"
    End If
End Sub

Private Sub Form_Load()
    Me.numDays = DEFAULT_SUGG_DAYS
    Me.numToPopulate = DEFAULT_NUM_TO_POP
    Me.numNewProd = DEFAULT_NEW_PRODS
    Me.numInCategory = DEFAULT_ITEMS_PER_CAT
    Dim tabs(4) As Long
    tabs(0) = 78                'item
    tabs(1) = tabs(0) + 142     'desc
    tabs(2) = tabs(1) + 42      'price
    tabs(3) = tabs(2) + 42      'qty
    tabs(4) = tabs(3) + 42      'orders
    Dim i As Long
    For i = Me.fpColumn.LBound To Me.fpColumn.UBound
        SetListTabStops Me.fpColumn(i).hwnd, tabs
    Next i
    requeryItems
    requeryColumns
    requeryLastItems
End Sub

Private Sub numDays_LostFocus()
    If Me.numDays = "" Or Not IsNumeric(Me.numDays) Then
        Me.numDays = DEFAULT_SUGG_DAYS
    End If
End Sub

Private Sub numInCategory_LostFocus()
    If Me.numInCategory = "" Or Not IsNumeric(Me.numInCategory) Then
        Me.numInCategory = DEFAULT_ITEMS_PER_CAT
    End If
End Sub

Private Sub numNewProd_LostFocus()
    If Me.numNewProd = "" Or Not IsNumeric(Me.numNewProd) Then
        Me.numNewProd = DEFAULT_NEW_PRODS
    End If
End Sub

Private Sub numToPopulate_LostFocus()
    If Me.numToPopulate = "" Or Not IsNumeric(Me.numToPopulate) Then
        Me.numToPopulate = DEFAULT_NUM_TO_POP
    End If
End Sub


Private Sub requeryItems()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT PartNumbers.ItemNumber FROM PartNumbers INNER JOIN vItemStatusBreakdown ON PartNumbers.ItemNumber=vItemStatusBreakdown.ItemNumber WHERE WebPublished=1 AND vItemStatusBreakdown.DC=0")
    Dim i As Long
    For i = Me.cmbAddItem.LBound To Me.cmbAddItem.UBound
        Me.cmbAddItem(i).Clear
    Next i
    While Not rst.EOF
        For i = Me.cmbAddItem.LBound To Me.cmbAddItem.UBound
            Me.cmbAddItem(i).AddItem rst("ItemNumber")
        Next i
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryColumns()
    fillingform = True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ColName, CurrentPos FROM WebFrontPage ORDER BY ID")
    Dim i As Long
    While Not rst.EOF
        For i = Me.fpColumn.LBound To Me.fpColumn.UBound
            Me.cmbColName(i).AddItem rst("ColName")
        Next i
        If rst("CurrentPos") >= Me.cmbColName.LBound And rst("CurrentPos") <= Me.cmbColName.UBound Then
            Me.cmbColName(rst("CurrentPos")) = rst("ColName")
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    fillingform = False
End Sub

Private Sub requeryLastItems()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT WhichCol, WhichPos, ItemNumber FROM WebFrontPageItems ORDER BY WhichCol, WhichPos")
    While Not rst.EOF
        AddToList rst("ItemNumber"), Me.fpColumn(rst("WhichCol"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub


Private Sub AddToList(itemNumber As String, ctl As ListBox, Optional rst As ADODB.Recordset = Nothing)
    If itemNumber <> "" Then
        'If MsgBox("Add " & ItemNumber & " to the front page?", vbYesNo) = vbYes Then
            Dim closerst As Boolean
            If rst Is Nothing Then
                closerst = True
                Set rst = DB.retrieve("exec spWebFrontPageSingleItem '" & itemNumber & "', " & Me.numDays)
            End If
            If Not rst.EOF Then
                ctl.AddItem rst("ItemNumber") & vbTab & _
                            rst("ItemDescription") & vbTab & _
                            rst("IDiscountMarkupPriceRate1") & vbTab & _
                            Nz(rst("QtySold")) & vbTab & _
                            rst("NumTransactions") & vbTab & _
                            rst("QtyOnHand")
            End If
            If closerst Then
                rst.Close
                Set rst = Nothing
            End If
        'End If
    End If
End Sub
