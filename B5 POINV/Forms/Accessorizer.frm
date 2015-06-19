VERSION 5.00
Begin VB.Form Accessorizer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accessorizer"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   9720
   StartUpPosition =   1  'CenterOwner
   Begin PoInv.ItemFinder AccessoriesFinder 
      Height          =   3855
      Left            =   60
      TabIndex        =   18
      Top             =   5040
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   6800
   End
   Begin PoInv.ItemFinder ItemsToDoFinder 
      Height          =   3915
      Left            =   60
      TabIndex        =   17
      Top             =   750
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   6906
   End
   Begin VB.CommandButton btnFlipper 
      Caption         =   "Flip"
      Height          =   315
      Left            =   6960
      TabIndex        =   16
      Top             =   60
      Width           =   765
   End
   Begin VB.CommandButton btnAccessoriesMoveLeftAll 
      Caption         =   "<< all"
      Height          =   465
      Left            =   7080
      TabIndex        =   15
      Top             =   6630
      Width           =   585
   End
   Begin VB.CommandButton btnItemsMoveLeftAll 
      Caption         =   "<< all"
      Height          =   465
      Left            =   7080
      TabIndex        =   14
      Top             =   2370
      Width           =   585
   End
   Begin VB.CommandButton btnAccessoriesMoveRightAll 
      Caption         =   ">> all"
      Height          =   465
      Left            =   7080
      TabIndex        =   11
      Top             =   6120
      Width           =   585
   End
   Begin VB.CommandButton btnItemsMoveRightAll 
      Caption         =   ">> all"
      Height          =   465
      Left            =   7080
      TabIndex        =   10
      Top             =   1860
      Width           =   585
   End
   Begin VB.CommandButton btnAccessoriesMoveLeft 
      Caption         =   "<<"
      Height          =   465
      Left            =   7080
      TabIndex        =   8
      Top             =   7920
      Width           =   585
   End
   Begin VB.CommandButton btnAccessoriesMoveRight 
      Caption         =   ">>"
      Height          =   465
      Left            =   7080
      TabIndex        =   7
      Top             =   7410
      Width           =   585
   End
   Begin VB.CommandButton btnItemsMoveLeft 
      Caption         =   "<<"
      Height          =   465
      Left            =   7080
      TabIndex        =   6
      Top             =   3660
      Width           =   585
   End
   Begin VB.CommandButton btnItemsMoveRight 
      Caption         =   ">>"
      Height          =   465
      Left            =   7080
      TabIndex        =   5
      Top             =   3150
      Width           =   585
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   3000
      TabIndex        =   4
      Top             =   8940
      Width           =   1395
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "GO!"
      Height          =   435
      Left            =   1410
      TabIndex        =   3
      Top             =   8940
      Width           =   1395
   End
   Begin VB.ListBox Accessories 
      Height          =   3765
      Left            =   7740
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   5010
      Width           =   1875
   End
   Begin VB.ListBox ItemsToDo 
      Height          =   3765
      Left            =   7740
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   750
      Width           =   1875
   End
   Begin VB.Label Label3 
      Caption         =   "Step 2: select accessories to add to each of the items selected above"
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   4740
      Width           =   4305
   End
   Begin VB.Label Label2 
      Caption         =   "Accessorizer"
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
      Left            =   120
      TabIndex        =   12
      Top             =   30
      Width           =   3825
   End
   Begin VB.Label lblStatusBar 
      Height          =   315
      Left            =   4740
      TabIndex        =   9
      Top             =   9060
      Width           =   4545
   End
   Begin VB.Label Label1 
      Caption         =   "Step 1: select items to add accessories to"
      Height          =   225
      Left            =   90
      TabIndex        =   2
      Top             =   510
      Width           =   3225
   End
End
Attribute VB_Name = "Accessorizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Accessories_DblClick()
    btnAccessoriesMoveLeft_Click
End Sub

Private Sub AccessoriesFinder_ListAction()
    btnAccessoriesMoveRight_Click
End Sub

Private Sub btnAccessoriesMoveLeft_Click()
    If Me.Accessories.SelCount <> 0 Then
        Me.Accessories.RemoveItem Me.Accessories.ListIndex
    End If
End Sub

Private Sub btnAccessoriesMoveLeftAll_Click()
    Me.Accessories.Clear
End Sub

Private Sub btnAccessoriesMoveRight_Click()
    Dim item As String
    item = Me.AccessoriesFinder.GetSelectedItem
    If Not InList(item, Me.Accessories, True) Then
        Me.Accessories.AddItem item
    End If
End Sub

Private Sub btnAccessoriesMoveRightAll_Click()
    Dim excDisc As Boolean
    excDisc = MsgBox("Exclude discontinued items?", vbYesNo) = vbYes
    ReDim items(Me.AccessoriesFinder.GetItemListCount(excDisc)) As String
    items = Me.AccessoriesFinder.GetItemList(excDisc)
    Dim i As Long
    For i = 0 To UBound(items)
        If Not InList(items(i), Me.Accessories) Then
            Me.Accessories.AddItem items(i)
        End If
    Next i
End Sub

Private Sub btnExit_Click()
    Unload Accessorizer
End Sub

Private Sub btnFlipper_Click()
    Dim i As Long
    If Me.ItemsToDo.ListCount > 0 Then
        ReDim ItemList(Me.ItemsToDo.ListCount - 1) As String
        ItemList = ListBoxAsArray(Me.ItemsToDo)
        Me.ItemsToDo.Clear
        For i = 0 To Me.Accessories.ListCount - 1
            Me.ItemsToDo.AddItem Me.Accessories.list(i)
        Next i
        Me.Accessories.Clear
        For i = 0 To UBound(ItemList)
            Me.Accessories.AddItem ItemList(i)
        Next i
    Else
        For i = 0 To Me.Accessories.ListCount - 1
            Me.ItemsToDo.AddItem Me.Accessories.list(i)
        Next i
        Me.Accessories.Clear
    End If
End Sub

Private Sub btnItemsMoveLeft_Click()
    If Me.ItemsToDo.SelCount <> 0 Then
        Me.ItemsToDo.RemoveItem Me.ItemsToDo.ListIndex
    End If
End Sub

Private Sub btnItemsMoveLeftAll_Click()
    Me.ItemsToDo.Clear
End Sub

Private Sub btnItemsMoveRight_Click()
    Dim item As String
    Dim warned As Boolean
    item = Me.ItemsToDoFinder.GetSelectedItem
    If Not InList(item, Me.ItemsToDo, True) Then
        If IsRouterBit(item) Then
            If Not warned Then
                MsgBox "Router bits don't get accessories!"
                warned = True
            End If
        Else
            Me.ItemsToDo.AddItem item
        End If
    End If
End Sub

Private Sub btnItemsMoveRightAll_Click()
    Dim excDisc As Boolean
    Dim warned As Boolean
    excDisc = MsgBox("Exclude discontinued items?", vbYesNo) = vbYes
    ReDim items(Me.ItemsToDoFinder.GetItemListCount(excDisc)) As String
    items = Me.ItemsToDoFinder.GetItemList(excDisc)
    Dim i As Long
    For i = 0 To UBound(items)
        If Not InList(items(i), Me.ItemsToDo) Then
            If IsRouterBit(items(i)) Then
                If Not warned Then
                    MsgBox "Router bits don't get accessories!"
                    warned = True
                End If
            Else
                Me.ItemsToDo.AddItem items(i)
            End If
        End If
    Next i
End Sub

Private Sub btnStart_Click()
    If Me.Accessories.ListCount = 0 Or Me.ItemsToDo.ListCount = 0 Then
        Exit Sub
    End If
    Dim i As Long, curritem As String, j As Long, curracc As String
    Dim acclist As Dictionary
    Dim rst As ADODB.Recordset
    For i = 0 To Me.ItemsToDo.ListCount - 1
        curritem = Me.ItemsToDo.list(i)
        Set rst = DB.retrieve("SELECT Accessory FROM PartNumberAccessories WHERE ItemNumber='" & curritem & "'")
        Set acclist = New Dictionary
        While Not rst.EOF
            acclist.Add rst("Accessory").value, "1"
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
        For j = 0 To Me.Accessories.ListCount - 1
            curracc = Me.Accessories.list(j)
            Me.lblStatusBar.caption = "Working on " & curritem & " - " & curracc
            DoEvents
            If Not acclist.exists(curracc) Then
                DB.Execute "INSERT INTO PartNumberAccessories ( ItemNumber, Accessory ) VALUES ( '" & curritem & "', '" & curracc & "' )"
            End If
        Next j
        Set acclist = Nothing
    Next i
    Me.lblStatusBar.caption = ""
    If IsFormLoaded("SignMaintenance") Then
        If InList(SignMaintenance.ItemNumber, Me.ItemsToDo) Then
            SignMaintenance.RefreshForm
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.ItemsToDoFinder.Initialize
    Me.AccessoriesFinder.Initialize
End Sub

Private Sub ItemsToDo_DblClick()
    btnItemsMoveLeft_Click
End Sub

Private Sub ItemsToDoFinder_ListAction()
    btnItemsMoveRight_Click
End Sub
