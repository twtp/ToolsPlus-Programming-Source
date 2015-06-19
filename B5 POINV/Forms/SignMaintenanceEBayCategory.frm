VERSION 5.00
Begin VB.Form SignMaintenanceEBayCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EBay Category"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox HiddenCategoryType 
      Height          =   315
      Left            =   4980
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Cat Type"
      Top             =   120
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox FilterText 
      Height          =   315
      Left            =   1440
      TabIndex        =   8
      Top             =   1320
      Width           =   4395
   End
   Begin VB.TextBox ItemNumber 
      Height          =   285
      Left            =   3630
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "ItemNumber"
      Top             =   120
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   465
      Left            =   4380
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4620
      Width           =   975
   End
   Begin VB.CommandButton btnCategorySet 
      Caption         =   "Set New Category"
      Height          =   435
      Left            =   780
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4650
      Width           =   1695
   End
   Begin VB.ListBox CategoryList 
      Height          =   2595
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1710
      Width           =   10995
   End
   Begin VB.TextBox CurrentCategory 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1470
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   780
      Width           =   9615
   End
   Begin VB.Label Label3 
      Caption         =   "Filter:"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   1350
      Width           =   585
   End
   Begin VB.Label Label2 
      Caption         =   "Current Category:"
      Height          =   285
      Left            =   150
      TabIndex        =   1
      Top             =   810
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "EBay Categories"
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
      Width           =   2895
   End
End
Attribute VB_Name = "SignMaintenanceEBayCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FullList As PerlArray
Private lastFilter As String
Private lastSelectionID As String

Public Sub LoadItem(item As String, categoryType As Long)
    Set FullList = New PerlArray
    Dim iter As Variant
    For Each iter In EBayCategoryHashIDtoName.keys
        If EBayCategoryHashIDtoName.item(CStr(iter))(0) = categoryType Or EBayCategoryHashIDtoName.item(CStr(iter))(0) + 2 = categoryType Then
            Me.CategoryList.AddItem EBayCategoryHashIDtoName.item(CStr(iter))(1) & vbTab & CStr(iter)
            FullList.Push Array(EBayCategoryHashIDtoName.item(CStr(iter))(1), CStr(iter))
        End If
    Next iter
    
    fillList
    
    Me.ItemNumber = item
    Me.HiddenCategoryType = categoryType
    
    Dim rst As ADODB.Recordset
    Select Case categoryType
        Case Is = 0
            Set rst = DB.retrieve("SELECT EBayCategoryID FROM PartNumbers WHERE ItemNumber='" & Me.ItemNumber & "'")
        Case Is = 1
            Set rst = DB.retrieve("SELECT EBayStoreCategoryID FROM PartNumbers WHERE ItemNumber='" & Me.ItemNumber & "'")
        Case Is = 2
            'grab override if present, if not then grab the alternative from primary...
            Set rst = DB.retrieve("SELECT SecondaryCategoryID AS EBayCategoryID FROM EBaySecondaryCategoryOverride WHERE ItemNumber='" & Me.ItemNumber & "'")
            If rst.RecordCount <= 0 Then
                Set rst = DB.retrieve("SELECT dbo.AlternativeCategoryID(EBayCategoryID) AS EBayCategoryID FROM PartNumbers WHERE ItemNumber='" & Me.ItemNumber & "'")
            End If
    End Select
    If Nz(rst(0), -1) = -1 Then
        Me.CurrentCategory = "<none>"
    Else
        'Me.CurrentCategory = EBayCategoryHashIDtoName.item(CStr(rst(0)))
        Me.CurrentCategory = EBayCategoryHashIDtoName.item(CStr(rst(0)))(1)
    End If
End Sub

Private Sub CategoryList_DblClick()
    btnCategorySet_Click
End Sub

Private Sub FilterText_Change()
    lastFilter = Me.FilterText
    fillList
End Sub

Private Sub Form_Load()
    lastFilter = ""
    
    'Dim iter As Variant
    'ReDim FullList(0 To EBayCategoryHashIDtoName.count - 1) As Variant
    'Dim i As Long
    'i = 0
    'For Each iter In EBayCategoryHashIDtoName.keys
    '    Dim cat As String
    '    cat = EBayCategoryHashIDtoName.item(CStr(iter))
    '    'Me.CategoryList.AddItem cat & vbTab & CStr(iter)
    '    FullList(i) = Array(cat, CLng(iter))
    '    i = i + 1
    'Next iter
    '
    'fillList
    
    Dim tabs(0) As Long
    tabs(0) = 600
    SetListTabStops Me.CategoryList.hWnd, tabs
End Sub

Private Sub btnCategorySet_Click()
    If Me.CategoryList.ListIndex = -1 Then
        MsgBox "Select something above!"
    Else
        Dim str As String
        str = ListBoxLineColumn(Me.CategoryList, -1, 0)
        If Left(str, 1) = "$" Then
            If MsgBox("This is an expensive category! Are you sure you want to use it?", vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        Dim id As String
        id = ListBoxLineColumn(Me.CategoryList, -1, 1)
        'DB.Execute "UPDATE PartNumbers SET EBayCategoryID=" & id & " WHERE ItemNumber='" & Me.ItemNumber & "'"
        Select Case Me.HiddenCategoryType
            Case Is = 0
                DB.Execute "UPDATE PartNumbers SET EBayCategoryID=" & id & " WHERE ItemNumber='" & Me.ItemNumber & "'"
            Case Is = 1
                DB.Execute "UPDATE PartNumbers SET EBayStoreCategoryID=" & id & " WHERE ItemNumber='" & Me.ItemNumber & "'"
            Case Is = 2
                Dim hasOverride As ADODB.Recordset
                Set hasOverride = DB.retrieve("SELECT ItemNumber FROM EBaySecondaryCategoryOverride WHERE ItemNumber='" & Me.ItemNumber & "'")
                If (hasOverride.RecordCount > 0) Then
                    DB.Execute "UPDATE EBaySecondaryCategoryOverride SET SecondaryCategoryID=" & id & " WHERE ItemNumber='" & Me.ItemNumber & "'"
                Else
                    DB.Execute "INSERT INTO EBaySecondaryCategoryOverride (ItemNumber,SecondaryCategoryID) VALUES('" & Me.ItemNumber & "'," & id & ")"
                End If
        End Select
        SignMaintenance.EBayCategorySet ListBoxLineColumn(Me.CategoryList), Me.HiddenCategoryType
        Unload Me
    End If
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub fillList()
    If Me.CategoryList.ListIndex <> -1 Then
        lastSelectionID = ListBoxLineColumn(Me.CategoryList, -1, 1)
    Else
        'on no selection, keep the last selection
    End If
    
    Me.CategoryList.Clear
    Dim i As Long
    'For i = 0 To UBound(FullList)
    '    If InStr(1, FullList(i)(0), lastFilter, vbTextCompare) <> 0 Then
    '        Me.CategoryList.AddItem FullList(i)(0) & vbTab & FullList(i)(1)
    '        If FullList(i)(1) = lastSelectionID Then
    '            Me.CategoryList = FullList(i)(0) & vbTab & FullList(i)(1)
    '        End If
    '    End If
    'Next i
    For i = 0 To FullList.Scalar - 1
        If InStr(1, FullList.Elem(i)(0), lastFilter, vbTextCompare) <> 0 Then
            If InStr(1, FullList.Elem(i)(0), "$$$") = 0 Then
                Me.CategoryList.AddItem FullList.Elem(i)(0) & vbTab & FullList.Elem(i)(1)
                If FullList.Elem(i)(1) = lastSelectionID Then
                    Me.CategoryList = FullList.Elem(i)(0) & vbTab & FullList.Elem(i)(1)
                End If
            End If
        End If
    Next i
End Sub
