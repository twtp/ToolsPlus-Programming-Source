VERSION 5.00
Begin VB.Form VendorDropshipChargeCfg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendor Dropship Charge Config"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   3840
      TabIndex        =   26
      Top             =   4020
      Width           =   705
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel Changes"
      Height          =   435
      Left            =   2100
      TabIndex        =   25
      Top             =   4020
      Width           =   1335
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save Changes"
      Height          =   435
      Left            =   600
      TabIndex        =   24
      Top             =   4020
      Width           =   1335
   End
   Begin VB.TextBox cost 
      Height          =   285
      Index           =   5
      Left            =   3180
      TabIndex        =   23
      Top             =   3450
      Width           =   1005
   End
   Begin VB.TextBox high 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3450
      Width           =   1005
   End
   Begin VB.TextBox low 
      Height          =   285
      Index           =   5
      Left            =   420
      TabIndex        =   21
      Top             =   3450
      Width           =   1005
   End
   Begin VB.TextBox cost 
      Height          =   285
      Index           =   4
      Left            =   3180
      TabIndex        =   20
      Top             =   3120
      Width           =   1005
   End
   Begin VB.TextBox high 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3120
      Width           =   1005
   End
   Begin VB.TextBox low 
      Height          =   285
      Index           =   4
      Left            =   420
      TabIndex        =   18
      Top             =   3120
      Width           =   1005
   End
   Begin VB.TextBox cost 
      Height          =   285
      Index           =   3
      Left            =   3180
      TabIndex        =   17
      Top             =   2790
      Width           =   1005
   End
   Begin VB.TextBox high 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2790
      Width           =   1005
   End
   Begin VB.TextBox low 
      Height          =   285
      Index           =   3
      Left            =   420
      TabIndex        =   15
      Top             =   2790
      Width           =   1005
   End
   Begin VB.TextBox cost 
      Height          =   285
      Index           =   2
      Left            =   3180
      TabIndex        =   14
      Top             =   2460
      Width           =   1005
   End
   Begin VB.TextBox high 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2460
      Width           =   1005
   End
   Begin VB.TextBox low 
      Height          =   285
      Index           =   2
      Left            =   420
      TabIndex        =   12
      Top             =   2460
      Width           =   1005
   End
   Begin VB.TextBox cost 
      Height          =   285
      Index           =   1
      Left            =   3180
      TabIndex        =   11
      Top             =   2130
      Width           =   1005
   End
   Begin VB.TextBox high 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2130
      Width           =   1005
   End
   Begin VB.TextBox low 
      Height          =   285
      Index           =   1
      Left            =   420
      TabIndex        =   9
      Top             =   2130
      Width           =   1005
   End
   Begin VB.TextBox cost 
      Height          =   285
      Index           =   0
      Left            =   3180
      TabIndex        =   7
      Top             =   1800
      Width           =   1005
   End
   Begin VB.TextBox high 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   1005
   End
   Begin VB.TextBox low 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   420
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "$0"
      Top             =   1800
      Width           =   1005
   End
   Begin VB.ComboBox LineCode 
      Height          =   315
      Left            =   2220
      TabIndex        =   1
      Top             =   870
      Width           =   1425
   End
   Begin VB.Label Label5 
      Caption         =   "Cost:"
      Height          =   225
      Left            =   3270
      TabIndex        =   8
      Top             =   1530
      Width           =   555
   End
   Begin VB.Label Label4 
      Caption         =   "To:"
      Height          =   225
      Left            =   1590
      TabIndex        =   6
      Top             =   1530
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "From:"
      Height          =   225
      Left            =   510
      TabIndex        =   5
      Top             =   1530
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Select Line Code:"
      Height          =   255
      Left            =   810
      TabIndex        =   2
      Top             =   900
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Vendor Dropship Charge Config"
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
      Left            =   330
      TabIndex        =   0
      Top             =   60
      Width           =   3945
   End
End
Attribute VB_Name = "VendorDropshipChargeCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FILE_DIR = "s:\mastest\importscripts\"
Private Const FILE_NAME = "dscosts.txt"

Private Sub btnCancel_Click()
    fillForm
End Sub

Private Sub btnExit_Click()
    Unload VendorDropshipChargeCfg
End Sub

Private Sub btnSave_Click()
    Dim foo As Integer
    foo = sanityCheck()
    If foo = 0 Then
        saveToFile
    ElseIf foo = 1 Then
        MsgBox "Invalid data, from's must go in ascending order"
    ElseIf foo = 2 Then
        MsgBox "Invalid data, you are missing a cost"
    End If
End Sub

Private Sub cost_LostFocus(Index As Integer)
    If Me.cost(Index) = "" Then
        'do nothing
    ElseIf validateCurrency(Me.low(Index)) Then
        Me.cost(Index) = Format(Me.cost(Index), "Currency")
    Else
        MsgBox "error, must be a valid number"
    End If
End Sub

Private Sub Form_Load()
    Set Mouse = New MouseHourglass

    requeryLineCodes
End Sub

Private Sub requeryLineCodes()
    Dim rst As ADODB.Recordset
    Set rst = retrieve("SELECT ProductLine FROM vRequeryLineCodes ORDER BY ProductLine")
    Me.LineCode.Clear
    While Not rst.EOF
        Me.LineCode.AddItem rst("ProductLine")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub fillForm()
    Open FILE_DIR & FILE_NAME For Input As #1
    Dim ln As String, found As Boolean
    While Not found And Not EOF(1)
        Line Input #1, ln
        If Left(ln, 5) = Me.LineCode & ": " Then
            found = True
        End If
    Wend
    Close #1
    If Not found Then
        clearForm
    Else
        Dim foo As Variant
        foo = Split(Mid(ln, 6), " ")
        fillItOut foo
    End If
End Sub

Private Sub fillItOut(hashpairs As Variant)
    clearForm
    Dim i As Integer
    For i = 0 To UBound(hashpairs)
        Dim key As String, value As String
        key = Left(hashpairs(i), InStr(hashpairs(i), "=") - 1)
        value = Mid(hashpairs(i), InStr(hashpairs(i), "=") + 1)
        If i <> 0 Then
            Me.low(i) = Format(Me.high(i - 1) + 0.01, "Currency")
        End If
        If key = 2 ^ 31 Then
            Me.high(i) = "and up"
        Else
            Me.high(i) = Format(key - 0.01, "Currency")
        End If
        Me.cost(i) = Format(value, "Currency")
    Next i
End Sub

Private Sub LineCode_Click()
    LineCode_LostFocus
End Sub

Private Sub LineCode_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.LineCode, KeyCode, Shift
End Sub

Private Sub LineCode_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.LineCode, KeyAscii
End Sub

Private Sub LineCode_LostFocus()
    AutoCompleteLostFocus Me.LineCode
    fillForm
End Sub

Private Sub clearForm()
    Dim i As Integer
    For i = 0 To 5
        Me.low(i) = ""
        Me.high(i) = ""
        Me.cost(i) = ""
    Next i
    Me.low(0) = "$0.00"
    Me.high(0) = "and up"
End Sub

Private Sub low_LostFocus(Index As Integer)
    If Me.low(Index) = "" Then
        'do nothing
    ElseIf validateCurrency(Me.low(Index)) Then
        Me.low(Index) = Format(Me.low(Index), "Currency")
        Me.high(Index - 1) = Format(Me.low(Index) - 0.01, "Currency")
        If Me.high(Index) = "" Then
            Me.high(Index) = "and up"
        End If
    Else
        MsgBox "error, must be a valid number"
    End If
End Sub

Private Function sanityCheck() As Integer
    Dim i As Integer
    sanityCheck = 0
    For i = 1 To Me.low.UBound
        If Me.low(i) = "" Then
            i = Me.low.UBound
        ElseIf CInt(Me.low(i)) < CInt(Me.low(i - 1)) Then
            sanityCheck = 1
            Exit Function
        ElseIf Me.cost(i) = "" Then
            sanityCheck = 2
            Exit Function
        End If
    Next i
End Function

Private Sub saveToFile()
    Open FILE_DIR & FILE_NAME For Input As #1
    Dim contents As PerlArray
    Set contents = New PerlArray
    Dim foo As String
    While Not EOF(1)
        Line Input #1, foo
        If Left(foo, 3) <> Me.LineCode Then
            contents.Push foo
        End If
    Wend
    Close #1
    foo = ""
    Dim i As Integer
    For i = 1 To Me.low.UBound
        If Me.low(i) = "" Then
            If foo <> "" Then
                foo = foo & " " & 2 ^ 31 & "=" & Replace(Replace(Me.cost(i - 1), ",", ""), "$", "")
            End If
            i = Me.low.UBound
        Else
            foo = foo & " " & Replace(Replace(Me.low(i), ",", ""), "$", "") & "=" & Replace(Replace(Me.cost(i - 1), ",", ""), "$", "")
        End If
    Next i
    
    Open FILE_DIR & FILE_NAME For Output As #1
    If foo <> "" Then
        Print #1, Me.LineCode & ":" & foo
    End If
    For i = 0 To contents.Scalar - 1
        Print #1, contents.Elem(i)
    Next i
    Close #1
End Sub
