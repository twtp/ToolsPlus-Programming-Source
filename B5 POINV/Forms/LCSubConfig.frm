VERSION 5.00
Begin VB.Form LCSubConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Line Code Substitution Config"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox tableName 
      Height          =   315
      Left            =   3420
      TabIndex        =   21
      Text            =   "hidden table name"
      Top             =   510
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   20
      Top             =   4470
      Width           =   1365
   End
   Begin VB.VScrollBar lcScrollBar 
      Height          =   2715
      LargeChange     =   8
      Left            =   4410
      TabIndex        =   19
      Top             =   1620
      Width           =   285
   End
   Begin VB.TextBox realLC 
      Height          =   285
      Index           =   7
      Left            =   2640
      TabIndex        =   18
      Top             =   3960
      Width           =   1635
   End
   Begin VB.TextBox realLC 
      Height          =   285
      Index           =   6
      Left            =   2640
      TabIndex        =   17
      Top             =   3630
      Width           =   1635
   End
   Begin VB.TextBox realLC 
      Height          =   285
      Index           =   5
      Left            =   2640
      TabIndex        =   16
      Top             =   3300
      Width           =   1635
   End
   Begin VB.TextBox realLC 
      Height          =   285
      Index           =   4
      Left            =   2640
      TabIndex        =   15
      Top             =   2970
      Width           =   1635
   End
   Begin VB.TextBox realLC 
      Height          =   285
      Index           =   3
      Left            =   2640
      TabIndex        =   14
      Top             =   2640
      Width           =   1635
   End
   Begin VB.TextBox realLC 
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   13
      Top             =   2310
      Width           =   1635
   End
   Begin VB.TextBox realLC 
      Height          =   285
      Index           =   1
      Left            =   2640
      TabIndex        =   12
      Top             =   1980
      Width           =   1635
   End
   Begin VB.TextBox alternateLC 
      Height          =   285
      Index           =   7
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3960
      Width           =   1635
   End
   Begin VB.TextBox alternateLC 
      Height          =   285
      Index           =   6
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3630
      Width           =   1635
   End
   Begin VB.TextBox alternateLC 
      Height          =   285
      Index           =   5
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3300
      Width           =   1635
   End
   Begin VB.TextBox alternateLC 
      Height          =   285
      Index           =   4
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2970
      Width           =   1635
   End
   Begin VB.TextBox alternateLC 
      Height          =   285
      Index           =   3
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2640
      Width           =   1635
   End
   Begin VB.TextBox alternateLC 
      Height          =   285
      Index           =   2
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2310
      Width           =   1635
   End
   Begin VB.TextBox alternateLC 
      Height          =   285
      Index           =   1
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1980
      Width           =   1635
   End
   Begin VB.TextBox realLC 
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   4
      Top             =   1650
      Width           =   1635
   End
   Begin VB.TextBox alternateLC 
      Height          =   285
      Index           =   0
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1650
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "remember that they all point to a common line code, so POW/POA/PER => JET, but JET stays as JET"
      Height          =   495
      Left            =   150
      TabIndex        =   22
      Top             =   510
      Width           =   3795
   End
   Begin VB.Line Line2 
      X1              =   300
      X2              =   4680
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Line Line1 
      X1              =   270
      X2              =   4680
      Y1              =   4350
      Y2              =   4350
   End
   Begin VB.Label generalLabel 
      Caption         =   "gets substituted with this:"
      Height          =   315
      Index           =   2
      Left            =   2490
      TabIndex        =   2
      Top             =   1140
      Width           =   2025
   End
   Begin VB.Label generalLabel 
      Caption         =   "This line code:"
      Height          =   285
      Index           =   1
      Left            =   390
      TabIndex        =   1
      Top             =   1140
      Width           =   1695
   End
   Begin VB.Label generalLabel 
      Caption         =   "Line Code Substitution Config"
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
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   3945
   End
End
Attribute VB_Name = "LCSubConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NUM_BOXES As Long = 8

'
' ============================================================
'  NOTE: this module uses the old product line prefix fields
'  for reading, but updates both.
' ============================================================
'

Dim changed As Boolean
Dim currentPos As Long
Dim lcList() As String

Private Sub btnExit_Click()
    Unload LCSubConfig
End Sub

Private Sub tableName_Change()
    requeryLCs
    fillLCs
End Sub




Private Sub requeryLCs()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM vRequeryLineCodes ORDER BY ProductLine")
    If rst.RecordCount > 0 Then
        ReDim lcList(rst.RecordCount - 1) As String
        Dim i As Long
        For i = 0 To rst.RecordCount - 1
            lcList(i) = rst("ProductLine")
            rst.MoveNext
        Next i
        Me.lcScrollBar.Min = 0
        Me.lcScrollBar.max = IIf(rst.RecordCount - NUM_BOXES < 0, 0, rst.RecordCount - NUM_BOXES)
    Else
        ReDim ItemList(0) As String
        ItemList(0) = "NORECORDS"
        Me.lcScrollBar.Min = 0
        Me.lcScrollBar.max = 0
    End If
    currentPos = 0
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub fillLCs(Optional offset As Long = 0)
    Dim rst As ADODB.Recordset
    Dim i As Long
    For i = 0 To NUM_BOXES - 1
        If UBound(lcList) >= i Then
            If lcList(offset + i) <> "NORECORDS" Then
                Set rst = DB.retrieve("SELECT AlternateLineCode, RealLineCode FROM " & Me.tableName & " WHERE AlternateLineCode='" & lcList(offset + i) & "'")
                Me.alternateLC(i).Visible = True
                Me.alternateLC(i) = lcList(offset + i)
                Me.realLC(i).Visible = True
                If rst.EOF Then
                    Me.realLC(i) = lcList(offset + i)
                Else
                    Me.realLC(i) = rst("RealLineCode")
                End If
                rst.Close
            Else
                While i < NUM_BOXES
                    Me.alternateLC(i).Visible = False
                    Me.realLC(i).Visible = False
                    i = i + 1
                Wend
            End If
        Else
            While i < NUM_BOXES
                Me.alternateLC(i).Visible = False
                Me.realLC(i).Visible = False
                i = i + 1
            Wend
        End If
    Next i
End Sub

Private Sub lcScrollBar_Change()
    Mouse.Hourglass True
    If lcList(0) <> "NORECORDS" Then
        If Abs(Me.lcScrollBar.value - currentPos) > 1 Then
            'multiple move, must redo entire form
            fillLCs Me.lcScrollBar.value
        Else
            'only a single up or down, shift things around
            If Me.lcScrollBar.value > currentPos Then 'down arrow
                shuffleUp
                addLCAfter Me.alternateLC(NUM_BOXES - 2)
            Else 'uparrow
                shuffleDown
                addLCBefore Me.alternateLC(1)
            End If
        End If
        currentPos = Me.lcScrollBar.value
    End If
    doHighlighting
    Mouse.Hourglass False
End Sub

Private Sub realLC_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            realLC_LostFocus Index
        Case Is = vbKeyDelete
            realLC_KeyPress Index, KeyCode
    End Select
End Sub

Private Sub realLC_KeyPress(Index As Integer, KeyAscii As Integer)
    changed = True
End Sub

Private Sub realLC_LostFocus(Index As Integer)
    If changed = True Then
        If Me.realLC(Index) = "" Then
            Me.realLC(Index) = Me.alternateLC(Index)
        End If
        Me.realLC(Index) = UCase(Me.realLC(Index))
        
        Dim altID As String, realID As String
        altID = DLookup("ID", "ProductLine", "ProductLine='" & Me.alternateLC(Index) & "'")
        realID = DLookup("ID", "ProductLine", "ProductLine='" & Me.realLC(Index) & "'")
        
        If PLExists(Me.realLC(Index)) Then
            If Me.realLC(Index) = Me.alternateLC(Index) Then
                DB.Execute "DELETE FROM " & Me.tableName & " WHERE ProductLineID=" & altID
            Else
                If DLookup("ProductLineID", Me.tableName, "ProductLineID=" & altID) = "" Then
                    DB.Execute "INSERT INTO " & Me.tableName & " ( ProductLineID, PurchasingProductLineID, AlternateLineCode, RealLineCode ) VALUES ( " & altID & ", " & realID & ", '" & Me.alternateLC(Index) & "', '" & Me.realLC(Index) & "' )"
                Else
                    DB.Execute "UPDATE " & Me.tableName & " SET PurchasingProductLineID=" & realID & ", RealLineCode='" & Me.realLC(Index) & "' WHERE ProductLineID=" & altID
                End If
            End If
        Else
            MsgBox "Invalid Line Code."
            Me.realLC(Index) = Me.alternateLC(Index)
        End If
        changed = False
        doHighlighting
    End If
End Sub

Private Sub shuffleUp()
    Dim i As Long
    For i = 0 To NUM_BOXES - 2
        shuffleLCs i + 1, i
    Next i
End Sub

Private Sub shuffleDown()
    Dim i As Long
    For i = NUM_BOXES - 1 To 1 Step -1
        shuffleLCs i - 1, i
    Next i
End Sub

Private Sub shuffleLCs(fromIndex As Long, toIndex As Long)
    Me.alternateLC(toIndex) = Me.alternateLC(fromIndex)
    Me.realLC(toIndex) = Me.realLC(fromIndex)
End Sub

Private Sub addLCAfter(lc As String)
    Dim i As Long
    i = bsearch(lc, lcList) + 1
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT AlternateLineCode, RealLineCode FROM " & Me.tableName & " WHERE AlternateLineCode='" & lcList(i) & "'")
    Me.alternateLC(NUM_BOXES - 1) = lcList(i)
    If rst.EOF Then
        Me.realLC(NUM_BOXES - 1) = lcList(i)
    Else
        Me.realLC(NUM_BOXES - 1) = rst("RealLineCode")
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub addLCBefore(lc As String)
    Dim i As Long
    i = bsearch(lc, lcList) - 1
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT AlternateLineCode, RealLineCode FROM " & Me.tableName & " WHERE AlternateLineCode='" & lcList(i) & "'")
    Me.alternateLC(0) = lcList(i)
    If rst.EOF Then
        Me.realLC(0) = lcList(i)
    Else
        Me.realLC(0) = rst("RealLineCode")
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Function doHighlighting()
    Dim i As Long
    For i = 0 To NUM_BOXES - 1
        If Me.alternateLC(i) = Me.realLC(i) Then
            Me.alternateLC(i).FontBold = False
            Me.realLC(i).FontBold = False
        Else
            Me.alternateLC(i).FontBold = True
            Me.realLC(i).FontBold = True
        End If
    Next i
End Function
