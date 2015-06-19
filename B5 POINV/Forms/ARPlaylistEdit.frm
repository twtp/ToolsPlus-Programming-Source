VERSION 5.00
Begin VB.Form ARPlaylistEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audio Runner Playlist Maker"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OpenExplorer 
      Caption         =   "E"
      Height          =   255
      Index           =   9
      Left            =   60
      TabIndex        =   71
      ToolTipText     =   "Open this path in Explorer"
      Top             =   4350
      Width           =   225
   End
   Begin VB.CommandButton OpenExplorer 
      Caption         =   "E"
      Height          =   255
      Index           =   8
      Left            =   60
      TabIndex        =   70
      ToolTipText     =   "Open this path in Explorer"
      Top             =   3960
      Width           =   225
   End
   Begin VB.CommandButton OpenExplorer 
      Caption         =   "E"
      Height          =   255
      Index           =   7
      Left            =   60
      TabIndex        =   69
      ToolTipText     =   "Open this path in Explorer"
      Top             =   3570
      Width           =   225
   End
   Begin VB.CommandButton OpenExplorer 
      Caption         =   "E"
      Height          =   255
      Index           =   6
      Left            =   60
      TabIndex        =   68
      ToolTipText     =   "Open this path in Explorer"
      Top             =   3180
      Width           =   225
   End
   Begin VB.CommandButton OpenExplorer 
      Caption         =   "E"
      Height          =   255
      Index           =   5
      Left            =   60
      TabIndex        =   67
      ToolTipText     =   "Open this path in Explorer"
      Top             =   2790
      Width           =   225
   End
   Begin VB.CommandButton OpenExplorer 
      Caption         =   "E"
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   66
      ToolTipText     =   "Open this path in Explorer"
      Top             =   2400
      Width           =   225
   End
   Begin VB.CommandButton OpenExplorer 
      Caption         =   "E"
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   65
      ToolTipText     =   "Open this path in Explorer"
      Top             =   2010
      Width           =   225
   End
   Begin VB.CommandButton OpenExplorer 
      Caption         =   "E"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   64
      ToolTipText     =   "Open this path in Explorer"
      Top             =   1620
      Width           =   225
   End
   Begin VB.CommandButton OpenExplorer 
      Caption         =   "E"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   63
      ToolTipText     =   "Open this path in Explorer"
      Top             =   1230
      Width           =   225
   End
   Begin VB.CommandButton OpenExplorer 
      Caption         =   "E"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   62
      ToolTipText     =   "Open this path in Explorer"
      Top             =   840
      Width           =   225
   End
   Begin VB.CommandButton DeleteLine 
      Caption         =   "X"
      Height          =   255
      Index           =   9
      Left            =   7950
      TabIndex        =   61
      ToolTipText     =   "Remove this line"
      Top             =   4350
      Width           =   225
   End
   Begin VB.CommandButton DeleteLine 
      Caption         =   "X"
      Height          =   255
      Index           =   8
      Left            =   7950
      TabIndex        =   60
      ToolTipText     =   "Remove this line"
      Top             =   3960
      Width           =   225
   End
   Begin VB.CommandButton DeleteLine 
      Caption         =   "X"
      Height          =   255
      Index           =   7
      Left            =   7950
      TabIndex        =   59
      ToolTipText     =   "Remove this line"
      Top             =   3570
      Width           =   225
   End
   Begin VB.CommandButton DeleteLine 
      Caption         =   "X"
      Height          =   255
      Index           =   6
      Left            =   7950
      TabIndex        =   58
      ToolTipText     =   "Remove this line"
      Top             =   3180
      Width           =   225
   End
   Begin VB.CommandButton DeleteLine 
      Caption         =   "X"
      Height          =   255
      Index           =   5
      Left            =   7950
      TabIndex        =   57
      ToolTipText     =   "Remove this line"
      Top             =   2790
      Width           =   225
   End
   Begin VB.CommandButton DeleteLine 
      Caption         =   "X"
      Height          =   255
      Index           =   4
      Left            =   7950
      TabIndex        =   56
      ToolTipText     =   "Remove this line"
      Top             =   2400
      Width           =   225
   End
   Begin VB.CommandButton DeleteLine 
      Caption         =   "X"
      Height          =   255
      Index           =   3
      Left            =   7950
      TabIndex        =   55
      ToolTipText     =   "Remove this line"
      Top             =   2010
      Width           =   225
   End
   Begin VB.CommandButton DeleteLine 
      Caption         =   "X"
      Height          =   255
      Index           =   2
      Left            =   7950
      TabIndex        =   54
      ToolTipText     =   "Remove this line"
      Top             =   1620
      Width           =   225
   End
   Begin VB.CommandButton DeleteLine 
      Caption         =   "X"
      Height          =   255
      Index           =   1
      Left            =   7950
      TabIndex        =   53
      ToolTipText     =   "Remove this line"
      Top             =   1230
      Width           =   225
   End
   Begin VB.CommandButton DeleteLine 
      Caption         =   "X"
      Height          =   255
      Index           =   0
      Left            =   7950
      TabIndex        =   52
      ToolTipText     =   "Remove this line"
      Top             =   840
      Width           =   225
   End
   Begin VB.TextBox PlaylistType 
      Height          =   255
      Left            =   4500
      TabIndex        =   51
      Text            =   "typedir"
      Top             =   60
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox Path 
      Height          =   315
      Index           =   9
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "PATH"
      Top             =   4320
      Width           =   2775
   End
   Begin VB.ComboBox CriteriaType 
      Height          =   315
      Index           =   9
      Left            =   4110
      TabIndex        =   49
      Text            =   "CRIT TYPE"
      Top             =   4320
      Width           =   1005
   End
   Begin VB.ComboBox Context 
      Height          =   315
      Index           =   9
      Left            =   3120
      TabIndex        =   48
      Text            =   "CONTEXT"
      Top             =   4320
      Width           =   945
   End
   Begin VB.TextBox Criteria 
      Height          =   315
      Index           =   9
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "CRITERIA"
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox Path 
      Height          =   315
      Index           =   8
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "PATH"
      Top             =   3930
      Width           =   2775
   End
   Begin VB.ComboBox CriteriaType 
      Height          =   315
      Index           =   8
      Left            =   4110
      TabIndex        =   45
      Text            =   "CRIT TYPE"
      Top             =   3930
      Width           =   1005
   End
   Begin VB.ComboBox Context 
      Height          =   315
      Index           =   8
      Left            =   3120
      TabIndex        =   44
      Text            =   "CONTEXT"
      Top             =   3930
      Width           =   945
   End
   Begin VB.TextBox Criteria 
      Height          =   315
      Index           =   8
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   43
      Text            =   "CRITERIA"
      Top             =   3930
      Width           =   2775
   End
   Begin VB.TextBox Path 
      Height          =   315
      Index           =   7
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "PATH"
      Top             =   3540
      Width           =   2775
   End
   Begin VB.ComboBox CriteriaType 
      Height          =   315
      Index           =   7
      Left            =   4110
      TabIndex        =   41
      Text            =   "CRIT TYPE"
      Top             =   3540
      Width           =   1005
   End
   Begin VB.ComboBox Context 
      Height          =   315
      Index           =   7
      Left            =   3120
      TabIndex        =   40
      Text            =   "CONTEXT"
      Top             =   3540
      Width           =   945
   End
   Begin VB.TextBox Criteria 
      Height          =   315
      Index           =   7
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "CRITERIA"
      Top             =   3540
      Width           =   2775
   End
   Begin VB.CommandButton AddNewPath 
      Caption         =   "Add New"
      Height          =   225
      Left            =   870
      TabIndex        =   38
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Preview 
      Caption         =   "Preview"
      Height          =   435
      Left            =   1530
      TabIndex        =   37
      Top             =   4740
      Width           =   1215
   End
   Begin VB.TextBox Criteria 
      Height          =   315
      Index           =   6
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "CRITERIA"
      Top             =   3150
      Width           =   2775
   End
   Begin VB.ComboBox Context 
      Height          =   315
      Index           =   6
      Left            =   3120
      TabIndex        =   35
      Text            =   "CONTEXT"
      Top             =   3150
      Width           =   945
   End
   Begin VB.ComboBox CriteriaType 
      Height          =   315
      Index           =   6
      Left            =   4110
      TabIndex        =   34
      Text            =   "CRIT TYPE"
      Top             =   3150
      Width           =   1005
   End
   Begin VB.TextBox Path 
      Height          =   315
      Index           =   6
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "PATH"
      Top             =   3150
      Width           =   2775
   End
   Begin VB.TextBox Criteria 
      Height          =   315
      Index           =   5
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "CRITERIA"
      Top             =   2760
      Width           =   2775
   End
   Begin VB.ComboBox Context 
      Height          =   315
      Index           =   5
      Left            =   3120
      TabIndex        =   31
      Text            =   "CONTEXT"
      Top             =   2760
      Width           =   945
   End
   Begin VB.ComboBox CriteriaType 
      Height          =   315
      Index           =   5
      Left            =   4110
      TabIndex        =   30
      Text            =   "CRIT TYPE"
      Top             =   2760
      Width           =   1005
   End
   Begin VB.TextBox Path 
      Height          =   315
      Index           =   5
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "PATH"
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox Criteria 
      Height          =   315
      Index           =   4
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "CRITERIA"
      Top             =   2370
      Width           =   2775
   End
   Begin VB.ComboBox Context 
      Height          =   315
      Index           =   4
      Left            =   3120
      TabIndex        =   27
      Text            =   "CONTEXT"
      Top             =   2370
      Width           =   945
   End
   Begin VB.ComboBox CriteriaType 
      Height          =   315
      Index           =   4
      Left            =   4110
      TabIndex        =   26
      Text            =   "CRIT TYPE"
      Top             =   2370
      Width           =   1005
   End
   Begin VB.TextBox Path 
      Height          =   315
      Index           =   4
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "PATH"
      Top             =   2370
      Width           =   2775
   End
   Begin VB.TextBox Criteria 
      Height          =   315
      Index           =   3
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "CRITERIA"
      Top             =   1980
      Width           =   2775
   End
   Begin VB.ComboBox Context 
      Height          =   315
      Index           =   3
      Left            =   3120
      TabIndex        =   23
      Text            =   "CONTEXT"
      Top             =   1980
      Width           =   945
   End
   Begin VB.ComboBox CriteriaType 
      Height          =   315
      Index           =   3
      Left            =   4110
      TabIndex        =   22
      Text            =   "CRIT TYPE"
      Top             =   1980
      Width           =   1005
   End
   Begin VB.TextBox Path 
      Height          =   315
      Index           =   3
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "PATH"
      Top             =   1980
      Width           =   2775
   End
   Begin VB.TextBox Criteria 
      Height          =   315
      Index           =   2
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "CRITERIA"
      Top             =   1590
      Width           =   2775
   End
   Begin VB.ComboBox Context 
      Height          =   315
      Index           =   2
      Left            =   3120
      TabIndex        =   19
      Text            =   "CONTEXT"
      Top             =   1590
      Width           =   945
   End
   Begin VB.ComboBox CriteriaType 
      Height          =   315
      Index           =   2
      Left            =   4110
      TabIndex        =   18
      Text            =   "CRIT TYPE"
      Top             =   1590
      Width           =   1005
   End
   Begin VB.TextBox Path 
      Height          =   315
      Index           =   2
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "PATH"
      Top             =   1590
      Width           =   2775
   End
   Begin VB.TextBox Criteria 
      Height          =   315
      Index           =   1
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "CRITERIA"
      Top             =   1200
      Width           =   2775
   End
   Begin VB.ComboBox Context 
      Height          =   315
      Index           =   1
      Left            =   3120
      TabIndex        =   15
      Text            =   "CONTEXT"
      Top             =   1200
      Width           =   945
   End
   Begin VB.ComboBox CriteriaType 
      Height          =   315
      Index           =   1
      Left            =   4110
      TabIndex        =   14
      Text            =   "CRIT TYPE"
      Top             =   1200
      Width           =   1005
   End
   Begin VB.TextBox Path 
      Height          =   315
      Index           =   1
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "PATH"
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox Criteria 
      Height          =   315
      Index           =   0
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "CRITERIA"
      Top             =   810
      Width           =   2775
   End
   Begin VB.TextBox Path 
      Height          =   315
      Index           =   0
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "PATH"
      Top             =   810
      Width           =   2775
   End
   Begin VB.ComboBox CriteriaType 
      Height          =   315
      Index           =   0
      Left            =   4110
      TabIndex        =   7
      Text            =   "CRIT TYPE"
      Top             =   810
      Width           =   1005
   End
   Begin VB.ComboBox Context 
      Height          =   315
      Index           =   0
      Left            =   3120
      TabIndex        =   6
      Text            =   "CONTEXT"
      Top             =   810
      Width           =   945
   End
   Begin VB.VScrollBar ScrollBar 
      Height          =   3915
      LargeChange     =   8
      Left            =   8400
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   750
      Width           =   285
   End
   Begin VB.CommandButton Save 
      Caption         =   "Save"
      Height          =   435
      Left            =   3180
      TabIndex        =   4
      Top             =   4740
      Width           =   1215
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   435
      Left            =   4860
      TabIndex        =   3
      Top             =   4740
      Width           =   1215
   End
   Begin VB.TextBox fname 
      Height          =   255
      Left            =   3930
      TabIndex        =   2
      Text            =   "fname"
      Top             =   60
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox PlaylistName 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "PLAYLIST NAME"
      Top             =   60
      Width           =   2505
   End
   Begin VB.Label Label4 
      Caption         =   "Criteria"
      Height          =   195
      Left            =   3960
      TabIndex        =   11
      Top             =   510
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "Play at:"
      Height          =   195
      Left            =   3030
      TabIndex        =   10
      Top             =   510
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Path:"
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   510
      Width           =   1245
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   8700
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8700
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Currently editing:"
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   1245
   End
End
Attribute VB_Name = "ARPlaylistEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NUM_BOXES = 10

Private currentpos As Long
Private lineList() As Dictionary
Private changed As Boolean

Private Sub AddNewPath_Click()
    Dim newpath As String
    Dim fp As FolderPicker
    Set fp = New FolderPicker
    fp.SetParent Me.hWnd
    fp.ShowFilesToo True
    newpath = fp.GetFolder
    Set fp = Nothing
    newpath = LCase(newpath)
    
    If newpath = "" Then
        Exit Sub
    End If
    
    On Error Resume Next
    ReDim Preserve lineList(UBound(lineList) + 1) As Dictionary
    If Err.Number = 9 Then
        Err.Clear
        ReDim Preserve lineList(0) As Dictionary
    End If
    On Error GoTo 0
    Set lineList(UBound(lineList)) = New Dictionary
    lineList(UBound(lineList)).Add "path", newpath
    lineList(UBound(lineList)).Add "context", "both"
    lineList(UBound(lineList)).Add "criteria_type", "none"
    lineList(UBound(lineList)).Add "criteria", "none"
    
    If UBound(lineList) < NUM_BOXES Then
        fillForm
    Else
        Me.ScrollBar.Max = Me.ScrollBar.Max + 1
        Me.ScrollBar.Value = UBound(lineList) - NUM_BOXES + 1
    End If
End Sub

Private Sub Context_Click(index As Integer)
    changed = True
    Context_LostFocus index
End Sub

Private Sub Context_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.Context(index), KeyCode, Shift
    changed = True
End Sub

Private Sub Context_KeyPress(index As Integer, KeyAscii As Integer)
    AutoCompleteKeyPress Me.Context(index), KeyAscii
    changed = True
End Sub

Private Sub Context_LostFocus(index As Integer)
    AutoCompleteLostFocus Me.Context(index)
    If changed Then
        changed = False
        If Me.Context(index) = "" Then
            Me.Context(index) = "none"
        End If
        lineList(currentpos + index).item("context") = Me.Context(index)
    End If
End Sub

Private Sub CriteriaType_Click(index As Integer)
    changed = True
    CriteriaType_LostFocus index
End Sub

Private Sub CriteriaType_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.CriteriaType(index), KeyCode, Shift
    changed = True
End Sub

Private Sub CriteriaType_KeyPress(index As Integer, KeyAscii As Integer)
    AutoCompleteKeyPress Me.CriteriaType(index), KeyAscii
    changed = True
End Sub

Private Sub CriteriaType_LostFocus(index As Integer)
    AutoCompleteLostFocus Me.CriteriaType(index)
    If changed = False Then
        Exit Sub
    End If
    changed = False
    If Me.CriteriaType(index) = "" Then
        Me.CriteriaType(index) = "none"
    End If
    lineList(currentpos + index).item("criteria_type") = Me.CriteriaType(index)
    Select Case Me.CriteriaType(index)
        Case Is = "none"
            Me.Criteria(index) = ""
            lineList(currentpos + index).item("criteria") = Me.Criteria(index)
        Case Is = "date"
            Dim startdate As String, enddate As String, annual As Boolean
            annual = CBool(MsgBox("YES = include year in criteria, so it'll only play for this year" & vbCrLf & "NO = don't include year, will be annual", vbYesNo) = vbNo)
            Do
                startdate = InputBox("Enter starting date:", , Format(Date, IIf(annual, "m/d", "m/d/yyyy")))
            Loop While Not IsDate(startdate) And startdate <> ""
            If startdate = "" Then
                Me.Criteria(index) = ""
                Me.CriteriaType(index) = "none"
                GoTo theend
            End If
            Do
                enddate = InputBox("Enter ending date:", , Format(Date, IIf(annual, "m/d", "m/d/yyyy")))
            Loop While Not IsDate(enddate) And enddate <> ""
            If enddate = "" Then
                Me.Criteria(index) = ""
                Me.CriteriaType(index) = "none"
                GoTo theend
            End If
            If annual Then
                Me.Criteria(index) = Format(startdate, "m/d") & " to " & Format(enddate, "m/d") & ", every year"
                lineList(currentpos + index).item("criteria") = "XXXX-" & Format(startdate, "mm-dd") & " " & "XXXX-" & Format(enddate, "mm-dd")
            Else
                Me.Criteria(index) = Format(startdate, "m/d/yyyy") & " to " & Format(enddate, "m/d/yyyy")
                lineList(currentpos + index).item("criteria") = Format(startdate, "yyyy-mm-dd") & " " & Format(enddate, "yyyy-mm-dd")
            End If
        Case Is = "quantity"
            Dim item As String, qty As String
            item = InputBox("Enter the itemnumber:" & vbCrLf & vbCrLf & "(sorry this isn't a selectbox)")
            If item = "" Then
                Me.Criteria(index) = ""
                Me.CriteriaType(index) = "none"
                GoTo theend
            End If
            qty = InputBox("Enter the quantity threshold (will play as long as inventory is above this #):")
            If qty = "" Or Not IsNumeric(qty) Then
                Me.Criteria(index) = ""
                Me.CriteriaType(index) = "none"
                GoTo theend
            End If
            Me.Criteria(index) = UCase(item) & " " & ">" & " " & qty
            lineList(currentpos + index).item("criteria") = Me.Criteria(index)
    End Select
    Exit Sub
theend:
    lineList(currentpos + index).item("criteria_type") = Me.CriteriaType(index)
End Sub

Private Sub fixPath(index As Long)
    If LCase(Left(Me.Path(index), 15)) = "h:\audio runner" Then
        Me.Path(index) = Replace(Me.Path(index), "h:\audio runner", "", , , vbTextCompare)
        If Left(Me.Path(index), 1) <> "\" Then
            Me.Path(index) = "\" & Me.Path(index)
        End If
    End If
End Sub

Private Sub fixCriteria(index As Long)
    Select Case Me.CriteriaType(index)
        Case Is = "none"
            'do nothing
        Case Is = "date"
            Dim startdate As String, enddate As String
            startdate = Left(Me.Criteria(index), InStr(Me.Criteria(index), " ") - 1)
            enddate = Mid(Me.Criteria(index), InStr(Me.Criteria(index), " ") + 1)
            Dim annual As Boolean
            annual = CBool(InStr(startdate, "XXXX-"))
            If annual Then
                startdate = Replace(startdate, "XXXX-", "")
                enddate = Replace(enddate, "XXXX-", "")
                Me.Criteria(index) = Format(startdate, "m/d") & " to " & Format(enddate, "m/d") & ", every year"
            Else
                Me.Criteria(index) = Format(startdate, "m/d/yyyy") & " to " & Format(enddate, "m/d/yyyy")
            End If
        Case Is = "quantity"
            'do nothing
    End Select
End Sub

Private Sub DeleteLine_Click(index As Integer)
    If vbYes = MsgBox("Are you sure you want to delete this line?", vbYesNo) Then
        Dim i As Long
        For i = currentpos + index + 1 To UBound(lineList)
            Set lineList(i - 1) = lineList(i)
        Next i
        If UBound(lineList) = 0 Then
            'ReDim lineList(-1) As Dictionary
            Erase lineList
        Else
            ReDim Preserve lineList(UBound(lineList) - 1) As Dictionary
        End If
        fillForm currentpos
    End If
End Sub

Private Sub Exit_Click()
    Load ARNewPlaylist
    ARNewPlaylist.Show
    Unload Me
End Sub

Private Sub fname_Change()
    loadFile Me.fname
    fillForm
End Sub

Private Sub Form_Load()
    requeryContext
    requeryCriteria
End Sub

Private Sub OpenExplorer_Click(index As Integer)
    Shell "explorer.exe " & IIf(Left(Me.Path(index), 1) = "\", "h:\audio runner", "") & Me.Path(index), vbNormalFocus
End Sub

Private Sub Preview_Click()
'TODO
    MsgBox "TO DO!"
End Sub

Private Sub Save_Click()
    If MsgBox("Save this playlist?", vbYesNo) = vbYes Then
        Open Me.fname For Output As #1
        Print #1, "path" & vbTab & "context" & vbTab & "criteria_type" & vbTab & "criteria"
        Dim i As Long
        For i = 0 To UBound(lineList)
            Print #1, lineList(i).item("path") & vbTab & lineList(i).item("context") & vbTab & lineList(i).item("criteria_type") & vbTab & lineList(i).item("criteria")
        Next i
        Close #1
    End If
End Sub

Private Sub requeryContext()
    Dim i As Integer
    For i = Me.CriteriaType.lBound To Me.CriteriaType.UBound
        Me.Context(i).Clear
        Me.Context(i).AddItem "both"
        Me.Context(i).AddItem "store"
        Me.Context(i).AddItem "phone"
        Me.Context(i).AddItem "none"
    Next i
End Sub

Private Sub requeryCriteria()
    Dim i As Integer
    For i = Me.CriteriaType.lBound To Me.CriteriaType.UBound
        Me.CriteriaType(i).Clear
        Me.CriteriaType(i).AddItem "none"
        Me.CriteriaType(i).AddItem "date"
        Me.CriteriaType(i).AddItem "quantity"
    Next i
End Sub

Private Sub loadFile(Filename As String)
    Dim lines As Variant
    lines = ReadXSV(Filename, vbTab)
    On Error Resume Next
    ReDim lineList(UBound(lines)) As Dictionary
    If Err.Number = 9 Then
        Err.Clear
        ReDim lineList(-1) As Dictionary
    End If
    On Error GoTo 0
    Dim i As Integer
    For i = 0 To UBound(lines)
        Set lineList(i) = lines(i)
    Next i
End Sub

Private Sub fillForm(Optional offset As Long = 0)
On Error GoTo errh
    Dim norecs As Boolean
    Dim i As Long
    For i = 0 To NUM_BOXES - 1
        If UBound(lineList) >= i Then
            setVisible i, True
            Me.Path(i) = lineList(offset + i).item("path")
            fixPath i
            Me.Context(i) = lineList(offset + i).item("context")
            Me.CriteriaType(i) = lineList(offset + i).item("criteria_type")
            Me.Criteria(i) = lineList(offset + i).item("criteria")
            fixCriteria i
        Else
clearform:
            While i < NUM_BOXES
                setVisible i, False
                i = i + 1
            Wend
        End If
    Next i
    Me.ScrollBar.Min = 0
    If norecs Then
        Me.ScrollBar.Max = 0
    Else
        Me.ScrollBar.Max = IIf(UBound(lineList) - NUM_BOXES <= 0, 0, UBound(lineList) - NUM_BOXES + 1)
    End If
    Exit Sub
errh:
    If Err.Number = 9 Then
        norecs = True
        Err.Clear
        GoTo clearform
    End If
End Sub

Private Sub setVisible(i As Long, onoff As Boolean)
    Me.Path(i).Visible = onoff
    Me.Context(i).Visible = onoff
    Me.CriteriaType(i).Visible = onoff
    Me.Criteria(i).Visible = onoff
    Me.DeleteLine(i).Visible = onoff
    Me.OpenExplorer(i).Visible = onoff
End Sub

Private Sub ScrollBar_Change()
    If Abs(Me.ScrollBar.Value - currentpos) > 1 Then
        'refill
        fillForm Me.ScrollBar.Value
    Else
        'slide
        If Me.ScrollBar.Value > currentpos Then 'down arrow
            shuffleUp
        Else
            shuffleDown
        End If
    End If
    currentpos = Me.ScrollBar.Value
End Sub

Private Sub shuffleUp()
    Dim i As Long
    For i = 0 To NUM_BOXES - 2
        shuffle i + 1, i
    Next i
    Me.Path(i) = lineList(Me.ScrollBar.Value + i).item("path")
    Me.Context(i) = lineList(Me.ScrollBar.Value + i).item("context")
    Me.CriteriaType(i) = lineList(Me.ScrollBar.Value + i).item("criteria_type")
    Me.Criteria(i) = lineList(Me.ScrollBar.Value + i).item("criteria")
    fixCriteria i
End Sub

Private Sub shuffleDown()
    Dim i As Long
    For i = NUM_BOXES - 1 To 1 Step -1
        shuffle i - 1, i
    Next i
    Me.Path(i) = lineList(Me.ScrollBar.Value + i).item("path")
    Me.Context(i) = lineList(Me.ScrollBar.Value + i).item("context")
    Me.CriteriaType(i) = lineList(Me.ScrollBar.Value + i).item("criteria_type")
    Me.Criteria(i) = lineList(Me.ScrollBar.Value + i).item("criteria")
    fixCriteria i
End Sub

Private Sub shuffle(fromIndex As Long, toIndex As Long)
    Me.Path(toIndex) = Me.Path(fromIndex)
    Me.Context(toIndex) = Me.Context(fromIndex)
    Me.CriteriaType(toIndex) = Me.CriteriaType(fromIndex)
    Me.Criteria(toIndex) = Me.Criteria(fromIndex)
End Sub
