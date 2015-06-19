VERSION 5.00
Begin VB.Form CalculateSalesRanking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate Sales Ranking"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4155
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnExit 
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   2190
      TabIndex        =   17
      Top             =   3450
      Width           =   1305
   End
   Begin VB.CommandButton btnGo 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   3450
      Width           =   1305
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Select % Thresholds:"
      Height          =   1125
      Index           =   2
      Left            =   180
      TabIndex        =   7
      Top             =   2160
      Width           =   3675
      Begin VB.TextBox RankDThreshold 
         Height          =   285
         Left            =   2160
         TabIndex        =   15
         Top             =   630
         Width           =   705
      End
      Begin VB.TextBox RankBThreshold 
         Height          =   285
         Left            =   750
         TabIndex        =   13
         Top             =   630
         Width           =   705
      End
      Begin VB.TextBox RankCThreshold 
         Height          =   285
         Left            =   2160
         TabIndex        =   11
         Top             =   270
         Width           =   705
      End
      Begin VB.TextBox RankAThreshold 
         Height          =   285
         Left            =   750
         TabIndex        =   9
         Top             =   270
         Width           =   705
      End
      Begin VB.Label generalLabel 
         Caption         =   "D:"
         Height          =   255
         Index           =   5
         Left            =   1860
         TabIndex        =   14
         Top             =   660
         Width           =   345
      End
      Begin VB.Label generalLabel 
         Caption         =   "B:"
         Height          =   255
         Index           =   4
         Left            =   450
         TabIndex        =   12
         Top             =   660
         Width           =   345
      End
      Begin VB.Label generalLabel 
         Caption         =   "C:"
         Height          =   255
         Index           =   3
         Left            =   1860
         TabIndex        =   10
         Top             =   300
         Width           =   345
      End
      Begin VB.Label generalLabel 
         Caption         =   "A:"
         Height          =   255
         Index           =   2
         Left            =   450
         TabIndex        =   8
         Top             =   300
         Width           =   345
      End
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Select Line Codes:"
      Height          =   1185
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   780
      Width           =   3675
      Begin VB.CheckBox LinesAll 
         Caption         =   "All Lines"
         Height          =   285
         Left            =   2220
         TabIndex        =   4
         Top             =   510
         Width           =   1005
      End
      Begin VB.ComboBox LinesEnd 
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Top             =   690
         Width           =   1155
      End
      Begin VB.ComboBox LinesStart 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label generalLabel 
         Caption         =   "To:"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Top             =   690
         Width           =   705
      End
      Begin VB.Label generalLabel 
         Caption         =   "From:"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   330
         Width           =   645
      End
   End
   Begin VB.Label lblStatusBar 
      Height          =   225
      Left            =   30
      TabIndex        =   18
      Top             =   3900
      Width           =   3765
   End
   Begin VB.Label Label2 
      Caption         =   "Calculate Sales Ranking"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3645
   End
End
Attribute VB_Name = "CalculateSalesRanking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload CalculateSalesRanking
End Sub

Private Sub btnGo_Click()
    If Me.LinesStart = "" Or Me.LinesEnd = "" Then
        Exit Sub
    End If
    
    Mouse.Hourglass True
    
    Dim conn As ADODB.Connection
    Set conn = openMSSQLConn
    
    Dim rstLines As ADODB.Recordset
    If Me.LinesAll Then
        Set rstLines = retrieve("SELECT * FROM vRequeryLineCodes ORDER BY ProductLine", conn)
    Else
        Set rstLines = retrieve("exec spLineCodesRange '" & Me.LinesStart & "', '" & Me.LinesEnd & "'", conn)
    End If

    Dim rstBaseline As ADODB.Recordset
    Dim rst As ADODB.Recordset
    Dim itemArrays As Dictionary
    
    Dim rankTypes() As Variant
    ' fields MUST be named similarly, assumes there's always going to be at least a *Percent and *SalesRank
    rankTypes = Array("ThisPeriodUnits", "ThisPeriodDollars", "LastPeriodUnits", "LastPeriodDollars", "YTDUnits", "YTDDollars", "LYRUnits", "LYRDollars")
    ReDim current(UBound(rankTypes)) As String
    
    Dim total As Double
    Dim strsql As String, i As Integer, j As Integer
    
    While Not rstLines.EOF
        If rstLines("ProductLine") <> "XXX" And rstLines("ProductLine") <> "ZZZ" Then
            Me.lblStatusBar.caption = "Processing " & rstLines("ProductLine")
            DoEvents
            ' crunch the numbers
            '    rstBaseline contains the total sales in units and dollars for each period we're interested in for entire line code
            Set rstBaseline = retrieve("exec spSalesRankBaseline '" & rstLines("ProductLine") & "'", conn)
    
            Set rst = retrieve("exec spSalesRankItemsByPL '" & rstLines("ProductLine") & "'", conn)
            
            dbExecute "DELETE FROM SalesRankings WHERE ItemNumber IN ( SELECT ItemNumber FROM InventoryMaster WHERE ProductLine='" & rstLines("ProductLine") & "')", conn
            dbExecute "INSERT INTO SalesRankings ( ItemNumber ) SELECT ItemNumber FROM InventoryMaster WHERE ProductLine='" & rstLines("ProductLine") & "'", conn
    
            While Not rst.EOF
                dbExecute "exec spSalesRankSetItemNumbers '" & rst("ItemNumber") & "', " & _
                                                               rst("SalesThisPeriod") & ", " & rst("SalesLastPeriod") & ", " & rst("SalesYTD") & ", " & rst("SalesLYR") & ", " & _
                                                               rstBaseline("SalesThisPeriod") & ", " & rstBaseline("SalesLastPeriod") & ", " & rstBaseline("SalesYTD") & ", " & rstBaseline("SalesLYR") & ", " & _
                                                               rst("DollarsThisPeriod") & ", " & rst("DollarsLastPeriod") & ", " & rst("DollarsYTD") & ", " & rst("DollarsLYR") & ", " & _
                                                               rstBaseline("DollarsThisPeriod") & ", " & rstBaseline("DollarsLastPeriod") & ", " & rstBaseline("DollarsYTD") & ", " & rstBaseline("DollarsLYR"), conn
                rst.MoveNext
            Wend
            rst.Close
            rstBaseline.Close
            
            ' calculate sales ranking letter
            '    anything in the top 70th (by default) percentile goes into A, 70-80 B, etc
            '    save everything in a hash of arrays, itemnumber -> array of letters for each metric
            Set itemArrays = New Dictionary
    
            For i = 0 To UBound(rankTypes)
                Set rst = retrieve("SELECT ItemNumber, " & rankTypes(i) & "Percent FROM SalesRankings WHERE LEFT(ItemNumber,3)='" & rstLines("ProductLine") & "' ORDER BY " & rankTypes(i) & "Percent DESC", conn)
                total = 0
                While Not rst.EOF
                    If itemArrays.exists(rst("ItemNumber").value) Then
                        current = itemArrays.item(rst("ItemNumber").value)
                    End If
                    If total < Me.RankAThreshold Then
                        current(i) = "A"
                    ElseIf total < Me.RankBThreshold Then
                        current(i) = "B"
                    ElseIf total < Me.RankCThreshold Then
                        current(i) = "C"
                    Else
                        current(i) = "D"
                    End If
                    total = total + rst(rankTypes(i) & "Percent")
                    itemArrays.item(rst("ItemNumber").value) = current
                    rst.MoveNext
                Wend
                rst.Close
            Next i
            
            ' update the table with the info from the HoA
            For i = 0 To UBound(itemArrays.Keys)
                strsql = "UPDATE SalesRankings SET "
                For j = 0 To UBound(rankTypes)
                    strsql = strsql & rankTypes(j) & "SalesRank='" & itemArrays.item(itemArrays.Keys(i))(j) & "'"
                    If j <> UBound(rankTypes) Then
                        strsql = strsql & ", "
                    End If
                Next j
                strsql = strsql & " WHERE ItemNumber='" & itemArrays.Keys(i) & "'"
                dbExecute strsql, conn
            Next i
            
            Set itemArrays = Nothing
        End If
        rstLines.MoveNext
    Wend
    rstLines.Close
    
    closeMSSQLConn conn
    
    Set conn = Nothing
    Set rstLines = Nothing
    Set rstBaseline = Nothing
    Set rst = Nothing
    
    Me.lblStatusBar.caption = ""
    
    Mouse.Hourglass False

End Sub

Private Sub Form_Load()
    requeryLines
    Me.LinesAll = 1
    Me.LinesStart.Enabled = False
    Me.LinesEnd.Enabled = False
    Me.LinesStart = "AAA"
    Me.LinesEnd = "ZZZ"
    Me.RankAThreshold = "70"
    Me.RankBThreshold = "80"
    Me.RankCThreshold = "95"
    Me.RankDThreshold = "100"
End Sub

Private Sub LinesAll_Click()
    If Me.LinesAll Then
        Disable Me.LinesStart
        Disable Me.LinesEnd
        Me.LinesStart = "AAA"
        Me.LinesEnd = "ZZZ"
    Else
        Enable Me.LinesStart
        Enable Me.LinesEnd
        Me.LinesStart = ""
        Me.LinesEnd = ""
    End If
End Sub

Private Sub LinesStart_Click()
    LinesStart_LostFocus
End Sub

Private Sub LinesStart_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.LinesStart, KeyCode, Shift
    If KeyCode = vbKeyReturn Then
        LinesStart_LostFocus
    End If
End Sub

Private Sub LinesStart_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.LinesStart, KeyAscii
End Sub

Private Sub LinesStart_LostFocus()
    AutoCompleteLostFocus Me.LinesStart
    Me.LinesEnd = Me.LinesStart
End Sub

Private Sub LinesEnd_Click()
    LinesEnd_LostFocus
End Sub

Private Sub LinesEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.LinesEnd, KeyCode, Shift
    If KeyCode = vbKeyReturn Then
        LinesEnd_LostFocus
    End If
End Sub

Private Sub LinesEnd_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.LinesEnd, KeyAscii
End Sub

Private Sub LinesEnd_LostFocus()
    AutoCompleteLostFocus Me.LinesEnd
End Sub

Private Sub requeryLines()
    Dim rst As ADODB.Recordset
    Set rst = retrieve("SELECT * FROM vRequeryLineCodes ORDER BY ProductLine")
    Me.LinesStart.Clear
    Me.LinesEnd.Clear
    While Not rst.EOF
        Me.LinesStart.AddItem rst("ProductLine")
        Me.LinesEnd.AddItem rst("ProductLine")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub
