VERSION 5.00
Begin VB.Form InventoryCountMgmt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory Count Management"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   14715
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton ViewChangesSinceSnapshot 
      Caption         =   "View Changes Since Snapshot"
      Height          =   735
      Left            =   120
      TabIndex        =   54
      Top             =   9570
      Width           =   1455
   End
   Begin VB.CommandButton ReSnapshot 
      Caption         =   "Re-Snapshot"
      Height          =   315
      Left            =   4200
      TabIndex        =   53
      Top             =   2850
      Width           =   1275
   End
   Begin VB.CommandButton CommitBatch 
      Caption         =   "Send Batch to MAS200"
      Enabled         =   0   'False
      Height          =   765
      Left            =   6690
      TabIndex        =   52
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton AdjustEntry 
      Caption         =   "Adjust This Entry"
      Enabled         =   0   'False
      Height          =   315
      Left            =   11070
      TabIndex        =   51
      Top             =   9360
      Width           =   1665
   End
   Begin VB.CommandButton DeleteJob 
      Caption         =   "Delete Job"
      Height          =   585
      Left            =   6030
      TabIndex        =   50
      Top             =   2070
      Width           =   705
   End
   Begin VB.CommandButton CreateNewBatch 
      Caption         =   "New"
      Height          =   225
      Left            =   4860
      TabIndex        =   49
      Top             =   9360
      Width           =   615
   End
   Begin VB.ListBox BatchList 
      Height          =   840
      Left            =   3960
      TabIndex        =   46
      Top             =   9570
      Width           =   2355
   End
   Begin VB.CommandButton AddItemAdjustments 
      Caption         =   "Adjust Quantity in MAS200 For Item(s)"
      Enabled         =   0   'False
      Height          =   765
      Left            =   1920
      TabIndex        =   45
      Top             =   9570
      Width           =   1845
   End
   Begin VB.Frame Frame2 
      Caption         =   "Off By counts from:"
      Height          =   585
      Left            =   4470
      TabIndex        =   42
      Top             =   3450
      Width           =   2445
      Begin VB.OptionButton OffByCountFrom 
         Caption         =   "Current"
         Height          =   225
         Index           =   1
         Left            =   1350
         TabIndex        =   44
         Top             =   270
         Width           =   975
      End
      Begin VB.OptionButton OffByCountFrom 
         Caption         =   "Snapshot"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   43
         Top             =   270
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Close"
      Height          =   435
      Left            =   5190
      TabIndex        =   39
      Top             =   0
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Create New Job"
      Height          =   3225
      Left            =   7050
      TabIndex        =   25
      Top             =   30
      Width           =   7605
      Begin VB.ComboBox NewJobLocation 
         Height          =   315
         Left            =   2070
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1230
         Width           =   2265
      End
      Begin VB.CheckBox NewJobSpecificLocation 
         Caption         =   "Specific Location:"
         Height          =   225
         Left            =   480
         TabIndex        =   37
         Top             =   1290
         Width           =   1605
      End
      Begin VB.ComboBox NewJobWarehouse 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   870
         Width           =   3195
      End
      Begin VB.CommandButton NewJobCreate 
         Caption         =   "Create New Job"
         Height          =   255
         Left            =   1410
         TabIndex        =   34
         Top             =   2910
         Width           =   2235
      End
      Begin VB.TextBox NewJobNotes 
         Height          =   1035
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   1830
         Width           =   4545
      End
      Begin VB.TextBox NewJobName 
         Height          =   285
         Left            =   150
         TabIndex        =   30
         Top             =   510
         Width           =   4545
      End
      Begin VB.ListBox NewJobLines 
         Height          =   2790
         Left            =   4830
         MultiSelect     =   2  'Extended
         TabIndex        =   29
         Top             =   360
         Width           =   945
      End
      Begin VB.ListBox NewJobAssignment 
         Height          =   2790
         Left            =   5850
         MultiSelect     =   2  'Extended
         TabIndex        =   26
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label generalLabel 
         Caption         =   "Where:"
         Height          =   195
         Index           =   7
         Left            =   570
         TabIndex        =   35
         Top             =   930
         Width           =   615
      End
      Begin VB.Label generalLabel 
         Caption         =   "Directions:"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   33
         Top             =   1620
         Width           =   825
      End
      Begin VB.Label generalLabel 
         Caption         =   "Description:"
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   31
         Top             =   300
         Width           =   825
      End
      Begin VB.Label generalLabel 
         Caption         =   "Lines:"
         Height          =   195
         Index           =   4
         Left            =   4860
         TabIndex        =   28
         Top             =   150
         Width           =   615
      End
      Begin VB.Label generalLabel 
         Caption         =   "Assign To:"
         Height          =   195
         Index           =   3
         Left            =   5880
         TabIndex        =   27
         Top             =   150
         Width           =   825
      End
   End
   Begin VB.CommandButton MarkJobClosed 
      Caption         =   "Mark Job(s) Closed"
      Height          =   885
      Left            =   6030
      TabIndex        =   24
      Top             =   1050
      Width           =   705
   End
   Begin VB.ComboBox DrillDownSelect 
      Height          =   315
      ItemData        =   "InventoryCountMgmt.frx":0000
      Left            =   1200
      List            =   "InventoryCountMgmt.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3600
      Width           =   3195
   End
   Begin VB.CheckBox ShowCompletedJobs 
      Caption         =   "Show Completed Jobs"
      Height          =   495
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1590
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox ShowInProgressJobs 
      Caption         =   "Show In-Progress Jobs"
      Height          =   495
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.ListBox DrillDown 
      Height          =   4935
      Index           =   2
      Left            =   9210
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   4350
      Width           =   5475
   End
   Begin VB.ListBox DrillDown 
      Height          =   4935
      Index           =   1
      Left            =   5640
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   4350
      Width           =   3525
   End
   Begin VB.ListBox DrillDown 
      Height          =   4935
      Index           =   0
      Left            =   90
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   4350
      Width           =   5505
   End
   Begin VB.CheckBox ShowClosedJobs 
      Caption         =   "Show Closed Jobs"
      Height          =   495
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2100
      Width           =   1335
   End
   Begin VB.ListBox JobList 
      Height          =   2010
      Left            =   1620
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   780
      Width           =   4365
   End
   Begin VB.Label generalLabel 
      Caption         =   "# items"
      Height          =   195
      Index           =   9
      Left            =   5760
      TabIndex        =   48
      Top             =   9360
      Width           =   585
   End
   Begin VB.Label generalLabel 
      Caption         =   "Batch #"
      Height          =   195
      Index           =   8
      Left            =   4020
      TabIndex        =   47
      Top             =   9360
      Width           =   705
   End
   Begin VB.Label DDListHeader 
      Caption         =   "Adj."
      Height          =   195
      Index           =   5
      Left            =   4980
      TabIndex        =   41
      Top             =   4140
      Width           =   345
   End
   Begin VB.Label DDListHeader 
      Caption         =   "Current QOH"
      Height          =   375
      Index           =   3
      Left            =   3390
      TabIndex        =   40
      Top             =   3960
      Width           =   675
   End
   Begin VB.Label DDListHeader 
      Caption         =   "Item"
      Height          =   195
      Index           =   6
      Left            =   5730
      TabIndex        =   23
      Top             =   4140
      Width           =   1095
   End
   Begin VB.Label DDListHeader 
      Caption         =   "Time"
      Height          =   195
      Index           =   12
      Left            =   12870
      TabIndex        =   22
      Top             =   4140
      Width           =   615
   End
   Begin VB.Label DDListHeader 
      Caption         =   "Count"
      Height          =   195
      Index           =   11
      Left            =   12030
      TabIndex        =   21
      Top             =   4140
      Width           =   615
   End
   Begin VB.Label DDListHeader 
      Caption         =   "Person"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   10860
      TabIndex        =   20
      Top             =   4140
      Width           =   1065
   End
   Begin VB.Label DDListHeader 
      Caption         =   "Item"
      Height          =   195
      Index           =   9
      Left            =   9240
      TabIndex        =   19
      Top             =   4140
      Width           =   1095
   End
   Begin VB.Label DDListHeader 
      Caption         =   "Count"
      Height          =   195
      Index           =   8
      Left            =   8430
      TabIndex        =   18
      Top             =   4140
      Width           =   615
   End
   Begin VB.Label DDListHeader 
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   7290
      TabIndex        =   17
      Top             =   4140
      Width           =   1065
   End
   Begin VB.Label DDListHeader 
      Caption         =   "Off By"
      Height          =   195
      Index           =   4
      Left            =   4200
      TabIndex        =   16
      Top             =   4140
      Width           =   615
   End
   Begin VB.Label DDListHeader 
      Caption         =   "Snapshot QOH"
      Height          =   375
      Index           =   2
      Left            =   2550
      TabIndex        =   15
      Top             =   3960
      Width           =   675
   End
   Begin VB.Label DDListHeader 
      Caption         =   "Count"
      Height          =   195
      Index           =   1
      Left            =   1710
      TabIndex        =   14
      Top             =   4140
      Width           =   615
   End
   Begin VB.Label DDListHeader 
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   13
      Top             =   4140
      Width           =   1065
   End
   Begin VB.Label generalLabel 
      Caption         =   "Drill Down By:"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   3630
      Width           =   1065
   End
   Begin VB.Label jobListHeader 
      Caption         =   "Status"
      Height          =   195
      Index           =   1
      Left            =   5130
      TabIndex        =   10
      Top             =   570
      Width           =   615
   End
   Begin VB.Label jobListHeader 
      Caption         =   "Job Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1680
      TabIndex        =   9
      Top             =   570
      Width           =   1035
   End
   Begin VB.Label generalLabel 
      Caption         =   "Inventory Count Management"
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
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   30
      Width           =   4575
   End
   Begin VB.Label generalLabel 
      Caption         =   "Select job(s) to view:"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   810
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   14730
      Y1              =   3390
      Y2              =   3390
   End
End
Attribute VB_Name = "InventoryCountMgmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private jobListDescending As Boolean
Private ddListDescending(0 To 2) As Boolean

Private lastDDS As String

Private badWhseCombo As Boolean

Private stopevents As Boolean


Private Sub Form_Load()
    ReDim tabs(1) As Long
    tabs(0) = 154
    tabs(1) = tabs(0) + 300 'off screen
    SetListTabStops Me.JobList.hwnd, tabs
    
    ReDim tabs(5) As Long
    tabs(0) = 72
    tabs(1) = tabs(0) + 36
    tabs(2) = tabs(1) + 36
    tabs(3) = tabs(2) + 36
    tabs(4) = tabs(3) + 36
    tabs(5) = tabs(4) + 36
    SetListTabStops Me.DrillDown(0).hwnd, tabs
    
    ReDim tabs(1) As Long
    tabs(0) = 72
    tabs(1) = tabs(0) + 54
    SetListTabStops Me.DrillDown(1).hwnd, tabs
    SetListTabStops Me.DrillDown(2).hwnd, tabs
    
    ReDim tabs(0) As Long
    tabs(0) = 300 'off screen
    SetListTabStops Me.NewJobAssignment.hwnd, tabs
    tabs(0) = 90
    SetListTabStops Me.BatchList.hwnd, tabs
    
    requeryDrillDownSelect
    lastDDS = Me.DrillDownSelect
    requeryJobList
    requeryNewJobAssignment
    requeryNewJobLines
    requeryNewJobWarehouses
    requeryBatchList
    Disable Me.NewJobLocation
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub


'------------------
' joblist area
'------------------
Private Sub MarkJobClosed_Click()
    If Me.JobList.SelCount = 0 Then
        'do nothing
    ElseIf vbYes = MsgBox("Mark selected job(s) closed?", vbYesNo) Then
        DB.Execute "UPDATE HandheldCountJobs SET Complete=1 WHERE ID IN ( " & ListAsCommaSep(ListBoxAsArray(Me.JobList, True, 2)) & " )"
        requeryJobList
    End If
End Sub

Private Sub DeleteJob_Click()
    If Me.JobList.SelCount = 0 Then
        'do nothing
    ElseIf Me.JobList.SelCount >= 2 Then
        'not a technical limitation, just a precaution
        MsgBox "Can only delete 1 job at a time!"
    Else
        If vbYes = MsgBox("Are you sure you want to delete the selected job?", vbYesNo) Then
            Dim s As Variant
            s = Split(Me.JobList, vbTab)
            If DLookup("COUNT(*)", "HandheldCountJobDetailLines", "JobID=" & s(2)) <> 0 Then
                MsgBox "Can't delete a job with detail lines! Ask Brian if you really really want to."
            Else
                DB.Execute "DELETE FROM HandheldCountJobs WHERE ID=" & s(2)
                DB.Execute "DELETE FROM HandheldCountJobAssignments WHERE JobID=" & s(2)
                DB.Execute "DELETE FROM HandheldCountJobInventorySnapshots WHERE JobID=" & s(2)
            End If
        End If
    End If
End Sub

Private Sub ReSnapshot_Click()
    If Me.JobList.SelCount = 0 Then
        MsgBox "You haven't selected a job!"
    ElseIf Me.JobList.SelCount = 1 Then
        If vbYes = MsgBox("Re-snapshot this job", vbYesNo) Then
            Dim JobID As String
            JobID = Mid(Me.JobList, InStrRev(Me.JobList, vbTab) + 1)
            DB.Execute "DELETE FROM HandheldCountJobInventorySnapshots WHERE JobID=" & JobID
            createJobSnapshot JobID, Split(DLookup("LineItemFilter", "HandheldCountJobs", "ID=" & JobID), ",")
        End If
    Else
        MsgBox "Can only re-snapshot one job at a time!"
    End If
End Sub

Private Sub ShowCompletedJobs_Click()
    requeryJobList
End Sub

Private Sub ShowInProgressJobs_Click()
    requeryJobList
End Sub

Private Sub ShowClosedJobs_Click()
    requeryJobList
End Sub

Private Sub JobList_Click()
    If Me.JobList.SelCount > 0 Then
        badWhseCombo = DLookup("COUNT(DISTINCT MasWhseCode)", "HandheldCountJobs INNER JOIN Warehouses ON HandheldCountJobs.WarehouseID=Warehouses.ID", "HandheldCountJobs.ID IN ( " & ListAsCommaSep(ListBoxAsArray(Me.JobList, True, 2)) & " )") <> 1
        If badWhseCombo Then
            Me.AddItemAdjustments.Enabled = False
        End If
    Else
        badWhseCombo = True
    End If
    
    requeryDrillDown 0
    
    enableAdjustmentMaybe
End Sub

Private Sub jobListHeader_Click(Index As Integer)
    If Me.jobListHeader(Index).FontBold Then
        jobListDescending = Not jobListDescending
    Else
        Dim i As Long
        For i = Me.jobListHeader.LBound To Me.jobListHeader.UBound
            Me.jobListHeader(i).FontBold = False
        Next i
        Me.jobListHeader(Index).FontBold = True
        jobListDescending = False
    End If
    requeryJobList
End Sub

Private Sub requeryJobList()
    Dim whereclause As String
    If Me.ShowInProgressJobs Then
        whereclause = " WHERE Status='In Progress'"
    End If
    If Me.ShowCompletedJobs Then
        whereclause = IIf(whereclause = "", " WHERE", whereclause & " OR") & " Status='Complete'"
    End If
    If Me.ShowClosedJobs Then
        whereclause = IIf(whereclause = "", " WHERE", whereclause & " OR") & " Status='Closed'"
    End If
    whereclause = IIf(whereclause = "", " WHERE", whereclause & " OR") & " Status='ERROR'"
    
    Dim orderby As String
    Select Case True
        Case Is = Me.jobListHeader(0).FontBold
            orderby = " ORDER BY Name" & IIf(jobListDescending, " DESC", "")
        Case Is = Me.jobListHeader(1).FontBold
            orderby = " ORDER BY Status" & IIf(jobListDescending, " DESC", "")
    End Select
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, Name, Status FROM vCountMgmtJobList" & whereclause & orderby)
    Me.JobList.Clear
    Me.DrillDown(0).Clear
    Me.DrillDown(1).Clear
    Me.DrillDown(2).Clear
    Me.AddItemAdjustments.Enabled = False
    While Not rst.EOF
        Me.JobList.AddItem rst("Name") & vbTab & rst("Status") & vbTab & rst("ID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub


'----------------
' drilldowns
'----------------
Private Sub DrillDown_Click(Index As Integer)
    If Index <> Me.DrillDown.UBound Then
        requeryDrillDown Index + 1
    End If
    If badWhseCombo Then
        Me.AddItemAdjustments.Enabled = False
    ElseIf Me.DrillDown(Index).SelCount = 0 Then
        Me.AddItemAdjustments.Enabled = False
    ElseIf getDrillBy(0) = "ItemNumber" Then
        Me.AddItemAdjustments.Enabled = True
    Else
        Me.AddItemAdjustments.Enabled = False
    End If
    
    enableAdjustmentMaybe
End Sub

Private Sub DDListHeader_Click(Index As Integer)
    If Index = 2 And Me.JobList.SelCount > 1 And getDrillBy(0) = "ItemNumber" Then
        'do nothing
    Else
        Dim dd As Long
        If Index <= ddHeaderUBound(0) Then
            dd = 0
        ElseIf Index <= ddHeaderUBound(1) Then
            dd = 1
        Else
            dd = 2
        End If
        
        If Me.DDListHeader(Index).FontBold Then
            ddListDescending(dd) = Not ddListDescending(dd)
        Else
            Dim i As Long
            For i = ddHeaderLBound(dd) To ddHeaderUBound(dd)
                Me.DDListHeader(i).FontBold = False
            Next i
            Me.DDListHeader(Index).FontBold = True
            ddListDescending(dd) = False
        End If
        If Me.DrillDown(dd).ListCount > 0 Then
            requeryDrillDown dd
        End If
    End If
End Sub

Private Sub DrillDownSelect_Click()
    Dim i As Long, found As Boolean

    Dim s As Variant
    s = Split(Me.DrillDownSelect, " -> ")
    
    Dim l As Variant
    l = Split(lastDDS, " -> ")
    
    lastDDS = Me.DrillDownSelect
    
    If s(0) <> l(0) Then
        Me.DDListHeader(ddHeaderLBound(0)) = s(0)
        If s(0) = "ItemNumber" Then
            For i = ddHeaderLBound(0) + 2 To ddHeaderLBound(0) + 5
                Me.DDListHeader(i).Visible = True
            Next i
            Me.AddItemAdjustments.Enabled = Not badWhseCombo
        Else
            For i = ddHeaderLBound(0) + 2 To ddHeaderLBound(0) + 5
                If Me.DDListHeader(i).FontBold Then
                    Me.DDListHeader(i).FontBold = False
                    Me.DDListHeader(ddHeaderLBound(0)).FontBold = True
                    ddListDescending(0) = False
                End If
                Me.DDListHeader(i).Visible = False
            Next i
            Me.AddItemAdjustments.Enabled = False
        End If
    End If
    
    If s(1) <> l(1) Then
        Me.DDListHeader(ddHeaderLBound(1) + 1) = s(1)
        If s(1) = "ItemNumber" Then
            For i = ddHeaderLBound(1) + 1 To ddHeaderLBound(1) + 1 'just in case we add more columns
                If Me.DDListHeader(i).FontBold Then
                    Me.DDListHeader(i).FontBold = False
                    Me.DDListHeader(ddHeaderLBound(1)).FontBold = True
                    ddListDescending(1) = False
                End If
                Me.DDListHeader(i).Visible = False
            Next i
        Else
            For i = ddHeaderLBound(1) + 1 To ddHeaderLBound(1) + 1
                Me.DDListHeader(i).Visible = True
            Next i
        End If
    End If
    
    If s(2) <> l(2) Then
        Me.DDListHeader(ddHeaderLBound(2) + 1) = s(2)
        If s(2) = "ItemNumber" Then
            For i = ddHeaderLBound(2) + 1 To ddHeaderLBound(2) + 1
                If Me.DDListHeader(i).FontBold Then
                    Me.DDListHeader(i).FontBold = False
                    Me.DDListHeader(ddHeaderLBound(2)).FontBold = True
                    ddListDescending(2) = False
                End If
                Me.DDListHeader(i).Visible = False
            Next i
        Else
            For i = ddHeaderLBound(2) + 1 To ddHeaderLBound(2) + 1
                Me.DDListHeader(i).Visible = True
            Next i
        End If
    End If
    
    If Me.JobList.SelCount > 0 Then
        requeryDrillDown IIf(s(0) <> l(0), 0, IIf(s(1) <> l(1), 1, 2))
    End If
End Sub

Private Sub OffByCountFrom_Click(Index As Integer)
    If Not stopevents Then
        If getDrillBy(0) = "ItemNumber" Then
            If Me.OffByCountFrom(1).value = True Then
                requeryDrillDown 0
            Else
                If Me.JobList.SelCount > 1 Then
                    'no snapshots for multiple jobs
                    stopevents = True
                    Me.OffByCountFrom(1) = True
                    stopevents = False
                Else
                    requeryDrillDown 0
                End If
            End If
        End If
    End If
End Sub

Private Function ddHeaderLBound(dd As Long) As Long
    Select Case dd
        Case Is = 0
            ddHeaderLBound = 0
        Case Is = 1
            ddHeaderLBound = 6
        Case Is = 2
            ddHeaderLBound = 9
    End Select
End Function

Private Function ddHeaderUBound(dd As Long) As Long
    Select Case dd
        Case Is = 0
            ddHeaderUBound = 5
        Case Is = 1
            ddHeaderUBound = 8
        Case Is = 2
            ddHeaderUBound = 12
    End Select
End Function

Private Sub requeryDrillDown(dx As Long)
    Select Case dx
        Case Is = 0
            requeryDrillDown0
        Case Is = 1
            requeryDrillDown1
        Case Is = 2
            requeryDrillDown2
        Case Else
            MsgBox "ERROR"
    End Select
End Sub

Private Sub requeryDrillDown0()
    If Me.DDListHeader(2).FontBold And getDrillBy(0) = "ItemNumber" And Me.JobList.SelCount > 1 Then
        Me.DDListHeader(2).FontBold = False
        Me.DDListHeader(0).FontBold = True
        ddListDescending(0) = False
    End If

    Dim selectclause As String, groupbyclause As String, orderbyclause As String
    
    Select Case True
        Case Is = Me.DDListHeader(ddHeaderLBound(0) + 0).FontBold
            orderbyclause = " ORDER BY " & getDrillBy(0) & IIf(ddListDescending(0), " DESC", "")
        Case Is = Me.DDListHeader(ddHeaderLBound(0) + 1).FontBold
            orderbyclause = " ORDER BY SUM(CountQuantity)" & IIf(ddListDescending(0), " DESC", "")
        'the following should only happen w/ ItemNumber drill-by
        Case Is = Me.DDListHeader(ddHeaderLBound(0) + 2).FontBold
            orderbyclause = " ORDER BY SnapshotQuantity" & IIf(ddListDescending(0), " DESC", "")
        Case Is = Me.DDListHeader(ddHeaderLBound(0) + 3).FontBold
            orderbyclause = " ORDER BY CurrentQuantity" & IIf(ddListDescending(0), " DESC", "")
        Case Is = Me.DDListHeader(ddHeaderLBound(0) + 4).FontBold
            orderbyclause = " ORDER BY SUM(CountQuantity)-" & Switch(Me.OffByCountFrom(0), "Snapshot", Me.OffByCountFrom(1), "Current") & "Quantity" & IIf(ddListDescending(0), " DESC", "")
        Case Is = Me.DDListHeader(ddHeaderLBound(0) + 5).FontBold
            orderbyclause = " ORDER BY Adjusted" & IIf(ddListDescending(0), " DESC", "")
    End Select
    
    Select Case getDrillBy(0)
        Case Is = "ItemNumber"
            If Me.JobList.SelCount = 1 Then
                'selectclause = "SELECT ItemNumber, SUM(CountQuantity), SnapshotQuantity, CurrentQuantity, SUM(CountQuantity)-" & Switch(Me.OffByCountFrom(0), "Snapshot", Me.OffByCountFrom(1), "Current") & "Quantity, Adjusted"
                'groupbyclause = " GROUP BY ItemNumber, SnapshotQuantity, CurrentQuantity, Adjusted"
                selectclause = "SELECT ItemNumber, SUM(CountQuantity), SnapshotQuantity, CurrentQuantity, SUM(CountQuantity)-" & Switch(Me.OffByCountFrom(0), "Snapshot", Me.OffByCountFrom(1), "Current") & "Quantity, Adjusted"
                groupbyclause = " GROUP BY ItemNumber, SnapshotQuantity, CurrentQuantity, Adjusted"
            Else
                If Me.OffByCountFrom(0).value = True Then
                    stopevents = True
                    Me.OffByCountFrom(1) = True
                    stopevents = False
                End If
                selectclause = "SELECT ItemNumber, SUM(CountQuantity), '', CurrentQuantity, SUM(CountQuantity)-" & Switch(Me.OffByCountFrom(0), "Snapshot", Me.OffByCountFrom(1), "Current") & "Quantity, Adjusted"
                groupbyclause = " GROUP BY ItemNumber, CurrentQuantity, Adjusted"
            End If
        Case Is = "Location"
            selectclause = "SELECT Location, SUM(CountQuantity), '', ''"
            groupbyclause = " GROUP BY Location"
        Case Is = "Person"
            selectclause = "SELECT Person, SUM(CountQuantity), '', ''"
            groupbyclause = " GROUP BY Person"
    End Select
    
    Me.DrillDown(0).Clear
    Me.DrillDown(1).Clear
    Me.DrillDown(2).Clear
    Me.AddItemAdjustments.Enabled = False
    If Me.JobList.SelCount <> 0 Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve(selectclause & " FROM vCountMgmtDrillDown WHERE " & filterJobs() & groupbyclause & orderbyclause)
        If getDrillBy(0) = "ItemNumber" Or Me.JobList.SelCount = 1 Then
            Dim itemcache As Dictionary
            Set itemcache = New Dictionary
            While Not rst.EOF
                If itemcache.exists(CStr(rst("ItemNumber"))) Then
                    If itemcache.item(CStr(rst("ItemNumber"))) <> CStr(rst("Adjusted")) Then 'always true?
                        Dim i As Long
                        For i = 0 To Me.DrillDown(0).ListCount - 1
                            If getCell("ItemNumber", i) = rst("ItemNumber") Then
                                Dim templine As Variant
                                templine = Split(Me.DrillDown(0).list(i), vbTab)
                                templine(1) = CLng(templine(1)) + CLng(rst(1))
                                templine(4) = CLng(templine(1)) - CLng(templine(3))
                                templine(5) = "/"
                                Me.DrillDown(0).list(i) = Join(templine, vbTab)
                                Exit For
                            End If
                        Next i
                    End If
                Else
                    Me.DrillDown(0).AddItem Join(RSFieldsAsArray(rst.Fields), vbTab)
                    itemcache.Add CStr(rst("ItemNumber")), CStr(rst("Adjusted"))
                End If
                rst.MoveNext
            Wend
        Else
            While Not rst.EOF
                Me.DrillDown(0).AddItem Join(RSFieldsAsArray(rst.Fields), vbTab)
                rst.MoveNext
            Wend
        End If
        rst.Close
        Set rst = Nothing
    End If
End Sub

Private Sub requeryDrillDown1()
    Dim selectclause As String, groupbyclause As String, orderbyclause As String
    
    Select Case True
        Case Is = Me.DDListHeader(ddHeaderLBound(1) + 0).FontBold
            orderbyclause = " ORDER BY ItemNumber" & IIf(ddListDescending(1), " DESC", "")
        Case Is = Me.DDListHeader(ddHeaderLBound(1) + 1).FontBold
            orderbyclause = " ORDER BY " & getDrillBy(1) & IIf(ddListDescending(1), " DESC", "")
        Case Is = Me.DDListHeader(ddHeaderLBound(1) + 2).FontBold
            orderbyclause = " ORDER BY SUM(CountQuantity)" & IIf(ddListDescending(1), " DESC", "")
    End Select
    
    Select Case getDrillBy(1)
        Case Is = "ItemNumber"
            selectclause = "SELECT ItemNumber, '', SUM(CountQuantity)"
            groupbyclause = " GROUP BY ItemNumber"
        Case Is = "Location"
            selectclause = "SELECT ItemNumber, Location, SUM(CountQuantity)"
            groupbyclause = " GROUP BY ItemNumber, Location"
        Case Is = "Person"
            selectclause = "SELECT ItemNumber, Person, SUM(CountQuantity)"
            groupbyclause = " GROUP BY ItemNumber, Person"
    End Select
    
    Me.DrillDown(1).Clear
    Me.DrillDown(2).Clear
    If Me.DrillDown(0).SelCount <> 0 Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve(selectclause & " FROM vCountMgmtDrillDown WHERE " & filterJobs() & " AND " & filterDrillDown(0) & groupbyclause & orderbyclause)
        While Not rst.EOF
            Me.DrillDown(1).AddItem rst(0) & vbTab & rst(1) & vbTab & rst(2)
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
    End If
End Sub

Private Sub requeryDrillDown2()
    Dim selectclause As String, groupbyclause As String, orderbyclause As String
    
    Select Case True
        Case Is = Me.DDListHeader(ddHeaderLBound(2) + 0).FontBold
            orderbyclause = " ORDER BY ItemNumber" & IIf(ddListDescending(2), " DESC", "")
        Case Is = Me.DDListHeader(ddHeaderLBound(2) + 1).FontBold
            orderbyclause = " ORDER BY " & getDrillBy(2) & IIf(ddListDescending(2), " DESC", "")
        Case Is = Me.DDListHeader(ddHeaderLBound(2) + 2).FontBold
            orderbyclause = " ORDER BY SUM(CountQuantity)" & IIf(ddListDescending(2), " DESC", "")
        Case Is = Me.DDListHeader(ddHeaderLBound(2) + 3).FontBold
            orderbyclause = " ORDER BY DateOfEntry" & IIf(ddListDescending(2), " DESC", "")
    End Select
    
    Select Case getDrillBy(2)
        Case Is = "ItemNumber"
            selectclause = "SELECT ItemNumber, '', SUM(CountQuantity), DateOfEntry"
            groupbyclause = " GROUP BY ItemNumber, DateOfEntry"
        Case Is = "Location"
            selectclause = "SELECT ItemNumber, Location, SUM(CountQuantity), DateOfEntry"
            groupbyclause = " GROUP BY ItemNumber, Location, DateOfEntry"
        Case Is = "Person"
            selectclause = "SELECT ItemNumber, Person, SUM(CountQuantity), DateOfEntry"
            groupbyclause = " GROUP BY ItemNumber, Person, DateOfEntry"
    End Select
    
    Me.DrillDown(2).Clear
    If Me.DrillDown(1).SelCount <> 0 Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve(selectclause & " FROM vCountMgmtDrillDown WHERE " & filterJobs() & " AND " & filterDrillDown(0) & " AND " & filterDrillDown(1) & groupbyclause & orderbyclause)
        While Not rst.EOF
            Me.DrillDown(2).AddItem rst(0) & vbTab & rst(1) & vbTab & rst(2) & vbTab & rst(3)
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
    End If
End Sub

Private Function getDrillBy(dx As Long) As String
    Dim s As Variant
    s = Split(Me.DrillDownSelect, " -> ")
    getDrillBy = s(dx)
End Function

Private Function whichDrillByIs(drillby As String) As Long
    Dim s As Variant
    s = Split(Me.DrillDownSelect, " -> ")
    Dim i As Long
    whichDrillByIs = -1
    For i = 0 To UBound(s)
        If s(i) = drillby Then
            whichDrillByIs = i
            Exit For
        End If
    Next i
End Function

Private Function filterJobs() As String
    filterJobs = "JobID IN (" & ListAsCommaSep(ListBoxAsArray(Me.JobList, True, 2)) & ")"
End Function

Private Function filterDrillDown(dx As Long) As String
    'dx should be 0 or 1, 2 doesn't get filtered for another one
    Select Case getDrillBy(dx)
        Case Is = "ItemNumber"
            filterDrillDown = "ItemNumber IN (" & ListAsCommaSep(ListBoxAsArray(Me.DrillDown(dx), True, 0), "'", True) & ")"
        Case Is = "Location"
            'drilldown index happens to match column index, so use it, but this may need to be changed in the future
            filterDrillDown = "ItemNumber IN (" & ListAsCommaSep(ListBoxAsArray(Me.DrillDown(dx), True, 0), "'", True) & ")" & _
                         " AND Location IN (" & ListAsCommaSep(ListBoxAsArray(Me.DrillDown(dx), True, dx), "'", True) & ")"
        Case Is = "Person"
            filterDrillDown = "ItemNumber IN (" & ListAsCommaSep(ListBoxAsArray(Me.DrillDown(dx), True, 0), "'", True) & ")" & _
                         " AND Person IN (" & ListAsCommaSep(ListBoxAsArray(Me.DrillDown(dx), True, dx), "'", True) & ")"
    End Select
End Function

Private Sub AdjustEntry_Click()
'    Load InventoryCountMgmtAdjustDlg
'    Dim i As Long, dd As Long, JobID As String, item As String, Location As String
'    For i = 0 To Me.JobList.ListCount - 1
'        If Me.JobList.Selected(i) Then
'            JobID = Mid(Me.JobList.list(i), InStrRev(Me.JobList.list(i), vbTab) + 1)
'            Exit For
'        End If
'    Next i
'    dd = whichDrillByIs("ItemNumber")
'    For i = 0 To Me.DrillDown(dd).ListCount - 1
'        If Me.DrillDown(dd).Selected(i) Then
'            'adjust enablement is based on a single item being selected, so we can
'            'stop once we get one item here
'            item = getCell("ItemNumber", i)
'            Exit For
'        End If
'    Next i
'    dd = whichDrillByIs("Location")
'    Location = ""
'    If Me.DrillDown(dd).SelCount <> 0 Then
'        For i = 0 To Me.DrillDown(dd).ListCount - 1
'            If Me.DrillDown(dd).Selected(i) Then
'                If Location = "" Then
'                    Location = getCell("Location", i)
'                Else
'                    'multiple locations selected, default to none
'                    Location = ""
'                    Exit For
'                End If
'            End If
'        Next i
'    End If
'    InventoryCountMgmtAdjustDlg.LoadStuff JobID, item, Location
'    InventoryCountMgmtAdjustDlg.Show MODAL
'    MsgBox "TODO: The form should requery here, but it's complicated, so do it yourself."
    Load InventoryCountMgmtAdjustDlg
    Dim dd As Long, i As Long, item As String, Location As String
    dd = whichDrillByIs("ItemNumber")
    For i = 0 To Me.DrillDown(dd).ListCount - 1
        If Me.DrillDown(dd).Selected(i) Then
            'adjust enablement is based on a single item being selected, so we can
            'stop once we get one item here
            item = getCell("ItemNumber", i)
            Exit For
        End If
    Next i
    InventoryCountMgmtAdjustDlg.LoadStuff ListBoxAsArray(Me.JobList, True, 2), item
    InventoryCountMgmtAdjustDlg.Show MODAL
    MsgBox "TODO: The form should requery here, but it's complicated, so do it yourself."
End Sub

Private Sub enableAdjustmentMaybe()
    'If Me.JobList.SelCount = 1 Then
        Dim dd As Long
        dd = whichDrillByIs("ItemNumber")
        Select Case Me.DrillDown(dd).SelCount
            Case Is = 0
                Me.AdjustEntry.Enabled = False
            Case Is = 1
                Me.AdjustEntry.Enabled = True
            Case Else
                Dim i As Long
                Dim firstitem As String
                Me.AdjustEntry.Enabled = True
                'you shouldn't see two of the same item, wtf was i doing here?
                'TODO: fix this, maybe?
                For i = 0 To Me.DrillDown(dd).ListCount - 1
                    If Me.DrillDown(dd).Selected(i) Then
                        Dim curr As String
                        curr = getCell("ItemNumber", i)
                        If firstitem = "" Then
                            firstitem = curr
                        Else
                            If firstitem <> curr Then
                                Me.AdjustEntry.Enabled = False
                                Exit For
                            End If
                        End If
                    End If
                Next i
        End Select
    'Else
    '    Me.AdjustEntry.Enabled = False
    'End If
End Sub

Private Function getCell(drillby As String, rownum As Long) As String
    Dim dd As Long
    dd = whichDrillByIs(drillby)
    Dim s As Variant
    Select Case drillby
        Case Is = "ItemNumber"
            'itemnumber is always first in all drill downs
            getCell = Left(Me.DrillDown(dd).list(rownum), InStr(Me.DrillDown(dd).list(rownum), vbTab) - 1)
        Case Is = "Location"
            'first column if dd=0, second otherwise
            s = Split(Me.DrillDown(dd).list(rownum), vbTab)
            Select Case dd
                Case Is = 0
                    getCell = s(0)
                Case Else
                    getCell = s(1)
            End Select
        Case Is = "Person"
            'first column if dd=0, second otherwise
            s = Split(Me.DrillDown(dd).list(rownum), vbTab)
            Select Case dd
                Case Is = 0
                    getCell = s(0)
                Case Else
                    getCell = s(1)
            End Select
    End Select
End Function

'-------------------
' batch functions
'-------------------
Private Sub AddItemAdjustments_Click()
    If Me.DrillDown(0).SelCount > 0 And Me.BatchList.SelCount = 1 Then
        If vbYes = MsgBox("Set inventory quantity to count quantity for all selected items?", vbYesNo) Then
            Dim i As Long
            Dim JobID As String, batchid As String
            For i = 0 To Me.JobList.ListCount - 1
                If Me.JobList.Selected(i) Then
                    JobID = Mid(Me.JobList.list(i), InStrRev(Me.JobList.list(i), vbTab) + 1)
                    Exit For
                End If
            Next i
            batchid = Left(Me.BatchList, InStr(Me.BatchList, vbTab) - 1)
            'batchid = Mid(batchid, 3)
            For i = 0 To Me.DrillDown(0).ListCount - 1
                If Me.DrillDown(0).Selected(i) Then
                    Dim s As Variant, item As String, qty As String
                    s = Split(Me.DrillDown(0).list(i), vbTab)
                    item = s(0)
                    qty = s(1) - s(3)
                    'TODO: fix for jx table
                    If s(5) = " " Then
                        DB.Execute "INSERT INTO HandheldCountJobAdjustments ( JobID, BatchID, ItemNumber, AdjustmentQuantity ) VALUES ( " & JobID & ", '" & CStr(CLng(batchid)) & "', '" & item & "', " & qty & " )"
                        Me.DrillDown(0).list(i) = s(0) & vbTab & s(1) & vbTab & s(2) & vbTab & s(3) & vbTab & s(4) & vbTab & "R"
                    End If
                End If
            Next i
            requeryBatchList
        End If
    ElseIf Me.BatchList.SelCount <> 1 Then
        MsgBox "Select a batch first!"
    End If
End Sub

Private Sub BatchList_Click()
    Me.CommitBatch.Enabled = True
End Sub

Private Sub CommitBatch_Click()
    Dim batchnum As String, qty As String
    batchnum = Left(Me.BatchList, InStr(Me.BatchList, vbTab) - 1)
    qty = Mid(Me.BatchList, InStr(Me.BatchList, vbTab) + 1)
    If qty = 0 Then
        MsgBox "There are no items in this batch!"
    ElseIf vbYes = MsgBox("Send this batch now?", vbYesNo) Then
    
        DB.Execute "UPDATE HandheldCountJobAdjustments SET ReadyToExport=1, TransactionDate='" & Format(Date, "mm/dd/yyyy") & "' WHERE BatchID=" & batchnum
    
        Dim fso As FileSystemObject
        Set fso = New FileSystemObject
        fso.CreateTextFile MUTEX_INV_CHANGE

        ShellWait "s:\mastest\mas90-signs\mas200_import_adjustments.bat", vbHide
        
        While fso.FileExists(MUTEX_INV_CHANGE)
            DoEvents
            Sleep 5000
        Wend
        
        If vbYes = MsgBox("Did MAS200 finish the import successfully?", vbYesNo) Then
            DB.Execute "UPDATE HandheldCountJobAdjustments SET Exported=1, ReadyToExport=0 WHERE ReadyToExport=1"
            Me.BatchList.RemoveItem Me.BatchList.ListIndex
            requeryDrillDown 0
        Else
            MsgBox "That sucks, Brian should probably fix this."
        End If
    End If
End Sub

Private Sub CreateNewBatch_Click()
    If vbYes = MsgBox("Create new batch?", vbYesNo) Then
        Dim batchid As String
        batchid = DLookup("MAX(BatchID)", "HandheldCountJobAdjustments")
        If batchid = "" Then
            batchid = 0
        End If
        batchid = batchid + 1
        If Me.BatchList.ListCount = 0 Then
            Me.BatchList.AddItem Format(batchid, "000000") & vbTab & "0"
        Else
            Dim temp As String
            temp = Me.BatchList.list(Me.BatchList.ListCount - 1)
            temp = Left(temp, InStr(temp, vbTab) - 1)
            'temp = Mid(temp, 3)
            Dim lastbatch As Long
            lastbatch = CLng(temp)
            If lastbatch = batchid Then
                MsgBox "An empty batch already exists!"
                Me.BatchList.Selected(Me.BatchList.ListCount - 1) = True
            Else
                Me.BatchList.AddItem Format(batchid, "000000") & vbTab & "0"
            End If
        End If
    End If
End Sub


'------------------
' new job area
'------------------
Private Sub NewJobSpecificLocation_Click()
    If Me.NewJobSpecificLocation Then
        Enable Me.NewJobLocation
    Else
        Disable Me.NewJobLocation
    End If
End Sub

Private Sub NewJobWarehouse_Click()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Name FROM WarehouseLocations WHERE WarehouseID=(SELECT ID FROM Warehouses WHERE Name='" & EscapeSQuotes(Me.NewJobWarehouse) & "')")
    Me.NewJobLocation.Clear
    While Not rst.EOF
        Me.NewJobLocation.AddItem rst("Name")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub NewJobCreate_Click()
    If Me.NewJobName = "" Then
        MsgBox "You must enter a description for this job!"
    ElseIf Me.NewJobWarehouse = "" Then
        MsgBox "You must select a warehouse for this job!"
    ElseIf Me.NewJobSpecificLocation And Me.NewJobLocation = "" Then
        MsgBox "If ""Specific Location"" is checked, you must select a location!"
    Else
        Dim i As Long
        Dim wid As String, lid As String, pid As String, Lines As String
        wid = DLookup("ID", "Warehouses", "Name='" & EscapeSQuotes(Me.NewJobWarehouse) & "'")
        If Me.NewJobSpecificLocation Then
            lid = DLookup("ID", "WarehouseLocations", "Name='" & EscapeSQuotes(Me.NewJobLocation) & "' AND WarehouseID=" & wid)
        Else
            lid = "NULL"
        End If
        pid = DLookup("ID", "Users", "NTUsername='" & EscapeSQuotes(Environ("UserName")) & "'")
        If Me.NewJobLines.SelCount = 0 Then
            Lines = "NULL"
        Else
            Lines = "'" & Replace(ListAsCommaSep(ListBoxAsArray(Me.NewJobLines, True)), " ", "") & "'"
        End If
        DB.Execute "INSERT INTO HandheldCountJobs ( Name, Notes, AssignedByID, WarehouseID, LocationID, LineItemFilter ) VALUES ( '" & EscapeSQuotes(Me.NewJobName) & "', '" & EscapeSQuotes(Me.NewJobNotes) & "', " & pid & ", " & wid & ", " & lid & ", " & Lines & " )"
        
        Dim newid As String
        newid = DLookup("@@IDENTITY", "HandheldCountJobs")
        
        If Me.NewJobAssignment.SelCount > 0 Then
            For i = 0 To Me.NewJobAssignment.ListCount - 1
                If Me.NewJobAssignment.Selected(i) Then
                    DB.Execute "INSERT INTO HandheldCountJobAssignments ( JobID, UserID ) VALUES ( " & newid & ", " & Mid(Me.NewJobAssignment.list(i), InStr(Me.NewJobAssignment.list(i), vbTab) + 1) & " )"
                End If
            Next i
        End If
        
        Me.NewJobName = ""
        Me.NewJobNotes = ""
        Me.NewJobWarehouse.ListIndex = -1
        Me.NewJobLocation.ListIndex = -1
        Me.NewJobSpecificLocation = 0
        For i = 0 To Me.NewJobAssignment.ListCount - 1
            If Me.NewJobAssignment.Selected(i) Then
                Me.NewJobAssignment.Selected(i) = False
            End If
        Next i
        For i = 0 To Me.NewJobLines.ListCount - 1
            If Me.NewJobLines.Selected(i) Then
                Me.NewJobLines.Selected(i) = False
            End If
        Next i
        
        createJobSnapshot newid, Split(IIf(Lines = "NULL", "", Replace(Lines, "'", "")), ",")
        
        requeryJobList
    End If
End Sub

Private Sub createJobSnapshot(JobID As String, Lines As Variant)
    Dim whse As String
    whse = DLookup("MasWhseCode", "Warehouses INNER JOIN HandheldCountJobs ON Warehouses.ID=WarehouseID", "HandheldCountJobs.ID=" & JobID)
    
    DB.Execute "INSERT INTO HandheldCountJobInventorySnapshots (JobID, ItemNumber, WhseCode, Quantity ) " & _
               "SELECT " & JobID & ", ItemNumber, WhseCode, QtyOnHand " & _
               "FROM InventoryQuantities " & _
               "WHERE WhseCode='" & whse & "' " & _
                 "AND QuantityOnHand<>0 " & _
                 "AND ItemNumber<'XX%' " & _
                 IIf(UBound(Lines) = -1, "", " AND LEFT(ItemNumber,3) IN ( '" & Join(Lines, "', '") & "' )")
    DB.Execute "UPDATE HandheldCountJobs SET SnapshotTime=GETDATE() WHERE JobID=" & JobID
End Sub

'-------------------
' data filling
'-------------------
Private Sub requeryNewJobAssignment()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ShortName, PersonID FROM vPermissions WHERE ModuleName='HandheldCount'")
    Me.NewJobAssignment.Clear
    While Not rst.EOF
        Me.NewJobAssignment.AddItem rst("ShortName") & vbTab & rst("PersonID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryNewJobLines()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ProductLine FROM ProductLine WHERE Hide=0")
    Me.NewJobLines.Clear
    While Not rst.EOF
        Me.NewJobLines.AddItem rst("ProductLine")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryNewJobWarehouses()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Name FROM Warehouses")
    Me.NewJobWarehouse.Clear
    While Not rst.EOF
        Me.NewJobWarehouse.AddItem rst("Name")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryDrillDownSelect()
    Me.DrillDownSelect.Clear
    Me.DrillDownSelect.AddItem "ItemNumber -> Location -> Person"
    Me.DrillDownSelect.AddItem "ItemNumber -> Person -> Location"
    Me.DrillDownSelect.AddItem "Location -> ItemNumber -> Person"
    Me.DrillDownSelect.AddItem "Location -> Person -> ItemNumber"
    Me.DrillDownSelect.AddItem "Person -> ItemNumber -> Location"
    Me.DrillDownSelect.AddItem "Person -> Location -> ItemNumber"
    lastDDS = Me.DrillDownSelect.list(0)
    Me.DrillDownSelect = Me.DrillDownSelect.list(0)
End Sub

Private Sub requeryBatchList()
    Dim sel As String
    If Me.BatchList.SelCount <> 0 Then
        sel = Left(Me.BatchList, InStr(Me.BatchList, vbTab) - 1)
    End If
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT BatchID, COUNT(*) FROM HandheldCountJobAdjustments WHERE Exported=0 GROUP BY BatchID")
    Me.BatchList.Clear
    Me.CommitBatch.Enabled = False
    While Not rst.EOF
        Me.BatchList.AddItem Format(rst(0), "000000") & vbTab & rst(1)
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Me.CommitBatch.Enabled = False
    If sel <> "" Then
        Dim i As Long
        For i = 0 To Me.BatchList.ListCount - 1
            If Left(Me.BatchList.list(i), InStr(Me.BatchList.list(i), vbTab) - 1) = sel Then
                Me.BatchList.Selected(i) = True
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub ViewChangesSinceSnapshot_Click()
    If getDrillBy(0) <> "ItemNumber" Then
        MsgBox "Must drill down by itemnumber first!"
    ElseIf Me.JobList.SelCount = 0 Then
        MsgBox "You must select a job!"
    ElseIf Me.JobList.SelCount > 1 Then
        MsgBox "Can only select one job at a time!"
    ElseIf Me.DrillDown(0).SelCount = 0 Then
        MsgBox "You must select an item!"
    ElseIf Me.DrillDown(0).SelCount > 1 Then
        MsgBox "Can only select one item at a time!"
    Else
        Dim item As String, JobID As String
        item = Left(Me.DrillDown(0), InStr(Me.DrillDown(0), vbTab) - 1)
        JobID = Mid(Me.JobList, InStrRev(Me.JobList, vbTab) + 1)
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT * FROM IM5_TransactionDetail WHERE ItemNumber='" & item & "' AND TransactionDate-1>(SELECT SnapshotTime FROM HandheldCountJobs WHERE ID=" & JobID & ")")
        Dim msg As String
        While Not rst.EOF
            msg = msg & rst("TransactionDate") & vbTab & rst("WarehouseCode") & vbTab & rst("TransactionCode") & vbTab & rst("TransactionQty") & vbCrLf
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
        MsgBox IIf(msg = "", "No changes for this item!", msg)
    End If
End Sub
