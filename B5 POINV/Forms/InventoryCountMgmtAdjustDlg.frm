VERSION 5.00
Begin VB.Form InventoryCountMgmtAdjustDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adjust Entry - Inventory Count Management"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox JobName 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   300
      Width           =   2655
   End
   Begin VB.TextBox AdjustTo 
      Height          =   315
      Left            =   960
      TabIndex        =   10
      Top             =   1860
      Width           =   1185
   End
   Begin VB.ComboBox Location 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   525
      Left            =   1950
      TabIndex        =   1
      Top             =   2490
      Width           =   1485
   End
   Begin VB.CommandButton AcceptButton 
      Caption         =   "Accept Change"
      Height          =   525
      Left            =   390
      TabIndex        =   0
      Top             =   2490
      Width           =   1485
   End
   Begin VB.Label generalLabel 
      Caption         =   "Adjust To:"
      Height          =   315
      Index           =   4
      Left            =   150
      TabIndex        =   9
      Top             =   1860
      Width           =   735
   End
   Begin VB.Label generalLabel 
      Caption         =   "Location:"
      Height          =   315
      Index           =   3
      Left            =   150
      TabIndex        =   7
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label generalLabel 
      Caption         =   "Count:"
      Height          =   315
      Index           =   2
      Left            =   150
      TabIndex        =   6
      Top             =   1470
      Width           =   735
   End
   Begin VB.Label CurrentCountLabel 
      Caption         =   "CURRENT_COUNT_LABEL"
      Height          =   315
      Left            =   960
      TabIndex        =   5
      Top             =   1470
      Width           =   2655
   End
   Begin VB.Label ItemLabel 
      Caption         =   "ITEM_LABEL"
      Height          =   315
      Left            =   960
      TabIndex        =   4
      Top             =   690
      Width           =   2655
   End
   Begin VB.Label generalLabel 
      Caption         =   "Item:"
      Height          =   315
      Index           =   1
      Left            =   150
      TabIndex        =   3
      Top             =   690
      Width           =   735
   End
   Begin VB.Label generalLabel 
      Caption         =   "Job:"
      Height          =   315
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   300
      Width           =   735
   End
End
Attribute VB_Name = "InventoryCountMgmtAdjustDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private locIDCache As Dictionary
Private countCache As Dictionary

Private jobListMap As Dictionary
Private locListMap As Dictionary
Private locNameMap As Dictionary

Private savedjobid As String

'Public Sub LoadStuff(JobID As String, item As String, Optional loc As String = "")
'    savedjobid = JobID
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT WarehouseID, Location, CountQuantity FROM vCountMgmtDrillDown WHERE JobID=" & JobID & " AND ItemNumber='" & item & "'")
'    Set locIDCache = New Dictionary
'    Set countCache = New Dictionary
'    Dim rst2 As ADODB.Recordset
'    Set rst2 = DB.retrieve("SELECT ID, Name FROM HandheldWarehouseLocations WHERE WarehouseID=" & rst("WarehouseID"))
'    While Not rst2.EOF
'        locIDCache.Add CStr(rst2("Name")), CStr(rst2("ID"))
'        countCache.Add CStr(rst2("Name")), 0
'        Me.location.AddItem rst2("Name")
'        rst2.MoveNext
'    Wend
'    rst2.Close
'    Set rst2 = Nothing
'    While Not rst.EOF
'        countCache.item(CStr(rst("Location"))) = countCache.item(CStr(rst("Location"))) + rst("CountQuantity")
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'
'    Me.JobLabel.caption = DLookup("Name", "HandheldCountJobs", "ID=" & JobID)
'    Me.ItemLabel.caption = item
'    If loc = "" Then
'        Me.CurrentCountLabel.caption = ""
'    Else
'        Me.location = loc
'        Me.CurrentCountLabel.caption = countCache.item(loc)
'        Me.AdjustTo = countCache.item(loc)
'    End If
Public Sub LoadStuff(JobID() As String, item As String)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, Name FROM HandheldCountJobs WHERE ID IN ( " & Join(JobID, ", ") & " )")
    Set jobListMap = New Dictionary
    Dim i As Long
    i = 0
    While Not rst.EOF
        Me.JobName.AddItem rst("Name")
        jobListMap.Add i, CStr(rst("ID"))
        rst.MoveNext
        i = i + 1
    Wend
    rst.Close
    Me.ItemLabel.caption = item
    Me.JobName.ListIndex = 0 'calls the rest of the updating functions
End Sub

Private Sub AcceptButton_Click()
    If Me.Location.ListIndex = -1 Then
        MsgBox "No location selected!"
    ElseIf Me.AdjustTo = "" Then
        MsgBox "No amount entered!"
    ElseIf Me.AdjustTo = Me.CurrentCountLabel.caption Then
        Unload Me
    Else
        DB.Execute "INSERT INTO HandheldCountJobDetailLines ( JobID, ItemNumber, LocationID, PersonID, Quantity ) VALUES ( " & jobListMap.item(Me.JobName.ListIndex) & ", '" & Me.ItemLabel.caption & "', " & locListMap.item(Me.Location.ListIndex) & ", " & GetCurrentUserID() & ", " & CLng(Me.AdjustTo) - CLng(Me.CurrentCountLabel.caption) & " )"
        Unload Me
    End If
End Sub

Private Sub AdjustTo_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set jobListMap = Nothing
    Set locListMap = Nothing
    Set locNameMap = Nothing
    Set countCache = Nothing
End Sub

Private Sub JobName_Click()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT LocationID, Location, SUM(CountQuantity) AS Qty FROM vCountMgmtDrillDown WHERE JobID=" & jobListMap.item(Me.JobName.ListIndex) & " AND ItemNumber='" & Me.ItemLabel.caption & "' GROUP BY LocationID, Location")
    Set locListMap = New Dictionary
    Set locNameMap = New Dictionary
    Set countCache = New Dictionary
    Dim i As Long
    i = 0
    Me.Location.Clear
    While Not rst.EOF
        locListMap.Add i, CStr(rst("LocationID"))
        locNameMap.Add i, CStr(rst("Location"))
        countCache.Add i, CStr(rst("Qty"))
        Me.Location.AddItem rst("Location")
        rst.MoveNext
        i = i + 1
    Wend
    rst.Close
    Set rst = Nothing
    If Me.Location.ListCount = 0 Then
        Me.CurrentCountLabel.caption = ""
        Me.AdjustTo = ""
    Else
        Me.Location.ListIndex = 0
    End If
End Sub

Private Sub Location_Click()
    Me.CurrentCountLabel.caption = countCache.item(Me.Location.ListIndex)
    Me.AdjustTo = countCache.item(Me.Location.ListIndex)
End Sub
