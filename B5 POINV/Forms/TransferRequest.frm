VERSION 5.00
Object = "{39C208C4-2615-4D49-911A-50F903B45C86}#5.2#0"; "TPControls.ocx"
Begin VB.Form TransferRequest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer Request"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OpenReconciliation 
      Caption         =   "Transfer Reconciliation"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2603
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5580
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Caption         =   "Request a new item:"
      Height          =   2085
      Left            =   330
      TabIndex        =   13
      Top             =   3420
      Width           =   6075
      Begin VB.TextBox Notes 
         Height          =   285
         Left            =   780
         TabIndex        =   4
         Top             =   1320
         Width           =   2685
      End
      Begin VB.ComboBox ItemPriority 
         Height          =   315
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   1365
      End
      Begin VB.ComboBox ItemNumber 
         Height          =   315
         Left            =   780
         TabIndex        =   1
         Top             =   270
         Width           =   2145
      End
      Begin VB.TextBox Quantity 
         Height          =   285
         Left            =   780
         TabIndex        =   2
         Top             =   630
         Width           =   825
      End
      Begin VB.CommandButton AddRequest 
         Caption         =   "Add Request"
         Height          =   315
         Left            =   570
         TabIndex        =   5
         Top             =   1680
         Width           =   1845
      End
      Begin VB.CheckBox IncludeDCItems 
         Caption         =   "Include Discontinued Items?"
         Height          =   195
         Left            =   3330
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   300
         Width           =   2325
      End
      Begin VB.Label generalLabel 
         Caption         =   "Notes:"
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   21
         Top             =   1350
         Width           =   705
      End
      Begin VB.Label lblQtyAvail 
         Caption         =   "QTY"
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   660
         Width           =   705
      End
      Begin VB.Label generalLabel 
         Caption         =   "Qty Avail:"
         Height          =   255
         Index           =   6
         Left            =   1800
         TabIndex        =   19
         Top             =   660
         Width           =   705
      End
      Begin VB.Label generalLabel 
         Caption         =   "Priority:"
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   18
         Top             =   990
         Width           =   705
      End
      Begin VB.Label generalLabel 
         Caption         =   "Item:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   525
      End
      Begin VB.Label generalLabel 
         Caption         =   "Quantity:"
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   15
         Top             =   660
         Width           =   705
      End
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Close"
      Height          =   495
      Left            =   4980
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5580
      Width           =   1485
   End
   Begin VB.CommandButton RefreshList 
      Caption         =   "Refresh List"
      Height          =   495
      Left            =   270
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5580
      Width           =   1485
   End
   Begin VB.ComboBox Location 
      Height          =   315
      Left            =   4170
      Style           =   2  'Dropdown List
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   60
      Width           =   2445
   End
   Begin VB.ComboBox RemoteLocation 
      Height          =   315
      Left            =   4170
      Style           =   2  'Dropdown List
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   420
      Width           =   2445
   End
   Begin TPControls.SimpleListView RequestList 
      Height          =   2475
      Left            =   60
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   840
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   4366
   End
   Begin VB.Label lblLocation 
      Caption         =   "Your Location:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2850
      TabIndex        =   9
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label lblRemoteLocation 
      Caption         =   "Warehouse:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2850
      TabIndex        =   8
      Top             =   480
      Width           =   1305
   End
   Begin VB.Label generalLabel 
      Caption         =   "Transfer Request"
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
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   2505
   End
End
Attribute VB_Name = "TransferRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents refreshTimer As Timer
Attribute refreshTimer.VB_VarHelpID = -1
Private refreshTimerMinutes As Long

Private locIDMap As Dictionary

Private Sub refreshTimer_Timer()
    refreshTimerMinutes = refreshTimerMinutes - 1
    If refreshTimerMinutes = 0 Then
        If Me.RequestList.SelIndex = -1 Then
            requeryRequestList ""
        Else
            requeryRequestList Me.RequestList.GetTextSelected(0)
        End If
        refreshTimerMinutes = 5
    End If
End Sub

Private Sub AddRequest_Click()
    If Me.itemNumber = "" Then
        MsgBox "Select an item!"
    ElseIf Me.Quantity = "" Then
        MsgBox "Enter a quantity!"
    ElseIf CLng(Me.Quantity) <= 0 Then
        MsgBox "Quantity must be positive, you idiot!"
    ElseIf Me.Location = "" Then
        MsgBox "You must select your location above!"
    ElseIf Me.RemoteLocation = "" Then
        MsgBox "You must select the warehouse location!" & vbCrLf & vbCrLf & "It'd be nice if this auto-selected a warehouse based on quantity available and distance."
    Else
        If vbYes = MsgBox("Request " & Me.Quantity & " " & Me.itemNumber & " to be shipped from " & Me.RemoteLocation & " to " & Me.Location & "?", vbYesNo) Then
            Mouse.Hourglass True
            DB.Execute "INSERT INTO HandheldTransfers ( LineType, ItemNumber, Priority, Quantity, PersonID, FromWhseID, ToWhseID, Notes ) VALUES ( 0, '" & Me.itemNumber & "', " & Me.ItemPriority.ListIndex & ", " & Me.Quantity & ", " & GetCurrentUserID() & ", " & locIDMap(Me.RemoteLocation.ListIndex) & ", " & locIDMap(Me.Location.ListIndex) & ", '" & EscapeSQuotes(Left(Me.Notes, 64)) & "' )"
            Me.itemNumber = ""
            Me.lblQtyAvail.caption = ""
            Me.Quantity = ""
            Me.ItemPriority.ListIndex = 0
            Me.Notes = ""
            requeryRequestList
            Mouse.Hourglass False
        End If
    End If
End Sub

Private Sub Location_Click()
    requeryRequestList
    Me.OpenReconciliation.Enabled = IsSupervisorForWhse(locIDMap.item(Me.Location.ListIndex))
    setLocationLabelWarnings Me.lblLocation, CBool(Me.Location.ListIndex = -1)
End Sub

Private Sub OpenReconciliation_Click()
    Load TransferReconciliation
    TransferReconciliation.Show
End Sub

Private Sub RemoteLocation_Click()
    requeryRequestList
    setLocationLabelWarnings Me.lblRemoteLocation, CBool(Me.RemoteLocation.ListIndex = -1)
End Sub

Private Sub RefreshList_Click()
    requeryRequestList
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Mouse.Hourglass True
    
    Set refreshTimer = Controls.Add("VB.Timer", "refreshTimer")
    refreshTimer.Interval = 60000
    refreshTimerMinutes = 5
    
    Set locIDMap = New Dictionary
    Me.RequestList.SetColumnNames Array("Item Number", "*", "Quantity", "Loaded", "Unloaded")
    Me.RequestList.SetColumnWidths Array("2000", "250", "1000", "1000", "1000")
    requeryLocations
    requeryItems
    requeryPriorities
    Dim locID As String
    locID = DLookup("DefaultWhse", "Users", "ID=" & GetCurrentUserID())
    Me.Visible = True
    If setLocationMaybe(locID) Then
        If setWarehouseMaybe(locID) Then
            Me.RefreshList.SetFocus
        Else
            Me.RemoteLocation.SetFocus
        End If
    Else
        Me.Location.SetFocus
    End If

    refreshTimer.Enabled = True
    
    Mouse.Hourglass False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set locIDMap = Nothing
End Sub

Private Sub IncludeDCItems_Click()
    requeryItems
End Sub

Private Sub requeryLocations()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Name, ID FROM Warehouses")
    Me.Location.Clear
    Me.RemoteLocation.Clear
    Dim i As Long
    i = 0
    While Not rst.EOF
        locIDMap.Add i, CStr(rst("ID"))
        Me.Location.AddItem rst("Name")
        Me.RemoteLocation.AddItem rst("Name")
        i = i + 1
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requeryItems()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber FROM " & IIf(Me.IncludeDCItems, "PartNumbers", "vNonDiscontinuedItems") & " ORDER BY ItemNumber")
    Me.itemNumber.Clear
    Me.lblQtyAvail.caption = ""
    While Not rst.EOF
        Me.itemNumber.AddItem rst("ItemNumber")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requeryPriorities()
    Me.ItemPriority.Clear
    Me.ItemPriority.AddItem "Normal"
    Me.ItemPriority.AddItem "High"
    Me.ItemPriority.ListIndex = 0
End Sub

Private Sub requeryRequestList(Optional selItem As String = "")
    If Me.Location <> "" And Me.RemoteLocation <> "" Then
        Mouse.Hourglass True
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("EXEC spTransferRequests " & locIDMap(Me.RemoteLocation.ListIndex) & ", " & locIDMap(Me.Location.ListIndex))
        Me.RequestList.Clear
        Me.RequestList.Add rst, , True
        rst.Close
        Set rst = Nothing
        If selItem <> "" Then
            Dim i As Long
            For i = 0 To Me.RequestList.ListCount - 1
                If Me.RequestList.GetText(i, 0) = selItem Then
                    Me.RequestList.SelectRow i
                    Exit For
                End If
            Next i
        End If
        Mouse.Hourglass False
        refreshTimer.Enabled = False
        refreshTimer.Enabled = True
        refreshTimerMinutes = 5
    End If
End Sub

Private Sub ItemNumber_Click()
    ItemNumber_LostFocus
End Sub

Private Sub ItemNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.itemNumber, KeyCode, Shift
End Sub

Private Sub ItemNumber_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.itemNumber, KeyAscii
End Sub

Private Sub ItemNumber_LostFocus()
    AutoCompleteLostFocus Me.itemNumber
    If Me.itemNumber = "" Then
        Me.lblQtyAvail.caption = ""
    Else
        Me.lblQtyAvail.caption = DLookup("QuantityOnHand", "InventoryQuantities", "ItemNumber='" & Me.itemNumber & "' AND WhseCode='" & DLookup("MasWhseCode", "Warehouses", "ID=" & locIDMap.item(Me.RemoteLocation.ListIndex)) & "'")
    End If
End Sub

Private Sub Quantity_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
End Sub

Private Function setLocationMaybe(locID As String) As Boolean
    setLocationMaybe = False
    If locID <> "" Then
        Dim i As Long
        For i = 0 To Me.Location.ListCount - 1
            If locIDMap.item(i) = locID Then
                Me.Location.ListIndex = i
                setLocationMaybe = True
                Exit For
            End If
        Next i
    End If
    setLocationLabelWarnings Me.lblLocation, CBool(Me.Location.ListIndex = -1)
End Function

Private Function setWarehouseMaybe(storeID As String) As Boolean
    Dim locID As String
    locID = DLookup("DefaultFillWhse", "Warehouses", "ID=" & storeID)
    setWarehouseMaybe = False
    If locID <> "" Then
        Dim i As Long
        For i = 0 To Me.RemoteLocation.ListCount - 1
            If locIDMap.item(i) = locID Then
                Me.RemoteLocation.ListIndex = i
                setWarehouseMaybe = True
                Exit For
            End If
        Next i
    End If
    setLocationLabelWarnings Me.lblRemoteLocation, CBool(Me.RemoteLocation.ListIndex = -1)
End Function

Private Sub RequestList_DblClick(i As Long, j As Long, txt As String)
    Select Case j
        Case Is = 0
            Mouse.Hourglass True
            Load TransferRequestHist
            TransferRequestHist.SetRequestHistoryParams txt, locIDMap.item(Me.RemoteLocation.ListIndex), locIDMap.item(Me.Location.ListIndex)
            Mouse.Hourglass False
            TransferRequestHist.Show MODAL
    End Select
End Sub

Private Sub setLocationLabelWarnings(lbl As Label, onoff As Boolean)
    If onoff Then
        lbl.FontBold = True
        lbl.ForeColor = RED
    Else
        lbl.FontBold = False
        lbl.ForeColor = BLACK
    End If
End Sub
