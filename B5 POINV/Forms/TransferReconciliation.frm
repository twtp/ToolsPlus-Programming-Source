VERSION 5.00
Object = "{39C208C4-2615-4D49-911A-50F903B45C86}#5.2#0"; "TPControls.ocx"
Begin VB.Form TransferReconciliation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer Reconciliation"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton MasTransfers 
      Caption         =   "MAS TRANSFER CAPTION"
      Height          =   435
      Left            =   2700
      TabIndex        =   14
      Top             =   5790
      Width           =   4545
   End
   Begin VB.CommandButton ArchiveOldRequests 
      Caption         =   "Archive Old Requests"
      Height          =   435
      Left            =   360
      TabIndex        =   13
      Top             =   5790
      Width           =   1815
   End
   Begin VB.CommandButton AutoReconcile 
      Caption         =   "Auto-Reconcile"
      Height          =   435
      Left            =   7770
      TabIndex        =   9
      Top             =   1350
      Width           =   1815
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Close"
      Height          =   435
      Left            =   8100
      TabIndex        =   8
      Top             =   30
      Width           =   1485
   End
   Begin TPControls.SimpleListView TransferList 
      Height          =   3675
      Left            =   60
      TabIndex        =   7
      Top             =   1830
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   6482
   End
   Begin VB.ComboBox Truck 
      Height          =   315
      Left            =   1110
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1380
      Width           =   2295
   End
   Begin VB.ComboBox RemoteLocation 
      Height          =   315
      Left            =   1110
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1020
      Width           =   2295
   End
   Begin VB.ComboBox Location 
      Height          =   315
      Left            =   1110
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   660
      Width           =   2295
   End
   Begin VB.Label lblNotifyLocation 
      Caption         =   "<-- SELECT A LOCATION"
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
      Left            =   3480
      TabIndex        =   12
      Top             =   690
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Label lblNotifyWhse 
      Caption         =   "<-- SELECT A WAREHOUSE"
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
      Left            =   3480
      TabIndex        =   11
      Top             =   1050
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Label lblNotifyTruck 
      Caption         =   "<-- SELECT A TRUCK"
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
      Left            =   3480
      TabIndex        =   10
      Top             =   1410
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Label generalLabel 
      Caption         =   "Truck:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label generalLabel 
      Caption         =   "Warehouse:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label generalLabel 
      Caption         =   "Location:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   945
   End
   Begin VB.Label generalLabel 
      Caption         =   "Transfer Reconciliation"
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
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   3165
   End
End
Attribute VB_Name = "TransferReconciliation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private locIDMap As Dictionary
Private whseCodeMap As Dictionary
Private truckIDMap As Dictionary

Private Const MAS_CAPTION As String = "Ready to transfer in MAS 200: %SKU_COUNT% SKUs, %PIECE_COUNT% eaches"

Private Const COL_ITEM As Long = 0
Private Const COL_PRIO As Long = 1
Private Const COL_REQD As Long = 2
Private Const COL_LOAD As Long = 3
Private Const COL_UNLD As Long = 4
Private Const COL_VARI As Long = 5
Private Const COL_DIFF As Long = 6
Private Const COL_RECN As Long = 7


Private Sub ArchiveOldRequests_Click()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT DISTINCT ItemNumber FROM HandheldTransfers ORDER BY ItemNumber")
    While Not rst.EOF
        Dim rst2 As ADODB.Recordset
        Set rst2 = DB.retrieve("SELECT SUM(CASE WHEN LineType=0 THEN Quantity WHEN LineType=3 THEN -1*Quantity END) AS QtyReq, SUM(CASE WHEN LineType=1 THEN Quantity WHEN LineType=3 THEN -1*Quantity END) AS QtyLoad, SUM(CASE WHEN LineType=2 THEN Quantity WHEN LineType=3 THEN -1*Quantity END) AS QtyUnload FROM HandheldTransfers WHERE ItemNumber='" & rst("ItemNumber") & "'")
        If Not rst2.EOF Then
            If Nz(rst2("QtyReq"), "0") = 0 And Nz(rst2("QtyLoad"), "0") = 0 And Nz(rst2("QtyUnload"), "0") = 0 Then
                DB.Execute "UPDATE HandheldTransfers SET ArchiveFlag=1 WHERE ItemNumber='" & rst("ItemNumber") & "'"
            End If
        End If
        rst2.Close
        Set rst2 = Nothing
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub AutoReconcile_Click()
    Dim i As Long
    Dim thisln As Variant
    For i = 0 To Me.TransferList.ListCount - 1
        thisln = Me.TransferList.GetRow(i)
        If thisln(COL_LOAD) = thisln(COL_UNLD) Then
            If CLng(thisln(COL_REQD)) = CLng(thisln(COL_LOAD)) Then
                'ok to reconcile
                reconcileLine CStr(thisln(COL_ITEM)), CStr(thisln(COL_PRIO)), CStr(thisln(COL_LOAD))
            ElseIf CLng(thisln(COL_REQD)) > CLng(thisln(COL_LOAD)) Then
                'still more requested, ask what to do
                If vbYes = MsgBox("Requested quantity for " & thisln(COL_ITEM) & " not met, roll remaining quantity over until next time?", vbYesNo) Then
                    reconcileLine CStr(thisln(COL_ITEM)), CStr(thisln(COL_PRIO)), CStr(thisln(COL_LOAD)), True
                End If
            Else 'reqd < load
                'need to adjust req, then ok
                If vbYes = MsgBox("Adjust request quantity to match loaded quantity for " & thisln(COL_ITEM) & "?", vbYesNo) Then
                    adjustRequest CStr(thisln(COL_ITEM)), CStr(thisln(COL_PRIO)), CStr(CLng(thisln(COL_LOAD)) - CLng(thisln(COL_REQD)))
                    reconcileLine CStr(thisln(COL_ITEM)), CStr(thisln(COL_PRIO)), CStr(thisln(COL_LOAD)), True
                End If
            End If
        Else
            'can't auto-reconcile, do nothing
        End If
    Next i
    requeryTransferList
    setMasCaption
End Sub

Private Sub Form_Load()
    Mouse.Hourglass True
    Set locIDMap = New Dictionary
    Set whseCodeMap = New Dictionary
    Set truckIDMap = New Dictionary
    requeryLocations
    Me.Location = TransferRequest.Location
    If TransferRequest.RemoteLocation <> "" Then
        Me.RemoteLocation = TransferRequest.RemoteLocation
    End If
    requeryTrucks
    Me.TransferList.SetColumnNames Array("ItemNumber", "Priority", "Qty Requested", "Qty Loaded", "Qty Unloaded", "Variance", "Difference", "Reconcile?")
    Me.TransferList.SetColumnWidths Array("1500", "850", "1250", "1250", "1250", "1000", "1000", "1000")
    setNotifierLabels
    Mouse.Hourglass False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set locIDMap = Nothing
    Set whseCodeMap = Nothing
    Set truckIDMap = Nothing
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub Location_Click()
    setNotifierLabels
    setMasCaption
    requeryTransferList
End Sub

Private Sub RemoteLocation_Click()
    setNotifierLabels
    setMasCaption
    requeryTransferList
End Sub

Private Sub TransferList_DblClick(i As Long, j As Long, txt As String)
    Dim item As String
    item = Me.TransferList.GetText(i, COL_ITEM)
    Dim newval As String
    Select Case j
        
        Case Is = COL_ITEM
            'do nothing
        
        Case Is = COL_PRIO
            'do nothing
        
        Case Is = COL_REQD
            newval = getNewQuantity("requested", txt)
            If newval <> "" Then
                adjustRequest item, CStr(Me.TransferList.GetText(i, COL_PRIO)), CStr(CLng(newval) - CLng(txt))
                Me.TransferList.Edit CStr(CLng(newval)), i, COL_REQD
                Me.TransferList.Edit CStr(CLng(Me.TransferList.GetText(i, COL_LOAD)) - CLng(newval)), i, COL_DIFF
                If Me.TransferList.GetText(i, COL_DIFF) = "0" Then
                    Me.TransferList.Edit "", i, COL_DIFF
                End If
            End If
        
        Case Is = COL_LOAD
            newval = getNewQuantity("loaded", txt)
            If newval <> "" Then
                adjustLoad item, CStr(Me.TransferList.GetText(i, COL_PRIO)), CStr(CLng(newval) - CLng(txt))
                Me.TransferList.Edit CStr(CLng(newval)), i, COL_LOAD
                Me.TransferList.Edit CStr(CLng(Me.TransferList.GetText(i, COL_UNLD)) - CLng(newval)), i, COL_VARI
                Me.TransferList.Edit CStr(CLng(newval) - CLng(Me.TransferList.GetText(i, COL_REQD))), i, COL_DIFF
                If Me.TransferList.GetText(i, COL_VARI) = "0" Then
                    Me.TransferList.Edit "", i, COL_VARI
                End If
                If Me.TransferList.GetText(i, COL_DIFF) = "0" Then
                    Me.TransferList.Edit "", i, COL_DIFF
                End If
            End If
        
        Case Is = COL_UNLD
            newval = getNewQuantity("unloaded", txt)
            If newval <> "" Then
                adjustUnload item, CStr(Me.TransferList.GetText(i, COL_PRIO)), CStr(CLng(newval) - CLng(txt))
                Me.TransferList.Edit CStr(CLng(newval)), i, COL_UNLD
                Me.TransferList.Edit CStr(CLng(newval) - CLng(Me.TransferList.GetText(i, COL_LOAD))), i, COL_VARI
                If Me.TransferList.GetText(i, COL_VARI) = "0" Then
                    Me.TransferList.Edit "", i, COL_VARI
                End If
            End If
        
        Case Is = COL_VARI
            'do nothing
        
        Case Is = COL_DIFF
            'do nothing
        
        Case Is = COL_RECN
            Dim thisln As Variant
            thisln = Me.TransferList.GetRow(i)
            If CLng(thisln(COL_REQD)) = CLng(thisln(COL_LOAD)) And CLng(thisln(COL_LOAD)) = CLng(thisln(COL_UNLD)) Then
                reconcileLine CStr(thisln(COL_ITEM)), CStr(thisln(COL_PRIO)), CStr(thisln(COL_LOAD))
                Me.TransferList.Remove i
            End If
        
    End Select
End Sub

Private Sub Truck_Click()
    setNotifierLabels
    requeryTransferList
End Sub

Private Sub requeryTransferList()
    Me.TransferList.Clear
    
    If Me.Location = "" Or Me.RemoteLocation = "" Or Me.Truck = "" Then
        Exit Sub
    End If
    
    If Not IsSupervisorForWhse(locIDMap.item(Me.Location.ListIndex)) Then
        MsgBox "You are not a supervisor for the " & Me.Location
        Exit Sub
    End If
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("EXEC spTransfersReconciliation " & locIDMap.item(Me.RemoteLocation.ListIndex) & ", " & locIDMap.item(Me.Location.ListIndex) & ", " & truckIDMap.item(Me.Truck.ListIndex))
    While Not rst.EOF
        Dim variance As String, diff As String
        variance = CStr(CLng(rst("QtyUnloaded")) - CLng(rst("QtyLoaded")))
        diff = CStr(CLng(rst("QtyLoaded")) - CLng(rst("QtyReq")))
        Me.TransferList.Add Array(CStr(rst("ItemNumber")), _
                                  CStr(rst("PriorityLabel")), _
                                  CStr(rst("QtyReq")), _
                                  CStr(rst("QtyLoaded")), _
                                  CStr(rst("QtyUnloaded")), _
                                  IIf(variance = "0", "", variance), _
                                  IIf(diff = "0", "", diff), _
                                  "")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryLocations()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Name, ID, MasWhseCode FROM Warehouses")
    Me.Location.Clear
    Me.RemoteLocation.Clear
    Dim i As Long
    i = 0
    While Not rst.EOF
        locIDMap.Add i, CStr(rst("ID"))
        whseCodeMap.Add i, CStr(rst("MasWhseCode"))
        Me.Location.AddItem rst("Name")
        Me.RemoteLocation.AddItem rst("Name")
        i = i + 1
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requeryTrucks()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Name, ID FROM WarehouseTrucks")
    Me.Truck.Clear
    Dim i As Long
    i = 0
    While Not rst.EOF
        truckIDMap.Add i, CStr(rst("ID"))
        Me.Truck.AddItem rst("Name")
        i = i + 1
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub adjustRequest(item As String, priority As String, qty As String)
    Dim Notes As String
    Notes = InputBox("Enter notes here (optional):")
    DB.Execute "INSERT INTO HandheldTransfers ( LineType, ItemNumber, Priority, Quantity, PersonID, FromWhseID, ToWhseID, Notes ) VALUES ( 0, '" & item & "', " & priorityToBit(priority) & ", " & qty & ", -" & GetCurrentUserID & ", " & locIDMap.item(Me.RemoteLocation.ListIndex) & ", " & locIDMap.item(Me.Location.ListIndex) & ", '" & EscapeSQuotes(Notes) & "' )"
End Sub

Private Sub adjustLoad(item As String, priority As String, qty As String)
    Dim Notes As String
    Notes = InputBox("Enter notes here (optional):")
    DB.Execute "INSERT INTO HandheldTransfers ( LineType, ItemNumber, Priority, Quantity, PersonID, FromWhseID, ToWhseID, TruckID, Notes ) VALUES ( 1, '" & item & "', " & priorityToBit(priority) & ", " & qty & ", -" & GetCurrentUserID & ", " & locIDMap.item(Me.RemoteLocation.ListIndex) & ", " & locIDMap.item(Me.Location.ListIndex) & ", " & truckIDMap.item(Me.Truck.ListIndex) & ", '" & EscapeSQuotes(Notes) & "' )"
End Sub

Private Sub adjustUnload(item As String, priority As String, qty As String)
    Dim Notes As String
    Notes = InputBox("Enter notes here (optional):")
    DB.Execute "INSERT INTO HandheldTransfers ( LineType, ItemNumber, Priority, Quantity, PersonID, FromWhseID, ToWhseID, TruckID, Notes ) VALUES ( 2, '" & item & "', " & priorityToBit(priority) & ", " & qty & ", -" & GetCurrentUserID & ", " & locIDMap.item(Me.RemoteLocation.ListIndex) & ", " & locIDMap.item(Me.Location.ListIndex) & ", " & truckIDMap.item(Me.Truck.ListIndex) & ", '" & EscapeSQuotes(Notes) & "' )"
End Sub

Private Sub reconcileLine(item As String, priority As String, qty As String, Optional skipMasCaption As Boolean = False)
    DB.Execute "INSERT INTO HandheldTransfers ( LineType, ItemNumber, Priority, Quantity, PersonID, FromWhseID, ToWhseID, TruckID ) VALUES ( 3, '" & item & "', " & priorityToBit(priority) & ", " & qty & ", " & GetCurrentUserID & ", " & locIDMap.item(Me.RemoteLocation.ListIndex) & ", " & locIDMap.item(Me.Location.ListIndex) & ", " & truckIDMap.item(Me.Truck.ListIndex) & " )"
    If whseCodeMap.item(Me.Location.ListIndex) = whseCodeMap.item(Me.RemoteLocation.ListIndex) Then
        'same warehouse, no transfer needed
    Else
        DB.Execute "INSERT INTO ExportInventoryTransfers ( FromWhse, ItemNumber, ToWhse, Quantity ) VALUES ( '" & whseCodeMap.item(Me.RemoteLocation.ListIndex) & "', '" & item & "', '" & whseCodeMap.item(Me.Location.ListIndex) & "', " & qty & " )"
        If Not skipMasCaption Then
            setMasCaption
        End If
    End If
End Sub

Private Function getNewQuantity(qtyType As String, dflt As String) As String
    Dim newval As String
    newval = InputBox("Change " & qtyType & " quantity to what?", , dflt)
    If newval = "" Then
        getNewQuantity = ""
    ElseIf Not IsNumeric(newval) Then
        MsgBox "Enter a number!"
        getNewQuantity = ""
    Else
        getNewQuantity = newval
    End If
End Function

Private Sub setNotifierLabels()
    Me.lblNotifyLocation.Visible = CBool(Me.Location.ListIndex = -1)
    Me.lblNotifyWhse.Visible = CBool(Me.RemoteLocation.ListIndex = -1)
    Me.lblNotifyTruck.Visible = CBool(Me.Truck.ListIndex = -1)
End Sub

Private Function priorityToBit(priority As String) As Long
    Select Case priority
        Case Is = "High"
            priorityToBit = 1
        Case Is = "Normal"
            priorityToBit = 0
        Case Else
            Err.Raise 100, "priorityToBit", "Can't find priority " & qq(priority) & " in lookup!"
    End Select
End Function

Private Sub setMasCaption()
    If Me.Location.ListIndex = -1 Or Me.RemoteLocation.ListIndex = -1 Then
        Me.MasTransfers.caption = "Select locations first"
        Me.MasTransfers.Enabled = False
    Else
        Me.MasTransfers.caption = MAS_CAPTION
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT COUNT(DISTINCT ItemNumber), SUM(Quantity) FROM ExportInventoryTransfers WHERE Exported=0 AND FromWhse='" & whseCodeMap.item(Me.RemoteLocation.ListIndex) & "' AND ToWhse='" & whseCodeMap.item(Me.RemoteLocation.ListIndex) & "'")
        Dim itemcount As Long
        Dim piececount As Long
        If rst.EOF Then
            Me.MasTransfers.Enabled = False
            itemcount = 0
            piececount = 0
        Else
            itemcount = Nz(rst(0), "0")
            piececount = Nz(rst(1), "0")
            Me.MasTransfers.Enabled = CBool(itemcount <> 0 And piececount <> 0)
        End If
        Me.MasTransfers.caption = Replace(Me.MasTransfers.caption, "%SKU_COUNT%", itemcount)
        Me.MasTransfers.caption = Replace(Me.MasTransfers.caption, "%PIECE_COUNT%", piececount)
        rst.Close
        Set rst = Nothing
    End If
End Sub
