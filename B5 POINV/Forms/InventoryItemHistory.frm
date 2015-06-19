VERSION 5.00
Begin VB.Form InventoryItemHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item History"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   8460
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkShowMAPP 
      Caption         =   "MAPP"
      Height          =   225
      Left            =   7200
      TabIndex        =   12
      Top             =   2280
      Width           =   1305
   End
   Begin VB.CheckBox chkShowDiscontinued 
      Caption         =   "D/C"
      Height          =   225
      Left            =   4650
      TabIndex        =   11
      Top             =   180
      Width           =   945
   End
   Begin VB.CheckBox chkShowCostNew 
      Caption         =   "TNC"
      Height          =   225
      Left            =   7200
      TabIndex        =   10
      Top             =   1320
      Width           =   1305
   End
   Begin VB.CheckBox chkShowCostStd 
      Caption         =   "Std Cost"
      Height          =   225
      Left            =   7200
      TabIndex        =   9
      Top             =   1560
      Width           =   1305
   End
   Begin VB.CheckBox chkShowEBay 
      Caption         =   "EBay Price"
      Height          =   225
      Left            =   7200
      TabIndex        =   8
      Top             =   2520
      Width           =   1305
   End
   Begin VB.CheckBox chkShowYahoo 
      Caption         =   "Yahoo Price"
      Height          =   225
      Left            =   7200
      TabIndex        =   7
      Top             =   2040
      Width           =   1305
   End
   Begin VB.CheckBox chkShowStore 
      Caption         =   "Store Price"
      Height          =   225
      Left            =   7200
      TabIndex        =   6
      Top             =   1800
      Width           =   1305
   End
   Begin VB.ComboBox ItemNumber 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2820
      TabIndex        =   4
      Top             =   3540
      Width           =   1305
   End
   Begin VB.CommandButton btnRevert 
      Caption         =   "Revert This Change"
      Height          =   495
      Left            =   1290
      TabIndex        =   3
      Top             =   3540
      Width           =   1305
   End
   Begin VB.ListBox changeList 
      Height          =   2205
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   6975
   End
   Begin VB.Label Label3 
      Caption         =   "ItemNumber:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   630
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Item Price/Cost/DC Change History"
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
      TabIndex        =   0
      Top             =   120
      Width           =   4365
   End
End
Attribute VB_Name = "InventoryItemHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload InventoryItemHistory
End Sub

Private Sub btnRevert_Click()
    If Me.changeList = "" Then
        Exit Sub
    End If
    
    MsgBox "This has been removed, because the implementation was awful."
    
    'MsgBox "chances are this button doesn't work, so you'll probably want to hit ""no"" here"
    'Dim str As String, one As String, two As String
    'str = Mid(Me.changeList, InStrRev(Me.changeList, vbTab) + 1)
    'Select Case Left(str, 15)
    '    Case Is = "set being disco" 'ntinued
    '        If MsgBox("Un-discontinue the " & Me.ItemNumber & "?", vbYesNo) = vbYes Then
    '            SetUnDiscontinued Me.ItemNumber
    '        End If
    '    Case Is = "discontinued in" 'the store
    '        If MsgBox("Un-discontinue the " & Me.ItemNumber & "?", vbYesNo) = vbYes Then
    '            SetUnDiscontinued Me.ItemNumber
    '        End If
    '    Case Is = "discontinued on" 'the web
    '        If MsgBox("Un-discontinue the " & Me.ItemNumber & "?", vbYesNo) = vbYes Then
    '            SetUnDiscontinued Me.ItemNumber
    '        End If
    '    Case Is = "unset discontin" 'ued
    '        If MsgBox("Discontinue the " & Me.ItemNumber & "?", vbYesNo) = vbYes Then
    '            HandleDiscontinuation Me.ItemNumber
    '        End If
    '    Case Is = "store price cha" 'nged from ? to ?
    '        one = Format(getNumber(str), "Currency")
    '        If MsgBox("Change the store price for " & Me.ItemNumber & " back to " & one & "?", vbYesNo) = vbYes Then
    '            LogItemStorePriceChanged Me.ItemNumber, one
    '            updateInventoryMaster "DiscountMarkupPriceRate1", one, Me.ItemNumber, ""
    '        End If
    '    Case Is = "web price chang" 'ed from ? to ?
    '        one = Format(getNumber(str), "Currency")
    '        If MsgBox("Change the web price for " & Me.ItemNumber & " back to " & one & "?", vbYesNo) = vbYes Then
    '            LogItemWebPriceChanged Me.ItemNumber, one
    '            updateInventoryMaster "IDiscountMarkupPriceRate1", one, Me.ItemNumber, ""
    '        End If
    '    Case Is = "ebay price chan" 'ged from ? to ?
    '        one = Format(getNumber(str), "Currency")
    '        If MsgBox("Change the ebay price for " & Me.ItemNumber & " back to " & one & "?", vbYesNo) = vbYes Then
    '            LogItemEBayPriceChanged Me.ItemNumber, one
    '            updateInventoryMaster "EBayPrice", one, Me.ItemNumber, ""
    '        End If
    '    Case Is = "cost changed fr" 'om ? to ?
    '        one = Format(getNumber(str, True), "Currency")
    '        two = Format(getNumber(str), "Currency")
    '        If MsgBox("Change the cost for " & Me.ItemNumber & " back to " & one & "?", vbYesNo) = vbYes Then
    '            LogItemCostChanged Me.ItemNumber, one
    '            updateInventoryMaster "StdCost", one, Me.ItemNumber, ""
    '            updateInventoryMaster "TNC", two, Me.ItemNumber, ""
    '        End If
    '    Case Is = "TNC changed fro" 'm ? to ?
    '        one = Format(getNumber(str), "Currency")
    '        If MsgBox("Change the TNC for " & Me.ItemNumber & " back to " & one & "?", vbYesNo) = vbYes Then
    '            LogItemTNCChanged Me.ItemNumber, one
    '            updateInventoryMaster "TNC", one, Me.ItemNumber, ""
    '        End If
    '    Case Is = "costs swapped, " 'was cost=?, TNC=?
    '        one = getNumber(str)
    '        two = getNumber(str, True)
    '        If MsgBox("Change the cost back to " & Format(two, "Currency") & " and TNC back to " & Format(one, "Currency") & " for " & Me.ItemNumber & " back to " & Format(one, "Currency") & "?", vbYesNo) = vbYes Then
    '            LogItemCostSwapped Me.ItemNumber
    '            DB.Execute "UPDATE InventoryMaster SET StdCost=" & two & ", TNC=" & one & " WHERE ItemNumber='" & Me.ItemNumber & "'"
    '        End If
    '    Case Else
    '        MsgBox "not sure what type of line that was...can't revert."
    '        Exit Sub
    'End Select
End Sub

Private Function getNumber(str As String, Optional First As Boolean = False) As String
    Dim re As RegExp
    Set re = New RegExp
    If First Then
        re.Pattern = "(\d+\.\d\d) to \d+\.\d\d"
    Else
        re.Pattern = "\d+\.\d\d to (\d+\.\d\d)"
    End If
    Dim matches As MatchCollection
    Set matches = re.Execute(str)
    getNumber = CStr(matches(0).SubMatches(0))
End Function

Private Sub changeList_Click()
    'Me.btnRevert.Enabled = Not Me.changeList.Selected(0)
End Sub

Private Sub chkShowCostNew_Click()
    requeryChangeList
End Sub

Private Sub chkShowCostStd_Click()
    requeryChangeList
End Sub

Private Sub chkShowDiscontinued_Click()
    requeryChangeList
End Sub

Private Sub chkShowEBay_Click()
    requeryChangeList
End Sub

Private Sub chkShowMAPP_Click()
    requeryChangeList
End Sub

Private Sub chkShowStore_Click()
    requeryChangeList
End Sub

Private Sub chkShowYahoo_Click()
    requeryChangeList
End Sub

Private Sub Form_Load()
    If Not HasPermissionsTo("InventoryMaintenanceWrite") Then
        Me.btnRevert.Enabled = False
    End If
'    requeryItemNumber
    Me.ItemNumber = InventoryMaintenance.ItemNumber
    Disable Me.ItemNumber
    SetListTabs Me.changeList, Array(30, 72)
    requeryChangeList
End Sub

'Private Sub ItemNumber_Click()
'    ItemNumber_LostFocus
'End Sub
'
'Private Sub ItemNumber_KeyDown(KeyCode As Integer, Shift As Integer)
'    AutoCompleteKeyDown Me.ItemNumber, KeyCode, Shift
'    If KeyCode = vbKeyReturn Then
'        Me.changeList.SetFocus
'    End If
'End Sub
'
'Private Sub ItemNumber_KeyPress(KeyAscii As Integer)
'    AutoCompleteKeyPress Me.ItemNumber, KeyAscii
'End Sub
'
'Private Sub ItemNumber_LostFocus()
'    AutoCompleteLostFocus Me.ItemNumber
'    requeryChangeList
'End Sub
'
'Private Sub requeryItemNumber()
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT ItemNumber FROM InventoryMaster WHERE ItemNumber<'XX%' ORDER BY ItemNumber")
'    Me.ItemNumber.Clear
'    While Not rst.EOF
'        Me.ItemNumber.AddItem rst("ItemNumber")
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'End Sub

Private Sub requeryChangeList()
    Dim q As PerlArray
    Set q = New PerlArray
    Dim q2 As PerlArray
    
    If Me.chkShowDiscontinued Then
        q.Push "SELECT TimeChanged, " & _
               "       ChangedBy, " & _
               "       CASE WHEN Status='being' THEN 'set being discontinued' " & _
               "            WHEN Status='store' THEN 'discontinued in the store' " & _
               "            WHEN Status='web'   THEN 'discontinued on the web' " & _
               "            WHEN Status='un'    THEN 'unset discontinued' " & _
               "            ELSE                     'invalid discontinued value' " & _
               "       END As StatusString " & _
               "FROM InventoryDiscontinuedLog " & _
               "WHERE ItemNumber='" & Me.ItemNumber & "'"
    End If
    If Me.chkShowMAPP Then
        q.Push "SELECT TimeChanged, " & _
               "       ChangedBy, " & _
               "       CASE WHEN IgnoreFlag=0 THEN 'mapp changed from ' + CAST(OldMAPP AS varchar) + ' to ' + CAST(NewMAPP AS varchar) " & _
               "            ELSE 'mapp ignore ' + CASE WHEN NewMAPP=1 THEN 'set' ELSE 'unset' END " & _
               "       END AS StatusString " & _
               "FROM InventoryMAPPLog " & _
               "WHERE ItemNumber='" & Me.ItemNumber & "'"
    End If
    If Me.chkShowCostNew Or Me.chkShowCostStd Then
        Set q2 = New PerlArray
        If Me.chkShowCostNew Then
            q2.Push "IsRegularCost=0"
        End If
        If Me.chkShowCostStd Then
            q2.Push "IsRegularCost=1"
        End If
        q.Push "SELECT TimeChanged, " & _
               "       ChangedBy, " & _
               "       CASE WHEN IsSwap=1 THEN 'costs swapped, was cost=' + CAST(oldcost AS varchar) + ', TNC=' + CAST(newcost AS varchar) " & _
               "            ELSE          CASE WHEN IsRegularCost=1 THEN 'cost' ELSE 'TNC' END + ' changed from ' + CAST(oldcost AS varchar) + ' to ' + CAST(newcost AS varchar) " & _
               "       END As StatusString " & _
               "FROM InventoryCostLog " & _
               "WHERE ItemNumber='" & Me.ItemNumber & "' " & _
               "  AND (IsSwap=1 OR " & q2.Join(" OR ") & ")"
    End If
    If Me.chkShowStore Or Me.chkShowYahoo Or Me.chkShowEBay Then
        Set q2 = New PerlArray
        If Me.chkShowStore Then
            q2.Push "IsStorePrice=1"
        End If
        If Me.chkShowYahoo Then
            q2.Push "IsStorePrice=0"
        End If
        If Me.chkShowEBay Then
            q2.Push "IsStorePrice=2"
        End If
        q.Push "SELECT TimeChanged, " & _
               "       ChangedBy, " & _
               "       CASE WHEN IsStorePrice=1 THEN 'store' WHEN IsStorePrice=2 THEN 'ebay' ELSE 'web' END + ' price changed from ' + CAST(oldprice AS varchar) + ' to ' + CAST(newprice AS varchar) AS StatusString " & _
               "FROM InventoryPriceLog " & _
               "WHERE ItemNumber='" & Me.ItemNumber & "' " & _
               "  AND (" & q2.Join(" OR ") & ")"
    End If
    
    Me.changeList.Clear
    
    If q.Scalar > 0 Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT TimeChanged, ChangedBy, StatusString FROM (" & q.Join(" UNION ") & ") AS Tmp ORDER BY TimeChanged DESC")
        While Not rst.EOF
            Me.changeList.AddItem rst("TimeChanged") & vbTab & rst("ChangedBy") & vbTab & rst("StatusString")
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
    End If
    
    'Dim rst As ADODB.Recordset
    'Set rst = DB.retrieve("SELECT TimeChanged, ChangedBy, StatusString FROM vInventoryHistory WHERE ItemNumber='" & Me.ItemNumber & "' ORDER BY TimeChanged DESC")
    'Me.changeList.Clear
    'While Not rst.EOF
    '    Me.changeList.AddItem rst("TimeChanged") & vbTab & rst("ChangedBy") & vbTab & rst("StatusString")
    '    rst.MoveNext
    'Wend
    'rst.Close
    'Set rst = Nothing
End Sub
