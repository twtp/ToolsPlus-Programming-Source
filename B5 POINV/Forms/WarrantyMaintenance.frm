VERSION 5.00
Begin VB.Form WarrantyMaintenance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Warranty Information Maintenance"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   9750
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   8250
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   0
      Width           =   1485
   End
   Begin VB.TextBox WarrantyID 
      Height          =   315
      Left            =   4710
      TabIndex        =   10
      Text            =   "WID"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   6075
      Left            =   2700
      TabIndex        =   9
      Top             =   570
      Width           =   6975
      Begin VB.Frame Frame4 
         Height          =   5055
         Left            =   5190
         TabIndex        =   27
         Top             =   960
         Width           =   1695
         Begin VB.ListBox WarrantyLineList 
            Height          =   1035
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   390
            Width           =   1515
         End
         Begin VB.CommandButton btnLineAssociate 
            Caption         =   "ã"
            BeginProperty Font 
               Name            =   "Wingdings 3"
               Size            =   12
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   240
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1500
            Width           =   555
         End
         Begin VB.CommandButton btnLineDeassociate 
            Caption         =   "ä"
            BeginProperty Font 
               Name            =   "Wingdings 3"
               Size            =   12
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   900
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1500
            Width           =   555
         End
         Begin VB.ListBox WarrantyLineAvailList 
            Height          =   2985
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1515
         End
         Begin VB.Label generalLabel 
            Caption         =   "Lines:"
            Height          =   255
            Index           =   4
            Left            =   150
            TabIndex        =   32
            Top             =   180
            Width           =   885
         End
      End
      Begin VB.TextBox GuaranteeInfo 
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   4260
         Width           =   4995
      End
      Begin VB.TextBox GuaranteeLength 
         Height          =   285
         Left            =   1890
         TabIndex        =   21
         Top             =   3930
         Width           =   765
      End
      Begin VB.CheckBox chkLifetimeGuarantee 
         Caption         =   "Lifetime"
         Height          =   225
         Left            =   1890
         TabIndex        =   20
         Top             =   3690
         Width           =   855
      End
      Begin VB.TextBox WarrantyInfo 
         Height          =   1695
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1770
         Width           =   4995
      End
      Begin VB.TextBox WarrantyLength 
         Height          =   285
         Left            =   870
         TabIndex        =   17
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox chkLifetimeWarranty 
         Caption         =   "Lifetime"
         Height          =   255
         Left            =   870
         TabIndex        =   16
         Top             =   1170
         Width           =   975
      End
      Begin VB.TextBox WarrantyManufURL 
         Height          =   285
         Left            =   780
         TabIndex        =   14
         Top             =   630
         Width           =   5955
      End
      Begin VB.CommandButton btnWarrantyURL 
         Caption         =   "WWW"
         Height          =   255
         Left            =   90
         TabIndex        =   13
         Top             =   630
         Width           =   675
      End
      Begin VB.TextBox WarrantyName 
         Height          =   285
         Left            =   780
         TabIndex        =   12
         Top             =   300
         Width           =   5955
      End
      Begin VB.Label generalLabel 
         Caption         =   "Satisfaction Guarantee:"
         Height          =   255
         Index           =   26
         Left            =   150
         TabIndex        =   23
         Top             =   3690
         Width           =   1695
      End
      Begin VB.Label generalLabel 
         Caption         =   "days"
         Height          =   255
         Index           =   27
         Left            =   2700
         TabIndex        =   22
         Top             =   3960
         Width           =   645
      End
      Begin VB.Line Line2 
         X1              =   150
         X2              =   5040
         Y1              =   3570
         Y2              =   3570
      End
      Begin VB.Label generalLabel 
         Caption         =   "months"
         Height          =   255
         Index           =   3
         Left            =   1770
         TabIndex        =   18
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label generalLabel 
         Caption         =   "Warranty:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   765
      End
      Begin VB.Line Line1 
         X1              =   150
         X2              =   5040
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Label generalLabel 
         Caption         =   "Name:"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   330
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6075
      Left            =   90
      TabIndex        =   1
      Top             =   570
      Width           =   2535
      Begin VB.CommandButton btnDeleteWarranty 
         Caption         =   "Delete Selected Warranty"
         Height          =   375
         Left            =   90
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   5610
         Width           =   2325
      End
      Begin VB.CommandButton btnAddNewWarranty 
         Caption         =   "Add New Warranty"
         Height          =   375
         Left            =   90
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   5220
         Width           =   2325
      End
      Begin VB.CheckBox chkShowOldLines 
         Caption         =   "Include old product lines"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   4920
         Width           =   2145
      End
      Begin VB.TextBox QuickSearch 
         Height          =   285
         Left            =   690
         TabIndex        =   6
         Top             =   750
         Width           =   1755
      End
      Begin VB.ListBox WarrantyList 
         Height          =   3765
         Left            =   90
         TabIndex        =   4
         Top             =   1050
         Width           =   2355
      End
      Begin VB.ComboBox LineFilter 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   420
         Width           =   2385
      End
      Begin VB.Label generalLabel 
         Caption         =   "Search:"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   780
         Width           =   585
      End
      Begin VB.Label Label2 
         Caption         =   "Select Line:"
         Height          =   255
         Left            =   90
         TabIndex        =   2
         Top             =   210
         Width           =   1275
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Warranty Information Maintenance"
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
      Width           =   4245
   End
End
Attribute VB_Name = "WarrantyMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private changed As Boolean
Private whichCtl As String
Private allWarranties() As String
Private lastQSearch As String
Private fillingForm As Boolean
Private lockCheckboxes As Boolean

Public Sub SetLineFilter(PL As String)
    Dim isHidden As Boolean
    isHidden = CBool(DLookup("Hide", "ProductLine", "ProductLine='" & PL & "'"))
    If isHidden And Me.chkShowOldLines.value = vbUnchecked Then
        Me.chkShowOldLines.value = vbChecked
    End If
    Dim i As Long
    For i = 0 To Me.LineFilter.ListCount - 1
        If Left(Me.LineFilter.list(i), 3) = PL Then
            Me.LineFilter.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Private Sub btnAddNewWarranty_Click()
    Dim wname As String
    wname = InputBox("Enter new warranty name [internal only]:")
    If wname = "" Then
        Exit Sub
    End If
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT COUNT(*) FROM Warranties WHERE WarrantyName='" & EscapeSQuotes(wname) & "'")
    If rst(0) > 0 Then
        If vbNo = MsgBox("A warranty with this name already exists, are you sure that this is a different one?", vbYesNo) Then
            rst.Close
            Set rst = Nothing
            Exit Sub
        End If
    End If
    rst.Close
    Set rst = Nothing
    DB.Execute "INSERT INTO Warranties ( WarrantyName ) VALUES ( '" & EscapeSQuotes(wname) & "' )"
    Dim newid As String
    newid = DLookup("@@IDENTITY", "Warranties")
    Me.WarrantyList.AddItem wname
    Dim i As Long
    For i = 0 To Me.WarrantyList.ListCount - 1
        If Me.WarrantyList.list(i) = wname Then
            Me.WarrantyList.Selected(i) = True
            Exit For
        End If
    Next i
End Sub

Private Sub btnLineAssociate_Click()
    If HasPermissionsTo("WarrantyWrite") Then
        If Me.WarrantyLineAvailList.SelCount <> 0 Then
            Dim linestr As String
            linestr = Me.WarrantyLineAvailList
            Dim pid As String
            pid = Mid(linestr, InStr(linestr, vbTab) + 1)
            DB.Execute "INSERT INTO ProductLineWarranties ( ProductLineID, WarrantyID ) VALUES ( " & pid & ", " & Me.WarrantyID & " )"
            Me.WarrantyLineList.AddItem linestr
            Me.WarrantyLineAvailList.RemoveItem Me.WarrantyLineAvailList.ListIndex
        End If
    End If
End Sub

Private Sub btnLineDeassociate_Click()
    If HasPermissionsTo("WarrantyWrite") Then
        If Me.WarrantyLineList.SelCount <> 0 Then
            Dim linestr As String
            linestr = Me.WarrantyLineList
            Dim pid As String
            pid = Mid(linestr, InStr(linestr, vbTab) + 1)
            DB.Execute "DELETE FROM ProductLineWarranties WHERE ProductLineID=" & pid & " AND WarrantyID=" & Me.WarrantyID
            Me.WarrantyLineAvailList.AddItem linestr
            Me.WarrantyLineList.RemoveItem Me.WarrantyLineList.ListIndex
        End If
    End If
End Sub

Private Sub btnDeleteWarranty_Click()
    If HasPermissionsTo("WarrantyWrite") Then
        If vbYes = MsgBox("Are you sure you want to delete this warranty?" & vbCrLf & vbCrLf & "This will remove it from all lines it is associated with!", vbYesNo) Then
            DB.Execute "DELETE FROM ProductLineWarranties WHERE WarrantyID=" & Me.WarrantyID
            DB.Execute "DELETE FROM Warranties WHERE ID=" & Me.WarrantyID
            Dim namestr As String
            namestr = buildNameString(Me.WarrantyName, Me.WarrantyID)
            Dim i As Long
            For i = 0 To Me.WarrantyList.ListCount - 1
                If Me.WarrantyList.list(i) = namestr Then
                    Me.WarrantyList.RemoveItem i
                    Exit For
                End If
            Next i
            clearForm
        End If
    End If
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnWarrantyURL_Click()
    If Me.WarrantyManufURL <> "" Then
        OpenDefaultApp Me.WarrantyManufURL
    End If
End Sub

Private Sub chkShowOldLines_Click()
    requeryLineList
End Sub

Private Sub Form_Load()
    Dim tabs(0) As Long
    tabs(0) = 300
    SetListTabStops Me.WarrantyList.hwnd, tabs
    SetListTabStops Me.WarrantyLineList.hwnd, tabs
    SetListTabStops Me.WarrantyLineAvailList.hwnd, tabs
    requeryLineList
    requeryWarrantyList
    
    If Not HasPermissionsTo("WarrantyWrite") Then
        Me.btnAddNewWarranty.Visible = False
        Me.WarrantyName.Locked = True
        Me.WarrantyManufURL.Locked = True
        'Me.chkLifetimeWarranty.Locked = True
        Me.WarrantyLength.Locked = True
        Me.WarrantyInfo.Locked = True
        'Me.chkLifetimeGuarantee.Locked = True
        Me.GuaranteeLength.Locked = True
        Me.GuaranteeInfo.Locked = True
        Me.btnDeleteWarranty.Visible = False
        Me.btnLineAssociate.Visible = False
        Me.btnLineDeassociate.Visible = False
        Me.WarrantyLineAvailList.Visible = False
        lockCheckboxes = True
    End If
End Sub

Private Sub LineFilter_Click()
    requeryWarrantyList
End Sub

Private Sub QuickSearch_Change()
    If Me.QuickSearch = lastQSearch Then
        'do nothing
    Else
        Me.WarrantyList.Clear
        Dim i As Long
        For i = 0 To UBound(allWarranties)
            If CBool(InStr(1, CStr(allWarranties(i)), Me.QuickSearch, vbTextCompare)) Then
                Me.WarrantyList.AddItem allWarranties(i)
            End If
        Next i
        lastQSearch = Me.QuickSearch
    End If
End Sub

Private Sub WarrantyLineAvailList_DblClick()
    If HasPermissionsTo("WarrantyWrite") Then
        btnLineAssociate_Click
    End If
End Sub

Private Sub WarrantyLineList_DblClick()
    If HasPermissionsTo("WarrantyWrite") Then
        btnLineDeassociate_Click
    End If
End Sub

Private Sub WarrantyList_Click()
    If changed = True Then
        Select Case whichCtl
            Case Is = "WarrantyName"
                WarrantyName_LostFocus
            Case Is = "WarrantyManufURL"
                WarrantyManufURL_LostFocus
            Case Is = "WarrantyLength"
                WarrantyLength_LostFocus
            Case Is = "WarrantyInfo"
                WarrantyInfo_LostFocus
            Case Is = "GuaranteeLength"
                GuaranteeLength_LostFocus
            Case Is = "GuaranteeInfo"
                GuaranteeInfo_LostFocus
            Case Else
                MsgBox "Forgot to add " & qq(whichCtl) & " to the select list!"
        End Select
    End If
    fillForm Mid(Me.WarrantyList, InStr(Me.WarrantyList, vbTab) + 1)
End Sub

Private Sub requeryLineList()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ProductLine FROM ProductLine " & IIf(Me.chkShowOldLines, "", "WHERE ProductLine.Hide=0 ") & "ORDER BY ProductLine")
    Me.LineFilter.Clear
    Me.LineFilter.AddItem "<All>"
    While Not rst.EOF
        Me.LineFilter.AddItem rst("ProductLine")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    ExpandDropDownToFit Me.LineFilter
    Me.LineFilter = "<All>"
End Sub

Private Sub requeryWarrantyList()
    Dim pid As String
    If Me.LineFilter = "<All>" Then
        pid = -1
    Else
        pid = DLookup("ID", "ProductLine", "ProductLine='" & EscapeSQuotes(Me.LineFilter) & "'")
    End If
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT DISTINCT Warranties.WarrantyName, Warranties.ID FROM Warranties LEFT OUTER JOIN ProductLineWarranties ON Warranties.ID=ProductLineWarranties.WarrantyID" & IIf(pid = -1, "", " WHERE ProductLineWarranties.ProductLineID=" & pid) & " ORDER BY Warranties.WarrantyName")
    Me.WarrantyList.Clear
    While Not rst.EOF
        Me.WarrantyList.AddItem buildNameString(Nz(rst("WarrantyName")), rst("ID"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    clearForm
    If Me.WarrantyList.ListCount = 0 Then
        allWarranties = Split("", "zerolengtharray")
    Else
        ReDim allNames(Me.WarrantyList.ListCount - 1) As String
        allWarranties = ListBoxAsArray(Me.WarrantyList)
    End If
End Sub

Private Sub fillForm(idnum As String)
    fillingForm = True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT WarrantyName, WarrantyURL, WarrantyLengthMonths, WarrantyInfo, SatisfactionGuaranteeLengthDays, SatisfactionGuaranteeInfo FROM Warranties WHERE ID=" & idnum)
    If Not rst.EOF Then
        Me.WarrantyID = idnum
        Me.WarrantyName = Nz(rst("WarrantyName"))
        Me.WarrantyManufURL = Nz(rst("WarrantyURL"))
        If IsNull(rst("WarrantyLengthMonths")) Then
            Me.chkLifetimeWarranty = 0
            Me.WarrantyLength = ""
            Enable Me.WarrantyLength
        ElseIf rst("WarrantyLengthMonths") = -1 Then
            Me.chkLifetimeWarranty = 1
            Me.WarrantyLength = ""
            Disable Me.WarrantyLength
        Else
            Me.chkLifetimeWarranty = 0
            Me.WarrantyLength = rst("WarrantyLengthMonths")
            Enable Me.WarrantyLength
        End If
        Me.WarrantyInfo = Nz(rst("WarrantyInfo"))
        If IsNull(rst("SatisfactionGuaranteeLengthDays")) Then
            Me.chkLifetimeGuarantee = 0
            Me.GuaranteeLength = ""
            Enable Me.GuaranteeLength
        ElseIf rst("SatisfactionGuaranteeLengthDays") = -1 Then
            Me.chkLifetimeGuarantee = 1
            Me.GuaranteeLength = ""
            Disable Me.GuaranteeLength
        Else
            Me.chkLifetimeGuarantee = 0
            Me.GuaranteeLength = rst("SatisfactionGuaranteeLengthDays")
            Enable Me.GuaranteeLength
        End If
        Me.GuaranteeInfo = Nz(rst("SatisfactionGuaranteeInfo"))
        Enable Me.WarrantyName
        Enable Me.WarrantyManufURL
        Enable Me.chkLifetimeWarranty
        Enable Me.WarrantyInfo
        Enable Me.chkLifetimeGuarantee
        Enable Me.GuaranteeInfo
        'warranty/guarantee length are already enabled/disabled appropriately
        Me.btnDeleteWarranty.Enabled = True
        fillLineSubform idnum
        Me.btnLineAssociate.Enabled = True
        Me.btnLineDeassociate.Enabled = True
    Else
        MsgBox "Can't find this warranty, wtf?"
        clearForm
    End If
    rst.Close
    Set rst = Nothing
    fillingForm = False
End Sub

Private Sub fillLineSubform(idnum As String)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ProductLine.ProductLine, ProductLine.ID FROM ProductLine INNER JOIN ProductLineWarranties ON ProductLine.ID=ProductLineWarranties.ProductLineID WHERE ProductLineWarranties.WarrantyID=" & Me.WarrantyID & IIf(Me.chkShowOldLines, "", " AND ProductLine.Hide=0"))
    Me.WarrantyLineList.Clear
    Me.WarrantyLineAvailList.Clear
    While Not rst.EOF
        Me.WarrantyLineList.AddItem rst("ProductLine") & vbTab & rst("ID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = DB.retrieve("SELECT PL1.ProductLine, PL1.ID FROM ProductLine AS PL1 WHERE PL1.ID NOT IN (SELECT PL2.ID FROM ProductLine AS PL2 INNER JOIN ProductLineWarranties ON PL2.ID=ProductLineWarranties.ProductLineID WHERE ProductLineWarranties.WarrantyID=" & Me.WarrantyID & ")" & IIf(Me.chkShowOldLines, "", " AND PL1.Hide=0"))
    While Not rst.EOF
        Me.WarrantyLineAvailList.AddItem rst("ProductLine") & vbTab & rst("ID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub clearForm()
    fillingForm = True
    Me.WarrantyName = ""
    Me.WarrantyManufURL = ""
    Me.WarrantyLength = ""
    Me.WarrantyInfo = ""
    Me.chkLifetimeGuarantee = 0
    Me.GuaranteeLength = ""
    Me.GuaranteeInfo = ""
    Me.WarrantyLineList.Clear
    Me.WarrantyLineAvailList.Clear
    Disable Me.WarrantyName
    Disable Me.WarrantyManufURL
    Disable Me.chkLifetimeWarranty
    Disable Me.WarrantyLength
    Disable Me.WarrantyInfo
    Disable Me.chkLifetimeGuarantee
    Disable Me.GuaranteeLength
    Disable Me.GuaranteeInfo
    Me.btnDeleteWarranty.Enabled = False
    Me.btnLineAssociate.Enabled = False
    Me.btnLineDeassociate.Enabled = False
End Sub

Private Sub fixNameList()
    Me.WarrantyList.list(Me.WarrantyList.ListIndex) = buildNameString(Me.WarrantyName, Me.WarrantyID)
End Sub

Private Function buildNameString(wname As String, id As String) As String
    buildNameString = wname & vbTab & id
End Function



'--------------------------------------
' UPDATING
'--------------------------------------

Private Sub WarrantyName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            WarrantyName_LostFocus
        Case Is = vbKeyDelete
            WarrantyName_KeyPress KeyCode
    End Select
End Sub

Private Sub WarrantyName_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "WarrantyName"
End Sub

Private Sub WarrantyName_LostFocus()
    If changed = True Then
        If Len(Me.WarrantyName) > 50 Then
            MsgBox "Name too long! Must be under 50 chars."
            Me.WarrantyName.SetFocus
        ElseIf Me.WarrantyName = "" Then
            MsgBox "Name cannot be blank!"
            Me.WarrantyName.SetFocus
        Else
            DB.Execute "UPDATE Warranties SET WarrantyName='" & EscapeSQuotes(Me.WarrantyName) & "' WHERE ID=" & Me.WarrantyID
            changed = False
            fixNameList
        End If
    End If
End Sub

Private Sub WarrantyManufURL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            WarrantyManufURL_LostFocus
        Case Is = vbKeyDelete
            WarrantyManufURL_KeyPress KeyCode
    End Select
End Sub

Private Sub WarrantyManufURL_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "WarrantyManufURL"
End Sub

Private Sub WarrantyManufURL_LostFocus()
    If changed = True Then
        If Len(Me.WarrantyManufURL) > 512 Then
            MsgBox "URL too long! Must be under 512 chars."
            Me.WarrantyManufURL.SetFocus
        Else
            DB.Execute "UPDATE Warranties SET WarrantyURL='" & EscapeSQuotes(Me.WarrantyManufURL) & "' WHERE ID=" & Me.WarrantyID
            changed = False
        End If
    End If
End Sub

Private Sub chkLifetimeWarranty_Click()
    If Not fillingForm Then
        If lockCheckboxes Then
            fillingForm = True
            Me.chkLifetimeWarranty = IIf(Me.chkLifetimeWarranty, 0, 1)
            fillingForm = False
        Else
            If Me.chkLifetimeWarranty Then
                DB.Execute "UPDATE Warranties SET WarrantyLengthMonths=-1 WHERE ID=" & Me.WarrantyID
                Me.WarrantyLength = ""
                Disable Me.WarrantyLength
            Else
                DB.Execute "UPDATE Warranties SET WarrantyLengthMonths=0 WHERE ID=" & Me.WarrantyID
                Enable Me.WarrantyLength
            End If
        End If
    End If
End Sub

Private Sub WarrantyLength_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            WarrantyLength_LostFocus
        Case Is = vbKeyDelete
            WarrantyLength_KeyPress KeyCode
    End Select
End Sub

Private Sub WarrantyLength_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
    If KeyAscii <> 0 Then
        changed = True
        whichCtl = "WarrantyLength"
    End If
End Sub

Private Sub WarrantyLength_LostFocus()
    If changed = True Then
        If IsNumeric(Me.WarrantyLength) Then
            DB.Execute "UPDATE Warranties SET WarrantyLengthMonths='" & EscapeSQuotes(Me.WarrantyLength) & "' WHERE ID=" & Me.WarrantyID
            changed = False
        ElseIf Me.WarrantyLength = "" Then
            DB.Execute "UPDATE Warranties SET WarrantyLengthMonths=NULL WHERE ID=" & Me.WarrantyID
            changed = False
        Else
            MsgBox "Length should be numeric! How did you do that?"
            Me.WarrantyLength.SetFocus
        End If
    End If
End Sub

Private Sub WarrantyInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            WarrantyInfo_LostFocus
        Case Is = vbKeyDelete
            WarrantyInfo_KeyPress KeyCode
    End Select
End Sub

Private Sub WarrantyInfo_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "WarrantyInfo"
End Sub

Private Sub WarrantyInfo_LostFocus()
    If changed = True Then
        DB.Execute "UPDATE Warranties SET WarrantyInfo='" & EscapeSQuotes(Me.WarrantyInfo) & "' WHERE ID=" & Me.WarrantyID
        changed = False
    End If
End Sub

Private Sub chkLifetimeGuarantee_Click()
    If Not fillingForm Then
        If lockCheckboxes Then
            fillingForm = True
            Me.chkLifetimeGuarantee = IIf(Me.chkLifetimeGuarantee, 0, 1)
            fillingForm = False
        Else
            If Me.chkLifetimeGuarantee Then
                DB.Execute "UPDATE Warranties SET SatisfactionGuaranteeLengthDays=-1 WHERE ID=" & Me.WarrantyID
                Me.GuaranteeLength = ""
                Disable Me.GuaranteeLength
            Else
                DB.Execute "UPDATE Warranties SET SatisfactionGuaranteeLengthDays=0 WHERE ID=" & Me.WarrantyID
                Enable Me.GuaranteeLength
            End If
        End If
    End If
End Sub

Private Sub GuaranteeLength_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            GuaranteeLength_LostFocus
        Case Is = vbKeyDelete
            GuaranteeLength_KeyPress KeyCode
    End Select
End Sub

Private Sub GuaranteeLength_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
    If KeyAscii <> 0 Then
        changed = True
        whichCtl = "GuaranteeLength"
    End If
End Sub

Private Sub GuaranteeLength_LostFocus()
    If changed = True Then
        If IsNumeric(Me.GuaranteeLength) Then
            DB.Execute "UPDATE Warranties SET SatisfactionGuaranteeLengthDays='" & EscapeSQuotes(Me.GuaranteeLength) & "' WHERE ID=" & Me.WarrantyID
            changed = False
        ElseIf Me.GuaranteeLength = "" Then
            DB.Execute "UPDATE Warranties SET SatisfactionGuaranteeLengthDays=NULL WHERE ID=" & Me.WarrantyID
            changed = False
        Else
            MsgBox "Length should be numeric! How did you do that?"
            Me.GuaranteeLength.SetFocus
        End If
    End If
End Sub

Private Sub GuaranteeInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            GuaranteeInfo_LostFocus
        Case Is = vbKeyDelete
            GuaranteeInfo_KeyPress KeyCode
    End Select
End Sub

Private Sub GuaranteeInfo_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "GuaranteeInfo"
End Sub

Private Sub GuaranteeInfo_LostFocus()
    If changed = True Then
        DB.Execute "UPDATE Warranties SET SatisfactionGuaranteeInfo='" & EscapeSQuotes(Me.GuaranteeInfo) & "' WHERE ID=" & Me.WarrantyID
        changed = False
    End If
End Sub
