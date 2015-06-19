VERSION 5.00
Begin VB.Form SalesRepMaintenance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Rep Maintenance"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   8880
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   6105
      Left            =   2700
      TabIndex        =   11
      Top             =   570
      Width           =   6105
      Begin VB.CommandButton btnDeleteSalesRep 
         Caption         =   "Delete Rep"
         Height          =   315
         Left            =   4290
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   900
         Width           =   1365
      End
      Begin VB.Frame Frame4 
         Height          =   4635
         Left            =   3720
         TabIndex        =   37
         Top             =   1470
         Width           =   2385
         Begin VB.CommandButton btnMakePrimaryContact 
            Caption         =   "Primary Contact"
            Height          =   255
            Left            =   930
            TabIndex        =   46
            Top             =   180
            Width           =   1365
         End
         Begin VB.ListBox RepLineAvailList 
            Height          =   2595
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   1920
            Width           =   2205
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
            Left            =   1290
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1530
            Width           =   555
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
            Left            =   630
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1530
            Width           =   555
         End
         Begin VB.ListBox RepLineList 
            Height          =   1035
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   450
            Width           =   2205
         End
         Begin VB.Label generalLabel 
            Caption         =   "Lines:"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   38
            Top             =   210
            Width           =   885
         End
      End
      Begin VB.TextBox SalesRepAddress 
         Height          =   285
         Left            =   930
         TabIndex        =   18
         Top             =   1410
         Width           =   2595
      End
      Begin VB.TextBox SalesRepFirstName 
         Height          =   285
         Left            =   930
         TabIndex        =   13
         Top             =   810
         Width           =   1275
      End
      Begin VB.TextBox SalesRepCity 
         Height          =   285
         Left            =   930
         TabIndex        =   20
         Top             =   1710
         Width           =   2595
      End
      Begin VB.TextBox SalesRepState 
         Height          =   285
         Left            =   930
         TabIndex        =   22
         Top             =   2010
         Width           =   2595
      End
      Begin VB.TextBox SalesRepZipCode 
         Height          =   285
         Left            =   930
         TabIndex        =   24
         Top             =   2310
         Width           =   2595
      End
      Begin VB.TextBox SalesRepPhone 
         Height          =   285
         Left            =   930
         TabIndex        =   26
         Top             =   2670
         Width           =   2595
      End
      Begin VB.TextBox SalesRepCell 
         Height          =   285
         Left            =   930
         TabIndex        =   28
         Top             =   2970
         Width           =   2595
      End
      Begin VB.TextBox SalesRepFax 
         Height          =   285
         Left            =   930
         TabIndex        =   31
         Top             =   3270
         Width           =   2595
      End
      Begin VB.CommandButton btnEmailRep 
         Caption         =   "Email"
         Height          =   255
         Left            =   180
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3570
         Width           =   705
      End
      Begin VB.TextBox SalesRepEmail 
         Height          =   285
         Left            =   930
         TabIndex        =   34
         Top             =   3570
         Width           =   2595
      End
      Begin VB.TextBox SalesRepNotes 
         Height          =   1665
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   4110
         Width           =   3345
      End
      Begin VB.TextBox SalesRepLastName 
         Height          =   285
         Left            =   2250
         TabIndex        =   14
         Top             =   810
         Width           =   1275
      End
      Begin VB.TextBox SalesRepPosition 
         Height          =   285
         Left            =   930
         TabIndex        =   17
         Top             =   1110
         Width           =   2595
      End
      Begin VB.TextBox TitleArea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   5835
      End
      Begin VB.Label generalLabel 
         Caption         =   "Name:"
         Height          =   255
         Index           =   25
         Left            =   180
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.Label generalLabel 
         Caption         =   "Address:"
         Height          =   255
         Index           =   29
         Left            =   180
         TabIndex        =   19
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label generalLabel 
         Caption         =   "City:"
         Height          =   255
         Index           =   30
         Left            =   180
         TabIndex        =   21
         Top             =   1740
         Width           =   735
      End
      Begin VB.Label generalLabel 
         Caption         =   "State:"
         Height          =   255
         Index           =   31
         Left            =   180
         TabIndex        =   23
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label generalLabel 
         Caption         =   "Zip Code:"
         Height          =   255
         Index           =   32
         Left            =   180
         TabIndex        =   25
         Top             =   2340
         Width           =   735
      End
      Begin VB.Label generalLabel 
         Caption         =   "Phone:"
         Height          =   255
         Index           =   33
         Left            =   180
         TabIndex        =   27
         Top             =   2700
         Width           =   735
      End
      Begin VB.Label generalLabel 
         Caption         =   "Cell:"
         Height          =   255
         Index           =   34
         Left            =   180
         TabIndex        =   29
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label generalLabel 
         Caption         =   "Fax:"
         Height          =   255
         Index           =   35
         Left            =   180
         TabIndex        =   32
         Top             =   3300
         Width           =   735
      End
      Begin VB.Label generalLabel 
         Caption         =   "Notes:"
         Height          =   255
         Index           =   36
         Left            =   210
         TabIndex        =   35
         Top             =   3900
         Width           =   735
      End
      Begin VB.Label generalLabel 
         Caption         =   "Position:"
         Height          =   255
         Index           =   39
         Left            =   180
         TabIndex        =   30
         Top             =   1140
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   60
      TabIndex        =   3
      Top             =   570
      Width           =   2535
      Begin VB.TextBox QuickSearch 
         Height          =   285
         Left            =   690
         TabIndex        =   45
         Top             =   750
         Width           =   1755
      End
      Begin VB.CheckBox chkShowOldLines 
         Caption         =   "Include old product lines"
         Height          =   255
         Left            =   180
         TabIndex        =   43
         Top             =   5130
         Width           =   2145
      End
      Begin VB.CommandButton btnAddNewRep 
         Caption         =   "Add New Sales Rep"
         Height          =   375
         Left            =   90
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   5430
         Width           =   2325
      End
      Begin VB.ComboBox LineFilter 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   420
         Width           =   2385
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sort Names By:"
         Height          =   855
         Left            =   90
         TabIndex        =   7
         Top             =   4260
         Width           =   2355
         Begin VB.OptionButton opNameStyle 
            Caption         =   "First name then last"
            Height          =   225
            Index           =   1
            Left            =   150
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   540
            Value           =   -1  'True
            Width           =   2025
         End
         Begin VB.OptionButton opNameStyle 
            Caption         =   "Last name then first"
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   270
            Width           =   1995
         End
      End
      Begin VB.ListBox SalesRepList 
         Height          =   3180
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1050
         Width           =   2355
      End
      Begin VB.Label generalLabel 
         Caption         =   "Search:"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   44
         Top             =   780
         Width           =   585
      End
      Begin VB.Label Label2 
         Caption         =   "Select Line:"
         Height          =   255
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   1275
      End
   End
   Begin VB.TextBox SalesRepID 
      Height          =   315
      Left            =   4290
      TabIndex        =   2
      Text            =   "SID"
      Top             =   150
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   7380
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "Sales Rep Maintenance"
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
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   3135
   End
End
Attribute VB_Name = "SalesRepMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private changed As Boolean
Private whichCtl As String
Private allNames() As String
Private lastQSearch As String
Private fillingForm As Boolean


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

Private Sub btnAddNewRep_Click()
    Dim fname As String, lname As String
    fname = InputBox("Enter new rep's first name:")
    If fname = "" Then
        Exit Sub
    End If
    lname = InputBox("Enter new rep's last name:")
    'If lname = "" Then
    '    Exit Sub
    'End If
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT COUNT(*) FROM SalesReps WHERE FirstName='" & EscapeSQuotes(fname) & "' AND LastName='" & EscapeSQuotes(lname) & "'")
    If rst(0) > 0 Then
        If vbNo = MsgBox("A sales rep with this name already exists, are you sure that this is a different guy?", vbYesNo) Then
            rst.Close
            Set rst = Nothing
            Exit Sub
        End If
    End If
    rst.Close
    Set rst = Nothing
    fname = UCase(Left(fname, 1)) & Mid(fname, 2)
    lname = UCase(Left(lname, 1)) & Mid(lname, 2)
    DB.Execute "INSERT INTO SalesReps ( FirstName, LastName ) VALUES ( '" & EscapeSQuotes(fname) & "', '" & EscapeSQuotes(lname) & "' )"
    Dim newid As String
    newid = DLookup("@@IDENTITY", "SalesReps")
    If Me.LineFilter <> "<All>" Then
        If vbYes = MsgBox("Add this rep to the " & Me.LineFilter & " line?", vbYesNo) Then
            Dim pid As String
            pid = DLookup("ID", "ProductLine", "ProductLine='" & Me.LineFilter & "'")
            DB.Execute "INSERT INTO ProductLineSalesReps ( ProductLineID, SalesRepID ) VALUES ( " & pid & ", " & newid & " )"
        End If
    End If
    Dim namestr As String
    namestr = buildNameString(fname, lname, newid)
    Me.SalesRepList.AddItem namestr
    Dim i As Long
    For i = 0 To Me.SalesRepList.ListCount - 1
        If Me.SalesRepList.list(i) = namestr Then
            Me.SalesRepList.Selected(i) = True
            Exit For
        End If
    Next i
End Sub

Private Sub btnLineAssociate_Click()
    If HasPermissionsTo("SalesRepsWrite") Then
        If Me.RepLineAvailList.SelCount <> 0 Then
            Dim linestr As String
            linestr = Me.RepLineAvailList
            Dim pid As String
            pid = Mid(linestr, InStr(linestr, vbTab) + 1)
            DB.Execute "INSERT INTO ProductLineSalesReps ( ProductLineID, SalesRepID ) VALUES ( " & pid & ", " & Me.SalesRepID & " )"
            Me.RepLineList.AddItem linestr
            Me.RepLineAvailList.RemoveItem Me.RepLineAvailList.ListIndex
        End If
    End If
End Sub

Private Sub btnLineDeassociate_Click()
    If HasPermissionsTo("SalesRepsWrite") Then
        If Me.RepLineList.SelCount <> 0 Then
            Dim linestr As String
            linestr = Me.RepLineList
            Dim pid As String
            pid = Mid(linestr, InStr(linestr, vbTab) + 1)
            DB.Execute "DELETE FROM ProductLineSalesReps WHERE ProductLineID=" & pid & " AND SalesRepID=" & Me.SalesRepID
            Me.RepLineAvailList.AddItem linestr
            Me.RepLineList.RemoveItem Me.RepLineList.ListIndex
        End If
        If Me.RepLineList.ListCount = 0 Then
            Me.btnMakePrimaryContact.Enabled = False
        End If
    End If
End Sub

Private Sub btnDeleteSalesRep_Click()
    If HasPermissionsTo("SalesRepsWrite") Then
        If vbYes = MsgBox("Are you sure you want to delete this sales rep?" & vbCrLf & vbCrLf & "This will remove him from all lines he is associated with!", vbYesNo) Then
            DB.Execute "DELETE FROM ProductLineSalesReps WHERE SalesRepID=" & Me.SalesRepID
            DB.Execute "DELETE FROM SalesReps WHERE ID=" & Me.SalesRepID
            Dim namestr As String
            namestr = buildNameString(Me.SalesRepFirstName, Me.SalesRepLastName, Me.SalesRepID)
            Dim i As Long
            For i = 0 To Me.SalesRepList.ListCount - 1
                If Me.SalesRepList.list(i) = namestr Then
                    Me.SalesRepList.RemoveItem i
                    Exit For
                End If
            Next i
            clearForm
        End If
    End If
End Sub

Private Sub btnEmailRep_Click()
    If Me.SalesRepEmail <> "" Then
        OpenEmailTo Me.SalesRepEmail
    End If
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnMakePrimaryContact_Click()
    If Me.RepLineList.ListIndex <> -1 Then
        Dim pid As String
        pid = Mid(Me.RepLineList, InStr(Me.RepLineList, vbTab) + 1)
        'DB.Execute "UPDATE ProductLine SET PrimarySalesRepID=" & Me.SalesRepID & " WHERE ID=" & pid
        DB.Execute "UPDATE ProductLineSalesReps SET IsPrimary=0 WHERE ProductLineID=" & pid
        DB.Execute "UPDATE ProductLineSalesReps SET IsPrimary=1 WHERE ProductLineID=" & pid & " AND SalesRepID=" & Me.SalesRepID
        Me.RepLineList.list(Me.RepLineList.ListIndex) = "*" & Me.RepLineList
    End If
End Sub

Private Sub chkShowOldLines_Click()
    requeryLineList
End Sub

Private Sub Form_Load()
    Dim tabs(0) As Long
    tabs(0) = 300
    SetListTabStops Me.SalesRepList.hWnd, tabs
    SetListTabStops Me.RepLineList.hWnd, tabs
    SetListTabStops Me.RepLineAvailList.hWnd, tabs
    requeryLineList
    'requerySalesRepList
    
    If Not HasPermissionsTo("SalesRepsWrite") Then
        Me.btnAddNewRep.Visible = False
        Me.SalesRepFirstName.Locked = True
        Me.SalesRepLastName.Locked = True
        Me.SalesRepPosition.Locked = True
        Me.SalesRepAddress.Locked = True
        Me.SalesRepCity.Locked = True
        Me.SalesRepState.Locked = True
        Me.SalesRepZipCode.Locked = True
        Me.SalesRepPhone.Locked = True
        Me.SalesRepCell.Locked = True
        Me.SalesRepFax.Locked = True
        Me.SalesRepEmail.Locked = True
        Me.SalesRepNotes.Locked = True
        Me.btnDeleteSalesRep.Visible = False
        Me.btnLineAssociate.Visible = False
        Me.btnLineDeassociate.Visible = False
        Me.RepLineAvailList.Visible = False
        Me.btnMakePrimaryContact.Visible = False
    End If
End Sub

Private Sub LineFilter_Click()
    requerySalesRepList
End Sub

Private Sub opNameStyle_Click(index As Integer)
    requerySalesRepList
End Sub

Private Sub QuickSearch_Change()
    If Me.QuickSearch = lastQSearch Then
        'do nothing
    Else
        'Me.SalesRepList.Clear
        'Dim i As Long
        'For i = 0 To UBound(allNames)
        '    If CBool(InStr(1, CStr(allNames(i)), Me.QuickSearch, vbTextCompare)) Then
        '        Me.SalesRepList.AddItem allNames(i)
        '    End If
        'Next i
        lastQSearch = Me.QuickSearch
        fillSalesRepList
    End If
End Sub

Private Sub RepLineAvailList_DblClick()
    If HasPermissionsTo("SalesRepsWrite") Then
        btnLineAssociate_Click
    End If
End Sub

Private Sub RepLineList_Click()
    If HasPermissionsTo("SalesRepsWrite") Then
        Me.btnMakePrimaryContact.Enabled = True
    End If
End Sub

Private Sub RepLineList_DblClick()
    If HasPermissionsTo("SalesRepsWrite") Then
        Me.btnMakePrimaryContact.Enabled = True
        btnLineDeassociate_Click
    End If
End Sub

Private Sub SalesRepList_Click()
    If Not fillingForm Then
        If changed = True Then
            Select Case whichCtl
                Case Is = "SalesRepFirstName"
                    SalesRepFirstName_LostFocus
                Case Is = "SalesRepLastName"
                    SalesRepLastName_LostFocus
                Case Is = "SalesRepPosition"
                    SalesRepPosition_LostFocus
                Case Is = "SalesRepAddress"
                    SalesRepAddress_LostFocus
                Case Is = "SalesRepCity"
                    SalesRepCity_LostFocus
                Case Is = "SalesRepState"
                    SalesRepState_LostFocus
                Case Is = "SalesRepZipCode"
                    SalesRepZipCode_LostFocus
                Case Is = "SalesRepPhone"
                    SalesRepPhone_LostFocus
                Case Is = "SalesRepCell"
                    SalesRepCell_LostFocus
                Case Is = "SalesRepFax"
                    SalesRepFax_LostFocus
                Case Is = "SalesRepEmail"
                    SalesRepEmail_LostFocus
                Case Is = "SalesRepNotes"
                    SalesRepNotes_LostFocus
                Case Else
                    MsgBox "Forgot to add " & qq(whichCtl) & " to select clause!"
            End Select
        End If
        fillForm Mid(Me.SalesRepList, InStr(Me.SalesRepList, vbTab) + 1)
    End If
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

Private Sub requerySalesRepList()
    Dim pid As String
    If Me.LineFilter = "<All>" Then
        pid = -1
    Else
        pid = DLookup("ID", "ProductLine", "ProductLine='" & EscapeSQuotes(Me.LineFilter) & "'")
    End If
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT DISTINCT SalesReps.FirstName, SalesReps.LastName, SalesReps.ID FROM SalesReps LEFT OUTER JOIN ProductLineSalesReps ON SalesReps.ID=ProductLineSalesReps.SalesRepID" & IIf(pid = -1, "", " WHERE ProductLineSalesReps.ProductLineID=" & pid))
    'Me.SalesRepList.Clear
    'While Not rst.EOF
    '    Me.SalesRepList.AddItem buildNameString(Nz(rst("FirstName")), Nz(rst("LastName")), rst("ID"))
    '    rst.MoveNext
    'Wend
    'rst.Close
    'Set rst = Nothing
    'clearForm
    'If Me.SalesRepList.ListCount = 0 Then
    '    allNames = Split("", "zerolengtharray")
    'Else
    '    ReDim allNames(Me.SalesRepList.ListCount - 1) As String
    '    allNames = ListBoxAsArray(Me.SalesRepList)
    'End If
    If rst.RecordCount = 0 Then
        allNames = Split("", "zerolengtharray")
    Else
        ReDim allNames(rst.RecordCount - 1) As String
        Dim i As Long
        For i = 0 To rst.RecordCount - 1
            allNames(i) = buildNameString(Nz(rst("FirstName")), Nz(rst("LastName")), rst("ID"))
            rst.MoveNext
        Next i
    End If
    fillSalesRepList
End Sub

Private Sub fillSalesRepList()
    Dim cursel As String
    If Me.SalesRepList.ListIndex = -1 Then
        cursel = ""
    Else
        cursel = Me.SalesRepList
    End If
    Dim foundsel As Boolean
    foundsel = False
    Me.SalesRepList.Clear
    Dim i As Long
    For i = 0 To UBound(allNames)
        If CBool(InStr(1, CStr(allNames(i)), lastQSearch, vbTextCompare)) Then
            Me.SalesRepList.AddItem allNames(i)
            If allNames(i) = cursel Then
                fillingForm = True
                Me.SalesRepList = allNames(i)
                fillingForm = False
                foundsel = True
            End If
        End If
    Next i
    If Not foundsel Then
        clearForm
    Else
        'TODO: make sure selection is in visible part of list
    End If
End Sub

Private Sub fillForm(idnum As String)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT FirstName, LastName, PositionTitle, Address, City, State, ZipCode, Phone, Cell, Fax, Email, Notes FROM SalesReps WHERE ID=" & idnum)
    If Not rst.EOF Then
        Me.SalesRepID = idnum
        Me.SalesRepFirstName = Nz(rst("FirstName"))
        Me.SalesRepLastName = Nz(rst("LastName"))
        Me.SalesRepPosition = Nz(rst("PositionTitle"))
        fixTitleArea
        Me.SalesRepAddress = Nz(rst("Address"))
        Me.SalesRepCity = Nz(rst("City"))
        Me.SalesRepState = Nz(rst("State"))
        Me.SalesRepZipCode = Nz(rst("ZipCode"))
        Me.SalesRepPhone = Nz(rst("Phone"))
        Me.SalesRepCell = Nz(rst("Cell"))
        Me.SalesRepFax = Nz(rst("Fax"))
        Me.SalesRepEmail = Nz(rst("Email"))
        Me.SalesRepNotes = Nz(rst("Notes"))
        Enable Me.SalesRepFirstName
        Enable Me.SalesRepLastName
        Enable Me.SalesRepPosition
        Enable Me.SalesRepAddress
        Enable Me.SalesRepCity
        Enable Me.SalesRepState
        Enable Me.SalesRepZipCode
        Enable Me.SalesRepPhone
        Enable Me.SalesRepCell
        Enable Me.SalesRepFax
        Enable Me.SalesRepEmail
        Enable Me.SalesRepNotes
        Me.btnDeleteSalesRep.Enabled = True
        fillLineSubform idnum
        Me.btnLineAssociate.Enabled = True
        Me.btnLineDeassociate.Enabled = True
    Else
        MsgBox "Can't find this sales rep, wtf?"
        clearForm
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub fillLineSubform(idnum As String)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ProductLine.ProductLine, ProductLine.ID, ProductLineSalesReps.IsPrimary FROM ProductLine INNER JOIN ProductLineSalesReps ON ProductLine.ID=ProductLineSalesReps.ProductLineID WHERE ProductLineSalesReps.SalesRepID=" & Me.SalesRepID & IIf(Me.chkShowOldLines, "", " AND ProductLine.Hide=0"))
    Me.btnMakePrimaryContact.Enabled = False
    Me.RepLineList.Clear
    Me.RepLineAvailList.Clear
    While Not rst.EOF
        Me.RepLineList.AddItem IIf(rst("IsPrimary"), "*", "") & rst("ProductLine") & vbTab & rst("ID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = DB.retrieve("SELECT PL1.ProductLine, PL1.ID FROM ProductLine AS PL1 WHERE PL1.ID NOT IN (SELECT PL2.ID FROM ProductLine AS PL2 INNER JOIN ProductLineSalesReps ON PL2.ID=ProductLineSalesReps.ProductLineID WHERE ProductLineSalesReps.SalesRepID=" & Me.SalesRepID & ")" & IIf(Me.chkShowOldLines, "", " AND PL1.Hide=0"))
    While Not rst.EOF
        Me.RepLineAvailList.AddItem rst("ProductLine") & vbTab & rst("ID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
End Sub

Private Sub clearForm()
    Me.TitleArea = ""
    Me.SalesRepFirstName = ""
    Me.SalesRepLastName = ""
    Me.SalesRepPosition = ""
    Me.SalesRepAddress = ""
    Me.SalesRepCity = ""
    Me.SalesRepState = ""
    Me.SalesRepZipCode = ""
    Me.SalesRepPhone = ""
    Me.SalesRepCell = ""
    Me.SalesRepFax = ""
    Me.SalesRepEmail = ""
    Me.SalesRepNotes = ""
    Me.RepLineList.Clear
    Me.RepLineAvailList.Clear
    Disable Me.SalesRepFirstName
    Disable Me.SalesRepLastName
    Disable Me.SalesRepPosition
    Disable Me.SalesRepAddress
    Disable Me.SalesRepCity
    Disable Me.SalesRepState
    Disable Me.SalesRepZipCode
    Disable Me.SalesRepPhone
    Disable Me.SalesRepCell
    Disable Me.SalesRepFax
    Disable Me.SalesRepEmail
    Disable Me.SalesRepNotes
    Me.btnDeleteSalesRep.Enabled = False
    Me.btnLineAssociate.Enabled = False
    Me.btnLineDeassociate.Enabled = False
    Me.btnMakePrimaryContact.Enabled = False
End Sub

Private Sub fixTitleArea()
    Me.TitleArea = Me.SalesRepFirstName & " " & Me.SalesRepLastName & IIf(Me.SalesRepPosition = "", "", " (" & Me.SalesRepPosition & ")")
End Sub

Private Sub fixNameList()
    Me.SalesRepList.list(Me.SalesRepList.ListIndex) = buildNameString(Me.SalesRepFirstName, Me.SalesRepLastName, Me.SalesRepID)
End Sub

Private Function buildNameString(fname As String, lname As String, id As String) As String
    If Me.opNameStyle(0) Then
        buildNameString = IIf(lname <> "", lname & ", ", "") & fname & vbTab & id
    Else
        buildNameString = fname & IIf(lname <> "", " " & lname, "") & vbTab & id
    End If
End Function


'--------------------------------------
' UPDATING
'--------------------------------------

Private Sub SalesRepFirstName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            SalesRepFirstName_LostFocus
        Case Is = vbKeyDelete
            SalesRepFirstName_KeyPress KeyCode
    End Select
End Sub

Private Sub SalesRepFirstName_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "SalesRepFirstName"
End Sub

Private Sub SalesRepFirstName_LostFocus()
    If changed = True Then
        If Len(Me.SalesRepFirstName) > 50 Then
            MsgBox "Name too long! Must be under 50 chars."
            Me.SalesRepFirstName.SetFocus
        ElseIf Me.SalesRepFirstName = "" Then
            MsgBox "First name cannot be blank!"
            Me.SalesRepFirstName.SetFocus
        Else
            DB.Execute "UPDATE SalesReps SET FirstName='" & EscapeSQuotes(Me.SalesRepFirstName) & "' WHERE ID=" & Me.SalesRepID
            changed = False
            fixTitleArea
            fixNameList
        End If
    End If
End Sub

Private Sub SalesRepLastName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            SalesRepLastName_LostFocus
        Case Is = vbKeyDelete
            SalesRepLastName_KeyPress KeyCode
    End Select
End Sub

Private Sub SalesRepLastName_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "SalesRepLastName"
End Sub

Private Sub SalesRepLastName_LostFocus()
    If changed = True Then
        If Len(Me.SalesRepLastName) > 50 Then
            MsgBox "Name too long! Must be under 50 chars."
            Me.SalesRepLastName.SetFocus
        Else
            DB.Execute "UPDATE SalesReps SET LastName='" & EscapeSQuotes(Me.SalesRepLastName) & "' WHERE ID=" & Me.SalesRepID
            changed = False
            fixTitleArea
            fixNameList
        End If
    End If
End Sub

Private Sub SalesRepPosition_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            SalesRepPosition_LostFocus
        Case Is = vbKeyDelete
            SalesRepPosition_KeyPress KeyCode
    End Select
End Sub

Private Sub SalesRepPosition_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "SalesRepPosition"
End Sub

Private Sub SalesRepPosition_LostFocus()
    If changed = True Then
        If Len(Me.SalesRepPosition) > 50 Then
            MsgBox "Position too long! Must be under 50 chars."
            Me.SalesRepPosition.SetFocus
        Else
            DB.Execute "UPDATE SalesReps SET PositionTitle='" & EscapeSQuotes(Me.SalesRepPosition) & "' WHERE ID=" & Me.SalesRepID
            changed = False
            fixTitleArea
        End If
    End If
End Sub

Private Sub SalesRepAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            SalesRepAddress_LostFocus
        Case Is = vbKeyDelete
            SalesRepAddress_KeyPress KeyCode
    End Select
End Sub

Private Sub SalesRepAddress_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "SalesRepAddress"
End Sub

Private Sub SalesRepAddress_LostFocus()
    If changed = True Then
        If Len(Me.SalesRepAddress) > 50 Then
            MsgBox "Address too long! Must be under 50 chars."
            Me.SalesRepAddress.SetFocus
        Else
            DB.Execute "UPDATE SalesReps SET Address='" & EscapeSQuotes(Me.SalesRepAddress) & "' WHERE ID=" & Me.SalesRepID
            changed = False
        End If
    End If
End Sub

Private Sub SalesRepCity_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            SalesRepCity_LostFocus
        Case Is = vbKeyDelete
            SalesRepCity_KeyPress KeyCode
    End Select
End Sub

Private Sub SalesRepCity_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "SalesRepCity"
End Sub

Private Sub SalesRepCity_LostFocus()
    If changed = True Then
        If Len(Me.SalesRepCity) > 50 Then
            MsgBox "City too long! Must be under 50 chars."
            Me.SalesRepCity.SetFocus
        Else
            DB.Execute "UPDATE SalesReps SET City='" & EscapeSQuotes(Me.SalesRepCity) & "' WHERE ID=" & Me.SalesRepID
            changed = False
        End If
    End If
End Sub

Private Sub SalesRepState_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            SalesRepState_LostFocus
        Case Is = vbKeyDelete
            SalesRepState_KeyPress KeyCode
    End Select
End Sub

Private Sub SalesRepState_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "SalesRepState"
End Sub

Private Sub SalesRepState_LostFocus()
    If changed = True Then
        If Len(Me.SalesRepState) > 50 Then
            MsgBox "State too long! Must be under 50 chars."
            Me.SalesRepState.SetFocus
        Else
            DB.Execute "UPDATE SalesReps SET State='" & EscapeSQuotes(Me.SalesRepState) & "' WHERE ID=" & Me.SalesRepID
            changed = False
        End If
    End If
End Sub

Private Sub SalesRepZipCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            SalesRepZipCode_LostFocus
        Case Is = vbKeyDelete
            SalesRepZipCode_KeyPress KeyCode
    End Select
End Sub

Private Sub SalesRepZipCode_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "SalesRepZipCode"
End Sub

Private Sub SalesRepZipCode_LostFocus()
    If changed = True Then
        If Len(Me.SalesRepZipCode) > 10 Then
            MsgBox "Zip code too long! Must be under 10 chars."
            Me.SalesRepZipCode.SetFocus
        Else
            DB.Execute "UPDATE SalesReps SET ZipCode='" & EscapeSQuotes(Me.SalesRepZipCode) & "' WHERE ID=" & Me.SalesRepID
            changed = False
        End If
    End If
End Sub

Private Sub SalesRepPhone_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            SalesRepPhone_LostFocus
        Case Is = vbKeyDelete
            SalesRepPhone_KeyPress KeyCode
    End Select
End Sub

Private Sub SalesRepPhone_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "SalesRepPhone"
End Sub

Private Sub SalesRepPhone_LostFocus()
    If changed = True Then
        If Len(Me.SalesRepPhone) > 17 Then
            MsgBox "Phone number too long! Must be under 17 chars."
            Me.SalesRepPhone.SetFocus
        Else
            DB.Execute "UPDATE SalesReps SET Phone='" & EscapeSQuotes(Me.SalesRepPhone) & "' WHERE ID=" & Me.SalesRepID
            changed = False
        End If
    End If
End Sub

Private Sub SalesRepCell_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            SalesRepCell_LostFocus
        Case Is = vbKeyDelete
            SalesRepCell_KeyPress KeyCode
    End Select
End Sub

Private Sub SalesRepCell_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "SalesRepCell"
End Sub

Private Sub SalesRepCell_LostFocus()
    If changed = True Then
        If Len(Me.SalesRepCell) > 17 Then
            MsgBox "Cell number too long! Must be under 17 chars."
            Me.SalesRepCell.SetFocus
        Else
            DB.Execute "UPDATE SalesReps SET Cell='" & EscapeSQuotes(Me.SalesRepCell) & "' WHERE ID=" & Me.SalesRepID
            changed = False
        End If
    End If
End Sub

Private Sub SalesRepFax_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            SalesRepFax_LostFocus
        Case Is = vbKeyDelete
            SalesRepFax_KeyPress KeyCode
    End Select
End Sub

Private Sub SalesRepFax_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "SalesRepFax"
End Sub

Private Sub SalesRepFax_LostFocus()
    If changed = True Then
        If Len(Me.SalesRepFax) > 17 Then
            MsgBox "Fax number too long! Must be under 17 chars."
            Me.SalesRepFax.SetFocus
        Else
            DB.Execute "UPDATE SalesReps SET Fax='" & EscapeSQuotes(Me.SalesRepFax) & "' WHERE ID=" & Me.SalesRepID
            changed = False
        End If
    End If
End Sub

Private Sub SalesRepEmail_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            SalesRepEmail_LostFocus
        Case Is = vbKeyDelete
            SalesRepEmail_KeyPress KeyCode
    End Select
End Sub

Private Sub SalesRepEmail_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "SalesRepEmail"
End Sub

Private Sub SalesRepEmail_LostFocus()
    If changed = True Then
        If Len(Me.SalesRepEmail) > 50 Then
            MsgBox "Email address too long! Must be under 50 chars."
            Me.SalesRepEmail.SetFocus
        Else
            DB.Execute "UPDATE SalesReps SET Email='" & EscapeSQuotes(Me.SalesRepEmail) & "' WHERE ID=" & Me.SalesRepID
            changed = False
        End If
    End If
End Sub

Private Sub SalesRepNotes_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            SalesRepNotes_LostFocus
        Case Is = vbKeyDelete
            SalesRepNotes_KeyPress KeyCode
    End Select
End Sub

Private Sub SalesRepNotes_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "SalesRepNotes"
End Sub

Private Sub SalesRepNotes_LostFocus()
    If changed = True Then
        DB.Execute "UPDATE SalesReps SET Notes='" & EscapeSQuotes(Me.SalesRepNotes) & "' WHERE ID=" & Me.SalesRepID
        changed = False
    End If
End Sub

