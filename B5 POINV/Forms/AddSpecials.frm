VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#3.5#0"; "TPControls.ocx"
Begin VB.Form AddSpecials 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Specials"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13470
   ControlBox      =   0   'False
   Icon            =   "AddSpecials.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   13470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Include Special on Feeds"
      Height          =   1815
      Left            =   120
      TabIndex        =   56
      Top             =   3840
      Width           =   5775
      Begin VB.CheckBox Check2 
         Caption         =   "Tool Parts Froogle (not setup)"
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tools Plus Froogle"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton btnMakeSaleReportString 
      Caption         =   "generate sale transaction report string"
      Height          =   405
      Left            =   7170
      TabIndex        =   55
      Top             =   7920
      Width           =   2055
   End
   Begin VB.TextBox FlatPercentDiscount 
      Height          =   285
      Left            =   990
      TabIndex        =   54
      Top             =   9210
      Width           =   945
   End
   Begin VB.CommandButton btnSalesLookup2 
      Caption         =   "free item #2"
      Height          =   315
      Left            =   9360
      TabIndex        =   52
      Top             =   7980
      Width           =   1545
   End
   Begin VB.TextBox PickingWarning 
      Height          =   315
      Left            =   1170
      MaxLength       =   64
      TabIndex        =   51
      Top             =   5880
      Width           =   5055
   End
   Begin VB.CommandButton btnSalesLookup 
      Caption         =   "lookup free item sales"
      Height          =   465
      Left            =   9420
      TabIndex        =   49
      Top             =   7440
      Width           =   1155
   End
   Begin VB.Frame Frame3 
      Height          =   555
      Left            =   3240
      TabIndex        =   47
      Top             =   3240
      Width           =   2595
      Begin VB.CheckBox chkFreeShipping 
         Caption         =   "Free Shipping (feed)"
         Height          =   255
         Left            =   90
         TabIndex        =   48
         ToolTipText     =   "Mark as free shipping in the feeds (you still need to set up a coupon or something!)"
         Top             =   210
         Width           =   2235
      End
   End
   Begin VB.CommandButton btnMakeCouponString 
      Caption         =   "generate Y! coupon string"
      Height          =   345
      Left            =   7170
      TabIndex        =   46
      Top             =   7500
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Height          =   1425
      Left            =   3240
      TabIndex        =   41
      Top             =   1920
      Width           =   2595
      Begin VB.OptionButton opShowOnTab 
         Caption         =   "Show above tab control"
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   45
         Top             =   1110
         Width           =   2235
      End
      Begin VB.OptionButton opShowOnTab 
         Caption         =   "Don't show anywhere"
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   44
         Top             =   210
         Width           =   2235
      End
      Begin VB.OptionButton opShowOnTab 
         Caption         =   "Show on Specials tab"
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   43
         Top             =   510
         Width           =   2055
      End
      Begin VB.OptionButton opShowOnTab 
         Caption         =   "Show on Notes tab"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   42
         Top             =   810
         Width           =   2235
      End
   End
   Begin VB.Frame Frame1 
      Height          =   885
      Left            =   3240
      TabIndex        =   38
      Top             =   1080
      Width           =   2595
      Begin VB.OptionButton opShowOnRebatePage 
         Caption         =   "Don't include on R&&S page"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   40
         Top             =   510
         Width           =   2235
      End
      Begin VB.OptionButton opShowOnRebatePage 
         Caption         =   "Include on R&&S page"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   210
         Width           =   2055
      End
   End
   Begin VB.TextBox Notes 
      Height          =   1245
      Left            =   6660
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   36
      Top             =   6090
      Width           =   4395
   End
   Begin VB.TextBox SpecialTitle 
      Height          =   285
      Left            =   1110
      TabIndex        =   34
      Top             =   6570
      Width           =   5115
   End
   Begin VB.TextBox RebateFormURL 
      Height          =   285
      Left            =   3930
      TabIndex        =   32
      Top             =   6240
      Width           =   2295
   End
   Begin TPControls.SimpleListView AppliedTo 
      Height          =   5445
      Left            =   6390
      TabIndex        =   30
      Top             =   300
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   9604
      MultiSelect     =   0   'False
      Sorted          =   0   'False
      Enabled         =   -1  'True
   End
   Begin VB.CheckBox chkHideDiscontinued 
      Caption         =   "Hide Discontinued"
      Height          =   255
      Left            =   11520
      TabIndex        =   29
      Top             =   6990
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Frame generalFrame 
      Height          =   1485
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   1170
      Width           =   3015
      Begin TPControls.DateDropdown SpecialEndDate 
         Height          =   315
         Left            =   960
         TabIndex        =   21
         Top             =   570
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
      End
      Begin TPControls.DateDropdown SpecialStartDate 
         Height          =   315
         Left            =   960
         TabIndex        =   22
         Top             =   210
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
      End
      Begin TPControls.DateDropdown SpecialPostmarkDate 
         Height          =   315
         Left            =   960
         TabIndex        =   24
         Top             =   930
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
      End
      Begin VB.Label generalLabel 
         Caption         =   "End Date:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   630
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Caption         =   "PM Date:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   795
      End
      Begin VB.Label generalLabel 
         Caption         =   "Start Date:"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Section Page"
      Height          =   1005
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   3015
      Begin VB.ComboBox SectionURL 
         Height          =   315
         Left            =   630
         TabIndex        =   17
         Top             =   600
         Width           =   2355
      End
      Begin VB.ComboBox SectionImage 
         Height          =   315
         Left            =   630
         TabIndex        =   16
         Top             =   240
         Width           =   2355
      End
      Begin VB.Label generalLabel 
         Caption         =   "URL:"
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   19
         Top             =   630
         Width           =   525
      End
      Begin VB.Label generalLabel 
         Caption         =   "Image:"
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   18
         Top             =   270
         Width           =   525
      End
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Filtering"
      Height          =   1065
      Index           =   0
      Left            =   11400
      TabIndex        =   14
      Top             =   5850
      Width           =   1935
      Begin VB.OptionButton opPNFilter 
         Caption         =   "All Lines"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   750
         Width           =   1665
      End
      Begin VB.OptionButton opPNFilter 
         Caption         =   "Related Lines"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "Show this line and other lines from the same manufacturer"
         Top             =   510
         Width           =   1665
      End
      Begin VB.OptionButton opPNFilter 
         Caption         =   "Current Line Only"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Show only this line code (and accessory lines, if any)"
         Top             =   270
         Value           =   -1  'True
         Width           =   1665
      End
   End
   Begin VB.CheckBox chkFilterSpecials 
      Caption         =   "Don't show old"
      Height          =   195
      Left            =   4080
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   390
      Value           =   1  'Checked
      Width           =   1395
   End
   Begin VB.CommandButton btnToggleText 
      Caption         =   "Text:"
      Height          =   225
      Left            =   150
      TabIndex        =   11
      Top             =   6990
      Width           =   675
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   2760
      TabIndex        =   9
      Top             =   9570
      Width           =   1365
   End
   Begin VB.CommandButton btnMoveRight 
      Caption         =   "=>"
      Height          =   495
      Left            =   10830
      TabIndex        =   6
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton btnMoveLeft 
      Caption         =   "<="
      Height          =   495
      Left            =   10830
      TabIndex        =   5
      Top             =   1380
      Width           =   495
   End
   Begin VB.ListBox AllPartNumbers 
      Height          =   5325
      Left            =   11340
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   300
      Width           =   1995
   End
   Begin VB.CommandButton btnNewSpecial 
      Caption         =   "Create New Special"
      Height          =   315
      Left            =   3990
      TabIndex        =   3
      Top             =   690
      Width           =   1605
   End
   Begin VB.TextBox SpecialText 
      Height          =   2265
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   6900
      Width           =   5385
   End
   Begin VB.ComboBox SpecialName 
      Height          =   315
      Left            =   990
      TabIndex        =   0
      Top             =   690
      Width           =   2955
   End
   Begin SHDocVwCtl.WebBrowser SpecialTextPreview 
      Height          =   2235
      Left            =   840
      TabIndex        =   10
      Top             =   6900
      Width           =   5385
      ExtentX         =   9499
      ExtentY         =   3942
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label generalLabel 
      Caption         =   "% Discount:"
      Height          =   255
      Index           =   11
      Left            =   60
      TabIndex        =   53
      Top             =   9240
      Width           =   885
   End
   Begin VB.Label generalLabel 
      Caption         =   "Pick Warning:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   50
      Top             =   5910
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Notes:"
      Height          =   195
      Left            =   6660
      TabIndex        =   37
      Top             =   5880
      Width           =   945
   End
   Begin VB.Label generalLabel 
      Caption         =   "Special Title:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   33
      Top             =   6600
      Width           =   1005
   End
   Begin VB.Label generalLabel 
      Caption         =   "http://p1.hostingprod.com/@tools-plus.com/rebates/"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   31
      Top             =   6270
      Width           =   3795
   End
   Begin VB.Label Label1 
      Caption         =   "Add/Edit Specials"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   90
      Width           =   2715
   End
   Begin VB.Label generalLabel 
      Caption         =   "Applied To:"
      Height          =   255
      Index           =   5
      Left            =   6420
      TabIndex        =   8
      Top             =   60
      Width           =   1185
   End
   Begin VB.Label generalLabel 
      Caption         =   "Item Numbers:"
      Height          =   255
      Index           =   4
      Left            =   11370
      TabIndex        =   7
      Top             =   60
      Width           =   1425
   End
   Begin VB.Label generalLabel 
      Caption         =   "Name:"
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   720
      Width           =   795
   End
End
Attribute VB_Name = "AddSpecials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private changed As Boolean
'Private itemAssocChanged As Boolean
Private fillingForm As Boolean

Private Const PNF_THIS As Long = 0
Private Const PNF_REL As Long = 1
Private Const PNF_ALL As Long = 2

Private Sub AllPartNumbers_DblClick()
    btnMoveLeft_Click
End Sub

Private Sub AppliedTo_DblClick(i As Long, j As Long, txt As String)
    If Me.SpecialName = "" Then
        Exit Sub
    End If
    Dim newval As String
    Select Case j
        Case Is = 0
            btnMoveRight_Click
        Case Is = 1
            newval = InputBox("Enter rebate value for " & Me.AppliedTo.GetText(i, 0) & ":", , txt)
            If newval <> "" Then
                If IsNumeric(newval) Then
                    newval = Format(newval, "Currency")
                    DB.Execute "UPDATE PartNumberSpecials SET RebateAmount=" & Replace(newval, "$", "") & " WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "' AND ItemNumber='" & Me.AppliedTo.GetText(i, 0) & "'"
                    Me.AppliedTo.Edit newval, i, j
                Else
                    MsgBox newval & " is not numeric!"
                End If
            End If
        Case Is = 2
            If txt = "N" Then
                DB.Execute "UPDATE PartNumberSpecials SET InstantRebate=1 WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "' AND ItemNumber='" & Me.AppliedTo.GetText(i, 0) & "'"
                Me.AppliedTo.Edit "Y", i, j
            Else
                DB.Execute "UPDATE PartNumberSpecials SET InstantRebate=0 WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "' AND ItemNumber='" & Me.AppliedTo.GetText(i, 0) & "'"
                Me.AppliedTo.Edit "N", i, j
            End If
        Case Is = 3
            newval = InputBox("Enter quantity available for the special (or 0 to remove):", , txt)
            If newval <> "" Then
                If IsNumeric(newval) Then
                    If newval = "0" Then
                        DB.Execute "UPDATE PartNumberSpecials SET AvailQuantity=NULL WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "' AND ItemNumber='" & Me.AppliedTo.GetText(i, 0) & "'"
                        Me.AppliedTo.Edit "", i, j
                    Else
                        DB.Execute "UPDATE PartNumberSpecials SET AvailQuantity=" & newval & " WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "' AND ItemNumber='" & Me.AppliedTo.GetText(i, 0) & "'"
                        Me.AppliedTo.Edit newval, i, j
                    End If
                Else
                    MsgBox "Value must be numeric!"
                End If
            End If
    End Select
End Sub

Private Sub btnMakeCouponString_Click()
    Dim couponstr As String
    couponstr = ""
    Dim i As Long
    For i = 0 To Me.AppliedTo.ListCount - 1
        couponstr = couponstr & YahooIDFunctions.CreateYahooID(Me.AppliedTo.GetText(i, 0)) & ","
    Next i
    If Len(couponstr) > 0 Then
        couponstr = Left(couponstr, Len(couponstr) - 1)
        Open DESTINATION_DIR & "couponstr.txt" For Output As #1
        Print #1, couponstr
        Close #1
        OpenDefaultApp DESTINATION_DIR & "couponstr.txt"
    Else
        MsgBox "No items?"
    End If
End Sub

Private Sub btnMakeSaleReportString_Click()
    Dim strString As String
    strString = ""
    Dim i As Long
    For i = 0 To Me.AppliedTo.ListCount - 1
        strString = strString & Me.AppliedTo.GetText(i, 0) & vbCrLf
    Next i
    Open DESTINATION_DIR & "couponstr.txt" For Output As #1
    Print #1, strString
    Close #1
    OpenDefaultApp DESTINATION_DIR & "couponstr.txt"
End Sub

'Private Sub btnApplyToAllItems_Click()
'    If MsgBox("Are you sure you want to apply this special to all items?", vbYesNo) = vbYes Then
'        Mouse.Hourglass True
'        Dim rst As ADODB.Recordset
'        Set rst = DB.retrieve("SELECT ItemNumber FROM PartNumbers WHERE WebDeleted=0")
'        While Not rst.EOF
'            DB.Execute "INSERT INTO PartNumberSpecials ( ItemNumber, SpecialName ) VALUES ( '" & rst("ItemNumber") & "', '" & EscapeSQuotes(Me.SpecialName) & "' )", True
'            rst.MoveNext
'            DoEvents
'        Wend
'        rst.Close
'        Set rst = Nothing
'        requeryAppliedTo
'        Mouse.Hourglass False
'    End If
'End Sub
'
'Private Sub btnApplyToLine_Click()
'    Dim lc As String
'    lc = InputBox("Apply this special to all items in which line code?")
'    If lc <> "" Then
'        Dim rst As ADODB.Recordset
'        Set rst = DB.retrieve("SELECT ItemNumber FROM PartNumbers WHERE LEFT(ItemNumber,3)='" & lc & "' AND WebDeleted=0")
'        While Not rst.EOF
'            DB.Execute "INSERT INTO PartNumberSpecials ( ItemNumber, SpecialName ) VALUES ( '" & rst("ItemNumber") & "', '" & EscapeSQuotes(Me.SpecialName) & "' )", True
'            rst.MoveNext
'        Wend
'        rst.Close
'        Set rst = Nothing
'        requeryAppliedTo
'    End If
'End Sub

Private Sub btnMoveLeft_Click()
'first we need to check if this item is ok to be added to current special as some feed checks
'are most likely checked off and this means only 1 special per feed per item.


'If Check1.value = vbChecked Then
'    Dim query As String
'    query = "SELECT COUNT(SpecialName) AS SpecialsCountInFeed FROM PartNumberSpecials WHERE ItemNumber='{ITEMNUMBER}' AND SpecialName IN (SELECT SpecialName FROM SpecialsFeedFilter WHERE FeedName='Feeds_Froogle')"
'    query = Replace(query, "{ITEMNUMBER}", Me.AllPartNumbers.list(Me.AllPartNumbers.ListIndex))
'    Dim Results As ADODB.Recordset
'    Set Results = DB.retrieve(query)
'    If Results.RecordCount > 0 Then
'        Do
'            Dim thecount As Integer
'            thecount = CInt(Results("SpecialsCountInFeed"))
'            If thecount > 0 Then
'                MsgBox "Item " & Me.AllPartNumbers.list(Me.AllPartNumbers.ListIndex) & " is already on a special under the Feeds_Froogle feed and therefore cannot be added."
'                Exit Sub
'            End If
'            Results.MoveNext
'        Loop Until Results.EOF = True Or Results.BOF = True
'
'    End If
'End If

    
    If Me.SpecialName = "" Then
        Exit Sub
    End If
    If Me.AllPartNumbers.SelCount > 0 Then
        Dim i As Long
        For i = Me.AllPartNumbers.ListCount - 1 To 0 Step -1
            If Me.AllPartNumbers.Selected(i) Then
                DB.Execute "INSERT INTO PartNumberSpecials ( ItemNumber, SpecialName ) VALUES ( '" & Me.AllPartNumbers.list(i) & "', '" & EscapeSQuotes(Me.SpecialName) & "' )"
                Me.AppliedTo.Add Array(Me.AllPartNumbers.list(i), "$0.00", "N")
                Me.AllPartNumbers.RemoveItem i
            End If
        Next i
        'itemAssocChanged = True
    End If
End Sub

Private Sub btnMoveRight_Click()
    If Me.SpecialName = "" Then
        Exit Sub
    End If
    If Me.AppliedTo.SelCount > 0 Then
        Dim rows As Variant, thisdx As Variant, item As String
        rows = Me.AppliedTo.SelIndex
        For Each thisdx In rows
            item = Me.AppliedTo.GetText(CLng(thisdx), 0)
            DB.Execute "DELETE FROM PartNumberSpecials WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "' AND ItemNumber='" & item & "'"
            
            'we will add to the full list even if that is filtered, since it
            'will make it easier to add back after a mis-click
            Me.AllPartNumbers.AddItem item
        Next thisdx
        
        'keep this out of the loop, not sure if the selindex list will
        'get messed up after removing items, it does in c#
        Me.AppliedTo.RemoveSelected
    End If
End Sub

Private Sub btnNewSpecial_Click()
    'If itemAssocChanged = True Then
    '    If MsgBox("Do you want to save item associations for this special first?", vbYesNo) = vbYes Then
    '        saveItemAssoc
    '    End If
    '    itemAssocChanged = False
    'End If
    Dim sname As String
    sname = InputBox("Enter the name for the special:")
    If sname <> "" Then
        Me.SpecialName = sname
        DB.Execute "INSERT INTO Specials ( SpecialName, SpecialStartDate, SpecialEndDate, SpecialPostmarkDate ) VALUES ( '" & EscapeSQuotes(sname) & "', '" & Format(Now(), "Short Date") & "', '" & Format(Now(), "Short Date") & "', '" & Format(Now(), "Short Date") & "' )"
        Me.SpecialName.AddItem sname
        fillForm
    End If
End Sub

Private Sub btnOK_Click()
    If Me.SpecialName = "" Then
        'ok to close
    Else
        'saveItemAssoc
        'If OptionGroupHasSelection(Me.opShowOnRebatePage) Then
        '    'continue
        'Else
        '    MsgBox "You need to select whether this should show on the rebates and specials page!"
        '    Exit Sub
        'End If
        '
        'If OptionGroupHasSelection(Me.opShowOnTab) Then
        '    'continue
        'Else
        '    MsgBox "You need to select which tab this will show up on!"
        '    Exit Sub
        'End If
    
        fixPostmarkDate
    End If
    
    Unload AddSpecials
End Sub

Private Sub btnSalesLookup_Click()
    If Me.AppliedTo.SelCount = 0 Then
        MsgBox "Nothing selected!"
        Exit Sub
    End If
    If Me.AppliedTo.GetTextSelected(3)(0) = "" Then
        MsgBox "Selected item doesn't have a Qty Avail!"
        Exit Sub
    End If
    
    Dim rst As ADODB.Recordset
    'this should be the same as the one in vSpecialsRemove
    Set rst = DB.retrieve("SELECT SUM(-1 * TransactionQuantity) " & _
                          "FROM IM_ItemTransactionHistory " & _
                          "WHERE ItemCode='" & Me.AppliedTo.GetTextSelected(0)(0) & "' " & _
                          "  AND TransactionDate>='" & Me.SpecialStartDate.value & "' " & _
                          "  AND WarehouseCode='000' " & _
                          "  AND TransactionCode IN ('P2','SO','SI')" & _
                          "  AND TransactionQuantity<0")
    MsgBox "Total sold since " & Me.SpecialStartDate.value & " is: " & Nz(rst(0), "?")
    rst.Close
    Set rst = Nothing
End Sub

Private Sub btnSalesLookup2_Click()
    Dim item As String
    item = InputBox("This is for free item specials that have the free item added into the Mas order." & vbCrLf & vbCrLf & "Enter the item number of the free item to see how many we've given away (separate multiple with commas):", "Enter itemnumber", "D-ADCB120,D-ADCB201,D-ADCB200")
    If item = "" Then
        Exit Sub
    End If
    
    Mouse.Hourglass True
    
    Dim masdate As String
    masdate = Format(Me.SpecialStartDate.value, "YYYY-MM-DD")
    
    Dim items As Variant
    items = Split(item, ",")
    Dim rst As ADODB.Recordset
    Dim iter As Variant
    Dim msg As String
    msg = "ItemNumber" & vbTab & "Shipped" & vbTab & "Pending" & vbTab & "Total"
    For Each iter In items
        Set rst = MASDB.retrieve("SELECT SUM(-1 * TransactionQty) " & _
                                 "FROM IM_ItemTransactionHistory " & _
                                 "WHERE ItemCode='" & EscapeSQuotes(CStr(iter)) & "' " & _
                                 "  AND TransactionDate>={ d '" & masdate & "' } " & _
                                 "  AND WarehouseCode='000' " & _
                                 "  AND TransactionCode IN ('P2','SO','SI')" & _
                                 "  AND TransactionQuantity<0 " & _
                                 "  AND UnitPrice=0")
        Dim comp As Long
        comp = CLng(Nz(rst(0), "0"))
        rst.Close
        
        Set rst = MASDB.retrieve("SELECT SUM(QuantityOrdered) " & _
                                 "FROM SO_SalesOrderHeader, SO_SalesOrderDetail " & _
                                 "WHERE SO_SalesOrderHeader.SalesOrderNo=SO_SalesOrderDetail.SalesOrderNo " & _
                                 "  AND SO_SalesOrderDetail.ItemCode='" & EscapeSQuotes(CStr(iter)) & "' " & _
                                 "  AND SO_SalesOrderHeader.OrderDate>={ d '" & masdate & "' } " & _
                                 "  AND SO_SalesOrderDetail.WarehouseCode='000' " & _
                                 "  AND SO_SalesOrderDetail.QuantityOrdered>0 " & _
                                 "  AND SO_SalesOrderDetail.UnitPrice=0")
        Dim pending As Long
        pending = CLng(Nz(rst(0), "0"))
        rst.Close
        
        msg = msg & vbCrLf & CStr(iter) & vbTab & comp & vbTab & pending & vbTab & (comp + pending)
    Next iter
    
    Mouse.Hourglass False
    
    MsgBox msg
End Sub

Private Sub btnToggleText_Click()
    If Me.SpecialText.Visible = True Then
        Me.SpecialText.Visible = False
        Me.SpecialTextPreview.Visible = True
        Open DESTINATION_DIR & "special.html" For Output As #1
        Print #1, "<html>" & vbCrLf & "<body bgcolor=""#CEDDE4"">" & vbCrLf & Me.SpecialText & vbCrLf & "</body>" & vbCrLf & "</html>"
        Close #1
        Me.SpecialTextPreview.Navigate2 DESTINATION_DIR & "special.html"
    Else
        Me.SpecialTextPreview.Visible = False
        Me.SpecialText.Visible = True
    End If
End Sub
Private Sub CheckFeedPromos()
    Check1.value = vbUnchecked
    'RemoveOldPromosFromFeeds
    
    Dim query As String
    
    query = "SELECT FeedName FROM SpecialsFeedFilter WHERE SpecialName='" & Me.SpecialName.Text & "'"
    Dim Results As ADODB.Recordset
    Set Results = DB.retrieve(query)
    If Results.RecordCount > 0 Then
    
    Do
        If Results("FeedName") = "Feeds_Froogle" Then Check1.value = vbChecked
        Results.MoveNext
    Loop Until Results.EOF = True Or Results.BOF = True
    End If
End Sub
Private Sub Check1_Click()
'MsgBox "temporarily disabled. not in use at the moment."

Feeds_Froogle True
End Sub
Private Sub Feeds_Froogle(userClicked As Boolean)
'Tools Plus Froogle Feed Filter Check

If Check1.value = vbChecked Then
Dim check As Boolean

check = CheckForExistingPromo("Feeds_Froogle")
'MsgBox CStr(check)
    If check = False Then
        DB.Execute "DELETE FROM SpecialsFeedFilter WHERE FeedName='Feeds_Froogle' AND SpecialName='" & Replace(Me.SpecialName.Text, "'", "''") & "'"
        DB.Execute "INSERT INTO SpecialsFeedFilter (FeedName,SpecialName) VALUES('Feeds_Froogle','" & Replace(Me.SpecialName.Text, "'", "''") & "')"
    End If
    If check = True Then
        MsgBox "Cannot add to Froogle Feed. Froogle Feed already has a promo on one of these items."
        Check1.value = vbUnchecked
    End If
Else
    DB.Execute "DELETE FROM SpecialsFeedFilter WHERE FeedName='Feeds_Froogle' AND SpecialName='" & Replace(Me.SpecialName.Text, "'", "''") & "'"
End If
End Sub
Private Sub RemoveOldPromosFromFeeds()
    Dim promos As ADODB.Recordset
    Set promos = DB.retrieve("SELECT SpecialName FROM Specials WHERE SpecialEndDate < DATEADD(day,1,GETDATE())")
    Do
        Dim promoToCheck As String
        promoToCheck = Replace(promos("SpecialName"), "'", "''")
        DB.Execute "DELETE FROM SpecialsFeedFilter WHERE SpecialName='" & promoToCheck & "'"
        promos.MoveNext
    Loop Until promos.EOF = True Or promos.BOF = True


End Sub
Private Function CheckForExistingPromo(tableName As String) As Boolean
'Only 1 promo per item per feed... need to check to make sure none others are selected for froogle.
Dim query As String
'Query = "SELECT PartNumberSpecials.ItemNumber FROM Specials INNER JOIN PartNumberSpecials ON PartNumberSpecials.SpecialName=Specials.SpecialName WHERE SpecialEndDate > GETDATE() AND ItemNumber='{ITEMNUMBER}' AND Specials.SpecialName IN (SELECT SpecialsFeedFilter.SpecialName FROM SpecialsFeedFilter WHERE FeedName='" & TableName & "')"
query = "SELECT SpecialName FROM PartNumberSpecials WHERE ItemNumber='{ITEMNUMBER}' AND SpecialName IN (SELECT SpecialName FROM SpecialsFeedFilter WHERE FeedName='" & tableName & "') AND SpecialName <> '" & Me.SpecialName.Text & "'"

Dim itm As Variant
Dim newQuery As String
Dim countY As Integer
Dim errStr As String

'For Each itm In AllPartNumbers.
'For countY = 0 To AllPartNumbers.ListCount - 1
'MsgBox "listcount: " & CStr(Me.AppliedTo.ListCount)
For countY = 0 To Me.AppliedTo.ListCount - 1
    Dim wtf As String
    wtf = AppliedTo.GetText(CLng(countY), 0)
    'MsgBox CStr(Me.AppliedTo.GetText(countY, 0))
    'newQuery = Replace(query, "{ITEMNUMBER}", AllPartNumbers.list(countY))
    newQuery = Replace(query, "{ITEMNUMBER}", wtf)
    Dim Results As ADODB.Recordset
    Set Results = DB.retrieve(newQuery)
    If Results.RecordCount > 0 Then
        errStr = errStr & wtf & ", "
    End If
Next
If Len(errStr) > 0 Then
    errStr = errStr & " items are on another promotion in the " & tableName & " feed and therefore cannot allow a second special on the feed."
    MsgBox errStr
    CheckForExistingPromo = True
Else
    CheckForExistingPromo = False
End If
End Function


Private Sub Check2_Click()
MsgBox "temporarily disabled. not in use at the moment."
End Sub

Private Sub chkFilterSpecials_Click()
    requerySpecials
End Sub

Private Sub chkFreeShipping_Click()
    If Me.SpecialName = "" Then
        Exit Sub
    End If
    If Not fillingForm Then
        DB.Execute "UPDATE Specials SET FreeShipping=" & SQLBool(Me.chkFreeShipping) & " WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
    End If
End Sub

Private Sub FlatPercentDiscount_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii, False, False)
    If KeyAscii <> 0 Then
        changed = True
    End If
End Sub

Private Sub FlatPercentDiscount_LostFocus()
    If changed Then
        Dim disc As Long
        If Me.FlatPercentDiscount = "" Then
            Me.FlatPercentDiscount = "0"
        End If
        disc = CLng(Me.FlatPercentDiscount)
        DB.Execute "UPDATE Specials SET FlatPercentDiscount=" & CStr(disc) & " WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
        If disc = 0 Then
            Me.FlatPercentDiscount = ""
        End If
        changed = False
    End If
End Sub

'Private Sub chkInfoSpecial_Click()
'    If Me.SpecialName = "" Then
'        Exit Sub
'    End If
'    If Not fillingform Then
'        DB.Execute "UPDATE Specials SET NotRebate=" & SQLBool(Me.chkInfoSpecial) & " WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
'    End If
'End Sub

Private Sub Form_Load()
    Me.AppliedTo.MultiSelect = True
    Me.AppliedTo.sorted = True
    Me.AppliedTo.SetColumnNames Array("ItemNumber", "Rebate Amt", "Instant", "Qty Avail")
    Me.AppliedTo.SetColumnWidths Array(1500, 1100, 700, 900)
    
    'itemAssocChanged = False
    requerySpecials
    requeryAllPartNumbers
    requerySectionURL
    requerySectionImage
    ExpandDropDownToFit Me.SpecialName
    ExpandDropDownToFit Me.SectionImage
    ExpandDropDownToFit Me.SectionURL
    If SignMaintenance.Specials.SelCount = 0 And SignMaintenance.Specials.ListCount = 1 Then
        SignMaintenance.Specials.Selected(0) = True
    End If
    If SignMaintenance.Specials.SelCount <> 0 Then
        Dim i As Long
        For i = 0 To SignMaintenance.Specials.ListCount - 1
            If SignMaintenance.Specials.Selected(i) Then
                Me.SpecialName = SignMaintenance.Specials.list(i)
                fillForm
                i = SignMaintenance.Specials.ListCount
            End If
        Next i
    End If
    
    CheckFeedPromos
End Sub


Private Sub requerySpecials()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT SpecialName FROM Specials" & IIf(Me.chkFilterSpecials, " WHERE SpecialEndDate>(GETDATE()-120)", ""))
    Me.SpecialName.Clear
    While Not rst.EOF
        Me.SpecialName.AddItem (rst("SpecialName"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub fillForm()
    fillingForm = True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT SpecialTitle, SpecialStartDate, SpecialEndDate, SpecialPostmarkDate, SpecialText, SectionURL, SectionImage, ShowOnRebatePage, ShowOnTab, RebateFormURL, FlatPercentDiscount, InternalNotes, FreeShipping, PickingWarning FROM Specials WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'")
    Me.SpecialTitle = Nz(rst("SpecialTitle"))
    Me.SpecialStartDate.value = Format(rst("SpecialStartDate"), "Short Date")
    Me.SpecialEndDate.value = Format(rst("SpecialEndDate"), "Short Date")
    Me.SpecialPostmarkDate.value = Format(rst("SpecialPostmarkDate"), "Short Date")
    Me.RebateFormURL = Nz(rst("RebateFormURL"))
    Me.SpecialText = Nz(rst("SpecialText"))
    Me.SpecialText.Visible = True
    Me.SpecialTextPreview.Visible = False
    Me.SectionImage = Nz(rst("SectionImage"))
    Me.SectionURL = Nz(rst("SectionURL"))
    Me.chkFreeShipping = SQLBool(rst("FreeShipping"))
    Me.PickingWarning = Nz(rst("PickingWarning"))
    Select Case rst("ShowOnRebatePage")
        Case Is = False 'do not show
            Me.opShowOnRebatePage(1) = 1
        Case Is = True 'show
            Me.opShowOnRebatePage(0) = 1
    End Select
    Select Case rst("ShowOnTab")
        Case Is = 0 'no tab
            Me.opShowOnTab(2) = 1
        Case Is = 1 'specials tab
            Me.opShowOnTab(0) = 1
        Case Is = 2 'info tab
            Me.opShowOnTab(1) = 1
        Case Is = 3 'above tabs
            Me.opShowOnTab(3) = 1
    End Select
    Me.Notes = Nz(rst("InternalNotes"))
    Me.FlatPercentDiscount = IIf(CLng(rst("FlatPercentDiscount")) = 0, "", CStr(CLng(rst("FlatPercentDiscount"))))
    requeryAppliedTo
    fillingForm = False
End Sub

Private Sub clearForm()
    Me.SpecialName = ""
    Me.SpecialStartDate.value = Date
    Me.SpecialEndDate.value = Date
    Me.SpecialText = ""
    'Me.chkInfoSpecial = 0
    Dim i As Long
    For i = 0 To Me.opShowOnRebatePage.UBound
        Me.opShowOnRebatePage(i) = 0
    Next i
    For i = 0 To Me.opShowOnTab.UBound
        Me.opShowOnTab(i) = 0
    Next i
    Me.SectionImage = ""
    Me.SectionURL = ""
    Me.RebateFormURL = ""
    Me.SpecialTitle = ""
    Me.AppliedTo.Clear
    requeryAllPartNumbers
End Sub

Private Sub opPNFilter_Click(index As Integer)
    requeryAllPartNumbers
End Sub

Private Sub chkHideDiscontinued_Click()
    requeryAllPartNumbers
End Sub

Private Sub opShowOnRebatePage_Click(index As Integer)
    If Not fillingForm Then
        If Me.SpecialName = "" Then
            Me.opShowOnRebatePage(index) = 0
            Exit Sub
        End If
        
        Select Case index
            Case Is = 0 'show
                DB.Execute "UPDATE Specials SET ShowOnRebatePage=1 WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
            Case Is = 1 'do not show
                DB.Execute "UPDATE Specials SET ShowOnRebatePage=0 WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
        End Select
    End If
End Sub

Private Sub opShowOnTab_Click(index As Integer)
    If Not fillingForm Then
        If Me.SpecialName = "" Then
            Me.opShowOnTab(index) = 0
            Exit Sub
        End If
        
        Select Case index
            Case Is = 0 'specials
                DB.Execute "UPDATE Specials SET ShowOnTab=1 WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
            Case Is = 1 'info
                DB.Execute "UPDATE Specials SET ShowOnTab=2 WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
            Case Is = 2 'none
                DB.Execute "UPDATE Specials SET ShowOnTab=0 WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
            Case Is = 3 'above tabs
                DB.Execute "UPDATE Specials SET ShowOnTab=3 WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
        End Select
    End If
End Sub

Private Sub PickingWarning_KeyDown(KeyCode As Integer, Shift As Integer)
    changed = True
End Sub

Private Sub PickingWarning_LostFocus()
    If Me.SpecialName = "" Then
        Me.RebateFormURL = ""
        changed = False
        Exit Sub
    End If
    If changed Then
        DB.Execute "UPDATE Specials SET PickingWarning='" & EscapeSQuotes(Me.PickingWarning) & "' WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
        changed = False
    End If
End Sub

Private Sub RebateFormURL_KeyDown(KeyCode As Integer, Shift As Integer)
    changed = True
End Sub

Private Sub RebateFormURL_LostFocus()
    If Me.SpecialName = "" Then
        Me.RebateFormURL = ""
        changed = False
        Exit Sub
    End If
    If changed Then
        If Left(Me.RebateFormURL, 50) = "http://p1.hostingprod.com/@tools-plus.com/rebates/" Then
            Me.RebateFormURL = Mid(Me.RebateFormURL, 51)
        End If
        DB.Execute "UPDATE Specials SET RebateFormURL='" & EscapeSQuotes(Me.RebateFormURL) & "' WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
        changed = False
    End If
End Sub

Private Sub SectionImage_Click()
    changed = True
    SectionImage_LostFocus
End Sub

Private Sub SectionImage_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.SectionImage, KeyCode, Shift
    Select Case KeyCode
        Case Is = vbKeyReturn
            SectionImage_LostFocus
        Case Is = vbKeyDelete
            SectionImage_KeyPress KeyCode
    End Select
End Sub

Private Sub SectionImage_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.SectionImage, KeyAscii
    changed = True
End Sub

Private Sub SectionImage_LostFocus()
    AutoCompleteLostFocus Me.SectionImage
    If Me.SpecialName = "" Then
        Me.SectionImage = ""
        changed = False
        Exit Sub
    End If
    If changed = True Then
        DB.Execute "UPDATE Specials SET SectionImage='" & EscapeSQuotes(Me.SectionImage) & "' WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
        changed = False
    End If
End Sub

Private Sub SectionURL_Click()
    changed = True
    SectionURL_LostFocus
End Sub

Private Sub SectionURL_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.SectionURL, KeyCode, Shift
    Select Case KeyCode
        Case Is = vbKeyReturn
            SectionURL_LostFocus
        Case Is = vbKeyDelete
            SectionURL_KeyPress KeyCode
    End Select
End Sub

Private Sub SectionURL_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.SectionURL, KeyAscii
    changed = True
End Sub

Private Sub SectionURL_LostFocus()
    AutoCompleteLostFocus Me.SectionURL
    If Me.SpecialName = "" Then
        Me.SectionURL = ""
        changed = False
        Exit Sub
    End If
    If changed = True Then
        DB.Execute "UPDATE Specials SET SectionURL='" & EscapeSQuotes(Me.SectionURL) & "' WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
        changed = False
    End If
End Sub

Private Sub SpecialPostmarkDate_LostFocus()
    If Me.SpecialName = "" Then
        Exit Sub
    End If
    fixPostmarkDate
    DB.Execute "UPDATE Specials SET SpecialPostmarkDate='" & Me.SpecialPostmarkDate.value & "' WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
End Sub

Private Sub SpecialStartDate_LostFocus()
    If Me.SpecialName = "" Then
        Exit Sub
    End If
    DB.Execute "UPDATE Specials SET SpecialStartDate='" & Me.SpecialStartDate.value & "' WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
End Sub

Private Sub SpecialEndDate_LostFocus()
    If Me.SpecialName = "" Then
        Exit Sub
    End If
    DB.Execute "UPDATE Specials SET SpecialEndDate='" & Me.SpecialEndDate.value & "' WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
End Sub

Private Sub SpecialName_Click()
    changed = True
    SpecialName_LostFocus
End Sub

Private Sub SpecialName_GotFocus()
    'If itemAssocChanged = True Then
    '    If MsgBox("Do you want to save item associations for this special first?", vbYesNo) = vbYes Then
    '        saveItemAssoc
    '    End If
    '    itemAssocChanged = False
    'End If
    
    If Me.SpecialName <> "" Then
        'If Not OptionGroupHasSelection(Me.opShowOnRebatePage) Then
        '    MsgBox "You need to select whether this should show on the rebates and specials page!"
        '    'Me.opShowOnRebatePage(0).SetFocus
        '    Me.SpecialText.SetFocus
        '    Exit Sub
        'End If
        'If Not OptionGroupHasSelection(Me.opShowOnTab) Then
        '    MsgBox "You need to select which tab this will show up on!"
        '    'Me.opShowOnTab(0).SetFocus
        '    Me.SpecialText.SetFocus
        '    Exit Sub
        'End If
    End If
    
    fixPostmarkDate
End Sub

Private Sub SpecialName_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.SpecialName, KeyCode, Shift
End Sub

Private Sub SpecialName_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.SpecialName, KeyAscii
    If KeyAscii = vbKeyReturn Then
        SpecialName_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub SpecialName_LostFocus()
    AutoCompleteLostFocus Me.SpecialName
    If changed = True Then
        If Me.SpecialName <> "" Then
            'CheckFeedPromos
            If InCombo(Me.SpecialName.Text, Me.SpecialName, True) Then
                CheckFeedPromos
                fillForm
            Else
                Me.SpecialName = ""
            End If
        Else
            clearForm
        End If
        changed = False
    End If
End Sub

'Private Sub saveItemAssoc()
'    Dim i As Long, strsql As String
'    strsql = "DELETE FROM PartNumberSpecials WHERE SpecialName='" & Me.SpecialName & "'"
    'If Me.AppliedTo.ListCount > 0 Then
    '    strsql = strsql & " AND ItemNumber NOT IN ( "
    '    For i = 0 To Me.AppliedTo.ListCount - 1
    '        strsql = strsql & "'" & Me.AppliedTo.list(i) & "', "
    '    Next i
    '    strsql = Left(strsql, Len(strsql) - 2) & " )"
    'End If
'    DB.Execute strsql
    
'    Dim rst As ADODB.Recordset
'    For i = 0 To Me.AppliedTo.ListCount - 1
        'Set rst = retrieve("SELECT ID FROM PartNumberSpecials WHERE ItemNumber='" & Me.AppliedTo.list(i) & "' AND SpecialName='" & escapeSQuotes(Me.SpecialName) & "'")
        'If rst.EOF Then
'            DB.Execute "INSERT INTO PartNumberSpecials ( ItemNumber, SpecialName ) VALUES ( '" & Me.AppliedTo.list(i) & "', '" & EscapeSQuotes(Me.SpecialName) & "' )"
        'End If
        'rst.Close
'    Next i
'    Set rst = Nothing
'    Dim i As Long
'    Dim sql As String
'    sql = "DELETE FROM PartNumberSpecials WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
'    If Me.AppliedTo.ListCount > 0 Then
'        sql = sql & " AND ItemNumber NOT IN ( '"
'        For i = 0 To Me.AppliedTo.ListCount - 1
'            sql = sql & Me.AppliedTo.GetText(i, 0) & "', '"
'        Next i
'        sql = Left(sql, Len(sql) - 3) & " )"
'    End
'    DB.Execute sql
'End Sub

Private Sub Notes_KeyDown(KeyCode As Integer, Shift As Integer)
    changed = True
End Sub

Private Sub Notes_LostFocus()
    If changed Then
        DB.Execute "UPDATE Specials SET InternalNotes='" & EscapeSQuotes(Me.Notes) & "' WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
        changed = False
    End If
End Sub

Private Sub SpecialText_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.SpecialText.selStart = 0
                Me.SpecialText.SelLength = Len(Me.SpecialText)
                KeyCode = 0
                Shift = 0
            End If
        Case Is = vbKeyReturn
            SpecialText_LostFocus
        Case Is = vbKeyDelete
            changed = True
        Case Else
            changed = True
    End Select
End Sub

Private Sub SpecialText_LostFocus()
    If Me.SpecialName = "" Then
        Me.SpecialText = ""
        changed = False
        Exit Sub
    End If
    If changed = True Then
        Me.SpecialText = UnFrontpage(Me.SpecialText)
        DB.Execute "UPDATE Specials SET SpecialText='" & EscapeSQuotes(Me.SpecialText) & "' WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
        changed = False
    End If
End Sub



Private Sub requerySectionURL()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT URL FROM SectionURL")
    Me.SectionURL.Clear
    While Not rst.EOF
        Me.SectionURL.AddItem (rst("URL"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requerySectionImage()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ImageName FROM SectionImage")
    Me.SectionImage.Clear
    While Not rst.EOF
        Me.SectionImage.AddItem (rst("ImageName"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub requeryAllPartNumbers()
    Dim sql As String
    Dim clc As String, altlc As String
    clc = Left(SignMaintenance.ItemNumber, 3)
    'Select Case True
    '    Case Is = Me.opPNFilter(PNF_ALL)
    '        sql = "SELECT ItemNumber FROM PartNumbers WHERE 1=1"
    '    Case Is = Me.opPNFilter(PNF_THIS)
    '        altlc = DLookup("RealLineCode", "ProductLineSub", "AlternateLineCode='" & clc & "'")
    '        If altlc = "" Then
    '            altlc = clc
    '        End If
    '        sql = "SELECT ItemNumber FROM PartNumbers WHERE (LEFT(ItemNumber,3)='" & altlc & "' OR LEFT(ItemNumber,3) IN (SELECT AlternateLineCode FROM ProductLineSub WHERE RealLineCode='" & altlc & "'))"
    '    Case Is = Me.opPNFilter(PNF_REL)
    '        altlc = DLookup("RealLineCode", "ProductLineSubForPOs", "AlternateLineCode='" & clc & "'")
    '        If altlc = "" Then
    '            altlc = clc
    '        End If
    '        sql = "SELECT ItemNumber FROM PartNumbers WHERE (LEFT(ItemNumber,3)='" & altlc & "' OR LEFT(ItemNumber,3) IN (SELECT AlternateLineCode FROM ProductLineSubForPOs WHERE RealLineCode='" & altlc & "'))"
    'End Select
    'sql = sql & IIf(Me.chkHideDeleted, " AND WebDeleted=0", "")
    sql = "SELECT PartNumbers.ItemNumber FROM PartNumbers"
    If Me.chkHideDiscontinued Then
        sql = sql & " INNER JOIN vItemStatusBreakdown ON PartNumbers.ItemNumber=vItemStatusBreakdown.ItemNumber"
    End If
    sql = sql & " WHERE WebDeleted=0"
    altlc = DLookup("RealLineCode", "ProductLineSub", "AlternateLineCode='" & clc & "'")
    If altlc = "" Then
        altlc = clc
    End If
    Select Case True
        Case Is = Me.opPNFilter(PNF_ALL)
            'nothing
        Case Is = Me.opPNFilter(PNF_THIS)
            sql = sql & " AND (LEFT(PartNumbers.ItemNumber,3)='" & altlc & "' OR LEFT(PartNumbers.ItemNumber,3) IN (SELECT AlternateLineCode FROM ProductLineSub WHERE RealLineCode='" & altlc & "'))"
        Case Is = Me.opPNFilter(PNF_REL)
            sql = sql & " AND (LEFT(PartNumbers.ItemNumber,3)='" & altlc & "' OR LEFT(PartNumbers.ItemNumber,3) IN (SELECT AlternateLineCode FROM ProductLineSubForPOs WHERE RealLineCode='" & altlc & "'))"
    End Select
    If Me.chkHideDiscontinued Then
        sql = sql & " AND vItemStatusBreakdown.DConWeb=0"
    End If
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve(sql)
    Me.AllPartNumbers.Clear
    While Not rst.EOF
    
        If Me.AppliedTo.InList(rst("ItemNumber")) Then
            'skip
            'Debug.Print "skip"
        Else
            Me.AllPartNumbers.AddItem (rst("ItemNumber"))
        End If
        
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryAppliedTo()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber, '$' + CONVERT(varchar(12),RebateAmount,1) AS RebateAmount, CASE WHEN InstantRebate=1 THEN 'Y' ELSE 'N' END AS InstantRebate, CASE WHEN AvailQuantity IS NULL THEN '' ELSE CAST(AvailQuantity AS varchar(4)) END FROM PartNumberSpecials WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'")
    Me.AppliedTo.Clear
    'While Not rst.EOF
    '    Me.AppliedTo.AddItem (rst("ItemNumber"))
        Me.AppliedTo.Add rst, -1, True
    '    rst.MoveNext
    'Wend
    rst.Close
    Set rst = Nothing
    
    requeryAllPartNumbers
End Sub

Private Sub fixPostmarkDate()
    If DateDiff("d", Me.SpecialEndDate.value, Me.SpecialPostmarkDate.value) < 0 Then
        Me.SpecialPostmarkDate.value = Me.SpecialEndDate.value
    End If
End Sub

Private Sub SpecialTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    changed = True
End Sub

Private Sub SpecialTitle_LostFocus()
    If changed Then
        DB.Execute "UPDATE Specials SET SpecialTitle='" & EscapeSQuotes(Me.SpecialTitle) & "' WHERE SpecialName='" & EscapeSQuotes(Me.SpecialName) & "'"
        changed = False
    End If
End Sub
