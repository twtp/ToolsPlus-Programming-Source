VERSION 5.00
Begin VB.Form GeneralSpreadsheetImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Spreadsheet Import"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6195
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkWarnInvalid 
      Caption         =   "warn on invalid itemnumber"
      Height          =   585
      Left            =   5010
      TabIndex        =   21
      Top             =   1530
      Width           =   1155
   End
   Begin VB.CommandButton btnGo 
      Caption         =   "Import"
      Height          =   435
      Left            =   1860
      TabIndex        =   20
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton btnGoTest 
      Caption         =   "Test"
      Height          =   435
      Left            =   420
      TabIndex        =   19
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   5040
      TabIndex        =   18
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton btnRemove 
      Caption         =   "Remove"
      Height          =   255
      Left            =   5310
      TabIndex        =   17
      Top             =   2730
      Width           =   735
   End
   Begin VB.ListBox ColumnList 
      Height          =   645
      ItemData        =   "GeneralSpreadsheetImport.frx":0000
      Left            =   3270
      List            =   "GeneralSpreadsheetImport.frx":0002
      TabIndex        =   16
      Top             =   2580
      Width           =   1995
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add"
      Height          =   285
      Left            =   2370
      TabIndex        =   15
      Top             =   2940
      Width           =   645
   End
   Begin VB.ComboBox FieldsAdd 
      Height          =   315
      Left            =   1050
      TabIndex        =   14
      Top             =   2910
      Width           =   1275
   End
   Begin VB.TextBox ColumnAdd 
      Height          =   285
      Left            =   180
      TabIndex        =   13
      Top             =   2940
      Width           =   765
   End
   Begin VB.TextBox PrependItemNumber 
      Height          =   285
      Left            =   4050
      TabIndex        =   11
      Top             =   1860
      Width           =   765
   End
   Begin VB.TextBox ColItemNumber 
      Height          =   285
      Left            =   4050
      TabIndex        =   9
      Top             =   1470
      Width           =   765
   End
   Begin VB.TextBox RowEnd 
      Height          =   285
      Left            =   1140
      TabIndex        =   7
      Top             =   1860
      Width           =   735
   End
   Begin VB.TextBox RowStart 
      Height          =   285
      Left            =   1140
      TabIndex        =   4
      Top             =   1530
      Width           =   735
   End
   Begin VB.CommandButton btnBrowseForFile 
      Caption         =   "Browse..."
      Height          =   285
      Left            =   3810
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   810
      Width           =   915
   End
   Begin VB.TextBox Filename 
      Height          =   285
      Left            =   810
      TabIndex        =   2
      Top             =   810
      Width           =   2985
   End
   Begin VB.Label lblStatusBar 
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   4230
      Width           =   4635
   End
   Begin VB.Label Label1 
      Caption         =   "Columns to Import:"
      Height          =   255
      Left            =   150
      TabIndex        =   12
      Top             =   2640
      Width           =   2865
   End
   Begin VB.Label generalLabel 
      Caption         =   "Add to beginning of ItemNumber (LC?):"
      Height          =   435
      Index           =   6
      Left            =   2460
      TabIndex        =   10
      Top             =   1830
      Width           =   1545
   End
   Begin VB.Label generalLabel 
      Caption         =   "ItemNumber column:"
      Height          =   285
      Index           =   5
      Left            =   2460
      TabIndex        =   8
      Top             =   1470
      Width           =   1545
   End
   Begin VB.Label generalLabel 
      Caption         =   "End at row:"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1860
      Width           =   975
   End
   Begin VB.Label generalLabel 
      Caption         =   "Start at row:"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1530
      Width           =   975
   End
   Begin VB.Label generalLabel 
      Caption         =   "File:"
      Height          =   285
      Index           =   1
      Left            =   420
      TabIndex        =   1
      Top             =   840
      Width           =   375
   End
   Begin VB.Label generalLabel 
      Caption         =   "General Spreadsheet Import"
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4215
   End
End
Attribute VB_Name = "GeneralSpreadsheetImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private testOnly As Boolean

Private Sub btnAdd_Click()
    If Me.ColumnAdd <> "" And Me.FieldsAdd <> "" Then
        Me.ColumnList.AddItem Me.ColumnAdd & vbTab & Me.FieldsAdd
    End If
End Sub

Private Sub btnBrowseForFile_Click()
    Dim dlg As FilePicker
    Set dlg = New FilePicker
    dlg.SetParent Me.hWnd
    dlg.AddFilter "Excel Documents", "*.xls"
    dlg.AddFilter "Comma-Separated Values Files", "*.csv"
    Me.Filename = dlg.ShowDialogOpen
    Set dlg = Nothing
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnGo_Click()
    testOnly = False
    runIt
End Sub

Private Sub btnGoTest_Click()
    testOnly = True
    runIt
End Sub

Private Sub btnRemove_Click()
    If Me.ColumnList.SelCount = 1 Then
        Me.ColumnList.RemoveItem Me.ColumnList.ListIndex
    End If
End Sub

Private Sub FieldsAdd_Click()
    FieldsAdd_LostFocus
End Sub

Private Sub FieldsAdd_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.FieldsAdd, KeyCode, Shift
End Sub

Private Sub FieldsAdd_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.FieldsAdd, KeyAscii
End Sub

Private Sub FieldsAdd_LostFocus()
    AutoCompleteLostFocus Me.FieldsAdd
End Sub

Private Sub Form_Load()
    SetListTabs Me.ColumnList, Array(18)
    requeryFields
End Sub

Private Sub requeryFields()
    Me.FieldsAdd.AddItem "Store Price"
    Me.FieldsAdd.AddItem "Web Price"
    Me.FieldsAdd.AddItem "EBay Price"
    Me.FieldsAdd.AddItem "MAPP"
    Me.FieldsAdd.AddItem "TNC"
    Me.FieldsAdd.AddItem "List Price"
    Me.FieldsAdd.AddItem "Weight"
    Me.FieldsAdd.AddItem "Std Pack"
End Sub

Private Function filledOut() As Boolean
    filledOut = False
    If Me.Filename = "" Then
        MsgBox "Need file!"
    ElseIf Me.RowStart = "" Then
        MsgBox "Need start row!"
    ElseIf Me.RowEnd = "" Then
        MsgBox "Need end row!"
    ElseIf Me.ColItemNumber = "" Then
        MsgBox "Need item number column!"
    ElseIf Me.ColumnList.ListCount < 1 Then
        MsgBox "Need data columns!"
    Else
        filledOut = True
    End If
End Function

Private Function checkAndUpdateNumeric(item As String, value As String, field As String, Optional field2 As String = "") As Boolean
    If IsNumeric(value) Then
        If testOnly Then
            MsgBox "Changing " & field & " for " & item & " to " & value
        Else
            'all get trunc to 2 dec places
            If InStr(value, ".") Then
                'value = Left(value, InStr(value, ".") + 2)
                value = Round(value, 2)
            End If
            Dim oldvalue As String
            oldvalue = DLookup(field, "InventoryMaster", "ItemNumber='" & item & "'")
            'yes, this works for integers too
            oldvalue = Left(oldvalue, InStr(value, ".") + 2)
            If value = oldvalue Then
                'Debug.Print item & " - no change"
            Else
                'Debug.Print item & " - now " & value
                Select Case field
                    Case Is = "StdPrice", "DiscountMarkupPriceRate1"
                        LogItemStorePriceChanged item, value
                    Case Is = "IDiscountMarkupPriceRate1"
                        LogItemWebPriceChanged item, value
                    Case Is = "EBayPrice"
                        LogItemEBayPriceChanged item, value
                    Case Is = "StdCost"
                        LogItemCostChanged item, value
DatabaseFunctions.ModifyItemCost GetItemIDByItemCode(item), "Std Cost", value, 3
                    Case Is = "TNC"
                        LogItemTNCChanged item, value
DatabaseFunctions.ModifyItemCost GetItemIDByItemCode(item), "New Cost", value, 3
                    Case Is = "MAPP"
                        LogItemMAPPChanged item, value, False
DatabaseFunctions.ModifyItemCost GetItemIDByItemCode(item), "MAPP", value, 3
                    Case Is = "ListPrice"
DatabaseFunctions.ModifyItemCost GetItemIDByItemCode(item), "List Price", value, 3
                End Select
                updateInventoryMaster field, value, item, ""
                If field2 <> "" Then
                    updateInventoryMaster field2, value, item, ""
                End If
                'If field = "MAPP" Then
                    'updateInventoryMaster "IgnoreMAPP", "0", item, ""
                    'Debug.Print ItemNumber & " - $" & value & ", was $" & DLookup("MAPP", "InventoryMaster", "ItemNumber='" & item & "'")
                'End If
                If field = "IDiscountMarkupPriceRate1" Then
                    updateDropshipItemWebPrice item, value
                End If
            End If
        End If
    Else
        MsgBox item & ": " & field & " has non-numeric value, not changing..."
    End If
End Function

Private Function checkAndUpdateStringish(item As String, value As String, field As String) As Boolean
    If testOnly Then
        MsgBox "Changing " & field & " for " & item & " to " & value
    Else
        Dim oldvalue As String
        oldvalue = DLookup(field, "InventoryMaster", "ItemNumber='" & item & "'")
        If value = oldvalue Then
            'Debug.Print ItemNumber & " - no change"
        Else
            'Debug.Print ItemNumber & " - now " & value
            Select Case field
                Case Is = "Weight"
                    'TODO: flag for truck freight? or question?
            End Select
            updateInventoryMaster field, EscapeSQuotes(value), item, "'"
            Select Case field
                Case Is = "Weight"
                    ExternalComponentDBSync item
            End Select
        End If
    End If
End Function

Private Sub RowEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
End Sub

Private Sub RowStart_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
End Sub

Private Sub ColItemNumber_KeyPress(KeyAscii As Integer)
    KeyAscii = ForceUppercase(LimitToLetters(KeyAscii))
End Sub

Private Sub ColumnAdd_KeyPress(KeyAscii As Integer)
    KeyAscii = ForceUppercase(LimitToLetters(KeyAscii))
End Sub


Private Sub runIt()
    If Not filledOut Then
        Exit Sub
    End If
getnum:
    Dim numToTest As String
    If testOnly Then
        numToTest = InputBox("Enter the number of lines to test:", , "10")
        If numToTest = "" Then
            Exit Sub
        ElseIf Not IsNumeric(numToTest) Then
            GoTo getnum
        ElseIf numToTest <= 0 Then
            GoTo getnum
        End If
    End If
    
    Dim excelapp As Excel.Application
    Set excelapp = New Excel.Application
    excelapp.Workbooks.Open Me.Filename
    
    Dim colField As Dictionary
    Set colField = New Dictionary
    ReDim colsA(Me.ColumnList.ListCount - 1) As String
    Dim col As String, field As String
    Dim i As Long, j As Long
    For i = 0 To Me.ColumnList.ListCount - 1
        col = Left(Me.ColumnList.list(i), InStr(Me.ColumnList.list(i), vbTab) - 1)
        field = Mid(Me.ColumnList.list(i), InStr(Me.ColumnList.list(i), vbTab) + 1)
        colField.Add col, field
        colsA(i) = col
    Next i

    Dim item As String, value As String
    Dim endline As Long
    If numToTest <> "" Then
        endline = CInt(Me.RowStart) + numToTest
    Else
        endline = CInt(Me.RowEnd)
    End If
    
    Dim warnings As PerlArray
    Set warnings = New PerlArray
    
    For i = CInt(Me.RowStart) To endline
        DoEvents
        Me.lblStatusBar.Caption = "Working on row " & i
        item = excelapp.ActiveWorkbook.ActiveSheet.Range(Me.ColItemNumber & i).value
        If item <> "" Then
            item = Me.PrependItemNumber & Trim$(item)
            item = Replace(item, "/", "-")
            item = Replace(item, "'", "")
            If Not ExistsInInventoryMaster(item) Then
                If Me.chkWarnInvalid Or testOnly Then
                    MsgBox "Couldn't find " & item & " in InventoryMaster, skipping..."
                End If
            Else
                For j = LBound(colsA) To UBound(colsA)
                    value = excelapp.ActiveWorkbook.ActiveSheet.Range(colsA(j) & i).value
                    If value <> "" Then
                        Select Case colField.item(colsA(j))
                            Case Is = "Store Price"
                                If IsItemCompletelyDC(item) Then
                                    warnings.Push (item & ": is d/c, skipping store price")
                                Else
                                    checkAndUpdateNumeric item, value, "StdPrice", "DiscountMarkupPriceRate1"
                                End If
                            Case Is = "Web Price"
                                If IsItemCompletelyDC(item) Then
                                    warnings.Push (item & ": is d/c, skipping web price")
                                Else
                                    checkAndUpdateNumeric item, value, "IDiscountMarkupPriceRate1"
                                End If
                            Case Is = "EBay Price"
                                If IsItemCompletelyDC(item) Then
                                    warnings.Push (item & ": is d/c, skipping ebay price")
                                Else
                                    checkAndUpdateNumeric item, value, "EBayPrice"
                                End If
                            Case Is = "MAPP"
                                checkAndUpdateNumeric item, value, "MAPP"
                            Case Is = "TNC"
                                If value <> DLookup("StdCost", "InventoryMaster", "ItemNumber='" & EscapeSQuotes(item) & "'") Then
                                    checkAndUpdateNumeric item, value, "TNC"
                                End If
                            Case Is = "List Price"
                                checkAndUpdateNumeric item, value, "ListPrice"
                            Case Is = "Weight"
                                checkAndUpdateStringish item, value, "Weight"
                            Case Is = "Std Pack"
                                checkAndUpdateNumeric item, value, "StdPack"
                        End Select
                    End If
                Next j
            End If
        End If
    Next i
    Me.lblStatusBar.Caption = ""
    excelapp.Quit
    Set excelapp = Nothing
    
    If warnings.Scalar > 0 Then
        MsgBox warnings.Join(vbCrLf)
    End If
End Sub
