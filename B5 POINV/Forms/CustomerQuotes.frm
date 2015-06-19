VERSION 5.00
Begin VB.Form CustomerQuotes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Quotes"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton FirstQuote 
      Caption         =   "First"
      Height          =   435
      Left            =   3090
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5460
      Width           =   765
   End
   Begin VB.CommandButton LastQuote 
      Caption         =   "Last"
      Height          =   435
      Left            =   5430
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5460
      Width           =   765
   End
   Begin VB.CommandButton MakeYahooFile 
      Caption         =   "Make quotes.js"
      Height          =   405
      Left            =   3960
      TabIndex        =   24
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   23
      Top             =   0
      Width           =   1155
   End
   Begin VB.TextBox LastQuoteID 
      Height          =   285
      Left            =   3330
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "LastID"
      Top             =   30
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CheckBox NewQuote 
      Caption         =   "New"
      Height          =   435
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5460
      Width           =   795
   End
   Begin VB.Frame MiscFrame 
      Caption         =   "Other Information:"
      Height          =   765
      Left            =   150
      TabIndex        =   20
      Top             =   4380
      Width           =   5865
      Begin VB.ComboBox RunTo 
         Height          =   315
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   300
         Width           =   975
      End
      Begin VB.ComboBox RunFrom 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   300
         Width           =   975
      End
      Begin VB.Label generalLabel 
         Caption         =   "to:"
         Height          =   195
         Index           =   6
         Left            =   2070
         TabIndex        =   27
         Top             =   360
         Width           =   315
      End
      Begin VB.Label generalLabel 
         Caption         =   "Use From:"
         Height          =   225
         Index           =   5
         Left            =   240
         TabIndex        =   25
         Top             =   330
         Width           =   795
      End
   End
   Begin VB.CommandButton DeleteQuote 
      Caption         =   "Delete"
      Height          =   435
      Left            =   1800
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5460
      Width           =   795
   End
   Begin VB.CommandButton SaveQuote 
      Caption         =   "Save"
      Height          =   435
      Left            =   990
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5460
      Width           =   795
   End
   Begin VB.CommandButton NextQuote 
      Caption         =   "Next"
      Height          =   435
      Left            =   4650
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5460
      Width           =   765
   End
   Begin VB.CommandButton PreviousQuote 
      Caption         =   "Previous"
      Height          =   435
      Left            =   3870
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5460
      Width           =   765
   End
   Begin VB.TextBox QuoteID 
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "ID"
      Top             =   30
      Width           =   765
   End
   Begin VB.Frame QuoteFrame 
      Caption         =   "Customer's Message:"
      Height          =   1935
      Left            =   150
      TabIndex        =   13
      Top             =   2340
      Width           =   5865
      Begin VB.TextBox Quote 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   240
         Width           =   5595
      End
   End
   Begin VB.Frame OrderInfoFrame 
      Caption         =   "Order Information:"
      Height          =   1725
      Left            =   3480
      TabIndex        =   10
      Top             =   510
      Width           =   2535
      Begin VB.ListBox ItemList 
         Height          =   1035
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   630
         Width           =   2205
      End
      Begin VB.ComboBox ItemSelector 
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   270
         Width           =   2205
      End
   End
   Begin VB.Frame CustomerInfoFrame 
      Caption         =   "Customer Information:"
      Height          =   1665
      Left            =   150
      TabIndex        =   0
      Top             =   540
      Width           =   3135
      Begin VB.TextBox CustomerCountry 
         Height          =   285
         Left            =   1050
         TabIndex        =   5
         Top             =   1260
         Width           =   1935
      End
      Begin VB.TextBox CustomerState 
         Height          =   285
         Left            =   1050
         TabIndex        =   4
         Top             =   930
         Width           =   1935
      End
      Begin VB.TextBox CustomerCity 
         Height          =   285
         Left            =   1050
         TabIndex        =   3
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox CustomerFirstName 
         Height          =   285
         Left            =   1050
         TabIndex        =   2
         Top             =   270
         Width           =   1935
      End
      Begin VB.Label generalLabel 
         Caption         =   "Country:"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   8
         Top             =   1290
         Width           =   915
      End
      Begin VB.Label generalLabel 
         Caption         =   "State:"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   7
         Top             =   960
         Width           =   915
      End
      Begin VB.Label generalLabel 
         Caption         =   "City:"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Top             =   630
         Width           =   915
      End
      Begin VB.Label generalLabel 
         Caption         =   "First Name:"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.Label generalLabel 
      Caption         =   "Customer Quotes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   60
      Width           =   2175
   End
End
Attribute VB_Name = "CustomerQuotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private changed As Boolean
Private fillingForm As Boolean

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    requeryItemSelector
    requeryMonths
    Me.QuoteID = DLookup("MIN(ID)", "CustomerQuotes")
    fillForm
End Sub

Private Sub CustomerCity_Change()
    changed = True
End Sub

Private Sub CustomerCountry_Change()
    changed = True
End Sub

Private Sub CustomerFirstName_Change()
    changed = True
End Sub

Private Sub CustomerState_Change()
    changed = True
End Sub

Private Sub CustomerState_LostFocus()
    Dim full As String
    full = FullStateName(Me.CustomerState)
    If full <> "" Then
        Me.CustomerState = full
    End If
End Sub

Private Sub requeryItemSelector()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber FROM vNonDiscontinuedItems ORDER BY ItemNumber")
    Me.ItemSelector.AddItem ""
    While Not rst.EOF
        Me.ItemSelector.AddItem rst("ItemNumber")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub ItemSelector_Click()
    ItemSelector_LostFocus
End Sub

Private Sub ItemSelector_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.ItemSelector, KeyCode, Shift
    If KeyCode = vbKeyReturn Then
        ItemSelector_LostFocus
    End If
End Sub

Private Sub ItemSelector_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.ItemSelector, KeyAscii
End Sub

Private Sub ItemSelector_LostFocus()
    AutoCompleteLostFocus Me.ItemSelector
    If Me.ItemSelector <> "" Then
        If Not InList(Me.ItemSelector, Me.ItemList, True) Then
            Me.ItemList.AddItem Me.ItemSelector
            changed = True
            Me.ItemSelector = ""
        End If
    End If
    Me.ItemList.SetFocus
End Sub

Private Sub ItemList_DblClick()
    If vbYes = MsgBox("Remove this item from the list?", vbYesNo) Then
        Me.ItemList.RemoveItem Me.ItemList.ListIndex
        changed = True
    End If
End Sub

Private Sub requeryMonths()
    Me.RunFrom.AddItem "Never": Me.RunTo.AddItem "Never"
    Me.RunFrom.AddItem "Always": Me.RunTo.AddItem "Always"
    Me.RunFrom.AddItem "Jan": Me.RunTo.AddItem "Jan"
    Me.RunFrom.AddItem "Feb": Me.RunTo.AddItem "Feb"
    Me.RunFrom.AddItem "Mar": Me.RunTo.AddItem "Mar"
    Me.RunFrom.AddItem "Apr": Me.RunTo.AddItem "Apr"
    Me.RunFrom.AddItem "May": Me.RunTo.AddItem "May"
    Me.RunFrom.AddItem "Jun": Me.RunTo.AddItem "Jun"
    Me.RunFrom.AddItem "Jul": Me.RunTo.AddItem "Jul"
    Me.RunFrom.AddItem "Aug": Me.RunTo.AddItem "Aug"
    Me.RunFrom.AddItem "Sep": Me.RunTo.AddItem "Sep"
    Me.RunFrom.AddItem "Oct": Me.RunTo.AddItem "Oct"
    Me.RunFrom.AddItem "Nov": Me.RunTo.AddItem "Nov"
    Me.RunFrom.AddItem "Dec": Me.RunTo.AddItem "Dec"
End Sub

Public Function convertMonth(mo As String) As String
    If IsNumeric(mo) Then
        Select Case mo
            Case Is = "-2"
                convertMonth = "Never"
            Case Is = "-1"
                convertMonth = "Always"
            Case Is = "0"
                convertMonth = "Jan"
            Case Is = "1"
                convertMonth = "Feb"
            Case Is = "2"
                convertMonth = "Mar"
            Case Is = "3"
                convertMonth = "Apr"
            Case Is = "4"
                convertMonth = "May"
            Case Is = "5"
                convertMonth = "Jun"
            Case Is = "6"
                convertMonth = "Jul"
            Case Is = "7"
                convertMonth = "Aug"
            Case Is = "8"
                convertMonth = "Sep"
            Case Is = "9"
                convertMonth = "Oct"
            Case Is = "10"
                convertMonth = "Nov"
            Case Is = "11"
                convertMonth = "Dec"
        End Select
    Else
        Select Case mo
            Case Is = "Never"
                convertMonth = "-2"
            Case Is = "Always"
                convertMonth = "-1"
            Case Is = "Jan"
                convertMonth = "0"
            Case Is = "Feb"
                convertMonth = "1"
            Case Is = "Mar"
                convertMonth = "2"
            Case Is = "Apr"
                convertMonth = "3"
            Case Is = "May"
                convertMonth = "4"
            Case Is = "Jun"
                convertMonth = "5"
            Case Is = "Jul"
                convertMonth = "6"
            Case Is = "Aug"
                convertMonth = "7"
            Case Is = "Sep"
                convertMonth = "8"
            Case Is = "Oct"
                convertMonth = "9"
            Case Is = "Nov"
                convertMonth = "10"
            Case Is = "Dec"
                convertMonth = "11"
        End Select
    End If
End Function



Private Sub MakeYahooFile_Click()
    Dim fname As String
    fname = Environ("UserProfile") & "\Desktop\quotes.js"
    Shell PERL & " s:\mastest\mas90-signs\generate_quote_array.pl " & qq(fname), vbNormalFocus
    MsgBox "Upload the file:" & vbCrLf & vbCrLf & qq(fname) & vbCrLf & vbCrLf & " to the hostingprod ftp."
End Sub

Private Sub Quote_Change()
    changed = True
End Sub

Private Sub NewQuote_Click()
    If Me.NewQuote Then
        Me.LastQuoteID = Me.QuoteID
        Me.QuoteID = "NEW"
        Me.DeleteQuote.Enabled = False
        Me.PreviousQuote.Enabled = False
        'Me.Search.Enabled = False
        Me.NextQuote.Enabled = False
    Else
        Me.QuoteID = Me.LastQuoteID
        Me.DeleteQuote.Enabled = True
        Me.PreviousQuote.Enabled = True
        'Me.Search.Enabled = True
        Me.NextQuote.Enabled = True
    End If
    fillForm
End Sub

Private Sub RunFrom_Click()
    If Not fillingForm Then
        fillingForm = True
        If Me.RunFrom = "Never" Then
            Me.RunTo = "Never"
        ElseIf Me.RunFrom = "Always" Then
            Me.RunTo = "Always"
        End If
        fillingForm = True
        changed = True
    End If
End Sub

Private Sub RunTo_Click()
    If Not fillingForm Then
        fillingForm = True
        If Me.RunTo = "Never" Then
            Me.RunFrom = "Never"
        ElseIf Me.RunTo = "Always" Then
            Me.RunFrom = "Always"
        End If
        fillingForm = False
        changed = True
    End If
End Sub

Private Sub SaveQuote_Click()
    If Not changed Then
        MsgBox "You haven't made any changes, what am I supposed to save?"
        Exit Sub
    ElseIf Me.CustomerFirstName = "" Then
        MsgBox "No customer first name given!"
        Exit Sub
    ElseIf CBool(InStr(Me.CustomerFirstName, " ")) Then
        If vbNo = MsgBox("Space found in name, this better not include a last name!!!" & vbCrLf & vbCrLf & "Are you sure you want to continue?", vbYesNo) Then
            Exit Sub
        End If
    ElseIf Me.CustomerCity = "" Then
        MsgBox "No customer city given!"
        Exit Sub
    ElseIf Me.CustomerState = "" Then
        'ok
    ElseIf Me.CustomerCountry = "" Then
        'ok
    ElseIf Me.ItemList.ListCount = 0 Then
        'ok
    ElseIf Me.Quote = "" Then
        MsgBox "No customer message given!"
        Exit Sub
    ElseIf (Me.RunFrom = "Never" And Me.RunTo <> "Never") Or (Me.RunFrom = "Always" And Me.RunTo <> "Always") Then
        MsgBox "From month specified, but no To month specified!"
        Exit Sub
    ElseIf (Me.RunTo = "Never" And Me.RunFrom <> "Never") Or (Me.RunTo = "Always" And Me.RunFrom <> "Always") Then
        MsgBox "To month specified, but no From month specified!"
        Exit Sub
    End If
    
    If Me.CustomerCountry = "USA" Or Me.CustomerCountry = "United States" Or Me.CustomerCountry = "" Then
        Me.CustomerCountry = "US"
    End If
    
    'If vbYes = MsgBox("Save this quote now?", vbYesNo) Then
        Mouse.Hourglass True
        Dim newid As String
        If Me.NewQuote Then
            newid = insertRecord
            updateItems newid
        Else
            updateRecord Me.QuoteID
            updateItems Me.QuoteID
        End If
        Mouse.Hourglass False
    'End If
    If Me.NewQuote Then
        Me.LastQuoteID = newid
        Me.NewQuote = 0
    End If
    changed = False
End Sub

Private Sub DeleteQuote_Click()
    If vbYes = MsgBox("Delete this quote?", vbYesNo) Then
        If Me.NewQuote Then
            Me.NewQuote = 1
        Else
            DB.Execute "DELETE FROM CustomerQuotes WHERE ID=" & Me.QuoteID
        End If
    End If
End Sub

Private Sub FirstQuote_Click()
    If Not Me.NewQuote Then
        If changed Then
            If vbNo = MsgBox("Discard changes?", vbYesNo) Then
                Exit Sub
            End If
            changed = False
        End If
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT TOP 1 ID FROM CustomerQuotes ORDER BY ID")
        If rst.EOF Then
            'that's funny
        Else
            Me.LastQuoteID = Me.QuoteID
            Me.QuoteID = rst("ID")
            fillForm
        End If
    End If
End Sub

Private Sub PreviousQuote_Click()
    If Not Me.NewQuote Then
        If changed Then
            If vbNo = MsgBox("Discard changes?", vbYesNo) Then
                Exit Sub
            End If
            changed = False
        End If
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT TOP 1 ID FROM CustomerQuotes WHERE ID<" & Me.QuoteID & " ORDER BY ID DESC")
        If rst.EOF Then
            'first one, do nothing
        Else
            Me.LastQuoteID = Me.QuoteID
            Me.QuoteID = rst("ID")
            fillForm
        End If
    End If
End Sub

Private Sub NextQuote_Click()
    If Not Me.NewQuote Then
        If changed Then
            If vbNo = MsgBox("Discard changes?", vbYesNo) Then
                Exit Sub
            End If
            changed = False
        End If
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT TOP 1 ID FROM CustomerQuotes WHERE ID>" & Me.QuoteID & " ORDER BY ID")
        If rst.EOF Then
            'last one, do nothing
        Else
            Me.LastQuoteID = Me.QuoteID
            Me.QuoteID = rst("ID")
            fillForm
        End If
    End If
End Sub

Private Sub LastQuote_Click()
    If Not Me.NewQuote Then
        If changed Then
            If vbNo = MsgBox("Discard changes?", vbYesNo) Then
                Exit Sub
            End If
            changed = False
        End If
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT TOP 1 ID FROM CustomerQuotes ORDER BY ID DESC")
        If rst.EOF Then
            'that's funny
        Else
            Me.LastQuoteID = Me.QuoteID
            Me.QuoteID = rst("ID")
            fillForm
        End If
    End If
End Sub

Private Sub fillForm()
    fillingForm = True
    If Me.QuoteID = "" Then
        Me.NewQuote = 1
    ElseIf Me.QuoteID = "NEW" Then
        Me.CustomerFirstName = ""
        Me.CustomerCity = ""
        Me.CustomerState = ""
        Me.CustomerCountry = ""
        Me.ItemList.Clear
        Me.Quote = ""
        Me.RunFrom = "Always"
        Me.RunTo = "Always"
    Else
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT * FROM CustomerQuotes WHERE ID=" & Me.QuoteID)
        Me.CustomerFirstName = rst("FirstName")
        Me.CustomerCity = rst("City")
        Me.CustomerState = Nz(rst("State"))
        Me.CustomerCountry = Nz(rst("Country"))
        Me.RunFrom = convertMonth(rst("RunFrom"))
        Me.RunTo = convertMonth(rst("RunTo"))
        
        Me.Quote = rst("Quote")
        rst.Close
        Set rst = DB.retrieve("SELECT ItemNumber FROM CustomerQuoteItems WHERE QuoteID=" & Me.QuoteID)
        Me.ItemList.Clear
        While Not rst.EOF
            Me.ItemList.AddItem rst("ItemNumber")
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
        changed = False
    End If
    fillingForm = False
End Sub

Private Function insertRecord(Optional newid As String = "") As String
    'insert, maybe with the id. if not given an ID, then get the autonumbered ID
    Dim sql As String
    sql = "INSERT INTO CustomerQuotes (" & IIf(newid = "", "", "ID, ") & _
                                        "FirstName, " & _
                                        "City, " & _
                                        IIf(Me.CustomerState = "", "", "State, ") & _
                                        IIf(Me.CustomerCountry = "", "", "Country, ") & _
                                        "Quote, " & _
                                        "RunFrom, RunTo " & _
                                     ") VALUES ( " & _
                                        IIf(newid = "", "", newid & ", ") & _
                                        "'" & EscapeSQuotes(Me.CustomerFirstName) & "', " & _
                                        "'" & EscapeSQuotes(Me.CustomerCity) & "', " & _
                                        IIf(Me.CustomerState = "", "", "'" & EscapeSQuotes(Me.CustomerState) & "', ") & _
                                        IIf(Me.CustomerCountry = "", "", "'" & EscapeSQuotes(Me.CustomerCountry) & "', ") & _
                                        "'" & EscapeSQuotes(Me.Quote) & "', " & _
                                        convertMonth(Me.RunFrom) & ", " & convertMonth(Me.RunTo) & _
                                    " )"
    DB.Execute sql

    If newid = "" Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT @@IDENTITY")
        newid = rst(0)
        rst.Close
        Set rst = Nothing
    End If
    insertRecord = newid
End Function

Private Sub updateRecord(QID As String)
    Dim sql As String
    sql = "UPDATE CustomerQuotes " & _
          "SET FirstName='" & EscapeSQuotes(Me.CustomerFirstName) & "', " & _
              "City='" & EscapeSQuotes(Me.CustomerCity) & "', " & _
              "State=" & IIf(Me.CustomerState = "", "NULL", "'" & EscapeSQuotes(Me.CustomerState) & "'") & ", " & _
              "Country=" & IIf(Me.CustomerCountry = "", "NULL", "'" & EscapeSQuotes(Me.CustomerCountry) & "'") & ", " & _
              "Quote='" & EscapeSQuotes(Me.Quote) & "', " & _
              "RunFrom=" & convertMonth(Me.RunFrom) & ", " & _
              "RunTo=" & convertMonth(Me.RunTo) & _
         " WHERE ID=" & QID
    DB.Execute sql
End Sub

Private Sub updateItems(QID As String)
    Dim itemH As Dictionary
    Set itemH = New Dictionary
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber FROM CustomerQuoteItems WHERE QuoteID=" & QID)
    While Not rst.EOF
        itemH.Add CStr(rst("ItemNumber")), "1"
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    Dim i As Integer
    For i = 0 To Me.ItemList.ListCount - 1
        If itemH.exists(CStr(Me.ItemList.list(i))) Then
            itemH.Remove CStr(Me.ItemList.list(i))
        Else
            DB.Execute "INSERT INTO CustomerQuoteItems ( QuoteID, ItemNumber ) VALUES ( " & QID & ", '" & Me.ItemList.list(i) & "' )"
        End If
    Next i
    
    For i = 0 To UBound(itemH.keys)
        DB.Execute "DELETE FROM CustomerQuoteItems WHERE QuoteID=" & QID & " AND ItemNumber='" & CStr(itemH.keys(i)) & "'"
    Next i
End Sub
