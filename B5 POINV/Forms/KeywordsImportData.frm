VERSION 5.00
Begin VB.Form KeywordsImportData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keywords - Import Data"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9870
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnFromFile 
      Caption         =   "From TSV File"
      Height          =   375
      Left            =   180
      TabIndex        =   10
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton btnImport2 
      Caption         =   "Import #2"
      Height          =   375
      Left            =   3300
      TabIndex        =   9
      Top             =   4830
      Width           =   1395
   End
   Begin VB.TextBox hiddenDate 
      Height          =   285
      Left            =   5100
      TabIndex        =   8
      Text            =   "date"
      Top             =   330
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton btnToday 
      Caption         =   "Today is: DATE"
      Height          =   525
      Left            =   3420
      TabIndex        =   7
      Top             =   210
      Width           =   1665
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   3660
      TabIndex        =   6
      Top             =   990
      Width           =   975
   End
   Begin VB.ListBox missingData 
      Height          =   4545
      Left            =   6330
      TabIndex        =   5
      Top             =   630
      Width           =   3465
   End
   Begin VB.CommandButton btnImport 
      Caption         =   "Import this"
      Height          =   465
      Left            =   900
      TabIndex        =   4
      Top             =   4740
      Width           =   1785
   End
   Begin VB.TextBox dataField 
      Height          =   2895
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1740
      Width           =   6135
   End
   Begin VB.CommandButton btnReset 
      Caption         =   "Clear all data now"
      Height          =   525
      Left            =   570
      TabIndex        =   0
      Top             =   270
      Width           =   1875
   End
   Begin VB.Label Label3 
      Caption         =   "Paste (tab-separated) report from GA here"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   990
      Width           =   3315
   End
   Begin VB.Label Label2 
      Caption         =   "Missing data from:"
      Height          =   225
      Left            =   6360
      TabIndex        =   1
      Top             =   390
      Width           =   1785
   End
End
Attribute VB_Name = "KeywordsImportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClear_Click()
    Me.dataField = ""
End Sub

Private Sub btnFromFile_Click()
    Dim fp As FilePicker
    Set fp = New FilePicker
    fp.AddFilter "TSV Files", "*.tsv"
    Dim fname As String
    fname = fp.ShowDialogOpen
    If fname = "" Then
        Exit Sub
    End If
    
    Open fname For Input As #1
    Dim buf As String
    buf = input(LOF(1), 1)
    Close #1
    
    Me.dataField = Replace(buf, vbLf, vbCrLf)
End Sub

Private Sub btnImport_Click()
    Dim Lines As Variant
    Lines = Split(Me.dataField, vbCrLf)
    
    Dim timeframe As String
    timeframe = importGetTimeFrame(Lines)
    If timeframe = "" Then
        Exit Sub
    End If
    
    Dim service As String
    service = importGetService(Lines)
    If service = "" Then
        Exit Sub
    End If
    
    If Not importGetIsCorrectSpreadsheet(Lines) Then
        MsgBox "This isn't the correct spreadsheet, you should be on the ""e-commerce"" tab!"
        Exit Sub
    End If
    
    'MsgBox timeframe
    'MsgBox service
    
    Mouse.Hourglass True
    
    Dim engineid As String, shopeng As Boolean
    engineid = DLookup("ID", "KeywordsEngines", "Engine='" & EscapeSQuotes(service) & "'")
    shopeng = CBool(DLookup("ProductsOnly", "KeywordsEngines", "Engine='" & EscapeSQuotes(service) & "'"))
    DB.Execute "DELETE FROM KeywordsStatistics WHERE EngineID=" & engineid & " AND TimeFrame=" & timeframe
    
    Dim i As Long
    Dim atkwds As Boolean
    Dim pieces As Variant
    Dim phrase As String, visits As String, orders As String
    Dim revenue As String
    For i = 0 To UBound(Lines)
        If Not atkwds Then
            If Left(Lines(i), 8) = "Keyword" & vbTab Then
                atkwds = True
            End If
        Else
            'keywords  visits  p/visit  g1/visit  ...  t/visit  $/visit
            pieces = Split(Lines(i), vbTab)
            If UBound(pieces) > 4 Then
                phrase = CStr(pieces(0))
                visits = CStr(CLng(pieces(1)))
                'orders = CStr(CLng((CDbl(pieces(3)) / 100) * visits))
                orders = CStr(CLng(pieces(3)))
                'revenue = CStr(CDbl(pieces(UBound(pieces)) * visits))
                revenue = CStr(CDbl(pieces(2)))
                If shopeng Then
                    DB.Execute "INSERT INTO KeywordsStatistics (SearchPhrase, EngineID, TimeFrame, Visits, Orders, Revenue, Clicks, Conversions, Impressions) VALUES ( '" & EscapeSQuotes(phrase) & "', " & engineid & ", " & timeframe & ", " & visits & ", " & orders & ", " & revenue & ", " & visits & ", " & orders & ", -1 )", True
                Else
                    'TODO
                End If
            End If
        End If
    Next i
    requeryMissingData
    
    'MsgBox "done"
    
    Me.dataField = ""
    
    Mouse.Hourglass False
End Sub

Private Sub btnImport2_Click()
    Mouse.Hourglass True
    
    Dim linesA As Variant
    linesA = Split(Me.dataField, vbCrLf)
    
    Dim redate As RegExp
    Set redate = New RegExp
    redate.Pattern = "^#\s+(\d{8})-(\d{8})$"
    Dim begindate As String, enddate As String, timeframe As String
    
    Dim i As Long
    For i = 0 To UBound(linesA)
        If redate.Test(linesA(i)) Then
            Dim matches As MatchCollection
            Set matches = redate.Execute(linesA(i))
            begindate = matches(0).SubMatches(0)
            enddate = matches(0).SubMatches(1)
            begindate = Format(Left(begindate, 4) & "-" & Mid(begindate, 5, 2) & "-" & Right(begindate, 2), "m-d-yy")
            enddate = Format(Left(enddate, 4) & "-" & Mid(enddate, 5, 2) & "-" & Right(enddate, 2), "m-d-yy")
            timeframe = DateDiff("d", begindate, enddate) + 1
            If (timeframe = 10 Or timeframe = 30 Or timeframe = 60 Or timeframe = 90) And enddate = Format(Date - 1, "m-d-yy") Then
                'ok, good timeframe
                Exit For
            Else
                MsgBox "This isn't one of the needed time frames!"
                i = UBound(linesA) + 1
            End If
        ElseIf Left(linesA(i), 14) = "Source/Medium" & vbTab Then
            'can't find date line
            MsgBox "Can't find date line!"
            i = UBound(linesA) + 1
        Else
            'next
        End If
    Next i
    
    Dim reheader As RegExp
    Set reheader = New RegExp
    reheader.Pattern = "^Source(/Medium)?\s+Keyword\s+Sessions\s+Revenue\s+Transactions"
    For i = i + 1 To UBound(linesA)
        If reheader.Test(linesA(i)) Then
            Exit For
        ElseIf Left(linesA(i), 8) = "Source/Medium" & vbTab Then
            'ok, forget it
            MsgBox "This is the wrong spreadsheet, click the e-commerce tab!"
            i = UBound(linesA) + 1
        Else
            'next
        End If
    Next i
    
    Dim engineIDMap As Dictionary
    Set engineIDMap = New Dictionary
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Engine, ID FROM KeywordsEngines")
    While Not rst.EOF
        engineIDMap.Add LCase(CStr(rst("Engine"))), CStr(rst("ID"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    Dim source As String, phrase As String, visits As String, orders As String, revenue As String
    For i = i + 1 To UBound(linesA)
        'source/medium, keywords, visits, revenue, transactions (, average value, ecommerce conversion rate, per visit value)
        Dim pieces As Variant
        pieces = Split(linesA(i), vbTab)
        If UBound(pieces) >= 4 Then
            source = CStr(pieces(0))
            source = Replace(LCase(source), " / shopeng", "")
            source = Replace(LCase(source), " / feed", "")
            phrase = CStr(pieces(1))
            visits = CStr(CLng(Replace(Replace(CStr(pieces(2)), ",", ""), """", "")))
            orders = CStr(CLng(Replace(Replace(CStr(pieces(4)), ",", ""), """", "")))
            revenue = CStr(CDbl(Replace(Replace(CStr(pieces(3)), ",", ""), """", "")))
            If engineIDMap.exists(source) Then
                If Len(phrase) > 50 Then
                    'weird, usually a ? or # url that doesn't get truncated properly.
                    'this hasn't happened in a while though, it'd be nice to truncate it
                    'ourselves and add it. except that would probably require an update
                    'statement too.
                    MsgBox "Skipping long phrase " & qq(phrase)
                Else
                    DB.Execute "INSERT INTO KeywordsStatistics (SearchPhrase, EngineID, TimeFrame, Visits, Orders, Revenue, Clicks, Conversions, Impressions) VALUES ( '" & EscapeSQuotes(phrase) & "', " & engineIDMap.item(source) & ", " & timeframe & ", " & visits & ", " & orders & ", " & revenue & ", " & visits & ", " & orders & ", -1 )", True
                End If
            Else
                'TODO
            End If
        End If
    Next i
    
    requeryMissingData
    
    'MsgBox "done"
    
    Me.dataField = ""
    
    Mouse.Hourglass False
End Sub

Private Sub btnReset_Click()
    If MsgBox("Clear data?", vbYesNo) = vbYes Then
        DB.Execute "TRUNCATE TABLE KeywordsStatistics"
        requeryMissingData
    End If
End Sub

Private Sub btnToday_Click()
    Load GenericCalendarPopUp
    GenericCalendarPopUp.SetUpCalendar Me.hiddenDate, "KeywordsImportData", "hiddenDate"
    GenericCalendarPopUp.Show MODAL
    Me.btnToday.Caption = "Today is: " & Me.hiddenDate
    requeryMissingData
End Sub

Private Sub Form_Load()
    Me.hiddenDate = Date
    Me.btnToday.Caption = "Today is: " & Date
    Dim tabs(1) As Long
    tabs(0) = 72
    tabs(1) = tabs(0) + 36
    SetListTabStops Me.missingData.hWnd, tabs
    requeryMissingData
End Sub

Private Sub requeryMissingData()
    Me.missingData.Clear
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Engine FROM KeywordsEngines WHERE Active=1 ORDER BY ID")
    While Not rst.EOF
        Me.missingData.AddItem rst("Engine") & vbTab & Format(DateAdd("d", -10, Me.hiddenDate), "m-d-yy") & vbTab & Format(DateAdd("d", -1, Me.hiddenDate), "m-d-yy")
        Me.missingData.AddItem rst("Engine") & vbTab & Format(DateAdd("d", -30, Me.hiddenDate), "m-d-yy") & vbTab & Format(DateAdd("d", -1, Me.hiddenDate), "m-d-yy")
        Me.missingData.AddItem rst("Engine") & vbTab & Format(DateAdd("d", -60, Me.hiddenDate), "m-d-yy") & vbTab & Format(DateAdd("d", -1, Me.hiddenDate), "m-d-yy")
        Me.missingData.AddItem rst("Engine") & vbTab & Format(DateAdd("d", -90, Me.hiddenDate), "m-d-yy") & vbTab & Format(DateAdd("d", -1, Me.hiddenDate), "m-d-yy")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = DB.retrieve("SELECT KeywordsEngines.Engine, KeywordsStatistics.TimeFrame, COUNT(*) FROM KeywordsStatistics INNER JOIN KeywordsEngines ON KeywordsStatistics.EngineID=KeywordsEngines.ID GROUP BY KeywordsEngines.Engine, KeywordsStatistics.TimeFrame")
    Dim i As Long
    While Not rst.EOF
        For i = 0 To Me.missingData.ListCount - 1
            Dim pieces As Variant
            pieces = Split(Me.missingData.list(i), vbTab)
            If rst("Engine") = pieces(0) Then
                If rst("TimeFrame") = DateDiff("d", pieces(1), Me.hiddenDate) Then
                    Me.missingData.RemoveItem i
                    i = Me.missingData.ListCount + 1
                End If
            End If
        Next i
        rst.MoveNext
    Wend
    rst.Close
End Sub

Private Function importGetTimeFrame(Lines As Variant) As String
    Dim redate As RegExp
    Set redate = New RegExp
    redate.Pattern = "(\w+\s\d{1,2},\s\d{4})\s+(\w+\s\d{1,2},\s\d{4})"
    
    Dim okDate As Boolean
    Dim matches As MatchCollection
    Dim begindate As String, enddate As String, timeframe As String
    Dim i As Long
    For i = 0 To UBound(Lines)
        'If Left(Lines(i), 13) = "# Date Range:" Then
            'ok, check what range it is
        If redate.Test(Lines(i)) Then
            Set matches = redate.Execute(Lines(i))
            begindate = matches(0).SubMatches(0)
            enddate = matches(0).SubMatches(1)
            begindate = Format(begindate, "m-d-yy")
            enddate = Format(enddate, "m-d-yy")
            timeframe = DateDiff("d", begindate, enddate) + 1
            If (timeframe = 10 Or timeframe = 30 Or timeframe = 60 Or timeframe = 90) And enddate = Format(Date - 1, "m-d-yy") Then
                'ok, good timeframe
                okDate = True
            Else
                MsgBox "This isn't one of the needed time frames!"
            End If
            'Else
            '    MsgBox "Invalid date from GA? WTF?"
            'End If
            i = UBound(Lines) + 1
        ElseIf Left(Lines(i), 8) = "Keyword" & vbTab Then
            'ok, forget it
            i = UBound(Lines) + 1
        Else
            'next
        End If
    Next i
    If okDate Then
        importGetTimeFrame = timeframe
    Else
        MsgBox "Could not find date range from GA! This sucks!"
        importGetTimeFrame = ""
    End If
    Set redate = Nothing
End Function

Private Function importGetService(Lines As Variant) As String
    Dim resvc As RegExp
    Set resvc = New RegExp
    resvc.Pattern = "Source\sMedium\sDetail:\s+(\w+)\s/\s(?:cpc|shopeng)"
    
    Dim oksvc As Boolean
    Dim matches As MatchCollection
    Dim svc As String
    Dim i As Long
    For i = 0 To UBound(Lines)
        If resvc.Test(Lines(i)) Then
            Set matches = resvc.Execute(Lines(i))
            svc = matches(0).SubMatches(0)
            oksvc = True
            i = UBound(Lines) + 1
        ElseIf Left(Lines(i), 8) = "Keyword" & vbTab Then
            'ok, forget it
            i = UBound(Lines) + 1
        Else
            'next
        End If
    Next i
    If oksvc Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT COUNT(*) FROM KeywordsEngines WHERE Engine='" & EscapeSQuotes(svc) & "'")
        If rst(0) = 0 Then
            MsgBox "Found service " & qq(svc) & ", but it's not in the database, add it there first!"
            importGetService = ""
        Else
            importGetService = svc
        End If
        rst.Close
        Set rst = Nothing
    Else
        MsgBox "Could not find the service!" & vbCrLf & vbCrLf & "You probably didn't do the right report, read the directions again."
        importGetService = ""
    End If
    Set resvc = Nothing
End Function

Private Function importGetIsCorrectSpreadsheet(Lines As Variant) As Boolean
    Dim reheader As RegExp
    Set reheader = New RegExp
    reheader.Pattern = "^Keyword\s+Visits\s+Revenue\s+Transactions\s+"
    Dim okXSheet As Boolean
    Dim i As Long
    For i = 0 To UBound(Lines)
        If reheader.Test(Lines(i)) Then
            okXSheet = True
            i = UBound(Lines) + 1
        ElseIf Left(Lines(i), 8) = "Keyword" & vbTab Then
            'ok, forget it
            i = UBound(Lines) + 1
        Else
            'next
        End If
    Next i
    importGetIsCorrectSpreadsheet = okXSheet
    Set reheader = Nothing
End Function
