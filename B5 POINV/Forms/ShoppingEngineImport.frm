VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form ShoppingEngineImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shopping Engines - Import Data"
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   10680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   465
      Left            =   8760
      TabIndex        =   13
      Top             =   10800
      Width           =   1365
   End
   Begin VB.CommandButton btnImport 
      Caption         =   "Import"
      Height          =   585
      Left            =   5370
      TabIndex        =   12
      Top             =   10740
      Width           =   1755
   End
   Begin VB.OptionButton opEng 
      Caption         =   "Nextag"
      Height          =   225
      Index           =   2
      Left            =   2850
      TabIndex        =   11
      Top             =   11190
      Width           =   1935
   End
   Begin VB.OptionButton opEng 
      Caption         =   "Shopzilla"
      Height          =   225
      Index           =   1
      Left            =   2850
      TabIndex        =   10
      Top             =   10890
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton opEng 
      Caption         =   "Shopping.com"
      Height          =   225
      Index           =   0
      Left            =   2850
      TabIndex        =   9
      Top             =   10590
      Width           =   1935
   End
   Begin VB.TextBox password 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   8
      Text            =   "summer01"
      Top             =   11190
      Width           =   1215
   End
   Begin VB.TextBox password 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Text            =   "summer01"
      Top             =   10890
      Width           =   1215
   End
   Begin VB.TextBox password 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Text            =   "summer01"
      Top             =   10590
      Width           =   1215
   End
   Begin VB.TextBox username 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Text            =   "esavelle"
      Top             =   11190
      Width           =   1455
   End
   Begin VB.TextBox username 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Text            =   "esavelle"
      Top             =   10890
      Width           =   1455
   End
   Begin VB.TextBox username 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Text            =   "eric@toolsplus.com"
      Top             =   10590
      Width           =   1455
   End
   Begin SHDocVwCtl.WebBrowser browser 
      Height          =   10155
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10605
      ExtentX         =   18706
      ExtentY         =   17912
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
   Begin VB.Label Label2 
      Caption         =   "password"
      Height          =   255
      Left            =   1590
      TabIndex        =   2
      Top             =   10290
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "username"
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   10290
      Width           =   1245
   End
End
Attribute VB_Name = "ShoppingEngineImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SHOPPING = 0
Private Const SHOPZILLA = 1
Private Const NEXTAG = 2

Private complete As Boolean

Private Sub browser_DocumentComplete(ByVal pDisp As Object, url As Variant)
    complete = True
End Sub

Private Sub btnExit_Click()
    Unload ShoppingEngineImport
End Sub

Private Sub btnImport_Click()
    Select Case True
        Case Is = Me.opEng(SHOPPING)
            importShopping
        Case Is = Me.opEng(SHOPZILLA)
            importShopzilla
        Case Is = Me.opEng(NEXTAG)
            importNextag
    End Select
End Sub

Private Sub importShopping()
    complete = False
    Mouse.Hourglass True
    Me.btnImport.Enabled = False

    Dim todaysdate As String, currdate As String
    todaysdate = Format(Date, "mm/dd/yyyy")
    currdate = Format(DateAdd("d", 1, getLastDate("Shopping")), "mm/dd/yyyy")
    If todaysdate = currdate Then
        MsgBox "already up to date!"
        Exit Sub
    End If
    
    Me.browser.Navigate2 "https://merchant.shopping.com/mac/app"
    waitForPage 60
    
    Me.browser.Document.Forms(0).item("loginName").value = Me.username(SHOPPING)
    Me.browser.Document.Forms(0).item("password").value = Me.password(SHOPPING)
    Me.browser.Document.Forms(0).submit
    waitForPage 60
    
    Me.browser.Navigate2 findLink("Reports")
    waitForPage 60
    
    Me.browser.Navigate2 findLink("Traffic and sales report")
    waitForPage 60
    
    While todaysdate <> currdate
    
        Me.browser.Document.Forms(0).item("$RadioGroup").item(1).Click
        Me.browser.Document.Forms(0).item("sDate").value = currdate
        Me.browser.Document.Forms(0).item("eDate").value = currdate
        Me.browser.Document.Forms(0).submit
        waitForPage 60
        
        Dim re As RegExp, m As MatchCollection
        Dim cost As String, rev As String, orders As String
        Set re = New RegExp
        re.IgnoreCase = True
        re.MultiLine = True
        re.Pattern = ">Grand Total</td>.*col4>\s*\$([\d,]+\.\d\d)\s*</td>.*col6>\s*([\d,]+)?\s*</td>.*col8>\s*\$([\d,]+\.\d\d)\s*</td>"
        If re.test(Replace(Replace(Me.browser.Document.body.innerHTML, vbCr, ""), vbLf, "")) Then
            Set m = re.Execute(Replace(Replace(Me.browser.Document.body.innerHTML, vbCr, ""), vbLf, ""))
            cost = Replace(m.item(0).SubMatches(0), ",", "")
            orders = Replace(m.item(0).SubMatches(1), ",", "")
            rev = Replace(m.item(0).SubMatches(2), ",", "")
            If orders = "" Then
                orders = "0"
            End If
            dbExecute "INSERT INTO ShoppingEngineStatistics ( Engine, SaleDate, TotalCost, TotalRevenue, TotalOrders ) VALUES ( 'Shopping', '" & currdate & "', " & cost & ", " & rev & ", " & orders & " )"
        Else
            'can't find "grand total" line, so maybe nothing happened that day?
            dbExecute "INSERT INTO ShoppingEngineStatistics ( Engine, SaleDate, TotalCost, TotalRevenue, TotalOrders ) VALUES ( 'Shopping', '" & currdate & "', 0, 0, 0 )"
        End If
        Set m = Nothing
        Set re = Nothing
        currdate = Format(DateAdd("d", 1, currdate), "mm/dd/yyyy")
        
        Sleep 5000
    Wend
    
    MsgBox "All done!"

    Me.btnImport.Enabled = True
    Mouse.Hourglass False
End Sub

Private Sub importShopzilla()
    complete = False
    Mouse.Hourglass True
    Me.btnImport.Enabled = False
    
    Dim todaysdate As String, currdate As String
    todaysdate = Format(Date, "mm/dd/yyyy")
    'shopzilla doesn't do yesterday, so today=yesterday
    todaysdate = Format(DateAdd("d", 1, todaysdate), "mm/dd/yyyy")
    currdate = Format(DateAdd("d", 1, getLastDate("ShopZilla")), "mm/dd/yyyy")
    If todaysdate = currdate Then
        MsgBox "already up to date!"
        Exit Sub
    End If
    
    Me.browser.Navigate2 "https://merchant.shopzilla.com/"
    waitForPage 60
    
    Me.browser.Document.Forms(0).item("uname").value = Me.username(SHOPZILLA)
    Me.browser.Document.Forms(0).item("passw").value = Me.password(SHOPZILLA)
    Me.browser.Document.Forms(0).submit
    waitForPage 60
    
    Me.browser.Navigate2 findLink("View Reports")
    waitForPage 60
    
    Me.browser.Navigate2 findLink("ROI & Conversion Report")
    waitForPage 60
    
    While todaysdate <> currdate
    
        setSelect Me.browser.Document.Forms(0).item("begin_day"), Format(currdate, "dd")
        setSelect Me.browser.Document.Forms(0).item("to_day"), Format(currdate, "dd")
        setSelect Me.browser.Document.Forms(0).item("begin_month"), Format(currdate, "mmm")
        setSelect Me.browser.Document.Forms(0).item("to_month"), Format(currdate, "mmm")
        setSelect Me.browser.Document.Forms(0).item("begin_year"), Format(currdate, "yyyy")
        setSelect Me.browser.Document.Forms(0).item("to_year"), Format(currdate, "yyyy")
        'setSelect Me.browser.Document.Forms(0).item("categories"), "All Level One Categories"
        '.click so it runs the js
        Me.browser.Document.Forms(0).item("categories").options(0).Click
        Me.browser.Document.Forms(0).item("submit_filters").Click
        waitForPage 60
        
        Dim re As RegExp, m As MatchCollection
        Dim cost As String, rev As String, orders As String
        Set re = New RegExp
        re.IgnoreCase = True
        re.MultiLine = True
        re.Pattern = ">Total:</td>.*?Total PPC Costs.*?>.*?([\d,]+\.\d\d)</span>.*?Orders.*?>(\d+)</td>.*?Sales Volume.*?>.*?([\d,]+\.\d\d)</span>"
        If re.test(Replace(Replace(Me.browser.Document.body.innerHTML, vbCr, ""), vbLf, "")) Then
            Set m = re.Execute(Replace(Replace(Me.browser.Document.body.innerHTML, vbCr, ""), vbLf, ""))
            cost = Replace(m.item(0).SubMatches(0), ",", "")
            orders = Replace(m.item(0).SubMatches(1), ",", "")
            rev = Replace(m.item(0).SubMatches(2), ",", "")
            dbExecute "INSERT INTO ShoppingEngineStatistics ( Engine, SaleDate, TotalCost, TotalRevenue, TotalOrders ) VALUES ( 'ShopZilla', '" & currdate & "', " & cost & ", " & rev & ", " & orders & " )"
        Else
            'can't find table?
            dbExecute "INSERT INTO ShoppingEngineStatistics ( Engine, SaleDate, TotalCost, TotalRevenue, TotalOrders ) VALUES ( 'ShopZilla', '" & currdate & "', 0, 0, 0 )"
        End If
        Set m = Nothing
        Set re = Nothing
        currdate = Format(DateAdd("d", 1, currdate), "mm/dd/yyyy")
        
        Sleep 5000
    Wend
    
    MsgBox "All done!"
    
    Me.btnImport.Enabled = True
    Mouse.Hourglass False
End Sub

Private Sub importNextag()
    complete = False
    Mouse.Hourglass True
    Me.btnImport.Enabled = False

    Dim todaysdate As String, currdate As String
    todaysdate = Format(Date, "mm/dd/yyyy")
    currdate = Format(DateAdd("d", 1, getLastDate("NexTag")), "mm/dd/yyyy")
    If todaysdate = currdate Then
        MsgBox "already up to date!"
        Exit Sub
    End If
    
    Me.browser.Navigate2 "https://merchants.nextag.com/"
    waitForPage 60
    
    Me.browser.Document.Forms(1).item("email").value = Me.username(NEXTAG)
    Me.browser.Document.Forms(1).item("password").value = Me.password(NEXTAG)
    Me.browser.Document.Forms(1).submit
    waitForPage 60
    
    While todaysdate <> currdate
    
        Me.browser.Navigate2 findLink("View Sales from NexTag")
        waitForPage 60
    
        Me.browser.Document.Forms(1).item("from").value = currdate
        Me.browser.Document.Forms(1).item("to").value = currdate
        Me.browser.Document.Forms(1).submit
        waitForPage 60
    
        Dim re As RegExp, m As MatchCollection, i As Long
        Dim cost As Double, rev As Double, orders As Long
        cost = 0
        rev = 0
        orders = 0
        Set re = New RegExp
        re.IgnoreCase = True
        re.MultiLine = True
        re.Global = True
        re.Pattern = "<tr>.*?<td.*?>\$([\d,]+\.\d\d)</td>.*?</tr>"
        If Not re.test(Replace(Replace(Me.browser.Document.body.innerHTML, vbCr, ""), vbLf, "")) Then
            rev = 0
            orders = 0
        Else
            Set m = re.Execute(Replace(Replace(Me.browser.Document.body.innerHTML, vbCr, ""), vbLf, ""))
            For i = 0 To m.count - 1
                rev = rev + m.item(i).SubMatches(0)
            Next i
            orders = m.count
            Set re = Nothing
            Set m = Nothing
        End If
        
        Me.browser.GoBack
        waitForPage 60
            
        Me.browser.GoBack
        waitForPage 60
            
        Me.browser.Navigate2 findLink("Balance History")
        waitForPage 60
        
        Me.browser.Document.Forms(1).item("startDate").value = currdate
        Me.browser.Document.Forms(1).item("endDate").value = currdate
        Me.browser.Document.Forms(1).submit
        waitForPage 60
            
        Set re = New RegExp
        Dim shortdate As String
        shortdate = Format(currdate, "m/d/yy")
        re.IgnoreCase = True
        re.MultiLine = True
        re.Pattern = "For [\d,]+ clicks on " & shortdate & ".*?>\$([\d,]+\.\d\d)</td>"
        If re.test(Replace(Replace(Me.browser.Document.body.innerHTML, vbCr, ""), vbLf, "")) Then
            Set m = re.Execute(Replace(Replace(Me.browser.Document.body.innerHTML, vbCr, ""), vbLf, ""))
            cost = m.item(0).SubMatches(0)
            dbExecute "INSERT INTO ShoppingEngineStatistics ( Engine, SaleDate, TotalCost, TotalRevenue, TotalOrders ) VALUES ( 'NexTag', '" & currdate & "', " & cost & ", " & rev & ", " & orders & " )"
        Else
            'no cost yet, do nothing
        End If
        
        Me.browser.GoBack
        waitForPage 60
        
        Me.browser.GoBack
        waitForPage 60
        
        Set m = Nothing
        Set re = Nothing
        currdate = Format(DateAdd("d", 1, currdate), "mm/dd/yyyy")
        
        Sleep 5000
    
    Wend
    
    MsgBox "All done!"
    
    Me.btnImport.Enabled = True
    Mouse.Hourglass False
End Sub


Private Function getLastDate(engine As String) As String
    Dim rst As ADODB.Recordset
    Set rst = retrieve("SELECT MAX(SaleDate) AS LastDate FROM ShoppingEngineStatistics WHERE Engine='" & engine & "'")
    If rst.EOF Then
        getLastDate = "06/22/06"
    Else
        If IsNull(rst("LastDate")) Then
            getLastDate = "06/22/06"
        Else
            getLastDate = rst("LastDate")
        End If
    End If
End Function

Private Function waitForPage(Optional timeout As Long = 60) As Boolean
    Mouse.Hourglass True
    Dim starttime As Double
    starttime = Timer
    'Me.LoadingTimer.BackColor = RED
    While Not complete
        DoEvents
    '    Me.LoadingTimer = Timer - starttime
        If Timer - starttime > timeout Then
            Me.browser.Stop
            complete = False
            waitForPage = False
            Exit Function
        End If
    Wend
    complete = False
    'Me.LoadingTimer = ""
    'Me.LoadingTimer.BackColor = GREEN
    waitForPage = True
    Mouse.Hourglass False
End Function

Private Function findLink(linkText As String) As String
    findLink = ""
    Dim i As Long
    For i = 0 To Me.browser.Document.Links.Length - 1
        If Me.browser.Document.Links(i).innerText = linkText Then
            findLink = Me.browser.Document.Links(i).href
            Exit For
        End If
    Next i
End Function

'Private Sub setSelect(selectElement As HTMLSelectElement, val As String)
Private Sub setSelect(selectElement As Object, val As String)
    Dim i As Long
    For i = 0 To selectElement.options.ALL.Length - 1
        If LCase(selectElement.options.ALL(i).Text) = LCase(val) Then
            selectElement.value = selectElement.options.ALL(i).value
            Exit For
        End If
    Next i
End Sub
