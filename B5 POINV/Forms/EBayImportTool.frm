VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form EBayImportTool 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "eBay Item Import Tool"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   12660
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox qtyToList 
      Height          =   255
      Left            =   11430
      TabIndex        =   10
      Text            =   "qty"
      Top             =   7470
      Width           =   1065
   End
   Begin VB.CheckBox sellingLocal 
      Caption         =   "sell local"
      Height          =   195
      Left            =   10020
      TabIndex        =   9
      Top             =   8070
      Width           =   1125
   End
   Begin VB.TextBox reservePrice 
      Height          =   255
      Left            =   10020
      TabIndex        =   8
      Text            =   "reserve price"
      Top             =   7740
      Width           =   1155
   End
   Begin VB.TextBox auctionPrice 
      Height          =   255
      Left            =   10020
      TabIndex        =   7
      Text            =   "auction price"
      Top             =   7470
      Width           =   1185
   End
   Begin VB.CommandButton btnFillLastPage 
      Caption         =   "Fill out last page"
      Height          =   615
      Left            =   5190
      TabIndex        =   6
      Top             =   7620
      Width           =   2175
   End
   Begin VB.TextBox LoadingTimer 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Height          =   315
      Left            =   1950
      TabIndex        =   5
      Top             =   7830
      Width           =   1035
   End
   Begin VB.TextBox ItemNumber 
      Height          =   255
      Left            =   8010
      TabIndex        =   4
      Text            =   "TOODRLBX"
      Top             =   7470
      Width           =   1215
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3630
      TabIndex        =   3
      Top             =   7860
      Width           =   1215
   End
   Begin VB.TextBox Password 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "V7V6s4C3"
      Top             =   7920
      Width           =   1395
   End
   Begin VB.TextBox UserName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "toolsplu"
      Top             =   7590
      Width           =   1395
   End
   Begin SHDocVwCtl.WebBrowser browser 
      Height          =   7395
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12645
      ExtentX         =   22304
      ExtentY         =   13044
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
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
      Location        =   ""
   End
End
Attribute VB_Name = "EBayImportTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim complete As Boolean
'Dim continuebutton As Boolean

Private Sub browser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    complete = False
    'Debug.Print "beforenav - " & complete
End Sub

Private Sub browser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    'Debug.Print pDisp.ReadyState
    complete = True
    'Debug.Print "doc complete - " & complete
End Sub

'Private Sub browser_DownloadBegin()
'    Debug.Print "downloadbegin"
'End Sub

Public Sub ListItemInEbay(ItemNumber As String, auctionPrice As String)
    Me.ItemNumber = ItemNumber
    Me.auctionPrice = auctionPrice
    Dim reservePrice As String, qty As String
setres:
    If MsgBox("Set a reserve price for this item?", vbYesNo) = vbYes Then
        reservePrice = InputBox("Enter the reserve price:")
        If Not IsNumeric(reservePrice) Then
            MsgBox "Reserve price is not a number!"
            GoTo setres
        ElseIf reservePrice < DLookup("EbayAuctionPrice", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'") Then
            MsgBox "Reserve price should be higher than auction price!"
            GoTo setres
        Else
            reservePrice = Format(reservePrice, "Currency")
        End If
    Else
        reservePrice = ""
    End If
    Me.reservePrice = reservePrice
    Me.sellingLocal = SQLBool(MsgBox("Put a ""selling locally"" disclaimer on the listing?", vbYesNo) = vbYes)
setqty:
    qty = InputBox("Enter the quantity to list:", , DLookup("QtyOnHand", "vActualWhseQty", "ItemNumber='" & Me.ItemNumber & "'"))
    If Not IsNumeric(qty) Then
        GoTo setqty
    End If
    Me.qtyToList = qty


    Me.browser.Navigate2 "http://my.prostores.com"
    waitForPage 60, "login"
    'Sleep 500
    logIn Me.UserName, Me.Password
    SleepActive 2 'refresh to another page, how to handle that?
    Me.browser.Navigate2 findLink("My Store")
    waitForPage 60, "enter store"
    'Sleep 500
    Me.browser.Navigate2 findLink("Store Administration")
    waitForPage 60, "enter admin"
    'Sleep 500
    Me.browser.Navigate2 findLink("Search")
    waitForPage 60, "enter search"
    'Sleep 500
    searchFor Me.ItemNumber
    'Sleep 500
    If clickFoundItem Then
        'Sleep 500
        clickAuctionButton
        'Sleep 500
        setCategories
    Else
        MsgBox "Unable to find the item in the store. Maybe it needs to be added to prostores first?"
    End If
    Me.btnExit.caption = "Close"
End Sub



Private Sub logIn(UserName, Password)
    Dim i As Long
    With Me.browser.Document.Forms(0)
        For i = 0 To .ALL.Length - 1
            If LCase(.ALL(i).nodename) = "input" Then
                Select Case LCase(.ALL(i).name)
                    Case Is = "unsecured_username"
                        .ALL(i).value = UserName
                    Case Is = "unsecured_password"
                        .ALL(i).value = Password
                    Case Is = "login"
                        .ALL(i).Click
                        waitForPage 60, "logging in"
                        Exit Sub
                End Select
            End If
        Next i
    End With
End Sub

Private Sub searchFor(ItemNumber As String)
    Dim i As Long
    With Me.browser.Document.Forms(0)
        For i = 0 To .ALL.Length - 1
            If LCase(.ALL(i).nodename) = "input" Then
                Select Case LCase(.ALL(i).name)
                    Case Is = "skusearch"
                        .ALL(i).value = ItemNumber
                    Case Is = "search"
                        .ALL(i).Click
                        waitForPage 60, "do search"
                        Exit Sub
                End Select
            End If
        Next i
    End With
End Sub

Private Function clickFoundItem() As Boolean
    Dim i As Long
    With Me.browser.Document
        For i = 0 To .ALL.Length - 1
            If LCase(.ALL(i).nodename) = "td" Then
                If LCase(.ALL(i).classname) = "reporttext" Then
                    If LCase(.ALL(i).Children(0).nodename) = "a" Then
                        Me.browser.Navigate2 findLink(.ALL(i).Children(0).Children(0).firstChild.nodeValue)
                        waitForPage
                        clickFoundItem = True
                        Exit Function
                    End If
                End If
            End If
        Next i
    End With
    clickFoundItem = False
End Function

Private Sub clickAuctionButton()
    Dim i As Long
    With Me.browser.Document.Forms(0)
        For i = 0 To .ALL.Length - 1
            If LCase(.ALL(i).nodename) = "input" Then
                If LCase(.ALL(i).name) = "auction" Then
                    .ALL(i).Click
                    waitForPage
                    Exit Sub
                End If
            End If
        Next i
    End With
End Sub

Private Sub setCategories()
    Dim cats(1) As String
    Dim rst As ADODB.Recordset
    Set rst = retrieve("SELECT WebPathName FROM vPNWebPaths WHERE PathLevel=1 AND ItemNumber='" & Me.ItemNumber & "'")
    If Not rst.EOF Then
        cats(0) = rst("WebPathName")
        rst.MoveNext
        cats(1) = IIf(rst.EOF, "None", rst("WebPathName"))
    Else
        cats(0) = "Other"
    End If
    rst.Close
    Set rst = Nothing
    ebayifyPaths cats
    Dim i As Long
    With Me.browser.Document.Forms(0)
        For i = 0 To .ALL.Length - 1
            If LCase(.ALL(i).nodename) = "select" Then
                Select Case LCase(.ALL(i).name)
                    Case Is = "$auctionitem.storecategory"
                        setSelect .ALL(i), cats(0)
                    Case Is = "$auctionitem.storecategory2"
                        setSelect .ALL(i), cats(1)
                End Select
            End If
        Next i
    End With
    MsgBox "Ok, you're on your own." & vbCrLf & vbCrLf & "double check the store categories, fill out the ebay category, click next, and finish that page." & vbCrLf & vbCrLf & "If you cancel, remember to clear the auction price from po/inv."
'    continuebutton = False
'    MsgBox "Enter the eBay category now, then press the CONTINUE button on the SIDE!" & vbCrLf & vbCrLf & "Don't press the next button on the page!!"
'    While Not continuebutton
'        Sleep 50
'        DoEvents
'    Wend
'    With Me.browser.Document.Forms(0)
'        For i = 0 To .ALL.length - 1
'            If LCase(.ALL(i).nodename) = "input" Then
'                If LCase(.ALL(i).name) = "price" Then 'price? i dunno
'                    .ALL(i).Click
'                    waitForPage
'                    Exit Sub
'                End If
'            End If
'        Next i
'    End With
End Sub

Private Sub ebayifyPaths(ByRef categories() As String)
    Dim i As Long
    For i = 0 To 1
        Select Case LCase(categories(i))
            Case Is = "1no-path"
                categories(i) = "Router Bits"
            Case Is = "closeouts & specials"
                categories(i) = "None"
            Case Is = "cordless tool ""accessories"""
                categories(i) = "Cordless Tools"
            Case Is = "layout & marking"
                categories(i) = "None"
            Case Is = "miscellaneous items"
                categories(i) = "Other"
            Case Is = "power tool ""accessories"""
                categories(i) = "Power Tools"
            Case Is = "power tools ""electric"""
                categories(i) = "Power Tools"
            Case Is = "woodworking mach. ""access"""
                categories(i) = "Woodworking Machinery"
        End Select
    Next i
    If categories(0) = "None" Then
        If categories(1) = "None" Then
            categories(0) = "Other"
        Else
            categories(0) = categories(1)
            categories(1) = "None"
        End If
    End If
End Sub

Private Sub setSelect(selectElement As Object, val As String)
    Dim i As Long
    For i = 0 To selectElement.Options.ALL.Length - 1
        If LCase(selectElement.Options.ALL(i).Text) = LCase(val) Then
            selectElement.value = selectElement.Options.ALL(i).value
            Exit For
        End If
    Next i
End Sub



Private Function waitForPage(Optional timeout As Long = 60, Optional pageName As String = "") As Boolean
    complete = False
    Debug.Print "Started wait" & IIf(pageName <> "", " for " & pageName, "")
    Mouse.Hourglass True
    Dim starttime As Double
    starttime = Timer
    Me.LoadingTimer.BackColor = RED
    'Debug.Print "started wait - " & complete
    While Not complete
        DoEvents
        Me.LoadingTimer = Timer - starttime
        If Timer - starttime > timeout Then
            Me.browser.Stop
            complete = False
            waitForPage = False
            Exit Function
        End If
    Wend
    If Timer - starttime < 2 Then
        Sleep (2 - (Timer - starttime)) * 1000
        'Debug.Print "extra wait"
    End If
    complete = False
    Me.LoadingTimer = ""
    Me.LoadingTimer.BackColor = GREEN
    waitForPage = True
    Mouse.Hourglass False
    Debug.Print "Finished wait" & IIf(pageName <> "", " for " & pageName, "")
End Function

Private Function findLink(linkText As String) As String
    findLink = ""
    Dim i As Long
    For i = 1 To Me.browser.Document.Links.Length - 1
        If Me.browser.Document.Links(i).innertext = linkText Then
            findLink = Me.browser.Document.Links(i).href
            Exit For
        End If
    Next i
End Function

Private Sub btnFillLastPage_Click()
    With Me.browser.Document.Forms(0)
        Dim i As Long
        If Me.sellingLocal Then
            
        End If
        For i = 0 To .ALL.Length - 1
            If LCase(.ALL(i).nodename) = "input" Then
                If InStrRev(.ALL(i).name, "#") Then
                    Select Case Right(LCase(.ALL(i).name), Len(.ALL(i).name) - InStrRev(.ALL(i).name, "#") + 1)
                        Case Is = "#minimumbid"
                            .ALL(i).value = Me.auctionPrice
                        Case Is = "#reserveprice"
                            .ALL(i).value = Me.reservePrice
                        Case Is = "#quantity"
                            .ALL(i).value = Me.qtyToList
                    End Select
                End If
            End If
            If LCase(.ALL(i).nodename) = "textarea" And Me.sellingLocal Then
                If LCase(.ALL(i).name) = "$auctionitem.description" Then
                    .ALL(i).value = .ALL(i).value & "<hr /><p>This item is also being sold locally, so this auction may be cancelled before it ends."
                End If
            End If
        Next i
    End With
End Sub
