VERSION 5.00
Begin VB.Form WebOffload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Offload"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   7755
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton jetOffloadBtn 
      Caption         =   "Jet.com Offload"
      Height          =   315
      Left            =   3180
      TabIndex        =   33
      Top             =   4100
      Width           =   1965
   End
   Begin VB.CommandButton btnEBayOffload 
      Caption         =   "EBay Offload - SQL"
      Height          =   315
      Index           =   3
      Left            =   5520
      TabIndex        =   32
      Top             =   4320
      Width           =   1965
   End
   Begin VB.CommandButton btnEBayOffload 
      Caption         =   "EBay Offload - no WMH"
      Height          =   315
      Index           =   1
      Left            =   3180
      TabIndex        =   31
      Top             =   3450
      Width           =   1965
   End
   Begin VB.CommandButton btnEBayOffload 
      Caption         =   "EBay Offload - Single Item"
      Height          =   315
      Index           =   2
      Left            =   3180
      TabIndex        =   30
      Top             =   3780
      Width           =   1965
   End
   Begin VB.CommandButton btnEBayOffloadUnlock 
      Caption         =   "Unlock EBay Offload"
      Height          =   285
      Left            =   5550
      TabIndex        =   29
      Top             =   3690
      Width           =   1785
   End
   Begin VB.CommandButton btnImageSyncEBay 
      Caption         =   "Re-sync Images (EBay overrides)"
      Height          =   345
      Left            =   120
      TabIndex        =   27
      Top             =   3510
      Width           =   2505
   End
   Begin VB.CommandButton btnImageSyncYahoo 
      Caption         =   "Re-sync Images (Yahoo/EBay)"
      Height          =   345
      Left            =   120
      TabIndex        =   26
      Top             =   3150
      Width           =   2505
   End
   Begin VB.CommandButton btnEBayOffload 
      Caption         =   "EBay Offload"
      Height          =   315
      Index           =   0
      Left            =   3180
      TabIndex        =   25
      Top             =   3120
      Width           =   1965
   End
   Begin VB.CommandButton btnWebsiteSelection 
      Caption         =   "Website: Tools Plus"
      Height          =   435
      Left            =   3150
      TabIndex        =   24
      Tag             =   "0"
      Top             =   60
      Width           =   2445
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Offload Sections:"
      Height          =   1005
      Index           =   3
      Left            =   90
      TabIndex        =   19
      Top             =   2040
      Width           =   7575
      Begin VB.OptionButton SectionOptions 
         Caption         =   "Offload ALL Sections"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   22
         Top             =   630
         Width           =   3615
      End
      Begin VB.OptionButton SectionOptions 
         Caption         =   "Offload Sections Changed Since Last Offload"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   21
         Top             =   300
         Width           =   3615
      End
      Begin VB.CommandButton btnOffloadSections 
         Caption         =   "Offload Sections"
         Height          =   495
         Left            =   5280
         TabIndex        =   20
         Top             =   210
         Width           =   2175
      End
      Begin VB.Label lblWebOffloadSectionCount 
         Caption         =   "(COUNT)"
         Height          =   255
         Left            =   3750
         TabIndex        =   23
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.Frame frameOptions 
      Caption         =   "General Options:"
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   2415
      Begin VB.CheckBox chkPrintReport 
         Caption         =   "Print Offload Report"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox chkDisplayReport 
         Caption         =   "Display Offload Report"
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   510
         Width           =   1935
      End
   End
   Begin VB.Frame frameFilePath 
      Caption         =   "Save File To:"
      Height          =   1335
      Left            =   2610
      TabIndex        =   9
      Top             =   4680
      Width           =   5055
      Begin VB.OptionButton opFilePath 
         Caption         =   "My Documents"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   2115
      End
      Begin VB.TextBox PathToSave 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   960
         Width           =   4095
      End
      Begin VB.OptionButton opFilePath 
         Caption         =   "C:\"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton opFilePath 
         Caption         =   "Desktop"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1005
      End
      Begin VB.OptionButton opFilePath 
         Caption         =   "Other"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   825
      End
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Offload Products:"
      Height          =   1425
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   570
      Width           =   7575
      Begin VB.CommandButton btnViewLast 
         Caption         =   "View Last Offload Report"
         Height          =   315
         Left            =   5280
         TabIndex        =   7
         Top             =   990
         Width           =   2175
      End
      Begin VB.CommandButton btnRollbackOffload 
         Caption         =   "Roll Back To This Morning"
         Height          =   315
         Left            =   3060
         TabIndex        =   6
         Top             =   990
         Width           =   2175
      End
      Begin VB.CommandButton btnOffloadProducts 
         Caption         =   "Offload Products"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   5280
         TabIndex        =   4
         Top             =   210
         Width           =   2175
      End
      Begin VB.OptionButton ProductOptions 
         Caption         =   "Offload ALL Products"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   630
         Width           =   5340
      End
      Begin VB.OptionButton ProductOptions 
         Caption         =   "Offload Products CHANGED Since Last Web Offload"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Value           =   -1  'True
         Width           =   4080
      End
      Begin VB.Label lblWebOffloadCount 
         Caption         =   "(COUNT)"
         Height          =   255
         Left            =   4260
         TabIndex        =   18
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   6720
      TabIndex        =   0
      Top             =   0
      Width           =   1035
   End
   Begin VB.Label lblEBayOffloadUnlock 
      Alignment       =   2  'Center
      Caption         =   "unlock text"
      Height          =   525
      Left            =   5220
      TabIndex        =   28
      Top             =   3120
      Width           =   2445
   End
   Begin VB.Label Label1 
      Caption         =   "Web Offload"
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
      TabIndex        =   8
      Top             =   60
      Width           =   2415
   End
   Begin VB.Label lblStatusBar 
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   7515
   End
End
Attribute VB_Name = "WebOffload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ByVal lpszPath As String) As Long

Private Const UPDATED = 0
Private Const ALL = 1

Private Const FILE_ROOT = 0
Private Const FILE_DESKTOP = 1
Private Const FILE_OTHER = 2
Private Const FILE_MYDOCS = 3

Dim globalFoundProductLinePage As String
Dim globalPathProblems As Dictionary




Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnOffloadSections_Click()
    
    Dim tempHash As Dictionary, i As Long
    
    Dim WebsiteID As Long
    WebsiteID = CLng(Me.btnWebsiteSelection.tag)
    
    If RunningThroughVPN Then
        If vbNo = MsgBox("Are you sure you want to offload through the VPN? This will probably take a while!", vbYesNo) Then
            Exit Sub
        End If
    End If
    
    Dim strsql As String
    Select Case True
        Case Is = Me.SectionOptions(UPDATED)
            strsql = "SELECT * FROM WebPaths WHERE WebsiteID=" & WebsiteID & " AND PublishStatus=1 AND LastChanged>=LastOffload"
        Case Is = Me.SectionOptions(ALL)
            strsql = "SELECT * FROM WebPaths WHERE WebsiteID=" & WebsiteID & " AND PublishStatus=1"
        Case Else
            MsgBox "Select an offload!"
            Exit Sub
    End Select
    
    Dim filenamepath As String
    filenamepath = getPath() & "weboff-sections." & WebsiteURLHash.item(WebsiteID) & ".csv"
    If Not testWriteAccess(filenamepath) Then
        MsgBox "You don't have write access to this directory!"
        Exit Sub
    End If
    
    Me.btnOffloadSections.Enabled = False
    Mouse.Hourglass True
    
    WebPathCache_Flush
    
    'Dim test As ADODB.Command
    'Set test = New ADODB.Command
    ' test.ActiveConnection = db.
    ' test.CommandTimeout = 50
    Dim StraightQuery As String
    StraightQuery = "SELECT Tmp.ItemNumber,ItemDescription,AVG(InternetPrice) AS InternetPrice,SUM(QtyOnHand) AS QtyOnHand,SUM(QtyOnOrder) AS QtyOnOrder,AVG(CAST(Reconditioned as int)) AS Reconditioned,AVG(Discontinued) AS Discontinued,SUM(QtySold) AS QtySold,SUM(NumTransactions) AS NumTransactions,AVG(dbo.PriceChangeOverRange(RealItemNumber,{STARTDATE},{ENDDATE})) AS PriceChange,COUNT(vSpecialsActive.SpecialName) AS ActiveSpecialCount,MAX(PublishedDate) AS PublishedDate "
    StraightQuery = StraightQuery & "FROM (SELECT ItemNumber,RealItemNumber,ItemDescription,InternetPrice,QtyOnHand,QtyOnOrder,Reconditioned,Discontinued,PublishedDate,SUM(CASE WHEN TransactionDate BETWEEN {STARTDATE} AND {ENDDATE} THEN TransactionQty ELSE 0 END) AS QtySold,SUM(CASE WHEN TransactionDate BETWEEN {STARTDATE} AND {ENDDATE} THEN 1 ELSE 0 END) AS NumTransactions FROM vProductSuggestions WHERE WebsiteID={WEBSITEID} AND TransactionType={TRANSTYPE} AND (({PL}='' AND 1=1) OR ({PL}<>'' AND LEFT(ItemNumber,3)={PL})) GROUP BY ItemNumber, RealItemNumber, ItemDescription, InternetPrice, QtyOnHand, QtyOnOrder, Reconditioned, Discontinued, PublishedDate) AS Tmp "
    StraightQuery = StraightQuery & "LEFT OUTER JOIN vSpecialsActive ON Tmp.ItemNumber=vSpecialsActive.ItemNumber AND vSpecialsActive.ShowOnRebatePage=1 GROUP BY Tmp.ItemNumber, ItemDescription HAVING {INCZERO}=1 OR ({INCZERO}=0 AND SUM(QtySold)>0) ORDER BY QtySold DESC, NumTransactions DESC"
    
    
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT GETDATE()")
    Dim datestamp As String
    datestamp = rst(0)
    rst.Close
    
    Dim StartDate As String
    Dim EndDate As String
    Dim PL As String
    Dim RowLimit As Integer
    Dim IncZero As Integer
    Dim SiteID As Integer
    Dim TransType As String
    
    StartDate = CStr(DateAdd("d", -90, Date))
    EndDate = CStr(Date)
    PL = ""
    RowLimit = 0
    IncZero = 0
    SiteID = WebsiteID
    TransType = "S"
    
    StraightQuery = Replace(StraightQuery, "{STARTDATE}", "'" & StartDate & "'")
    StraightQuery = Replace(StraightQuery, "{ENDDATE}", "'" & EndDate & "'")
    StraightQuery = Replace(StraightQuery, "{PL}", "'" & PL & "'")
    StraightQuery = Replace(StraightQuery, "{ROWLIMIT}", CStr(RowLimit))
    StraightQuery = Replace(StraightQuery, "{INCZERO}", CStr(IncZero))
    StraightQuery = Replace(StraightQuery, "{WEBSITEID}", CStr(SiteID))
    StraightQuery = Replace(StraightQuery, "{TRANSTYPE}", "'" & TransType & "'")
    
    
    
    
    
    
    'cache top seller info
    Dim qry As String
    qry = "EXEC spProductsSales '" & DateAdd("d", -90, Date) & "', '" & Date & "', '', 0, 0, " & WebsiteID & ", 'S'"
     'Set rst = test.Execute(qry)
    DB.TimeoutPeriod = 240
   
    'Set rst = DB.retrieve(qry)
    Set rst = DB.retrieve(StraightQuery)
    'DB.TimeoutPeriod = 30
    Dim topsellers As Dictionary
    Set topsellers = New Dictionary
    While Not rst.EOF
        topsellers.Add CStr(rst("ItemNumber")), RSFieldsAsHash(rst.fields)
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    'cache item-path info
    Set rst = DB.retrieve("SELECT PartNumberWebPaths.ItemNumber, PartNumberWebPaths.WebPathID " & _
                          "FROM PartNumberWebPaths " & _
                          "  INNER JOIN PartNumbers ON PartNumberWebPaths.ItemNumber=PartNumbers.ItemNumber " & _
                          "  INNER JOIN vItemStatusBreakdown ON PartNumbers.ItemNumber=vItemStatusBreakdown.ItemNumber " & _
                          "WHERE PartNumbers.WebPublished=1 " & _
                          "  AND vItemStatusBreakdown.DConWeb=0 " & _
                          "  AND vItemStatusBreakdown.WebsiteID=" & WebsiteID)
    Dim pathItems As Dictionary
    Set pathItems = New Dictionary
    While Not rst.EOF
        If Not pathItems.exists(CStr(rst("WebPathID"))) Then
            pathItems.Add CStr(rst("WebPathID")), New PerlArray
        End If
        pathItems.item(CStr(rst("WebPathID"))).Push CStr(rst("ItemNumber"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    Set rst = DB.retrieve(strsql)
    
    DB.Execute "TRUNCATE TABLE WebOffloadSections"
    
    Dim fieldnames As String
    fieldnames = "Name, Description3, Code, Path, MetaKeywords, MetaDescription, MetaTitleStyle, MetaTitle, Caption, SectionBottomCaption, TopSellers, PageLayoutSections, PageLayoutItems, SortStyleItems, ManufReconURL, ManufContactInfo, IsSectionPage, IsItemPage, IsManufacturerPage"
    Dim fieldvalues(0 To 19) As String
    
    'shop by brand
    fieldvalues(0) = "-1"
    fieldvalues(1) = "'" & "Shop by Brand" & "'"
    fieldvalues(2) = "'" & "Shop by Brand" & "'"
    fieldvalues(3) = "'" & "search-by-manufacturer" & "'"
    fieldvalues(4) = "''"
    fieldvalues(5) = "''"
    fieldvalues(6) = "''"
    fieldvalues(7) = "2"
    fieldvalues(8) = "'Shop by Brand'"
    fieldvalues(9) = "''"
    fieldvalues(10) = "''"
    fieldvalues(11) = "''"
    fieldvalues(12) = "2"
    fieldvalues(13) = "-1"
    fieldvalues(14) = "0"
    fieldvalues(15) = "''"
    fieldvalues(16) = "''"
    fieldvalues(17) = "'Yes'"
    fieldvalues(18) = "'No'"
    fieldvalues(19) = "'No'"
    DB.Execute "INSERT INTO WebOffloadSections ( DatabaseID, " & fieldnames & " ) VALUES ( " & Join(fieldvalues, ", ") & " )"
    
    While Not rst.EOF
        
        Me.lblStatusBar.Caption = "Offloading " & WebPathCache_GetURLFull(rst("ID"))
        DoEvents
        
        Dim topsellersString As String
        Dim pathsToDo As PerlArray, itemsToInclude As Dictionary
        Set pathsToDo = New PerlArray
        Set itemsToInclude = New Dictionary
        pathsToDo.Push CStr(rst("ID"))
        While pathsToDo.Scalar > 0
            Dim thisPath As String
            thisPath = pathsToDo.Shift
            Dim pathKids As Variant
            pathKids = WebPathCache_GetChildList(thisPath)
            If UBound(pathKids) >= 0 Then
                For i = 0 To UBound(pathKids)
                    pathsToDo.Push CStr(pathKids(i))
                Next i
            End If
            If pathItems.exists(thisPath) Then
                For i = 0 To pathItems.item(thisPath).Scalar - 1
                    If Not itemsToInclude.exists(CStr(pathItems.item(thisPath).Elem(i))) Then
                        itemsToInclude.Add CStr(pathItems.item(thisPath).Elem(i)), "1"
                    End If
                Next i
            End If
        Wend
        If itemsToInclude.count = 0 Then
            topsellersString = ""
        Else
            ReDim toptemp(itemsToInclude.count - 1) As Dictionary
            Dim count As Long
            count = 0
            For i = 0 To itemsToInclude.count - 1
                If topsellers.exists(CStr(itemsToInclude.keys(i))) Then
                    Set toptemp(count) = topsellers.item(CStr(itemsToInclude.keys(i)))
                    count = count + 1
                End If
            Next i
            If count = 0 Then
                topsellersString = ""
            Else
                ReDim Preserve toptemp(count - 1) As Dictionary
                QuickSortHashes toptemp, Array(Array("QtySold", False), Array("NumTransactions", False))
                
                Dim tscount As Long
                Select Case itemsToInclude.count
                    Case Is <= 3
                        tscount = 0
                    Case Is <= 8
                        tscount = 2
                    Case Is <= 12
                        tscount = 4
                    Case Else
                        tscount = 8
                End Select
                topsellersString = ""
                For i = 0 To IIf(tscount < UBound(toptemp), tscount - 1, UBound(toptemp))
                    topsellersString = IIf(topsellersString = "", "", topsellersString & " ") & CreateYahooID(toptemp(i).item("ItemNumber"))
                Next i
            End If
        End If
        
        'database id
        fieldvalues(0) = rst("ID")
        'web path name
        fieldvalues(1) = "'" & EscapeSQuotes(WebPathCache_GetName(rst("ID"))) & "'"
        'description 3 (short name)
        fieldvalues(2) = "'" & EscapeSQuotes(WebPathCache_GetName(rst("ID"), False, False)) & "'"
        'code
        fieldvalues(3) = "'" & WebPathCache_GetURLFull(rst("ID")) & "'"
        'path
        fieldvalues(4) = "'" & WebPathCache_GetPath(rst("ID")) & "'"
        'meta keywords
        fieldvalues(5) = "'" & EscapeSQuotes(Nz(rst("MetaKeywords"))) & "'"
        'meta description
        fieldvalues(6) = "'" & EscapeSQuotes(Nz(rst("MetaDescription"))) & "'"
        'meta title style
        fieldvalues(7) = rst("MetaTitleStyle")
        'meta title
        fieldvalues(8) = "'" & EscapeSQuotes(Nz(rst("MetaTitle"))) & "'"
        'top caption
        fieldvalues(9) = "'" & EscapeSQuotes(Nz(rst("CaptionTop"))) & "'"
        'bottom caption
        fieldvalues(10) = "'" & EscapeSQuotes(Nz(rst("CaptionBottom"))) & "'"
        'top sellers
        fieldvalues(11) = "'" & topsellersString & "'"
        'section layout
        fieldvalues(12) = rst("PageLayoutSections")
        'item layout
        If rst("PageLayoutSections") = "1" Then
            fieldvalues(13) = "-1"
        Else
            fieldvalues(13) = rst("PageLayoutItems")
        End If
        'item sort
        fieldvalues(14) = rst("SortStyleItems")
        'reconditioned url
        'contact info block
        'is section
        'is item
        'is manufacturer
        If CBool(rst("IsManufacturerPage")) Then
            fieldvalues(15) = "'" & EscapeSQuotes(DLookup("ReconditionedURL", "ProductLine", "WebSectionID=" & rst("ID"))) & "'"
            fieldvalues(16) = "'" & EscapeSQuotes(WebOffloadFunctions.GetHTMLContactInfoFor(rst("ID"))) & "'"
            fieldvalues(17) = "'Yes'"
            fieldvalues(18) = "'No'"
            fieldvalues(19) = "'Yes'"
        Else
            fieldvalues(15) = "''"
            fieldvalues(16) = "''"
            fieldvalues(17) = "'Yes'"
            fieldvalues(18) = "'No'"
            fieldvalues(19) = "'No'"
        End If
        
        DB.Execute "INSERT INTO WebOffloadSections ( DatabaseID, " & fieldnames & " ) VALUES ( " & Join(fieldvalues, ", ") & " )"
        
        rst.MoveNext
    Wend
    Me.lblStatusBar = ""
    rst.Close
    Set rst = Nothing
    
    ExportAsCSV "SELECT " & fieldnames & " FROM WebOffloadSections", filenamepath, Array("Name", "Description-3", "Code", "Path", "Meta-Keywords", "Meta-Description", "Meta-Title-Style", "Meta-Title", "Caption", "Section-Bottom-Caption", "Top-Sellers", "Page-Layout-Sections", "Page-Layout-Items", "Section-Itemlist-Sort", "Manuf-Recon-Url", "Manuf-Contact-Info", "Is-Section-Page", "Is-Item-Page", "Is-Manufacturer-Page")
    
    DB.Execute "UPDATE WebPaths SET LastOffload='" & datestamp & "' WHERE ID IN (SELECT DatabaseID FROM WebOffloadSections)"
    
    Mouse.Hourglass False
    Me.btnOffloadSections.Enabled = True
    
    UpdateWebOffloadChangedQuantity
    
    MsgBox "Offload complete!"
    
End Sub

Private Sub btnRollbackOffload_Click()
    If MsgBox("Are you sure you want to roll back today's offload?", vbYesNo) = vbYes Then
        Dim Y As Long, m As Long, d As Long
        Y = DatePart("yyyy", Now())
        m = DatePart("m", Now())
        d = DatePart("d", Now())
        DB.Execute "exec spWebOffloadRollBack " & Y & ", " & m & ", " & d
        UpdateWebOffloadChangedQuantity
    End If
End Sub

Private Sub btnOffloadProducts_Click()
    
    If RunningThroughVPN Then
        If vbNo = MsgBox("Are you sure you want to offload through the VPN? This will probably take a while!", vbYesNo) Then
            Exit Sub
        End If
    End If
    
    Dim WebsiteID As Long
    WebsiteID = CLng(Me.btnWebsiteSelection.tag)
    
    Dim strsql As String
    Select Case True
        Case Is = Me.ProductOptions(UPDATED)
            strsql = "SELECT * FROM vWebOffload WHERE WebsiteID=" & WebsiteID & " AND LastChanged>=LastOffload"
        Case Is = Me.ProductOptions(ALL)
            strsql = "SELECT * FROM vWebOffload WHERE WebsiteID=" & WebsiteID
        Case Else
            MsgBox "Select an offload first."
            Exit Sub
    End Select
    
    Mouse.Hourglass True
    Me.btnOffloadProducts.Enabled = False
    
    Dim filenamepath As String
    filenamepath = getPath() & "weboff-items." & WebsiteURLHash.item(WebsiteID) & ".csv"
    If Not testWriteAccess(filenamepath) Then
        MsgBox "You don't have write access to this directory!"
        GoTo theend
    End If
    
    Me.lblStatusBar.Caption = "Retrieving items..."
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT GETDATE()")
    Dim datestamp As String
    datestamp = rst(0)
    rst.Close
    
    Set rst = DB.retrieve(strsql)
    
    If rst.EOF Then
        MsgBox "No items need offloading!"
        GoTo theend
    End If
    
    Me.lblStatusBar.Caption = "Caching Info..."
    DoEvents
    'if we change a manufacturer name, and don't close po/inv, it breaks the website,
    'so let's re-hash everything, on the 1 in a 1000 chance it changed. everything
    'else in these hashes is static enough.
    CreateHashes "ManufHash"
    'flush some other caches
    WebPathCache_Flush
    SpecialsCache_Initialize
    GroupItemCache_Initialize
    
    Me.lblStatusBar.Caption = "Preprocessing..."
    'group items:
    '  any sub-items flagged for offload get template flagged as well, force yahoo to publish stock/price changes
    '  any templates flagged for offload get all sub items flagged as well, possible spec changes to propagate
    While Not rst.EOF
        If GroupItemCache_IsGroupedSubItem(rst("ItemNumber")) Then
            DB.Execute "UPDATE PartNumbers SET LastChanged=DATEADD(s, 1, LastOffload) WHERE LastChanged<LastOffload AND ItemNumber='" & GroupItemCache_GetMasterItemNumberForSubItem(rst("ItemNumber")) & "'"
        ElseIf GroupItemCache_IsGroupMasterItem(rst("ItemNumber")) Then
            DB.Execute "UPDATE PartNumbers SET LastChanged=DATEADD(s, 1, LastOffload) WHERE LastChanged<LastOffload AND ItemNumber IN (SELECT DISTINCT ItemNumber FROM PartNumberGroupLines WHERE GroupID=" & GroupItemCache_GetGroupDatabaseIDByMasterItemNumber(rst("ItemNumber")) & ")"
        End If
        rst.MoveNext
    Wend
    rst.Close
    'reopen to get any changes
    Set rst = DB.retrieve(strsql)
    
    DB.Execute "TRUNCATE TABLE WebOffload"
    
    Set globalPathProblems = New Dictionary
    While Not rst.EOF
        Me.lblStatusBar.Caption = "Processing " & rst("ItemNumber")
        DoEvents
        webOffloadSingleItem rst
        rst.MoveNext
    Wend
    
    If WebsiteID = 0 Then
        'tp only
        addSpecialsPageItemLine
    End If
    
    Me.lblStatusBar.Caption = "Generating CSV"
    Dim colNames As Variant
'XXX: temporarily keep oos around until we can migrate
    colNames = Array("Line-Code", "Name", "Part-Number", "Real-ItemNumber", "Code", "Family", "Logo", "Path", "Manufacturer-Short-Name", "Description-1", "Description-2", "Description-3", "Description-4", "Partial-Specs", "Full-Specs", "Price", "Rebate-Amount", "Is-Instant-Rebate", "Orderable", "Out-Of-Stock", "Stock-Enum", "Availability", "Item-Section-Specials", "Related-Sections", "Accessories", "Similar-Items", "Contents", "Caption", "Ext-Spec", "Meta-Keywords", "Request-Form", "Reconditioned", "MAPP", "Sale-Discount", "Is-Section-Page", "Is-Item-Page", "Shipping-Type", "Is-Meta-No-Index", "Grouped-Subitems", "Top-Sellers", "Grouped-Masteritem", "Name-Mobile", "H-Lite1", "H-Lite2")
    ExportAsCSV "SELECT LineCode, ItemName, PartNumber, RealItemNumber, Code, Family, Logo, Path, ManufacturerShortName, Description1, Description2, Description3, Description4, PartialSpecs, FullSpecs, Price, RebateAmount, IsInstantRebate, Orderable, OutOfStock, StockEnum, Availability, ItemSectionSpecials, RelatedSections, Accessories, SimilarItems, Contents, Caption, ExtSpec, MetaKeywords, RequestForm, IsReconditioned, MAPP, SaleDiscount, IsSectionPage, IsItemPage, ShippingType, IsMetaNoIndex, GroupedSubitems, TopSellers, GroupedMasteritem, NameMobile, HLite1, HLite2 FROM WebOffload", filenamepath, colNames
    
    Me.lblStatusBar.Caption = "Updating offload timestamps"
    'this won't update base items correctly, do it the slow way
    'DB.Execute "UPDATE PartNumbers SET LastOffload='" & datestamp & "' WHERE ItemNumber IN (SELECT RealItemNumber FROM WebOffload)"
    rst.MoveFirst
    While Not rst.EOF
        DoEvents
        DB.Execute "exec spWebOffloadUpdateItemOffloadStamp '" & rst("ItemNumber") & "', '" & datestamp & "'"
        rst.MoveNext
    Wend
    
    Dim rstCheck As ADODB.Recordset
    Set rstCheck = DB.retrieve("SELECT DISTINCT RealItemNumber FROM WebOffload WHERE IsItemPage='Yes' AND Price<>'0.00' AND Availability='This item is not stocked or has been discontinued'")
    While Not rstCheck.EOF
        If Not globalPathProblems.exists(CStr(rstCheck("RealItemNumber"))) Then
            globalPathProblems.Add CStr(rstCheck("RealItemNumber")), "has price, but availability says d/c"
        Else
            globalPathProblems.item(CStr(rstCheck("RealItemNumber"))) = globalPathProblems.item(CStr(rstCheck("RealItemNumber"))) & ", has price, but availability says d/c"
        End If
        rstCheck.MoveNext
    Wend
    rstCheck.Close
    Set rstCheck = DB.retrieve("SELECT DISTINCT RealItemNumber FROM WebOffload WHERE IsItemPage='Yes' AND Price='0.00' AND Availability<>'This item is not stocked or has been discontinued'")
    While Not rstCheck.EOF
        If GroupItemCache_IsGroupMasterItem(rstCheck("RealItemNumber")) Then
            'skip
        Else
            If Not globalPathProblems.exists(CStr(rstCheck("RealItemNumber"))) Then
                globalPathProblems.Add CStr(rstCheck("RealItemNumber")), "no price, but availability is not d/c"
            Else
                globalPathProblems.item(CStr(rstCheck("RealItemNumber"))) = globalPathProblems.item(CStr(rstCheck("RealItemNumber"))) & ", no price, but availability is not d/c"
            End If
        End If
        rstCheck.MoveNext
    Wend
    rstCheck.Close
    Set rstCheck = DB.retrieve("SELECT DISTINCT RealItemNumber FROM WebOffload WHERE IsItemPage='Yes' AND Price='0.00' AND Orderable='Yes'")
    While Not rstCheck.EOF
        If Not globalPathProblems.exists(CStr(rstCheck("RealItemNumber"))) Then
            globalPathProblems.Add CStr(rstCheck("RealItemNumber")), "no price, but orderable flag is checked"
        Else
            globalPathProblems.item(CStr(rstCheck("RealItemNumber"))) = globalPathProblems.item(CStr(rstCheck("RealItemNumber"))) & ", no price, but orderable flag is checked"
        End If
        rstCheck.MoveNext
    Wend
    rstCheck.Close
    Set rstCheck = Nothing
    
    If globalPathProblems.count > 0 Then
        Open getPath() & "badrecs.lst" For Output As #1
        Print #1, Now() & vbCrLf '& "The following sections are unpublished:"
        Dim iter As Variant
        For Each iter In globalPathProblems.keys
            If IsNumeric(iter) Then
                Print #1, DLookup("dbo.WebPathPrintableName(" & CStr(iter) & ")") & ": " & globalPathProblems.item(CStr(iter))
            Else
                Print #1, iter & ": " & globalPathProblems.item(CStr(iter))
            End If
        Next iter
        Close #1
        MsgBox "Found errors in this offload!"
        Shell "notepad " & getPath() & "badrecs.lst", vbNormalFocus
    End If
    
    On Error GoTo errh
    'If Me.chkDisplayReport Or Me.chkPrintReport Then
    '    Me.lblStatusBar.Caption = "Getting report information"
    '    Unload RptWebOffload
    '    Load RptWebOffload
    '    Set RptWebOffload.DataSource = DB.retrieve("SELECT * FROM vWebOffloadReportSource ORDER BY Changes, LineCode, PartNumber")
    '    'RptWebOffload.Orientation = rptOrientLandscape
    '    RptWebOffload.TopMargin = 0.25
    '    RptWebOffload.BottomMargin = 0.25
    '    RptWebOffload.LeftMargin = 0.5
    '    RptWebOffload.RightMargin = 0.5
    '    If Me.chkPrintReport Then
    '        RptWebOffload.PrintReport False
    '    End If
    '    If Me.chkDisplayReport Then
    '        RptWebOffload.Show
    '    End If
    'Else
        MsgBox "Offload complete!"
    'End If
    On Error GoTo 0
    
theend:
    Me.lblStatusBar = ""
    If Not rst Is Nothing Then
        rst.Close
        Set rst = Nothing
    End If
    
    SpecialsCache_Clear
    GroupItemCache_Clear
    Mouse.Hourglass False
    Me.btnOffloadProducts.Enabled = True
    Me.btnOffloadProducts.SetFocus
    UpdateWebOffloadChangedQuantity
    Exit Sub
errh:
    If Err.Number = 372 Then
        MsgBox "Can't display offload report!" & vbCrLf & vbCrLf & "Brian should fix this, but he's lazy. It's probably a permissions issue or something"
    Else
        MsgBox "Error " & Err.Number & " displaying offload report!" & vbCrLf & vbCrLf & Err.Description
    End If
    Err.Clear
    GoTo theend
End Sub

Public Sub btnEBayOffload_Click(index As Integer)
    If IsProcessFlagSet("ebayoffload") Then
        MsgBox "An EBay offload is already running!" & vbCrLf & vbCrLf & DLookup("UserName", "ProcessFlags", "Task='ebayoffload'") & " at " & DLookup("DateStamp", "ProcessFlags", "Task='ebayoffload'")
        Exit Sub
    End If
    
    Dim whereClause As String
    Select Case index
        Case Is = 0 ' default
            whereClause = "EBayOffloadRequired=1"
        Case Is = 1 ' wmh
            whereClause = "EBayOffloadRequired=1 AND LEFT(ItemNumber,3) NOT IN ('JEA','JET','POA','POW','PER','WIL')"
        Case Is = 2 ' single
            whereClause = InputBox("Enter item number:")
            If whereClause = "" Then
                Exit Sub
            End If
            whereClause = "ItemNumber='" & EscapeSQuotes(whereClause) & "'"
            
        Case Is = 3 ' sql
            whereClause = InputBox("Enter SQL clause:")
            If whereClause = "" Then
                Exit Sub
            End If
        Case Is = 4 ' Tom's Code to allow offload from inventory maintenance form...
            whereClause = "ItemNumber='" & InventoryMaintenance.ItemNumber.Text & "'"
            index = 2
            If whereClause = "" Then
                Exit Sub
            End If
    End Select
    
    SetProcessFlag "ebayoffload"
    
    Dim stopFlag As Boolean
    stopFlag = False
    Dim skipAddRelistFlag As Boolean
    skipAddRelistFlag = False
    
    Me.lblStatusBar.Caption = "Initializing"
    GroupItemCache_Initialize
    SpecialsCache_Initialize
    
    Me.lblStatusBar.Caption = "Loading EBay data"
    Dim offloadLines As PerlArray
    Mouse.Hourglass True
    Set offloadLines = EBay.EBayOffloadItemsAll(whereClause)
    
    GroupItemCache_Clear
    SpecialsCache_Clear
    
    If offloadLines.Scalar = 0 Then
        Mouse.Hourglass False
        MsgBox "Nothing to do!"
        Me.lblStatusBar.Caption = ""
        ClearProcessFlag "ebayoffload"
        Exit Sub
    End If
    
    
    Me.lblStatusBar.Caption = "Preprocessing"
    Dim i As Long, def As Dictionary
    Dim todoReport As String
    todoReport = ""
    For i = 0 To offloadLines.Scalar - 1
        Set def = offloadLines.Elem(i)
        todoReport = todoReport & vbCrLf & def.item("ItemNumber") & ": " & def.item("Action")
    Next i
    
    If vbNo = MsgBox("Continue with the following " & offloadLines.Scalar & " actions?" & vbCrLf & todoReport, vbYesNo) Then
        Mouse.Hourglass False
        Me.lblStatusBar.Caption = ""
        ClearProcessFlag "ebayoffload"
        Exit Sub
    End If
    
    Dim successCount As Long, errorcount As Long
    successCount = 0
    errorcount = 0
    Dim errmsg As String
    errmsg = ""
    
    For i = 0 To offloadLines.Scalar - 1
        Set def = offloadLines.Elem(i)
        Dim success As Boolean
        Debug.Print def.item("ItemNumber") & ": " & def.item("Action")
        Me.lblStatusBar.Caption = def.item("ItemNumber") & ": " & def.item("Action")
        DoEvents
        Select Case def.item("Action")
            Case Is = "Add"
                If skipAddRelistFlag = False Then
                    success = EBay.EBayAPIAddFixedPriceItem(def)
                End If
            Case Is = "Revise"
                success = EBay.EBayAPIReviseFixedPriceItem(def)
            Case Is = "Relist"
                If skipAddRelistFlag = False Then
                    success = EBay.EBayAPIRelistFixedPriceItem(def)
                    If (success = False) Then
                        success = EBay.EBayAPIReviseFixedPriceItem(def)
                    End If
                    
                End If
            Case Is = "End"
                success = EBay.EBayAPIEndFixedPriceItem(def)
            Case Else
                MsgBox "unhandled ebay action type '" & def.item("Action") & "'"
        End Select
        Debug.Print def.item("ItemNumber") & ": " & def.item("Action") & " -> " & IIf(success, "success", "failure")
        If success Then
            successCount = successCount + 1
        Else
            errorcount = errorcount + 1
            errmsg = errmsg & def.item("ItemNumber") & ": " & EBay.EBayGetLastUnhandledErrorMessage & vbCrLf
        End If
        
        If EBay.EBaySellingLimitReached Then
            Select Case MsgBox("Selling limit has been reached!" & vbCrLf & vbCrLf & "Click ""Yes"" to keep processing normally, click ""No"" continue but skip add/relist actions, or click ""Cancel"" to stop processing entirely.", vbYesNoCancel)
                Case Is = vbYes
                    'nothing to do
                Case Is = vbNo
                    skipAddRelistFlag = True
                Case Is = vbCancel
                    stopFlag = True
            End Select
        End If
        
        If EBay.EBayAPICallLimitReached Then
            MsgBox "API call limit has been reached!" & vbCrLf & vbCrLf & "This is totally going to fuck up orders and tracking info for the rest of the day."
            stopFlag = True
        End If
        
        If stopFlag = True Then
            Exit For
        End If
    Next i
    
    ClearProcessFlag "ebayoffload"
    
    If errorcount > 0 Then
        Dim msg As String
        If stopFlag = True Then
            msg = "Offload aborted!" & vbCrLf & vbCrLf
        ElseIf skipAddRelistFlag = True Then
            msg = "Offload partially aborted!" & vbCrLf & vbCrLf
        ElseIf successCount > 0 Then
            msg = "Offload completed, but with errors!" & vbCrLf & vbCrLf
        Else
            msg = "Error offloading!" & vbCrLf & vbCrLf
        End If
        
        If successCount > 0 Then
            MsgBox msg & successCount & " items updated, but " & errorcount & " items had problems." & vbCrLf & vbCrLf & errmsg
        Else
            MsgBox msg & "Unable to offload " & errorcount & " items!" & vbCrLf & vbCrLf & errmsg
        End If
        
        SendEmailTo "kirk@tools-plus.com", "ebay offload report", errmsg
    ElseIf successCount > 0 Then
        MsgBox "Offload completed without errors!" & vbCrLf & vbCrLf & successCount & " items updated."
    Else
        MsgBox "Zero errors, zero successes? That's really weird!"
    End If
    
    Dim missingImages As String
    missingImages = EBay.EBayGetLastImageErrors()
    If missingImages <> "" Then
        SendEmailTo "brandon@tools-plus.com; kirk@tools-plus.com", "ebay missing images", missingImages
    End If
    
    Mouse.Hourglass False
    Me.lblStatusBar.Caption = ""
End Sub

Private Sub btnEBayOffloadUnlock_Click()
    If vbYes = MsgBox("Unset the ""running"" flag? Only do this if the ebay offload crashed last time!", vbYesNo) Then
        ClearProcessFlag "ebayoffload"
        Me.lblEBayOffloadUnlock.Visible = False
        Me.btnEBayOffloadUnlock.Visible = False
    End If
End Sub

Private Sub checkEBayOffloadRunning()
    Dim pflag As String
    pflag = DLookup("UserName", "ProcessFlags", "Task='ebayoffload'")
    If pflag = "NOTRUNNING" Then
        Me.lblEBayOffloadUnlock.Visible = False
        Me.btnEBayOffloadUnlock.Visible = False
    Else
        Me.lblEBayOffloadUnlock.Caption = pflag & " - " & DLookup("DateStamp", "ProcessFlags", "Task='ebayoffload'")
        Me.lblEBayOffloadUnlock.Visible = True
    End If
End Sub

Private Sub btnImageSyncYahoo_Click()
    Shell PERL & " -MToolsPlus::ImageDB -e ""ToolsPlus::ImageDB::synchronize_auth(verbose=>4);""", vbNormalFocus
End Sub

Private Sub btnImageSyncEBay_Click()
    Shell PERL & " -MToolsPlus::ImageDB -e ""ToolsPlus::ImageDB::synchronize_noauth(verbose=>4);""", vbNormalFocus
    MsgBox "Shopping engines will update with the new images automatically, but EBay requires an offload!"
End Sub


Private Sub webOffloadSingleItem(rst As ADODB.Recordset)
    Dim itemName As String
    Dim path As Variant
    Dim desc1 As String, desc2 As String, desc3 As String, desc4 As String
    'Dim options As String
    Dim capt As String, crossSellLinks As String, contents As String
    Dim Accessories As String, similarItems As String
    'Dim showURL As String, showImage As String
    Dim sectionSpecialStanza As String
    Dim orderable As String, outofstock As String, stockEnum As Long, Availability As String
    Dim price As String, saleDisc As Long, mappprice As String
    Dim metas As String
    Dim mobileName As String, hlite1 As String, hlite2 As String 'for practical data
    
    'at this point, we assume that all YahooIDs are unique
    'they should be, if everything's been going through po/inv correctly
    Dim newitemnumber As String
    If Not GroupItemCache_IsGroupMasterItem(rst("ItemNumber")) Then
        newitemnumber = CreateYahooID(rst("ItemNumber"))
    Else
        newitemnumber = GroupItemCache_GetGroupYahooIDByMasterItemNumber(rst("ItemNumber"))
    End If
    
    Dim i As Long
    
    Dim allWebPaths As PerlArray
    globalFoundProductLinePage = "" 'this fucking blows
    If Not GroupItemCache_IsGroupedSubItem(rst("ItemNumber")) Then
        Set allWebPaths = getWebPathsForItem(rst("ItemNumber"))
    Else
        'ok, this is fucking awful. get the paths for the parent template,
        'just so we can get the globalFoundProductPage set, so we can have
        'a manuf logo on the page. then wipe it out, and put in a single
        'empty path so it'll offload something.
        Set allWebPaths = getWebPathsForItem(GroupItemCache_GetMasterItemNumberForSubItem(rst("ItemNumber")))
        Set allWebPaths = Nothing
        Set allWebPaths = New PerlArray
        allWebPaths.Push ""
    End If
    If allWebPaths.Scalar = 0 Then
        Exit Sub
    End If
    
    desc1 = EscapeHTMLEntities(Trim(Nz(rst("Desc1"))))
    desc2 = EscapeHTMLEntities(Trim(Nz(rst("Desc2"))))
    desc3 = EscapeHTMLEntities(Trim(Nz(rst("Desc3"))))
    desc4 = EscapeHTMLEntities(Trim(Nz(rst("Desc4"))))
    
    orderable = rst("WebOrderable")
    If GroupItemCache_IsGroupMasterItem(rst("ItemNumber")) Then
        orderable = "No"
    End If
    If rst("ShowRequestForm") = "Yes" Then  'request forms w/ orderable makes
        orderable = "Yes"                   'price appear on section pages
    End If
    'options = ""
    Availability = GetAvailabilityText(rst("ItemStatus"), rst("QtyOnHand"), rst("QtyOnSO"), rst("QtyOnPO"), rst("QtyOrdered"), rst("AvailabilityLimit"), IIf(rst("ShowRequestForm") = "Yes", True, False), Nz(rst("AvailString")))
    itemName = Nz(rst("ManufFullNameWeb")) & " " & desc2 & " " & desc3
    mobileName = TrimWhitespace(Nz(rst("ManufFullNameWeb")) & " " & desc2 & " " & desc1 & " " & desc3, True, True, True, True)
    
    contents = ""
    If rst("ShowDCSpecs") Then
        Accessories = ""
        similarItems = ""
    Else
        Accessories = GetItemAccessoriesAsString(rst("ItemNumber"))
        'similarItems = GetItemSimilarItemsAsString(rst("ItemNumber"))
        similarItems = ""
    End If
    
    If CBool(rst("ShowDCSpecs")) Then
        capt = "<!--nosearch-->" & vbCrLf
    End If
    capt = capt & TrimWhitespace(desc1 & " " & desc3 & " " & desc4 & IIf(rst("IsReconditioned") = "Yes", " Reconditioned", ""), True, True, True) & vbCrLf & vbCrLf
    capt = capt & WebOffloadFunctions.GenerateKeywordAlternatesFromPathing(rst("ItemNumber"))
    
    
    If rst("MAPP") <> 0 And rst("MAPP") > rst("IDiscountMarkupPriceRate1") Then
        capt = "<!--noyshopping-->" & vbCrLf & capt
    End If

    metas = GenerateYahooMetaKeywords(rst("ItemNumber"))
    
    crossSellLinks = GenerateCrossSellText(rst("ItemNumber"))
    
    'section image
    'showImage = SpecialsCache_GetSectionImage(rst("ItemNumber"))
    'showURL = SpecialsCache_GetSectionURL(rst("ItemNumber"))
    sectionSpecialStanza = SpecialsCache_GetSectionSpecialStanza(rst("ItemNumber"))
    ''$6.50 freight only appears if no other 'better' specials
    'If rst("ShippingType") = 2 And showImage = "" And Not rst("ShowDCSpecs") Then
    '    showImage = "650-freight-icon.png"
    '    showURL = newitemnumber & ".html#truck-freight-disclaimer"
    'End If
    
    'sale percentage discounts
    saleDisc = SpecialsCache_GetFlatPercentDiscount(rst("ItemNumber"))
    
    'mapp
    mappprice = CStr(rst("MAPP"))
    
    price = rst("IDiscountMarkupPriceRate1")
    If saleDisc <> 0 Then
        'price is 0 here for d/c items, so we don't have to worry about that.
        price = FormatNumber(price * (1 - (saleDisc / 100)), 2, vbTrue, vbFalse, vbFalse)
    ElseIf SpecialsCache_GetRebateInstantFlag(rst("ItemNumber")) Then
        'only need instant rebates here, they affect price, mail-in doesn't
        price = FormatNumber(price - SpecialsCache_GetRebateAmount(rst("ItemNumber")), 2, vbTrue, vbFalse, vbFalse)
    Else
        'regular and mail-in rebate pricing
        price = FormatNumber(price, 2, vbTrue, vbFalse, vbFalse)
    End If
    For i = 1 To 3
        If rst("IBreakQty" & i) = 999999 Or rst("IDiscountMarkupPriceRate" & i + 1) = 0 Then
            i = 3
        Else
            price = price & " " & CStr(rst("IBreakQty" & i) + 1) & " " & FormatNumber(rst("IDiscountMarkupPriceRate" & i + 1) * (rst("IBreakQty" & i) + 1), 2, vbTrue, vbFalse, vbFalse)
        End If
    Next i
    
    'now do all rebates
    Dim rebateamt As String, isinstant As String
    rebateamt = SpecialsCache_GetRebateAmount(rst("ItemNumber"))
    isinstant = IIf(SpecialsCache_GetRebateInstantFlag(rst("ItemNumber")), "Yes", "No")
    
    stockEnum = getStockEnum(rst)
'XXX: temporarily keep oos around until we can migrate
    outofstock = IIf(stockEnumMeansOutOfStock(stockEnum), "Yes", "No")
    
    'Debug.Print rst("ItemNumber") & ": " & outofstock
    
    'outofstock = getOutOfStockStatus(rst)
    If rst("IsMASKit") Then
'XXX: temporarily keep oos around until we can migrate
        'correct the outofstock flag here
        'we don't update the lastchanged date though
        Dim laststockstatus As String
        laststockstatus = DLookup("OutOfStock", "PartNumbers", "ItemNumber='" & rst("ItemNumber") & "'")
        If (laststockstatus = True And outofstock = "No") Or _
           (laststockstatus = False And outofstock = "Yes") Then
            DB.Execute "UPDATE PartNumbers SET Changes=2, OutOfStock=" & IIf(outofstock = "Yes", "1", "0") & " WHERE ItemNumber='" & rst("ItemNumber") & "'"
        End If
        'Dim laststockstatus As Long
        'laststockstatus = DLookup("StockEnum", "PartNumbers", "ItemNumber='" & rst("ItemNumber") & "'")
        'If laststockstatus <> stockEnum Then
        '    DB.Execute "UPDATE PartNumbers SET Changes=2, StockEnum=" & stockEnum & " WHERE ItemNumber='" & rst("ItemNumber") & "'"
        'End If
    End If
    
    'main spec contents, everything special (group items, inserting specials, etc) is handled there
    Dim extspec As String
    extspec = WebOffloadFunctions.GenerateItemPageContents(rst)
    hlite1 = WebOffloadFunctions.GenerateHighlightByPosition(rst, 1)
    hlite2 = WebOffloadFunctions.GenerateHighlightByPosition(rst, 2)
    
    'section spec contents, handles group items, etc
    Dim partspecs As String
    partspecs = GenerateSectionPageSpecsList(rst)
    
    Dim groupedMasterItem As String
    groupedMasterItem = ""
    If GroupItemCache_IsGroupedSubItem(rst("ItemNumber")) Then
        groupedMasterItem = GroupItemCache_GetGroupYahooIDBySubItemNumber(rst("ItemNumber"))
    End If
    
    Dim groupTemplateBlock As String
    groupTemplateBlock = ""
    Dim groupTemplateItemCount As Long
    groupTemplateItemCount = 0
    If GroupItemCache_IsGroupMasterItem(rst("ItemNumber")) Then
        'get the group items, and build block
        Dim rstGroup As ADODB.Recordset
        Set rstGroup = DB.retrieve("SELECT ItemNumber, ColumnValue, dbo.CreateYahooID(ItemNumber) AS YahooID FROM vGroupedItemsActive WHERE GroupID=" & GroupItemCache_GetGroupDatabaseIDByMasterItemNumber(rst("ItemNumber")) & " ORDER BY SortOrder, ColumnIndex")
        Dim currentItemNumber As String
        currentItemNumber = ""
        Dim contentsArray As PerlArray
        Set contentsArray = New PerlArray
        While Not rstGroup.EOF
            If currentItemNumber = "" Then
                'initial loop, add the header line
                currentItemNumber = rstGroup("ItemNumber")
                groupTemplateBlock = IIf(GroupItemCache_GetGroupThumbnailSettingByMasterItemNumber(rst("ItemNumber")) = True, "IMAGE%", "") & _
                                     IIf(GroupItemCache_GetGroupAvailabilityColumnSettingByMasterItemNumber(rst("ItemNumber")) = False, "NOAVAILCOL%", "") & _
                                     "ITEMNUMBER" & "~" & IIf(rstGroup("ColumnValue") = "", "&nbsp;", rstGroup("ColumnValue"))
            ElseIf currentItemNumber = rstGroup("ItemNumber") Then
                'item column
                groupTemplateBlock = groupTemplateBlock & "~" & IIf(rstGroup("ColumnValue") = "", "&nbsp;", rstGroup("ColumnValue"))
                groupTemplateItemCount = groupTemplateItemCount + 1
            Else
                'new item and item column
                currentItemNumber = rstGroup("ItemNumber")
                groupTemplateBlock = groupTemplateBlock & vbCrLf & CreateYahooID(rstGroup("ItemNumber")) & "~" & IIf(rstGroup("ColumnValue") = "", "&nbsp;", rstGroup("ColumnValue"))
                contentsArray.Push rstGroup("YahooID")
            End If
            rstGroup.MoveNext
        Wend
        rstGroup.Close
        Set rstGroup = Nothing
        
        'contents so breadcrumbs work
        contents = contentsArray.Join(" ")
    End If
    
    Dim topsellers As String
    topsellers = ""
    If GroupItemCache_IsGroupMasterItem(rst("ItemNumber")) Then
        Dim rstGroupTop As ADODB.Recordset
        Set rstGroupTop = DB.retrieve("SELECT dbo.CreateYahooID(ItemNumber) FROM vGroupedItemsTopSellers WHERE GroupID=" & GroupItemCache_GetGroupDatabaseIDByMasterItemNumber(rst("ItemNumber")) & " ORDER BY Sales DESC, NEWID() DESC")
        Dim tops As PerlArray
        Set tops = New PerlArray
        While Not rstGroupTop.EOF
            tops.Push CStr(rstGroupTop(0))
            rstGroupTop.MoveNext
        Wend
        rstGroupTop.Close
        Set rstGroupTop = Nothing
        topsellers = tops.Join(" ")
    End If
    
    ' SHIPPING TYPE
    '   0 = standard
    '   1 = truck
    '   2 = truck 6.50 (deprecated)
    '   3 = free
    ' on the tools plus site, any item over $50 individually is considered free shipping, and needs
    ' to be switched to that if it isn't (but not if it's truck freight or common carrier). the
    ' "and shippingTypeID=0" will need to change if we add more shipping types that could change
    ' over to free.
    Dim shippingTypeID As Long
    shippingTypeID = CLng(rst("ShippingType"))
    If rst("WebsiteID") = 0 And shippingTypeID = 0 Then
        Dim basePrice As Long
        If CBool(InStr(price, " ")) Then
            basePrice = Left(price, InStr(price, " ") - 1)
        Else
            basePrice = price
        End If
        If basePrice >= 50 Then
            shippingTypeID = 3
        End If
    End If
    
    Dim pathNum As Long
    For pathNum = 0 To allWebPaths.Scalar - 1
        path = allWebPaths.Elem(pathNum)
        With New ADODB.Command
            .name = "spWebOffloadAddItem"
            .CommandText = "spWebOffloadAddItem"
            .CommandType = adCmdStoredProc
            .NamedParameters = True
            .ActiveConnection = DB.handle
            .Parameters.Append .CreateParameter("@offloadnum", adInteger, adParamInput, , CLng(1))
            .Parameters.Append .CreateParameter("@rlc", adVarChar, adParamInput, 4, CStr(rst("RealProductLine")))
            .Parameters.Append .CreateParameter("@realitem", adVarChar, adParamInput, 30, CStr(rst("ItemNumber")))
            .Parameters.Append .CreateParameter("@code", adVarChar, adParamInput, 100, CStr(newitemnumber))
            .Parameters.Append .CreateParameter("@itemname", adVarChar, adParamInput, 512, itemName)
            .Parameters.Append .CreateParameter("@prodline", adVarChar, adParamInput, 4, CStr(rst("ProductLine")))
            .Parameters.Append .CreateParameter("@plpage", adVarChar, adParamInput, 32, CStr(globalFoundProductLinePage))
            .Parameters.Append .CreateParameter("@partnum", adVarChar, adParamInput, 30, Mid(rst("ItemNumber"), 4))
            
            .Parameters.Append .CreateParameter("@path", adVarChar, adParamInput, 100, CStr(path))
            
            .Parameters.Append .CreateParameter("@manuf", adVarChar, adParamInput, 50, CStr(rst("ManufFullNameWeb")))
            .Parameters.Append .CreateParameter("@desc1", adVarChar, adParamInput, 512, desc1)
            .Parameters.Append .CreateParameter("@desc2", adVarChar, adParamInput, 512, desc2)
            .Parameters.Append .CreateParameter("@desc3", adVarChar, adParamInput, 512, desc3)
            .Parameters.Append .CreateParameter("@desc4", adVarChar, adParamInput, 512, desc4)
            
            .Parameters.Append .CreateParameter("@partspecs", adLongVarChar, adParamInput, IIf(Len(partspecs) = 0, 1, Len(partspecs)), partspecs)
            '.Parameters.Append .CreateParameter("@fullspecs", adLongVarChar, adParamInput, 1, "")
            .Parameters.Append .CreateParameter("@extspec", adLongVarChar, adParamInput, IIf(Len(extspec) = 0, 1, Len(extspec)), extspec)
            '.Parameters.Append .CreateParameter("@options", adLongVarChar, adParamInput, IIf(Len(options) = 0, 1, Len(options)), options)
            .Parameters.Append .CreateParameter("@relsects", adLongVarChar, adParamInput, IIf(Len(crossSellLinks) = 0, 1, Len(crossSellLinks)), crossSellLinks)
            .Parameters.Append .CreateParameter("@caption", adLongVarChar, adParamInput, IIf(Len(capt) = 0, 1, Len(capt)), capt)
            .Parameters.Append .CreateParameter("@contents", adLongVarChar, adParamInput, IIf(Len(contents) = 0, 1, Len(contents)), contents)
            
            .Parameters.Append .CreateParameter("@accessories", adLongVarChar, adParamInput, IIf(Len(Accessories) = 0, 1, Len(Accessories)), Accessories)
            .Parameters.Append .CreateParameter("@similaritems", adLongVarChar, adParamInput, IIf(Len(similarItems) = 0, 1, Len(similarItems)), similarItems)
            
            .Parameters.Append .CreateParameter("@groupedsubitems", adLongVarChar, adParamInput, IIf(Len(groupTemplateBlock) = 0, 1, Len(groupTemplateBlock)), groupTemplateBlock)
            .Parameters.Append .CreateParameter("@topsellers", adLongVarChar, adParamInput, IIf(Len(topsellers) = 0, 1, Len(topsellers)), topsellers)
            .Parameters.Append .CreateParameter("@groupedmasteritem", adLongVarChar, adParamInput, IIf(Len(groupedMasterItem) = 0, 1, Len(groupedMasterItem)), groupedMasterItem)

            .Parameters.Append .CreateParameter("@price", adVarChar, adParamInput, 50, price)
            .Parameters.Append .CreateParameter("@rebateamount", adVarChar, adParamInput, 10, rebateamt)
            .Parameters.Append .CreateParameter("@instantrebate", adVarChar, adParamInput, 3, isinstant)
            .Parameters.Append .CreateParameter("@saledisc", adInteger, adParamInput, , saleDisc)
            .Parameters.Append .CreateParameter("@mapp", adVarChar, adParamInput, 10, mappprice)
            
            .Parameters.Append .CreateParameter("@orderable", adVarChar, adParamInput, 3, orderable)
            .Parameters.Append .CreateParameter("@outofstock", adVarChar, adParamInput, 3, outofstock)
            .Parameters.Append .CreateParameter("@stockenum", adInteger, adParamInput, , stockEnum)
            .Parameters.Append .CreateParameter("@availability", adVarChar, adParamInput, 128, Availability)
            '.Parameters.Append .CreateParameter("@showurl", adVarChar, adParamInput, 255, showURL)
            '.Parameters.Append .CreateParameter("@showimage", adVarChar, adParamInput, 255, showImage)
            .Parameters.Append .CreateParameter("@sectionspecial", adLongVarChar, adParamInput, IIf(Len(sectionSpecialStanza) = 0, 1, Len(sectionSpecialStanza)), sectionSpecialStanza)
            '512 is also hard coded as length in the generate function, changes should be
            'made in both places
            .Parameters.Append .CreateParameter("@metakeywords", adVarChar, adParamInput, 512, metas)
            .Parameters.Append .CreateParameter("@req", adVarChar, adParamInput, 3, CStr(rst("ShowRequestForm")))
            .Parameters.Append .CreateParameter("@recon", adVarChar, adParamInput, 3, CStr(rst("IsReconditioned")))
            .Parameters.Append .CreateParameter("@shiptype", adInteger, adParamInput, 4, shippingTypeID)
            
            .Parameters.Append .CreateParameter("@issection", adVarChar, adParamInput, 3, "No")
            .Parameters.Append .CreateParameter("@isitem", adVarChar, adParamInput, 3, "Yes")
            
            .Parameters.Append .CreateParameter("@isnoindex", adVarChar, adParamInput, 3, IIf(CBool(rst("MetaNoIndex")), "Yes", "No"))
            
            .Parameters.Append .CreateParameter("@mobilename", adVarChar, adParamInput, IIf(Len(mobileName) = 0, 1, Len(mobileName)), mobileName)
            .Parameters.Append .CreateParameter("@hlite1", adVarChar, adParamInput, IIf(Len(hlite1) = 0, 1, Len(hlite1)), hlite1)
            .Parameters.Append .CreateParameter("@hlite2", adVarChar, adParamInput, IIf(Len(hlite2) = 0, 1, Len(hlite2)), hlite2)
            
            .Execute
        End With
    Next pathNum
    
    Set allWebPaths = Nothing
End Sub

'Private Sub cacheAcronyms()
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT Abbreviation, Explanation FROM Abbreviations")
'    Set abbrevList = New Dictionary
'    While Not rst.EOF
'        If Right(rst("Abbreviation"), 2) = "\." Then
'            abbrevList.Add "\b" & rst("Abbreviation") & "(\s)", CStr(rst("Explanation"))
'            abbrevList.Add "\b" & rst("Abbreviation") & "\z", CStr(rst("Explanation"))
'        Else
'            abbrevList.Add "\b" & rst("Abbreviation") & "\b", CStr(rst("Explanation"))
'        End If
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'End Sub

'Public Function getOutOfStockStatus(rst As ADODB.Recordset, Optional verifyQty As Long = -1) As String
'    'anything with an availability override gets handled there
'    ' 1 means always out
'    ' 0 means always in
'    ' -1 (or anything else) means skip, and calculate normally
'    'discontinued is never out of stock, breaks template
'    ' 2013-02-21: being d/c were always in stock, now changing to normal handling
'    'if it's a kit then check kit contents
'    'dropshippable and over BOBelow amount never out of stock
'    'otherwise, whatever it says
'    '
'    ' A VARIATION OF THIS LOGIC IS IN FEEDS.PM
'    '
'    Dim retval As String
'    If Not IsNull(rst("AvailabilityOverride")) And Nz(rst("OverrideOutOfStock")) = "1" Then
'        retval = "Yes"
'    ElseIf Not IsNull(rst("AvailabilityOverride")) And Nz(rst("OverrideOutOfStock")) = "0" Then
'        retval = "No"
'    'ElseIf IsItemDC(rst("ItemStatus")) And rst("QtyOnPO") = 0 Then
'    '    retval = "No"
'    ElseIf IsItemCompletelyDC(rst("ItemStatus")) Then
'        retval = "No"
'    ElseIf rst("IsMASKit") Then
'        retval = "Yes"
'        Dim kitlist As ADODB.Recordset
'        Set kitlist = DB.retrieve("SELECT ComponentItemCode, QuantityPerAssembly FROM IM_SalesKitDetail WHERE SalesKitNo='" & rst("ItemNumber") & "'")
'        While Not kitlist.EOF
'            Dim subItem As ADODB.Recordset
'            Set subItem = DB.retrieve("SELECT * FROM vWebOffloadBackend WHERE ItemNumber='" & kitlist("ComponentItemCode") & "'")
'            retval = getOutOfStockStatus(subItem, kitlist("QuantityPerAssembly"))
'            If retval = "Yes" Then
'                kitlist.MoveLast    'a component is out, so the kit is out
'                kitlist.MoveNext
'            Else
'                kitlist.MoveNext    'this component is in stock, keep checking the rest
'            End If
'            subItem.Close
'            Set subItem = Nothing
'        Wend
'        kitlist.Close
'        Set kitlist = Nothing
'    ElseIf IsItemDSable(rst("ItemStatus")) And (rst("StdCost") = 0 Or rst("StdCost") > rst("BOBelow")) Then
'        retval = "No"
'    Else
'        If verifyQty = -1 Then
'            'we don't care about the qty, just use the flag
'            retval = rst("OutOfStock")
'        Else
'            'check the qty avail against our verifyQty
'            'If (CLng(rst("QtyOnHand")) - CLng(rst("AvailabilityLimit"))) >= verifyQty Then
'            If Logic.YahooGetAvailableQuantityFor(rst("ItemNumber"), rst("QtyOnHand"), rst("QtyOnSO"), rst("AvailabilityLimit")) >= verifyQty Then
'                retval = "No"
'            Else
'                retval = "Yes"
'            End If
'        End If
'    End If
'    getOutOfStockStatus = retval
'End Function

Private Function getStockEnum(rst As ADODB.Recordset, Optional verifyQty As Long = 1) As Long
    'STOCK ENUM
    ' 0: out of stock
    ' 1: in stock
    ' 2: dropship only, in stock
    ' 3: dropship only, out of stock
    ' 4: dropship only, unknown
    ' 5: preorder
    ' 6: special order
    ' 7: built to order (TODO)
    ' 8: discontinued
    Dim retval As Long
    'anything with an availability override gets handled here
    ' 1 means always out
    ' 0 means always in
    ' -1 (or anything else) means skip, and calculate normally
    'discontinued is never out of stock, breaks template
    ' 2013-02-21: being d/c were always in stock, now changing to normal handling
    'if it's a kit then check kit contents
    'dropshippable and over BOBelow amount never out of stock
    'otherwise, whatever it says
    '
    ' A VARIATION OF THIS LOGIC IS IN FEEDS.PM
    '
    If Not IsNull(rst("AvailabilityOverride")) And Nz(rst("OverrideOutOfStock")) = "1" Then      'forced out of stock
        retval = 0
    ElseIf Not IsNull(rst("AvailabilityOverride")) And Nz(rst("OverrideOutOfStock")) = "0" Then  'forced into stock
        retval = 1
    ElseIf IsItemCompletelyDC(rst("ItemStatus")) Then                                            'discontinued
        retval = 8
    ElseIf rst("IsMASKit") Then
        retval = getStockEnumForKit(rst, verifyQty)
    ElseIf ItemStatusBitFlags.IsItemDSonly(rst("ItemStatus")) Then
        If IsNull(rst("VendorQuantityOnHand")) Then                                              'dropship only with no vendor link
            retval = 4
        ElseIf rst("VendorQuantityOnHand") - rst("AvailabilityLimit") > 0 Then                   'dropship only, with quantity per vendor
            retval = 2
        Else                                                                                     'dropship only, no quantity or below availability limit
            retval = 3
        End If
    ElseIf ItemStatusBitFlags.IsItemStocked(rst("ItemStatus")) = False Then                      'not dropshipped, not stocked, special order
        retval = 6
    ElseIf Logic.YahooGetAvailableQuantityFor( _
            rst("ItemNumber"), _
            rst("QtyOnHand"), _
            Nz(rst("VendorQuantityOnHand"), 0), _
            rst("QtyOnSO"), _
            rst("DSonly"), _
            rst("DSable"), _
            rst("AvailabilityLimit")) >= verifyQty Then                                          'in stock!
        retval = 1
    ElseIf IsNull(rst("LastReceiptDate")) Then                                                  'preorder
        retval = 5
    Else                                                                                         'none of the above, so out of stock
        retval = 0
    End If
    getStockEnum = retval
End Function

Private Function getStockEnumForKit(rst As ADODB.Recordset, verifyQty As Long) As Long
    Dim retval As Long
    retval = 1 'assume in stock
    Dim kitlist As ADODB.Recordset
    Set kitlist = DB.retrieve("SELECT ComponentItemCode, QuantityPerAssembly FROM IM_SalesKitDetail WHERE SalesKitNo='" & rst("ItemNumber") & "'")
    While Not kitlist.EOF
        Dim subItem As ADODB.Recordset
        Set subItem = DB.retrieve("SELECT * FROM vWebOffloadBackend WHERE ItemNumber='" & kitlist("ComponentItemCode") & "'")
        retval = getStockEnum(subItem, kitlist("QuantityPerAssembly") * verifyQty)
        If stockEnumMeansOutOfStock(retval) Then
            kitlist.MoveLast    'a component is out, so the kit is out
            kitlist.MoveNext
        Else
            kitlist.MoveNext    'this component is in stock, keep checking the rest
        End If
        subItem.Close
        Set subItem = Nothing
    Wend
    kitlist.Close
    Set kitlist = Nothing
    getStockEnumForKit = retval
End Function

Private Function stockEnumMeansInStock(stockEnum As Long) As Boolean
    'in, d/s in, d/s unknown, spec order, built to order
    stockEnumMeansInStock = stockEnum = 1 Or stockEnum = 2 Or stockEnum = 4 Or stockEnum = 6 Or stockEnum = 7
End Function

Private Function stockEnumMeansOutOfStock(stockEnum As Long) As Boolean
    'process of elimination: out, d/s out, prerder, d/c
    stockEnumMeansOutOfStock = Not stockEnumMeansInStock(stockEnum)
End Function



Private Sub btnViewLast_Click()
    'Unload RptWebOffload
    'Load RptWebOffload
    'Set RptWebOffload.DataSource = DB.retrieve("SELECT * FROM vWebOffloadReportSource ORDER BY Changes, Path")
    ''RptWebOffload.Orientation = rptOrientLandscape
    'RptWebOffload.TopMargin = 0.25
    'RptWebOffload.BottomMargin = 0.25
    'RptWebOffload.LeftMargin = 0.5
    'RptWebOffload.RightMargin = 0.5
    'RptWebOffload.Show
End Sub

Private Sub btnWebsiteSelection_Click()
    Dim newsite As Long
    newsite = CLng(Me.btnWebsiteSelection.tag) + 1
    If Not WebsiteNameHash.exists(newsite) Then
        newsite = 0
    End If
    Me.btnWebsiteSelection.tag = newsite
    Me.btnWebsiteSelection.Caption = "Website: " & WebsiteNameHash.item(newsite)
    UpdateWebOffloadChangedQuantity
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Me.btnWebsiteSelection.tag = CURRENT_WEBSITE_ID
    Me.btnWebsiteSelection.Caption = "Website: " & WebsiteNameHash.item(CURRENT_WEBSITE_ID)
    
    UpdateWebOffloadChangedQuantity
    
    checkEBayOffloadRunning
    
    Me.btnEBayOffload(3).Visible = CBool(GetCurrentUserNickname() = "Kirk")
End Sub

Private Sub jetOffloadBtn_Click()
JetDotCom.JetOffloadAll
End Sub

Private Sub opFilePath_Click(index As Integer)
    Select Case index
        Case Is = FILE_ROOT
            Me.PathToSave = ""
        Case Is = FILE_DESKTOP
            Me.PathToSave = ""
        Case Is = FILE_OTHER
            Dim fp As FolderPicker, foo As String
            Set fp = New FolderPicker
            fp.SetParent WebOffload.hWnd
            foo = fp.GetFolder()
            Set fp = Nothing
            If foo = "" Then
                opFilePath(FILE_ROOT) = True
            Else
                Me.PathToSave = foo
            End If
        Case Is = FILE_MYDOCS
            Me.PathToSave = ""
    End Select
End Sub



Private Function getPath() As String
    Select Case True
        Case Is = Me.opFilePath(FILE_ROOT)
            getPath = "c:\"
        Case Is = Me.opFilePath(FILE_DESKTOP)
            getPath = Environ("UserProfile") & "\Desktop\"
        Case Is = Me.opFilePath(FILE_OTHER)
            getPath = Me.PathToSave & IIf(Right(Me.PathToSave, 1) = "\", "", "\")
        Case Is = Me.opFilePath(FILE_MYDOCS)
            getPath = getMyDocsFolder() & "\"
    End Select
End Function

Private Function testWriteAccess(filepathname As String) As Boolean
On Error GoTo errh
openfile:
    Open filepathname For Append As #1
    Close #1
    testWriteAccess = True
    Exit Function
    
errh:
    If Err.Number = 70 Then 'permission denied
        MsgBox "Unable to open " & qq(filepathname) & ", you probably have the file open in Excel." & vbCrLf & vbCrLf & "Close Excel and click OK."
        Err.Clear
        GoTo openfile
    Else                    'don't care what the error is, but we can't open for write
        If GetCurrentUserID = 11 Then 'kmehlin
            MsgBox "PERMISSIONS ERROR DEBUGGING INFO:" & vbCrLf & vbCrLf & "Error #:" & Err.Number & vbCrLf & vbCrLf & Err.Description
        End If
        Err.Clear
        testWriteAccess = False
    End If
End Function

Private Function getMyDocsFolder() As String
    Dim path As String
    path = String(260, 0)
    Call SHGetFolderPath(0, &H5, 0, 0, path)
    getMyDocsFolder = Replace(path, Chr(0), "")
End Function

Private Sub addSpecialsPageItemLine()

    Dim capt As String
    capt = "<!-- NOTE: DON'T CHANGE THIS TEXT DIRECTLY, IT WILL BE OVERWRITTEN BY EACH OFFLOAD! -->" & vbCrLf
    capt = capt & "<h2 align=""center"">Rebates and Specials Headquarters</h2>" & vbCrLf
    capt = capt & "<p>Welcome to our Rebates and Specials Headquarters. Below you will find the latest promotions from the industries top manufacturers, along with <b>EXCLUSIVE</b> Tools Plus deals that you won't find anywhere else.</p>" & vbCrLf
    capt = capt & "<p>We know that everyone loves a great deal, so we make it our business to hunt down every special we can find and pass it on to you right away. Consider this your &quot;One Stop Shop&quot; for savings and check back regularly.</p>" & vbCrLf
    capt = capt & "<p>Specials are grouped by manufacturer, so scroll down to see the full list!</p>" & vbCrLf
    capt = capt & "<p><span class=""important"">Mail-in Rebates</span> come to you as a form that should be included with your shipment. In the event we mess up and forget to send it, don't panic. Either <a class=""important"" href=""javascript:window.open('http://p1.hostingprod.com/@tools-plus.com/contact_shipping.html', 'form', 'scrollbars=no,toolbar=no,status=no,width=500,height=320 top=20,left=20');void(0);"">email us</a> or call us here at 1-800-222-6133 with your information and your request, or go to that manufacturers web site and you can usually print a copy from them.</p>" & vbCrLf
    capt = capt & "<p><span class=""important""><strong>Note for International Customers</strong></span>: Most <strong>mail-in</strong> rebate offers are valid for US residents only. Please check the rebate details and forms closely before ordering.</p>" & vbCrLf
    capt = capt & "<p><span class=""important"">Instant Rebates</span> are rebates that are deducted at time of purchase. The price that appears in your shopping cart reflects the instant rebate.</p>" & vbCrLf
    
    Dim contents As String
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT DISTINCT dbo.CreateYahooID(ItemNumber) AS YahooID FROM vSpecialsActive WHERE ShowOnRebatePage=1 AND WebPublished=1 ORDER BY YahooID")
    While Not rst.EOF
        contents = contents & " " & rst("YahooID")
        rst.MoveNext
    Wend
    If contents <> "" Then
        contents = Mid(contents, 2)
    End If
    
    Dim rebatePathNames As Variant
    rebatePathNames = Array("deals")
    Dim iter As Variant
    For Each iter In rebatePathNames
        With New ADODB.Command
            .name = "spWebOffloadAddItem"
            .CommandText = "spWebOffloadAddItem"
            .CommandType = adCmdStoredProc
            .NamedParameters = True
            .ActiveConnection = DB.handle
            .Parameters.Append .CreateParameter("@offloadnum", adInteger, adParamInput, , 1)
            .Parameters.Append .CreateParameter("@rlc", adVarChar, adParamInput, 4, "")
            .Parameters.Append .CreateParameter("@realitem", adVarChar, adParamInput, 30, "")
            .Parameters.Append .CreateParameter("@code", adVarChar, adParamInput, 100, CStr(iter))
            .Parameters.Append .CreateParameter("@itemname", adVarChar, adParamInput, 512, "Rebates and Bonuses")
            .Parameters.Append .CreateParameter("@prodline", adVarChar, adParamInput, 4, "")
            .Parameters.Append .CreateParameter("@plpage", adVarChar, adParamInput, 32, "")
            .Parameters.Append .CreateParameter("@partnum", adVarChar, adParamInput, 30, "")
            
            .Parameters.Append .CreateParameter("@path", adVarChar, adParamInput, 100, "")
            
            .Parameters.Append .CreateParameter("@manuf", adVarChar, adParamInput, 50, "")
            .Parameters.Append .CreateParameter("@desc1", adVarChar, adParamInput, 512, "")
            .Parameters.Append .CreateParameter("@desc2", adVarChar, adParamInput, 512, "")
            .Parameters.Append .CreateParameter("@desc3", adVarChar, adParamInput, 512, "")
            .Parameters.Append .CreateParameter("@desc4", adVarChar, adParamInput, 512, "")
            
            .Parameters.Append .CreateParameter("@partspecs", adLongVarChar, adParamInput, 1, "")
            '.Parameters.Append .CreateParameter("@fullspecs", adLongVarChar, adParamInput, 1, "")
            .Parameters.Append .CreateParameter("@extspec", adLongVarChar, adParamInput, 1, "")
            '.Parameters.Append .CreateParameter("@options", adLongVarChar, adParamInput, 1, "")
            .Parameters.Append .CreateParameter("@relsects", adLongVarChar, adParamInput, 1, "")
            .Parameters.Append .CreateParameter("@caption", adLongVarChar, adParamInput, IIf(Len(capt) = 0, 1, Len(capt)), capt)
            .Parameters.Append .CreateParameter("@contents", adLongVarChar, adParamInput, IIf(Len(contents) = 0, 1, Len(contents)), contents)
            
            .Parameters.Append .CreateParameter("@accessories", adLongVarChar, adParamInput, 1, "")
            .Parameters.Append .CreateParameter("@similaritems", adLongVarChar, adParamInput, 1, "")
            
            .Parameters.Append .CreateParameter("@groupedsubitems", adLongVarChar, adParamInput, 1, "")
            .Parameters.Append .CreateParameter("@topsellers", adLongVarChar, adParamInput, 1, "")
            .Parameters.Append .CreateParameter("@groupedmasteritem", adLongVarChar, adParamInput, 1, "")
            
            .Parameters.Append .CreateParameter("@price", adVarChar, adParamInput, 50, "")
            .Parameters.Append .CreateParameter("@rebateamount", adVarChar, adParamInput, 10, "")
            .Parameters.Append .CreateParameter("@instantrebate", adVarChar, adParamInput, 3, "No")
            .Parameters.Append .CreateParameter("@saledisc", adInteger, adParamInput, , 0)
            .Parameters.Append .CreateParameter("@mapp", adVarChar, adParamInput, 10, "0")
            
            .Parameters.Append .CreateParameter("@orderable", adVarChar, adParamInput, 3, "No")
            .Parameters.Append .CreateParameter("@outofstock", adVarChar, adParamInput, 3, "No")
            .Parameters.Append .CreateParameter("@stockenum", adInteger, adParamInput, , 1)
            .Parameters.Append .CreateParameter("@availability", adVarChar, adParamInput, 128, "")
            '.Parameters.Append .CreateParameter("@showurl", adVarChar, adParamInput, 255, "")
            '.Parameters.Append .CreateParameter("@showimage", adVarChar, adParamInput, 255, "")
            .Parameters.Append .CreateParameter("@sectionspecial", adLongVarChar, adParamInput, 1, "")
            .Parameters.Append .CreateParameter("@metakeywords", adVarChar, adParamInput, 512, "")
            .Parameters.Append .CreateParameter("@req", adVarChar, adParamInput, 3, "No")
            .Parameters.Append .CreateParameter("@recon", adVarChar, adParamInput, 3, "No")
            .Parameters.Append .CreateParameter("@shiptype", adInteger, adParamInput, 4, 0)
            
            .Parameters.Append .CreateParameter("@issection", adVarChar, adParamInput, 3, "Yes")
            .Parameters.Append .CreateParameter("@isitem", adVarChar, adParamInput, 3, "No")
            
            .Parameters.Append .CreateParameter("@isnoindex", adVarChar, adParamInput, 3, "No")
            
            .Parameters.Append .CreateParameter("@mobilename", adVarChar, adParamInput, 1, "")
            .Parameters.Append .CreateParameter("@hlite1", adVarChar, adParamInput, 1, "")
            .Parameters.Append .CreateParameter("@hlite2", adVarChar, adParamInput, 1, "")
            
            .Execute
        End With
    Next iter
End Sub


Private Function getWebPathsForItem(item As String) As PerlArray
    Dim webpathstrings As PerlArray
    Set webpathstrings = New PerlArray
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT WebPathID FROM PartNumberWebPaths WHERE ItemNumber='" & item & "'")
    While Not rst.EOF
        Dim wpid As String, pathstr As String
        wpid = rst("WebPathID")
        If WebPathCache_GetPublishStatus(wpid) = 2 Then
            'skipme
            If Not globalPathProblems.exists(wpid) Then
                globalPathProblems.Add CStr(wpid), "skipping deleted path for " & item
            End If
        Else
            pathstr = ""
            While wpid <> "-1"
                pathstr = WebPathCache_GetURLPart(wpid) & IIf(pathstr = "", "", ":" & pathstr)
                If globalFoundProductLinePage = "" And WebPathCache_IsManufacturerPage(wpid) Then
                    globalFoundProductLinePage = WebPathCache_GetURLPart(wpid)
                End If
                Select Case WebPathCache_GetPublishStatus(wpid)
                    Case Is = 0
                        If Not globalPathProblems.exists(wpid) Then
                            globalPathProblems.Add CStr(wpid), "t/b published"
                        End If
                    Case Is = 1
                        'ok
                    Case Is = 2
                        If Not globalPathProblems.exists(wpid) Then
                            globalPathProblems.Add CStr(wpid), "deleted"
                        End If
                End Select
                wpid = WebPathCache_GetParent(wpid)
            Wend
            webpathstrings.Push pathstr
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    If webpathstrings.Scalar = 0 Then
        globalPathProblems.Add CStr(item), "item has no paths associated"
    End If
    
    Set getWebPathsForItem = webpathstrings
End Function
