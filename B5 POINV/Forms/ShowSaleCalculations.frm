VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#3.5#0"; "TPControls.ocx"
Begin VB.Form ShowSaleCalculations 
   Caption         =   "Show/Sale Calculations"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   9525
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   46
      Text            =   "ShowSaleCalculations.frx":0000
      Top             =   360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton btnMacroToolarama14 
      Caption         =   "Tool-A-Rama 14"
      Height          =   255
      Left            =   8100
      TabIndex        =   44
      Top             =   2640
      Width           =   1365
   End
   Begin VB.CommandButton btnMacroToolarama13 
      Caption         =   "Tool-A-Rama 13"
      Height          =   255
      Left            =   8100
      TabIndex        =   43
      Top             =   2370
      Width           =   1365
   End
   Begin VB.CommandButton btnMacroToolarama12 
      Caption         =   "Tool-A-Rama 12"
      Height          =   255
      Left            =   8100
      TabIndex        =   42
      Top             =   2100
      Width           =   1365
   End
   Begin VB.CommandButton btnMacroToolarama11 
      Caption         =   "Tool-A-Rama 11"
      Height          =   255
      Left            =   6720
      TabIndex        =   41
      Top             =   3540
      Width           =   1365
   End
   Begin VB.CommandButton btnMacroToolarama10 
      Caption         =   "Tool-A-Rama 10"
      Height          =   255
      Left            =   6720
      TabIndex        =   40
      Top             =   2910
      Width           =   1365
   End
   Begin VB.CommandButton btnMacroToolarama09 
      Caption         =   "Tool-A-Rama 09"
      Height          =   255
      Left            =   6720
      TabIndex        =   39
      Top             =   2640
      Width           =   1365
   End
   Begin VB.CheckBox chkIncludeVendorInSpreadsheet 
      Caption         =   "Incl. Vendor in XSheet"
      Height          =   255
      Left            =   2820
      TabIndex        =   38
      Top             =   2850
      Value           =   1  'Checked
      Width           =   2355
   End
   Begin VB.CommandButton btnMacroToolarama07 
      Caption         =   "Tool-A-Rama 07"
      Height          =   255
      Left            =   6720
      TabIndex        =   37
      Top             =   2100
      Width           =   1365
   End
   Begin VB.CommandButton btnImport 
      Caption         =   "DO IMPORT"
      Height          =   375
      Left            =   3930
      TabIndex        =   36
      Top             =   30
      Width           =   2625
   End
   Begin VB.CommandButton btnMacroToolarama08 
      Caption         =   "Tool-A-Rama 08"
      Height          =   255
      Left            =   6720
      TabIndex        =   35
      Top             =   2370
      Width           =   1365
   End
   Begin VB.CommandButton btnMacroNEWW08 
      Caption         =   "NE WW Expo 08"
      Height          =   255
      Left            =   5340
      TabIndex        =   34
      Top             =   2370
      Width           =   1365
   End
   Begin VB.CheckBox chkEnablePOSP 
      Caption         =   "Point of Sale Professional"
      Height          =   225
      Left            =   1380
      TabIndex        =   20
      Top             =   3690
      Value           =   1  'Checked
      Width           =   2115
   End
   Begin VB.CheckBox chkEnableAR 
      Caption         =   "Completed Invoices"
      Height          =   255
      Left            =   1350
      TabIndex        =   19
      Top             =   4980
      Value           =   1  'Checked
      Width           =   1725
   End
   Begin VB.CheckBox chkEnableSO 
      Caption         =   "Open Sales Orders"
      Height          =   225
      Left            =   1350
      TabIndex        =   18
      Top             =   6120
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin VB.CheckBox chkShowTotals 
      Caption         =   "Calculate Totals"
      Height          =   255
      Left            =   2820
      TabIndex        =   14
      Top             =   2550
      Value           =   1  'Checked
      Width           =   2355
   End
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   240
      TabIndex        =   11
      Top             =   2130
      Width           =   2295
      Begin VB.OptionButton opMetric 
         Caption         =   "By Sphere 1 Category"
         Height          =   225
         Index           =   3
         Left            =   90
         TabIndex        =   45
         Top             =   980
         Width           =   2055
      End
      Begin VB.OptionButton opMetric 
         Caption         =   "Individual Sales Orders"
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   15
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton opMetric 
         Caption         =   "Quantities By Item"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   13
         Top             =   450
         Width           =   1875
      End
      Begin VB.OptionButton opMetric 
         Caption         =   "Sales $ By Line"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Top             =   180
         Value           =   -1  'True
         Width           =   1875
      End
   End
   Begin VB.CheckBox chkIncludeSpecial 
      Caption         =   "Include XXX/ZZZ"
      Height          =   255
      Left            =   2820
      TabIndex        =   10
      Top             =   2250
      Value           =   1  'Checked
      Width           =   2355
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   405
      Left            =   8190
      TabIndex        =   9
      Top             =   0
      Width           =   1095
   End
   Begin VB.Frame frameSO 
      Height          =   1095
      Left            =   1290
      TabIndex        =   8
      Top             =   6090
      Width           =   2385
      Begin TPControls.DateDropdown SOToDate 
         Height          =   315
         Left            =   540
         TabIndex        =   30
         Top             =   660
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
      End
      Begin TPControls.DateDropdown SOFromDate 
         Height          =   315
         Left            =   540
         TabIndex        =   29
         Top             =   300
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
      End
      Begin VB.Label generalLabel 
         Caption         =   "To:"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   24
         Top             =   690
         Width           =   495
      End
      Begin VB.Label generalLabel 
         Caption         =   "From:"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   21
         Top             =   330
         Width           =   495
      End
   End
   Begin VB.Frame frameAR 
      Height          =   1035
      Left            =   1290
      TabIndex        =   7
      Top             =   4980
      Width           =   2385
      Begin TPControls.DateDropdown ARToDate 
         Height          =   315
         Left            =   540
         TabIndex        =   28
         Top             =   660
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
      End
      Begin TPControls.DateDropdown ARFromDate 
         Height          =   315
         Left            =   540
         TabIndex        =   27
         Top             =   300
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
      End
      Begin VB.Label generalLabel 
         Caption         =   "To:"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   23
         Top             =   690
         Width           =   495
      End
      Begin VB.Label generalLabel 
         Caption         =   "From:"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   17
         Top             =   330
         Width           =   495
      End
   End
   Begin VB.ListBox Lines 
      Height          =   3960
      Left            =   90
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   3810
      Width           =   1065
   End
   Begin VB.ListBox Results 
      Height          =   3570
      Left            =   3780
      TabIndex        =   5
      Top             =   3810
      Width           =   5715
   End
   Begin VB.CommandButton btnCalculate 
      Caption         =   "Calculate!"
      Height          =   495
      Left            =   1710
      TabIndex        =   4
      Top             =   7290
      Width           =   1605
   End
   Begin VB.CommandButton btnOpenInExcel 
      Caption         =   "Open In Excel"
      Enabled         =   0   'False
      Height          =   405
      Left            =   7800
      TabIndex        =   3
      Top             =   7440
      Width           =   1725
   End
   Begin VB.Frame framePOSP 
      Height          =   1095
      Left            =   1290
      TabIndex        =   2
      Top             =   3810
      Width           =   2385
      Begin TPControls.DateDropdown POSPToDate 
         Height          =   315
         Left            =   540
         TabIndex        =   26
         Top             =   660
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
      End
      Begin TPControls.DateDropdown POSPFromDate 
         Height          =   315
         Left            =   540
         TabIndex        =   25
         Top             =   300
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
      End
      Begin VB.Label generalLabel 
         Caption         =   "To:"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   22
         Top             =   690
         Width           =   495
      End
      Begin VB.Label generalLabel 
         Caption         =   "From:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   330
         Width           =   495
      End
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status:"
      Height          =   255
      Left            =   3600
      TabIndex        =   47
      Top             =   7560
      Width           =   4215
   End
   Begin VB.Label Label5 
      Caption         =   "Open Sales Orders are all phone and internet orders that haven't been shipped yet. The date is when the sales order was opened."
      Height          =   225
      Left            =   60
      TabIndex        =   33
      Top             =   1800
      Width           =   9195
   End
   Begin VB.Label Label4 
      Caption         =   $"ShowSaleCalculations.frx":0006
      Height          =   435
      Left            =   60
      TabIndex        =   32
      Top             =   1320
      Width           =   9195
   End
   Begin VB.Label Label3 
      Caption         =   "POSP includes all front counter transactions, including orders and special orders. The date is the actual date of the transaction."
      Height          =   255
      Left            =   60
      TabIndex        =   31
      Top             =   1020
      Width           =   9195
   End
   Begin VB.Label Label2 
      Caption         =   $"ShowSaleCalculations.frx":00B9
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   450
      Width           =   9135
   End
   Begin VB.Label Label1 
      Caption         =   "Show/Sale Calculations"
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
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   3045
   End
End
Attribute VB_Name = "ShowSaleCalculations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const METRIC_SALES As Long = 0
Private Const METRIC_QTYS As Long = 1
Private Const METRIC_SO As Long = 2

Private Type SphereCatVendorItem
    ItemNumber As String
    PrimaryVendor As String
    CategoryID As String
End Type



Private Sub btnImport_Click()
    If RunningThroughVPN Then
        MsgBox "Don't run this through the vpn!!"
        Exit Sub
    End If
    
    If vbNo = MsgBox("You should probably run the perl import, this one is 90% likely to crash because Mas sucks. Continue?", vbYesNo + vbDefaultButton2) Then
        Exit Sub
    End If
    
    Mouse.Hourglass True
    
    Dim m As Mas200ImportExport
    Set m = New Mas200ImportExport
    m.ImportShowSaleData
    MsgBox "DONE!!!"
    
    Mouse.Hourglass False
End Sub

Private Sub selectList(linesA As Variant)
    Dim thisline As Variant, i As Long, found As Boolean
    For i = 0 To Me.Lines.ListCount - 1
        Me.Lines.Selected(i) = False
    Next i
    For Each thisline In linesA
        found = False
        For i = 0 To Me.Lines.ListCount - 1
            If Me.Lines.list(i) = thisline Then
                Me.Lines.Selected(i) = True
                found = True
                i = Me.Lines.ListCount
            End If
        Next i
        If Not found Then
            MsgBox "Couldn't find LC " & thisline & " in list!"
        End If
    Next thisline
End Sub

Private Sub btnCalculate_Click()

If Me.opMetric(3).value = True Then
    BySphereOneCategoryV2
Exit Sub
End If

    
    
    If Me.Lines.SelCount = 0 Then
        MsgBox "You didn't select any line codes!"
        Exit Sub
    End If
    Mouse.Hourglass True
    ReDim sellines(Me.Lines.SelCount - 1) As String
    sellines = ListBoxAsArray(Me.Lines, True)
    Dim sql As String
    Dim rst As ADODB.Recordset
    
    Select Case True
        Case Is = Me.opMetric(METRIC_SO)
            Me.Results.Clear
            Me.Results.AddItem "Order" & vbTab & "Item" & vbTab & "Quantity" & vbTab & "Subtotal"
            If Me.chkEnableAR Then
                Set rst = DB.retrieve(buildSQLString("AR", sellines))
                While Not rst.EOF
                    Me.Results.AddItem rst("OrderNumber") & vbTab & rst("ItemNumber") & vbTab & rst("Quantity") & vbTab & rst("Extension")
                    rst.MoveNext
                Wend
                rst.Close
            End If
            If Me.chkEnableSO Then
                Set rst = DB.retrieve(buildSQLString("SO", sellines))
                While Not rst.EOF
                    Me.Results.AddItem rst("OrderNumber") & vbTab & rst("ItemNumber") & vbTab & rst("Quantity") & vbTab & rst("Extension")
                    rst.MoveNext
                Wend
                rst.Close
            End If
            If Me.chkEnablePOSP Then
                Set rst = DB.retrieve(buildSQLString("POSP", sellines))
                While Not rst.EOF
                    Me.Results.AddItem rst("OrderNumber") & vbTab & rst("ItemNumber") & vbTab & rst("Quantity") & vbTab & rst("Extension")
                    rst.MoveNext
                Wend
                rst.Close
            End If
            If Me.chkShowTotals Then
                MsgBox "Totals not implemented for View By SO"
            End If
            Me.btnOpenInExcel.Enabled = True
        Case Else
            Dim selstring As String
            selstring = "PLorItem"
            DB.Execute "TRUNCATE TABLE ShowSaleCalculations"
            If Me.chkEnablePOSP Then
                selstring = selstring & ", POSPSales AS 'Point of Sale'"
                sql = buildSQLString("POSP", sellines)
                DB.Execute sql
            End If
            If Me.chkEnableAR Then
                selstring = selstring & ", ARSales AS Invoiced"
                sql = buildSQLString("AR", sellines)
                DB.Execute sql
            End If
            If Me.chkEnableSO Then
                selstring = selstring & ", SOSales AS OpenOrders"
                sql = buildSQLString("SO", sellines)
                DB.Execute sql
            End If
    
            Set rst = DB.retrieve("SELECT DISTINCT " & selstring & " FROM vShowSaleCalcsDenormalized ORDER BY PLorItem")
            Me.Results.Clear
            Dim line As String
            Dim i As Long
            line = getHeaderName()
            For i = 1 To rst.fields.count - 1
                line = line & vbTab & rst.fields(i).name
            Next i
            Me.Results.AddItem line
            ReDim totals(rst.fields.count - 2) As Double
            While Not rst.EOF
                line = rst("PLorItem")
                For i = 1 To rst.fields.count - 1
                    line = line & vbTab & formatForDisplay(Nz(rst(i), "0"))
                    totals(i - 1) = totals(i - 1) + Nz(rst(i), "0")
                Next i
                Me.Results.AddItem line
                rst.MoveNext
            Wend
            rst.Close
            Set rst = Nothing
            If Me.chkShowTotals Then
                Me.Results.AddItem ""
                line = "Totals"
                Dim grandtotal As Double
                For i = LBound(totals) To UBound(totals)
                    line = line & vbTab & formatForDisplay(CStr(totals(i)))
                    grandtotal = grandtotal + totals(i)
                Next i
                Me.Results.AddItem line
                Me.Results.AddItem ""
                Me.Results.AddItem "Overall" & vbTab & formatForDisplay(CStr(grandtotal))
            End If
            Me.btnOpenInExcel.Enabled = True
    End Select
    
    setTabs
    
    Mouse.Hourglass False
End Sub

Private Sub BySphereOneCategoryV2()
    Dim showWebPriceInsteadOfStdPrice As Boolean
    showWebPriceInsteadOfStdPrice = True
    MsgBox "This version calcs the previous year's sales on Sales Orders. Using the " & IIf(showWebPriceInsteadOfStdPrice = True, "WebPrice", "StdPrice") & " to calculate spreadsheet."
    Results.Clear
    Dim IncludeSpecials As Boolean
    Dim ShowTotals As Boolean
    Dim IncludeVendorInSpreadsheet As Boolean
    Dim EnablePOSP As Boolean
    Dim EnableAR As Boolean
    Dim EnableSO As Boolean
    Dim POStartDate As Date
    Dim POEndDate As Date
    Dim ARStartDate As Date
    Dim AREndDate As Date
    Dim SOStartDate As Date
    Dim SOEndDate As Date
    
    IncludeSpecials = Me.chkIncludeSpecial.value
    ShowTotals = Me.chkShowTotals.value
    IncludeVendorInSpreadsheet = Me.chkIncludeVendorInSpreadsheet.value
    EnablePOSP = Me.chkEnablePOSP.value
    EnableAR = Me.chkEnableAR.value
    EnableSO = Me.chkEnableSO.value
    POStartDate = Me.POSPFromDate.value
    POEndDate = Me.POSPToDate.value
    ARStartDate = Me.ARFromDate.value
    AREndDate = Me.ARToDate.value
    SOStartDate = Me.SOFromDate.value
    SOEndDate = Me.SOToDate.value
    
    Dim tLines() As String
    Dim tLineTotal() As Currency
    ReDim tLines(Me.Lines.SelCount - 1) As String
    ReDim tLineTotal(Me.Lines.SelCount - 1) As Currency
    tLines = ListBoxAsArray(Me.Lines, True)
    
    Dim SphereCatList() As String
    Dim sqlCatResults As ADODB.Recordset
    'The first set line below gets a list of only active categories, although faster, not what we want
    'Set sqlCatResults = DB.retrieve("SELECT DISTINCT ConnectorID FROM MasterCategoryConnectors WHERE ConnectorType=4")
    Set sqlCatResults = DB.retrieve("SELECT DISTINCT ID FROM SphereOneCategories ORDER BY ID ASC")
    ReDim SphereCatList(sqlCatResults.RecordCount - 1) As String
    Dim count As Integer
    Do
        SphereCatList(count) = sqlCatResults("ID")
        sqlCatResults.MoveNext
        count = count + 1
    Loop Until sqlCatResults.EOF = True Or sqlCatResults.BOF = True
    
    Dim SOTable, ARTable, POSPTable As String
    SOTable = "vShowSaleCalcSO"
    ARTable = "vShowSaleCalcAR"
    POSPTable = "vShowSaleCalcPOSP"
    
    Dim query As String
    query = "SELECT SUM(StdCost * (SOSalesAmount + P2SalesAmount)) AS 'StdTotal' ,SUM(IDiscountMarkupPriceRate1 * (SOSalesAmount + P2SalesAmount)) AS 'WebTotal' FROM " & _
            "(" & _
            "SELECT InventoryStatistics.ItemNumber,P2SalesAmount,SOSalesAmount,ReturnsAmount,DropshipsAmount, StdCost,ListPrice,IDiscountMarkupPriceRate1,EBayPrice,DiscountMarkupPriceRate1 " & _
            "FROM InventoryStatistics " & _
            "INNER JOIN InventoryMaster ON InventoryMaster.ItemNumber=InventoryStatistics.ItemNumber " & _
            "WHERE InventoryStatistics.ItemNumber LIKE '{LINECODE}%' And PeriodType=0 AND Period=5 " & _
            "AND InventoryMaster.ItemNumber IN " & _
            "(" & _
            "SELECT InventoryMaster.ItemNumber FROM InventoryMaster " & _
            "INNER JOIN PartNumberWebPaths ON PartNumberWebPaths.ItemNumber=InventoryMaster.ItemNumber " & _
            "INNER JOIN WebPaths ON WebPaths.ID = PartNumberWebPaths.WebPathID " & _
            "INNER JOIN MasterCategoryConnectors ON MasterCategoryConnectors.MasterID = PartNumberWebPaths.WebPathID " & _
            "INNER JOIN SphereOneCategories ON SphereOneCategories.ID = MasterCategoryConnectors.ConnectorID " & _
            "WHERE MasterCategoryConnectors.ConnectorType=4 AND InventoryMaster.ItemNumber LIKE '{LINECODE}%' AND SphereOneCategories.ID={S1ID} AND IDiscountMarkupPriceRate1 <> 99999.99 " & _
            ")" & _
            ") rabble"
       
        Dim FirstPlaceLine() As String
        Dim FirstPlaceAmount() As Currency
        Dim SecondPlaceLine() As String
        Dim SecondPlaceAmount() As Currency
        Dim ThirdPlaceLine() As String
        Dim ThirdPlaceAmount() As Currency
        Dim StdCostTotal() As Currency
        Dim WebCostTotal() As Currency
        ReDim FirstPlaceLine(sqlCatResults.RecordCount - 1) As String
        ReDim FirstPlaceAmount(sqlCatResults.RecordCount - 1) As Currency
        ReDim SecondPlaceLine(sqlCatResults.RecordCount - 1) As String
        ReDim SecondPlaceAmount(sqlCatResults.RecordCount - 1) As Currency
        ReDim ThirdPlaceLine(sqlCatResults.RecordCount - 1) As String
        ReDim ThirdPlaceAmount(sqlCatResults.RecordCount - 1) As Currency
        ReDim StdCostTotal(sqlCatResults.RecordCount - 1) As Currency
        ReDim WebCostTotal(sqlCatResults.RecordCount - 1) As Currency
        Dim cat As Variant
        Dim catCount As Integer
        Dim prevStatus As String
        Text1.Text = query
        For Each cat In SphereCatList
            lblStatus.Caption = "Status: Working on category #" & CStr(cat)
            lblStatus.Refresh
            DoEvents
            prevStatus = lblStatus.Caption & " "
            Dim ln As Variant
                Dim soQuery As String
                soQuery = query
                soQuery = Replace(soQuery, "{TABLENAME}", SOTable)
                soQuery = Replace(soQuery, "{STARTDATE}", SOStartDate)
                soQuery = Replace(soQuery, "{ENDDATE}", SOEndDate)
                For Each ln In tLines
                Dim stdTot As Currency
                Dim webTot As Currency
                    Dim lineCatResults As ADODB.Recordset
                    Dim qStr As String
                    qStr = Replace(Replace(soQuery, "{LINECODE}", CStr(ln)), "{S1ID}", CStr(cat))
                    Set lineCatResults = DB.retrieve(qStr)
                    If IsNull(lineCatResults("StdTotal")) = False Then
                        stdTot = CDec(lineCatResults("StdTotal"))
                        StdCostTotal(catCount) = StdCostTotal(catCount) + stdTot
                        If showWebPriceInsteadOfStdPrice = False Then
                            If stdTot > FirstPlaceAmount(catCount) Then
                                FirstPlaceLine(catCount) = CStr(ln)
                                FirstPlaceAmount(catCount) = CDec(stdTot)
                            Else
                                If stdTot > SecondPlaceAmount(catCount) Then
                                    SecondPlaceLine(catCount) = CStr(ln)
                                    SecondPlaceAmount(catCount) = CDec(stdTot)
                                Else
                                    If stdTot > ThirdPlaceAmount(catCount) Then
                                        ThirdPlaceLine(catCount) = CStr(ln)
                                        ThirdPlaceAmount(catCount) = CDec(stdTot)
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If IsNull(lineCatResults("WebTotal")) = False Then
                        webTot = CDec(lineCatResults("WebTotal"))
                        WebCostTotal(catCount) = WebCostTotal(catCount) + webTot
                        If showWebPriceInsteadOfStdPrice = True Then
                            If webTot > FirstPlaceAmount(catCount) Then
                                FirstPlaceLine(catCount) = CStr(ln)
                                FirstPlaceAmount(catCount) = CDec(webTot)
                            Else
                                If webTot > SecondPlaceAmount(catCount) Then
                                    SecondPlaceLine(catCount) = CStr(ln)
                                    SecondPlaceAmount(catCount) = CDec(webTot)
                                Else
                                    If webTot > ThirdPlaceAmount(catCount) Then
                                        ThirdPlaceLine(catCount) = CStr(ln)
                                        ThirdPlaceAmount(catCount) = CDec(webTot)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next ln
            catCount = catCount + 1
        Next cat
        lblStatus.Caption = "Status: Writing Data... Please Wait."
        lblStatus.Refresh
        Dim testCount As Integer
        testCount = 0
        Results.AddItem "Category" & vbTab & "1st Place" & vbTab & "1st Amount" & vbTab & "2nd Place" & vbTab & "2nd Amount" & vbTab & "3rd Place" & vbTab & "3rd Amount" & vbTab & "Total Std" & vbTab & "Total Web"
        Do
            Dim catName As ADODB.Recordset
            Set catName = DB.retrieve("SELECT Name FROM SphereOneCategories WHERE ID=" & CStr(CInt(testCount + 1)))
            Dim cName As String
            cName = catName("Name")
            Dim line As String
            line = cName & vbTab & CStr(FirstPlaceLine(testCount)) & vbTab & "$" & CStr(FirstPlaceAmount(testCount)) & vbTab & CStr(SecondPlaceLine(testCount)) & vbTab & "$" & CStr(SecondPlaceAmount(testCount)) & vbTab & CStr(ThirdPlaceLine(testCount)) & vbTab & CStr(ThirdPlaceAmount(testCount)) & vbTab & "$" & CStr(StdCostTotal(testCount)) & vbTab & "$" & CStr(WebCostTotal(testCount))
            Results.AddItem line
            testCount = testCount + 1
        Loop Until testCount = UBound(FirstPlaceLine)
        Me.btnOpenInExcel.Enabled = True
        btnOpenInExcel_Click
    
    
    
End Sub

Private Sub BySphereOneCategory()
    Results.Clear
    Dim IncludeSpecials As Boolean
    Dim ShowTotals As Boolean
    Dim IncludeVendorInSpreadsheet As Boolean
    Dim EnablePOSP As Boolean
    Dim EnableAR As Boolean
    Dim EnableSO As Boolean
    Dim POStartDate As Date
    Dim POEndDate As Date
    Dim ARStartDate As Date
    Dim AREndDate As Date
    Dim SOStartDate As Date
    Dim SOEndDate As Date
    
    IncludeSpecials = Me.chkIncludeSpecial.value
    ShowTotals = Me.chkShowTotals.value
    IncludeVendorInSpreadsheet = Me.chkIncludeVendorInSpreadsheet.value
    EnablePOSP = Me.chkEnablePOSP.value
    EnableAR = Me.chkEnableAR.value
    EnableSO = Me.chkEnableSO.value
    POStartDate = Me.POSPFromDate.value
    POEndDate = Me.POSPToDate.value
    ARStartDate = Me.ARFromDate.value
    AREndDate = Me.ARToDate.value
    SOStartDate = Me.SOFromDate.value
    SOEndDate = Me.SOToDate.value
    
    Dim tLines() As String
    Dim tLineTotal() As Currency
    ReDim tLines(Me.Lines.SelCount - 1) As String
    ReDim tLineTotal(Me.Lines.SelCount - 1) As Currency
    tLines = ListBoxAsArray(Me.Lines, True)
    
    Dim SphereCatList() As String
    Dim sqlCatResults As ADODB.Recordset
    'The first set line below gets a list of only active categories, although faster, not what we want
    'Set sqlCatResults = DB.retrieve("SELECT DISTINCT ConnectorID FROM MasterCategoryConnectors WHERE ConnectorType=4")
    Set sqlCatResults = DB.retrieve("SELECT DISTINCT ID FROM SphereOneCategories ORDER BY ID ASC")
    ReDim SphereCatList(sqlCatResults.RecordCount - 1) As String
    Dim count As Integer
    Do
        SphereCatList(count) = sqlCatResults("ID")
        sqlCatResults.MoveNext
        count = count + 1
    Loop Until sqlCatResults.EOF = True Or sqlCatResults.BOF = True
    
    Dim SOTable, ARTable, POSPTable As String
    SOTable = "vShowSaleCalcSO"
    ARTable = "vShowSaleCalcAR"
    POSPTable = "vShowSaleCalcPOSP"
    
    Dim query As String
    query = "SELECT SUM(StdTotal) AS 'StdTotal',SUM(WebTotal) AS 'WebTotal' FROM " & _
    "(SELECT DISTINCT SphereOneCategories.ID,ProductLine.PrimaryVendorNumber,InventoryMaster.ItemNumber," & _
    "{TABLENAME}.Quantity,InventoryMaster.StdCost,InventoryMaster.IDiscountMarkupPriceRate1," & _
    "InventoryMaster.StdCost * {TABLENAME}.Quantity AS 'StdTotal'," & _
    "InventoryMaster.IDiscountMarkupPriceRate1 * {TABLENAME}.Quantity AS 'WebTotal' " & _
    "FROM InventoryMaster " & _
    "INNER JOIN PartNumberWebPaths ON PartNumberWebPaths.ItemNumber=InventoryMaster.ItemNumber " & _
    "INNER JOIN WebPaths ON WebPaths.ID = PartNumberWebPaths.WebPathID " & _
    "INNER JOIN MasterCategoryConnectors ON MasterCategoryConnectors.MasterID = PartNumberWebPaths.WebPathID " & _
    "INNER JOIN SphereOneCategories ON SphereOneCategories.ID = MasterCategoryConnectors.ConnectorID " & _
    "INNER JOIN {TABLENAME} ON {TABLENAME}.ItemNumber=InventoryMaster.ItemNumber " & _
    "INNER JOIN ProductLine ON InventoryMaster.ProductLine=ProductLine.ProductLine " & _
    "WHERE MasterCategoryConnectors.ConnectorType=4 " & _
    "AND {TABLENAME}.ProductLine IN (SELECT ProductLine FROM ProductLine WHERE PrimaryVendorNumber IN (SELECT PrimaryVendorNumber FROM ProductLine WHERE ProductLine LIKE '{LINECODE}')) " & _
    "AND {TABLENAME}.TransactionDate BETWEEN '{ENDDATE}' AND '{STARTDATE}' " & _
    "AND SphereOneCategories.ID={S1ID}) rabble"
       
        Dim FirstPlaceLine() As String
        Dim FirstPlaceAmount() As Currency
        Dim SecondPlaceLine() As String
        Dim SecondPlaceAmount() As Currency
        Dim ThirdPlaceLine() As String
        Dim ThirdPlaceAmount() As Currency
        Dim StdCostTotal() As Currency
        Dim WebCostTotal() As Currency
        ReDim FirstPlaceLine(sqlCatResults.RecordCount - 1) As String
        ReDim FirstPlaceAmount(sqlCatResults.RecordCount - 1) As Currency
        ReDim SecondPlaceLine(sqlCatResults.RecordCount - 1) As String
        ReDim SecondPlaceAmount(sqlCatResults.RecordCount - 1) As Currency
        ReDim ThirdPlaceLine(sqlCatResults.RecordCount - 1) As String
        ReDim ThirdPlaceAmount(sqlCatResults.RecordCount - 1) As Currency
        ReDim StdCostTotal(sqlCatResults.RecordCount - 1) As Currency
        ReDim WebCostTotal(sqlCatResults.RecordCount - 1) As Currency
        Dim cat As Variant
        Dim catCount As Integer
        Dim prevStatus As String
        Text1.Text = query
        For Each cat In SphereCatList
            lblStatus.Caption = "Status: Working on category #" & CStr(cat)
            lblStatus.Refresh
            DoEvents
            prevStatus = lblStatus.Caption & " "
            Dim ln As Variant
            If Me.chkEnableSO = vbChecked Then
                Dim soQuery As String
                soQuery = query
                soQuery = Replace(soQuery, "{TABLENAME}", SOTable)
                soQuery = Replace(soQuery, "{STARTDATE}", SOStartDate)
                soQuery = Replace(soQuery, "{ENDDATE}", SOEndDate)
                'Dim prevStatus As String
                
                Dim count2 As Integer
                For Each ln In tLines
                'count2 = count2 + 1
                lblStatus.Caption = prevStatus & "(SO#" & CStr(count) & "/" & CStr(UBound(tLines)) & ")"
                lblStatus.Refresh
                Dim stdTot As Currency
                Dim webTot As Currency
                    Dim lineCatResults As ADODB.Recordset
                    Dim qStr As String
                    qStr = Replace(Replace(soQuery, "{LINECODE}", CStr(ln)), "{S1ID}", CStr(cat))
                    Set lineCatResults = DB.retrieve(qStr)
                    If IsNull(lineCatResults("StdTotal")) = False Then
                        stdTot = CDec(lineCatResults("StdTotal"))
                        StdCostTotal(catCount) = StdCostTotal(catCount) + stdTot
                        
                    End If
                    If IsNull(lineCatResults("WebTotal")) = False Then
                        webTot = CDec(lineCatResults("WebTotal"))
                        WebCostTotal(catCount) = WebCostTotal(catCount) + webTot
                        If webTot > FirstPlaceAmount(catCount) Then
                            FirstPlaceLine(catCount) = CStr(ln)
                            FirstPlaceAmount(catCount) = CDec(webTot)
                        Else
                            If webTot > SecondPlaceAmount(catCount) Then
                                SecondPlaceLine(catCount) = CStr(ln)
                                SecondPlaceAmount(catCount) = CDec(webTot)
                            Else
                                If webTot > ThirdPlaceAmount(catCount) Then
                                    ThirdPlaceLine(catCount) = CStr(ln)
                                    ThirdPlaceAmount(catCount) = CDec(webTot)
                                End If
                            End If
                        End If
                    End If
                Next ln
            End If
            If Me.chkEnableAR = vbChecked Then
                Dim arQuery As String
                arQuery = query
                arQuery = Replace(arQuery, "{STARTDATE}", ARStartDate)
                arQuery = Replace(arQuery, "{ENDDATE}", AREndDate)
                arQuery = Replace(arQuery, "{TABLENAME}", ARTable)
                
               
                Dim count3 As Integer
                For Each ln In tLines
                'count3 = count3 + 1
                lblStatus.Caption = prevStatus & "(AR#" & CStr(count) & "/" & CStr(UBound(tLines)) & ")"
                lblStatus.Refresh
                Dim stdTot2 As Currency
                Dim webTot2 As Currency
                    Dim lineCatResults2 As ADODB.Recordset
                    Dim qStr2 As String
                    qStr2 = Replace(Replace(arQuery, "{LINECODE}", CStr(ln)), "{S1ID}", CStr(cat))
                    Set lineCatResults2 = DB.retrieve(qStr2)
                    If IsNull(lineCatResults2("StdTotal")) = False Then
                        stdTot2 = CDec(lineCatResults2("StdTotal"))
                        StdCostTotal(catCount) = StdCostTotal(catCount) + stdTot2
                        
                    End If
                    If IsNull(lineCatResults2("WebTotal")) = False Then
                        webTot2 = CDec(lineCatResults2("WebTotal"))
                        WebCostTotal(catCount) = WebCostTotal(catCount) + webTot2
                        If webTot2 > FirstPlaceAmount(catCount) Then
                            FirstPlaceLine(catCount) = CStr(ln)
                            FirstPlaceAmount(catCount) = CDec(webTot2)
                        Else
                            If webTot2 > SecondPlaceAmount(catCount) Then
                                SecondPlaceLine(catCount) = CStr(ln)
                                SecondPlaceAmount(catCount) = CDec(webTot2)
                            Else
                                If webTot2 > ThirdPlaceAmount(catCount) Then
                                    ThirdPlaceLine(catCount) = CStr(ln)
                                    ThirdPlaceAmount(catCount) = CDec(webTot2)
                                End If
                            End If
                        End If
                    End If
                Next ln
            End If
            If Me.chkEnablePOSP = vbChecked Then
                Dim pospQuery As String
                pospQuery = query
                pospQuery = Replace(pospQuery, "{TABLENAME}", POSPTable)
                pospQuery = Replace(pospQuery, "{STARTDATE}", POStartDate)
                pospQuery = Replace(pospQuery, "{ENDDATE}", POEndDate)
                
               
                Dim count4 As Integer
                For Each ln In tLines
                'count4 = count4 + 1
                lblStatus.Caption = prevStatus & "(POSP#" & CStr(count) & "/" & CStr(UBound(tLines)) & ")"
                lblStatus.Refresh
                Dim stdTot3 As Currency
                Dim webTot3 As Currency
                    Dim lineCatResults3 As ADODB.Recordset
                    Dim qStr3 As String
                    qStr3 = Replace(Replace(pospQuery, "{LINECODE}", CStr(ln)), "{S1ID}", CStr(cat))
                    Set lineCatResults3 = DB.retrieve(qStr3)
                    If IsNull(lineCatResults3("StdTotal")) = False Then
                        stdTot3 = CDec(lineCatResults3("StdTotal"))
                        StdCostTotal(catCount) = StdCostTotal(catCount) + stdTot3
                        
                    End If
                    If IsNull(lineCatResults3("WebTotal")) = False Then
                        webTot3 = CDec(lineCatResults3("WebTotal"))
                        WebCostTotal(catCount) = WebCostTotal(catCount) + webTot3
                        If webTot3 > FirstPlaceAmount(catCount) Then
                            FirstPlaceLine(catCount) = CStr(ln)
                            FirstPlaceAmount(catCount) = CDec(webTot3)
                        Else
                            If webTot3 > SecondPlaceAmount(catCount) Then
                                SecondPlaceLine(catCount) = CStr(ln)
                                SecondPlaceAmount(catCount) = CDec(webTot3)
                            Else
                                If webTot3 > ThirdPlaceAmount(catCount) Then
                                    ThirdPlaceLine(catCount) = CStr(ln)
                                    ThirdPlaceAmount(catCount) = CDec(webTot3)
                                End If
                            End If
                        End If
                    End If
                Next ln
            End If
            catCount = catCount + 1
        Next cat
        lblStatus.Caption = "Status: Writing Data... Please Wait."
        lblStatus.Refresh
        Dim testCount As Integer
        testCount = 0
        Results.AddItem "Category" & vbTab & "1st Place" & vbTab & "1st Amount" & vbTab & "2nd Place" & vbTab & "2nd Amount" & vbTab & "3rd Place" & vbTab & "3rd Amount" & vbTab & "Total Std" & vbTab & "Total Web"
        Do
            Dim catName As ADODB.Recordset
            Set catName = DB.retrieve("SELECT Name FROM SphereOneCategories WHERE ID=" & CStr(CInt(testCount + 1)))
            Dim cName As String
            cName = catName("Name")
            Dim line As String
            line = cName & vbTab & CStr(FirstPlaceLine(testCount)) & vbTab & "$" & CStr(FirstPlaceAmount(testCount)) & vbTab & CStr(SecondPlaceLine(testCount)) & vbTab & "$" & CStr(SecondPlaceAmount(testCount)) & vbTab & CStr(ThirdPlaceLine(testCount)) & vbTab & CStr(ThirdPlaceAmount(testCount)) & vbTab & "$" & CStr(StdCostTotal(testCount)) & vbTab & "$" & CStr(WebCostTotal(testCount))
            Results.AddItem line
            testCount = testCount + 1
        Loop Until testCount = UBound(FirstPlaceLine)
        Me.btnOpenInExcel.Enabled = True
        btnOpenInExcel_Click
End Sub



Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnMacroNEWW08_Click()
    selectList Array("G-I", "G-M")
    Me.chkEnableAR = vbChecked
    Me.ARFromDate.value = "5/2/2008"
    Me.ARToDate.value = "5/6/2008"
    Me.chkEnablePOSP = vbChecked
    Me.POSPFromDate.value = "5/2/2008"
    Me.POSPToDate.value = "5/6/2008"
    Me.chkEnableSO = vbChecked
    Me.SOFromDate.value = "5/2/2008"
    Me.SOToDate.value = "5/6/2008"
End Sub

Private Sub btnMacroToolarama07_Click()
    selectList Array("AST", "BER", "CAH", "COI", "PBL", "DON", "MDM", "DEL", "DEA", "DEP", "D-W", _
                     "D-A", "D-P", "ELM", "EXC", "FEI", "FRE", "G-I", "G-M", "HTC", "HTP", "IRW", _
                     "VIS", "HAN", "QUI", "S-L", "STR", "PSN", "JET", "JEA", "KES", "LAC", "LEN", _
                     "MAK", "MAA", "MAP", "MAS", "MET", "MLW", "MLA", "MLP", "NEW", "OCC", "PAC", _
                     "P-C", "P-A", "P-P", "POW", "POA", "QAC", "ROB", "ROO", "SAI", "SEN", "STB", _
                     "STI", "TPC", "COP", "ROL")
    Me.chkEnableAR = vbChecked
    Me.ARFromDate.value = "11/8/2007"
    Me.ARToDate.value = "11/12/2007"
    Me.chkEnablePOSP = vbChecked
    Me.POSPFromDate.value = "11/8/2007"
    Me.POSPToDate.value = "11/12/2007"
    Me.chkEnableSO = vbChecked
    Me.SOFromDate.value = "11/8/2007"
    Me.SOToDate.value = "11/12/2007"
End Sub

Private Sub btnMacroToolarama08_Click()
    selectList Array("BER", "CAH", "COI", "PBL", "DON", "MDM", "DEL", "DEA", "DEP", "D-W", "D-A", _
                     "EKL", "ELM", "EXC", "FRE", "G-I", "G-M", "HTC", "HTP", "IRW", "VIS", "HAN", _
                     "QUI", "STR", "PSN", "LEN", "JET", "JEA", "LAC", "MAK", "MAA", "MAS", "MET", _
                     "MLW", "MLA", "NEW", "OCC", "P-C", "P-A", "P-P", "POW", "POA", "QAC", "SEN", _
                     "SHR", "STB", "STI", "TPC", "SAI", "ROL", "COP")
    Me.chkEnableAR = vbChecked
    Me.ARFromDate.value = "11/12/2008"
    Me.ARToDate.value = "11/19/2008"
    Me.chkEnablePOSP = vbChecked
    Me.POSPFromDate.value = "11/12/2008"
    Me.POSPToDate.value = "11/19/2008"
    Me.chkEnableSO = vbChecked
    Me.SOFromDate.value = "11/12/2008"
    Me.SOToDate.value = "11/19/2008"
End Sub

Private Sub btnMacroToolarama09_Click()
    selectList Array("CAH", "COI", "DAR", "D-C", "PBL", "DON", "MDM", "DEL", "DEA", "DEP", "D-W", _
                     "D-A", "EKL", "ELM", "EXC", "FOO", "FRE", "G-I", "IRW", "VIS", "HAN", "STR", _
                     "LEN", "JET", "JEA", "LAC", "NEW", "P-C", "P-A", "P-P", "POW", "POA", "SHR", _
                     "STB", "TPC", "SAI", "WOR", "ROL", "COP", "MLW", "MLA", "HTC", "HTP", "MAK", _
                     "MAA", "STI", "SEN")
    Me.chkEnableAR = vbChecked
    Me.ARFromDate.value = "11/5/2009"
    Me.ARToDate.value = "11/9/2009"
    Me.chkEnablePOSP = vbChecked
    Me.POSPFromDate.value = "11/5/2009"
    Me.POSPToDate.value = "11/9/2009"
    Me.chkEnableSO = vbChecked
    Me.SOFromDate.value = "11/5/2009"
    Me.SOToDate.value = "11/9/2009"
End Sub

Private Sub btnMacroToolarama10_Click()
    selectList Array("FRE", "Q-D", "TPC", "WOR", "BEN", "PBL", "CEP", "CLC", "E-C", "DEL", "DEA", _
                     "DEP", "D-W", "D-A", "D-P", "DAR", "DRI", "EKL", "ELM", "FAS", "FOO", "G-I", _
                     "G-M", "GEN", "HAN", "HIT", "IRW", "VIS", "TAR", "TRP", "LAC", "L-C", "LEN", _
                     "MEG", "MLW", "MLA", "MLP", "NEW", "P-C", "P-A", "P-P", "QUI", "SAI", "SND", _
                     "SHR", "STA", "STI", "STR", "SWA", "WOO")
    Me.chkEnableAR = vbChecked
    Me.ARFromDate.value = "11/12/2010"
    Me.ARToDate.value = "11/14/2010"
    Me.chkEnablePOSP = vbChecked
    Me.POSPFromDate.value = "11/12/2010"
    Me.POSPToDate.value = "11/14/2010"
    Me.chkEnableSO = vbChecked
    Me.SOFromDate.value = "11/12/2010"
    Me.SOToDate.value = "11/14/2010"
End Sub

Private Sub btnMacroToolarama11_Click()
    selectList Array("AST", "AZP", "BCH", "BIG", "BOA", "BOH", "BOR", "BOS", "CEP", "CHA", "CLC", _
                     "COI", "CRE", "CUL", "D-A", "DAR", "D-C", "DEA", "DEL", "D-H", "DRI", "D-W", _
                     "E-C", "EKL", "ELM", "ERB", "EST", "FAL", "FAS", "FOO", "FRE", "GEN", "G-I", _
                     "GNR", "HAN", "HIT", "HTC", "I-A", "IMP", "IRW", "JEA", "JET", "KLE", "LAC", _
                     "L-C", "LEN", "MAA", "MAK", "MEG", "MLA", "MLW", "NEW", "NWB", "OSC", "P-A", _
                     "P-C", "POA", "POW", "ROC", "ROL", "SAI", "SEN", "SHR", "SKL", "SND", "STA", _
                     "STB", "STI", "STR", "SWA", "TAJ", "VIR", "WOO", "WOR", "WRX")
    Me.chkEnableAR = vbChecked
    Me.ARFromDate.value = "11/2/2011"
    Me.ARToDate.value = "11/9/2011"
    Me.chkEnablePOSP = vbChecked
    Me.POSPFromDate.value = "11/2/2011"
    Me.POSPToDate.value = "11/9/2011"
    Me.chkEnableSO = vbChecked
    Me.SOFromDate.value = "11/2/2011"
    Me.SOToDate.value = "11/9/2011"
End Sub

Private Sub btnMacroToolarama12_Click()
    selectList Array("JEA", "JET", "PER", "POA", "POW", "WIL", "BOA", "BOH", "BOP", "BOS", "CLC", _
                     "D-A", "D-C", "D-H", "D-P", "D-W", "DEA", "DEL", "DEP", "EKL", "FRE", "GNR", _
                     "HTC", "HTP", "LAC", "MLA", "MLW", "MLP", "P-A", "P-C", "P-P", "ROL", "SAI", _
                     "SEN", "STA", "STI", "BCA", "BCH", "BCP", "D-G", "DRI", "FAL", "MAA", "MAK", _
                     "MAP", "Q-D", "STB")
    Me.chkEnableAR = vbChecked
    Me.ARFromDate.value = "11/7/2012"
    Me.ARToDate.value = "11/14/2012"
    Me.chkEnablePOSP = vbChecked
    Me.POSPFromDate.value = "11/7/2012"
    Me.POSPToDate.value = "11/14/2012"
    Me.chkEnableSO = vbChecked
    Me.SOFromDate.value = "11/7/2012"
    Me.SOToDate.value = "11/14/2012"
End Sub

Private Sub btnMacroToolarama13_Click()
    selectList Array("JEA", "JET", "POA", "POW", "WIL", "BOA", "BOH", "BOS", "CLC", "D-A", "D-C", _
                     "D-H", "D-W", "DRI", "EDG", "ELM", "EXC", "FRE", "HAN", "HTC", "HTP", "IRW", _
                     "LAC", "LEN", "MET", "MLA", "MLW", "NEW", "P-A", "P-C", "Q-D", "QUI", "ROL", _
                     "SAI", "STA", "STB", "STI", "STR", "VIS", "FAL")
    Me.chkEnableAR = vbChecked
    Me.ARFromDate.value = "10/30/2013"
    Me.ARToDate.value = "11/6/2013"
    Me.chkEnablePOSP = vbChecked
    Me.POSPFromDate.value = "10/30/2013"
    Me.POSPToDate.value = "11/6/2013"
    Me.chkEnableSO = vbChecked
    Me.SOFromDate.value = "10/30/2013"
    Me.SOToDate.value = "11/6/2013"
End Sub

Private Sub btnMacroToolarama14_Click()
    selectList Array("GNR", "GNA", "CLC", "BCH", "BCA", "D-W", "D-A", "SKL", "AST", "B-A", "B-D", "BOA", "BOS", "D-C", "D-H", "DRE", "HAN", "HTC", "HTP", "IMP", "IRW", "JEA", "JET", "LAC", "LEN", "MET", "MLA", "MLW", "NOV", "P-A", "P-C", "POA", "POW", "Q-D", "QUI", "ROL", "ROT", "SAI", "SEN", "SHR", "STA", "STB", "STI", "STR", "VIS", "WIL")
    Me.chkEnableAR = vbChecked
    Me.ARFromDate.value = "10/22/2014"
    Me.ARToDate.value = "10/28/2014"
    Me.chkEnablePOSP = vbChecked
    Me.POSPFromDate.value = "10/22/2014"
    Me.POSPToDate.value = "10/28/2014"
    Me.chkEnableSO = vbChecked
    Me.SOFromDate.value = "10/22/2014"
    Me.SOToDate.value = "10/28/2014"
End Sub

Private Sub btnOpenInExcel_Click()
    Mouse.Hourglass True

    Dim line As Long, found As Boolean
    line = 0
    found = False
    While Not found And line < Me.Results.ListCount - 1
        If Left(Me.Results.list(line), 5) = "Total" Then
            found = True
        Else
            line = line + 1
        End If
    Wend
    If Not found Then
        line = line + 1
    End If
    
    Dim excelapp As Excel.Application
    Set excelapp = New Excel.Application
    excelapp.Workbooks.Add
    Dim row As Long, col As Long
    row = 1
    
    Dim i As Long, j As Long
    For i = 0 To Me.Results.ListCount
        col = 1
        If Me.Results.list(i) = "" Then
            Exit For
        End If
        Dim cells As Variant
        cells = Split(Me.Results.list(i), vbTab)
        If i = 0 Then
            If Me.chkIncludeVendorInSpreadsheet Then
                excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(col) & row).value = "Vendor"
                col = col + 1
            End If
            For j = 0 To UBound(cells)
                excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(col) & row).value = cells(j)
                col = col + 1
            Next j
            If Me.chkShowTotals Then
                excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(col) & row).value = "Total"
            End If
        Else
            If Me.chkIncludeVendorInSpreadsheet Then
                excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(col) & row).value = lookupLineCodePrimaryVendor(CStr(cells(0)))
                col = col + 1
            End If
            For j = 0 To UBound(cells)
                excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(col) & row).value = cells(j)
                col = col + 1
            Next j
            If Me.chkShowTotals Then
                excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(col) & row).value = "=SUM(" & IIf(Me.chkIncludeVendorInSpreadsheet, "C", "B") & row & ":" & ColumnIndexToLetter(col - 1) & row & ")"
            End If
        End If
        row = row + 1
    Next i
    If Me.chkShowTotals Then
        excelapp.ActiveWorkbook.ActiveSheet.Range("A" & row).value = "Totals:"
        col = IIf(Me.chkIncludeVendorInSpreadsheet, 3, 2)
        Dim minCol As String, maxCol As String
        minCol = ColumnIndexToLetter(col)
        While excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(col) & "1").value <> "Total"
            excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(col) & row).value = "=SUM(" & ColumnIndexToLetter(col) & "2:" & ColumnIndexToLetter(col) & (row - 1) & ")"
            maxCol = ColumnIndexToLetter(col)
            col = col + 1
        Wend
        row = row + 1
        excelapp.ActiveWorkbook.ActiveSheet.Range("A" & row).value = "Overall Total:"
        excelapp.ActiveWorkbook.ActiveSheet.Range("B" & row).value = "=SUM(" & minCol & (row - 1) & ":" & maxCol & (row - 1) & ")"
    End If
    
    excelapp.Visible = True
    
    Mouse.Hourglass False
End Sub

Private Sub chkEnableAR_Click()
    If Me.chkEnableAR Then
        'Enable Me.btnARFromDate
        'Enable Me.btnARToDate
        'Enable Me.ARFromDate
        'Enable Me.ARToDate
        Me.frameAR.Enabled = True
    Else
        'Disable Me.btnARFromDate
        'Disable Me.btnARToDate
        'Disable Me.ARFromDate
        'Disable Me.ARToDate
        Me.frameAR.Enabled = False
    End If
End Sub

Private Sub chkEnablePOSP_Click()
    If Me.chkEnablePOSP Then
        'Enable Me.btnPOSPFromDate
        'Enable Me.btnPOSPToDate
        'Enable Me.POSPFromDate
        'Enable Me.POSPToDate
        Me.framePOSP.Enabled = True
    Else
        'Disable Me.btnPOSPFromDate
        'Disable Me.btnPOSPToDate
        'Disable Me.POSPFromDate
        'Disable Me.POSPToDate
        Me.framePOSP.Enabled = False
    End If
End Sub

Private Sub chkEnableSO_Click()
    If Me.chkEnableSO Then
        'Enable Me.btnSOFromDate
        'Enable Me.btnSOToDate
        'Enable Me.SOFromDate
        'Enable Me.SOToDate
        Me.frameSO.Enabled = True
    Else
        'Disable Me.btnSOFromDate
        'Disable Me.btnSOToDate
        'Disable Me.SOFromDate
        'Disable Me.SOToDate
        Me.frameSO.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
Dim rst As ADODB.Recordset
Set rst = MASDB.retrieve("SELECT TOP 10 * FROM {oj SO_SalesOrderHistoryHeader LEFT OUTER JOIN SO_SalesOrderHistoryDetail ON SO_SalesOrderHistoryHeader.SalesOrderNo=SO_SalesOrderHistoryDetail.SalesOrderNo}")
MsgBox rst.RecordCount
End Sub

Private Sub Form_Load()
    Me.POSPFromDate.value = Date
    Me.POSPToDate.value = Date
    Me.ARFromDate.value = Date
    Me.ARToDate.value = Date
    Me.SOFromDate.value = Date
    Me.SOToDate.value = Date
    requeryLines
End Sub

Private Sub requeryLines()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ProductLine FROM ProductLine WHERE Hide=0 ORDER BY ProductLine")
    Me.Lines.Clear
    While Not rst.EOF
        Me.Lines.AddItem rst("ProductLine")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Function buildSQLString(source As String, sellines() As String) As String
    Dim sql As String
    Dim PL As String
    PL = IIf(Me.chkIncludeSpecial, "RealProductLine", "ProductLine")
    Dim dateclause As String
    Select Case source
        Case Is = "POSP"
            dateclause = "TransactionDate BETWEEN '" & Me.POSPFromDate.value & "' AND '" & Me.POSPToDate.value & "'"
        Case Is = "AR"
            dateclause = "TransactionDate BETWEEN '" & Me.ARFromDate.value & "' AND '" & Me.ARToDate.value & "'"
        Case Is = "SO"
            dateclause = "TransactionDate BETWEEN '" & Me.SOFromDate.value & "' AND '" & Me.SOToDate.value & "'"
    End Select
    Select Case True
        Case Is = Me.opMetric(METRIC_SO)
            sql = "SELECT OrderNumber, ItemNumber, Quantity, Extension "
            sql = sql & "FROM vShowSaleCalc" & source & " "
            sql = sql & "WHERE " & dateclause & " "
            sql = sql & "AND " & PL & " IN ( " & ListAsCommaSep(sellines, "'", True) & " ) "
            sql = sql & "ORDER BY OrderNumber, ItemNumber"
        Case Else
            sql = "INSERT INTO ShowSaleCalculations ( Source, PLorItem, TotalAmount ) "
            sql = sql & "SELECT '" & source & "', "
            Select Case True
                Case Is = Me.opMetric(METRIC_SALES)
                    sql = sql & PL & ", SUM(Extension) "
                Case Is = Me.opMetric(METRIC_QTYS)
                    sql = sql & "ItemNumber, SUM(Quantity) "
            End Select
            sql = sql & "FROM vShowSaleCalc" & source & " "
            sql = sql & "WHERE " & dateclause & " "
            sql = sql & "AND " & PL & " IN ( " & ListAsCommaSep(sellines, "'", True) & " ) "
            Select Case True
                Case Is = Me.opMetric(METRIC_SALES)
                    sql = sql & "GROUP BY " & PL
                Case Is = Me.opMetric(METRIC_QTYS)
                    sql = sql & "GROUP BY ItemNumber"
            End Select
    End Select
    buildSQLString = sql
End Function

Private Function formatForDisplay(data As String) As String
    Select Case True
        Case Is = Me.opMetric(METRIC_SALES)
            formatForDisplay = Format(data, "Currency")
        Case Is = Me.opMetric(METRIC_QTYS)
            formatForDisplay = data
    End Select
End Function

Private Function getHeaderName() As String
    Select Case True
        Case Is = Me.opMetric(METRIC_SALES)
            getHeaderName = "LineCode"
        Case Is = Me.opMetric(METRIC_QTYS)
            getHeaderName = "Item"
    End Select
End Function

Private Sub setTabs()
    Dim tabs(2) As Long
    Select Case True
        Case Is = Me.opMetric(METRIC_SO)
            tabs(0) = 66
            tabs(1) = tabs(0) + 48
            tabs(2) = tabs(1) + 48
            SetListTabStops Me.Results.hWnd, tabs
        Case Else
            tabs(0) = 66
            tabs(1) = tabs(0) + 48
            tabs(2) = tabs(1) + 48
            SetListTabStops Me.Results.hWnd, tabs
    End Select
End Sub

'Private Sub addVendorToResults()
'    Dim i As Long, currLine As String, vendor As String
'    Me.Results.list(0) = Replace(Me.Results.list(0), vbTab, vbTab & "Vendor" & vbTab, 1, 1)
'    For i = 1 To Me.Results.ListCount - 1
'        currLine = Left(Me.Results.list(i), 3)
'        vendor = DLookup("PrimaryVendorNumber", "ProductLine", "ProductLine='" & currLine & "'")
'        Me.Results.list(i) = Replace(Me.Results.list(i), vbTab, vbTab & vendor & vbTab, 1, 1)
'    Next i
'End Sub
