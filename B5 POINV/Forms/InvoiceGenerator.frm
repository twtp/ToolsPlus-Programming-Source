VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#3.5#0"; "TPControls.ocx"
Begin VB.Form InvoiceGenerator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Generator"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame generalFrame 
      Caption         =   "Shipping"
      Height          =   1215
      Index           =   4
      Left            =   2910
      TabIndex        =   29
      Top             =   3360
      Width           =   2025
      Begin VB.OptionButton ShippingOption 
         Caption         =   "Don't Show"
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   32
         Top             =   810
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton ShippingOption 
         Caption         =   "Recalculate"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   31
         Top             =   540
         Width           =   1155
      End
      Begin VB.OptionButton ShippingOption 
         Caption         =   "From Order"
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   30
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Discount Application"
      Height          =   975
      Index           =   3
      Left            =   120
      TabIndex        =   26
      Top             =   4890
      Width           =   2715
      Begin VB.OptionButton DiscountApplicationOption 
         Caption         =   "Applies Per-Item"
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   28
         Top             =   540
         Width           =   1665
      End
      Begin VB.OptionButton DiscountApplicationOption 
         Caption         =   "Applies Only Once"
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   27
         Top             =   270
         Value           =   -1  'True
         Width           =   1665
      End
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Header"
      Height          =   765
      Index           =   2
      Left            =   6000
      TabIndex        =   18
      Top             =   660
      Width           =   2055
      Begin TPControls.NumericUpDown HeaderRow 
         Height          =   315
         Left            =   840
         TabIndex        =   20
         Top             =   330
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Value           =   "1"
         Minimum         =   1
         Maximum         =   100
         Increment       =   1
         DisplayFormat   =   ""
         AllowDecimals   =   0   'False
         TimeBeforeHoldDown=   500
         TimeBetweenRepeats=   50
         Enabled         =   -1  'True
         BackColor       =   -2147483643
         DefaultValue    =   ""
      End
      Begin VB.Label generalLabel 
         Caption         =   "On Row:"
         Height          =   255
         Index           =   9
         Left            =   150
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Discount"
      Height          =   1485
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   2715
      Begin VB.OptionButton DiscountOption 
         Caption         =   "Override to:"
         Height          =   285
         Index           =   3
         Left            =   90
         TabIndex        =   24
         Top             =   1080
         Width           =   1395
      End
      Begin VB.OptionButton DiscountOption 
         Caption         =   "File, Default to:"
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   23
         Top             =   810
         Width           =   1395
      End
      Begin VB.OptionButton DiscountOption 
         Caption         =   "From File"
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   22
         Top             =   540
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.OptionButton DiscountOption 
         Caption         =   "No Discount"
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   21
         Top             =   270
         Width           =   1395
      End
      Begin TPControls.NumericUpDown DiscountAmount 
         Height          =   315
         Left            =   1530
         TabIndex        =   25
         Top             =   900
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Value           =   "50"
         Minimum         =   1
         Maximum         =   1000
         Increment       =   1
         DisplayFormat   =   ""
         AllowDecimals   =   -1  'True
         TimeBeforeHoldDown=   500
         TimeBetweenRepeats=   50
         Enabled         =   -1  'True
         BackColor       =   -2147483643
         DefaultValue    =   ""
      End
   End
   Begin VB.CommandButton GenerateButton 
      Caption         =   "Generate Invoice HTML"
      Height          =   465
      Left            =   4260
      TabIndex        =   16
      Top             =   6000
      Width           =   2115
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Type"
      Height          =   1215
      Index           =   0
      Left            =   6000
      TabIndex        =   12
      Top             =   1590
      Width           =   2055
      Begin VB.OptionButton TypeOption 
         Caption         =   "All Point of Sale"
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   15
         Top             =   810
         Width           =   1785
      End
      Begin VB.OptionButton TypeOption 
         Caption         =   "All Sales Orders"
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   14
         Top             =   540
         Width           =   1785
      End
      Begin VB.OptionButton TypeOption 
         Caption         =   "Use ""Type"" Column"
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   13
         Top             =   270
         Value           =   -1  'True
         Width           =   1785
      End
   End
   Begin VB.CommandButton ValidateFileButton 
      Caption         =   "Validate File"
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   2850
      Width           =   1785
   End
   Begin VB.CommandButton BrowseButton 
      Caption         =   "Browse"
      Height          =   315
      Left            =   6990
      TabIndex        =   2
      Top             =   210
      Width           =   1065
   End
   Begin VB.TextBox ExcelFileName 
      Height          =   315
      Left            =   1410
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   210
      Width           =   5565
   End
   Begin VB.Label generalLabel 
      Caption         =   """Unit Price"": required"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   11
      Top             =   2550
      Width           =   5685
   End
   Begin VB.Label generalLabel 
      Caption         =   """Quantity"": required"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   5685
   End
   Begin VB.Label generalLabel 
      Caption         =   """Item Number"": required"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   9
      Top             =   2010
      Width           =   5685
   End
   Begin VB.Label generalLabel 
      Caption         =   """Customer Name"": optional, can be looked up by the order number "
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   8
      Top             =   1740
      Width           =   5685
   End
   Begin VB.Label generalLabel 
      Caption         =   """Order Date"": optional, can be looked up by the order number "
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   7
      Top             =   1470
      Width           =   5685
   End
   Begin VB.Label generalLabel 
      Caption         =   """Order Number"": required"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   5685
   End
   Begin VB.Label generalLabel 
      Caption         =   """Type"": optional, see right, either ""s/o"" or ""pos"""
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   930
      Width           =   5685
   End
   Begin VB.Label generalLabel 
      Caption         =   "NOTE: Header row (set on right) must have the following columns names:"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   660
      Width           =   5805
   End
   Begin VB.Label generalLabel 
      Caption         =   "Drag XLS Here:"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   240
      Width           =   1275
   End
End
Attribute VB_Name = "InvoiceGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private excelApp As Object
Private itemDescriptionCache As Scripting.Dictionary
Private DB As DBConn.DatabaseSingleton
Private columnMapping As Scripting.Dictionary
Private fso As Scripting.FileSystemObject
Private mouse As MouseHourglass



Private Sub Form_Load()
    Set excelApp = CreateObject("Excel.Application")
    excelApp.DisplayAlerts = False
    Set fso = New Scripting.FileSystemObject
    Set mouse = New MouseHourglass
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not excelApp Is Nothing Then
        excelApp.Workbooks.Close
        excelApp.Quit
        Set excelApp = Nothing
    End If
    Set itemDescriptionCache = Nothing
End Sub

Private Sub ExcelFileName_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetFormat(vbCFFiles) Then
        Dim fname As String
        fname = Data.Files(1)
        If Right(fname, 4) = ".xls" Or Right(fname, 4) = ".csv" Or Right(fname, 5) = ".xlsx" Then
            Me.ExcelFileName.Text = fname
            excelApp.Workbooks.Close
        Else
            MsgBox "This doesn't look like an Excel file!"
            Exit Sub
        End If
        
        On Error Resume Next
        excelApp.Workbooks.Open filename:=fname, ReadOnly:=True
        If Err.Number <> 0 Then
            MsgBox "Error opening spreadsheet: " & Err.Description
            Err.Clear
        Else
            excelApp.Workbooks.Close
        End If
    End If
End Sub

Private Sub BrowseButton_Click()
    Dim fp As FilePicker
    Set fp = New FilePicker
    fp.AddFilter "Excel Documents (*.xls)", "*.xls"
    fp.AddFilter "Excel Documents (*.xlsx)", "*.xlsx"
    fp.AddFilter "Excel Documents (*.csv)", "*.csv"
    Dim fname As String
    Me.ExcelFileName = fp.ShowDialogOpen
End Sub

Private Sub ValidateFileButton_Click()
    mouse.Hourglass True
    If Me.ExcelFileName = "" Then
        MsgBox "Select a file!"
        GoTo endsub
    End If
    If fso.FileExists(Me.ExcelFileName) = False Then
        MsgBox "File does not exist!"
        GoTo endsub
    End If
    
    excelApp.Workbooks.Open filename:=Me.ExcelFileName, ReadOnly:=True
    
    Dim sheet As Object
    If excelApp.ActiveWorkbook Is Nothing Then
        MsgBox "Load a spreadsheet!"
        GoTo endsub
    End If
    Set sheet = excelApp.ActiveWorkbook.ActiveSheet
    
    Set columnMapping = generateColumnMapping(sheet, CLng(Me.HeaderRow.Value))
    
    Dim errorMessage As String
    errorMessage = validateColumnMapping()
    If errorMessage <> "" Then
        MsgBox errorMessage
        GoTo endsub
    End If
    
    Dim consecutiveEmptyLines As Long, lastDataLine As Long
    consecutiveEmptyLines = 0
    lastDataLine = 0
    
    Dim row As Long
    row = CLng(Me.HeaderRow.Value) + 1
    While consecutiveEmptyLines < 10
        If sheet.Range(columnMapping.Item("Order Number") & row).Value = "" Then
            consecutiveEmptyLines = consecutiveEmptyLines + 1
        Else
            consecutiveEmptyLines = 0
            lastDataLine = row
        End If
        row = row + 1
    Wend
    
    If lastDataLine = 0 Then
        MsgBox "Header looks good, but I didn't find any data?"
        GoTo endsub
    End If
    
    MsgBox "Looks good! Found data rows up to row " & lastDataLine
    GoTo endsub
    
endsub:
    excelApp.Workbooks.Close
    mouse.Hourglass False
End Sub

Private Sub GenerateButton_Click()
    mouse.Hourglass True
    If Me.ExcelFileName = "" Then
        MsgBox "Select a file!"
        GoTo endsub
    End If
    If fso.FileExists(Me.ExcelFileName) = False Then
        MsgBox "File does not exist!"
        GoTo endsub
    End If
    
    excelApp.Workbooks.Open filename:=Me.ExcelFileName, ReadOnly:=True
    
    Dim sheet As Object
    If excelApp.ActiveWorkbook Is Nothing Then
        MsgBox "Load a spreadsheet!"
        GoTo endsub
    End If
    Set sheet = excelApp.ActiveWorkbook.ActiveSheet
    
    Set columnMapping = generateColumnMapping(sheet, CLng(Me.HeaderRow.Value))
    
    Dim errorMessage As String
    errorMessage = validateColumnMapping()
    If errorMessage <> "" Then
        MsgBox errorMessage
        GoTo endsub
    End If
    
    Dim cache As Scripting.Dictionary
    Set cache = buildOrderCache(sheet, CLng(Me.HeaderRow.Value) + 1)
    
    ReDim fakeInvoices(0 To cache.Count - 1) As String
    
    Dim iter As String, i As Long
    For i = 0 To cache.Count - 1
        iter = cache.Keys(i)
        Dim thisInvoice As String
        
        thisInvoice = generateHtmlHeaderArea(cache.Item(CStr(iter))) & _
                      generateHtmlAddressArea(cache.Item(CStr(iter))) & _
                      generateHtmlItemsArea(cache.Item(CStr(iter)))
        thisInvoice = thisInvoice & _
                      "<table width=""600"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf & _
                      " <tr>" & vbCrLf & _
                      "  <td>" & vbCrLf & _
                      "   <p>Tools Plus Order Processing Department<br />" & vbCrLf & _
                      "      <a href=""mailto:orderprocessing@tools-plus.com"">orderprocessing@tools-plus.com</a><br />" & vbCrLf & _
                      "      60 Scott Road, Prospect, CT 06712<br />" & vbCrLf & _
                      "      Telephone: (203) 573-0750 or 1-800-222-6133<br />" & vbCrLf & _
                      "      Visit our Retail Store:<br />" & vbCrLf & _
                      "      153 Meadow Street, Waterbury, CT 06702</p>" & vbCrLf & _
                      "  </td>" & vbCrLf & _
                      " </tr>" & vbCrLf & _
                      "</table>" & vbCrLf
        fakeInvoices(i) = "<div class=""order-summary"">" & vbCrLf & thisInvoice & "</div>" & vbCrLf
    Next i
    
    Dim html As String
    html = Join(fakeInvoices, "")
    
    Dim filename As String
retryfilename:
    'filename = excelApp.Application.GetSaveAsFilename(fileFilter:="HTML Documents (*.html), *.html")
    Dim fp As FilePicker
    Set fp = New FilePicker
    fp.SetParent Me.hWnd
    fp.AddFilter "HTML Documents", "*.html"
    filename = fp.ShowDialogSave
    If filename = "" Then
        MsgBox "I've already run through and generated everything, so I'm not letting you cancel this. Pick a damn file name."
        GoTo retryfilename
    End If
    
    Open filename For Output As #1
    Print #1, "<!DOCTYPE html>" & vbCrLf & _
              "<html>" & vbCrLf & _
              "<head>" & vbCrLf & _
              "<style type=""text/css"">" & vbCrLf & _
              "body { font-size: 12px; }" & vbCrLf & _
              "div.order-summary { page-break-after: always; }" & vbCrLf & _
              "</style>" & vbCrLf & _
              "</head>" & vbCrLf & _
              "<body>" & vbCrLf
    Print #1, html
    Print #1, "</body>" & vbCrLf & _
              "</html>" & vbCrLf
    Close #1
    
endsub:
    excelApp.Workbooks.Close
    mouse.Hourglass False
End Sub











'----------------------------------------
'      SPREADSHEET/DATABASE HEAVY LIFTING
'----------------------------------------
Private Function buildOrderCache(sheet As Object, startRow As Long)
    Dim cache As Scripting.Dictionary, temp As Scripting.Dictionary
    Set cache = New Scripting.Dictionary
    
    If DB Is Nothing Then
        Set DB = New DBConn.DatabaseSingleton
        DB.CurrentDB = DBENG_MAS90
    End If
    
    Dim row As Long
    row = startRow
    Dim consecutiveEmptyLines As Long
    consecutiveEmptyLines = 0
    While consecutiveEmptyLines < 10
        Dim orderType As String, orderNumber As String
        orderType = getOrderType(sheet, row)
        orderNumber = getValueFromSheet("Order Number", sheet, row)
        
        If orderNumber = "" Then
            consecutiveEmptyLines = consecutiveEmptyLines + 1
            GoTo endofprocessing
        End If
        
        consecutiveEmptyLines = 0
        
        Dim key As String
        key = orderType & orderNumber
        If Not cache.Exists(key) Then
            'order header required, db lookup, add
            Dim rst As ADODB.Recordset
            Select Case orderType
                Case Is = "s/o"
                    Set rst = DB.retrieve("SELECT OrderDate, BillToName, BillToCity, BillToState, BillToZipCode, BillToCountryCode, ShipToName, ShipToCity, ShipToState, ShipToZipCode, ShipToCountryCode, ShipVia, PaymentType FROM SO_SalesOrderHistoryHeader WHERE SalesOrderNo='" & orderNumber & "'")
                Case Is = "pos"
                    Set rst = DB.retrieve("SELECT PaymentType, Name FROM P2_TransactionHistoryPayments WHERE TransactionNo='" & orderNumber & "'")
                Case Else
                    MsgBox "Missing handler for type """ & orderType & """ on row " & row & ", skipping"
                    GoTo endofprocessing
            End Select
            
            Set temp = New Scripting.Dictionary
            temp.Add "Type", orderType
            temp.Add "Order Number", orderNumber
            temp.Add "Order Date", getValueFromSheet("Order Date", sheet, row) ' can be empty string, will fill from db
            temp.Add "Customer Name", Trim(getValueFromSheet("Customer Name", sheet, row)) ' can be empty
            temp.Add "Items", New Scripting.Dictionary
            temp.Add "Discounts", CDec(0)
            'temp.Add "Order Type", "Standard"
            Select Case orderType
                Case Is = "s/o"
                    temp.Add "Ship Via", CStr(nz(rst("ShipVia"), "UPS INET"))
                    If temp.Item("Ship Via") = "DROP SHIP" Then
                        temp.Item("Ship Via") = "UPS INET"
                    End If
                    temp.Add "Payment Type", CStr(rst("PaymentType"))
                    temp.Add "Bill To City", CStr(nz(rst("BillToCity")))
                    temp.Add "Bill To State", CStr(nz(rst("BillToState")))
                    temp.Add "Bill To Zip", CStr(nz(rst("BillToZipCode")))
                    temp.Add "Bill To Country", CStr(nz(rst("BillToCountryCode"), "US"))
                    temp.Add "Ship To Name", CStr(nz(rst("ShipToName")))
                    temp.Add "Ship To City", CStr(nz(rst("ShipToCity")))
                    temp.Add "Ship To State", CStr(nz(rst("ShipToState")))
                    temp.Add "Ship To Zip", CStr(nz(rst("ShipToZipCode")))
                    temp.Add "Ship To Country", CStr(nz(rst("ShipToCountryCode"), "US"))
                    If temp.Item("Order Date") = "" Then
                        temp.Item("Order Date") = CStr(Format(rst("OrderDate"), "m/d/yyyy"))
                    End If
                    If temp.Item("Customer Name") = "" Then
                        temp.Item("Customer Name") = Trim(CStr(rst("BillToName")))
                    End If
                Case Is = "pos"
                    temp.Add "Ship Via", CStr("Store Pickup")
                    temp.Add "Payment Type", CStr(rst("PaymentType"))
                    If temp.Item("Customer Name") = "" And rst("PaymentType") <> "CASH" Then
                        temp.Item("Customer Name") = Trim(CStr(nz(rst("Name"))))
                    End If
                    If temp.Item("Customer Name") = "" Then
                        temp.Item("Customer Name") = "<not recorded>"
                    End If
            End Select
            
            rst.Close
            Set rst = Nothing
            
            Select Case orderType
                Case Is = "s/o"
                    Set rst = DB.retrieve("SELECT LastExtensionAmt FROM SO_SalesOrderHistoryDetail WHERE SalesOrderNo='" & orderNumber & "' AND ItemCode='/SHIP'")
                    If rst.EOF Then
                        temp.Add "Shipping Cost", CDec(0)
                    Else
                        temp.Add "Shipping Cost", CDec(rst("LastExtensionAmt"))
                    End If
                    rst.Close
                    Set rst = Nothing
                Case Is = "pos"
                    temp.Add "Shipping Cost", Null
            End Select
            
            cache.Add key, temp
            Set temp = Nothing
        End If
        
        'now handle the item line
        Dim itemNumber As String, unitPrice As String
        itemNumber = getValueFromSheet("Item Number", sheet, row)
        itemNumber = Replace(itemNumber, " ", "")
        unitPrice = getValueFromSheet("Unit Price", sheet, row)
        
        Dim newOrdered As Long
        
        If Left(itemNumber, 1) <> "/" Then
            'actual item
            newOrdered = CLng(getValueFromSheet("Quantity", sheet, row))
            
            If cache.Item(key).Item("Items").Exists(itemNumber) Then
                cache.Item(key).Item("Items").Item(itemNumber).Item("Ordered") = cache.Item(key).Item("Items").Item(itemNumber).Item("Ordered") + newOrdered
                cache.Item(key).Item("Items").Item(itemNumber).Item("Subtotal") = CDec(cache.Item(key).Item("Items").Item(itemNumber).Item("Subtotal") + (newOrdered * CDec(unitPrice)))
            Else
                Set temp = New Scripting.Dictionary
                temp.Add "Ordered", newOrdered
                temp.Add "Subtotal", CDec(newOrdered * CDec(unitPrice))
                Set cache.Item(key).Item("Items").Item(itemNumber) = temp
                Set temp = Nothing
            End If
        ElseIf itemNumber = "/GFT" Or itemNumber = "/GC" Then
            'gift certificates don't count
        Else
            'misc line
            cache.Item(key).Item("Discounts") = CDec(cache.Item(key).Item("Discounts") + Abs(CDec(unitPrice)))
        End If
        
endofprocessing:
        row = row + 1
    Wend
    
    Dim orderIter As Variant
    Select Case True
        Case Is = Me.ShippingOption(0) ' from order
            ' no changes needed
        Case Is = Me.ShippingOption(1) ' recalculate
            ' TODO, option is currently disabled in gui
        Case Is = Me.ShippingOption(2) ' do not show
            For Each orderIter In cache.Keys
                cache.Item(CStr(orderIter)).Remove "Shipping Cost"
            Next orderIter
    End Select
    
    Set buildOrderCache = cache
End Function




'----------------------------------------
'      HTML GENERATION
'----------------------------------------
Private Function generateHtmlHeaderArea(orderCache As Scripting.Dictionary) As String
    Dim thisInvoice As String
    'HEADER BLOCK
    thisInvoice = "<table width=""600"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf & _
                  " <tr>" & vbCrLf & _
                  "  <td width=""150""><img src=""http://lib.store.yahoo.net/lib/toolsplus/invoice-logo.gif"" alt=""Tools Plus Logo"" /></td>" & vbCrLf & _
                  "  <td width=""450"">" & vbCrLf & _
                  "   <table width=""450"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf & _
                  "    <tr>" & vbCrLf & _
                  "     <h2>" & getSummaryTypeFromType(orderCache.Item("Type")) & "</h2>" & vbCrLf
    'BASIC ORDER INFO
    thisInvoice = thisInvoice & _
                  "     <td valign=""top"" width=""225"">" & vbCrLf & _
                  "      <b>Order Number:</b> " & orderCache.Item("Order Number") & "<br />" & vbCrLf & _
                  "      <b>Order Date:</b> " & orderCache.Item("Order Date") & "<br />" & vbCrLf & _
                  "      <b>Order Type:</b> " & orderCache.Item("Order Type") & "<br />" & vbCrLf & _
                  "      <b>Delivery Method:</b> " & orderCache.Item("Ship Via") & "<br />" & vbCrLf & _
                  "      <b>Payment Method:</b> " & orderCache.Item("Payment Type") & vbCrLf & _
                  "     </td>" & vbCrLf
    'TP ADDRESS
    thisInvoice = thisInvoice & _
                  "     <td valign=""top"" width=""225"" style=""text-align:right;"">" & vbCrLf & _
                  "      <b>Tools-Plus.com</b><br />" & vbCrLf & _
                  "      60 Scott Road<br />" & vbCrLf & _
                  "      Prospect, CT 06712<br />" & vbCrLf & _
                  "      Toll-Free: (800) 222-6133<br />" & vbCrLf & _
                  "      Fax: (203) 753-9042<br />" & vbCrLf & _
                  "      E-Mail: <a href=""mailto:orderprocessing@tools-plus.com"">orderprocessing@tools-plus.com</a><br />" & vbCrLf & _
                  "      Visit our Retail Store:<br />" & vbCrLf & _
                  "      153 Meadow Street<br />" & vbCrLf & _
                  "      Waterbury, CT 06702" & vbCrLf & _
                  "     </td>" & vbCrLf
    'END OF HEADER
    thisInvoice = thisInvoice & _
                  "    </tr>" & vbCrLf & _
                  "   </table>" & vbCrLf & _
                  "  </td>" & vbCrLf & _
                  " </tr>" & vbCrLf & _
                  "</table>" & vbCrLf & _
                  "<hr width=""600"" align=""left"" />" & vbCrLf & _
                  "<br />" & vbCrLf
    
    generateHtmlHeaderArea = thisInvoice
End Function

Private Function generateHtmlAddressArea(orderCache As Scripting.Dictionary) As String
    Dim thisInvoice As String
    'CUSTOMER ADDRESS
    thisInvoice = thisInvoice & _
                  "<table width=""600"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf & _
                  " <tr>" & vbCrLf
    Select Case orderCache.Item("Type")
        Case Is = "s/o"
            'SALES ORDER BILLING
            thisInvoice = thisInvoice & _
                          "  <td width=""295"" valign=""top"">" & vbCrLf & _
                          "   <b>Billing Address</b>" & vbCrLf & _
                          "   <hr />" & vbCrLf & _
                          "   <table width=""295"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf & _
                          "    <tr><td><b>Name:</b></td><td>" & encodeEntities(orderCache.Item("Customer Name")) & "</td></tr>" & vbCrLf & _
                          "    <tr><td><b>Address 1:</b></td><td>" & "xxxxx" & "</td></tr>" & vbCrLf & _
                          "    <tr><td><b>Address 2:</b></td><td>" & "xxxxx" & "</td></tr>" & vbCrLf & _
                          "    <tr><td><b>City:</b></td><td>" & encodeEntities(orderCache.Item("Bill To City")) & "</td></tr>" & vbCrLf & _
                          "    <tr><td><b>State:</b></td><td>" & encodeEntities(orderCache.Item("Bill To State")) & "</td></tr>" & vbCrLf & _
                          "    <tr><td><b>Zip/Postal Code:</b></td><td>" & encodeEntities(orderCache.Item("Bill To Zip")) & "</td></tr>" & vbCrLf & _
                          "    <tr><td><b>Country:</b></td><td>" & encodeEntities(orderCache.Item("Bill To Country")) & "</td></tr>" & vbCrLf & _
                          "    <tr><td><b>Phone:</b></td><td>" & "xxxxx" & "</td></tr>" & vbCrLf & _
                          "   </table>" & vbCrLf & _
                          "  </td>" & vbCrLf
            'SALES ORDER SHIPPING
            thisInvoice = thisInvoice & _
                          "  <td width=""295"" valign=""top"">" & vbCrLf & _
                          "   <b>Shipping Address</b>" & vbCrLf & _
                          "   <hr />" & vbCrLf & _
                          "   <table width=""295"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf & _
                          "    <tr><td><b>Name:</b></td><td>" & encodeEntities(orderCache.Item("Customer Name")) & "</td></tr>" & vbCrLf & _
                          "    <tr><td><b>Address 1:</b></td><td>" & "xxxxx" & "</td></tr>" & vbCrLf & _
                          "    <tr><td><b>Address 2:</b></td><td>" & "xxxxx" & "</td></tr>" & vbCrLf & _
                          "    <tr><td><b>City:</b></td><td>" & encodeEntities(orderCache.Item("Ship To City")) & "</td></tr>" & vbCrLf & _
                          "    <tr><td><b>State:</b></td><td>" & encodeEntities(orderCache.Item("Ship To State")) & "</td></tr>" & vbCrLf & _
                          "    <tr><td><b>Zip/Postal Code:</b></td><td>" & encodeEntities(orderCache.Item("Ship To Zip")) & "</td></tr>" & vbCrLf & _
                          "    <tr><td><b>Country:</b></td><td>" & encodeEntities(orderCache.Item("Ship To Country")) & "</td></tr>" & vbCrLf & _
                          "    <tr><td><b>Phone:</b></td><td>" & "xxxxx" & "</td></tr>" & vbCrLf & _
                          "   </table>" & vbCrLf & _
                          "  </td>" & vbCrLf
            thisInvoice = thisInvoice & _
                          " </tr>" & vbCrLf & _
                          "</table>" & vbCrLf & _
                          "<br />" & vbCrLf
        Case Is = "pos"
            'POINT OF SALE, CUSTOMER NAME ONLY
            thisInvoice = thisInvoice & _
                          "  <td width=""295"" valign=""top""><b>Customer Name:</b> " & encodeEntities(orderCache.Item("Customer Name")) & "</td>" & vbCrLf & _
                          " </tr>" & vbCrLf & _
                          "</table>" & vbCrLf & _
                          "<br />" & vbCrLf
    End Select
    
    generateHtmlAddressArea = thisInvoice
End Function

Private Function generateHtmlItemsArea(orderCache As Scripting.Dictionary) As String
    Dim orderType As String, itemsCache As Scripting.Dictionary
    orderType = orderCache.Item("Type")
    Set itemsCache = orderCache.Item("Items")
    
    Dim thisInvoice As String
    'ITEMS BLOCK
    Select Case orderType
        Case Is = "s/o"
            'SALES ORDERS - ORDERED/SHIPPED/BACKORDERED
            thisInvoice = thisInvoice & _
                          "<table width=""600"" border=""1"" cellspacing=""0"" cellpadding=""1"">" & vbCrLf & _
                          " <tr>" & vbCrLf & _
                          "  <th>Item Number</th>" & vbCrLf & _
                          "  <th>Description</th>" & vbCrLf & _
                          "  <th>Ordered</th>" & vbCrLf & _
                          "  <th>Shipped</th>" & vbCrLf & _
                          "  <th>Backordered</th>" & vbCrLf & _
                          "  <th>Subtotal</th>" & vbCrLf & _
                          " </tr>" & vbCrLf
        Case Is = "pos"
            'POINT OF SALE - SOLD ONLY
            thisInvoice = thisInvoice & _
                          "<table width=""600"" border=""1"" cellspacing=""0"" cellpadding=""1"">" & vbCrLf & _
                          " <tr>" & vbCrLf & _
                          "  <th>Item Number</th>" & vbCrLf & _
                          "  <th>Description</th>" & vbCrLf & _
                          "  <th>Sold</th>" & vbCrLf & _
                          "  <th>Subtotal</th>" & vbCrLf & _
                          " </tr>" & vbCrLf
    End Select
    Dim iter As Variant
    Dim itemsTotal As Variant, itemsQty As Long
    itemsTotal = CDec(0)
    itemsQty = 0
    'ITERATE ITEMS
    For Each iter In itemsCache.Keys
        Select Case orderType
            Case Is = "s/o"
                thisInvoice = thisInvoice & _
                              " <tr>" & vbCrLf & _
                              "  <td>" & iter & "</td>" & vbCrLf & _
                              "  <td>" & encodeEntities(getItemDescription(CStr(iter))) & "</td>" & vbCrLf & _
                              "  <td align=""center"">" & itemsCache.Item(CStr(iter)).Item("Ordered") & "</td>" & vbCrLf & _
                              "  <td align=""center"">" & itemsCache.Item(CStr(iter)).Item("Ordered") & "</td>" & vbCrLf & _
                              "  <td align=""center"">0</td>" & vbCrLf & _
                              "  <td align=""right"">" & Format(itemsCache.Item(CStr(iter)).Item("Subtotal"), "Currency") & "</td>" & vbCrLf & _
                              " </tr>" & vbCrLf
            Case Is = "pos"
                thisInvoice = thisInvoice & _
                              " <tr>" & vbCrLf & _
                              "  <td>" & iter & "</td>" & vbCrLf & _
                              "  <td>" & encodeEntities(getItemDescription(CStr(iter))) & "</td>" & vbCrLf & _
                              "  <td align=""center"">" & itemsCache.Item(CStr(iter)).Item("Ordered") & "</td>" & vbCrLf & _
                              "  <td align=""right"">" & Format(itemsCache.Item(CStr(iter)).Item("Subtotal"), "Currency") & "</td>" & vbCrLf & _
                              " </tr>" & vbCrLf
        End Select
        itemsTotal = CDec(itemsTotal + itemsCache.Item(CStr(iter)).Item("Subtotal"))
        itemsQty = itemsQty + CLng(itemsCache.Item(CStr(iter)).Item("Ordered"))
    Next iter
    
    'SUBTOTAL AMOUNT
    Dim colspan As Long
    Select Case orderType
        Case Is = "s/o"
            colspan = 5
        Case Is = "pos"
            colspan = 3
    End Select
    thisInvoice = thisInvoice & _
                  " <tr>" & vbCrLf & _
                  "  <td colspan=""" & colspan & """ align=""right"">Items Subtotal:</td>" & vbCrLf & _
                  "  <td align=""right"">" & Format(itemsTotal, "Currency") & "</td>" & vbCrLf & _
                  " </tr>" & vbCrLf
    
    'DISCOUNT AMOUNT
    Dim discountTotal As Variant
    Select Case True
        Case Is = Me.DiscountOption(0) 'no discount
            discountTotal = 0
        Case Is = Me.DiscountOption(1) 'file only, no default
            discountTotal = CDec(orderCache.Item("Discounts"))
        Case Is = Me.DiscountOption(2) 'file, fallback
            discountTotal = CDec(orderCache.Item("Discounts"))
            If discountTotal = 0 Then
                discountTotal = CDec(Me.DiscountAmount.Value)
                Select Case True
                    Case Is = Me.DiscountApplicationOption(0) 'applied singly
                        'no nothing
                    Case Is = Me.DiscountApplicationOption(1) 'applied per-item
                        discountTotal = discountTotal * itemsQty
                End Select
            End If
        Case Is = Me.DiscountOption(3) 'always override
            discountTotal = CDec(Me.DiscountAmount.Value)
            Select Case True
                Case Is = Me.DiscountApplicationOption(0) 'applied singly
                    'no nothing
                Case Is = Me.DiscountApplicationOption(1) 'applied per-item
                    discountTotal = discountTotal * itemsQty
            End Select
    End Select
    If discountTotal <> 0 Then
        thisInvoice = thisInvoice & _
                      " <tr>" & vbCrLf & _
                      "  <td colspan=""" & colspan & """ align=""right"">Discount:</td>" & vbCrLf & _
                      "  <td align=""right"">- " & Format(discountTotal, "Currency") & "</td>" & vbCrLf & _
                      " </tr>" & vbCrLf
    End If

    'SHIPPING AMOUNT
    Dim shippingTotal As Variant
    If orderCache.Exists("Shipping Cost") = False Then
        'nonexistent means do not display
        shippingTotal = 0
    ElseIf IsNull(orderCache.Item("Shipping Cost")) Then
        'null means p2, do not display
        shippingTotal = 0
    Else
        'standard handling
        shippingTotal = CDec(orderCache.Item("Shipping Cost"))
        Dim shippingTotalPrintable As String
        If shippingTotal = 0 Then
            shippingTotalPrintable = "FREE"
        Else
            shippingTotalPrintable = Format(shippingTotal, "Currency")
        End If
        thisInvoice = thisInvoice & _
                      " <tr>" & vbCrLf & _
                      "  <td colspan=""" & colspan & """ align=""right"">Shipping:</td>" & vbCrLf & _
                      "  <td align=""right"">" & shippingTotalPrintable & "</td>" & vbCrLf & _
                      " </tr>" & vbCrLf
    End If
    
    'TAX AMOUNT
    Dim taxTotal As Variant
    taxTotal = getTaxAmount(orderCache, itemsTotal, CDec(shippingTotal))
    If taxTotal <> 0 Then
        thisInvoice = thisInvoice & _
                      " <tr>" & vbCrLf & _
                      "  <td colspan=""" & colspan & """ align=""right"">Tax:</td>" & vbCrLf & _
                      "  <td align=""right"">" & Format(taxTotal, "Currency") & "</td>" & vbCrLf & _
                      " </tr>" & vbCrLf
    End If
    
    'TOTAL AMOUNT
    Dim totalTotal As Variant
    totalTotal = itemsTotal - discountTotal + shippingTotal + taxTotal
    thisInvoice = thisInvoice & _
                  " <tr>" & vbCrLf & _
                  "  <td colspan=""" & colspan & """ align=""right"">Total:</td>" & vbCrLf & _
                  "  <td align=""right"">" & Format(totalTotal, "Currency") & "</td>" & vbCrLf & _
                  " </tr>" & vbCrLf
    
    thisInvoice = thisInvoice & _
                  "</table>" & vbCrLf & _
                  "<br />" & vbCrLf
    
    generateHtmlItemsArea = thisInvoice
End Function



'----------------------------------------
'      SPREADSHEET GETTERS
'----------------------------------------
Private Function getOrderType(sheet As Object, row As Long) As String
    Select Case True
        Case Is = Me.TypeOption(0) ' column only
            getOrderType = getValueFromSheet("Type", sheet, row)
        Case Is = Me.TypeOption(1), Me.TypeOption(2) ' force s/o or pos
            getOrderType = Replace(CStr(columnMapping.Item("Type")), "*", "")
    End Select
End Function

Private Function getValueFromSheet(valueName As String, sheet As Object, row As Long) As String
    If Left(columnMapping.Item(valueName), 1) = "*" Then
        getValueFromSheet = ""
    Else
        getValueFromSheet = sheet.Range(CStr(columnMapping.Item(valueName)) & row).Value
    End If
End Function



'----------------------------------------
'      DATABASE ACCESS
'----------------------------------------
Private Function getItemDescription(itemNumber As String) As String
    If itemDescriptionCache Is Nothing Then
        Set itemDescriptionCache = New Scripting.Dictionary
    End If
    If DB Is Nothing Then
        Set DB = New DatabaseSingleton
        DB.CurrentDB = DBENG_MAS90
    End If
    If Not itemDescriptionCache.Exists(itemNumber) Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT ItemCodeDesc FROM CI_Item WHERE ItemCode='" & itemNumber & "'")
        If rst.EOF Then
            itemDescriptionCache.Add itemNumber, "ERROR"
        Else
            itemDescriptionCache.Add itemNumber, Trim(CStr(rst("ItemCodeDesc")))
        End If
        rst.Close
        Set rst = Nothing
    End If
    
    getItemDescription = itemDescriptionCache.Item(itemNumber)
End Function




'----------------------------------------
'      DISPLAY AND CALCULATIONS
'----------------------------------------
Private Function getSummaryTypeFromType(invoiceType As String) As String
    Select Case invoiceType
        Case Is = "s/o"
            getSummaryTypeFromType = "ORDER SUMMARY"
        Case Is = "pos"
            getSummaryTypeFromType = "TRANSACTION SUMMARY"
    End Select
End Function

Private Function getTaxAmount(orderDef As Scripting.Dictionary, itemsTotal As Variant, shippingAmount) As Variant
    itemsTotal = CDec(itemsTotal)
    Select Case orderDef.Item("Type")
        Case Is = "s/o"
            If Left(orderDef.Item("Ship To Country"), 2) <> "US" Then
                getTaxAmount = 0
            ElseIf Left(orderDef.Item("Ship To State"), 2) <> "CT" Then
                getTaxAmount = 0
            Else
                getTaxAmount = Round((itemsTotal + shippingAmount) * 0.0635, 2)
            End If
        Case Is = "pos"
            getTaxAmount = Round((itemsTotal + shippingAmount) * 0.0635, 2)
    End Select
End Function


'----------------------------------------
'      COLUMN HEADER MAPPING
'----------------------------------------
Private Function generateColumnMapping(sheet As Object, headerRowIndex As Long) As Scripting.Dictionary
    Dim col As Long, row As Long
    col = 1
    
    Dim retval As Scripting.Dictionary
    Set retval = New Dictionary
    
    While sheet.Range(columnIndexToLetter(col) & headerRowIndex).Value <> ""
        retval.Add sheet.Range(columnIndexToLetter(col) & headerRowIndex).Value, columnIndexToLetter(col)
        col = col + 1
    Wend
    
    Select Case True
        Case Is = Me.TypeOption(0) ' no override
            'do nothing, this will error eventually
        Case Is = Me.TypeOption(1) ' all s/o
            If retval.Exists("Type") = False Then
                retval.Add "Type", "*s/o"
            Else
                retval.Item("Type") = "*s/o"
            End If
        Case Is = Me.TypeOption(2) ' all pos
            If retval.Exists("Type") = False Then
                retval.Add "Type", "*pos"
            Else
                retval.Item("Type") = "*pos"
            End If
    End Select
    
    If retval.Exists("Order Date") = False Then
        retval.Add "Order Date", "*db"
    End If
    
    If retval.Exists("Customer Name") = False Then
        retval.Add "Customer Name", "*db"
    End If
    
    Set generateColumnMapping = retval
End Function

Private Function validateColumnMapping() As String
    Dim requiredColumns As Variant
    requiredColumns = Array("Type", "Order Number", "Order Date", "Customer Name", "Item Number", "Quantity", "Unit Price")
    
    Dim retval As String
    retval = ""
    
    Dim iter As Variant
    For Each iter In requiredColumns
        If Not columnMapping.Exists(CStr(iter)) Then
            retval = IIf(retval = "", "", retval & vbCrLf) & "Couldn't find column """ & iter & """"
        End If
    Next iter
    
    validateColumnMapping = retval
End Function





'----------------------------------------
'      GENERAL UTILITY FUNCTIONS
'----------------------------------------
Private Function columnIndexToLetter(dx As Long) As String
    Dim first As Long
    first = dx \ 26
    Dim second As Long
    second = dx Mod 26
    If second = 0 Then
        second = 26
        first = first - 1
    End If
    columnIndexToLetter = IIf(first = 0, "", Chr(64 + first)) & IIf(second = 0, "Z", Chr(64 + second))
End Function

Private Function nz(txt As Variant, Optional default As String = "") As String
    nz = IIf(IsNull(txt), default, txt)
End Function

Private Function encodeEntities(str As String) As String
    Dim retval As String
    retval = ""
    
    Dim i As Long
    For i = 1 To Len(str)
        Dim c As String
        c = Mid(str, i, 1)
        If c = "<" Then
            retval = retval & "&lt;"
        ElseIf c = ">" Then
            retval = retval & "&gt;"
        ElseIf c = "&" Then
            retval = retval & "&amp;"
        ElseIf c = """" Then
            retval = retval & "&quot;"
        ElseIf c = "'" Then
            retval = retval & "&apos;"
        Else
            retval = retval & c
        End If
    Next i
    
    encodeEntities = retval
End Function
