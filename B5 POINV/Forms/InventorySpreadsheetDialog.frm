VERSION 5.00
Begin VB.Form InventorySpreadsheetDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory - Generate Spreadsheet"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   6600
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton VendorPreset 
      Caption         =   "Joe Brakhage"
      Height          =   315
      Index           =   5
      Left            =   5250
      TabIndex        =   46
      Tag             =   "297"
      Top             =   600
      Width           =   1305
   End
   Begin VB.CommandButton VendorPreset 
      Caption         =   "Rob Desimone"
      Height          =   315
      Index           =   4
      Left            =   5250
      TabIndex        =   45
      Tag             =   "172"
      Top             =   1920
      Width           =   1305
   End
   Begin VB.CommandButton VendorPreset 
      Caption         =   "Frank Bednarz"
      Height          =   315
      Index           =   3
      Left            =   5250
      TabIndex        =   44
      Tag             =   "95"
      Top             =   270
      Width           =   1305
   End
   Begin VB.CommandButton VendorPreset 
      Caption         =   "Karl Bertelmann"
      Height          =   315
      Index           =   2
      Left            =   5250
      TabIndex        =   43
      Tag             =   "272"
      Top             =   930
      Width           =   1305
   End
   Begin VB.CommandButton VendorPreset 
      Caption         =   "Paul Reynolds"
      Height          =   315
      Index           =   1
      Left            =   5250
      TabIndex        =   42
      Tag             =   "289"
      Top             =   1590
      Width           =   1305
   End
   Begin VB.CommandButton VendorPreset 
      Caption         =   "Mike Engel"
      Height          =   315
      Index           =   0
      Left            =   5250
      TabIndex        =   41
      Tag             =   "308"
      Top             =   1260
      Width           =   1305
   End
   Begin VB.ListBox LineCodeList 
      Height          =   5130
      Left            =   5220
      MultiSelect     =   2  'Extended
      TabIndex        =   40
      Top             =   2640
      Width           =   1305
   End
   Begin VB.Frame Frame3 
      Caption         =   "Select Additional Options:"
      Height          =   615
      Left            =   180
      TabIndex        =   18
      Top             =   7590
      Width           =   4515
      Begin VB.CheckBox chkSkipDiscontinued 
         Caption         =   "Skip Discontinued Items"
         Height          =   285
         Left            =   150
         TabIndex        =   19
         Top             =   240
         Width           =   2805
      End
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   525
      Left            =   2640
      TabIndex        =   17
      Top             =   8310
      Width           =   1665
   End
   Begin VB.CommandButton btnGenerate 
      Caption         =   "Generate"
      Height          =   525
      Left            =   810
      TabIndex        =   16
      Top             =   8310
      Width           =   1665
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Columns to Fill:"
      Height          =   6045
      Left            =   210
      TabIndex        =   3
      Top             =   1530
      Width           =   4695
      Begin VB.CheckBox chkColumn 
         Caption         =   "Sphere One Category"
         Height          =   285
         Index           =   37
         Left            =   150
         TabIndex        =   57
         Top             =   5640
         Width           =   2055
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Availability Limit (Displays)"
         Height          =   255
         Index           =   22
         Left            =   2430
         TabIndex        =   56
         Top             =   1440
         Width           =   2205
      End
      Begin VB.TextBox creditPercBox 
         Height          =   285
         Left            =   4110
         TabIndex        =   55
         Text            =   "0"
         ToolTipText     =   "Credit percent"
         Top             =   5610
         Width           =   405
      End
      Begin VB.TextBox salePercBox 
         Height          =   285
         Left            =   3690
         TabIndex        =   54
         Text            =   "0"
         ToolTipText     =   "Sale percent"
         Top             =   5610
         Width           =   405
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Sale Percent"
         Height          =   285
         Index           =   36
         Left            =   2430
         TabIndex        =   53
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "D/S Charges (Comm.)"
         Height          =   255
         Index           =   10
         Left            =   150
         TabIndex        =   51
         Top             =   3240
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Replaced By/For"
         Height          =   255
         Index           =   35
         Left            =   2430
         TabIndex        =   50
         Top             =   5340
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "EBay Commission (Web)"
         Height          =   255
         Index           =   17
         Left            =   150
         TabIndex        =   49
         Top             =   5340
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "EBay Price"
         Height          =   255
         Index           =   15
         Left            =   150
         TabIndex        =   48
         Top             =   4740
         Width           =   2115
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "EBay Commission (EBay)"
         Height          =   255
         Index           =   16
         Left            =   150
         TabIndex        =   47
         Top             =   5040
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Sales Figures (365 only)"
         Height          =   255
         Index           =   24
         Left            =   2430
         TabIndex        =   38
         Top             =   2040
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Quantity On SO"
         Height          =   255
         Index           =   21
         Left            =   2430
         TabIndex        =   37
         Top             =   1140
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Follies Status"
         Height          =   255
         Index           =   34
         Left            =   2430
         TabIndex        =   36
         Top             =   5040
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Shipping Type"
         Height          =   255
         Index           =   33
         Left            =   2430
         TabIndex        =   35
         Top             =   4740
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Average Freight"
         Height          =   255
         Index           =   32
         Left            =   2430
         TabIndex        =   34
         Top             =   4440
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Additional Info (EPN)"
         Height          =   255
         Index           =   31
         Left            =   2430
         TabIndex        =   33
         Top             =   4140
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "D/S Status"
         Height          =   255
         Index           =   30
         Left            =   2430
         TabIndex        =   32
         Top             =   3840
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "D/C Status"
         Height          =   255
         Index           =   29
         Left            =   2430
         TabIndex        =   31
         Top             =   3540
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Components"
         Height          =   255
         Index           =   28
         Left            =   2430
         TabIndex        =   30
         Top             =   3240
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Entered/Published Dates"
         Height          =   255
         Index           =   27
         Left            =   2430
         TabIndex        =   29
         Top             =   2940
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Triad Code"
         Height          =   255
         Index           =   26
         Left            =   2430
         TabIndex        =   28
         Top             =   2640
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "PO Comment"
         Height          =   255
         Index           =   25
         Left            =   2430
         TabIndex        =   27
         Top             =   2340
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Sales Figures"
         Height          =   255
         Index           =   23
         Left            =   2430
         TabIndex        =   26
         Top             =   1740
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Quantity On Order"
         Height          =   255
         Index           =   20
         Left            =   2430
         TabIndex        =   25
         Top             =   840
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Quantity On PO"
         Height          =   255
         Index           =   19
         Left            =   2430
         TabIndex        =   24
         Top             =   540
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Current QOH"
         Height          =   255
         Index           =   18
         Left            =   2430
         TabIndex        =   23
         Top             =   240
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Average Cost"
         Height          =   255
         Index           =   9
         Left            =   150
         TabIndex        =   22
         Top             =   2940
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "MAPP"
         Height          =   255
         Index           =   14
         Left            =   150
         TabIndex        =   21
         Top             =   4440
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Web Price"
         Height          =   255
         Index           =   13
         Left            =   150
         TabIndex        =   15
         Top             =   4140
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Store Price"
         Height          =   255
         Index           =   12
         Left            =   150
         TabIndex        =   14
         Top             =   3840
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "List Price"
         Height          =   255
         Index           =   11
         Left            =   150
         TabIndex        =   13
         Top             =   3540
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "New Cost"
         Height          =   255
         Index           =   8
         Left            =   150
         TabIndex        =   12
         Top             =   2640
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Standard Cost"
         Height          =   255
         Index           =   7
         Left            =   150
         TabIndex        =   11
         Top             =   2340
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Weight / Dimensions"
         Height          =   255
         Index           =   6
         Left            =   150
         TabIndex        =   10
         Top             =   2040
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Extended Specs"
         Height          =   255
         Index           =   5
         Left            =   150
         TabIndex        =   9
         Top             =   1740
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Specs"
         Height          =   255
         Index           =   4
         Left            =   150
         TabIndex        =   8
         Top             =   1440
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Barcode"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   7
         Top             =   1140
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Description"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Top             =   840
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Part Number"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   5
         Top             =   540
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.CheckBox chkColumn 
         Caption         =   "Line Code"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   240
         Value           =   1  'Checked
         Width           =   2085
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Spreadsheet Type:"
      Height          =   1425
      Left            =   600
      TabIndex        =   0
      Top             =   60
      Width           =   3585
      Begin VB.OptionButton opXSheetType 
         Caption         =   "Create Pre-Sale Analysis Spreadsheet"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   52
         Top             =   1140
         Width           =   3345
      End
      Begin VB.OptionButton opXSheetType 
         Caption         =   "Create Vendor Spreadsheet"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   39
         Top             =   840
         Width           =   2835
      End
      Begin VB.OptionButton opXSheetType 
         Caption         =   "Create Regular Spreadsheet"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   540
         Value           =   -1  'True
         Width           =   2835
      End
      Begin VB.OptionButton opXSheetType 
         Caption         =   "Create New Products Spreadsheet"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   2835
      End
   End
   Begin VB.Label lblStatusBar 
      Height          =   225
      Left            =   0
      TabIndex        =   20
      Top             =   8880
      Width           =   3615
   End
End
Attribute VB_Name = "InventorySpreadsheetDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NEW_PROD_XSHEET_LOCATION As String = "h:\Common Documents\New Products Spreadsheet (Empty).xls"

Private Const DX_LINE_CODE As Long = 0
Private Const DX_PART_NUMBER As Long = 1

Private stopNow As Boolean
'Private AllItems As PerlArray

Private bigAssLookup As Dictionary

Private Function lookupFill() As Boolean
    Set bigAssLookup = New Dictionary
    
    'arguments are:
    '  chkColumn index for control
    '  whether this control is "special" (disabled and always-on)
    '  variant/array of column headings. when spaces are removed, these match the field name in the view.
    '  column number to use, -999 means the next column (and also switches enabledness for new prod spreadsheet)
    '  whether this control should be auto-checked when switching to the vendor spreadsheet
    '  whether this control should be auto-checked when switching to the analysis spreadsheet
    
    'line code
    lookupAdd 0, True, Array("Line Code"), 1, True, True
    'part number
    lookupAdd 1, True, Array("Part Number"), 2, True, True
    'description
    lookupAdd 2, False, Array("Description"), 4, True, True
    'barcode
    lookupAdd 3, False, Array("Barcode"), 3, True, False
    'specs
    lookupAdd 4, False, Array("Feature 1", "Feature 2", "Feature 3", "Feature 4", "Feature 5", "Feature 6", "Feature 7", "Feature 8"), 6, False, False
    'ext specs
    lookupAdd 5, False, Array("Tech Specifications"), 14, False, False
    'wt/dims
    lookupAdd 6, False, Array("Weight", "Dimensions"), 16, True, False
    'std cost
    lookupAdd 7, False, Array("Standard Cost"), 22, True, True
    'new cost
    lookupAdd 8, False, Array("New Cost"), 21, True, True
    'avg cost
    lookupAdd 9, False, Array("Average Cost"), -999, False, True
    'cost + d/s adders
    lookupAdd 10, False, Array("D/S Charges (Commercial)"), -999, False, True
    'list price
    lookupAdd 11, False, Array("List Price"), 22, True, False
    'store price
    lookupAdd 12, False, Array("Store Price"), 24, True, True
    'web price
    lookupAdd 13, False, Array("Web Price"), 25, True, True
    'mapp
    lookupAdd 14, False, Array("MAPP"), 28, True, True
    'ebay price
    lookupAdd 15, False, Array("EBay Price", "EBay Best Offer", "EBay Best Offer Auto Accept"), -999, False, False
    'ebay commission (ebay)
    lookupAdd 16, False, Array("EBay Commission EBay"), -999, False, False
    'ebay commission (web)
    lookupAdd 17, False, Array("EBay Commission Web"), -999, False, False
    'current qoh
    lookupAdd 18, False, Array("Quantity On Hand"), -999, True, False
    'qty on po
    lookupAdd 19, False, Array("Quantity On PO"), -999, True, False
    'qty on order
    lookupAdd 20, False, Array("QuantityOnOrder"), 27, True, False
    'qty on so
    lookupAdd 21, False, Array("Quantity On SO"), -999, True, False
    'availability limit
    lookupAdd 22, False, Array("Availability Limit Store", "Availability Limit EBay"), -999, False, True
    'sales
    lookupAdd 23, False, Array("SO Sales Last 30 Days", "SO Sales Last 90 Days", "SO Sales Last 365 Days", "P2 Sales Last 30 Days", "P2 Sales Last 90 Days", "P2 Sales Last 365 Days"), -999, True, False
    'sales 365
    lookupAdd 24, False, Array("SO Sales Last 365 Days", "P2 Sales Last 365 Days", "Total Sales Last 365 Days"), -999, False, False
    'po comment
    lookupAdd 25, False, Array("PO Comment"), -999, True, False
    'triad code
    lookupAdd 26, False, Array("Triad Code"), -999, True, False
    'entered/pub date
    lookupAdd 27, False, Array("Entered Date", "Published Date"), -999, True, False
    'components
    lookupAdd 28, False, Array("Components"), -999, False, False
    'd/c
    lookupAdd 29, False, Array("Is Being DC", "Is DC On Web", "Is DC"), -999, True, False
    'd/s
    lookupAdd 30, False, Array("Is Dropshippable", "Is Dropship Only"), -999, False, True
    'epn
    lookupAdd 31, False, Array("Additional Info"), -999, False, False
    'avg freight
    lookupAdd 32, False, Array("Average Freight"), -999, False, True
    'shipping type
    lookupAdd 33, False, Array("Shipping Type"), -999, False, True
    'follies
    lookupAdd 34, False, Array("Follies Status"), -999, True, False
    'replaced by/for
    lookupAdd 35, False, Array("Replaced By", "Replacement For"), -999, False, False
    'sale percent discounts
    lookupAdd 36, False, Array("Sale Percent"), -999, False, True
    'sphere one category
    lookupAdd 37, False, Array("Sphere One Category"), -999, False, False
    
    lookupFill = True
End Function

Private Function lookupAdd(dx As Long, isSpecial As Boolean, ColumnList As Variant, columnNo As Long, vendorSheetChecked As Boolean, preSaleSheetChecked As Boolean) As Boolean
    bigAssLookup.Add CStr(dx), New Dictionary
    bigAssLookup.item(CStr(dx)).Add "Special", isSpecial
    bigAssLookup.item(CStr(dx)).Add "Columns", ColumnList
    bigAssLookup.item(CStr(dx)).Add "ColumnNo", columnNo
    bigAssLookup.item(CStr(dx)).Add "VendorCheck", vendorSheetChecked
    bigAssLookup.item(CStr(dx)).Add "PreSaleCheck", preSaleSheetChecked
    
    lookupAdd = True
End Function

Private Function lookupSpecialness(dx As Long) As Variant
    lookupSpecialness = bigAssLookup.item(CStr(dx)).item("Special")
End Function

Private Function lookupColumnList(dx As Long) As Variant
    lookupColumnList = bigAssLookup.item(CStr(dx)).item("Columns")
End Function

Private Function lookupColumnIndex(dx As Long) As Long
    lookupColumnIndex = bigAssLookup.item(CStr(dx)).item("ColumnNo")
End Function

Private Function lookupNewProdEnabled(dx As Long) As Boolean
    lookupNewProdEnabled = CBool(bigAssLookup.item(CStr(dx)).item("ColumnNo") <> -999)
End Function

Private Function lookupVendorAutoCheck(dx As Long) As Boolean
    lookupVendorAutoCheck = bigAssLookup.item(CStr(dx)).item("VendorCheck")
End Function

Private Function lookupPreSaleAutoCheck(dx As Long) As Boolean
    lookupPreSaleAutoCheck = bigAssLookup.item(CStr(dx)).item("PreSaleCheck")
End Function

Private Sub btnCancel_Click()
    If Me.btnCancel.Caption = "Cancel" Then
        Unload Me
    Else
        stopNow = True
    End If
End Sub

Private Sub btnGenerate_Click()
On Error GoTo errh
    Mouse.Hourglass True
    
    stopNow = False
    lockForm
    
    Me.lblStatusBar.Caption = "Preparing..."
    DoEvents
    
    Dim sourceItems As PerlArray
    Set sourceItems = loadSourceItems()
    
    Dim excelapp As Excel.Application
    Set excelapp = New Excel.Application
    
    Dim currRow As Long, currCol As Long
    Dim i As Long, j As Long
    Dim colNames As Variant, thisColName As Variant
    
    Select Case True
        Case Is = Me.opXSheetType(0) 'new products spreadsheet
            excelapp.Workbooks.Open NEW_PROD_XSHEET_LOCATION
            currRow = 18
        Case Is = Me.opXSheetType(1), Me.opXSheetType(2), Me.opXSheetType(3) 'regular spreadsheet, vendor spreadsheet, sale spreadsheet
            excelapp.Workbooks.Add
            currCol = 1
            For i = 0 To Me.chkColumn.UBound
                If Me.chkColumn(i) Then
                    colNames = resolveControlIndex(i, 0)
                    For Each thisColName In colNames
                        excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(currCol) & "1").value = thisColName
                        currCol = currCol + 1
                    Next thisColName
                End If
            Next i
            excelapp.ActiveWorkbook.ActiveSheet.Columns("A:B").NumberFormat = "@" 'lc and pn are text-only
            currRow = 2
        Case Else
            MsgBox "missing spreadsheet type in select in generation header!"
    End Select
    
    'sql & sqlcol to xsheetcol mapping
    Dim sqlCols As PerlArray
    Set sqlCols = New PerlArray
    sqlCols.Push "DISCONTINUED"
    Dim xsheetMap As Dictionary
    Set xsheetMap = New Dictionary
    Dim startingCol As Long
    startingCol = 1
    For i = 0 To Me.chkColumn.UBound
        colNames = resolveControlIndex(i, 1)
        If Me.chkColumn(i) Then
            If colNames(0) = "Barcode" Then
                'barcode is handled differently
                xsheetMap.Add CStr(colNames(0)), startingCol
                startingCol = startingCol + 1
            ElseIf colNames(0) = "D/SCharges(Commercial)" Then
                'd/s commercial
                xsheetMap.Add CStr(colNames(0)), startingCol
                startingCol = startingCol + 1
            ElseIf colNames(0) = "SalePercent" Then
                'xsheetMap.Add CStr(colNames(0)), startingCol
                'startingCol = startingCol + 1
            ElseIf colNames(0) = "SphereOneCategory" Then
            sqlCols.Push "SphereOneCategories.Name AS 'SphereOneCategory'"
            xsheetMap.Add CStr(colNames(0)), startingCol
            startingCol = startingCol + 1
            '    colNames(0) = "SphereOneCategories.Name"
            
            Else
                Select Case True
                    Case Is = Me.opXSheetType(0) 'new products spreadsheet
                        startingCol = lookupColumnIndex(i)
                    Case Is = Me.opXSheetType(1), Me.opXSheetType(2), Me.opXSheetType(3) 'regular spreadsheet, vendor spreadsheet, sale spreadsheet
                        'startingCol is ok
                    Case Else
                        MsgBox "missing spreadsheet type in select in generation mapping!"
                End Select
                For Each thisColName In colNames
                    sqlCols.Push CStr(thisColName)
                    xsheetMap.Add CStr(thisColName), startingCol
                    startingCol = startingCol + 1
                Next thisColName
            End If
        End If
    Next i
    Dim sql As String
    sql = "SELECT " & sqlCols.Join(", ") & " FROM vInventorySpreadsheet WHERE vInventorySpreadsheet.ITEMNUMBER="
    If InStr(sql, "SphereOneCategories") > 0 Then
        sql = Split(sql, "WHERE")(0) & "LEFT OUTER JOIN PartNumberWebPaths ON PartNumberWebPaths.ItemNumber=vInventorySpreadsheet.ItemNumber LEFT OUTER JOIN MasterCategoryConnectors ON MasterCategoryConnectors.MasterID=PartNumberWebPaths.WebPathID LEFT OUTER JOIN SphereOneCategories ON SphereOneCategories.ID=MasterCategoryConnectors.ConnectorID WHERE " & Split(sql, "WHERE")(1)
    End If
    
    
    Me.lblStatusBar.Caption = ""
    DoEvents
    
    If stopNow Then
        excelapp.ActiveWorkbook.Close False
        excelapp.Quit
        GoTo theend
    End If
    
    For i = 0 To sourceItems.Scalar - 1
        If stopNow Then
            Exit For
        End If
        Me.lblStatusBar.Caption = i * 100 \ sourceItems.Scalar & "% complete"
        DoEvents
        Dim rst As ADODB.Recordset
        Dim query As String
        query = sql & "'" & sourceItems.Elem(i) & "'"
        'MsgBox query
        Set rst = DB.retrieve(query)
        If rst.EOF Then
            'what?
        Else
            If rst("DISCONTINUED") And Me.chkSkipDiscontinued Then
                'skip
            Else
                For j = 0 To rst.fields.count - 1
                    If rst(j).name = UCase(rst(j).name) And rst(j).name <> "MAPP" Then
                        'skip col, all uppercase is a control column
                    Else
                        currCol = xsheetMap.item(CStr(rst(j).name))
                        If currCol = -999 Then
                            'skip col, not on xsheet
                        Else
                            excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(currCol) & currRow).value = Nz(rst(j))
                        End If
                    End If
                Next j
                If xsheetMap.exists("Barcode") Then
                    'barcodes are handled differently
                    currCol = xsheetMap.item("Barcode")
                    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(currCol) & currRow).value = BarcodeFunctions.FindBestBarcodeFor(sourceItems.Elem(i))
                End If
                If xsheetMap.exists("D/SCharges(Commercial)") Then
                    'dropship costs are handled differently
                    currCol = xsheetMap.item("D/SCharges(Commercial)")
                    excelapp.ActiveWorkbook.ActiveSheet.Range(ColumnIndexToLetter(currCol) & currRow).value = CalculateDropshipChargesFor(sourceItems.Elem(i), 0)
                End If
                currRow = currRow + 1
            End If
        End If
        rst.Close
        Set rst = Nothing
    Next i
    
    If Me.opXSheetType(3) Then 'pre-sales analysis
        Dim webSalePriceCol As String
        webSalePriceCol = ColumnIndexToLetter(startingCol)
        Dim webDistCreditCol As String
        webDistCreditCol = ColumnIndexToLetter(startingCol + 1)
        Dim webShippingChargeCol As String
        webShippingChargeCol = ColumnIndexToLetter(startingCol + 2)
        Dim webPaymentProcessorPercentCol As String
        webPaymentProcessorPercentCol = ColumnIndexToLetter(startingCol + 3)
        Dim webHostingPercentCol As String
        webHostingPercentCol = ColumnIndexToLetter(startingCol + 4)
        Dim webShippingCostCol As String
        webShippingCostCol = ColumnIndexToLetter(startingCol + 5)
        Dim webTotalCostCol As String
        webTotalCostCol = ColumnIndexToLetter(startingCol + 6)
        Dim webProfitLossCol As String
        webProfitLossCol = ColumnIndexToLetter(startingCol + 7)
        
        Dim storeSalePriceCol As String
        storeSalePriceCol = ColumnIndexToLetter(startingCol + 8)
        Dim storeDistCreditCol As String
        storeDistCreditCol = ColumnIndexToLetter(startingCol + 9)
        Dim storePaymentProcessorPercentCol As String
        storePaymentProcessorPercentCol = ColumnIndexToLetter(startingCol + 10)
        Dim storeTotalCostCol As String
        storeTotalCostCol = ColumnIndexToLetter(startingCol + 11)
        Dim storeProfitLossCol As String
        storeProfitLossCol = ColumnIndexToLetter(startingCol + 12)
        
        Dim webSalePriceFormula As String
        webSalePriceFormula = ""
        If xsheetMap.exists("WebPrice") Then
            webSalePriceFormula = "=ROUND(" & ColumnIndexToLetter(xsheetMap.item("WebPrice")) & "%ROW%*" & ((100 - Me.salePercBox) / 100) & ", 2)"
        End If
        Dim storeSalePriceFormula As String
        storeSalePriceFormula = ""
        If xsheetMap.exists("StorePrice") Then
            storeSalePriceFormula = "=ROUND(" & ColumnIndexToLetter(xsheetMap.item("StorePrice")) & "%ROW%*" & ((100 - Me.salePercBox) / 100) & ", 2)"
        End If
        
        Dim webDistCreditFormula As String
        webDistCreditFormula = ""
        If xsheetMap.exists("WebPrice") Then
            webDistCreditFormula = "=ROUND(" & webSalePriceCol & "%ROW%*" & (Me.creditPercBox / 100) & ", 2)"
        End If
        Dim storeDistCreditFormula As String
        storeDistCreditFormula = ""
        If xsheetMap.exists("StorePrice") Then
            storeDistCreditFormula = "=ROUND(" & storeSalePriceCol & "%ROW%*" & (Me.creditPercBox / 100) & ", 2)"
        End If
        
        Dim webShippingChargeFormula As String
        webShippingChargeFormula = ""
        If xsheetMap.exists("ShippingType") Then
            webShippingChargeFormula = "=IF(OR(" & ColumnIndexToLetter(xsheetMap.item("ShippingType")) & "%ROW%=""Free""," & webSalePriceCol & "%ROW%>=150),0,8.5)"
        End If
        
        Dim webPaymentProcessorPercentFormula As String
        webPaymentProcessorPercentFormula = ""
        If xsheetMap.exists("WebPrice") Then
            webPaymentProcessorPercentFormula = "=ROUND(" & webSalePriceCol & "%ROW%*0.03" & ", 2)"
        End If
        Dim storePaymentProcessorPercentFormula As String
        storePaymentProcessorPercentFormula = ""
        If xsheetMap.exists("StorePrice") Then
            storePaymentProcessorPercentFormula = "=ROUND(" & storeSalePriceCol & "%ROW%*0.025" & ", 2)"
        End If
        
        Dim webHostingPercentFormula As String
        webHostingPercentFormula = ""
        If xsheetMap.exists("WebPrice") Then
            webHostingPercentFormula = "=ROUND(" & webSalePriceCol & "%ROW%*0.0075" & ", 2)"
        End If
        'no store
        
        Dim webShippingCostFormula As String
        webShippingCostFormula = ""
        If xsheetMap.exists("AverageFreight") And xsheetMap.exists("D/SCharges(Commercial)") And xsheetMap.exists("IsDropshippable") And xsheetMap.exists("IsDropshipOnly") Then
            'if dropship only:
            '  d/s
            'else
            '  if dropshippable and d/s filled:
            '    if average filled:
            '      min(d/s, average)
            '    else:
            '      d/s
            '  else (not d/s or d/s empty):
            '    average
            webShippingCostFormula = "=ROUND(IF(" & ColumnIndexToLetter(xsheetMap.item("IsDropshipOnly")) & "%ROW%=""Y""," & ColumnIndexToLetter(xsheetMap.item("D/SCharges(Commercial)")) & "%ROW%,IF(AND(" & ColumnIndexToLetter(xsheetMap.item("IsDropshippable")) & "%ROW%=""Y""," & ColumnIndexToLetter(xsheetMap.item("D/SCharges(Commercial)")) & "%ROW%>0),IF(" & ColumnIndexToLetter(xsheetMap.item("AverageFreight")) & "%ROW%>0,MIN(" & ColumnIndexToLetter(xsheetMap.item("D/SCharges(Commercial)")) & "%ROW%," & ColumnIndexToLetter(xsheetMap.item("AverageFreight")) & "%ROW%)," & ColumnIndexToLetter(xsheetMap.item("D/SCharges(Commercial)")) & "%ROW%)," & ColumnIndexToLetter(xsheetMap.item("AverageFreight")) & "%ROW%)),2)"
        ElseIf xsheetMap.exists("D/SCharges(Commercial)") Then
            webShippingCostFormula = "=ROUND(" & ColumnIndexToLetter(xsheetMap.item("D/SCharges(Commercial)")) & "%ROW%,2)"
        ElseIf xsheetMap.exists("AverageFreight") Then
            webShippingCostFormula = "=ROUND(" & ColumnIndexToLetter(xsheetMap.item("AverageFreight")) & "%ROW%,2)"
        End If
        
        Dim webTotalCostFormula As String
        Dim webCostParts As PerlArray
        Set webCostParts = New PerlArray
        If xsheetMap.exists("StandardCost") Then
            webCostParts.Push ColumnIndexToLetter(xsheetMap.item("StandardCost")) & "%ROW%"
        End If
        webCostParts.Push webPaymentProcessorPercentCol & "%ROW%"
        webCostParts.Push webHostingPercentCol & "%ROW%"
        webCostParts.Push webShippingCostCol & "%ROW%"
        webTotalCostFormula = "=ROUND(" & webCostParts.Join("+") & "-" & webDistCreditCol & "%ROW%, 2)"
        Dim storeTotalCostFormula As String
        Dim storeCostParts As PerlArray
        Set storeCostParts = New PerlArray
        If xsheetMap.exists("StandardCost") Then
            storeCostParts.Push ColumnIndexToLetter(xsheetMap.item("StandardCost")) & "%ROW%"
        End If
        storeCostParts.Push storePaymentProcessorPercentCol & "%ROW%"
        storeTotalCostFormula = "=ROUND(" & storeCostParts.Join("+") & "-" & storeDistCreditCol & "%ROW%, 2)"
        
        Dim webProfitLossFormula As String
        webProfitLossFormula = "=ROUND(" & webSalePriceCol & "%ROW%+" & webShippingChargeCol & "%ROW%-" & webTotalCostCol & "%ROW%, 2)"
        Dim storeProfitLossFormula As String
        storeProfitLossFormula = "=ROUND(" & storeSalePriceCol & "%ROW%-" & storeTotalCostCol & "%ROW%, 2)"
        
        excelapp.ActiveWorkbook.ActiveSheet.Range(webSalePriceCol & "1").value = "Web Sale Price"
        excelapp.ActiveWorkbook.ActiveSheet.Range(webDistCreditCol & "1").value = "Web Distributor Credit"
        excelapp.ActiveWorkbook.ActiveSheet.Range(webShippingChargeCol & "1").value = "Web Shipping Charge"
        excelapp.ActiveWorkbook.ActiveSheet.Range(webPaymentProcessorPercentCol & "1").value = "Web Payment Processor 3%"
        excelapp.ActiveWorkbook.ActiveSheet.Range(webHostingPercentCol & "1").value = "Web Hosting 0.75%"
        excelapp.ActiveWorkbook.ActiveSheet.Range(webShippingCostCol & "1").value = "Web Shipping Cost"
        excelapp.ActiveWorkbook.ActiveSheet.Range(webTotalCostCol & "1").value = "Web Total Cost"
        excelapp.ActiveWorkbook.ActiveSheet.Range(webProfitLossCol & "1").value = "Web Profit/Loss"
        
        excelapp.ActiveWorkbook.ActiveSheet.Range(storeSalePriceCol & "1").value = "Store Sale Price"
        excelapp.ActiveWorkbook.ActiveSheet.Range(storeDistCreditCol & "1").value = "Store Distributor Credit"
        excelapp.ActiveWorkbook.ActiveSheet.Range(storePaymentProcessorPercentCol & "1").value = "Store Payment Processor 2.5%"
        excelapp.ActiveWorkbook.ActiveSheet.Range(storeTotalCostCol & "1").value = "Store Total Cost"
        excelapp.ActiveWorkbook.ActiveSheet.Range(storeProfitLossCol & "1").value = "Store Profit/Loss"
        
        currRow = 2
        While excelapp.ActiveWorkbook.ActiveSheet.Range("A" & currRow).value <> ""
            excelapp.ActiveWorkbook.ActiveSheet.Range(webSalePriceCol & currRow).value = Replace(webSalePriceFormula, "%ROW%", currRow)
            excelapp.ActiveWorkbook.ActiveSheet.Range(webDistCreditCol & currRow).value = Replace(webDistCreditFormula, "%ROW%", currRow)
            excelapp.ActiveWorkbook.ActiveSheet.Range(webShippingChargeCol & currRow).value = Replace(webShippingChargeFormula, "%ROW%", currRow)
            excelapp.ActiveWorkbook.ActiveSheet.Range(webPaymentProcessorPercentCol & currRow).value = Replace(webPaymentProcessorPercentFormula, "%ROW%", currRow)
            excelapp.ActiveWorkbook.ActiveSheet.Range(webHostingPercentCol & currRow).value = Replace(webHostingPercentFormula, "%ROW%", currRow)
            excelapp.ActiveWorkbook.ActiveSheet.Range(webShippingCostCol & currRow).value = Replace(webShippingCostFormula, "%ROW%", currRow)
            excelapp.ActiveWorkbook.ActiveSheet.Range(webTotalCostCol & currRow).value = Replace(webTotalCostFormula, "%ROW%", currRow)
            excelapp.ActiveWorkbook.ActiveSheet.Range(webProfitLossCol & currRow).value = Replace(webProfitLossFormula, "%ROW%", currRow)
            
            excelapp.ActiveWorkbook.ActiveSheet.Range(storeSalePriceCol & currRow).value = Replace(storeSalePriceFormula, "%ROW%", currRow)
            excelapp.ActiveWorkbook.ActiveSheet.Range(storeDistCreditCol & currRow).value = Replace(storeDistCreditFormula, "%ROW%", currRow)
            excelapp.ActiveWorkbook.ActiveSheet.Range(storePaymentProcessorPercentCol & currRow).value = Replace(storePaymentProcessorPercentFormula, "%ROW%", currRow)
            excelapp.ActiveWorkbook.ActiveSheet.Range(storeTotalCostCol & currRow).value = Replace(storeTotalCostFormula, "%ROW%", currRow)
            excelapp.ActiveWorkbook.ActiveSheet.Range(storeProfitLossCol & currRow).value = Replace(storeProfitLossFormula, "%ROW%", currRow)
            
            currRow = currRow + 1
        Wend
    End If
    
    excelapp.ActiveWorkbook.ActiveSheet.Columns(ColumnIndexToLetter(xsheetMap.item("PartNumber"))).HorizontalAlignment = Excel.Constants.xlLeft
    excelapp.Windows(1).SplitRow = 1
    excelapp.Windows(1).FreezePanes = True
    excelapp.ActiveWorkbook.ActiveSheet.Columns.AutoFit
    
    excelapp.Visible = True
    
    If Not stopNow Then
        If getItemSource() = "inventorymaintenance" Then
            If vbYes = MsgBox("Clear item list?", vbYesNo) Then
                InventoryMaintenance.ClearXSheetItemList
                Unload Me
            End If
        End If
    End If
    
theend:
    Me.lblStatusBar.Caption = ""
    lockForm False
    Mouse.Hourglass False
    If Not excelapp.Visible Then
        excelapp.Visible = True
    End If
    Set excelapp = Nothing
    Exit Sub
errh:
    MsgBox "ERROR, call Brian right now, otherwise Excel might be messed up until you reboot: " & vbCrLf & vbCrLf & Err.Number & vbCrLf & vbCrLf & Err.Description
    lockForm False
    Mouse.Hourglass False
    If Not excelapp Is Nothing Then
        excelapp.Visible = True
        Set excelapp = Nothing
    End If
End Sub

Private Function resolveControlIndex(dx As Long, resolveAs As Long) As Variant
    Dim retval As Variant
    retval = lookupColumnList(dx)

    Dim i As Long
    Select Case resolveAs
        'header names
        Case Is = 0
            resolveControlIndex = retval
        'sql column names
        Case Is = 1
            For i = 0 To UBound(retval)
                retval(i) = Replace(retval(i), " ", "")
            Next i
            resolveControlIndex = retval
        Case Else
            Err.Raise 124, "resolveControlIndex", "TELL BRIAN: missing resolve index " & resolveAs
    End Select
End Function

Private Sub lockForm(Optional yesno As Boolean = True)
    Me.Frame1.Enabled = Not yesno
    Me.Frame2.Enabled = Not yesno
    Me.Frame3.Enabled = Not yesno
    Me.btnGenerate.Enabled = Not yesno
    Me.btnCancel.Caption = IIf(yesno, "Stop Generation", "Cancel")
End Sub

Private Sub chkColumn_Click(index As Integer)
    Select Case index
        Case Is = 30
            If Me.chkColumn(30) Then    ' ebay commission requires ebay price
                Me.chkColumn(31) = 1
            End If
    End Select
End Sub

Private Sub Form_Load()
    lookupFill
    
    requeryLineCodes
    Select Case True
        Case Is = Me.opXSheetType(0) 'new products spreadsheet
            opXSheetType_Click 0
        Case Is = Me.opXSheetType(1) 'regular spreadsheet
            opXSheetType_Click 1
        Case Is = Me.opXSheetType(2) 'vendor spreadsheet
            opXSheetType_Click 2
        Case Is = Me.opXSheetType(3) 'pre-sales spreadsheet
            opXSheetType_Click 3
        Case Else
            MsgBox "missing spreadsheet type in select in form load!"
    End Select
End Sub

Private Sub opXSheetType_Click(index As Integer)
    setColumnsEnabled
    setCheckedness
    setSideAreaVisibility
End Sub

Private Sub setColumnsEnabled()
    Dim ctl As Control
    For Each ctl In Me.chkColumn
        If lookupSpecialness(ctl.index) Then
            'special, always disabled/checked
        Else
            Select Case True
                Case Is = Me.opXSheetType(0) 'new products spreadsheet
                    If lookupNewProdEnabled(ctl.index) Then
                        Enable ctl
                    Else
                        ctl.value = 0
                        Disable ctl
                    End If
                Case Is = Me.opXSheetType(1), Me.opXSheetType(2), Me.opXSheetType(3) 'regular spreadsheet, vendor spreadsheet, pre-sales spreadsheet
                    Enable ctl
                Case Else
                    MsgBox "missing spreadsheet type in select in setColumnsEnabled()!"
            End Select
        End If
    Next ctl
End Sub

Private Sub setCheckedness()
    Dim ctl As Control
    For Each ctl In Me.chkColumn
        If lookupSpecialness(ctl.index) Then
            'special, always disabled/checked
        Else
            Select Case True
                Case Is = Me.opXSheetType(0), Me.opXSheetType(1) 'new products spreadsheet or regular spreadsheet
                    'do nothing, just leave the way it is
                Case Is = Me.opXSheetType(2) 'vendor spreadsheet
                    ctl.value = IIf(lookupVendorAutoCheck(ctl.index), 1, 0)
                Case Is = Me.opXSheetType(3) 'pre-sale analysis spreadsheet
                    ctl.value = IIf(lookupPreSaleAutoCheck(ctl.index), 1, 0)
                Case Else
                    MsgBox "missing spreadsheet type in select in setColumnsEnabled()!"
            End Select
        End If
    Next ctl
End Sub

Private Sub setSideAreaVisibility()
    Select Case True
        Case Is = Me.opXSheetType(0), Me.opXSheetType(1), Me.opXSheetType(3) 'new products spreadsheet, regular spreadsheet, presales spreadsheet
            Me.width = 5205
        Case Is = Me.opXSheetType(2) 'vendor spreadsheet
            Me.width = 6690
        Case Else
            MsgBox "missing spreadsheet type in select in setSideAreaVisibility()!"
    End Select
End Sub

Private Sub requeryLineCodes()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ProductLine FROM vRequeryLineCodes ORDER BY ProductLine")
    Me.LineCodeList.Clear
    While Not rst.EOF
        Me.LineCodeList.AddItem rst("ProductLine")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub VendorPreset_Click(index As Integer)
    Dim repID As Long
    repID = Me.VendorPreset(index).tag
    
    Dim i As Long
    For i = 0 To Me.LineCodeList.ListCount - 1
        Me.LineCodeList.Selected(i) = False
    Next i
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ProductLine FROM ProductLine WHERE ID IN (SELECT ProductLineID FROM ProductLineSalesReps WHERE SalesRepID=" & repID & ") ORDER BY ProductLine")
    While Not rst.EOF
        For i = 0 To Me.LineCodeList.ListCount - 1
            'silently drops unfound line codes? should only be # hidden ones anyway
            If Me.LineCodeList.list(i) = rst("ProductLine") Then
                Me.LineCodeList.Selected(i) = True
                Exit For
            End If
        Next i
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Function getItemSource() As String
    Select Case True
        Case Is = Me.opXSheetType(0), Me.opXSheetType(1), Me.opXSheetType(3) 'new products spreadsheet, regular spreadsheet, pre-sales spreadsheet
            getItemSource = "inventorymaintenance"
        Case Is = Me.opXSheetType(2) 'vendor spreadsheet
            getItemSource = "linecodelist"
        Case Else
            MsgBox "missing spreadsheet type in select in getItemSource()!"
    End Select
End Function

Private Function loadSourceItems() As PerlArray
    Dim items As PerlArray
    'Set items = New PerlArray
    Dim i As Long
    Select Case getItemSource()
        Case Is = "inventorymaintenance"
            'Dim temp() As String
            'temp = InventoryMaintenance.GetXSheetItemList()
            'For i = 0 To UBound(temp)
            '    If temp(i) <> "" Then
            '        items.Push CStr(temp(i))
            '    End If
            'Next i
            Set items = InventoryMaintenance.GetXSheetItemList()
        Case Is = "linecodelist"
            Set items = New PerlArray
            Dim rst As ADODB.Recordset
            For i = 0 To Me.LineCodeList.ListCount - 1
                If Me.LineCodeList.Selected(i) Then
                    Set rst = DB.retrieve("SELECT ItemNumber FROM InventoryMaster WHERE Hide=0 AND ProductLine='" & Me.LineCodeList.list(i) & "'")
                    While Not rst.EOF
                        items.Push CStr(rst("ItemNumber"))
                        rst.MoveNext
                    Wend
                    rst.Close
                    Set rst = Nothing
                End If
            Next i
    End Select
    Set loadSourceItems = items
End Function
