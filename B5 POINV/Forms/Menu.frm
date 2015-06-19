VERSION 5.00
Begin VB.Form Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Path Editor"
      Height          =   375
      Left            =   2730
      TabIndex        =   8
      Top             =   570
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton btnGeneric 
      Caption         =   "generic button for stuff"
      Height          =   315
      Left            =   210
      TabIndex        =   7
      Top             =   2610
      Width           =   1905
   End
   Begin VB.CommandButton btnBackExit 
      Caption         =   "exit or back"
      Height          =   435
      Left            =   1200
      TabIndex        =   4
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton btnSomething 
      Caption         =   "button"
      Height          =   435
      Index           =   0
      Left            =   1230
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblGeneric2 
      Caption         =   "generic label for stuff"
      Height          =   255
      Left            =   270
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblGeneric1 
      Caption         =   "generic label for stuff"
      Height          =   255
      Left            =   270
      TabIndex        =   5
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "menu title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   270
      TabIndex        =   2
      Top             =   90
      Width           =   4305
   End
   Begin VB.Label lblUserOverride 
      Alignment       =   1  'Right Justify
      Height          =   225
      Left            =   3420
      TabIndex        =   1
      Top             =   5430
      Width           =   1425
   End
   Begin VB.Label lblStatusBar 
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   5430
      Width           =   3195
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ImportTimer As Timer
Attribute ImportTimer.VB_VarHelpID = -1

' number of buttons in a column
Private Const COL_MAX As Long = 6

' y-axis starting position, offset between buttons, and button width
Private Const START_Y As Long = 1200
Private Const OFFSET_Y As Long = 600
Private Const WIDTH_Y As Long = 435

' x-axis positions/offsets, each set is for different #'s of columns
Private Const START_X_1 As Long = 1230
Private Const OFFSET_X_1 As Long = 0
Private Const WIDTH_X_1 As Long = 2415

Private Const START_X_2 As Long = 210
Private Const OFFSET_X_2 As Long = 2490
Private Const WIDTH_X_2 As Long = 1965

Private Const START_X_3 As Long = 60
Private Const OFFSET_X_3 As Long = 1635
Private Const WIDTH_X_3 As Long = 1485


Private Sub btnBackExit_Click()
    If Me.btnBackExit.tag = "_Exit" Then
        Unload Me
    Else
        switchToMenu Mid(Me.btnBackExit.tag, 2)
    End If
End Sub

Private Sub btnGeneric_Click()
    Select Case Me.Caption
        Case Is = "MAS 200 Import / Export"
            MasUnlockClick
        Case Else
            MsgBox "Error in click event: function mapping not defined for " & qq(Me.Caption) & " menu."
    End Select
End Sub

Private Sub btnSomething_Click(index As Integer)
    handleButton CLng(index)
End Sub

Private Sub Command1_Click()
    Load AddWebPaths2
    AddWebPaths2.Show
End Sub

'Private Sub Command2_Click()
'    Load Patherizer
'    Patherizer.Show
'End Sub

Private Sub Form_Load()
    'If USE_ALT_DATABASE Then
    '    Me.BackColor = &H80&
    '    Me.Command2.Visible = True
    '    Me.Command1.Visible = True
    'End If
    
    switchToMenu "Main"
    Me.lblTitle.ToolTipText = "Version: " & app.Major & "." & app.Minor & "." & app.Revision
    Me.lblUserOverride.Caption = Environ("UserName")
    
    setAutoImportLabel
    
    Set ImportTimer = Controls.Add("VB.Timer", "ImportTimer")
    If WARN_ON_IMPORT Then
        ImportTimer_Timer 'start off the event right now
        ImportTimer.Interval = 60000
        ImportTimer.Enabled = True
    Else
        ImportTimer.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    For Each frm In Forms
        If frm.name <> Me.name Then
            If vbNo = MsgBox(frm.Caption & " is still open, close anyway?", vbYesNo) Then
                Cancel = 1
                Exit Sub
            End If
        End If
    Next frm
    For Each frm In Forms
        If frm.name <> Me.name Then
            Unload frm
            Set frm = Nothing
        End If
    Next frm
    
    'If Not RunningInIDE Then
    '    LogOutPoInv POINV_SESSION_ID
    'End If
    
    cleanup
End Sub

Private Sub switchToMenu(menuName)
    clearGenerics
    ' this uses three parallel arrays to pass button information.
    ' formNameArray is either the name of the form to load, or an
    ' internal menu name to load, prefaced with an underscore.
    ' buttonCaptionArray is the list of captions for each button
    ' in sequence. moduleRequiredArray is the permissions module
    ' needed to access the form/menu. see the UserModules table
    ' for a list of possible modules.
    '
    ' everything with an underscore should have an entry in the
    ' switch below.
    '
    ' everything with an asterisk should have an entry in
    ' handleButton()'s switch.
    '
    ' menus using the generic controls should set everything
    ' (size/vis/etc) in a separate function, then add to the
    ' event handlers if necessary.
    Dim formNameArray As Variant
    Dim buttonCaptionArray As Variant
    Dim moduleRequiredArray As Variant
    Select Case menuName
        Case Is = "Main"
            Me.Caption = "Main Menu"
            Me.lblTitle.Caption = "Main Menu"
            Me.btnBackExit.Caption = "Exit"
            Me.btnBackExit.tag = "_Exit"
            formNameArray = Array("InventoryMaintenance", _
                                  "SignMaintenance", _
                                  "_MAS200", _
                                  "_CustomerService", _
                                  "_Website", _
                                  "_Administrative")
            buttonCaptionArray = Array("POinv", _
                                       "Signs", _
                                       "MAS 200 Import/Export", _
                                       "Customer Service", _
                                       "Website", _
                                       "Administrative")
            moduleRequiredArray = Array("InventoryMaintenanceRead", _
                                        "SignMaintenanceRead", _
                                        "MasImport", _
                                        "", _
                                        "", _
                                        "")
        Case Is = "MAS200"
            setMas200Generics
            Me.Caption = "MAS 200 Import / Export"
            Me.lblTitle.Caption = "MAS 200 Import / Export"
            Me.btnBackExit.Caption = "Back to Main Menu"
            Me.btnBackExit.tag = "_Main"
            formNameArray = Array("*MASImport", _
                                  "*MASExport", _
                                  "*MASRetry")
            buttonCaptionArray = Array("Import Data From MAS 200", _
                                       "Export Changes To MAS 200", _
                                       "Retry Last Export To MAS 200")
            moduleRequiredArray = Array("MasImport", _
                                        "MasImport", _
                                        "MasImport")
        Case Is = "CustomerService"
            Me.Caption = "Customer Service Menu"
            Me.lblTitle.Caption = "Customer Service Menu"
            Me.btnBackExit.Caption = "Back to Main Menu"
            Me.btnBackExit.tag = "_Main"
            formNameArray = Array("*PrintSigns", _
                                  "Barcode", _
                                  "TransferRequest", _
                                  "*FreightCalc", _
                                  "QuoteEntry", _
                                  "BlogManager", _
                                  "AddCrossSells", _
                                  "SalesRepMaintenance", _
                                  "WarrantyMaintenance")
            buttonCaptionArray = Array("Print Signs", _
                                       "Barcode", _
                                       "Inventory Transfer Request", _
                                       "Freight Calculator", _
                                       "Order Quote", _
                                       "Blog Manager", _
                                       "Accessories", _
                                       "Sales Rep Information", _
                                       "Warranty Information")
            moduleRequiredArray = Array("PrintSigns", _
                                        "Barcode", _
                                        "InventoryTransfer", _
                                        "", _
                                        "OrderQuote", _
                                        "Blogs", _
                                        "Accessorize", _
                                        "", _
                                        "")
        Case Is = "Website"
            Me.Caption = "Website Menu"
            Me.lblTitle.Caption = "Website Menu"
            Me.btnBackExit.Caption = "Back to Main Menu"
            Me.btnBackExit.tag = "_Main"
            formNameArray = Array("WebOffload", _
                                  "ProductSuggestionUtility", _
                                  "_NewProducts", _
                                  "CustomerQuotes", _
                                  "_Keywords", _
                                  "FAQBuilder", _
                                  "AddWebPaths2", _
                                  "GroupItemEditor")
            buttonCaptionArray = Array("Web Offload", _
                                       "Product Suggestions", _
                                       "Import New Products", _
                                       "Customer Quotes", _
                                       "Keywords", _
                                       "FAQ Builder", _
                                       "Web Path Editor", _
                                       "Group Item Page Editor")
            moduleRequiredArray = Array("WebOffload", _
                                        "WebOffload", _
                                        "NewProducts", _
                                        "WebOffload", _
                                        "Keywords", _
                                        "WebOffload", _
                                        "SignMaintenanceWrite", _
                                        "SignMaintenanceWrite")
        Case Is = "Administrative"
            Me.Caption = "Administrative Menu"
            Me.lblTitle.Caption = "Administrative Menu"
            Me.btnBackExit.Caption = "Back to Main Menu"
            Me.btnBackExit.tag = "_Main"
            formNameArray = Array("UserPermissions", _
                                  "WarehouseLocationMaintenance", _
                                  "InventoryCountMgmt", _
                                  "ShowSaleCalculations", _
                                  "GeneralSpreadsheetImport", _
                                  "*FreightCalculation", _
                                  "*WebPriceComparison", _
                                  "InventorySalesView")
            buttonCaptionArray = Array("Users and Permissions", _
                                       "Warehouse and Location Maintenance", _
                                       "Inventory Count Management", _
                                       "Show/Sale Calculations", _
                                       "General Spreadsheet Importer", _
                                       "Run Freight Calculation", _
                                       "Run Web Price Comparison", _
                                       "Inventory Sales Report")
            moduleRequiredArray = Array("UserMaintenance", _
                                        "UserMaintenance", _
                                        "InventoryCount", _
                                        "UserMaintenance", _
                                        "NewProducts", _
                                        "FreightEstimation", _
                                        "PriceComparison", _
                                        "")
        Case Is = "Keywords"
            DB.Execute "EXEC spKeywordsChildrenCleanup"
            Me.Caption = "Keywords Menu"
            Me.lblTitle = "Keywords Menu"
            Me.btnBackExit.Caption = "Back to Web Menu"
            Me.btnBackExit.tag = "_Website"
            formNameArray = Array("KeywordsPhraseList", _
                                  "KeywordsCostRevenue", _
                                  "KeywordsImportData", _
                                  "_KeywordsOverture", _
                                  "_KeywordsAdwords")
            buttonCaptionArray = Array("Keyword Maintenance", _
                                       "Cost / Revenue Analysis", _
                                       "Import New Data", _
                                       "Overture", _
                                       "AdWords")
            moduleRequiredArray = Array("Keywords", _
                                        "Keywords", _
                                        "Keywords", _
                                        "Keywords", _
                                        "Keywords")
        Case Is = "KeywordsOverture"
            Me.Caption = "Keywords Menu - Overture"
            Me.lblTitle = "Keywords Menu - Overture"
            Me.btnBackExit.Caption = "Back to Keywords Menu"
            Me.btnBackExit.tag = "_Keywords"
            formNameArray = Array("*KeywordsExportToOverture", _
                                  "KeywordsClearOvertureFlags", _
                                  "KeywordsToBeDeleted")
            buttonCaptionArray = Array("Export to Overture", _
                                       "Clear Overture Flags", _
                                       "To Be Deleted")
            moduleRequiredArray = Array("Keywords", _
                                        "Keywords", _
                                        "Keywords")
        Case Is = "KeywordsAdwords"
            Me.Caption = "Keywords Menu - AdWords"
            Me.lblTitle = "Keywords Menu - AdWords"
            Me.btnBackExit.Caption = "Back to Keywords Menu"
            Me.btnBackExit.tag = "_Keywords"
            formNameArray = Array("*KeywordsExportToAdwords")
            buttonCaptionArray = Array("Export to AdWords")
            moduleRequiredArray = Array("Keywords")
        Case Is = "NewProducts"
            Me.Caption = "Import New Products Menu"
            Me.lblTitle.Caption = "Import New Products Menu"
            Me.btnBackExit.Caption = "Back to Web Menu"
            Me.btnBackExit.tag = "_Website"
            formNameArray = Array("*ImportNewSpreadsheet", _
                                  "NewProdDetailForm", _
                                  "*OpenExcelSpreadsheet")
            buttonCaptionArray = Array("Import New Spreadsheet", _
                                       "View Current Items", _
                                       "Open Excel Spreadsheet")
            moduleRequiredArray = Array("NewProducts", _
                                        "NewProducts", _
                                        "NewProducts")
        Case Else
            MsgBox "Error in Menu.switchToMenu(menuName := " & menuName & "): menu with that name is not defined."
    End Select
    
    createButtons UBound(formNameArray)
    Dim i As Long
    For i = 0 To UBound(formNameArray)
        Me.btnSomething(i).tag = formNameArray(i) & "|" & moduleRequiredArray(i)
        Me.btnSomething(i).Caption = buttonCaptionArray(i)
        If moduleRequiredArray(i) = "" Then
            Me.btnSomething(i).Enabled = True
        Else
            Me.btnSomething(i).Enabled = HasPermissionsTo(CStr(moduleRequiredArray(i)))
        End If
    Next i
    
    formatMenu
    
    'set focus to the first button available (but not the first time, only subsequent loads)
    If Me.Visible Then
        For i = 0 To Me.btnSomething.UBound
            If Me.btnSomething(i).Enabled Then
                Me.btnSomething(i).SetFocus
                Exit For
            End If
        Next i
    End If
    
    'then, set tab order to something sane
    For i = 0 To Me.btnSomething.UBound
        If Me.btnSomething(i).Enabled Then
            Me.btnSomething(i).tabIndex = i
        End If
    Next i
    Me.btnBackExit.tabIndex = i
End Sub

Private Sub formatMenu()
    Dim i As Long
    Select Case Me.btnSomething.count
        Case Is <= COL_MAX
            formatColumn 0, Me.btnSomething.UBound, START_X_1, START_Y, WIDTH_X_1, WIDTH_Y, OFFSET_X_1, OFFSET_Y
        Case Is <= COL_MAX * 2
            formatColumn 0, COL_MAX - 1, START_X_2, START_Y, WIDTH_X_2, WIDTH_Y, OFFSET_X_2, OFFSET_Y
            formatColumn COL_MAX, Me.btnSomething.UBound, START_X_2, START_Y, WIDTH_X_2, WIDTH_Y, OFFSET_X_2, OFFSET_Y
        Case Is <= COL_MAX * 3
            formatColumn 0, COL_MAX - 1, START_X_3, START_Y, WIDTH_X_3, WIDTH_Y, OFFSET_X_3, OFFSET_Y
            formatColumn COL_MAX, 2 * COL_MAX - 1, START_X_3, START_Y, WIDTH_X_3, WIDTH_Y, OFFSET_X_3, OFFSET_Y
            formatColumn 2 * COL_MAX, Me.btnSomething.UBound, START_X_3, START_Y, WIDTH_X_3, WIDTH_Y, OFFSET_X_3, OFFSET_Y
        Case Else
            MsgBox "Error in Menu.formatMenu(): can't handle more than " & COL_MAX * 3 & " buttons on a menu."
    End Select
End Sub

Private Sub formatColumn(dxFrom As Long, dxTo As Long, xstart As Long, ystart As Long, xwidth As Long, ywidth As Long, xoffset As Long, yoffset As Long)
    Dim i As Long
    For i = dxFrom To dxTo
        Me.btnSomething(i).width = xwidth
        Me.btnSomething(i).Height = ywidth
        
        Me.btnSomething(i).Left = xstart + ((i \ COL_MAX) * xoffset)
        Me.btnSomething(i).Top = ystart + (((i - dxFrom) Mod COL_MAX) * yoffset)
        
        Me.btnSomething(i).Visible = True
    Next i
End Sub

Private Sub createButtons(max As Long)
    Dim i As Long
    For i = 0 To max
        If i <> 0 And Me.btnSomething.UBound < max And i > Me.btnSomething.UBound Then '0 always loaded
            Load Me.btnSomething(i)
        End If
    Next i
    For i = i To Me.btnSomething.UBound
        If i = 0 Then 'what happens if you remove everything from the collection?
            Me.btnSomething(i).Visible = False
        Else
            Unload Me.btnSomething(i)
        End If
    Next i
End Sub

Private Sub handleButton(dx As Long)
    Dim formName As String, moduleReqd As String
    formName = Left(Me.btnSomething(dx).tag, InStr(Me.btnSomething(dx).tag, "|") - 1)
    moduleReqd = Mid(Me.btnSomething(dx).tag, InStr(Me.btnSomething(dx).tag, "|") + 1)
    
    If moduleReqd <> "" Then
        If Not HasPermissionsTo(moduleReqd) Then
            MsgBox "You do not have permissions to access this module. Talk to Brian or Kirk."
            Exit Sub
        End If
    End If
    
    If formName = "_Exit" Then
        Unload Me
    ElseIf Left(formName, 1) = "_" Then
        'these are internal menu switches
        Mouse.Hourglass True
        switchToMenu Mid(formName, 2)
        Mouse.Hourglass False
    ElseIf Left(formName, 1) = "*" Then
        'other miscellaneous functions go here
        Select Case Mid(formName, 2)
            Case Is = "MASImport"
                MasImportClick
            Case Is = "MASExport"
                MasExportClick
            Case Is = "MASRetry"
                MasRetryClick
            Case Is = "FreightCalc"
                Shell FQ_STANDALONE_APP, vbNormalFocus
            Case Is = "KeywordsExportToOverture"
                Shell PERL & " " & KWD_Y_OFFLOAD, vbNormalFocus
            Case Is = "KeywordsExportToAdwords"
                ExportDataToAdwords
            Case Is = "OpenExcelSpreadsheet"
                OpenNewProductsSpreadsheet
            Case Is = "ImportNewSpreadsheet"
                ImportNewProductsSpreadsheet
            Case Is = "WebPriceComparison"
                Shell WPERL & " " & WPC_GUI_APP, vbNormalFocus
            Case Is = "FreightCalculation"
                Shell WPERL & " " & FQ_GUI_APP, vbNormalFocus
            Case Is = "PrintSigns"
                Shell WPERL & " " & SIGNS_APP, vbNormalFocus
            Case Else
                MsgBox "Error in Menu.handleButton(dx := " & dx & "): can't find pseudofunction " & qq(Mid(formName, 2))
        End Select
    Else
        'form loading
        Mouse.Hourglass True
        If IsFormLoaded(formName) Then
            If IsFormMinimized(formName) Then
                UnMinimizeForm formName
            Else
                FocusOnForm formName
            End If
        Else
            loadAndShowForm formName
        End If
        Mouse.Hourglass False
    End If
End Sub

Private Sub setAutoImportLabel()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT TimeStarted, TimeFinished, Success FROM EventLog WHERE Event='autoimport' AND TimeStarted>'" & Format(Now(), "Short Date") & "'")
    If rst.EOF Then
        Me.lblStatusBar.Caption = "Today's Automatic Import: FAILURE!"
        Me.lblStatusBar.FontBold = True
        Me.lblStatusBar.ToolTipText = "Autoimport was not found in the Event Log! No import done this morning?"
    Else
        If rst("Success") Then
            Me.lblStatusBar.Caption = "Today's Automatic Import: Success!"
            Me.lblStatusBar.ToolTipText = "Started: " & Format(rst("TimeStarted"), "Long Time") & "    Finished: " & Format(rst("TimeFinished"), "Long Time")
        Else
            Me.lblStatusBar.Caption = "Today's Automatic Import: FAILURE!"
            Me.lblStatusBar.FontBold = True
            Me.lblStatusBar.ToolTipText = "Autoimport was not successful! Either crashed, or just not finished yet?"
        End If
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub lblUserOverride_DblClick()
    If USING_SECURITY Then
        Dim passwd As String
        passwd = InputBox("Enter administrative password:")
        If passwd = "" Then
            'do nothing
        ElseIf passwd = "foobar" Then
            USING_SECURITY = False
            MsgBox "Lockout functions disabled, close and reopen all windows. I hope you know what you're doing."
        Else
            MsgBox "Invalid password, stop playing with this."
        End If
    Else
        USING_SECURITY = True
        MsgBox "Security functions enabled, close and reopen all windows."
    End If
End Sub

Private Sub ImportTimer_Timer()
    'some long-running jobs may want to turn off the timer temporarily, so check that first
    If WARN_ON_IMPORT Then
        ImportTimer.Enabled = False
        ReDim flags(1) As Boolean
        flags = ImportExportProcessFlags
        If IMPORT_GOING And Not flags(0) Then
            'import finished
            IMPORT_GOING = False
            MsgBox "Import has finished!"
        ElseIf EXPORT_GOING And Not flags(1) Then
            'export finished
            EXPORT_GOING = False
            If IsFormLoaded("InventoryMaintenance") Then
                InventoryMaintenance.LockForExport False
            End If
            If IsFormLoaded("PurchaseOrdersForm") Then
                PurchaseOrdersForm.LockForExport False
            End If
        ElseIf Not IMPORT_GOING And flags(0) Then
            'import started
            IMPORT_GOING = True
            MsgBox "An import has been started! Any changes you make will not flag the item to export to MAS 200! If you don't know what that means, talk to Brian or Kirk."
        ElseIf Not EXPORT_GOING And flags(1) Then
            'export started
            EXPORT_GOING = True
            If IsFormLoaded("InventoryMaintenance") Then
                InventoryMaintenance.LockForExport True
            End If
            If IsFormLoaded("PurchaseOrdersForm") Then
                PurchaseOrdersForm.LockForExport True
            End If
        End If
        ImportTimer.Enabled = True
    End If
End Sub

Private Sub clearGenerics()
    Me.lblGeneric1.Visible = False
    Me.lblGeneric2.Visible = False
    Me.btnGeneric.Visible = False
End Sub

Private Sub setMas200Generics()
    'checks the process flags, not the global booleans, just in case they're out of sync
    ReDim pflags(1) As Boolean
    pflags = ImportExportProcessFlags
    If pflags(0) Or pflags(1) Then
        'Me.btnImport.Enabled = False
        'Me.btnExport.Enabled = False
        'Me.btnRetryExport.Enabled = False
        'Me.lblCurrentlyRunning.Visible = True
        'Me.lblUserDate.Visible = True
        'Me.btnUnlock.Visible = True
        'Me.lblUserDate.caption = Replace(Me.lblUserDate.caption, "USERNAME", DLookup("UserName", "ProcessFlags", "Task='" & IIf(pflags(0), "import", "export") & "'"))
        'Me.lblUserDate.caption = Replace(Me.lblUserDate.caption, "DATETIME", DLookup("DateStamp", "ProcessFlags", "Task='" & IIf(pflags(0), "import", "export") & "'"))
        'Me.lblCurrentlyRunning.caption = Replace(Me.lblCurrentlyRunning.caption, "IMP/EXP", IIf(pflags(0), "import", "export"))
        Me.btnSomething(0).Enabled = False
        Me.btnSomething(1).Enabled = False
        Me.btnSomething(2).Enabled = False
        Me.lblGeneric1.Caption = "An " & IIf(pflags(0), "import", "export") & " is currently running."
        Me.lblGeneric1.FontBold = True
        Me.lblGeneric1.Alignment = vbCenter
        Me.lblGeneric1.Left = 690
        Me.lblGeneric1.Top = 3330
        Me.lblGeneric1.width = 3555
        Me.lblGeneric1.Height = 255
        Me.lblGeneric1.Visible = True
        Me.lblGeneric2.Caption = "Started by " & DLookup("UserName", "ProcessFlags", "Task='" & IIf(pflags(0), "import", "export") & "'") & " at " & DLookup("DateStamp", "ProcessFlags", "Task='" & IIf(pflags(0), "import", "export") & "'")
        Me.lblGeneric2.FontBold = False
        Me.lblGeneric2.Alignment = vbCenter
        Me.lblGeneric2.Left = 120
        Me.lblGeneric2.Top = 3630
        Me.lblGeneric2.width = 4695
        Me.lblGeneric2.Height = 255
        Me.lblGeneric2.Visible = True
        Me.btnGeneric.Caption = "Unlock"
        Me.btnGeneric.FontBold = False
        Me.btnGeneric.Left = 1830
        Me.btnGeneric.Top = 3960
        Me.btnGeneric.width = 1095
        Me.btnGeneric.Height = 225
        Me.btnGeneric.Visible = True
    End If
End Sub

Private Sub loadAndShowForm(formName As String)
    Select Case formName
        Case Is = "InventoryMaintenance"
            Load InventoryMaintenance
            InventoryMaintenance.Show
        Case Is = "SignMaintenance"
            Load SignMaintenance
            SignMaintenance.Show
        'Case Is = "PrintSigns"
        '    Load PrintSigns
        '    PrintSigns.Show
        Case Is = "Barcode"
            Load Barcode
            Barcode.Show
        'Case Is = "TransferRequest"
        '    Load TransferRequest
        '    TransferRequest.Show
        'Case Is = "QuoteInquiry"
        '    Load QuoteInquiry
        '    QuoteInquiry.Show
        Case Is = "QuoteEntry"
            Load QuoteEntry
            QuoteEntry.Show
        'Case Is = "BlogManager"
        '    Load BlogManager
        '    BlogManager.Show
        Case Is = "AddCrossSells"
            Load AddCrossSells
            AddCrossSells.Show
        Case Is = "WebOffload"
            Load WebOffload
            WebOffload.Show
        'Case Is = "WebSiteFrontPage"
        '    Load WebSiteFrontPage
        '    WebSiteFrontPage.Show
        Case Is = "CustomerQuotes"
            Load CustomerQuotes
            CustomerQuotes.Show
        Case Is = "FAQBuilder"
            Load FAQBuilder
            FAQBuilder.Show
        Case Is = "UserPermissions"
            Load UserPermissions
            UserPermissions.Show
        'Case Is = "WarehouseLocationMaintenance"
        '    Load WarehouseLocationMaintenance
        '    WarehouseLocationMaintenance.Show
        'Case Is = "InventoryCountMgmt"
        '    Load InventoryCountMgmt
        '    InventoryCountMgmt.Show
        Case Is = "ShowSaleCalculations"
            Load ShowSaleCalculations
            ShowSaleCalculations.Show
        Case Is = "GeneralSpreadsheetImport"
            Load GeneralSpreadsheetImport
            GeneralSpreadsheetImport.Show
        Case Is = "KeywordsPhraseList"
            Load KeywordsPhraseList
            KeywordsPhraseList.Show
        Case Is = "KeywordsCostRevenue"
            Load KeywordsCostRevenue
            KeywordsCostRevenue.Show
        Case Is = "KeywordsImportData"
            Load KeywordsImportData
            KeywordsImportData.Show
        Case Is = "KeywordsClearOvertureFlags"
            Load KeywordsClearOvertureFlags
            KeywordsClearOvertureFlags.Show
        Case Is = "KeywordsToBeDeleted"
            Load KeywordsToBeDeleted
            KeywordsToBeDeleted.Show
        Case Is = "NewProdDetailForm"
            Load NewProdDetailForm
            NewProdDetailForm.Show
        'Case Is = "WeightUpdate"
        '    Load WeightUpdate
        '    WeightUpdate.Show
        Case Is = "SalesRepMaintenance"
            Load SalesRepMaintenance
            SalesRepMaintenance.Show
        'Case Is = "SectionPageCaptionEditor"
        '    Load SectionPageCaptionEditor
        '    SectionPageCaptionEditor.Show
        'Case Is = "WriteupEditor"
        '    Load WriteupEditor
        '    WriteupEditor.Show
        Case Is = "WarrantyMaintenance"
            Load WarrantyMaintenance
            WarrantyMaintenance.Show
        Case Is = "ProductSuggestionUtility"
            Load ProductSuggestionUtility
            ProductSuggestionUtility.Show
        Case Is = "AddWebPaths2"
            Load AddWebPaths2
            AddWebPaths2.Show
        Case Is = "GroupItemEditor"
            Load GroupItemEditor
            GroupItemEditor.Show
        'Case Is = "GeneratorQuotes"
        '    Load GeneratorQuotes
        '    GeneratorQuotes.Show
        Case Is = "InventorySalesView"
            Load InventorySalesView
            InventorySalesView.Show
        Case Else
            MsgBox "Error in Menu.loadAndShowForm(formName := " & formName & "): form mapping not defined."
    End Select
End Sub

