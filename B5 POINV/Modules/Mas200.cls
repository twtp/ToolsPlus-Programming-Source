VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mas200ImportExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Mas200ImportExport
' DateTime  : 9/8/2005 12:11
' Author    : briandonorfio
' Purpose   : Mas200 import/export logic wrapped in a class, so it can fire off events
'             to update the progress.
'
'             Dependencies:
'               - Microsoft ActiveX Data Objects 2.8 Library
'               - Microsoft Scripting Runtime
'               - DatabaseFunctions + DBConn
'               - Logic
'               - SearchFunctions
'               - SyncShell
'               - Utilities
'---------------------------------------------------------------------------------------

Option Explicit

Private Const MUTEX_COMPLETE As String = "s:\mastest\mas90-signs\export\poinv_price.txt"
Private Const MUTEX_INV_COMP As String = "s:\mastest\mas90-signs\export\poinv_item.txt"
Private Const PO_EXPORT_FILE As String = "s:\mastest\mas90-signs\export\poinv_purch_ord.csv"
Private Const KIT_EXPORT_FILE As String = "s:\mastest\mas90-signs\export\poinv_kits.csv"

Private Const DROPSHIP_EXE As String = "s:\mastest\mas90-signs\purchasing\SendDropshipPO\SendDropshipPO.exe"

Private Const SYNC_POS_NEEDS_ESCAPE As Long = 0
Private Const SYNC_POS_IS_KEY As Long = 1
Private Const SYNC_POS_NEEDS_TRIM As Long = 2
Private Const SYNC_POS_FIELDNAME_SRC As Long = 3
Private Const SYNC_POS_FIELDNAME_DST As Long = 4

Public Event StatusChanged(status As String)
Public Event StepChanged(step As String)

Public Function GetSalesOrderInfo(soNum As String) As Variant
    Dim rst As ADODB.Recordset
    Set rst = MASDB.retrieve("SELECT ShipToName, ShipToAddress1, ShipToAddress2, ShipToCity, ShipToState, ShipToZipCode, EmailAddress, PaymentType, DepositAmt, UDF_SHIP_PHONE FROM SO_SalesOrderHeader WHERE SalesOrderNo='" & soNum & "'")
    If rst.EOF Then
        GetSalesOrderInfo = Null
    Else
        Dim retval(0 To 9) As Variant
        retval(0) = "" 'TODO: some way to tell if it's right SO/PO combo?
        retval(1) = Nz(rst("ShipToName"))
        retval(2) = Nz(rst("ShipToAddress1"))
        retval(3) = Nz(rst("ShipToAddress2"))
        retval(4) = Nz(rst("ShipToCity"))
        retval(5) = Nz(rst("ShipToState"))
        retval(6) = Nz(rst("ShipToZipCode"))
        retval(7) = Nz(rst("UDF_SHIP_PHONE"))
        retval(8) = Nz(rst("DepositAmt"), "0")
        retval(9) = Nz(rst("EmailAddress"))
        'If retval(7) = "" Then
        '    'ok, find the phone number a different way?
        '    rst.Close
        '    Set rst = MASDB.retrieve("SELECT CommentText FROM SO_SalesOrderDetail WHERE SalesOrderNo='" & soNum & "' AND ItemType='4'")
        '    While Not rst.EOF
        '        If Nz(rst("CommentText")) <> "" Then
        '            If InStr(1, rst("CommentText"), "shipping", vbTextCompare) <> 0 And InStr(1, rst("CommentText"), "phone", vbTextCompare) <> 0 Then
        '                Dim re As RegExp
        '                Set re = New RegExp
        '                re.Pattern = "\d{3}[-./ ]{0,1}\d{3}[-./ ]{0,1}\d{4}"
        '                If re.Test(rst("CommentText")) Then
        '                    Dim matches As MatchCollection
        '                    Set matches = re.Execute(rst("CommentText"))
        '                    retval(7) = matches.item(0).value
        '                    rst.MoveLast
        '                End If
        '            End If
        '        End If
        '        rst.MoveNext
        '    Wend
        'End If
        GetSalesOrderInfo = retval
    End If
    rst.Close
    Set rst = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : StartImport
' DateTime  : 9/21/2005 11:19
' Author    : briandonorfio
' Purpose   : Starts a regular, user-attended import. The export is considered part of
'             the overall import, and shouldn't be called beforehand.
'---------------------------------------------------------------------------------------
'
Public Sub StartImport()
    Dim starttime As Double
    starttime = Timer
    
    Dim eventID As Long
    eventID = EventLogEntry("import")
    
    SetProcessFlag "import"

    If Not StartExport(True) Then
        MsgBox "Export cancelled, so import cancelled."
        Exit Sub
    End If

    'DB.Execute "exec spImportPurgeAll"
    DB.Execute "exec spImportDisableTriggers"

    RaiseEvent StepChanged("Addresses")
    importVendorAddresses
    RaiseEvent StepChanged("PL")
    importProductLines
    RaiseEvent StepChanged("Vendors")
    importVendors
    importVendorTerms
    RaiseEvent StepChanged("Whse")
    importWarehouses
    RaiseEvent StepChanged("Inv")
    importInventory
    importItemKits
    RaiseEvent StepChanged("InvQ")
    importInvQtys
    RaiseEvent StepChanged("InvP")
    importInvPrices
    RaiseEvent StepChanged("SalesHist")
    calcSalesHistory

    'done automatically on the export now (no chance for user to cancel), so don't need to duplicate
    'dbExecute "exec spExportSaveReorderPoint"
    'dbExecute "exec spExportResetFlags"
    'dbExecute "exec spExportResetPLFlags"
    'dbExecute "exec spExportClearPOItems"
    'dbExecute "exec spExportClearPOs"
    
    DB.Execute "exec spImportEnableTriggers"
    
    ClearProcessFlag "import"
    
    EventLogUpdate eventID, True
    
    MsgBox "Import finished successfully!" & vbCrLf & vbCrLf & "Time taken: " & (Timer - starttime) / 60 & " minutes."
    
    RaiseEvent StepChanged("Done")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : StartAutomaticImport
' DateTime  : 9/21/2005 11:20
' Author    : briandonorfio
' Purpose   : Starts an automatic import. Similar to the regular import, but includes a
'             few flags on what to run (probably from a conf file).
'---------------------------------------------------------------------------------------
'
Public Sub StartAutomaticImport(Optional updateFreight As Boolean = True, _
                                Optional updateSalesRank As Boolean = True, _
                                Optional threshold As Dictionary = Nothing, _
                                Optional triadCodeTime As String = "", _
                                Optional triadCodeMetric As String = "", _
                                Optional forceOrderPointCalc As Boolean = False)

    Dim eventID As Long
    eventID = EventLogEntry("autoimport")
    
    SetProcessFlag "import"
    
    If DEBUG_MODE Then
        DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Starting Export' )"
    End If

    StartAutomaticExport forceOrderPointCalc
    
    If DEBUG_MODE Then
        DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Export Finished' )"
    End If

    'DB.Execute "exec spImportPurgeAll"
    DB.Execute "exec spImportDisableTriggers"
    
    If DEBUG_MODE Then
        DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Triggers Disabled' )"
    End If

    RaiseEvent StepChanged("Addresses")
    importVendorAddresses
    RaiseEvent StepChanged("PL")
    importProductLines
    RaiseEvent StepChanged("Vendors")
    importVendors
    importVendorTerms
    RaiseEvent StepChanged("Whse")
    importWarehouses
    
    If DEBUG_MODE Then
        DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Tables Imported' )"
    End If
    
    RaiseEvent StepChanged("Inv")
    importInventory
    importItemKits
    
    If DEBUG_MODE Then
        DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Inventory Imported' )"
    End If
    
    RaiseEvent StepChanged("InvQ")
    importInvQtys
    
    If DEBUG_MODE Then
        DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Quantities Imported' )"
    End If
    
    RaiseEvent StepChanged("InvP")
    importInvPrices
    
    If DEBUG_MODE Then
        DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Prices Imported' )"
    End If
    
    'all set with updating inventory, reset the flags, re-enable triggers,
    'and tell everyone it's ok to go back in.
    RaiseEvent StepChanged("InvDone")
    'DB.Execute "exec spExportSaveReorderPoint"
    DB.Execute "exec spExportResetFlags"
    DB.Execute "exec spExportResetPLFlags"
    DB.Execute "exec spImportEnableTriggers"
    ClearProcessFlag "import"
    
    If DEBUG_MODE Then
        DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Cleaned Up' )"
    End If
    
    RaiseEvent StepChanged("SalesHist")
    calcSalesHistory
    
    If DEBUG_MODE Then
        DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'History Done' )"
    End If
    
    'auto import does actual freight, regular import doesn't
    If updateFreight Then
        RaiseEvent StepChanged("FreightActual")
        calcFreightActual
    End If
    
    If DEBUG_MODE Then
        DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Freight Done' )"
    End If
    
    If updateSalesRank And Not threshold Is Nothing Then
        RaiseEvent StepChanged("SalesRank")
        calcSalesRank threshold
    End If
    
    If DEBUG_MODE Then
        DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Ranking Done' )"
    End If
    
    'completely done now, so log it.
    EventLogUpdate eventID, True

End Sub

'---------------------------------------------------------------------------------------
' Procedure : StartExport
' DateTime  : 9/21/2005 11:21
' Author    : briandonorfio
' Purpose   : Starts a user-attended export. Requires arguments as to whether it's part
'             of an export or not. Optional argument to skip processing, and just call
'             mas200 for the export with the old cached data.
'
'             Perform logic to know when the inventory part is finished, after that we
'             can run other functions, like the purchase orders.
'---------------------------------------------------------------------------------------
'
Public Function StartExport(fromImport As Boolean, Optional retry As Boolean = False) As Boolean
    Dim doPOs As Boolean
    
    Dim eventID As Long
    eventID = EventLogEntry("export")
    
    If Not fromImport Then
        SetProcessFlag "export"
    End If

    RaiseEvent StepChanged("ExportPrep")
    If Not retry Then
        DB.Execute "exec spExportPurgeAll"

        DB.Execute "exec spExportProductLine"
        DB.Execute "exec spExportInventory"
        'exportItemWhseLines
        DB.Execute "exec spExportInventoryVendor"
        'DB.Execute "exec spExportPriceLevel1"
        'DB.Execute "exec spExportPriceLevel2"
        'DB.Execute "exec spExportPriceLevelA"
        'DB.Execute "exec spExportPriceLevelB"
        'DB.Execute "exec spExportPriceLevelG"
        'DB.Execute "exec spExportPriceLevelH"
        DB.Execute "exec spExportPriceLevels"
        'DB.Execute "exec spExportReorderPoint"
        Call exportKits
        If exportPOs() Then
            doPOs = True
        Else
            doPOs = False
        End If
    Else
        doPOs = CBool(DLookup("COUNT(*)", "ExportPurchaseOrders"))
    End If
    
    If MsgBox(getQtyToExportStr(), vbYesNo) = vbYes Then
        RaiseEvent StepChanged("Export")
        
        Dim fso As FileSystemObject
        Set fso = New FileSystemObject
        fso.CreateTextFile MUTEX_COMPLETE
        'fso.CreateTextFile MUTEX_INV_COMP

        ShellWait "s:\mastest\mas90-signs\mas200_import_items.bat", vbHide
        If doPOs Then
            'as of mas 4.5, we get "another task is attempting to start up" errors when
            'trying to run this simultaneously. so instead, it's all one at a time.
            While fso.FileExists(MUTEX_COMPLETE)
                DoEvents
                RaiseEvent StatusChanged("Waiting for inventory to finish...")
                Sleep 5000
                If DEBUG_MODE Then
                    DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Waiting For Mas200' )"
                End If
            Wend
            ShellWait "s:\mastest\mas90-signs\mas200_import_pos.bat", vbHide
'            'Dim foo As Boolean
'            'foo = exportItemWhseLines
'            'wait for the mutex file to be deleted
'            While fso.FileExists(MUTEX_INV_COMP)
'                DoEvents
'                RaiseEvent StatusChanged("Waiting for inventory to finish...")
'                Sleep 5000
'                If DEBUG_MODE Then
'                    DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Waiting For Mas200' )"
'                End If
'            Wend
'            ''ok, now we can do the other stuff
'            'If foo Then
'            '    ShellWait "s:\mastest\mas90-signs\mas200_import_item_whses.bat", vbHide
'            'End If
'            ShellWait "s:\mastest\mas90-signs\mas200_import_pos.bat", vbHide
        End If

        MsgBox "Wait until all MAS 200 import windows are finished before continuing."

        If Not fromImport Then
            If MsgBox("Do you wish to reset the export flags and clear finalized POs?" & vbCrLf & vbCrLf & _
                      "Selecting yes will prevent the items just exported from being re-exported" & vbCrLf & vbCrLf & _
                      "Note for Kirk: if you closed the PO window without exporting them, you DEFINITELY want to click NO here!", vbYesNo) = vbYes Then
                DoEvents
                'DB.Execute "exec spExportSaveReorderPoint"
                DB.Execute "exec spExportResetFlags"
                DB.Execute "exec spExportResetPLFlags"
                'DB.Execute "exec spExportClearPOItems"
                DB.Execute "exec spExportClearPOs"
                DoEvents
            End If
        End If
        
        StartExport = True
    
    Else
        StartExport = False
    End If
    
    If doPOs Then
        generateAndSendPOExportEmail
    End If
    
    If Not fromImport Then
        ClearProcessFlag "export"
    End If
    
    EventLogUpdate eventID, True
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : StartAutomaticExport
' DateTime  : 9/21/2005 11:23
' Author    : briandonorfio
' Purpose   : Starts an automatic export. The EXP6_UDF_SQL import script (the last one
'             in the chain) calls a providex logic file that deletes a blank text file
'             that acts as a mutex. When VB realizes the file is gone, it continues with
'             with the rest of the function.
'
'             PVX file is s:\mastest\mas200\version4\mas90\vi-delete-logic\poinv6_orderpt.pl
'
'             Contents for the PVX file:
'                ENTER
'                ERASE "s:\mastest\mas90-signs\export\export_delete_file.txt",ERR=*NEXT
'                EXIT
'
'             If anything is ever added after the UDF import script, the perform logic
'             should be moved into the new one. In VI, perform button, on failure and
'             on success should call ..\vi-delete-logic\poinv6_orderpt.pl
'---------------------------------------------------------------------------------------
'
Public Sub StartAutomaticExport(forceOrderPointCalc As Boolean) 'really no reason why this should be public...
    'autoexport runs the reorderpoint calculations on sunday mornings
    If forceOrderPointCalc Or DatePart("w", Date) = vbSunday Then
        CalculateReorderPoints
    End If

    RaiseEvent StepChanged("ExportPrep")
    DB.Execute "exec spExportPurgeAll"
    
    DB.Execute "exec spExportProductLine"
    DB.Execute "exec spExportInventory"
    'exportItemWhseLines
    DB.Execute "exec spExportInventoryVendor"
    'DB.Execute "exec spExportPriceLevel1"
    'DB.Execute "exec spExportPriceLevel2"
    'DB.Execute "exec spExportPriceLevelA"
    'DB.Execute "exec spExportPriceLevelB"
    'DB.Execute "exec spExportPriceLevelG"
    'DB.Execute "exec spExportPriceLevelH"
    DB.Execute "exec spExportPriceLevels"
    'DB.Execute "exec spExportReorderPoint"
    Call exportKits
    If DEBUG_MODE Then
        DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Export Prep Done' )"
    End If
    
    RaiseEvent StepChanged("Export")
    
    'turn mutex on, UDF import script deletes the file when it finishes (mutex off)
    'if another import gets chained, move the perform logic to that one.
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    fso.CreateTextFile MUTEX_COMPLETE
    
    If DEBUG_MODE Then
        DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Calling Mas200' )"
    End If
    
    ShellWait "s:\mastest\mas90-signs\mas200_import_items.bat", vbHide

    While fso.FileExists(MUTEX_COMPLETE)
        DoEvents
        RaiseEvent StatusChanged("Waiting for mutex...")
        Sleep 30000
        If DEBUG_MODE Then
            DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Waiting For Mas200' )"
        End If
    Wend
    Set fso = Nothing
    
    If DEBUG_MODE Then
        DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Export Done' )"
    End If

    'don't need to do this, import will take care of it.
    'DoEvents
    'dbExecute "exec spExportSaveReorderPoint", , , 60
    'dbExecute "exec spExportResetFlags", , , 60
    'dbExecute "exec spExportResetPLFlags", , , 60
    'dbExecute "exec spExportClearPOItems", , , 60
    'dbExecute "exec spExportClearPOs", , , 60
    'DoEvents
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RefreshQuantities
' DateTime  : 9/21/2005 10:20
' Author    : briandonorfio
' Purpose   : Just updates the quantities, so we can get the latest updates up to the
'             web site.
'---------------------------------------------------------------------------------------
'
Public Sub RefreshQuantities(Optional singleItem As String = "")
    importInvQtys singleItem, False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RefreshVendors
' DateTime  : 4/20/2006 09:29
' Author    : briandonorfio
' Purpose   : Just a public wrapper for importVendors.
'---------------------------------------------------------------------------------------
'
Public Sub RefreshVendors()
    importVendors
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ExportSinglePO
' DateTime  : 9/21/2005 10:20
' Author    : briandonorfio
' Purpose   : Exports a single PO, if it's set to finalized.
'---------------------------------------------------------------------------------------
'
'Public Sub ExportSinglePO(POID As String)
'    If Not exportPOs(POID) Then
'        MsgBox "Could not find PO, or PO is  already exported or not finalized?"
'    End If
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : CalculateReorderPoints
' DateTime  : 9/21/2005 09:58
' Author    : briandonorfio
' Purpose   : Recalculates the given items order points. Formula+calc table apparently
'             originally from triad, seems to work well enough. Items are given by a
'             range of line codes and vendors, this doesn't get used so I'm not really
'             sure exactly how it's supposed to work. Called with no arguments, it
'             defaults to all items. That's probably what you should do.
'---------------------------------------------------------------------------------------
'
'Public Function CalculateReorderPoints(Optional PLstart As String = "AAA", Optional PLend As String = "XXA", Optional Vstart As String = "0", Optional Vend As String = "zzzzzzz")
'    DB.Execute "exec spReorderPointConfigUpdate"
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("exec spReorderPointItems '" & PLstart & "', '" & PLend & "', '" & Vstart & "', '" & Vend & "'")
'    Dim rstConfig As ADODB.Recordset
'    Set rstConfig = DB.retrieve("SELECT * FROM ReorderPointConfig")
'    Dim curr As Double, Last As Double, ytd As Double, lyr As Double, total As Double, weeks As Long, NewOP As Long
'    While Not rst.EOF
'        RaiseEvent StatusChanged("Calculating Reorder Point for " & rst("ItemNumber"))
'        DoEvents
'        weeks = rstConfig("ShelfSupplyWks") + rstConfig("LeadTimeWks") + rstConfig("SafetySupplyWks")
'        curr = (rst("SalesThisPeriod") / rstConfig("WeeksCurPer")) * (rstConfig("WeightCurPer") / 100)
'        Last = (rst("SalesLastPeriod") / rstConfig("WeeksLastPer")) * (rstConfig("WeightLastPer") / 100)
'        ytd = (rst("SalesYTD") / rstConfig("WeeksYTD")) * (rstConfig("WeightYTD") / 100)
'        lyr = (rst("SalesLYR") / rstConfig("WeeksLastYear")) * (rstConfig("WeightLastYear") / 100)
'        total = curr + Last + ytd + lyr
'        NewOP = Int(total * weeks) + 1
'        If NewOP = rst("ReorderPoint") Then 'same as old op
'            DB.Execute "UPDATE InventoryMaster SET OldReorderPoint=ReorderPoint WHERE ItemNumber='" & rst("ItemNumber") & "'"
'        Else
'            DB.Execute "UPDATE InventoryMaster SET OldReorderPoint=ReorderPoint, ReorderPoint=" & NewOP & ", NewReorderPoint=" & NewOP & " WHERE ItemNumber='" & rst("ItemNumber") & "'"
'        End If
'        rst.MoveNext
'    Wend
'    RaiseEvent StatusChanged("")
'    rst.Close
'    Set rst = Nothing
'    rstConfig.Close
'    Set rstConfig = Nothing
'End Function
Public Function CalculateReorderPoints(Optional singleItem As String = "") As Boolean
    
    Dim whses As Variant, whseIter As Variant
    whses = Array("1", "2", "5")
    
    Dim rst As ADODB.Recordset
    Dim whereClause As String
    
    RaiseEvent StatusChanged("Caching Order Point Transactions")
    If singleItem = "" Then
        whereClause = ""
    Else
        whereClause = "WHERE ItemNumber='" & singleItem & "'"
    End If
    
    'currently we split the warehouses as p2 and so
    '  1 == p2
    '  2 == always 0
    '  5 == so
    'get some ranges for each, and weight each accordingly
    '    0 -  30 days == 48%
    '   31 - 120 days == 48%
    '  121 - 365 days == 1%
    '  366 - 730 days == 1%
    'then round off to an integer. anything 0.2 or higher rounds up
    Set rst = DB.retrieve("SELECT ItemNumber, WH1_1, WH1_2, WH1_3, WH1_4, " & _
                                             "WH5_1, WH5_2, WH5_3, WH5_4, " & _
                                             "WH2_1, WH2_2, WH2_3, WH2_4 " & _
                          "FROM vImportReorderPointsDataSource " & whereClause)
    Dim dataIndex As Dictionary
    Set dataIndex = New Dictionary
    While Not rst.EOF
        Dim subHash As Dictionary
        Set subHash = New Dictionary
        
        For Each whseIter In whses
            Dim temp As Double
            temp = CDbl(CLng(rst("WH" & whseIter & "_1")) * 0.48) + _
                   CDbl(CLng(rst("WH" & whseIter & "_2")) * 0.48) + _
                   CDbl(CLng(rst("WH" & whseIter & "_3")) * 0.01) + _
                   CDbl(CLng(rst("WH" & whseIter & "_4")) * 0.01)
            subHash.Add CLng(whseIter), CLng(Fix(temp + 0.8))
        Next whseIter
        
        dataIndex.Add CStr(rst("ItemNumber")), subHash
        rst.MoveNext
    Wend
    rst.Close
    
    If singleItem = "" Then
        whereClause = ""
    Else
        whereClause = "WHERE ItemNumber='" & singleItem & "'"
    End If
    Set rst = DB.retrieve("SELECT ItemNumber, ItemStatus, WarehouseID, ClearOrderPoint, StockHere, LockOrderPoint, OrderPoint FROM vImportReorderPointsToDoList " & whereClause)
    While Not rst.EOF
        RaiseEvent StatusChanged("Calculating Reorder Point for " & rst("ItemNumber"))
        DoEvents
        
        Dim item As String, whse As Long, shouldClearOP As Boolean, stockHere As Boolean, lockedOP As Boolean, currentOPNull As Boolean, currentOP As Long, sqlInsertNeeded As Boolean
        item = CStr(rst("ItemNumber"))
        whse = CLng(rst("WarehouseID"))
        shouldClearOP = CBool(rst("ClearOrderPoint") = 1) 'type is adInteger, wtf
        If IsNull(rst("StockHere")) Then 'no data means assume stocked for whse, not stocked for store
            Select Case whse
                Case Is = 1
                    stockHere = False
                Case Else
                    stockHere = True
            End Select
        Else
            stockHere = CBool(rst("StockHere"))
        End If
        If IsNull(rst("LockOrderPoint")) Then 'no data means assume not locked
            lockedOP = False
        Else
            lockedOP = CBool(rst("LockOrderPoint"))
        End If
        currentOPNull = CBool(IsNull(rst("OrderPoint")))
        If IsNull(rst("OrderPoint")) Then 'no data means zero OP (actually null, should check currentOPNull first)
            currentOP = 0
        Else
            currentOP = CLng(rst("OrderPoint"))
        End If
        sqlInsertNeeded = IsNull(rst("LockOrderPoint")) 'check non-nullable field for existence to see insert/update reqd
        
        'items at the store have a minimum order point of 1
        'so if this time around is for the store then force the OP up if necessary
        'totally ignores the warehouse's stock flag, everything gets an OP
        If whse = 1 Then
            If Not dataIndex.exists(item) Then
                dataIndex.Add item, New Dictionary
            End If
            If Not dataIndex.item(item).exists(whse) Then
                dataIndex.item(item).Add whse, 0
            End If
            If dataIndex.item(item).item(whse) < 1 Then
                If CLng(rst("ItemStatus")) = 0 Then
                    '2013-01-03: items with no high-level status (stock, dropship, d/c) don't
                    'get the defaulted 1, because i guess that's not important now.
                Else
                    dataIndex.item(item).item(whse) = 1
                End If
            End If
        End If
        
        Dim hasHistoricalData As Boolean, newOrderPoint As Long
        If dataIndex.exists(item) Then
            If dataIndex.item(item).exists(whse) Then
                hasHistoricalData = True
                newOrderPoint = dataIndex.item(item).item(whse)
            Else
                hasHistoricalData = False
                newOrderPoint = 0
            End If
        Else
            hasHistoricalData = False
            newOrderPoint = 0
        End If
        
        If shouldClearOP = True Then
            If currentOPNull = True And sqlInsertNeeded = True Then
                'null is correct, except if the row is missing entirely
                DB.Execute "INSERT INTO InventoryLocationInfo ( ItemNumber, WarehouseID, StockHere ) VALUES ( '" & item & "', " & whse & ", " & SQLBool(stockHere) & " )"
            ElseIf currentOPNull = False Then
                'current order point is not null, need to fix that
                DB.Execute "UPDATE InventoryLocationInfo SET OrderPoint=NULL, LockOrderPoint=0 WHERE ItemNumber='" & item & "' AND WarehouseID=" & whse
            End If
        ElseIf hasHistoricalData = False Then
            'no data to make an order point calculation
            If lockedOP = True Then
                'locked in current position, do nothing
            ElseIf currentOPNull = True Then
                'already null, do nothing
            ElseIf sqlInsertNeeded Then
                'add line with null OP
                DB.Execute "INSERT INTO InventoryLocationInfo ( ItemNumber, WarehouseID, StockHere, OrderPoint ) VALUES ( '" & item & "', " & whse & ", " & SQLBool(stockHere) & ", NULL )"
            Else
                'update to null OP
                DB.Execute "UPDATE InventoryLocationInfo SET OrderPoint=NULL WHERE ItemNumber='" & item & "' AND WarehouseID=" & whse
            End If
        Else
            'we have historical data and a new order point, check to see if it's different
            Dim doChange As Boolean
            doChange = False
            If lockedOP = True And newOrderPoint > currentOP Then
                'if the order point is locked, but it wants to be higher, then
                'it's ok to set it higher, ratchet effect
                doChange = True
            ElseIf lockedOP = False And newOrderPoint <> currentOP Then
                'if floating order point is different than current, set it
                doChange = True
            ElseIf sqlInsertNeeded = True Then
                'no line currently set, so try to set it?
                doChange = True
            End If
            
            If doChange = True Then
                If sqlInsertNeeded = True Then
                    DB.Execute "INSERT INTO InventoryLocationInfo ( ItemNumber, WarehouseID, StockHere, OrderPoint ) VALUES ( '" & item & "', " & whse & ", " & SQLBool(stockHere) & ", " & newOrderPoint & " )"
                Else
                    DB.Execute "UPDATE InventoryLocationInfo SET OrderPoint=" & newOrderPoint & " WHERE ItemNumber='" & item & "' AND WarehouseID=" & whse
                End If
            End If
        End If
        rst.MoveNext
    Wend
    rst.Close
    
    Set rst = Nothing
    CalculateReorderPoints = True
End Function

Public Function ImportShowSaleData() As Boolean
    
    'DB.Execute "TRUNCATE TABLE ShowSaleCalcSOHistHeader"
    'DB.Execute "TRUNCATE TABLE ShowSaleCalcSOHistDetail"
    'DB.Execute "TRUNCATE TABLE ShowSaleCalcARHeader"
    'DB.Execute "TRUNCATE TABLE ShowSaleCalcARDetail"
    'DB.Execute "TRUNCATE TABLE ShowSaleCalcSOHeader"
    'DB.Execute "TRUNCATE TABLE ShowSaleCalcSODetail"
    'DB.Execute "TRUNCATE TABLE ShowSaleCalcPOSPHeader"
    'DB.Execute "TRUNCATE TABLE ShowSaleCalcPOSPDetail"
    
    synchronizeMasTable "SO_SalesOrderHistoryHeader", _
                        "ShowSaleCalcSOHistHeader", _
                        Array(Array(True, True, False, "SalesOrderNo", "SalesOrderNo"), _
                              Array(True, False, False, "OrderDate", "OrderDate"), _
                              Array(True, False, False, "OrderStatus", "OrderStatus") _
                             )
    'synchronizeMasTable "SO_SalesOrderHistoryDetail", _
    '                    "ShowSaleCalcSOHistDetail", _
    '                    Array(Array(True, True, False, "SalesOrderNo", "SalesOrderNo"), _
    '                          Array(True, True, False, "SequenceNo", "SequenceNo"), _
    '                          Array(True, False, False, "LineKey", "LineKey"), _
    '                          Array(True, False, False, "CancelledLine", "CancelledLine"), _
    '                          Array(True, False, False, "ItemType", "ItemType"), _
    '                          Array(True, False, False, "ItemCode", "ItemCode"), _
    '                          Array(True, False, False, "QuantityShipped", "QuantityShipped"), _
    '                          Array(True, False, False, "LastExtensionAmt", "LastExtensionAmt") _
    '                        )
    
    synchronizeMasTable "AR_InvoiceHistoryHeader", _
                        "ShowSaleCalcARHeader", _
                        Array(Array(True, True, False, "InvoiceNo", "InvoiceNo"), _
                              Array(True, True, False, "HeaderSeqNo", "HeaderSeqNo"), _
                              Array(True, False, False, "SalesOrderNo", "SalesOrderNo"), _
                              Array(True, False, False, "OrderDate", "OrderDate"), _
                              Array(True, False, False, "InvoiceType", "InvoiceType") _
                             )
    synchronizeMasTable "AR_InvoiceHistoryDetail", _
                        "ShowSaleCalcARDetail", _
                        Array(Array(True, True, False, "InvoiceNo", "InvoiceNo"), _
                              Array(True, True, False, "HeaderSeqNo", "HeaderSeqNo"), _
                              Array(True, True, False, "DetailSeqNo", "DetailSeqNo"), _
                              Array(True, False, False, "ItemCode", "ItemCode"), _
                              Array(True, False, False, "ItemType", "ItemType"), _
                              Array(True, False, False, "ProductLine", "ProductLine"), _
                              Array(False, False, False, "QuantityShipped", "QuantityShipped"), _
                              Array(False, False, False, "UnitPrice", "UnitPrice") _
                             ), _
                        "ItemType<>'4'"
    
    synchronizeMasTable "SO_SalesOrderHeader", _
                        "ShowSaleCalcSOHeader", _
                        Array(Array(True, True, False, "SalesOrderNo", "SalesOrderNo"), _
                              Array(True, False, False, "OrderDate", "OrderDate"), _
                              Array(True, False, False, "OtherPaymentTypeRefNo", "OtherPaymentTypeRefNo") _
                             )
    synchronizeMasTable "SO_SalesOrderDetail", _
                        "ShowSaleCalcSODetail", _
                        Array(Array(True, True, False, "SalesOrderNo", "SalesOrderNo"), _
                              Array(True, True, False, "LineKey", "LineKey"), _
                              Array(True, True, False, "LineSeqNo", "LineSeqNo"), _
                              Array(True, False, False, "ItemType", "ItemType"), _
                              Array(True, False, False, "ItemCode", "ItemCode"), _
                              Array(False, False, False, "QuantityOrdered", "QuantityOrdered"), _
                              Array(False, False, False, "ExtensionAmt", "ExtensionAmt") _
                             ), _
                        "ItemType<>'4'"
    
    synchronizeMasTable "P2_TransactionHistoryHeader", _
                        "ShowSaleCalcPOSPHeader", _
                        Array(Array(True, True, False, "TransactionNo", "TransactionNo"), _
                              Array(True, False, False, "TransactionDate", "TransactionDate"), _
                              Array(True, False, False, "CashierID", "CashierID"), _
                              Array(True, False, False, "RegisterNo", "RegisterNo"), _
                              Array(True, False, False, "TransactionStatus", "TransactionStatus"), _
                              Array(True, False, False, "ReferenceNo", "ReferenceNo") _
                             )
    synchronizeMasTable "P2_TransactionHistoryDetail", _
                        "ShowSaleCalcPOSPDetail", _
                        Array(Array(True, True, False, "TransactionNo", "TransactionNo"), _
                              Array(True, True, False, "LineKey", "LineKey"), _
                              Array(True, False, False, "ItemType", "ItemType"), _
                              Array(True, False, False, "ItemCode", "ItemCode"), _
                              Array(True, False, False, "ProductLine", "ProductLine"), _
                              Array(False, False, False, "QuantityOrdered", "QuantityOrdered"), _
                              Array(False, False, False, "ExtensionAmt", "ExtensionAmt") _
                             )
    
End Function



'---------------------------------------------------------------------------------------
' Procedure : getQtyToExportStr
' DateTime  : 9/7/2005 09:32
' Author    : briandonorfio
' Purpose   : Returns a string with the count from each export table, suitable for a y/n
'             message box.
'---------------------------------------------------------------------------------------
'
Private Function getQtyToExportStr() As String
    Dim str As String, i As Long, tablesA() As Variant
    tablesA = Array("ExportInventory", "ExportInventoryVendor", "ExportNonDefaultWhses", "ExportPriceLevelMAS45", "ExportProductLine")
    Dim rst As ADODB.Recordset
    For i = 0 To UBound(tablesA)
        Set rst = DB.retrieve("SELECT COUNT(*) AS NumToExp FROM " & tablesA(i))
        str = str & vbCrLf & Mid(tablesA(i), 7) & ": " & rst("NumToExp")
        rst.Close
    Next i
    Set rst = DB.retrieve("SELECT COUNT(*) AS NumLines, COUNT(DISTINCT PONumber) AS NumPOs FROM ExportPurchaseOrders")
    str = str & vbCrLf & "PurchaseOrders: " & rst("NumPOs") & " POs, " & rst("NumLines") & " lines"
    rst.Close
    Set rst = Nothing
    getQtyToExportStr = "Do you want to run the following export?" & vbCrLf & str
End Function

''---------------------------------------------------------------------------------------
'' Procedure : exportItemWhseLines
'' DateTime  : 3/6/2006 13:54
'' Author    : briandonorfio
'' Purpose   : Goes through the to-be-exported purchase orders to find any items going to
''             a warehouse other than 000. Then queries mas200 for each item/whse on that
''             list, and adds it to an export table if it's missing.
''
''             This is needed for a workaround with Mas200, where items shipping to a
''             warehouse that didn't exist in IM2 would cause an error rebuilding the PO
''             sort files. It's probably just some perform logic that needs to be added.
''---------------------------------------------------------------------------------------
''
'Private Function exportItemWhseLines() As Boolean
'    Dim retval As Boolean
'    retval = False
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT DISTINCT ItemNumber, WhseCode FROM ExportPurchaseOrders WHERE ItemNumber<>'COMMENT' AND WhseCode<>'000'")
'    Dim masrst As ADODB.Recordset
'    While Not rst.EOF
'        Set masrst = MASDB.retrieve("SELECT COUNT(*) FROM IM_ItemWarehouse WHERE ItemCode='" & rst("ItemNumber") & "' AND WarehouseCode='" & rst("WhseCode") & "'")
'        If masrst(0) = 0 Then
'            DB.Execute "INSERT INTO ExportNonDefaultWhses ( ItemNumber, WhseCode ) VALUES ( '" & rst("ItemNumber") & "', '" & rst("WhseCode") & "' )"
'            retval = True
'        End If
'        DoEvents
'        masrst.Close
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'    exportItemWhseLines = retval
'End Function

Private Function exportKits() As Boolean
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT SalesKitNo, ComponentItemCode, QuantityPerAssembly " & _
                          "FROM IM_SalesKitDetail " & _
                          "WHERE SalesKitNo IN (SELECT ItemNumber FROM InventoryMaster WHERE NewItem=1 AND IsMASKit=1) " & _
                          "ORDER BY SalesKitNo, ComponentItemCode")
    
    Open KIT_EXPORT_FILE For Output As #1
    Print #1, "LineType,SalesKitNo,ComponentItemCode,QuantityPerAssembly"
    
    Dim lastKit As String
    lastKit = ""
    While Not rst.EOF
        If lastKit <> rst("SalesKitNo") Then
            Print #1, "H," & rst("SalesKitNo")
            lastKit = rst("SalesKitNo")
        End If
        
        Print #1, "L,," & rst("ComponentItemCode") & "," & rst("QuantityPerAssembly")
        
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    Close #1
    
    exportKits = True
End Function

'---------------------------------------------------------------------------------------
' Procedure : exportPOs
' DateTime  : 7/12/2005 15:31
' Author    : briandonorfio
' Purpose   : Adds any finalized purchase orders to the ExportPurchaseOrders table.
'             Comments and items are all individual line-items in Mas200. If supplied
'             with a POID, then it will export just that PO, still has error checking so
'             we can't specify one that's already been exported, or hasn't yet been
'             finalized. Returns a boolean, true if there were any purchase orders
'             exported, false otherwise.
'---------------------------------------------------------------------------------------
'
Private Function exportPOs(Optional POID As String = "") As Boolean
    Dim rst As ADODB.Recordset
    If POID = "" Then
        Set rst = DB.retrieve("SELECT ID, PONumber, POAmount, POTerms, VendorNumber, DueDate, FreightFree, POComment, ShipToCode, ShipToName, ShipToAddress1, ShipToAddress2, ShipToCity, ShipToState, ShipToZip, MiscNotes, SalesOrderNo, ShipToTelephoneNo FROM PurchaseOrders WHERE Finalized=1 AND Exported=0")
    Else
        Set rst = DB.retrieve("SELECT ID, PONumber, POAmount, POTerms, VendorNumber, DueDate, FreightFree, POComment, ShipToCode, ShipToName, ShipToAddress1, ShipToAddress2, ShipToCity, ShipToState, ShipToZip, MiscNotes, SalesOrderNo, ShipToTelephoneNo FROM PurchaseOrders WHERE Finalized=1 AND Exported=0 AND ID=" & POID)
    End If

    If rst.EOF Then
        exportPOs = False
        Exit Function
    End If

'    While Not rst.EOF
'        If rst("PONumber") = "" Then    'don't know how it'd get here without
'            CreatePONumber rst("ID")    'a po #, but check it anyway
'        End If
'        rst.MoveNext
'    Wend
'    rst.MoveFirst
'
'    Dim Comment As String
'    Dim commA As Variant, word As Variant
'    While Not rst.EOF
'        If rst("FreightFree") Then
'            Comment = "Freight Free Order"
'        Else
'            Comment = "Tools Plus pays freight"
'        End If
'        DB.Execute "exec spExportPOTempComment '" & rst("ID") & "', '" & EscapeSQuotes(Comment) & "', ''"
'
'        'if this changes, also change the query in the purchase_orders.mdb
'        If rst("ShipToCode") = "0002" Then
'            DB.Execute "exec spExportPOTempComment '" & rst("ID") & "', 'Deliveries require an appointment set up in', 'advance. Unscheduled deliveries may be refused,'"
'            DB.Execute "exec spExportPOTempComment '" & rst("ID") & "', 'Tools Plus is not responsible for any added costs.', ''"
'            'DB.Execute "exec spExportPOTempComment '" & rst("ID") & "', '" & SHIPPING_GUY_PHONE & " or 203-573-0750 ext 831.', ''"
'            If SHIPPING_GUY_PHONE_1 = "" And SHIPPING_GUY_PHONE_2 = "" Then
'                'do nothing, this is probably bad
'            ElseIf SHIPPING_GUY_PHONE_2 = "" Then
'                DB.Execute "exec spExportPOTempComment '" & rst("ID") & "', 'Contact Johnny to schedule delivery.', '" & SHIPPING_GUY_PHONE_1 & "'"
'            ElseIf SHIPPING_GUY_PHONE_1 = "" Then
'                DB.Execute "exec spExportPOTempComment '" & rst("ID") & "', 'Contact Johnny to schedule delivery.', '" & SHIPPING_GUY_PHONE_2 & "'"
'            Else
'                DB.Execute "exec spExportPOTempComment '" & rst("ID") & "', 'Contact Johnny to schedule delivery.', '" & SHIPPING_GUY_PHONE_1 & " or " & SHIPPING_GUY_PHONE_2 & "'"
'            End If
'        End If
'
'        'splits the po comment field into multiple mas90 size comments
'        formatAndAddCommentLines rst("ID"), Nz(rst("POComment"))
'        formatAndAddCommentLines rst("ID"), Nz(rst("MiscNotes"))
'
'        rst.MoveNext
'    Wend
'    DB.Execute "exec spExportPOFinalComment"
'    rst.Close
'
'    Set rst = DB.retrieve("SELECT * FROM vExportPOGetItems ORDER BY ItemNumber")
'    While Not rst.EOF
'        DB.Execute "exec spExportPOAppendItem '" & rst("ItemNumber") & "'"
'        DB.Execute "exec spExportPOItemComment '" & rst("ItemNumber") & "'"
'        rst.MoveNext
'    Wend

    Dim hasAPIHandleableDropshipPOs As Boolean
    hasAPIHandleableDropshipPOs = False
    
    Open PO_EXPORT_FILE For Output As #1
    'TODO: print header line "LineType,..."
    
    While Not rst.EOF
        If rst("PONumber") = "" Then    'don't know how it'd get here without
            CreatePONumber rst("ID")    'a po #, but check it anyway
        End If
        
        If rst("ShipToCode") = "1234" Then
            If rst("VendorNumber") = "JET" Or rst("VendorNumber") = "BLAC01" Then
                hasAPIHandleableDropshipPOs = True
            End If
        End If
        
        poInsertHeader rst
        
        poInsertComment rst, IIf(rst("FreightFree"), "Freight Free Order", "Tools Plus pays freight"), ""
        
        ' add comments on regular POs to scott rd
        If rst("ShipToCode") = "0003" Then
            poInsertComment rst, "Truck freight require an appointment set up in", ""
            poInsertComment rst, "advance. Unscheduled deliveries may be refused.", ""
            poInsertComment rst, "Tools Plus is not responsible for any added costs.", ""
            If SHIPPING_GUY_PHONE_1 = "" And SHIPPING_GUY_PHONE_2 = "" Then
                'do nothing, this is probably bad
            ElseIf SHIPPING_GUY_PHONE_2 = "" Then
                poInsertComment rst, "Contact Johnny to schedule delivery.", ""
                poInsertComment rst, SHIPPING_GUY_PHONE_1, ""
            ElseIf SHIPPING_GUY_PHONE_1 = "" Then
                poInsertComment rst, "Contact Johnny to schedule delivery.", ""
                poInsertComment rst, SHIPPING_GUY_PHONE_2, ""
            Else
                poInsertComment rst, "Contact Johnny to schedule delivery.", ""
                poInsertComment rst, SHIPPING_GUY_PHONE_1 & " or " & SHIPPING_GUY_PHONE_2, ""
            End If
        End If
        
        ' add comments to dropship POs
        If rst("ShipToCode") = "1234" Then
            poInsertComment rst, "Unless otherwise noted, Tools Plus is not", ""
            poInsertComment rst, "responsible for state Sales Taxes on this order.", ""
            
            If rst("VendorNumber") = "DELT03" Then
                poInsertComment rst, "Truck line must call 24 hours before delivery", ""
            End If
        End If
        
        If Nz(rst("POComment")) <> "" Then
            formatAndAddCommentLines rst, rst("POComment")
        End If
        If rst("ShipToCode") = "1234" Then
            formatAndAddCommentLines rst, "CUSTOMER PHONE NUMBER " & Nz(rst("ShipToTelephoneNo")) & vbCrLf & "ORDER NUMBER " & Nz(rst("SalesOrderNo"))
        End If
        If Nz(rst("MiscNotes")) <> "" Then
            formatAndAddCommentLines rst, Nz(rst("MiscNotes"))
        End If
        
        Dim rst2 As ADODB.Recordset
        Set rst2 = DB.retrieve("SELECT InventoryMaster.ItemNumber, InventoryMaster.StdCost, PurchaseOrderLines.Quantity, InventoryMaster.POComment, InventoryMaster.PrimaryVendor FROM PurchaseOrderLines INNER JOIN InventoryMaster ON PurchaseOrderLines.ItemNumber=InventoryMaster.ItemNumber WHERE PurchaseOrderLines.HeaderID=" & rst("ID"))
        While Not rst2.EOF
            poInsertItem rst, rst2
            
            'If Nz(rst2("POComment")) <> "" Then
            '    poInsertComment rst, rst2("POComment"), ""
            'End If
            
            rst2.MoveNext
        Wend
        rst2.Close
        Set rst2 = Nothing
        
        rst.MoveNext
    Wend
    
    rst.Close
    Set rst = Nothing
    
    Close #1
    
    exportPOs = True
    
    If hasAPIHandleableDropshipPOs Then
        On Error Resume Next
        Shell DROPSHIP_EXE, vbNormalFocus
        On Error GoTo 0
    End If
    
End Function

Private Function poInsertHeader(rst As ADODB.Recordset) As Boolean
    Dim fields(16) As String, i As Long
    fields(0) = "H"
    fields(1) = rst("PONumber")
    fields(2) = Format(Now(), "YYYYMMDD")
    fields(3) = Format(rst("DueDate"), "YYYYMMDD")
    fields(4) = rst("VendorNumber")
    fields(5) = rst("ShipToCode")
    fields(6) = rst("ShipToName")
    fields(7) = rst("ShipToAddress1")
    fields(8) = Nz(rst("ShipToAddress2"))
    fields(9) = rst("ShipToCity")
    fields(10) = rst("ShipToState")
    fields(11) = rst("ShipToZip")
    fields(12) = rst("POTerms")
    For i = 13 To UBound(fields)
        fields(i) = ""
    Next i
    For i = LBound(fields) To UBound(fields)
        fields(i) = FormatStrForCSV(fields(i), True)
    Next i
    
    Print #1, Join(fields, ",")
    
    poInsertHeader = True
End Function

Private Function poInsertComment(rst As ADODB.Recordset, c1 As String, c2 As String) As Boolean
    DB.Execute "INSERT INTO ExportPurchaseOrders ( PONumber, Comment, Comment2, Vendor, " & _
               "                                   Terms, ReqDate, ShipToCode, WhseCode ) " & _
               "VALUES ( '" & rst("PONumber") & "', '" & EscapeSQuotes(c1) & "', '" & EscapeSQuotes(c2) & "', '" & rst("VendorNumber") & "', " & _
               "         '" & rst("POTerms") & "', '" & Format(rst("DueDate"), "mm/dd/yyyy") & "', '" & rst("ShipToCode") & "', '000' )"
    
    Dim fields(16) As String, i As Long
    fields(0) = "L"
    For i = 1 To 12
        fields(i) = ""
    Next i
    fields(13) = "/C"
    fields(14) = ""
    fields(15) = ""
    fields(16) = c1
    For i = LBound(fields) To UBound(fields)
        fields(i) = FormatStrForCSV(fields(i), True)
    Next i
    
    Print #1, Join(fields, ",")
    
    poInsertComment = True
End Function

Private Function poInsertItem(rstHeader As ADODB.Recordset, rstItem As ADODB.Recordset) As Boolean
    DB.Execute "INSERT INTO ExportPurchaseOrders ( DateOfPO, Vendor, ItemNumber, QtyOrdered, UnitCost, ItemTotal, Terms, ReqDate, PONumber, OrderTotal, " & _
               "                                   ShipToCode, ShipToName, ShipToAddress1, ShipToAddress2, ShipToCity, ShipToState, ShipToZip, WhseCode ) " & _
               "VALUES ( CONVERT(VARCHAR(10), GETDATE(), 101), '" & rstItem("PrimaryVendor") & "', '" & rstItem("ItemNumber") & "', " & rstItem("Quantity") & ", " & rstItem("StdCost") & ", " & rstItem("Quantity") * rstItem("StdCost") & ", '" & rstHeader("POTerms") & "', '" & Format(rstHeader("DueDate"), "mm/dd/yyyy") & "', '" & rstHeader("PONumber") & "', " & rstHeader("POAmount") & ", " & _
               "         '" & rstHeader("ShipToCode") & "', '" & EscapeSQuotes(rstHeader("ShipToName")) & "', '" & EscapeSQuotes(rstHeader("ShipToAddress1")) & "', '" & EscapeSQuotes(rstHeader("ShipToAddress2")) & "', '" & EscapeSQuotes(rstHeader("ShipToCity")) & "', '" & rstHeader("ShipToState") & "', '" & rstHeader("ShipToZip") & "', '000' )"
    
    Dim fields(16) As String, i As Long
    fields(0) = "L"
    For i = 1 To 12
        fields(i) = ""
    Next i
    fields(13) = rstItem("ItemNumber")
    fields(14) = rstItem("Quantity")
    fields(15) = rstItem("StdCost")
    fields(16) = Nz(rstItem("POComment"))
    
    Print #1, Join(fields, ",")
    
    poInsertItem = True
End Function

Private Function generateAndSendPOExportEmail() As Boolean
    Dim poHash As Dictionary
    Set poHash = New Dictionary
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ExportPurchaseOrders.PONumber, PurchaseOrders.SalesOrderNo, ExportPurchaseOrders.Vendor, ExportPurchaseOrders.QtyOrdered, ExportPurchaseOrders.ItemNumber, ExportPurchaseOrders.UnitCost FROM ExportPurchaseOrders INNER JOIN PurchaseOrders ON ExportPurchaseOrders.PONumber=PurchaseOrders.PONumber WHERE ExportPurchaseOrders.ShipToCode='1234' AND ExportPurchaseOrders.ItemNumber<>'COMMENT' ORDER BY ExportPurchaseOrders.ID")
    While Not rst.EOF
        If Not poHash.exists(CStr(rst("PONumber"))) Then
            poHash.Add CStr(rst("PONumber")), "Purchase Order: " & rst("PONumber") & vbCrLf & _
                                              "Sales Order: " & Nz(rst("SalesOrderNo"), "<unknown>") & vbCrLf & _
                                              "Vendor: " & rst("Vendor") & vbCrLf & _
                                              vbCrLf & _
                                              "(" & rst("QtyOrdered") & ") " & rst("ItemNumber") & " @ " & Format(rst("UnitCost"), "Currency") & vbCrLf
        Else
            poHash.item(CStr(rst("PONumber"))) = poHash.item(CStr(rst("PONumber"))) & "(" & rst("QtyOrdered") & ") " & rst("ItemNumber") & " @ " & Format(rst("UnitCost"), "Currency") & vbCrLf
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    If poHash.count > 0 Then
        Dim emailBody As String
        emailBody = ""
        Dim iter As Variant
        For Each iter In poHash.keys
            emailBody = emailBody & poHash.item(CStr(iter)) & vbCrLf
        Next iter
        
        SendEmailTo "kathleen@tools-plus.com", "New dropship purchase orders entered into Mas", emailBody
        generateAndSendPOExportEmail = True
    Else
        generateAndSendPOExportEmail = False
    End If
    Set poHash = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : formatAndAddCommentLines
' DateTime  : 11/5/2007 14:41
' Author    : briandonorfio
' Purpose   : Splits comments into lines of 50 characters or less, adding them to the
'             database. Theoretically handles newlines as well, but probably won't
'             stack multiple newlines in a row for fancy formatting.
'---------------------------------------------------------------------------------------
'
'Private Sub formatAndAddCommentLines(POID As String, fullComment As String)
Private Sub formatAndAddCommentLines(rst As ADODB.Recordset, fullComment As String)
    Dim Comment(0 To 1) As String
    Comment(0) = ""
    Comment(1) = ""
    Dim X As Long
    X = 0
    Dim linesArray As Variant
    linesArray = Split(fullComment, vbCrLf)
    Dim thisline As Variant
    For Each thisline In linesArray
        Dim wordsArray As Variant
        wordsArray = Split(thisline, " ")
        Dim thisWord As Variant
        For Each thisWord In wordsArray
            If Len(Comment(X) & IIf(Comment(X) = "", "", " ") & thisWord) <= 50 Then
                Comment(X) = Comment(X) & IIf(Comment(X) = "", "", " ") & thisWord
            Else
                If X = 0 Then
                    X = 1
                Else
                    'DB.Execute "exec spExportPOTempComment '" & POID & "', '" & EscapeSQuotes(Comment(0)) & "', '" & EscapeSQuotes(Comment(1)) & "'"
                    poInsertComment rst, Comment(0), Comment(1)
                    Comment(0) = ""
                    Comment(1) = ""
                    X = 0
                End If
                Comment(X) = thisWord
            End If
        Next thisWord
        If Comment(0) <> "" Or Comment(1) <> "" Then
            'DB.Execute "exec spExportPOTempComment '" & POID & "', '" & EscapeSQuotes(Comment(0)) & "', '" & EscapeSQuotes(Comment(1)) & "'"
            If Comment(0) <> "" Then
                poInsertComment rst, Comment(0), ""
            End If
            If Comment(1) <> "" Then
                poInsertComment rst, Comment(1), ""
            End If
            Comment(0) = ""
            Comment(1) = ""
            X = 0
        End If
    Next thisline
    If Comment(0) <> "" Or Comment(1) <> "" Then
        'DB.Execute "exec spExportPOTempComment '" & POID & "', '" & EscapeSQuotes(Comment(0)) & "', '" & EscapeSQuotes(Comment(1)) & "'"
        If Comment(0) <> "" Then
            poInsertComment rst, Comment(0), ""
        End If
        If Comment(1) <> "" Then
            poInsertComment rst, Comment(1), ""
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : importVendorAddresses
' DateTime  : 7/12/2005 14:26
' Author    : briandonorfio
' Purpose   : Imports the vendor purchase addresses (POB_VendorPurchaseAddress) and ship
'             to address codes (POA), both needed by the PurchaseOrders form.
'---------------------------------------------------------------------------------------
'
Private Sub importVendorAddresses()
    RaiseEvent StatusChanged("Querying MAS200 For Vendor Addresses")
    synchronizeMasTable "PO_VendorPurchaseAddress", _
                        "PO_VendorPurchaseAddress", _
                        Array(Array(True, True, False, "APDivisionNo", "APDivisionNo"), _
                              Array(True, True, False, "VendorNo", "VendorNo"), _
                              Array(True, True, False, "PurchaseAddressCode", "PurchaseAddressCode"), _
                              Array(True, False, False, "PurchaseName", "PurchaseName"), _
                              Array(True, False, False, "PurchaseAddress1", "PurchaseAddress1"), _
                              Array(True, False, False, "PurchaseAddress2", "PurchaseAddress2"), _
                              Array(True, False, False, "PurchaseAddress3", "PurchaseAddress3"), _
                              Array(True, False, False, "PurchaseCity", "PurchaseCity"), _
                              Array(True, False, False, "PurchaseState", "PurchaseState"), _
                              Array(True, False, False, "PurchaseZipCode", "PurchaseZipCode"), _
                              Array(True, False, False, "PurchaseCountryCode", "PurchaseCountryCode"), _
                              Array(True, False, False, "BatchFax", "BatchFax"), _
                              Array(True, False, False, "TelephoneNo", "TelephoneNo"), _
                              Array(True, False, False, "TelephoneExt", "TelephoneExt"), _
                              Array(True, False, False, "FaxNo", "FaxNo"), _
                              Array(True, False, False, "ContactCode", "ContactCode"), _
                              Array(True, False, False, "TaxSchedule", "TaxSchedule"), _
                              Array(True, False, False, "ExemptionNoOnFile", "ExemptionNoOnFile"))
    RaiseEvent StatusChanged("Querying MAS200 For Shipping Addresses")
    synchronizeMasTable "PO_ShipToAddress", _
                        "PO_ShipToAddress", _
                        Array(Array(True, True, False, "ShipToCode", "ShipToCode"), _
                              Array(True, False, False, "ShipToCodeDesc", "ShipToCodeDesc"), _
                              Array(True, False, False, "ShipToName", "ShipToName"), _
                              Array(True, False, False, "ShipToAddress1", "ShipToAddress1"), _
                              Array(True, False, False, "ShipToAddress2", "ShipToAddress2"), _
                              Array(True, False, False, "ShipToAddress3", "ShipToAddress3"), _
                              Array(True, False, False, "ShipToCity", "ShipToCity"), _
                              Array(True, False, False, "ShipToState", "ShipToState"), _
                              Array(True, False, False, "ShipToZipCode", "ShipToZipCode"), _
                              Array(True, False, False, "ShipToCountryCode", "ShipToCountryCode"))
End Sub

'---------------------------------------------------------------------------------------
' Procedure : importProductLines
' DateTime  : 7/12/2005 14:27
' Author    : briandonorfio
' Purpose   : Imports product line info from IMA_ProductLine. In-place update, currently
'             only checks for new product lines, doesn't update old product lines with
'             changed information.
'---------------------------------------------------------------------------------------
'
Private Sub importProductLines()
    RaiseEvent StatusChanged("Querying MAS200 For Product Lines")
    Dim rst As ADODB.Recordset
    Set rst = MASDB.retrieve("SELECT ProductLine FROM IM_ProductLine")
    Dim rstPL As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    While Not rst.EOF
        RaiseEvent StatusChanged("Processing " & rst("ProductLine"))
        Set rstPL = DB.retrieve("SELECT ProductLine FROM ProductLine WHERE ProductLine='" & rst("ProductLine") & "'")
        If rstPL.EOF Then
            Set rst2 = DB.retrieve("spImportLookupProdLineVendor '" & rst("ProductLine") & "'")
            If Not rst2.EOF Then
                DB.Execute "exec spImportProdLineNew '" & Trim(rst("ProductLine")) & "', '" & Trim(rst2("PrimaryVendor")) & "'"
            Else
                DB.Execute "exec spImportProdLineNew '" & Trim(rst("ProductLine")) & "', ''"
            End If
            rst2.Close
        Else
            'product lines can be updated here, if we want to
        End If
        rstPL.Close
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : importInventory
' DateTime  : 7/12/2005 14:29
' Author    : briandonorfio
' Purpose   : Imports basic inventory information (IM1). In-place updating, handles both
'             inserting new items and updating existing. Not sure if there's a way of
'             renaming an item from Mas200, it wouldn't handle that, since it'd leave
'             the old item there, which would then be duplicated after an export.
'
'             MUST DISABLE TRIGGERS, or everything will be set to offload.
'---------------------------------------------------------------------------------------
'
Private Sub importInventory()
    RaiseEvent StatusChanged("Querying MAS200 For Inventory Items")
    synchronizeMasTable "CI_Item", _
                        "InventoryMaster", _
                        Array(Array(True, True, False, "ItemCode", "ItemNumber"), _
                              Array(True, False, False, "ItemCodeDesc", "ItemDescription"), _
                              Array(True, False, False, "ProductLine", "ProductLine"), _
                              Array(True, False, False, "PrimaryVendorNo", "PrimaryVendor"), _
                              Array(False, False, False, "StandardUnitCost", "StdCost"), _
                              Array(False, False, False, "AverageUnitCost", "AvgCost"), _
                              Array(False, False, False, "StandardUnitPrice", "StdPrice"), _
                              Array(False, False, False, "SuggestedRetailPrice", "ListPrice"), _
                              Array(True, False, False, "LastReceiptDate", "LastReceiptDate") _
                             ), _
                        "ItemType='1'", _
                        True, _
                        False
    
    RaiseEvent StatusChanged("Querying MAS200 For Inventory Items")
    synchronizeMasTable "CI_Item", _
                        "PartNumbers", _
                        Array(Array(True, True, False, "ItemCode", "ItemNumber")), _
                        "ItemType='1' AND ItemCode<'XXX%'", _
                        True, _
                        False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : importInvQtys
' DateTime  : 7/12/2005 14:31
' Author    : briandonorfio
' Purpose   : Imports quantity information (IM2). Saved to table InventoryQuantities.
'             Creates new item/whse if needed, otherwise in-place updating. Optional
'             argument doReorderPoints, defaults to true. If false, then it skips the
'             reorder points and all items past XXX*. If false, it also doesn't need the
'             trigger disabled. Good for a mid-day refresh of the qty on so/hand.
'
'             MUST DISABLE TRIGGERS if importing the reorder points (default), or
'             everything updated will be set to offload.
'---------------------------------------------------------------------------------------
'
Private Sub importInvQtys(Optional singleItem As String = "", Optional doReorderPoints As Boolean = True)
    RaiseEvent StatusChanged("Querying MAS200 For Quantities")
    Dim whereClause As String
    If singleItem = "" And doReorderPoints = True Then
        whereClause = ""
    ElseIf singleItem = "" And doReorderPoints = False Then
        whereClause = "ItemCode<'XXX%'"
    Else
        whereClause = "ItemCode='" & singleItem & "'"
    End If
    synchronizeMasTable "IM_ItemWarehouse", _
                        "InventoryQuantities", _
                        Array(Array(True, True, False, "ItemCode", "ItemNumber"), _
                              Array(True, True, False, "WarehouseCode", "WhseCode"), _
                              Array(False, False, False, "QuantityOnHand", "QuantityOnHand"), _
                              Array(False, False, False, "QuantityOnBackOrder", "QuantityOnBackOrder"), _
                              Array(False, False, False, "QuantityOnPurchaseOrder", "QuantityOnPurchaseOrder"), _
                              Array(False, False, False, "QuantityOnSalesOrder", "QuantityOnSalesOrder")), _
                        whereClause, _
                        True, _
                        False   'this has to be false, or all XXX/ZZZ get killed on a refresh qtys
                                'the synchronizeMasTable function should be checking for this now,
                                'in case i fuck up again
    
    'this sucks, but there's not really a better way of handling it?
    'there might be a minute or two of null-ness for the itemid column,
    'but that's only the first time the row exists (and 000 rows are
    'pre-made, so good there). basically, just when an item hits the
    '900 for the first time will there be an oddity here.
    '
    'for new poinv: have this as a callback in the row creation for
    'whatever the new version of synchronizeMasTable ends up being.
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT DISTINCT ItemNumber FROM InventoryQuantities WHERE ItemID IS NULL")
    While Not rst.EOF
        DB.Execute "UPDATE InventoryQuantities SET ItemID=" & Utilities.GetItemIDByItemCode(rst("ItemNumber")) & " WHERE ItemNumber='" & rst("ItemNumber") & "' AND ItemID IS NULL"
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
'2010-02-18 - no longer sync'ing reorder points back
'    'this will definitely need to change when we moved to multiple
'    'warehouses, since this is pretty useless
'    If doReorderPoints Then
'        RaiseEvent StatusChanged("Querying MAS200 For Order Points")
'        Dim rst As ADODB.Recordset
'        Set rst = MASDB.retrieve("SELECT ItemNumber, WhseCode, ReorderPointQty FROM IM2_InventoryItemWhseDetl WHERE WhseCode='000'")
'        While Not rst.EOF
'            RaiseEvent StatusChanged("Processing " & rst("ItemNumber"))
'            DoEvents
'            Dim rst2 As ADODB.Recordset
'            Set rst2 = DB.retrieve("SELECT ReorderPoint FROM InventoryMaster WHERE ItemNumber='" & rst("ItemNumber") & "'")
'            If rst2.EOF Then
'                'extra line in the im2 table that's not an actual part number?
'                'just ignore it, i guess
'            Else
'                If rst2("ReorderPoint") <> rst("ReorderPointQty") Then
'                    DB.Execute "UPDATE InventoryMaster SET ReorderPoint=" & Trim$(EscapeSQuotes(rst("ReorderPointQty"))) & ", OldReorderPoint=" & Trim$(EscapeSQuotes(rst("ReorderPointQty"))) & ", NewReorderPoint=-1 WHERE ItemNumber='" & Trim$(EscapeSQuotes(rst("ItemNumber"))) & "'"
'                End If
'            End If
'            rst2.Close
'            Set rst2 = Nothing
'            rst.MoveNext
'        Wend
'        rst.Close
'        Set rst = Nothing
'    End If
    
    RaiseEvent StatusChanged("Refreshing In-Stock / BackOrder")
'    CheckInStockOrBackOrder
CheckInStockOrBackOrder2 singleItem
End Sub

'---------------------------------------------------------------------------------------
' Procedure : importInvPrices
' DateTime  : 7/12/2005 14:32
' Author    : briandonorfio
' Purpose   : Imports both store and web pricing information (IMB_PriceCode). In-place
'             updating, does not create new items or delete missing items.
'
'             MUST DISABLE TRIGGERS, or things that get updated will be set to offload.
'---------------------------------------------------------------------------------------
'
Private Sub importInvPrices()
    RaiseEvent StatusChanged("Querying MAS200 For Store And Web Prices")
    synchronizeMasTable "IM_PriceCode", _
                        "InventoryMaster", _
                        Array(Array(True, True, False, "ItemCode", "ItemNumber"), _
                              Array(False, False, False, "BreakQuantity1", "BreakQty1"), _
                              Array(False, False, False, "BreakQuantity2", "BreakQty2"), _
                              Array(False, False, False, "BreakQuantity3", "BreakQty3"), _
                              Array(False, False, False, "BreakQuantity4", "BreakQty4"), _
                              Array(False, False, False, "BreakQuantity5", "BreakQty5"), _
                              Array(False, False, False, "DiscountMarkup1", "DiscountMarkupPriceRate1"), _
                              Array(False, False, False, "DiscountMarkup2", "DiscountMarkupPriceRate2"), _
                              Array(False, False, False, "DiscountMarkup3", "DiscountMarkupPriceRate3"), _
                              Array(False, False, False, "DiscountMarkup4", "DiscountMarkupPriceRate4"), _
                              Array(False, False, False, "DiscountMarkup5", "DiscountMarkupPriceRate5")), _
                        "CustomerPriceLevel IS NULL", _
                        False, _
                        False
    synchronizeMasTable "IM_PriceCode", _
                        "InventoryMaster", _
                        Array(Array(True, True, False, "ItemCode", "ItemNumber"), _
                              Array(False, False, False, "BreakQuantity1", "IBreakQty1"), _
                              Array(False, False, False, "BreakQuantity2", "IBreakQty2"), _
                              Array(False, False, False, "BreakQuantity3", "IBreakQty3"), _
                              Array(False, False, False, "BreakQuantity4", "IBreakQty4"), _
                              Array(False, False, False, "BreakQuantity5", "IBreakQty5"), _
                              Array(False, False, False, "DiscountMarkup1", "IDiscountMarkupPriceRate1"), _
                              Array(False, False, False, "DiscountMarkup2", "IDiscountMarkupPriceRate2"), _
                              Array(False, False, False, "DiscountMarkup3", "IDiscountMarkupPriceRate3"), _
                              Array(False, False, False, "DiscountMarkup4", "IDiscountMarkupPriceRate4"), _
                              Array(False, False, False, "DiscountMarkup5", "IDiscountMarkupPriceRate5")), _
                        "CustomerPriceLevel='2'", _
                        False, _
                        False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : importVendors
' DateTime  : 7/12/2005 14:34
' Author    : briandonorfio
' Purpose   : Imports vendor information from AP_Vendor.
'---------------------------------------------------------------------------------------
'
Private Sub importVendors()
    RaiseEvent StatusChanged("Querying MAS200 For Vendor Info")
    synchronizeMasTable "AP_Vendor", _
                        "AP_Vendor", _
                        Array(Array(True, True, False, "APDivisionNo", "APDivisionNo"), _
                              Array(True, True, False, "VendorNo", "VendorNo"), _
                              Array(True, False, False, "VendorName", "VendorName"), _
                              Array(True, False, False, "AddressLine1", "AddressLine1"), _
                              Array(True, False, False, "AddressLine2", "AddressLine2"), _
                              Array(True, False, False, "City", "City"), _
                              Array(True, False, False, "State", "State"), _
                              Array(True, False, False, "ZipCode", "ZipCode"), _
                              Array(True, False, False, "CountryCode", "CountryCode"), _
                              Array(True, False, False, "TelephoneNo", "TelephoneNo"), _
                              Array(True, False, False, "TermsCode", "TermsCode"), _
                              Array(True, False, False, "Reference", "Reference")), _
                        "AccountKey='140000000' OR AccountKey='99990001B'"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : importVendorTerms
' DateTime  : 7/31/2008 15:29
' Author    : briandonorfio
' Purpose   : Imports purchase order terms code details from AP_TermsCode.
'---------------------------------------------------------------------------------------
'
Private Sub importVendorTerms()
    RaiseEvent StatusChanged("Querying MAS200 For Terms Info")
    synchronizeMasTable "AP_TermsCode", _
                        "AP_TermsCode", _
                        Array(Array(True, True, False, "TermsCode", "TermsCode"), _
                              Array(True, False, False, "TermsCodeDesc", "TermsCodeDesc"), _
                              Array(False, False, False, "DaysBeforeDue", "DaysBeforeDue"), _
                              Array(True, False, False, "DueDateADayOfTheMonth", "DueDateADayOfTheMonth"), _
                              Array(False, False, False, "MinimumDaysAllowedInv", "MinimumDaysAllowedInv"), _
                              Array(False, False, False, "DaysDiscountAllowed", "DaysDiscountAllowed"), _
                              Array(True, False, False, "DiscountDateADayOfTheMo", "DiscountDateADayOfTheMo"), _
                              Array(False, False, False, "MinimumDaysAllowedDisc", "MinimumDaysAllowedDisc"), _
                              Array(False, False, False, "DiscountRate", "DiscountRate"))
                              
End Sub

'---------------------------------------------------------------------------------------
' Procedure : importWarehouses
' DateTime  : 7/12/2005 15:28
' Author    : briandonorfio
' Purpose   : Imports warehouse information from IMC_WarehouseCode.
'---------------------------------------------------------------------------------------
'
Private Sub importWarehouses()
    RaiseEvent StatusChanged("Querying MAS200 For Warehouses")
    synchronizeMasTable "IM_Warehouse", _
                        "IM_Warehouse", _
                        Array(Array(True, True, False, "WarehouseCode", "WarehouseCode"), _
                              Array(True, False, False, "WarehouseDesc", "WarehouseDesc"), _
                              Array(True, False, False, "WarehouseName", "WarehouseName"), _
                              Array(True, False, False, "WarehouseAddress1", "WarehouseAddress1"), _
                              Array(True, False, False, "WarehouseAddress2", "WarehouseAddress2"), _
                              Array(True, False, False, "WarehouseAddress3", "WarehouseAddress3"), _
                              Array(True, False, False, "WarehouseCity", "WarehouseCity"), _
                              Array(True, False, False, "WarehouseState", "WarehouseState"), _
                              Array(True, False, False, "WarehouseZipCode", "WarehouseZipCode"), _
                              Array(True, False, False, "WarehouseCountryCode", "WarehouseCountryCode"))
End Sub

'---------------------------------------------------------------------------------------
' Procedure : importItemKits
' DateTime  : 2/21/2008 11:12
' Author    : briandonorfio
' Purpose   : Imports sales kit information from IMO_SalesKitDetail. Do we need to
'             include IMN_SalesKitMasterfile for any reason? Inserts only, so everything
'             should be deleted beforehand.
'---------------------------------------------------------------------------------------
'
Public Sub importItemKits()
    RaiseEvent StatusChanged("Querying MAS200 For Item Kits")
    synchronizeMasTable "IM_SalesKitDetail", _
                        "IM_SalesKitDetail", _
                        Array(Array(True, True, False, "SalesKitNo", "SalesKitNo"), _
                              Array(True, True, False, "ComponentItemCode", "ComponentItemCode"), _
                              Array(True, False, False, "QuantityPerAssembly", "QuantityPerAssembly")), _
                        "SalesKitNo<>''"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : calcSalesHistory
' DateTime  : 8/17/2005 13:02
' Author    : briandonorfio
' Purpose   : Updates the sales history for all items. All new records from the MAS 200
'             transaction log are moved to the SQL table IM5_TransactionDetail, then
'             queried from there for speed. Most of the calculation takes place in the
'             spImportCalcSalesHist*Period stored procedures. Most of this could have
'             probably been an INSERT...SELECT, but having the screen update is nice.
'
'             Does not require triggers to be disabled.
'---------------------------------------------------------------------------------------
'
Private Sub calcSalesHistory()
    RaiseEvent StatusChanged("Querying MAS200 For Sales History Info")
    
    'This will only do a partial sync of whatever's needed. If
    'something ever gets out of whack, we can just truncate the
    'IM5 and InventoryStatistics tables, and run this with the
    'arguments for the where clause and add/delete removed.
    synchronizeMasTable "IM_ItemTransactionHistory", _
                        "IM_ItemTransactionHistory", _
                        Array(Array(True, True, False, "ItemCode", "ItemCode"), _
                              Array(True, True, False, "WarehouseCode", "WarehouseCode"), _
                              Array(True, True, False, "TransactionDate", "TransactionDate"), _
                              Array(True, True, False, "TransactionCode", "TransactionCode"), _
                              Array(True, True, False, "EntryNo", "EntryNo"), _
                              Array(True, True, False, "SequenceNo", "SequenceNo"), _
                              Array(False, False, False, "TransactionQty", "TransactionQty")), _
                        "TransactionDate > { d '" & Format(DLookup("MAX(TransactionDate)", "IM_ItemTransactionHistory", "TransactionCode='SO'"), "yyyy-mm-dd") & "' }", _
                        True, _
                        False
    
    RaiseEvent StatusChanged("Caching Dropship Mappings")
    
    Dim rst As ADODB.Recordset, i As Long
    Set rst = DB.retrieve("SELECT ItemNumber FROM InventoryMaster WHERE ItemNumber<'X%'")
    Dim dropshipCache As Dictionary
    Set dropshipCache = New Dictionary
    While Not rst.EOF
        Dim zzz As String
        zzz = ZZZifyItemNumber(rst("ItemNumber"))
        If Not dropshipCache.exists(zzz) Then
            dropshipCache.Add zzz, New PerlArray
        End If
        dropshipCache.item(zzz).Push CStr(rst("ItemNumber"))
        rst.MoveNext
    Wend
    rst.Close
    
    RaiseEvent StatusChanged("Calculating Daily Sales")
    
    Set rst = DB.retrieve("SELECT ItemCode, WarehouseCode, TransactionDate, TransactionCode, EntryNo, TransactionQty FROM IM_ItemTransactionHistory WHERE TransactionDate>=(SELECT MAX(TransactionDate) FROM InventoryStatisticsByDay)")
    Dim dayCache As Dictionary
    Set dayCache = New Dictionary
    While Not rst.EOF
        'Dim transactionType As String
        'transactionType = rst("TransactionCode")
        'If transactionType = "SI" Then
        '    'If Right(rst("EntryNo"), 3) = "-IN" Then
        '    '    transactionType = "SO"
        '    'Else
        '    '    transactionType = "P2"
        '    'End If
        '    'actually, P2 counts kits as P2, so it's only SO here
        '    transactionType = "SO"
        'End If
        'If transactionType = "SO" And rst("TransactionQty") < 0 Then
        '    transactionType = "S"
        'ElseIf transactionType = "SO" Then
        '    transactionType = "R"
        'ElseIf transactionType = "P2" And rst("TransactionQty") < 0 Then
        '    transactionType = "P"
        'ElseIf transactionType = "P2" Then
        '    transactionType = "R"
        'Else
        '    transactionType = ""
        'End If
        
        'P2 is no longer using its own module code, what the fuck!
        Dim transactionType As String
        transactionType = rst("TransactionCode")
        If transactionType = "SI" Then
            transactionType = "SO"
        End If
        If transactionType = "SO" And rst("TransactionQty") > 0 Then 'return, not separated by sales channel
            transactionType = "R"
        ElseIf transactionType = "SO" And (Right(rst("EntryNo"), 3) = "-IN" Or Right(rst("EntryNo"), 3) = "-CM" Or Right(rst("EntryNo"), 3) = "DM") Then
            transactionType = "S"
        ElseIf transactionType = "SO" And Right(rst("EntryNo"), 1) = "-" Then
            transactionType = "P"
        Else
            'unknown
            transactionType = ""
        End If
        
        If transactionType <> "" Then
            Dim whseCode As String
            whseCode = rst("WarehouseCode")
            Dim item As String
            item = rst("ItemCode")
            Dim day As String
            day = rst("TransactionDate")
            
            Dim dayKey As String
            dayKey = item & Chr(0) & whseCode & Chr(0) & day & Chr(0) & transactionType
            
            If Not dayCache.exists(dayKey) Then
                dayCache.Add dayKey, Abs(CLng(rst("TransactionQty")))
            Else
                dayCache.item(dayKey) = dayCache.item(dayKey) + Abs(CLng(rst("TransactionQty")))
            End If
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Dim iter As Variant
    For Each iter In dayCache.keys
        If dayCache.item(CStr(iter)) <> 0 Then
            Dim parts As Variant
            parts = Split(CStr(iter), Chr(0))
            
            Dim rstTemp As ADODB.Recordset
            Set rstTemp = DB.retrieve("SELECT TransactionAmount FROM InventoryStatisticsByDay WHERE ItemNumber='" & parts(0) & "' AND WarehouseCode='" & parts(1) & "' AND TransactionDate='" & parts(2) & "' AND TransactionType='" & parts(3) & "'")
            If rstTemp.EOF Then
'Debug.Print "INSERT INTO InventoryStatisticsByDay ( ItemNumber, WarehouseCode, TransactionDate, TransactionType, TransactionAmount ) VALUES ( '" & parts(0) & "', '" & parts(1) & "', '" & parts(2) & "', '" & parts(3) & "', " & dayCache.item(CStr(iter)) & " )"
                DB.Execute "INSERT INTO InventoryStatisticsByDay ( ItemNumber, WarehouseCode, TransactionDate, TransactionType, TransactionAmount ) VALUES ( '" & parts(0) & "', '" & parts(1) & "', '" & parts(2) & "', '" & parts(3) & "', " & dayCache.item(CStr(iter)) & " )"
            ElseIf rstTemp(0) <> dayCache.item(CStr(iter)) Then
'Debug.Print "UPDATE InventoryStatisticsByDay SET TransactionAmount=" & dayCache.item(CStr(iter)) & " WHERE ItemNumber='" & parts(0) & "' AND WarehouseCode='" & parts(1) & "' AND TransactionDate='" & parts(2) & "' AND TransactionType='" & parts(3) & "'"
                DB.Execute "UPDATE InventoryStatisticsByDay SET TransactionAmount=" & dayCache.item(CStr(iter)) & " WHERE ItemNumber='" & parts(0) & "' AND WarehouseCode='" & parts(1) & "' AND TransactionDate='" & parts(2) & "' AND TransactionType='" & parts(3) & "'"
            Else
                'ok
            End If
            rstTemp.Close
            
            'only count dropships that were P2
            'what about returns of dropships?
            'If Left(parts(0), 3) = "ZZZ" And parts(3) = "P" Then
            'kathleen is finalizing lots of dropships in s/o now, count those too
            If Left(parts(0), 3) = "ZZZ" And (parts(3) = "P" Or parts(3) = "S") Then
                If dropshipCache.exists(CStr(parts(0))) Then
                    For i = 0 To dropshipCache.item(CStr(parts(0))).Scalar - 1
                        
                        Dim rstTempD As ADODB.Recordset
                        Set rstTempD = DB.retrieve("SELECT TransactionAmount FROM InventoryStatisticsByDay WHERE ItemNumber='" & dropshipCache.item(CStr(parts(0))).Elem(i) & "' AND WarehouseCode='" & parts(1) & "' AND TransactionDate='" & parts(2) & "' AND TransactionType='D'")
                        If rstTempD.EOF Then
'Debug.Print "INSERT INTO InventoryStatisticsByDay ( ItemNumber, WarehouseCode, TransactionDate, TransactionType, TransactionAmount ) VALUES ( '" & dropshipCache.item(CStr(parts(0))).Elem(i) & "', '" & parts(1) & "', '" & parts(2) & "', 'D', " & dayCache.item(CStr(iter)) & " )"
                            DB.Execute "INSERT INTO InventoryStatisticsByDay ( ItemNumber, WarehouseCode, TransactionDate, TransactionType, TransactionAmount ) VALUES ( '" & dropshipCache.item(CStr(parts(0))).Elem(i) & "', '" & parts(1) & "', '" & parts(2) & "', 'D', " & dayCache.item(CStr(iter)) & " )"
                        ElseIf rstTempD(0) <> dayCache.item(CStr(iter)) Then
'Debug.Print "UPDATE InventoryStatisticsByDay SET TransactionAmount=" & dayCache.item(CStr(iter)) & " WHERE ItemNumber='" & dropshipCache.item(CStr(parts(0))).Elem(i) & "' AND WarehouseCode='" & parts(1) & "' AND TransactionDate='" & parts(2) & "' AND TransactionType='D'"
                            DB.Execute "UPDATE InventoryStatisticsByDay SET TransactionAmount=" & dayCache.item(CStr(iter)) & " WHERE ItemNumber='" & dropshipCache.item(CStr(parts(0))).Elem(i) & "' AND WarehouseCode='" & parts(1) & "' AND TransactionDate='" & parts(2) & "' AND TransactionType='D'"
                        Else
                            'ok
                        End If
                        rstTempD.Close
                    Next i
                End If
            End If
        End If
    Next iter
    
    Dim rst2 As ADODB.Recordset
    Dim j As Long
    Dim updateClause As String
    Dim existingRows As Dictionary
    
    RaiseEvent StatusChanged("Calculating Sales History")
    
    For i = 0 To 1
        For j = 0 To 5
            Set rst = DB.retrieve("EXEC spImportCalcSalesHist" & IIf(i = 0, "Static", "Rolling") & "Period " & j)
            Set existingRows = New Dictionary
            While Not rst.EOF
                DoEvents
                'check for existence/changedness, update/insert as needed
                Set rst2 = DB.retrieve("SELECT P2SalesAmount, SOSalesAmount, ReturnsAmount, DropshipsAmount FROM InventoryStatistics WHERE ItemNumber='" & rst("ItemNumber") & "' AND WhseCode='" & rst("WarehouseCode") & "' AND PeriodType=" & i & " AND Period=" & j)
                If rst2.EOF Then
                    'insert
'Debug.Print "INSERT INTO InventoryStatistics ( ItemNumber, WhseCode, PeriodType, Period, P2SalesAmount, SOSalesAmount, ReturnsAmount, DropshipsAmount ) VALUES ( '" & rst("ItemNumber") & "', '" & rst("WarehouseCode") & "', " & i & ", " & j & ", " & rst("P2Qty") & ", " & rst("SOQty") & ", " & rst("ReturnQty") & ", " & rst("DropshipQty") & " )"
                    DB.Execute "INSERT INTO InventoryStatistics ( ItemNumber, WhseCode, PeriodType, Period, P2SalesAmount, SOSalesAmount, ReturnsAmount, DropshipsAmount ) VALUES ( '" & rst("ItemNumber") & "', '" & rst("WarehouseCode") & "', " & i & ", " & j & ", " & rst("P2Qty") & ", " & rst("SOQty") & ", " & rst("ReturnQty") & ", " & rst("DropshipQty") & " )"
                Else
                    'update?
                    updateClause = ""
                    If rst("P2Qty") <> rst2("P2SalesAmount") Then
                        updateClause = "P2SalesAmount=" & rst("P2Qty")
                    End If
                    If rst("SOQty") <> rst2("SOSalesAmount") Then
                        updateClause = IIf(updateClause <> "", updateClause & ", ", "") & "SOSalesAmount=" & rst("SOQty")
                    End If
                    If rst("ReturnQty") <> rst2("ReturnsAmount") Then
                        updateClause = IIf(updateClause <> "", updateClause & ", ", "") & "ReturnsAmount=" & rst("ReturnQty")
                    End If
                    If rst("DropshipQty") <> rst2("DropshipsAmount") Then
                        updateClause = IIf(updateClause <> "", updateClause & ", ", "") & "DropshipsAmount=" & rst("DropshipQty")
                    End If
                    
                    If updateClause <> "" Then
'Debug.Print "UPDATE InventoryStatistics SET " & updateClause & " WHERE ItemNumber='" & rst("ItemNumber") & "' AND WhseCode='" & rst("WarehouseCode") & "' AND PeriodType=" & i & " AND Period=" & j
                        DB.Execute "UPDATE InventoryStatistics SET " & updateClause & " WHERE ItemNumber='" & rst("ItemNumber") & "' AND WhseCode='" & rst("WarehouseCode") & "' AND PeriodType=" & i & " AND Period=" & j
                    End If
                End If
                rst2.Close
                existingRows.Add rst("ItemNumber") & Chr(0) & rst("WarehouseCode"), 1
                rst.MoveNext
            Wend
            rst.Close
            
            'sales will eventually get stale if nothing is bought for more than the longest
            'period, this should get rid of anything that isn't updated
            Set rst = DB.retrieve("SELECT ItemNumber, WhseCode FROM InventoryStatistics WHERE PeriodType=" & i & " AND Period=" & j)
            While Not rst.EOF
                If Not existingRows.exists(rst("ItemNumber") & Chr(0) & rst("WhseCode")) Then
                    DB.Execute "DELETE FROM InventoryStatistics WHERE ItemNumber='" & rst("ItemNumber") & "' AND WhseCode='" & rst("WhseCode") & "' AND PeriodType=" & i & " AND Period=" & j
                End If
                rst.MoveNext
            Wend
            rst.Close
            
            Set existingRows = Nothing
        Next j
    Next i
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : calcFreightActual
' DateTime  : 7/21/2005 12:13
' Author    : briandonorfio
' Purpose   : Inserts new information into the FreightActual table. Info comes from the
'             following Mas200 tables:
'
'              SO_03SOHistoryHeader  <- SalesOrderNumber -> SO_04SOHistoryDetail
'                 /                       |                                \
'      SalesOrderNumber                   |                            ItemNumber
'      OrderDate                          V
'      ShipToZip             SO_37SODailyShipment <- InvoiceNumber -> SO_38SODailyShipmentPackage
'                                 /                                           \
'                            ShipDate                                     FreightAmount
'
'             General idea is to grab all recent (from newest ship date in FreightActual
'             to now) sales orders with 1 item ordered, 1 qty, and determine the freight
'             cost. Function as a single query is saved as vFreightActualCalculation,
'             but couldn't get it to work in Mas200 directly (time out errors), maybe it
'             will work in Mas200 SQL or Mas500.
'
'             ~15 min. to do 1/1/2005 to 7/21/2005
'             ~12 min. to do 7/21/2005 to 8/5/2005
'
'             Does not require triggers to be disabled.
'---------------------------------------------------------------------------------------
'
Private Sub calcFreightActual()
    RaiseEvent StatusChanged("(1 of 2) Getting invoices from MAS200")
    Dim rstInv As ADODB.Recordset
    Set rstInv = MASDB.retrieve("SELECT InvoiceNo, HeaderSeqNo, TransactionDate, SalesOrderNo, OrderDate, ShipToZipCode FROM AR_InvoiceHistoryHeader WHERE ShippingInvoice='Y' AND SalesOrderNo IS NOT NULL AND ShipVia='UPS INET' AND TransactionDate>={ d '" & Format(DLookup("MAX(DateOfSale)", "FreightActual"), "yyyy-mm-dd") & "' }")
    'Set rstInv = MASDB.MASRetrieveViaJSON("SELECT InvoiceNo, HeaderSeqNo, TransactionDate, SalesOrderNo, OrderDate, ShipToZipCode FROM AR_InvoiceHistoryHeader WHERE ShippingInvoice='Y' AND SalesOrderNo IS NOT NULL AND ShipVia='UPS INET' AND TransactionDate>={ d '2012-02-15' } AND TransactionDate<={ d '2012-12-31' }")
    
    While Not rstInv.EOF
        RaiseEvent StatusChanged("(2 of 2) Checking invoice " & rstInv("InvoiceNo") & "-" & rstInv("HeaderSeqNo"))
        'Debug.Print "(2 of 2) Checking invoice " & rstInv("InvoiceNo") & "-" & rstInv("HeaderSeqNo")
        If 0 = DLookup("COUNT(*)", "FreightActual", "DateOfShipment='" & rstInv("TransactionDate") & "' AND SalesOrderNumber='" & rstInv("SalesOrderNo") & "'") Then
            Dim rstLines As ADODB.Recordset
            'ExplodedKitItem = "N" for items, "Y" for kits, "C" for sub-items, only look for items and kits
            Set rstLines = MASDB.retrieve("SELECT ItemCode, QuantityShipped, UnitPrice FROM AR_InvoiceHistoryDetail WHERE InvoiceNo='" & rstInv("InvoiceNo") & "' AND HeaderSeqNo='" & rstInv("HeaderSeqNo") & "' AND ItemType='1' AND QuantityShipped>0 AND (ExplodedKitItem='Y' OR ExplodedKitItem='N')")
            If rstLines.RecordCount = 1 Then
                If rstLines("QuantityShipped") = 1 Then
                    Dim rstPack As ADODB.Recordset
                    Set rstPack = MASDB.retrieve("SELECT SUM(FreightAmt) FROM AR_InvoiceHistoryTracking WHERE InvoiceNo='" & rstInv("InvoiceNo") & "' AND HeaderSeqNo='" & rstInv("HeaderSeqNo") & "'")
                    Dim amt As String
                    amt = Nz(rstPack(0), "0")
                    If amt > 0 Then
                        DB.Execute "INSERT INTO FreightActual ( ItemID, ItemNumber, ZipCode, Cost, DateOfShipment, DateOfSale, SalesOrderNumber, ItemUnitPrice ) VALUES ( " & Utilities.GetItemIDByItemCode(rstLines("ItemCode")) & ", '" & EscapeSQuotes(rstLines("ItemCode")) & "', '" & Left(rstInv("ShipToZipCode"), 5) & "', " & amt & ", '" & rstInv("TransactionDate") & "', '" & rstInv("OrderDate") & "', '" & rstInv("SalesOrderNo") & "', " & rstLines("UnitPrice") & " )", True
                    End If
                    rstPack.Close
                    Set rstPack = Nothing
                End If
            End If
            rstLines.Close
            Set rstLines = Nothing
        End If
        rstInv.MoveNext
    Wend
    rstInv.Close
    Set rstInv = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : calcSalesRank
' DateTime  : 2/14/2006 11:05
' Author    : briandonorfio
' Purpose   : Calculates sales rankings for each item. Generally, it ranks items as ABCD
'             within the line code, basing it on dollars sold or units sold, in separate
'             periods of time.
'
'             Does not require triggers to be disabled.
'---------------------------------------------------------------------------------------
'
Private Sub calcSalesRank(threshold As Dictionary)
    Dim rstLines As ADODB.Recordset
    Set rstLines = DB.retrieve("SELECT * FROM vRequeryLineCodes ORDER BY ProductLine")

    Dim rstBaseline As ADODB.Recordset
    Dim rst As ADODB.Recordset
    Dim itemArrays As Dictionary
    
    Dim rankTypes() As Variant
    ' fields MUST be named similarly, assumes there's always going to be at least a *Percent and *SalesRank
    rankTypes = Array("ThisPeriodUnits", "ThisPeriodDollars", "LastPeriodUnits", "LastPeriodDollars", "YTDUnits", "YTDDollars", "LYRUnits", "LYRDollars")
    ReDim current(UBound(rankTypes)) As String
    
    Dim total As Double
    Dim strsql As String, i As Long, j As Long
    
    While Not rstLines.EOF
        If rstLines("ProductLine") <> "XXX" And rstLines("ProductLine") <> "ZZZ" Then
            RaiseEvent StatusChanged("Calculating Sales Rankings for " & rstLines("ProductLine"))
            DoEvents
            ' crunch the numbers
            '    rstBaseline contains the total sales in units and dollars for each period we're interested in for the line code in aggregate
            Set rstBaseline = DB.retrieve("exec spSalesRankBaseline '" & rstLines("ProductLine") & "'")
    
            Set rst = DB.retrieve("exec spSalesRankItemsByPL '" & rstLines("ProductLine") & "'")
            
            DB.Execute "DELETE FROM SalesRankings WHERE ItemNumber IN ( SELECT ItemNumber FROM InventoryMaster WHERE ProductLine='" & rstLines("ProductLine") & "')"
            DB.Execute "INSERT INTO SalesRankings ( ItemNumber ) SELECT ItemNumber FROM InventoryMaster WHERE ProductLine='" & rstLines("ProductLine") & "'"
    
            While Not rst.EOF
                DB.Execute "exec spSalesRankSetItemNumbers '" & rst("ItemNumber") & "', " & _
                                                                rst("SalesThisPeriod") & ", " & rst("SalesLastPeriod") & ", " & rst("SalesYTD") & ", " & rst("SalesLYR") & ", " & _
                                                                rstBaseline("SalesThisPeriod") & ", " & rstBaseline("SalesLastPeriod") & ", " & rstBaseline("SalesYTD") & ", " & rstBaseline("SalesLYR") & ", " & _
                                                                rst("DollarsThisPeriod") & ", " & rst("DollarsLastPeriod") & ", " & rst("DollarsYTD") & ", " & rst("DollarsLYR") & ", " & _
                                                                rstBaseline("DollarsThisPeriod") & ", " & rstBaseline("DollarsLastPeriod") & ", " & rstBaseline("DollarsYTD") & ", " & rstBaseline("DollarsLYR")
                rst.MoveNext
            Wend
            rst.Close
            rstBaseline.Close
            
            ' calculate sales ranking letter
            '    anything in the top 70th (by default) percentile goes into A, 70-80 B, etc
            '    save everything in a hash of arrays, itemnumber -> array of letters for each metric
            Set itemArrays = New Dictionary
    
            For i = 0 To UBound(rankTypes)
                Set rst = DB.retrieve("SELECT ItemNumber, " & rankTypes(i) & "Percent FROM SalesRankings WHERE LEFT(ItemNumber,3)='" & rstLines("ProductLine") & "' ORDER BY " & rankTypes(i) & "Percent DESC")
                total = 0
                While Not rst.EOF
                    If itemArrays.exists(rst("ItemNumber").value) Then
                        current = itemArrays.item(rst("ItemNumber").value)
                    End If
                    If total < threshold("A") Then
                        current(i) = "A"
                    ElseIf total < threshold("B") Then
                        current(i) = "B"
                    ElseIf total < threshold("C") Then
                        current(i) = "C"
                    Else
                        current(i) = "D"
                    End If
                    total = total + rst(rankTypes(i) & "Percent")
                    itemArrays.item(rst("ItemNumber").value) = current
                    rst.MoveNext
                Wend
                rst.Close
            Next i
            
            ' update the table with the info from the HoA
            For i = 0 To UBound(itemArrays.keys)
                strsql = "UPDATE SalesRankings SET "
                For j = 0 To UBound(rankTypes)
                    strsql = strsql & rankTypes(j) & "SalesRank='" & itemArrays.item(itemArrays.keys(i))(j) & "'"
                    If j <> UBound(rankTypes) Then
                        strsql = strsql & ", "
                    End If
                Next j
                strsql = strsql & " WHERE ItemNumber='" & itemArrays.keys(i) & "'"
                DB.Execute strsql
            Next i
            
            Set itemArrays = Nothing
        End If
        rstLines.MoveNext
    Wend
    rstLines.Close

    Set rstLines = Nothing
    Set rstBaseline = Nothing
    Set rst = Nothing

End Sub

Private Sub synchronizeMasTable(masTableName As String, sqlTableName As String, fieldArray As Variant, Optional whereClause As String = "", Optional insertNew As Boolean = True, Optional deleteMissing As Boolean = True)
    Dim errorMessages As String
    errorMessages = ""
    
    Dim existingRows As Dictionary
    Set existingRows = New Dictionary
    
    Dim thisfield As Variant
    Dim keystr As String
    
    Dim rstSource As ADODB.Recordset
    If Not RunningThroughVPN Then
        Set rstSource = MASDB.retrieve(getSyncSelectClause(fieldArray) & " FROM " & masTableName & IIf(whereClause = "", "", " WHERE " & whereClause))
    Else
        Set rstSource = MASDB.MASRetrieveViaJSON(getSyncSelectClause(fieldArray) & " FROM " & masTableName & IIf(whereClause = "", "", " WHERE " & whereClause))
    End If
    
    While Not rstSource.EOF
        keystr = getSyncKeyStringified(fieldArray, rstSource)
        RaiseEvent StatusChanged("Processing " & Replace(keystr, Chr(0), " "))
        
        Dim rstDest As ADODB.Recordset
        Set rstDest = DB.retrieve(getSyncSelectClause(fieldArray, False, True) & " FROM " & sqlTableName & getSyncKeyWhereClause(fieldArray, rstSource, True))
        If rstDest.EOF Then
            If insertNew Then
                Dim stmt As String
                stmt = getSyncInsertSQL(sqlTableName, fieldArray, rstSource)
'Debug.Print stmt
                If DB.Execute(stmt) Then
                    'good insert
                Else
                    'insert failure, log and continue
                    'this is a workaround for an error -2147217913 date out of range
                    'error refuses to propagate up to the error handler here, so doing it this way
                    errorMessages = errorMessages & "can't run " & qq(stmt) & vbCrLf
                End If
            End If
        Else
            Dim dirty As PerlArray
            Set dirty = New PerlArray
            For Each thisfield In fieldArray
                Dim val1 As String
                val1 = Nz(rstSource(CStr(thisfield(SYNC_POS_FIELDNAME_DST))))
                Dim val2 As String
                val2 = Nz(rstDest(CStr(thisfield(SYNC_POS_FIELDNAME_DST))))
                If thisfield(SYNC_POS_NEEDS_TRIM) Then
                    val1 = Trim(val1)
                    val2 = Trim(val2)
                End If
                '4.5 changed the size of the sentinel value 999999 -> 99999999
                'for now, replace when found. export prices sproc has a similar
                'replacement in reverse.
                If masTableName = "IM_PriceCode" Then
                    If thisfield(SYNC_POS_FIELDNAME_SRC) = "BreakQuantity1" Or thisfield(SYNC_POS_FIELDNAME_SRC) = "BreakQuantity2" Or thisfield(SYNC_POS_FIELDNAME_SRC) = "BreakQuantity3" Or thisfield(SYNC_POS_FIELDNAME_SRC) = "BreakQuantity4" Or thisfield(SYNC_POS_FIELDNAME_SRC) = "BreakQuantity5" Then
                        If val1 = "99999999" Then
                            val1 = "999999"
                        End If
                    End If
                End If
                If val1 <> val2 Then
                    dirty.Push thisfield
                End If
            Next thisfield
            If dirty.Scalar <> 0 Then
'Debug.Print getSyncUpdateSQL(sqlTableName, dirty, fieldArray, rstSource)
                DB.Execute getSyncUpdateSQL(sqlTableName, dirty, fieldArray, rstSource)
            End If
        End If
        rstDest.Close
        
        If existingRows.exists(keystr) Then
            'mas keys are messed up for table, what do we do?
        Else
            existingRows.Add keystr, 1
        End If
        
        DoEvents
        
        rstSource.MoveNext
    Wend
    rstSource.Close
    Set rstDest = Nothing
    Set rstSource = Nothing
    
    '2009-10-09 - removing whereClause check, hopefully this doesn't break everything
    'If deleteMissing And whereClause = "" Then
    If deleteMissing Then
        Dim hasHideField As Boolean
        Dim rstCol As ADODB.Recordset
        Set rstCol = DB.retrieve("SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='" & sqlTableName & "' AND COLUMN_NAME='Hide'")
        If rstCol(0) = 0 Then
            hasHideField = False
        Else
            hasHideField = True
        End If
        rstCol.Close
        Set rstCol = Nothing
        
        RaiseEvent StatusChanged("Processing deleted records...")
        Set rstSource = DB.retrieve(getSyncSelectClause(fieldArray, True, True) & " FROM " & sqlTableName)
        While Not rstSource.EOF
            If existingRows.exists(getSyncKeyStringified(fieldArray, rstSource)) Then
                'ok
            Else
                Dim doDelete As Boolean
                doDelete = False
                
                If hasHideField Then
                    doDelete = Not CBool(DLookup("Hide", sqlTableName, getSyncKeyWhereClause(fieldArray, rstSource, True)))
                Else
                    doDelete = True
                End If
                
                If doDelete And masTableName = "CI_Item" Then
                    errorMessages = errorMessages & "Not deleting item from CI_Item, key is " & qq(getSyncKeyWhereClause(fieldArray, rstSource, True)) & vbCrLf
                    doDelete = False
                End If
                If doDelete And masTableName = "IM_ItemWarehouse" Then
                    errorMessages = errorMessages & "Not deleting item from IM_ItemWarehouse, key is " & qq(getSyncKeyWhereClause(fieldArray, rstSource, True)) & vbCrLf
                    doDelete = False
                End If
                
                If doDelete Then
'Debug.Print "DELETE FROM " & sqlTableName & getSyncKeyWhereClause(fieldArray, rstSource, True)
                    DB.Execute "DELETE FROM " & sqlTableName & getSyncKeyWhereClause(fieldArray, rstSource, True)
                End If
            End If
            
            rstSource.MoveNext
        Wend
        rstSource.Close
        Set rstSource = Nothing
        
        DoEvents
    End If
    
    If errorMessages <> "" Then
        'Open Environ("TEMP") & "\poinv\errormessages.txt" For Output As #1
        'Print #1, errorMessages
        'Close #1
        'SendEmailTo "brian@tools-plus.com", "sync mas table error", Environ("TEMP") & "\poinv\errormessages.txt"
        SendEmailTo "brian@tools-plus.com", "sync mas table error", errorMessages
    End If
End Sub

Private Function getSyncSelectClause(fullFieldList As Variant, Optional keysOnly As Boolean = False, Optional useDest As Boolean = False) As String
    Dim sql As String
    Dim thisfield As Variant
    For Each thisfield In fullFieldList
        If Not keysOnly Or thisfield(SYNC_POS_IS_KEY) Then
            sql = IIf(sql = "", "", sql & ", ")
            If useDest Or thisfield(SYNC_POS_FIELDNAME_DST) = thisfield(SYNC_POS_FIELDNAME_SRC) Then
                sql = sql & thisfield(SYNC_POS_FIELDNAME_DST)
            Else
                sql = sql & thisfield(SYNC_POS_FIELDNAME_SRC) & " AS " & thisfield(SYNC_POS_FIELDNAME_DST)
            End If
        End If
    Next thisfield
    
    getSyncSelectClause = "SELECT " & sql
End Function

Private Function getSyncKeyWhereClause(fullFieldList As Variant, rst As ADODB.Recordset, Optional useDest As Boolean = False) As String
    Dim whereClause As String
    Dim thisfield As Variant
    For Each thisfield In fullFieldList
        If thisfield(SYNC_POS_IS_KEY) Then
            whereClause = IIf(whereClause = "", "", whereClause & " AND ") & CStr(thisfield(IIf(useDest, SYNC_POS_FIELDNAME_DST, SYNC_POS_FIELDNAME_SRC))) & "="
            Dim fname As String
            fname = Nz(rst(CStr(thisfield(SYNC_POS_FIELDNAME_DST))))
            If thisfield(SYNC_POS_NEEDS_TRIM) Then
                fname = Trim(fname)
            End If
            If thisfield(SYNC_POS_NEEDS_ESCAPE) Then
                whereClause = whereClause & "'" & EscapeSQuotes(fname) & "'"
            Else
                whereClause = whereClause & fname
            End If
        End If
    Next thisfield
    
    getSyncKeyWhereClause = IIf(whereClause = "", "", " WHERE " & whereClause)
End Function

Private Function getSyncInsertSQL(tableName As String, fullFieldList As Variant, rst As ADODB.Recordset) As String
    Dim cols As String, vals As String
    Dim thisfield As Variant
    For Each thisfield In fullFieldList
        cols = IIf(cols = "", "", cols & ", ") & CStr(thisfield(SYNC_POS_FIELDNAME_DST))
        vals = IIf(vals = "", "", vals & ", ")
        Dim thisVal As String
        thisVal = Nz(rst(CStr(thisfield(SYNC_POS_FIELDNAME_DST))))
        If thisfield(SYNC_POS_NEEDS_TRIM) Then
            thisVal = Trim(thisVal)
        End If
        If thisfield(SYNC_POS_NEEDS_ESCAPE) Then
            vals = vals & "'" & EscapeSQuotes(thisVal) & "'"
        Else
            vals = vals & thisVal
        End If
    Next thisfield
    
    getSyncInsertSQL = "INSERT INTO " & tableName & " ( " & cols & " ) VALUES ( " & vals & " )"
End Function

Private Function getSyncUpdateSQL(tableName As String, dirtyList As PerlArray, fullFieldList As Variant, rst As ADODB.Recordset) As String
    Dim fieldNum As Long
    Dim thisfield As Variant
    Dim sql As String
    'For Each thisfield In dirtyList.DataArray
    For fieldNum = 0 To dirtyList.Scalar - 1
        thisfield = dirtyList.Elem(fieldNum)
        sql = IIf(sql = "", "", sql & ", ") & CStr(thisfield(SYNC_POS_FIELDNAME_DST)) & "="
        Dim thisVal As String
        thisVal = Nz(rst(CStr(thisfield(SYNC_POS_FIELDNAME_DST))))
        If thisfield(SYNC_POS_NEEDS_TRIM) Then
            thisVal = Trim(thisVal)
        End If
        If thisfield(SYNC_POS_NEEDS_ESCAPE) Then
            sql = sql & "'" & EscapeSQuotes(thisVal) & "'"
        Else
            sql = sql & thisVal
        End If
    'Next thisfield
    Next fieldNum
    
    getSyncUpdateSQL = "UPDATE " & tableName & " SET " & sql & getSyncKeyWhereClause(fullFieldList, rst, True)
End Function

Private Function getSyncKeyStringified(fullFieldList As Variant, rst As ADODB.Recordset)
    Dim thisfield As Variant
    Dim keystr As String
    For Each thisfield In fullFieldList
        If thisfield(SYNC_POS_IS_KEY) Then
            Dim thisVal As String
            thisVal = Nz(rst(CStr(thisfield(SYNC_POS_FIELDNAME_DST))))
            If thisfield(SYNC_POS_NEEDS_TRIM) Then
                thisVal = Trim(thisVal)
            End If
            keystr = IIf(keystr = "", "", keystr & Chr(0)) & thisVal
        End If
    Next thisfield
    
    getSyncKeyStringified = keystr
End Function

