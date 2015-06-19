Attribute VB_Name = "Logic"
'---------------------------------------------------------------------------------------
' Module    : Logic
' DateTime  : 9/7/2005 10:37
' Author    : briandonorfio
' Purpose   : General business logic functions, similar to utilities but not really.
'
'             Dependencies:
'               - Microsoft ActiveX Data Objects 2.8 Library
'               - Microsoft Scripting Runtime
'               - DatabaseFunctions + DBConn
'               - MouseHourglass + global Mouse object
'               - SyncShell
'               - Utilities
'               - EBay
'               - PerlArray
'---------------------------------------------------------------------------------------

Option Explicit

Private productLineDropshipChargeCache As Dictionary 'used in CalculateDropshipCostFor

'---------------------------------------------------------------------------------------
' Procedure : GetPurchaseOrderKeys
' Author    : briandonorfio
' Date      : 3/8/2011
' Purpose   : Returns a variant array of the unique (well, sort of) identifiers that
'             make up a PO.
'---------------------------------------------------------------------------------------
'
Public Function GetPurchaseOrderKeys(item As String) As Variant
    ' this is duplicated in the warehousing program TP.BusinessObject.Inventory.Item.PurchaseOrderKeys,
    ' any changes here or there should be carried over
    Mouse.Hourglass True
    
    Dim retval(0 To 2) As String
    'retval 1 = vendor
    'retval 2 = coop vendor
    'retval 3 = line code
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber, PrimaryVendor, ProductLine FROM InventoryMaster WHERE ItemNumber='" & item & "'")
    Dim v As String, cv As String, prodLine As String
    v = rst("PrimaryVendor") 'po vendor is the ITEM's primary vendor, which can be a coop
    prodLine = rst("ProductLine")
    
    'fix xxx/zzz items
    If rst("ProductLine") = "XXX" Or rst("ProductLine") = "ZZZ" Then
        prodLine = Mid(rst("ItemNumber"), 4, 3)
    End If
    
    Dim rstPL As ADODB.Recordset
    Set rstPL = DB.retrieve("SELECT Coop, RealVendor FROM ProductLine WHERE ProductLine='" & prodLine & "'")
    '6/17/06 - error w/ missing PO, probably mostly for XXX items?
    If rstPL.EOF Then
        MsgBox "Unable to find the line code " & prodLine & ", can't create PO!"
        GetPurchaseOrderKeys = Nothing 'probably crash, but oh well, this doesn't happen often, if ever during normal usage
        Mouse.Hourglass False
        Exit Function
    End If
    If rstPL("Coop") Then
        cv = rstPL("RealVendor") 'coop vendor is the real vendor, or the regular vendor
    Else
        cv = v
    End If
    rstPL.Close
    Set rstPL = Nothing
    
    'lc is ZZZ, or the lc-subbed line code
    Dim rstPO As ADODB.Recordset
    Dim altlc As String, lc As String
    If rst("ProductLine") = "XXX" Then
        altlc = Mid(item, 4, 3)
    Else
        altlc = rst("ProductLine")
    End If
    lc = DLookup("RealLineCode", "ProductLineSubForPOs", "AlternateLineCode='" & altlc & "'")
    If lc = "" Then
        lc = altlc
    End If
    
    retval(0) = v
    retval(1) = cv
    retval(2) = lc
    GetPurchaseOrderKeys = retval
    
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : CreatePurchaseOrder
' DateTime  : 6/16/2005 16:25
' Author    : briandonorfio
' Purpose   : Creates a new purchase order for the given item's primary vendor. Returns
'             the Purchase Order ID. This is way more complicated than it has any right
'             to be.
'
'             ZZZ DropShip PO's must be done one at a time for each vendor, no way of
'             differentiating two separate ones. Nothing bad will happen, just that the
'             item gets stuck to the first one.
'
'             May return a finalized PO, should check for this before trying to add an
'             item to a purchase order.
'---------------------------------------------------------------------------------------
'
Public Function CreatePurchaseOrder(item As String) As String
    'If DLookup("COUNT(*)", "BaseItemList", "BaseItem='" & item & "'") <> 0 Then
    '    MsgBox "This is a BASE ITEM, and does not go on a PO!!"
    '    CreatePurchaseOrder = ""
    '    Exit Function
    'End If
    'TODO: check for web pointered-ness, give warning?
    
    'Dim rst As ADODB.Recordset
    'Set rst = DB.retrieve("exec spInventoryMasterByItem '" & item & "'")
    'Dim vendor As String, CoopVendor As String
    'vendor = rst("PrimaryVendor")
    'Dim rstPL As ADODB.Recordset
    'Dim pl As String
    'If rst("ProductLine") = "XXX" Or rst("ProductLine") = "ZZZ" Then
    '    pl = Mid(rst("ItemNumber"), 4, 3)
    'Else
    '    pl = rst("ProductLine")
    'End If
    'Set rstPL = DB.retrieve("SELECT Coop, RealVendor FROM ProductLine WHERE ProductLine='" & pl & "'")
    ''6/17/06 - error w/ missing PO, probably mostly for XXX items?
    'If rstPL.EOF Then
    '    MsgBox "Unable to find the line code " & pl & ", can't create PO!"
    '    CreatePurchaseOrder = ""
    '    Exit Function
    'End If
    'If rstPL("Coop") Then
    '    CoopVendor = rstPL("RealVendor")
    'Else
    '    CoopVendor = vendor
    'End If
    'rstPL.Close
    'Set rstPL = Nothing
    'Dim rstPO As ADODB.Recordset
    'Dim altlc As String, lc As String
    'If rst("ProductLine") = "XXX" Then
    '    altlc = Mid(item, 4, 3)
    'Else
    '    altlc = rst("ProductLine")
    'End If
    'lc = DLookup("RealLineCode", "ProductLineSubForPOs", "AlternateLineCode='" & altlc & "'")
    'If lc = "" Then
    '    lc = altlc
    'End If
    Dim pokeys As Variant
    pokeys = GetPurchaseOrderKeys(item)
    
    Mouse.Hourglass True
    Dim rstPO As ADODB.Recordset
    Set rstPO = DB.retrieve("SELECT ID FROM PurchaseOrders WHERE Exported=0 AND VendorNumber='" & pokeys(0) & "' AND CoopVendor='" & pokeys(1) & "' AND ProductLineCode='" & pokeys(2) & "'")
    If rstPO.EOF Then
        Dim POTerms As String
        POTerms = DLookup("TermsCode", "AP_Vendor", "VendorNo='" & pokeys(0) & "'")
        If POTerms = "" Then
            POTerms = "00"
        End If
        DB.Execute "INSERT INTO PurchaseOrders ( VendorNumber, CoopVendor, ProductLineCode, POTerms ) VALUES ( '" & pokeys(0) & "', '" & pokeys(1) & "', '" & pokeys(2) & "', '" & POTerms & "' )"
        Dim newid As String
        newid = DLookup("ID", "PurchaseOrders", "Exported=0 AND VendorNumber='" & pokeys(0) & "' AND CoopVendor='" & pokeys(1) & "' AND ProductLineCode='" & pokeys(2) & "'")
        '10/1/09 - use table defaults (prospect whse)
        CreatePurchaseOrder = newid
    Else
        CreatePurchaseOrder = rstPO("ID")
    End If
    rstPO.Close
    Set rstPO = Nothing
    Mouse.Hourglass False
End Function

Public Function PurchaseOrderLineEdit(POID As Long, item As String, newQty As Long) As Boolean
    If newQty = 0 Then
        DB.Execute "DELETE FROM PurchaseOrderLines WHERE HeaderID=" & POID & " AND ItemNumber='" & item & "'"
    ElseIf 0 = DLookup("COUNT(*)", "PurchaseOrderLines", "HeaderID=" & POID & " AND ItemNumber='" & item & "'") Then
        'DB.Execute "INSERT INTO PurchaseOrderLines ( HeaderID, ItemNumber, Quantity ) VALUES ( " & POID & ", '" & item & "', " & newQty & " )"
        DB.Execute "INSERT INTO PurchaseOrderLines ( HeaderID, ItemNumber, ItemID, Quantity ) VALUES ( " & POID & ", '" & item & "', " & Utilities.GetItemIDByItemCode(item) & ", " & newQty & " )"
    Else
        DB.Execute "UPDATE PurchaseOrderLines SET Quantity=" & newQty & " WHERE HeaderID=" & POID & " AND ItemNumber='" & item & "'"
    End If
End Function

Public Function ConvertToDropshipPO(POID As Long) As Boolean
    Dim soNum As String
    soNum = InputBox("Enter Sales Order number:")
    If soNum = "" Then
        ConvertToDropshipPO = False
        Exit Function
    End If
    soNum = UCase(soNum)
    If IsNumeric(Mid(soNum, 1, 1)) Then
        While Len(soNum) < 7
            soNum = "0" & soNum
        Wend
    End If
    
    Dim m As Mas200ImportExport
    Set m = New Mas200ImportExport
    Dim so As Variant
    so = m.GetSalesOrderInfo(soNum)
    Set m = Nothing
    
    If IsNull(so) Then
        MsgBox "Can't find sales order " + soNum + " in Mas S/O Entry!"
        ConvertToDropshipPO = False
        Exit Function
    End If
    
    If CDbl(so(8)) = 0 Then
        MsgBox "Warning: no deposit has been taken for this order!" & vbCrLf & vbCrLf & "TODO: brian should probably check the payment type here too."
    End If
    
    Dim Notes As String
    'notes = "CUSTOMER PHONE NUMBER " & so(7) & vbCrLf & _
    '        "ORDER NUMBER " + soNum
    ' margie doesn't want the phone number entered, always manual?
    'Notes = "CUSTOMER PHONE NUMBER " & vbCrLf & _
    '        "ORDER NUMBER " + soNum
    DB.Execute "UPDATE PurchaseOrders SET ShiptoCode='1234', " & _
                                         "ShipToName='" & EscapeSQuotes(CStr(so(1))) & "', " & _
                                         "ShipToAddress1='" & EscapeSQuotes(CStr(so(2))) & "', " & _
                                         "ShipToAddress2='" & EscapeSQuotes(CStr(so(3))) & "', " & _
                                         "ShipToCity='" & EscapeSQuotes(CStr(so(4))) & "', " & _
                                         "ShipToState='" & EscapeSQuotes(CStr(so(5))) & "', " & _
                                         "ShipToZip='" & EscapeSQuotes(CStr(so(6))) & "', " & _
                                         "MiscNotes='" & EscapeSQuotes(Notes) & "', " & _
                                         "SalesOrderNo='" & EscapeSQuotes(soNum) & "', " & _
                                         "CustomerEmailAddress='" & Left(EscapeSQuotes(CStr(so(9))), 35) & "', " & _
                                         "LiftGateRequired=1, " & _
                                         "AddressZoningType=1 " & _
               "WHERE ID=" & POID
    ConvertToDropshipPO = True
End Function

'---------------------------------------------------------------------------------------
' Procedure : CreateItem
' DateTime  : 8/14/2006 15:01
' Author    : briandonorfio
' Purpose   : Creates a new item in InventoryMaster and related necessary tables. If
'             invalid item number or error inserting, then returns false, else true.
'---------------------------------------------------------------------------------------
'
Public Function CreateItem(newitem As String, websiteID As Long) As Boolean
    newitem = UCase(newitem)
    Dim exists As Boolean
    exists = ExistsInInventoryMaster(newitem)
    If Not PLExists(Left$(newitem, 3)) Then
        MsgBox "Could not find valid line code. Create a new line code first."
    ElseIf Not validateItemNumber(newitem) Then
        MsgBox "Must be between 4 and 27 chars, only alphanumeric or -"
    ElseIf exists Then
        MsgBox "Item number already exists."
    Else
        If Len(newitem) > 27 Then
            If MsgBox("This item number is longer than 27 characters, and will make a weird dropship item, continue?", vbYesNo) = vbNo Then
                Exit Function
            End If
        End If
        'If Right(newitem, 4) = "-KIT" Then
        '    If vbNo = MsgBox("Since this ends in ""-KIT"", it will be set up as a kit item. You also need to do this manually in Mas! Do you want to continue?", vbYesNo) Then
        '        Exit Function
        '    End If
        'End If
        If Not DB.Execute("INSERT INTO InventoryMaster ( ItemNumber, WebsiteID, ProductLine, PrimaryVendor ) VALUES ( '" & newitem & "', " & websiteID & ", '" & Left$(newitem, 3) & "', '" & lookupLineCodePrimaryVendor(IIf(Left$(newitem, 3) = "XXX" Or Left$(newitem, 3) = "ZZZ", Mid$(newitem, 4, 3), Left$(newitem, 3))) & "' )") Then
            MsgBox "Error creating " & newitem
        Else
            Dim itemID As String
            itemID = DLookup("@@IDENTITY", "InventoryMaster")
            If Not DB.Execute("INSERT INTO InventoryQuantities ( ItemID, ItemNumber, WhseCode ) VALUES ( " & itemID & ", '" & newitem & "', '000' )") Then
                MsgBox "Error creating " & newitem & " in InventoryQuantities."
            Else
                DB.Execute "INSERT INTO InventoryComponents ( Component ) VALUES ( '" & newitem & "' )"
                Dim cid As String
                cid = DLookup("@@IDENTITY", "InventoryComponents")
                DB.Execute "INSERT INTO InventoryComponentMap ( ItemID, ItemNumber, ComponentID ) VALUES ( " & itemID & ", '" & newitem & "', " & cid & " )"
                If Left(newitem, 3) <> "XXX" And Left(newitem, 3) <> "ZZZ" Then
                    Dim availLimit As String
                    availLimit = DLookup("AvailLimitDefault", "ProductLine", "ProductLine='" & Left$(newitem, 3) & "'")
                    If Not DB.Execute("INSERT INTO PartNumbers ( ItemNumber, Desc2, AvailabilityLimit ) VALUES ( '" & newitem & "', '" & EscapeSQuotes(Mid(newitem, 4)) & "', " & availLimit & " )") Then
                        MsgBox "Error creating " & newitem & " in part numbers."
                    Else
                        '2013-03-28: set ebay category by line code, if possible
                        Dim id As String
                        id = EBay.EBayGetProbableStoreCategoryIDFor(newitem)
                        If id <> "" Then
                            DB.Execute "UPDATE PartNumbers SET EBayStoreCategoryID=" & id & " WHERE ItemNumber='" & newitem & "'"
                        End If
                    End If
                    If DLookup("DefaultDropShippable", "ProductLine", "ProductLine='" & Left(newitem, 3) & "'") Then
                        updateInventoryMaster "ItemStatus", SetItemDSable(newitem), newitem, ""
                    End If
                ElseIf Left(newitem, 3) = "XXX" Then
                    updateInventoryMaster "ItemStatus", "0", newitem, ""
                ElseIf Left(newitem, 3) = "ZZZ" Then
                    updateInventoryMaster "ItemStatus", "4", newitem, ""
                End If
                CreateItem = True
            End If
        End If
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : CreateReplacePart
' DateTime  : 6/14/2005 11:06
' Author    : briandonorfio
' Purpose   : Creates a new item in po/inv and signs, with description, specs, web paths
'             and other general information copied into the new item. If new item
'             already exists, no action taken. Returns true if it creates a new record
'             in IM or PN, otherwise false.
'---------------------------------------------------------------------------------------
'
Public Function CreateReplacePart(newitem As String, olditem As String) As Boolean
    newitem = UCase(newitem)
    olditem = UCase(olditem)
    Mouse.Hourglass True
    
    If CreateItem(newitem, CLng(DLookup("WebsiteID", "InventoryMaster", "ItemNumber='" & olditem & "'"))) Then
        Dim i As Long
        
        updateInventoryMaster "ItemDescription", "REPLACES " & olditem & " " & DLookup("ItemDescription", "InventoryMaster", "ItemNumber='" & olditem & "'"), newitem, "'"
        updateInventoryMaster "ReplacementFor", olditem, newitem, "'"
        
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT SignID, AvailabilityLimit, Desc1, Desc3, Desc4, Spec1HTML, Spec2HTML, Spec3HTML, Spec4HTML, Spec5HTML, Spec6HTML, Spec7HTML, Spec8HTML FROM PartNumbers WHERE ItemNumber='" & olditem & "'")
        Dim sql As String
        For i = 0 To 1
            sql = sql & rst(i).name & "=" & rst(i).value & ", "
        Next i
        For i = 2 To rst.fields.count - 1
            If Nz(rst(i)) <> "" Then
                sql = sql & rst(i).name & "='" & IIf(Left(rst(i).name, 4) = "Spec", "COPIED--", "") & EscapeSQuotes(rst(i)) & "', "
            End If
        Next i
        sql = Left(sql, Len(sql) - 2)
        DB.Execute "UPDATE PartNumbers SET " & sql & " WHERE ItemNumber='" & newitem & "'"
        rst.Close
        
        Set rst = DB.retrieve("SELECT WebPathID FROM PartNumberWebPaths WHERE ItemNumber='" & olditem & "'")
        While Not rst.EOF
            DB.Execute "INSERT INTO PartNumberWebPaths ( ItemNumber, WebPathID ) VALUES ( '" & newitem & "', " & rst("WebPathID") & " )"
            rst.MoveNext
        Wend
        rst.Close
        
        Set rst = Nothing
        CreateReplacePart = True
    Else
        CreateReplacePart = False
    End If
    
    Mouse.Hourglass False
End Function

Public Function CreateItemAsKit(item As String, websiteID As Long, kitArray As PerlArray) As Boolean 'perlarray of variant arrays (item (str), qty (long))
    'basic validation
    item = UCase(item)
    CreateItemAsKit = False
    If Not PLExists(Left(item, 3)) Then
        MsgBox "Could not find valid line code. Create a new line code first."
        Exit Function
    End If
    If Left(item, 3) = "XXX" Or Left(item, 3) = "ZZZ" Then
        MsgBox "Can't create special order or dropship kits"
        Exit Function
    End If
    If Not validateItemNumber(item) Then
        MsgBox "Must be between 4 and 27 chars, only alphanumeric or -"
        Exit Function
    End If
    Dim exists As Boolean
    exists = ExistsInInventoryMaster(item)
    If exists Then
        MsgBox "Item number already exists."
        Exit Function
    End If
    
    'insert item and quantity lines
    If Not DB.Execute("INSERT INTO InventoryMaster ( ItemNumber, WebsiteID, ProductLine, PrimaryVendor, IsMASKit ) VALUES ( '" & item & "', " & websiteID & ", '" & Left(item, 3) & "', '" & lookupLineCodePrimaryVendor(Left(item, 3)) & "', 1 )") Then
        MsgBox "Error creating " & item
        Exit Function
    End If
    Dim itemID As String
    itemID = DLookup("@@IDENTITY", "InventoryMaster")
    If Not DB.Execute("INSERT INTO InventoryQuantities ( ItemID, ItemNumber, WhseCode ) VALUES ( " & itemID & ", '" & item & "', '000' )") Then
        MsgBox "Error creating " & item & " in InventoryQuantities"
        Exit Function
    End If
    
    'if it's a single item/multiple quantity kit, check for quantity breaks at this tier, apply price
    If kitArray.Scalar = 1 Then
        Dim rstPB As ADODB.Recordset
        Set rstPB = DB.retrieve("SELECT IBreakQty1, IBreakQty2, IBreakQty3, IBreakQty4, IBreakQty5, IDiscountMarkupPriceRate1, IDiscountMarkupPriceRate2, IDiscountMarkupPriceRate3, IDiscountMarkupPriceRate4, IDiscountMarkupPriceRate5 FROM InventoryMaster WHERE ItemNumber='" & kitArray.Elem(0)(0) & "'")
        Dim pb As Long
        For pb = 1 To 5
            If rstPB("IBreakQty" & pb) >= 999999 Or rstPB("IBreakQty" & pb) + 1 = kitArray.Elem(0)(1) Then
                'no breaks found, so just apply the lowest price
                ' OR
                'exact break found, use this
                DB.Execute "UPDATE InventoryMaster SET IDiscountMarkupPriceRate1=" & kitArray.Elem(0)(1) * rstPB("IDiscountMarkupPriceRate" & pb) & ", EBayPrice=" & kitArray.Elem(0)(1) * rstPB("IDiscountMarkupPriceRate" & pb) & " WHERE ItemNumber='" & item & "'"
                Exit For
            End If
        Next pb
        rstPB.Close
        Set rstPB = Nothing
    End If
    
    'item description
    Dim kitDesc As String
    If kitArray.Scalar = 1 Then
        'if it's a single item/multiple quantity kit, change the description to "{x} PACK {item desc}"
        kitDesc = kitArray.Elem(0)(1) & " PK " & DLookup("ItemDescription", "InventoryMaster", "ItemNumber='" & kitArray.Elem(0)(0) & "'")
    Else
        'multiple item kit, description defaults to "KIT - {first item desc}"
        kitDesc = "KIT - " & DLookup("ItemDescription", "InventoryMaster", "ItemNumber='" & kitArray.Elem(0)(0) & "'")
    End If
    updateInventoryMaster "ItemDescription", kitDesc, item, "'"
    
    'since we're filtering non-XXX/ZZZ, add to PN by default
    Dim availLimit As String, desc1 As String, desc2 As String, desc3 As String
    availLimit = DLookup("AvailLimitDefault", "ProductLine", "ProductLine='" & Left(item, 3) & "'")
    If kitArray.Scalar = 1 Then
        'single item/multiple quantity, "{x} PACK {desc2} {desc3}"
        desc1 = kitArray.Elem(0)(1) & " Pk"
        desc2 = DLookup("Desc2", "PartNumbers", "ItemNumber='" & kitArray.Elem(0)(0) & "'")
        desc3 = DLookup("Desc3", "PartNumbers", "ItemNumber='" & kitArray.Elem(0)(0) & "'")
    Else
        'multiple item pack, "{itemnumber} {1st desc3} KIT"
        desc1 = ""
        desc2 = EscapeSQuotes(Mid(item, 4))
        desc3 = DLookup("Desc3", "PartNumbers", "ItemNumber='" & kitArray.Elem(0)(0) & "'") & " Kit"
    End If
    If Not DB.Execute("INSERT INTO PartNumbers ( ItemNumber, Desc1, Desc2, Desc3, AvailabilityLimit ) VALUES ( '" & item & "', '" & EscapeSQuotes(desc1) & "', '" & EscapeSQuotes(desc2) & "', '" & EscapeSQuotes(desc3) & "', " & availLimit & " )") Then
        MsgBox "Error creating " & item & " in part numbers."
        Exit Function
    End If
    Dim id As String
    id = EBay.EBayGetProbableStoreCategoryIDFor(item)
    If id <> "" Then
        DB.Execute "UPDATE PartNumbers SET EBayStoreCategoryID=" & id & " WHERE ItemNumber='" & item & "'"
    End If
    
    'add sub-items to IM_SalesKitDetail so they're usable immediately (offload errors/OOS otherwise)
    'the mas export actually pulls directly from IM_SalesKitDetail, so that's actually kind of
    'convenient?
    'sum up the sub-items and convert to components...
    Dim i As Long, subItem As String, subQty As String
    Dim componentHash As Dictionary
    Set componentHash = New Dictionary
    For i = 0 To kitArray.Scalar - 1
        subItem = kitArray.Elem(i)(0)
        subQty = kitArray.Elem(i)(1)
        
        DB.Execute "INSERT INTO IM_SalesKitDetail ( SalesKitNo, ComponentItemCode, QuantityPerAssembly, CoreComponent ) VALUES ( '" & item & "', '" & subItem & "', " & subQty & ", " & IIf(i = 0, "1", "0") & " )"
        
        Dim rst As ADODB.Recordset
        Dim subItemID As String
        subItemID = Utilities.GetItemIDByItemCode(subItem)
        Set rst = DB.retrieve("SELECT ComponentID, Quantity FROM InventoryComponentMap WHERE ItemID=" & subItemID)
        While Not rst.EOF
            If componentHash.exists(CLng(rst("ComponentID"))) Then
                componentHash.item(CLng(rst("ComponentID"))) = CLng(componentHash.item(CLng(rst("ComponentID"))) + CLng(rst("Quantity")))
            Else
                componentHash.Add CLng(rst("ComponentID")), CLng(rst("Quantity") * subQty)
            End If
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
    Next i
    
    '...so we can add them to the component map
    Dim iter As Variant
    For Each iter In componentHash.keys
        DB.Execute "INSERT INTO InventoryComponentMap ( ItemID, ItemNumber, ComponentID, Quantity ) VALUES ( " & itemID & ", '" & item & "', " & CStr(iter) & ", " & CStr(componentHash.item(CLng(iter))) & " )"
    Next iter
    
    'now we add up the includes...
    Dim sortIter As Long
    sortIter = 0
    For i = 0 To kitArray.Scalar - 1
        subItem = kitArray.Elem(i)(0)
        subQty = kitArray.Elem(i)(1)
        Set rst = DB.retrieve("SELECT IncludeText FROM PartNumberIncludes WHERE ItemNumber='" & subItem & "' ORDER BY SortOrder")
        While Not rst.EOF
            DB.Execute "INSERT INTO PartNumberIncludes (ItemNumber, IncludeText, SortOrder) VALUES ('" & item & "', '" & EscapeSQuotes(subQty & "x " & rst("IncludeText")) & "', " & sortIter & ")"
            sortIter = sortIter + 1
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
    Next i
    
    '...cross sells...
    Dim crossSellDict As Dictionary
    Set crossSellDict = New Dictionary
    For i = 0 To kitArray.Scalar - 1
        subItem = kitArray.Elem(i)(0)
        Set rst = DB.retrieve("SELECT ShowText, URL FROM PartNumberCrossSell WHERE ItemNumber='" & subItem & "' ORDER BY SortOrder")
        While Not rst.EOF
            If Not crossSellDict.exists(CStr(rst("ShowText"))) Then
                crossSellDict.Add CStr(rst("ShowText")), CStr(rst("URL"))
            End If
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
    Next i
    sortIter = 0
    For Each iter In crossSellDict.keys
        DB.Execute "INSERT INTO PartNumberCrossSell (ItemNumber, ShowText, URL, SortOrder) VALUES ('" & item & "', '" & EscapeSQuotes(CStr(iter)) & "', '" & EscapeSQuotes(crossSellDict.item(CStr(iter))) & "', " & sortIter & ")"
        sortIter = sortIter + 1
    Next iter
    
    '...and extended specs
    Dim specsDict As Dictionary
    Set specsDict = New Dictionary
    Dim fieldsArray As Variant
    fieldsArray = Array("ExtendedSpecsHTML", "WriteupHTML", "FeaturesHTML", "TechSpecsHTML", "MediaHTML", "NotesHTML", "QuestionsHTML")
    For Each iter In fieldsArray
        specsDict.Add CStr(iter), ""
    Next iter
    For i = 0 To kitArray.Scalar - 1
        subItem = kitArray.Elem(i)(0)
        Set rst = DB.retrieve("SELECT " & Join(fieldsArray, ",") & " FROM PartNumbers WHERE ItemNumber='" & subItem & "'")
        For Each iter In fieldsArray
            If Nz(rst(CStr(iter))) <> "" Then
                specsDict.item(CStr(iter)) = specsDict.item(CStr(iter)) & rst(CStr(iter)) & vbCrLf
            End If
        Next iter
    Next i
    For Each iter In fieldsArray
        If specsDict.item(CStr(iter)) <> "" Then
            DB.Execute "UPDATE PartNumbers SET " & CStr(iter) & "='" & EscapeSQuotes(specsDict.item(CStr(iter))) & "' WHERE ItemNumber='" & item & "'"
        End If
    Next iter
    
    CreateItemAsKit = True
End Function

Public Function CountPriceDifferent() As Long
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT NumDifferent FROM vPriceDifferent")
    CountPriceDifferent = rst("NumDifferent")
    rst.Close
    Set rst = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : PriceDifferent
' DateTime  : 6/14/2005 11:01
' Author    : briandonorfio
' Purpose   : Checks for differences between the StdPrice and Store Price Break. Run
'             when InventoryMaster form opens up.
'---------------------------------------------------------------------------------------
'
'Public Function PriceDifferent() As Boolean
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT NumDifferent FROM vPriceDifferent")
'    If rst("NumDifferent") <> 0 Then
'        MsgBox "The number of Items where the Store Price & the QtyBreakPrice differs are: " & rst("NumDifferent") & "." & vbCrLf & "Please press All Items when you finish changing the prices."
'        PriceDifferent = True
'    Else
'        PriceDifferent = False
'    End If
'    rst.Close
'    Set rst = Nothing
'End Function

'---------------------------------------------------------------------------------------
' Procedure : CheckForExpiredSpecials
' DateTime  : 7/8/2005 15:16
' Author    : briandonorfio
' Purpose   : Checks the PartNumberSpecials table to find any expired specials. Expired
'             but postmarked, or new and need to go up, get flagged for offload. Expired
'             completely or after postmark, just display a msgbox.
'---------------------------------------------------------------------------------------
'
Public Sub CheckForExpiredSpecials()
    Mouse.Hourglass True
    DB.Execute "EXEC spSpecialsExpiredSetOffload"
    DB.Execute "EXEC spSpecialsNewSetOffload"
    Dim txt As String
    txt = ""
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber, SpecialName FROM vSpecialsRemove")
    If Not rst.EOF Then
        Dim temp As Dictionary
        Set temp = New Dictionary
        While Not rst.EOF
            If temp.exists(CStr(rst("SpecialName"))) Then
                temp.item(CStr(rst("SpecialName"))) = 1
            Else
                temp.item(CStr(rst("SpecialName"))) = temp.item(CStr(rst("SpecialName"))) + 1
            End If
            rst.MoveNext
        Wend
        Dim iter As Variant
        For Each iter In temp.keys
            txt = txt & iter & " with " & temp.item(CStr(iter)) & " associated part numbers" & vbCrLf
        Next iter
        MsgBox "The following expired specials are still associated with part numbers:" & vbCrLf & vbCrLf & txt
        Set temp = Nothing
    End If
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CheckInStockOrBackOrder
' DateTime  : 8/15/2005 09:36
' Author    : briandonorfio
' Purpose   : Refreshes the instock/backorder flags for Yahoo. Save the availability
'             amounts and outofstock flags to two hashes, then loop through everything
'             in the inventoryquantities table, finding everything that's been changed.
'             If given an itemnumber, checks just that one, otherwise defaults to all.
'---------------------------------------------------------------------------------------
'
Public Sub CheckInStockOrBackOrder2(Optional singleItem As String = "")
    Dim rst As ADODB.Recordset
    Dim temp As Dictionary
    Dim item As String
    Dim subItem As String
    Dim itemAdjustedQOHYahoo As Long
    Dim itemAdjustedQOHEBay As Long
    Dim itemAvailLimit As Long
    Dim itemEBayPublished As Boolean
    Dim itemEBayStatus As EBayStatus
    Dim itemOutOfStock As Boolean
    Dim itemEBayReserve As Long
    Dim itemEBayAvailable As Long
    Dim itemEBaySold As Long
    Dim currentQuantityAvailableAtEBay As Long
    Dim iter As Variant, iter2 As Variant
    
    Mouse.Hourglass True
    Set rst = DB.retrieve("SELECT ItemNumber, AvailabilityLimit, OutOfStock, EBayPublished, EBayStatusID, EBayReserveQty, EBayAvailableQty, EBaySoldQty FROM PartNumbers WHERE ItemNumber<'XX%'" & IIf(singleItem = "", "", " AND ItemNumber='" & singleItem & "'"))
    Dim itemMiscH As Dictionary
    Set itemMiscH = New Dictionary
    While Not rst.EOF
        itemMiscH.Add CStr(rst("ItemNumber")), Array(CLng(rst("AvailabilityLimit")), _
                                                     CBool(rst("OutOfStock")), _
                                                     CBool(rst("EBayPublished")), _
                                                     CLng(rst("EBayStatusID")), _
                                                     CLng(rst("EBayReserveQty")), _
                                                     CLng(rst("EBayAvailableQty")), _
                                                     CLng(rst("EBaySoldQty")) _
                                                    )
        rst.MoveNext
    Wend
    rst.Close
    
    'kitQtyHash:      subitem -> kit     -> qty in this kit
    'kitYStockHash:       kit -> subitem -> qty of kit we can make with this subitem (default 0)
    'kitEStockHash:       kit -> subitem -> qty of kit we can make with this subitem (default 0)
    Dim kitQtyHash As Dictionary, kitYStockHash As Dictionary, kitEStockHash As Dictionary
    Set kitQtyHash = New Dictionary
    Set kitYStockHash = New Dictionary
    Set kitEStockHash = New Dictionary
    If singleItem = "" Then 'on single items, no kit processing at all
        Set rst = DB.retrieve("SELECT SalesKitNo, ComponentItemCode, QuantityPerAssembly FROM IM_SalesKitDetail WHERE SalesKitNo<'XX%'")
        While Not rst.EOF
            item = rst("SalesKitNo")
            subItem = rst("ComponentItemCode")
            
            If kitQtyHash.exists(subItem) Then
                Set temp = kitQtyHash.item(subItem)
            Else
                Set temp = New Dictionary
            End If
            temp.Add item, CLng(rst("QuantityPerAssembly"))
            Set kitQtyHash.item(subItem) = temp
            Set temp = Nothing
            
            If kitYStockHash.exists(item) Then
                Set temp = kitYStockHash.item(item)
            Else
                Set temp = New Dictionary
            End If
            temp.Add subItem, 0
            Set kitYStockHash.item(item) = temp
            Set temp = Nothing
            
            If kitEStockHash.exists(item) Then
                Set temp = kitEStockHash.item(item)
            Else
                Set temp = New Dictionary
            End If
            temp.Add subItem, 0
            Set kitEStockHash.item(item) = temp
            Set temp = Nothing
            
            rst.MoveNext
        Wend
        rst.Close
    End If
    
    'ALL ITEMS, CHECK NEW QOH/QOSO AGAINST LIMITS, UPDATE FLAGS
    Set rst = DB.retrieve("SELECT InventoryMaster.ItemNumber, vActualWhseQty.QtyOnHand, vActualWhseQty.QtyOnSO, dbo.StatusIsDSonly(InventoryMaster.ItemStatus) AS DropshipOnly, dbo.StatusIsDSable(InventoryMaster.ItemStatus) AS Dropshippable, InventoryMaster.VendorQuantityOnHand " & _
                          "FROM InventoryMaster INNER JOIN vActualWhseQty ON InventoryMaster.ItemNumber=vActualWhseQty.ItemNumber INNER JOIN PartNumbers ON InventoryMaster.ItemNumber=PartNumbers.ItemNumber " & _
                          "WHERE PartNumbers.WebPointered=0 AND InventoryMaster.ItemNumber<'XX%'" & IIf(singleItem = "", "", " AND InventoryMaster.ItemNumber='" & singleItem & "'"))
    While Not rst.EOF
        DoEvents
        item = rst("ItemNumber")
        
        If itemMiscH.exists(item) Then
            itemAvailLimit = itemMiscH.item(item)(0)
            itemOutOfStock = itemMiscH.item(item)(1)
            itemEBayPublished = itemMiscH.item(item)(2)
            itemEBayStatus = EBay.EBayCurrentStatus(item, CLng(itemMiscH.item(item)(3)))
            itemEBayReserve = itemMiscH.item(item)(4)
            itemEBayAvailable = itemMiscH.item(item)(5)
            itemEBaySold = itemMiscH.item(item)(6)
        Else
            itemAvailLimit = 0
            itemOutOfStock = False
            itemEBayPublished = False
            itemEBayStatus = EBayStatus.EBayStatusDown
            itemEBayReserve = 0
            itemEBayAvailable = 0
            itemEBaySold = 0
        End If
        
        itemAdjustedQOHYahoo = Logic.YahooGetAvailableQuantityFor( _
            item, _
            rst("QtyOnHand"), _
            Nz(rst("VendorQuantityOnHand"), 0), _
            rst("QtyOnSO"), _
            rst("Dropshippable"), _
            rst("DropshipOnly"), _
            itemAvailLimit _
          )
        'If rst("DropshipOnly") = False Then
        '    itemAdjustedQOHEBay = EBay.EBayGetAvailableQuantityFor(item, rst("QtyOnHand"), rst("QtyOnSO"), itemAvailLimit, itemEBayReserve)
        'Else
        '    itemAdjustedQOHEBay = EBay.EBayGetAvailableQuantityFor(item, Nz(rst("VendorQuantityOnHand"), "0"), 0, itemAvailLimit, itemEBayReserve)
        'End If
        'currentQuantityAvailableAtEBay = itemEBayAvailable - itemEBaySold
        itemAdjustedQOHEBay = EBay.EBayGetAvailableQuantityFor( _
            item, _
            rst("QtyOnHand"), _
            Nz(rst("VendorQuantityOnHand"), "0"), _
            rst("QtyOnSO"), _
            rst("Dropshippable"), _
            rst("DropshipOnly"), _
            itemAvailLimit, _
            itemEBayReserve _
          )
        currentQuantityAvailableAtEBay = itemEBayAvailable - itemEBaySold
        
        If kitYStockHash.exists(item) = True Then
            'is a kit, skip this check, we check all kit stock flags after running
            'regular items
        ElseIf itemMiscH.exists(item) = False Then
            'not found in partnumbers? very weird!
            'this generally happens on wrong part numbers that get added and quickly
            'removed, but hang around in im2. pretty ignorable, i think.
            Debug.Print "weird: " & item & " found in vActualWhseQty but not PartNumbers!"
        Else
            'good item, check for required adjustments to status
            
            'YAHOO STORES
            checkInStockOrBackorderItemCheckYahoo item, itemAdjustedQOHYahoo, itemOutOfStock
            
            'EBAY
            If itemEBayPublished = True Then
                checkInStockOrBackOrderItemCheckEBay item, itemAdjustedQOHEBay, currentQuantityAvailableAtEBay, itemEBayStatus
            End If
        End If
        
        'check through kits for any containing this item
        'if found then check quantity against required qty for kit, and
        'set how many we can make. remember these are hashed differently!
        'qty by sub then master, stock by master then sub
        If kitQtyHash.exists(item) Then 'it's a component of at least one kit
            For Each iter In kitQtyHash.item(item).keys
                'iter is the kit
                'item is the subitem
                Dim qtyInThisKit As Long
                qtyInThisKit = kitQtyHash.item(item).item(CStr(iter))
                kitYStockHash.item(CStr(iter)).item(item) = CLng(Int(itemAdjustedQOHYahoo / qtyInThisKit))
                kitEStockHash.item(CStr(iter)).item(item) = CLng(Int(itemAdjustedQOHEBay / qtyInThisKit))
            Next iter
        End If
        rst.MoveNext
    Wend
    rst.Close
    
    'check kits now (this will skip (no keys) if single item mode)
    'Y/E are the same keys, so combine here
    For Each iter In kitYStockHash.keys
        item = CStr(iter)
        
        Dim numKitsCreatable As Dictionary
        Set numKitsCreatable = New Dictionary
        numKitsCreatable.Add "Y", 99999
        numKitsCreatable.Add "E", 99999
        For Each iter2 In kitYStockHash.item(item).keys
            subItem = CStr(iter2)
            If kitYStockHash.item(item).item(subItem) = 0 Then
                numKitsCreatable.item("Y") = 0
                Exit For
            ElseIf kitYStockHash.item(item).item(subItem) < numKitsCreatable.item("Y") Then
                numKitsCreatable.item("Y") = kitYStockHash.item(item).item(subItem)
            End If
        Next iter2
        For Each iter2 In kitEStockHash.item(item).keys
            subItem = CStr(iter2)
            If kitEStockHash.item(item).item(subItem) = 0 Then
                numKitsCreatable.item("E") = 0
                Exit For
            ElseIf kitEStockHash.item(item).item(subItem) < numKitsCreatable.item("E") Then
                numKitsCreatable.item("E") = kitEStockHash.item(item).item(subItem)
            End If
        Next iter2
        
        If itemMiscH.exists(item) Then
            'avail limit not used for kits
            itemOutOfStock = itemMiscH.item(item)(1)
            itemEBayPublished = itemMiscH.item(item)(2)
            itemEBayStatus = EBay.EBayCurrentStatus(item, CLng(itemMiscH.item(item)(3)))
            itemEBayReserve = itemMiscH.item(item)(4)
            itemEBayAvailable = itemMiscH.item(item)(5)
            itemEBaySold = itemMiscH.item(item)(6)
        Else
            'avail limit not used for kits
            itemOutOfStock = False
            itemEBayPublished = False
            itemEBayStatus = EBayStatus.EBayStatusDown
            itemEBayReserve = 0
            itemEBayAvailable = 0
            itemEBaySold = 0
        End If
        
        itemAdjustedQOHYahoo = numKitsCreatable.item("Y") 'no qtyonso/availlimit applied (already factored in subitems)
        itemAdjustedQOHEBay = numKitsCreatable.item("E") 'same
        currentQuantityAvailableAtEBay = itemEBayAvailable - itemEBaySold
        
        'YAHOO STORES
        checkInStockOrBackorderItemCheckYahoo item, itemAdjustedQOHYahoo, itemOutOfStock
        
        'EBAY
        If itemEBayPublished = True Then
            checkInStockOrBackOrderItemCheckEBay item, itemAdjustedQOHEBay, currentQuantityAvailableAtEBay, itemEBayStatus
        End If
    Next iter
    
    Mouse.Hourglass False
End Sub

Private Sub checkInStockOrBackorderItemCheckYahoo(item As String, itemAdjustedQOH As Long, itemOutOfStock As Boolean)
    If itemAdjustedQOH > 0 Then        'is currently in stock
        If itemOutOfStock = True Then  'but is marked out of stock
            Debug.Print "yahoo: " & item & " is in stock"
            DB.Execute "UPDATE PartNumbers SET LastChanged='" & Now() & "', Changes=2, OutOfStock=0 WHERE ItemNumber='" & item & "'"
        End If
    Else                               'is currently out of stock
        If itemOutOfStock = False Then 'but is marked as in stock
            Debug.Print "yahoo: " & item & " is out of stock"
            DB.Execute "UPDATE PartNumbers SET LastChanged='" & Now() & "', Changes=2, OutOfStock=1 WHERE ItemNumber='" & item & "'"
        End If
    End If
End Sub

Private Sub checkInStockOrBackOrderItemCheckEBay(item As String, itemAdjustedQOHEBay As Long, currentQuantityAvailableAtEBay As Long, itemEBayStatus As EBayStatus)
    If itemAdjustedQOHEBay > 0 Then                                                                     'and is currently in stock
        If EBay.EBayIsItemQtyUpdateRequired(currentQuantityAvailableAtEBay, itemAdjustedQOHEBay) Then   'and is inconsistent with ebay
            Debug.Print "ebay:  " & item & " is in stock"
            EBay.EBayMarkRevisionRequired item, itemEBayStatus                                          'revision will adjust quantity
        End If
    Else                                                                                                'otherwise is currently out of stock
        If itemEBayStatus = EBay.EBayStatusUp Then
            If currentQuantityAvailableAtEBay > 0 Then                                                  'and is available at ebay
                Debug.Print "ebay:  " & item & " is out of stock"
                EBay.EBayMarkRevisionRequired item, itemEBayStatus                                      'revision will be to remove
            End If
        End If
    End If
End Sub

Public Function YahooGetAvailableQuantityFor(item As String, Optional qoh As Long = -99999, Optional vendorQOH As Long = -99999, Optional qoso As Long = -99999, Optional dsOnly As Long = -99999, Optional dsable As Long = -99999, Optional availLimit As Long = -99999) As Long
    ' this logic for in/out thresholds (specifically about the QtyOnSO being included or not) is
    ' duplicated in EBay.EBayGetAvailableQuantityFor() and Feeds.pm:_get_stock_quantity(). any
    ' changes here should also be made there.
    Dim rst As ADODB.Recordset
    If qoh = -99999 Or qoso = -99999 Then
        Set rst = DB.retrieve("SELECT QtyOnHand, QtyOnSO FROM vActualWhseQty WHERE ItemNumber='" & item & "'")
        If qoh = -99999 Then
            qoh = rst("QtyOnHand")
        End If
        If qoso = -99999 Then
            qoso = rst("QtyOnSO")
        End If
        rst.Close
        Set rst = Nothing
    End If
    If availLimit = -99999 Then
        Set rst = DB.retrieve("SELECT AvailabilityLimit FROM PartNumbers WHERE ItemNumber='" & item & "'")
        availLimit = rst("AvailabilityLimit")
        rst.Close
        Set rst = Nothing
    End If
    If vendorQOH = -99999 Then
        Set rst = DB.retrieve("SELECT VendorQuantityOnHand FROM InventoryMaster WHERE ItemNumber='" & item & "'")
        vendorQOH = rst("VendorQuantityOnHand")
        rst.Close
        Set rst = Nothing
    End If
    If dsOnly = -99999 Or dsable = -99999 Then
        Set rst = DB.retrieve("SELECT DSonly, DSable FROM vItemStatusBreakdown WHERE ItemNumber='" & item & "'")
        If dsOnly = -99999 Then
            dsOnly = rst("DSonly")
        End If
        If dsable = -99999 Then
            dsable = rst("DSable")
        End If
        rst.Close
        Set rst = Nothing
    End If
    
    If dsOnly Then
        YahooGetAvailableQuantityFor = vendorQOH - availLimit
    ElseIf dsable Then
        YahooGetAvailableQuantityFor = qoh + vendorQOH - qoso - availLimit
    Else
        YahooGetAvailableQuantityFor = qoh - qoso - availLimit
    End If
End Function

Public Function YahooGetAvailableQuantityForKit(kit As String) As Long
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT K.ComponentItemCode, K.QuantityPerAssembly, I.QtyOnHand, I.QtyOnSO, P.AvailabilityLimit, M.VendorQuantityOnHand, B.DSonly, B.DSable " & _
                          "FROM IM_SalesKitDetail AS K " & _
                          "       INNER JOIN vActualWhseQty AS I ON K.ComponentItemCode=I.ItemNumber " & _
                          "       INNER JOIN PartNumbers AS P ON K.ComponentItemCode=P.ItemNumber " & _
                          "       INNER JOIN InventoryMaster AS M ON K.ComponentItemCode=M.ItemNumber " & _
                          "       INNER JOIN vItemStatusBreakdown AS B ON K.ComponentItemCode=B.ItemNumber " & _
                          "WHERE K.SalesKitNo='" & kit & "'")
    Dim retval As Long
    retval = 999999
    While Not rst.EOF
        Dim thisQty As Long
        thisQty = YahooGetAvailableQuantityFor( _
            rst("ComponentItemCode"), _
            rst("QtyOnHand"), _
            Nz(rst("VendorQuantityOnHand"), 0), _
            rst("QtyOnSO"), _
            rst("DSonly"), _
            rst("DSable"), _
            rst("AvailabilityLimit") _
          )
        thisQty = Int(thisQty / CLng(rst("QuantityPerAssembly")))
        If thisQty < retval Then
            retval = thisQty
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    If retval = 999999 Then
        YahooGetAvailableQuantityForKit = 0
    Else
        YahooGetAvailableQuantityForKit = retval
    End If
End Function




'---------------------------------------------------------------------------------------
' Procedure : PriceCompareSingleItemOrLC
' DateTime  : 8/3/2005 09:45
' Author    : briandonorfio
' Purpose   : Deletes entries from PriceComparisions, then calls perl script to run.
'             If given itemOrLC is three characters, then assumes it's a line code, and
'             runs for every item in that line code, otherwise assumes that it's a
'             single item, and runs just for that. Optional booleans to close window
'             when finished, defaults to keep open, hide window completely, defaults to
'             show. If given values to hide window and wait, then it overrides hide.
'---------------------------------------------------------------------------------------
'
Public Sub PriceCompareSingleItemOrLC(itemOrLC As String, Optional closeWhenDone As Boolean = False, Optional hideDOSBox As Boolean = False)
    Mouse.Hourglass True
    If Not closeWhenDone And hideDOSBox Then
        hideDOSBox = False
    End If
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT UserName, DateStamp FROM ProcessFlags WHERE Task='wpc'")
    If rst("UserName") <> "NOTRUNNING" Then
        MsgBox "A price comparison process was already started by " & rst("UserName") & " at " & rst("DateStamp") & ", wait until that one finishes."
    Else
        DB.Execute "UPDATE ProcessFlags SET UserName='" & Environ("UserName") & "' WHERE Task='wpc'"
        'If Len(itemOrLC) = 3 Then
        '    DB.Execute "DELETE FROM PriceComparisons WHERE ItemNumber LIKE '" & itemOrLC & "%'"
        '    ShellWait PERL & " " & WPC & " -f " & itemOrLC & " -t " & itemOrLC & " -c " & IIf(closeWhenDone, "1", "0"), IIf(hideDOSBox, vbHide, vbNormalFocus)
        'Else
        '    DB.Execute "DELETE FROM PriceComparisons WHERE ItemNumber='" & itemOrLC & "'"
        '    ShellWait PERL & " " & WPC & " -i " & itemOrLC & " -c " & IIf(closeWhenDone, "1", "0"), IIf(hideDOSBox, vbHide, vbNormalFocus)
        'End If
        ShellWait PERL & " " & WPC & " " & itemOrLC, IIf(hideDOSBox, vbHide, vbNormalFocus)
        DB.Execute "UPDATE ProcessFlags SET UserName='NOTRUNNING' WHERE Task='wpc'"
    End If
    Mouse.Hourglass False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FixQtyBreaks
' DateTime  : 6/14/2005 11:01
' Author    : briandonorfio
' Purpose   : Sets all quantity breaks and prices to what they should be when an item is
'             discontinued. Had some problems with quantity breaks not always clearing
'             correctly when an item was discontinued, seem to be cleared up now, but
'             leaving this here.
'---------------------------------------------------------------------------------------
'
Public Function FixQtyBreaks()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber FROM InventoryMaster WHERE IDiscountMarkupPriceRate1=0 and IBreakQty2<>0")
    While Not rst.EOF
        DB.Execute ("UPDATE InventoryMaster SET IDiscountMarkupPriceRate1=0, IDiscountMarkupPriceRate2=0, IDiscountMarkupPriceRate3=0, IDiscountMarkupPriceRate4=0, IDiscountMarkupPriceRate5=0, IBreakQty1=999999, IBreakQty2=0, IBreakQty3=0, IBreakQty4=0, IBreakQty5=0 WHERE ItemNumber='" & rst("ItemNumber") & "'")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = DB.retrieve("SELECT ItemNumber FROM InventoryMaster WHERE DiscountMarkupPriceRate1=88888.88 and BreakQty2<>0")
    While Not rst.EOF
        DB.Execute ("UPDATE InventoryMaster SET DiscountMarkupPriceRate1=88888.88, DiscountMarkupPriceRate2=0, DiscountMarkupPriceRate3=0, DiscountMarkupPriceRate4=0, DiscountMarkupPriceRate5=0, BreakQty1=999999, BreakQty2=0, BreakQty3=0, BreakQty4=0, BreakQty5=0 WHERE ItemNumber='" & rst("ItemNumber") & "'")
       rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : CreatePONumber
' DateTime  : 7/27/2005 13:29
' Author    : briandonorfio
' Purpose   : Creates a PO number for a given purchase order POID, returns PO number.
'---------------------------------------------------------------------------------------
'
'Public Function CreatePONumber(POID As String) As String
    'Dim nextPONumber As Long
    'If DLookup("PONumber", "PurchaseOrders", "ID=" & POID) = "" Then
    '    nextPONumber = DLookup("NextPurchaseOrderNumber", "PurchaseOrderConfig")
    '    DB.Execute "exec spExportCreatePONumber '" & "A" & nextPONumber & "', '" & POID & "'"
    '    DB.Execute "exec spExportSetNextPONumber '" & nextPONumber + 1 & "'"
    '    CreatePONumber = "A" & nextPONumber
    'End If
Public Function CreatePONumber(purchaseOrderID As String) As String
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT AutoNumberValue, FormatCSharp FROM MasAutoNumber WHERE AutoNumberName='PO_PurchaseOrder_PurchaseOrderNo'")
    Dim newPONumber As String
    newPONumber = Replace(rst("FormatCSharp"), "{0}", rst("AutoNumberValue"))
    DB.Execute "UPDATE PurchaseOrders SET PONumber='" & newPONumber & "' WHERE ID=" & purchaseOrderID
    DB.Execute "UPDATE MasAutoNumber SET AutoNumberValue=AutoNumberValue+1 WHERE AutoNumberName='PO_PurchaseOrder_PurchaseOrderNo'"
    rst.Close
    Set rst = Nothing
    CreatePONumber = newPONumber
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetAvgFreightCost
' DateTime  : 8/23/2005 14:45
' Author    : briandonorfio
' Purpose   : Returns the average actual freight cost for an item.
'---------------------------------------------------------------------------------------
'
Public Function GetAvgFreightCost(item As String) As String
    Dim rst As ADODB.Recordset
    'Set rst = DB.retrieve("exec spFreightActualAvg '" & item & "'")
    Set rst = DB.retrieve("SELECT AVG(Cost) AS AvgCost FROM FreightActual WHERE ItemNumber='" & item & "' AND DateOfShipment > DATEADD(YEAR, -1, GETDATE())")
    GetAvgFreightCost = IIf(IsNull(rst("Avgcost")), 0, rst("AvgCost"))
    rst.Close
    Set rst = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : ExistsInInventoryMaster
' DateTime  : 8/19/2005 16:06
' Author    : briandonorfio
' Purpose   : Simple test to see if an item exists in inventory master.
'---------------------------------------------------------------------------------------
'
Public Function ExistsInInventoryMaster(item As String) As Boolean
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("exec spExistsInInventoryMaster '" & item & "'")
    If rst("NumOccurrences") = 0 Then
        ExistsInInventoryMaster = False
    Else
        ExistsInInventoryMaster = True
    End If
    rst.Close
    Set rst = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : ExistsInPartNumbers
' DateTime  : 8/19/2005 16:08
' Author    : briandonorfio
' Purpose   : Simple test to see if an item exists in part numbers.
'---------------------------------------------------------------------------------------
'
Public Function ExistsInPartNumbers(item As String) As Boolean
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("exec spExistsInPartNumbers '" & item & "'")
    If rst("NumOccurrences") = 0 Then
        ExistsInPartNumbers = False
    Else
        ExistsInPartNumbers = True
    End If
    rst.Close
    Set rst = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : ExistsInInventoryQuantities
' DateTime  : 8/19/2005 16:11
' Author    : briandonorfio
' Purpose   : Simple test to see if an item exists in inventory quantities. Requires the
'             item number and a warehouse code.
'---------------------------------------------------------------------------------------
'
Public Function ExistsInInventoryQuantities(item As String, whseCode As String) As Boolean
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("exec spExistsInInventoryQuantities '" & item & "', '" & whseCode & "'")
    If rst("NumOccurrences") = 0 Then
        ExistsInInventoryQuantities = False
    Else
        ExistsInInventoryQuantities = True
    End If
    rst.Close
    Set rst = Nothing
End Function

Public Function CountCustomerInStockRequestsAged() As Long
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT COUNT(*) FROM vCustomerAlertsStockItemInfo WHERE InitialDaysAgo>60")
    CountCustomerInStockRequestsAged = rst(0)
    rst.Close
    Set rst = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : CountInventoryQuantityTriggers
' DateTime  : 10/4/2006 12:35
' Author    : briandonorfio
' Purpose   : Returns the number of active qty triggers
'---------------------------------------------------------------------------------------
'
Public Function CountInventoryQuantityTriggers(Optional triggerFlagged As Long = -1, Optional triggerTypeStr As String = "") As Long
    Dim wc As PerlArray
    Set wc = New PerlArray
    If triggerFlagged <> -1 Then '-1 = all, 0 = untriggered, 1 = triggered
        wc.Push "Triggered=" & triggerFlagged
    End If
    If triggerTypeStr <> "" Then
        wc.Push "TriggerType='" & EscapeSQuotes(triggerTypeStr) & "'"
    End If
    Dim whereClause As String
    If wc.Scalar = 0 Then
        whereClause = ""
    Else
        whereClause = " WHERE " & wc.Join(" AND ")
    End If
    Dim rst As ADODB.Recordset
    'Set rst = DB.retrieve("SELECT COUNT(*) FROM vCheckInventoryQuantityTriggers" & IIf(TriggerType = "", "", " WHERE TriggerType='" & EscapeSQuotes(TriggerType) & "'"))
    Set rst = DB.retrieve("SELECT COUNT(*) FROM vInventoryTriggers" & whereClause)
    CountInventoryQuantityTriggers = rst(0)
    rst.Close
    Set rst = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : CheckInventoryQuantityTriggers
' DateTime  : 3/6/2006 12:39
' Author    : briandonorfio
' Purpose   : Checks for items with a currently-fired quantity trigger. Displays a basic
'             yes/no msgbox for clearing the trigger.
'---------------------------------------------------------------------------------------
'
'Public Function CheckInventoryQuantityTriggers()
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT ItemNumber, QtyThreshold, AboveOrBelow, QtyOnHand, Notes FROM vCheckInventoryQuantityTriggers ORDER BY ItemNumber")
'    Dim msg As String
'    While Not rst.EOF
'        msg = "Inventory Quantity Threshold Alert!" & vbCrLf & vbCrLf
'        msg = msg & rst("ItemNumber") & " is " & rst("AboveOrBelow") & " the threshold of " & rst("QtyThreshold") & " (currently at " & rst("QtyOnHand") & ")" & vbCrLf & vbCrLf
'        msg = msg & "Notes:" & vbCrLf & rst("Notes") & vbCrLf & vbCrLf & "Clear this alert?"
'        If MsgBox(msg, vbYesNo + vbDefaultButton2) = vbYes Then
'            DB.Execute "DELETE FROM InventoryQuantityTriggers WHERE ItemNumber='" & rst("ItemNumber") & "'"
'        End If
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'End Function

'---------------------------------------------------------------------------------------
' Procedure : ZZZifyItemNumber
' DateTime  : 3/6/2006 12:42
' Author    : briandonorfio
' Purpose   : Returns a dropship (ZZZ) itemnumber for the given itemnumber. If it's
'             short enough, just add the ZZZ. Otherwise, strip out dashes. If that still
'             isn't enough, start removing characters from the end till it fits.
'---------------------------------------------------------------------------------------
'
Public Function ZZZifyItemNumber(item As String) As String
'    Dim retval As String
'    retval = item
'    If Len(retval) > 12 Then
'        If CBool(InStr(4, retval, "-")) Then
'            retval = Left(retval, 3) & Replace(Mid(retval, 4), "-", "")
'        End If
'    End If
'    If Len(retval) > 12 Then
'        retval = Left(retval, 12)
'    End If
'    ZZZifyItemNumber = "ZZZ" & retval
    If Left(item, 3) = "ZZZ" Then
        ZZZifyItemNumber = item
    ElseIf Left(item, 3) = "XXX" Then
        ZZZifyItemNumber = "ZZZ" & Mid(item, 4)
    Else
        ZZZifyItemNumber = "ZZZ" & item 'items limited to 27 chars, so this will always work
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : CreateZZZClone
' DateTime  : 3/6/2006 12:43
' Author    : briandonorfio
' Purpose   : Creates a new item in the database, with the equivalent ZZZ itemnumber to
'             the given itemnumber. All basic info is copied. EPN, components?
'---------------------------------------------------------------------------------------
'
Public Function CreateZZZClone(item As String) As Boolean
    If Left(item, 3) = "XXX" Or Left(item, 3) = "ZZZ" Then
        CreateZZZClone = False
    Else
        Dim newitem As String
        newitem = ZZZifyItemNumber(item)
        If newitem = "" Then
            CreateZZZClone = False
        ElseIf ExistsInInventoryMaster(newitem) Then
            CreateZZZClone = False
        Else
            
            Dim rst As ADODB.Recordset
            Set rst = DB.retrieve("SELECT * FROM InventoryMaster WHERE ItemNumber='" & item & "'")
            
            DB.Execute "INSERT INTO InventoryMaster (ItemNumber, WebsiteID, ItemDescription, ItemStatus, ProductLine, PrimaryVendor, IsMASKit, StdPack, ShippingType) " & _
                       "VALUES ('" & newitem & "', " & rst("WebsiteID") & ", '" & EscapeSQuotes(rst("ItemDescription")) & "', " & rst("ItemStatus") & ", 'ZZZ', '" & rst("PrimaryVendor") & "', " & IIf(rst("IsMASKit"), "1", "0") & ", " & rst("StdPack") & ", " & rst("ShippingType") & " )"
            Dim newitemID As String
            newitemID = DLookup("@@IDENTITY", "InventoryMaster")
            
            If rst("IsMASKit") Then
                'DB.Execute "INSERT INTO IM_SalesKitDetail (SalesKitNo, ComponentItemCode, QuantityPerAssembly, CoreComponent) " & _
                '           "SELECT '" & newitem & "', ComponentItemCode, QuantityPerAssembly, CoreComponent " & _
                '           "FROM IM_SalesKitDetail " & _
                '           "WHERE SalesKitNo='" & rst("ItemNumber") & "'"
                Dim rst2 As ADODB.Recordset
                Set rst2 = DB.retrieve("SELECT ComponentItemCode, QuantityPerAssembly, CoreComponent FROM IM_SalesKitDetail WHERE SalesKitNo='" & rst("ItemNumber") & "'")
                While Not rst2.EOF
                    If Not ExistsInInventoryMaster(ZZZifyItemNumber(rst2("ComponentItemCode"))) Then
                        CreateZZZClone rst2("ComponentItemCode")
                    End If
                    DB.Execute "INSERT INTO IM_SalesKitDetail (SalesKitNo, ComponentItemCode, QuantityPerAssembly, CoreComponent) VALUES ('" & newitem & "', '" & rst2("ComponentItemCode") & "', " & rst2("QuantityPerAssembly") & ", " & rst2("CoreComponent") & ")"
                    rst2.MoveNext
                Wend
                rst2.Close
            End If
            
            DB.Execute "INSERT INTO InventoryComponentMap ( ItemID, ItemNumber, ComponentID, Quantity ) " & _
                       "SELECT " & newitemID & ", '" & newitem & "', ComponentID, Quantity " & _
                       "FROM InventoryComponentMap " & _
                       "WHERE ItemID=" & rst("ItemID")
            
            If rst("DiscountMarkupPriceRate1") <> 0 Then
                LogItemStorePriceChanged newitem, rst("DiscountMarkupPriceRate1")
                DB.Execute "UPDATE InventoryMaster " & _
                           "SET BreakQty1=" & rst("BreakQty1") & ", BreakQty2=" & rst("BreakQty2") & ", BreakQty3=" & rst("BreakQty3") & ", BreakQty4=" & rst("BreakQty4") & ", BreakQty5=" & rst("BreakQty5") & ", " & _
                               "DiscountMarkupPriceRate1=" & rst("DiscountMarkupPriceRate1") & ", DiscountMarkupPriceRate2=" & rst("DiscountMarkupPriceRate2") & ", DiscountMarkupPriceRate3=" & rst("DiscountMarkupPriceRate3") & ", DiscountMarkupPriceRate4=" & rst("DiscountMarkupPriceRate4") & ", DiscountMarkupPriceRate5=" & rst("DiscountMarkupPriceRate5") & ", " & _
                               "StdPrice=" & rst("StdPrice") & " " & _
                           "WHERE ItemID=" & newitemID
            End If
            
            If rst("IDiscountMarkupPriceRate1") <> 0 Then
                LogItemWebPriceChanged newitem, rst("IDiscountMarkupPriceRate1")
                DB.Execute "UPDATE InventoryMaster " & _
                           "SET IBreakQty1=" & rst("IBreakQty1") & ", IBreakQty2=" & rst("IBreakQty2") & ", IBreakQty3=" & rst("IBreakQty3") & ", IBreakQty4=" & rst("IBreakQty4") & ", IBreakQty5=" & rst("IBreakQty5") & ", " & _
                               "IDiscountMarkupPriceRate1=" & rst("IDiscountMarkupPriceRate1") & ", IDiscountMarkupPriceRate2=" & rst("IDiscountMarkupPriceRate2") & ", IDiscountMarkupPriceRate3=" & rst("IDiscountMarkupPriceRate3") & ", IDiscountMarkupPriceRate4=" & rst("IDiscountMarkupPriceRate4") & ", IDiscountMarkupPriceRate5=" & rst("IDiscountMarkupPriceRate5") & " " & _
                           "WHERE ItemID=" & newitemID
            End If
            
            If rst("EBayPrice") <> 0 Then
                LogItemEBayPriceChanged newitem, rst("EBayPrice")
                DB.Execute "UPDATE InventoryMaster " & _
                           "SET EBayPrice=" & rst("EBayPrice") & " " & _
                           "WHERE ItemID=" & newitemID
            End If
            
            If rst("StdCost") <> 0 Then
                LogItemCostChanged newitem, rst("StdCost")
                DatabaseFunctions.ModifyItemCost CLng(newitemID), "Std Cost", rst("StdCost")
                DB.Execute "UPDATE InventoryMaster " & _
                           "SET StdCost=" & rst("StdCost") & " " & _
                           "WHERE ItemID=" & newitemID
            End If
            
            If rst("TNC") <> 0 Then
                LogItemTNCChanged newitem, rst("TNC")
                DatabaseFunctions.ModifyItemCost CLng(newitemID), "New Cost", rst("StdCost")
                DB.Execute "UPDATE InventoryMaster " & _
                           "SET TNC=" & rst("TNC") & " " & _
                           "WHERE ItemID=" & newitemID
            End If
            
            If rst("ListPrice") <> 0 Then
                'no logging
                DatabaseFunctions.ModifyItemCost CLng(newitemID), "List Price", rst("ListPrice")
                DB.Execute "UPDATE InventoryMaster " & _
                           "SET ListPrice=" & rst("ListPrice") & " " & _
                           "WHERE ItemID=" & newitemID
            End If
            
            'skips mapp?
            
            CreateZZZClone = True
        End If
    End If
End Function

Public Function updateDropshipItemWebPrice(item As String, newprice As String) As Boolean
    updateDropshipItemWebPrice = False
    If Left(item, 3) <> "XXX" And Left(item, 3) <> "ZZZ" Then
        If IsItemDSable(item) Then
            Dim zzz As String
            zzz = ZZZifyItemNumber(item)
            If zzz <> "" Then
                updateInventoryMaster "IDiscountMarkupPriceRate1", newprice, zzz, ""
                updateDropshipItemWebPrice = True
            End If
        End If
    End If
End Function

Public Function updateDropshipItemEBayPrice(item As String, newprice As String) As Boolean
    updateDropshipItemEBayPrice = False
    If Left(item, 3) <> "XXX" And Left(item, 3) <> "ZZZ" Then
        If IsItemDSable(item) Then
            Dim zzz As String
            zzz = ZZZifyItemNumber(item)
            If zzz <> "" Then
                updateInventoryMaster "EBayPrice", newprice, zzz, ""
                updateDropshipItemEBayPrice = True
            End If
        End If
    End If
End Function

Public Function IsGroupedSubitem(item As String) As Boolean
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumberPointer FROM PartNumberGroupMaster WHERE ID=(SELECT DISTINCT GroupID FROM PartNumberGroupLines WHERE ItemNumber='" & item & "')")
    If rst.EOF Then
        IsGroupedSubitem = False
    Else
        IsGroupedSubitem = True
    End If
    rst.Close
    Set rst = Nothing
End Function

Public Function IsGroupedMasteritem(item As String) As Boolean
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumberPointer FROM PartNumberGroupMaster WHERE ItemNumberPointer='" & item & "'")
    If rst.EOF Then
        IsGroupedMasteritem = False
    Else
        IsGroupedMasteritem = True
    End If
    rst.Close
    Set rst = Nothing
End Function

Public Function GetGroupedMasteritemForSubitem(item As String) As String
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumberPointer FROM PartNumberGroupMaster WHERE ID=(SELECT DISTINCT GroupID FROM PartNumberGroupLines WHERE ItemNumber='" & item & "')")
    If rst.EOF Then
        GetGroupedMasteritemForSubitem = ""
    Else
        GetGroupedMasteritemForSubitem = rst(0)
    End If
    rst.Close
    Set rst = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : CalculateDropshipChargesFor
' Author    : briandonorfio
' Date      : 1/22/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function CalculateDropshipChargesFor(item As String, addressType As Long, Optional costAmt As Variant = -1, Optional totalWeight As Double = -1, Optional avgFreightAmt As Variant = -1) As Variant
    Dim itemID As String
    itemID = Utilities.GetItemIDByItemCode(item)
    
    Select Case addressType
        Case Is = 0
            'commercial, okay
        Case Is = 1
            'residential, okay
        Case Else
            Err.Raise 123, "CalculateDropshipChargesFor", "bad addressType"
    End Select
    
    If productLineDropshipChargeCache Is Nothing Then
        Set productLineDropshipChargeCache = New Dictionary
    End If
    
    Dim i As Long
    Dim PL As String
    PL = Left(item, 3)
    
    'cache if this is the first time the product line has been hit
    '
    'cache by product line -> "adder"  -> "0" or "1" addressType -> ProductLineDropshipCharges fields
    '                      -> "tiered" -> ProductLineDropshipFlatFees array of variant/arrays
    If Not productLineDropshipChargeCache.exists(PL) Then
        Dim temp1 As Dictionary
        Set temp1 = New Dictionary
        
        temp1.Add "adder", New Dictionary
        temp1.item("adder").Add 0, Nothing 'more address types go here
        temp1.item("adder").Add 1, Nothing
        
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT AddressType, LiftGateCharge, LiftGateWeightThreshold, AdderAmount, AdderType, AdderThreshold, AdderMaximum " & _
                              "FROM ProductLineDropshipCharges " & _
                              "WHERE ProductLineID=(SELECT ID FROM ProductLine WHERE ProductLine='" & EscapeSQuotes(PL) & "') ")
        While Not rst.EOF
            Dim temp2 As Dictionary
            Set temp2 = New Dictionary
            temp2.Add "LiftGateCharge", CStr(Nz(rst("LiftGateCharge"), "<NULL>"))
            temp2.Add "LiftGateWeightThreshold", CStr(Nz(rst("LiftGateWeightThreshold"), "<NULL>"))
            temp2.Add "AdderAmount", CStr(Nz(rst("AdderAmount"), "<NULL>"))
            temp2.Add "AdderType", CLng(rst("AdderType"))
            temp2.Add "AdderThreshold", CStr(Nz(rst("AdderThreshold"), "<NULL>"))
            temp2.Add "AdderMaximum", CStr(Nz(rst("AdderMaximum"), "<NULL>"))
            Set temp1.item("adder").item(CLng(rst("AddressType"))) = temp2
            rst.MoveNext
        Wend
        
        Set rst = DB.retrieve("SELECT ThresholdCost, AmountCalculationType, Amount " & _
                              "FROM ProductLineDropshipFlatFees " & _
                              "WHERE ProductLineID=(SELECT ID FROM ProductLine WHERE ProductLine='" & EscapeSQuotes(PL) & "') " & _
                              "ORDER BY ThresholdCost ASC")
        If rst.EOF Then
            temp1.Add "tiered", Split("", "zerolengtharray")
        Else
            ReDim temp3(0 To rst.RecordCount - 1) As Variant
            i = 0
            While Not rst.EOF
                temp3(i) = Array(CDec(rst("ThresholdCost")), CLng(rst("AmountCalculationType")), CDec(rst("Amount")))
                i = i + 1
                rst.MoveNext
            Wend
            temp1.Add "tiered", temp3
        End If
        
        productLineDropshipChargeCache.Add PL, temp1
    End If
    
    'check for a weight being received in args (optional), fill if not given
    If totalWeight = -1 Then
        totalWeight = CDbl(DLookup("SUM(CAST(Weight as float))", "InventoryComponents", "ID IN (SELECT ComponentID FROM InventoryComponentMap WHERE ItemID=" & itemID & ")"))
    End If
    
    'check for a cost
    If costAmt = -1 Then
        Dim rst2 As ADODB.Recordset
        Set rst2 = DB.retrieve("SELECT TNC, StdCost FROM InventoryMaster WHERE ItemNumber='" & EscapeSQuotes(item) & "'")
        If Nz(rst2("TNC"), "0") <> "0" Then
            costAmt = CDec(rst2("TNC"))
        Else
            costAmt = CDec(rst2("StdCost"))
        End If
        rst2.Close
    End If
    
    'don't check for the average freight amount, only do that later if required
    
    Dim retvalAmt As Variant, retvalMsg As String
    retvalAmt = CDec(0)
    retvalMsg = ""
    
    'now the actual calculations
    'start with adders
    If productLineDropshipChargeCache.item(PL).item("adder").item(addressType) Is Nothing Then
        'no adders defined for this address type
        retvalMsg = retvalMsg & "lift gate charge not defined" & vbCrLf & "percentage adder not defined" & vbCrLf
    Else
        'adders defined, let's check
        'start with lift gate charge
        If productLineDropshipChargeCache.item(PL).item("adder").item(addressType).item("LiftGateCharge") = "<NULL>" Then
            'no lift gate charges
            retvalMsg = retvalMsg & "lift gate charge not defined" & vbCrLf
        Else
            'lift gate charges, maybe
            If productLineDropshipChargeCache.item(PL).item("adder").item(addressType).item("LiftGateWeightThreshold") = "<NULL>" Then
                'no lift gate threshold, so it must apply
                retvalAmt = CDec(CDec(retvalAmt) + CDec(productLineDropshipChargeCache.item(PL).item("adder").item(addressType).item("LiftGateCharge")))
                retvalMsg = retvalMsg & "lift gate weight minimum not defined, so assuming lift gate charge applies" & vbCrLf
                retvalMsg = retvalMsg & "lift gate charge is: " & Format(CDec(productLineDropshipChargeCache.item(PL).item("adder").item(addressType).item("LiftGateCharge")), "Currency") & vbCrLf
            Else
                'check against threshold, see if it applies
                If CDec(totalWeight) < CDec(productLineDropshipChargeCache.item(PL).item("adder").item(addressType).item("LiftGateWeightThreshold")) Then
                    'nope
                Else
                    'yup
                    retvalAmt = CDec(CDec(retvalAmt) + CDec(productLineDropshipChargeCache.item(PL).item("adder").item(addressType).item("LiftGateCharge")))
                    retvalMsg = retvalMsg & "lift gate charge is: " & Format(CDec(productLineDropshipChargeCache.item(PL).item("adder").item(addressType).item("LiftGateCharge")), "Currency") & vbCrLf
                End If
            End If
        End If
        
        'now go to the basic $/% adder
        If productLineDropshipChargeCache.item(PL).item("adder").item(addressType).item("AdderAmount") = "<NULL>" Then
            'not defined
            retvalMsg = retvalMsg & "% adder not defined" & vbCrLf
        Else
            'defined, so see if it applies
            'only applies if threshold is undefined, or cost of item is below threshold
            Dim adderApplies As Boolean
            If productLineDropshipChargeCache.item(PL).item("adder").item(addressType).item("AdderThreshold") = "<NULL>" Then
                retvalMsg = retvalMsg & "adder cost limit not defined, so assuming it applies" & vbCrLf
                adderApplies = True
            ElseIf CDec(costAmt) < CDec(productLineDropshipChargeCache.item(PL).item("adder").item(addressType).item("AdderThreshold")) Then
                adderApplies = True
            Else
                adderApplies = False
            End If
            If adderApplies Then
                Dim adderCost As Variant
                Select Case productLineDropshipChargeCache.item(PL).item("adder").item(addressType).item("AdderType")
                    Case Is = 0
                        'calculate via percentage of cost
                        adderCost = CDec(CDec(costAmt) * CDec(productLineDropshipChargeCache.item(PL).item("adder").item(addressType).item("AdderAmount")) / 100)
                    Case Is = 1
                        'flat charge on top
                        adderCost = CDec(productLineDropshipChargeCache.item(PL).item("adder").item(addressType).item("AdderAmount"))
                End Select
                'now check against the max, if there is one
                If productLineDropshipChargeCache.item(PL).item("adder").item(addressType).item("AdderMaximum") = "<NULL>" Then
                    'no maximum, full adderCost applies
                Else
                    If adderCost > CDec(productLineDropshipChargeCache.item(PL).item("adder").item(addressType).item("AdderMaximum")) Then
                        'over the max, bring it down
                        retvalMsg = retvalMsg & "% adder maximum reached" & vbCrLf
                        adderCost = CDec(productLineDropshipChargeCache.item(PL).item("adder").item(addressType).item("AdderMaximum"))
                    End If
                End If
                'finally, add the adder to the running total
                retvalMsg = retvalMsg & "adder charge is: " & Format(adderCost, "Currency") & vbCrLf
                retvalAmt = CDec(CDec(retvalAmt) + CDec(adderCost))
            End If
        End If
    End If
    'now run the tiered charges
    If UBound(productLineDropshipChargeCache.item(PL).item("tiered")) = -1 Then
        ''no tiered charges, just try our estimated shipping, this is probably dumb
        'If avgFreightAmt = -1 Then
        '    'need to grab from db, not supplied
        '    avgFreightAmt = GetAvgFreightCost(item)
        'End If
        'retvalMsg = retvalMsg & "tiered shipping charge not defined, using avg shipping estimate of " & avgFreightAmt & vbCrLf
        'retvalAmt = CDec(CDec(retvalAmt) + CDec(avgFreightAmt))
        retvalMsg = retvalMsg & "tiered shipping charge not defined"
    Else
        'tiered charges are defined, use the correct tier based on cost
        Dim tieredCost As Variant
        For i = 0 To UBound(productLineDropshipChargeCache.item(PL).item("tiered"))
            If CDec(costAmt) < CDec(productLineDropshipChargeCache.item(PL).item("tiered")(i)(0)) Then
                Select Case CLng(productLineDropshipChargeCache.item(PL).item("tiered")(i)(1))
                    Case 0 'percentage of cost
                        tieredCost = costAmt * (CDec(productLineDropshipChargeCache.item(PL).item("tiered")(i)(2)) / 100)
                    Case 1 'flat charge
                        tieredCost = CDec(productLineDropshipChargeCache.item(PL).item("tiered")(i)(2))
                    Case Else 'error
                        Err.Raise 123, "CalculateDropshipChargesFor", "bad addressType"
                End Select
                Exit For
            End If
        Next i
        retvalAmt = CDec(CDec(retvalAmt) + CDec(tieredCost))
        retvalMsg = retvalMsg & "tiered shipping charge is: " & Format(tieredCost, "Currency") & vbCrLf
    End If
    
    CalculateDropshipChargesFor = Array(CDec(retvalAmt), retvalMsg)
End Function

Public Sub CalculateDropshipChargesClearCacheFor(ProductLine As String)
    If productLineDropshipChargeCache Is Nothing Then
        'skip
    ElseIf productLineDropshipChargeCache.exists(ProductLine) Then
        productLineDropshipChargeCache.Remove ProductLine
    End If
End Sub
