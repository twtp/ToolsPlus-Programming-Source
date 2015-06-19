Attribute VB_Name = "NewProductsImport"
'---------------------------------------------------------------------------------------
' Module    : NewProductsImport
' DateTime  : 6/14/2005 10:34
' Author    : briandonorfio
' Purpose   : Spreadsheet -> NewProdImport, NewProdImport -> Po/Inv, and some temporary
'             additions to mas200 copies.
'
'             Dependencies:
'               - Microsoft ActiveX Data Objects 2.8 Library
'               - Microsoft Excel 9.0 Object Library
'               - DatabaseFunctions + DBConn
'               - ValidationFunctions
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : offload_single_item
' DateTime  : 6/14/2005 11:13
' Author    : briandonorfio
' Purpose   : Main function for the NewProd -> Po/Inv import. Attempts to transfer item
'             information from NewProdXSheetDetail to both the InventoryMaster table
'             and the PartNumbers table. Returns true if successful, false otherwise.
'---------------------------------------------------------------------------------------
'
Public Function offload_single_item(item As String, websiteID As Long) As Boolean
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("exec spNewProdHeaderAndDetailsByItem '" & item & "'")
    If rst.EOF Then
        MsgBox "Could not find item " & item & "? That shouldn't happen."
        offload_single_item = False
        GoTo done
    ElseIf Trim$(rst("VendorNumber")) = "" Then
        MsgBox "Error, need Vendor Number."
        offload_single_item = False
        GoTo done
    End If
    
    If Not ExistsInInventoryMaster(item) Then
        If Not IM_new_item(item, websiteID, rst) Then
            MsgBox "Error creating new item " & item & " in InventoryMaster"
            offload_single_item = False
            GoTo done
        End If
    End If
    
    If Not ExistsInPartNumbers(item) Then
        If Not PN_new_item(item, rst) Then
            MsgBox "Error creating new item " & item & " in PartNumbers"
            offload_single_item = False
            GoTo done
        End If
    End If
    
    If Not updateBoth(item, rst) Then
        MsgBox "Error updating item " & item
        offload_single_item = False
        GoTo done
    End If
    
    Dim itemID As String
    itemID = Utilities.GetItemIDByItemCode(item)
    
    If DLookup("COUNT(*)", "InventoryQuantities", "ItemNumber='" & item & "' AND WhseCode='000'") = "0" Then
        DB.Execute "INSERT INTO InventoryQuantities ( ItemID, ItemNumber, WhseCode ) VALUES ( " & itemID & ", '" & item & "', '000' )"
    End If
    
    'If rst("Weight") <> "0" Or rst("Dimensions") <> "0x0x0" Then
        Dim compCount As Long, cid As String
        compCount = CLng(DLookup("COUNT(*)", "InventoryComponentMap", "ItemID=" & itemID))
        If compCount = 0 Then
            DB.Execute "INSERT INTO InventoryComponents ( Component ) VALUES ( '" & item & "' )"
            cid = DLookup("@@IDENTITY", "InventoryComponents")
            DB.Execute "INSERT INTO InventoryComponentMap ( ItemID, ItemNumber, ComponentID ) VALUES ( " & itemID & ", '" & item & "', " & cid & " )"
        ElseIf compCount = 1 Then
            cid = DLookup("ComponentID", "InventoryComponentMap", "ItemID=" & itemID)
        Else
            cid = "-1"
        End If
        If cid <> "-1" Then
            Dim dims As Variant
            dims = Split(rst("Dimensions"), "x")
            DB.Execute "UPDATE InventoryComponents SET Length='" & dims(0) & "', Width='" & dims(1) & "', Height='" & dims(2) & "', Weight='" & rst("Weight") & "' WHERE ID=" & cid
        End If
    'End If
    
    If rst("UPC") <> "" Then
        'SaveBarCode item, rst("UPC")
        AddNewBarcodeForItem item, itemID, CStr(rst("UPC"))
    End If
    
    'updateExternalShippingDB item
    ExternalComponentDBSync item
    
    If rst("DropShippable") Then
        updateInventoryMaster "ItemStatus", SetItemDSable(item), item, ""
        CreateZZZClone item
    End If

    offload_single_item = True

done:
    rst.Close
    Set rst = Nothing

End Function

'---------------------------------------------------------------------------------------
' Procedure : IM_new_item
' DateTime  : 6/14/2005 11:14
' Author    : briandonorfio
' Purpose   : Creates a new stub item in InventoryMaster. Called by offload_single_item.
'---------------------------------------------------------------------------------------
'
Private Function IM_new_item(item As String, websiteID As Long, rst As ADODB.Recordset) As Boolean
    Dim sqlCol As String, sqlVal As String, strsql As String
    sqlCol = "ItemNumber, WebsiteID, ItemDescription, ProductLine, PrimaryVendor"
    sqlVal = "'" & item & "', " & websiteID & ", '" & EscapeSQuotes(rst("ItemDescription")) & "', '" & Left$(item, 3) & "', '" & EscapeSQuotes(rst("VendorNumber")) & "'"
    strsql = "INSERT INTO InventoryMaster ( " & sqlCol & " ) VALUES ( " & sqlVal & " )"
    If Not DB.Execute(strsql) Then
        MsgBox "Error creating item " & item & " in InventoryMaster. SQL string follows:" & vbCrLf & vbCrLf & strsql
        IM_new_item = False
    Else
    IM_new_item = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : PN_new_item
' DateTime  : 6/14/2005 11:15
' Author    : briandonorfio
' Purpose   : Creates a new stub item in PartNumbers. Called by offload_single_item.
'---------------------------------------------------------------------------------------
'
Private Function PN_new_item(item As String, rst As ADODB.Recordset) As Boolean
    Dim sqlCol As String, sqlVal As String, strsql As String
    sqlCol = "ItemNumber, PrintSign, SignID, WebToBePublished, WebOrderable"
    sqlVal = "'" & item & "', 1, " & SIGN_MASTERSIGN_ID & ", 1, 1"
    strsql = "INSERT INTO PartNumbers ( " & sqlCol & " ) VALUES ( " & sqlVal & " )"
    If Not DB.Execute(strsql) Then
        MsgBox "Error creating item " & item & " in PartNumbers. SQL string follows:" & vbCrLf & vbCrLf & strsql
        PN_new_item = False
    Else
        PN_new_item = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : updateBoth
' DateTime  : 6/14/2005 11:15
' Author    : briandonorfio
' Purpose   : Handles all the actual data movement from table to table. Called by
'             offload_single_item.
'---------------------------------------------------------------------------------------
'
Private Function updateBoth(item As String, rst As ADODB.Recordset) As Boolean
    Dim itemID As String
    itemID = Utilities.GetItemIDByItemCode(item)
    
    Dim strsql As String, cost As Variant

    strsql = "ItemInfoChanged=1, InstorePriceChanged=1, InternetPriceChanged=1, UDFChanged=1, NewItem=1"
    strsql = "UPDATE InventoryMaster SET " & strsql & " WHERE ItemNumber='" & item & "'"
    If Not DB.Execute(strsql) Then
        MsgBox "Error during first update for " & item & " to InventoryMaster. SQL string follows:" & vbCrLf & vbCrLf & strsql
        updateBoth = False
        GoTo done
    End If

    strsql = ""
    
    'If rst("DropShippable") = 1 And rst("DropShipTerms") <> "" Then
    '    strsql = strsql & ", EPN='" & IIf(Nz(rst("EPN")) = "", "", Nz(rst("EPN")) & " ") & EscapeSQuotes(rst("EPN")) & "DSTerms: " & EscapeSQuotes(rst("DropShipTerms")) & "'"
    'ElseIf Len(rst("EPN")) <> 0 Then
    '    strsql = strsql & ", EPN='" & EscapeSQuotes(rst("EPN")) & "'"
    'End If
    
    If rst("PromoCost") <> 0 Then
        If rst("RegularCost") = 0 Then
            cost = CDec(rst("PromoCost"))
        ElseIf rst("PromoCost") <> 0 And rst("PromoCost") < rst("RegularCost") Then
            cost = CDec(rst("PromoCost"))
        Else
            cost = CDec(rst("RegularCost"))
        End If
    ElseIf rst("RegularCost") <> 0 Then
        cost = CDec(rst("RegularCost"))
    Else
        cost = CDec(0)
    End If
    If cost <> 0 Then
        strsql = strsql & ", StdCost=" & CStr(cost)
        LogItemCostChanged item, CStr(cost)
        DatabaseFunctions.ModifyItemCost CLng(itemID), "Std Cost", CStr(cost), 4
    End If
    
    If rst("StorePrice") <> 0 Then
        strsql = strsql & ", StdPrice=" & rst("StorePrice") & ", DiscountMarkupPriceRate1=" & rst("StorePrice")
    End If
    
    If rst("ListPrice") <> 0 Then
        strsql = strsql & ", ListPrice=" & rst("ListPrice")
        DatabaseFunctions.ModifyItemCost CLng(itemID), "List Price", rst("ListPrice"), 4
    End If
    
    If rst("WebPrice") <> 0 Then
        strsql = strsql & ", IDiscountMarkupPriceRate1=" & rst("WebPrice")
    End If
    
    'If Len(rst("ShippingClass")) <> 0 Then
    '    strsql = strsql & ", Class='" & EscapeSQuotes(rst("ShippingClass")) & "'"
    'End If
    '
    'If rst("Weight") <> "0" Then
    '    strsql = strsql & ", Weight='" & EscapeSQuotes(rst("Weight")) & "'"
    '    If Not CBool(InStr(rst("Weight"), ";")) Then
    '        If CDbl(rst("Weight")) > 150 Then
    '            strsql = strsql & ", TruckFreight=1"
    '        End If
    '    End If
    'End If
    '
    'If rst("Dimensions") <> "0x0x0" Then
    '    strsql = strsql & ", Dimensions='" & EscapeSQuotes(rst("Dimensions")) & "'"
    'End If
    
    strsql = strsql & ", StdPack=" & rst("StdPack")
    
    If Len(rst("Components")) <> 0 Then
        strsql = strsql & ", Components='" & EscapeSQuotes(rst("Components")) & "'"
    End If
    
    'strsql = strsql & ", BoxID=" & DetermineBoxSize(rst("Dimensions"), rst("Weight"))
    
    strsql = Mid$(strsql, 3)
    strsql = "UPDATE InventoryMaster SET " & strsql & " WHERE ItemNumber='" & item & "'"
    If Not DB.Execute(strsql) Then
       MsgBox "Error during second update for " & item & " to InventoryMaster. SQL string follows:" & vbCrLf & vbCrLf & strsql
       updateBoth = False
       GoTo done
    End If
    
    'do this after IM gets updated, since it needs the info in there to access it.
    If rst("QtyToOrder") <> 0 Then
        'updateInventoryMaster "QtyOrdered", rst("QtyToOrder"), item
        'updateInventoryMaster "POID", CreatePurchaseOrder(item), item, ""
        MsgBox "po creation broken! shut the fuck up about it!"
    End If
    'updateInventoryMaster "BoxID", DetermineBoxSize(rst("Dimensions"), rst("Weight")), item, ""
    'If rst("DropShippable") Then
    '    updateInventoryMaster "ItemStatus", SetItemDSable(item), item, ""
    '    CreateZZZClone item
    'End If

    
    Dim i As Long
    strsql = "Desc2='" & Mid(item, 4) & "'"
    strsql = strsql & ", Desc3='" & EscapeSQuotes(rst("ItemDescription")) & "'"
    For i = 1 To 8
        If Len(rst("Spec" & i)) <> 0 Then
            strsql = strsql & ", Spec" & i & "HTML='" & EscapeSQuotes(UnFrontpage(rst("Spec" & i))) & "'"
        End If
    Next i
    If Len(rst("ExtendedSpecs")) <> 0 Then
        strsql = strsql & ", ExtendedSpecsHTML='" & EscapeSQuotes(UnFrontpage(rst("ExtendedSpecs"))) & "'"
    End If
    
    Dim availLimit As String
    availLimit = DLookup("AvailLimitDefault", "ProductLine", "ProductLine='" & Left$(item, 3) & "'")
    strsql = strsql & ", AvailabilityLimit=" & availLimit
    
    Dim ebayStoreCategoryID As String
    ebayStoreCategoryID = EBay.EBayGetProbableStoreCategoryIDFor(item)
    If ebayStoreCategoryID <> "" Then
        strsql = strsql & ", EBayStoreCategoryID=" & ebayStoreCategoryID
    End If
    
    If strsql <> "" Then
        strsql = "UPDATE PartNumbers SET " & strsql & " WHERE ItemNumber='" & item & "'"
        If Not DB.Execute(strsql) Then
            MsgBox "Error during second update for " & item & " to PartNumbers. SQL string follows:" & vbCrLf & vbCrLf & strsql
            updateBoth = False
            GoTo done
        End If
    End If
    
    updateBoth = True

done:

End Function

'---------------------------------------------------------------------------------------
' Procedure : parseXSheetHeader
' DateTime  : 6/14/2005 11:16
' Author    : briandonorfio
' Purpose   : Parses the header of the new products spreadsheet. Returns the vendor
'             number, to be used later on when parsing the detail section. Inserts all
'             information into the NewProdXSheetHeader table. Formatting information for
'             the spreadsheet is in NewProdXSheetLayout. Any changes to the spreadsheet
'             must be duplicated there.
'---------------------------------------------------------------------------------------
'
Public Function parseXSheetHeader(wks As Excel.Worksheet) As String
    'get vendornumber and vendorname
    Dim VendorNumber As String
    Dim rstHeader As ADODB.Recordset
    Set rstHeader = DB.retrieve("SELECT Row, Col FROM NewProdXSheetLayout WHERE FieldName='VendorNumber'")
    VendorNumber = EscapeSQuotes(wks.Range(rstHeader("Col") & rstHeader("Row")).value)
    Dim rstVendor As ADODB.Recordset
    Set rstVendor = DB.retrieve("SELECT DISTINCT VendorNo, VendorName FROM AP_Vendor WHERE VendorNo='" & VendorNumber & "'")
    If rstVendor.EOF Then
        DB.Execute ("INSERT INTO NewProdXSheetHeader ( VendorNumber ) VALUES ( '" & VendorNumber & "' )")
        rstHeader.Close
        Set rstHeader = DB.retrieve("SELECT Row, Col FROM NewProdXSheetLayout WHERE FieldName='VendorName'")
        If wks.Range(rstHeader("Col") & rstHeader("Row")).value <> "" Then
            DB.Execute ("UPDATE NewProdXSheetHeader SET VendorName='" & EscapeSQuotes(wks.Range(rstHeader("Col") & rstHeader("Row")).value) & "' WHERE VendorNumber='" & VendorNumber & "'")
        End If
    Else
        DB.Execute ("DELETE FROM NewProdXSheetHeader WHERE VendorNumber='" & VendorNumber & "'")
        DB.Execute ("INSERT INTO NewProdXSheetHeader ( VendorNumber, VendorName ) VALUES ( '" & VendorNumber & "', '" & rstVendor("VendorName") & "' )")
    End If
    rstVendor.Close
    Set rstVendor = Nothing
    rstHeader.Close
    
    
    'get rest of header information
    Set rstHeader = DB.retrieve("SELECT Row, Col, FieldName FROM NewProdXSheetLayout WHERE Row<18 AND FieldName<>'VendorName' AND FieldName<>'VendorNumber' ORDER BY Row, Col")
    Dim txt As String, strsql As String
    While Not rstHeader.EOF
        txt = EscapeSQuotes(wks.Range(rstHeader("Col") & rstHeader("Row")).value)
        If Len(txt) <> 0 Then
            strsql = strsql & ", " & rstHeader("FieldName") & "='" & txt & "'"
        End If
        rstHeader.MoveNext
    Wend
    If strsql <> "" Then
        strsql = Mid$(strsql, 3)
        DB.Execute ("UPDATE NewProdXSheetHeader SET " & strsql & " WHERE VendorNumber='" & VendorNumber & "'")
    End If
    
    rstHeader.Close
    Set rstHeader = Nothing
    
    parseXSheetHeader = VendorNumber
End Function

'---------------------------------------------------------------------------------------
' Procedure : parseXSheetItem
' DateTime  : 6/14/2005 11:18
' Author    : briandonorfio
' Purpose   : Parses the detail lines of the new products spreadsheet. Arguments are the
'             Excel worksheet object, a recordset containing the Column->Name lookups
'             from NewProdXSheetLayout, line number to parse, and the vendor number
'             of the item, from the header portion of the spreadsheet. Returns true if
'             successfully parses and inserts into NewProdXSheetDetail, otherwise false.
'---------------------------------------------------------------------------------------
'
Public Function parseXSheetItem(wks As Excel.Worksheet, rstItem As ADODB.Recordset, line As Long, VendorNumber As String) As Boolean

    Dim sqlCol As String, sqlVal As String, txt As String
    Dim lcode_pnum As String
    Dim delim As String
    Dim shipweight As Double
    'assumes line code and part # are 1 and 2, if null, end of spreadsheet
    If IsNull(wks.Range(rstItem(1) & line).value) Or wks.Range(rstItem(1) & line).value = "" Then
        GoTo done
    End If
    lcode_pnum = UCase$(EscapeSQuotes(wks.Range("A" & line).value & wks.Range("B" & line).value))
    If Not validateItemNumber(lcode_pnum) Then
        MsgBox ("Error on line " & line & " of spreadsheet. Line Code or Part Number is invalid (between 4 and 27 chars, only alphanumeric or -).")
        GoTo done
    End If
    sqlCol = "ItemNumber"
    sqlVal = "'" & lcode_pnum & "'"
    
    Dim firstTimeThrough As Boolean
    firstTimeThrough = True
    
    While Not rstItem.EOF
        shipweight = -1
        txt = ""
        delim = ""
        Select Case rstItem("FieldName")
            
            Case Is = "UPC"
                txt = wks.Range(rstItem("Col") & line).value
                txt = Replace(txt, "-", "")
                txt = Replace(txt, " ", "")
                txt = Replace(txt, vbCr, "")
                txt = Replace(txt, vbLf, "")
                If txt = "" Or txt = "n/a" Then
                    'ok
                ElseIf Not validateUPC(txt) Then
                    MsgBox ("Error on line " & line & " of spreadsheet. UPC is not numeric.")
                    GoTo done
                End If
                delim = "'"
                
            Case Is = "ItemDescription"
                txt = UCase$(wks.Range(rstItem("Col") & line).value)
                txt = Replace(txt, """", "''") 'replace d-q w/ 2 s-q's
                txt = Replace(txt, ",", "")    'remove commas
                If Not validateItemDescription(txt) Then
                    MsgBox ("Error on Line " & line & " of spreadsheet. Item Description is too long.")
                    GoTo done
                End If
                delim = "'"
                
            Case Is = "EPN"
                txt = UCase$(wks.Range(rstItem("Col") & line).value)
                If Not validateEPN(txt) Then
                    MsgBox ("Error on line " & line & " of spreadsheet. Item Comments (EPN) is too long.")
                    GoTo done
                End If
                delim = "'"
                
            Case Is = "KitIncludes"
                'TODO: how is this supposed to work? comma-delim?
                If firstTimeThrough Then
                    MsgBox "INCLUDES: not implemented yet!"
                End If
                txt = ""
            
            Case Is = "Dimensions"
                txt = UCase$(StripDQuotes(wks.Range(rstItem("Col") & line).value))
                If txt = "" Then
                    txt = "0x0x0"
                End If
                txt = validateDimensions(txt)
                If Left(txt, 5) = "ERROR" Then
                    MsgBox ("Error on line " & line & " of spreadsheet. Can't interpret dimensions.")
                    GoTo done
                End If
                delim = "'"
            
            Case Is = "ShippingClass"
                txt = wks.Range(rstItem("Col") & line).value
                If txt = "" Then
                    txt = "UPS"
                End If
                If Not validateShippingClass(txt) Then
                    If shipweight < 0 Then
                        txt = ""
                    ElseIf shipweight < 150 Then
                        txt = "UPS"
                    Else
                        txt = "85"
                    End If
                End If
                delim = "'"
                    
            Case Is = "Weight"
                txt = wks.Range(rstItem("Col") & line).value
                If txt = "" Then
                    txt = "0"
                End If
                txt = validateWeight(txt)
                If txt = "" Then
                    MsgBox ("Error on line " & line & " of spreadsheet. Can't interpret weight.")
                    GoTo done
                End If
                shipweight = CDbl(txt)
                delim = "'"
                    
            Case Is = "DropShippable"
                txt = UCase$(wks.Range(rstItem("Col") & line).value)
                If txt = "" Then
                    txt = "N"
                End If
                txt = SQLBool(Left$(txt, 1) = "Y")
                delim = ""
                
            Case Is = "DropShipTerms"
                txt = wks.Range(rstItem("Col") & line).value
                If Not validateDropShipTerms(txt) Then
                    MsgBox ("Error on line " & line & " of spreadsheet. Drop-Ship Terms is too long.")
                    GoTo done
                End If
                delim = "'"
                    
            Case Is = "StdPack"
                txt = wks.Range(rstItem("Col") & line).value
                If Not validateGeneralInteger(txt) Then
                    MsgBox ("Error on line " & line & " of spreadsheet. Standard Pack is not an integer.")
                    GoTo done
                End If
                delim = ""
                    
            Case Is = "QtyToOrder"
                txt = wks.Range(rstItem("Col") & line).value
                If txt = "" Then
                    txt = "0"
                End If
                If Not validateGeneralInteger(txt) Then
                    MsgBox ("Error on line " & line & " of spreadsheet. Quantity To Order is not an integer.")
                    GoTo done
                End If
                delim = ""
                
            Case Is = "MAPP"
                'TODO
                If firstTimeThrough Then
                    MsgBox "MAPP: not implemented yet!"
                End If
                txt = ""
                    
            Case Else
                
                If InStr(rstItem("FieldName"), "Spec") Then
                    txt = wks.Range(rstItem("Col") & line).value
                    If Not CBool(InStr(rstItem("FieldName"), "Extended")) Then
                        If Not validateSpecs(txt) Then
                            MsgBox ("Error on line " & line & " of spreadsheet. Spec in column " & rstItem("Col") & " is too long.")
                            GoTo done
                        End If
                    End If
                    delim = "'"
                    If StrComp(UCase(txt), txt, vbBinaryCompare) = 0 Then 'all caps -> de-cap
                        txt = UCase(Left(txt, 1)) & LCase(Mid(txt, 2))
                    End If
                    
                ElseIf InStr(rstItem("FieldName"), "Cost") Or InStr(rstItem("FieldName"), "Price") Then
                    txt = wks.Range(rstItem("Col") & line).value
                    If txt = "" Then
                        txt = "0"
                    End If
                    If Not validateCurrency(txt) Then
                        MsgBox ("Error on Line " & line & " of spreadsheet. Cost or Price in column " & rstItem("Col") & " is invalid.")
                        GoTo done
                    End If
                    txt = Format(txt, "Currency")
                    txt = Replace(txt, ",", "")
                    delim = ""
                    
                End If

        End Select

        txt = Trim$(txt)
        txt = EscapeSQuotes(txt)
        If Len(txt) <> 0 Then
            sqlCol = sqlCol & ", " & rstItem("FieldName")
            sqlVal = sqlVal & ", " & delim & txt & delim
        End If

        rstItem.MoveNext
        
        firstTimeThrough = False

    Wend
    
    sqlCol = sqlCol & ", VendorNumber"
    sqlVal = sqlVal & ", '" & VendorNumber & "'"

    DB.Execute "DELETE FROM NewProdXSheetDetail WHERE ItemNumber='" & lcode_pnum & "'"
    DB.Execute "INSERT INTO NewProdXSheetDetail ( " & sqlCol & " ) VALUES ( " & sqlVal & " )"
    
    parseXSheetItem = True
    Exit Function
    
done:
    parseXSheetItem = False

End Function



Public Sub OpenNewProductsSpreadsheet()
    Dim dlg As FilePicker
    Set dlg = New FilePicker
    'dlg.SetParent Me.hwnd
    dlg.AddFilter "Excel Documents", "*.xls"
    Dim fname As String
    fname = dlg.ShowDialogOpen
    Set dlg = Nothing
    If fname <> "" Then
        Dim excelapp As Excel.Application
        Set excelapp = New Excel.Application
        Dim wks As Excel.Worksheet
        Set wks = excelapp.Workbooks.Open(fname).Worksheets(1)
        excelapp.Visible = True
        Set wks = Nothing
        Set excelapp = Nothing
    End If
End Sub

Public Sub ImportNewProductsSpreadsheet()
    Dim dlg As FilePicker
    Set dlg = New FilePicker
    'dlg.SetParent Me.hwnd
    dlg.AddFilter "Excel Documents", "*.xls"
    Dim fname As String
    fname = dlg.ShowDialogOpen
    Set dlg = Nothing
    If fname = "" Then
        Exit Sub
    End If
    
    Dim origStatus As String
    origStatus = Menu.lblStatusBar.Caption
    
    Mouse.Hourglass True
    Menu.lblStatusBar.Caption = "Initializing..."
    Dim app As Excel.Application
    Set app = New Excel.Application
    
    Dim wks As Excel.Worksheet
    Set wks = app.Workbooks.Open(fname).Worksheets(1)

    Menu.lblStatusBar.Caption = "Parsing Header..."
    Dim VendorNumber As String
    VendorNumber = parseXSheetHeader(wks)

    Dim rstItem As ADODB.Recordset
    Set rstItem = DB.retrieve("SELECT Row, Col, FieldName FROM NewProdXSheetLayout WHERE Row=18 ORDER BY Col")
    Dim X As Long, goodData As Boolean
    goodData = True
    For X = 18 To 32000
        If IsNull(wks.Range("B" & X).value) Or wks.Range("B" & X).value = "" Then
            Exit For
        Else
            If Len(wks.Range("B" & X).value) > 24 Then
                MsgBox "Whoa, item on line " & X & " is longer than 27 chars and won't go over to Mas! You should fix this."
                goodData = False
                'keep going, alert about any overage
            End If
        End If
    Next X
    If goodData Then
        For X = 18 To 32000
            rstItem.MoveFirst
            Menu.lblStatusBar.Caption = "Processing item " & X - 17
            DoEvents
            If Not parseXSheetItem(wks, rstItem, X, VendorNumber) Then
                X = 32000
            End If
        Next X
    End If

    rstItem.Close
    Set rstItem = Nothing
    Menu.lblStatusBar.Caption = origStatus
    Set wks = Nothing
    app.Quit
    Set app = Nothing
    Mouse.Hourglass False
End Sub
