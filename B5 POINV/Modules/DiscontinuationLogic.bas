Attribute VB_Name = "DiscontinuationLogic"
'---------------------------------------------------------------------------------------
' Module    : DiscontinuationLogic
' DateTime  : 3/9/2006 12:02
' Author    : briandonorfio
' Purpose   : Business logic specifically about how items are discontinued, attributes
'             for discontinued items, etc. Hopefully, all the important stuff, like
'             whether to d/c on sales order or on invoice, etc., is handled here, and
'             not anywhere else in the program.
'
'             Dependencies:
'               - Microsoft ActiveX Data Objects 2.8 Library
'               - DatabaseFunctions + DBConn
'               - MouseHourglass + global Mouse object
'               - PerlArray
'               - Utilities
'---------------------------------------------------------------------------------------

Option Explicit

Public Function CountBeingDCItems() As Long
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT COUNT(*) FROM vCheckForBeingDCItemsStore UNION SELECT COUNT(*) FROM vCheckForBeingDCItemsWeb")
    Dim retval As Long
    retval = rst(0)
    rst.MoveNext
    If Not rst.EOF Then
        retval = retval + rst(0)
    End If
    rst.Close
    Set rst = Nothing
    CountBeingDCItems = retval
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : CheckForBeingDCItems
' DateTime  : 6/14/2005 11:03
' Author    : briandonorfio
' Purpose   : Checks for items marked being discontinued that need to be fully
'             discontinued. If Qty < AvailabilityLimit then it gets discontinued at the
'             web site, if < 0 then it gets discontinued at the store as well. Prompts
'             before doing anything.
'---------------------------------------------------------------------------------------
'
Public Sub CheckForBeingDCItems()
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Dim storeA As PerlArray, webA As PerlArray, dsA As PerlArray, baseA As PerlArray, kitH As Dictionary, reportHash As Dictionary
    Set storeA = New PerlArray
    Set webA = New PerlArray
    Set dsA = New PerlArray
    Set baseA = New PerlArray
    Set kitH = New Dictionary
    Set reportHash = New Dictionary
    reportHash.Add "ds-discontinue", New PerlArray
    reportHash.Add "ds-dropship", New PerlArray
    reportHash.Add "ds-skip", New PerlArray
    reportHash.Add "storeweb", New PerlArray
    reportHash.Add "web", New PerlArray
    reportHash.Add "kit", New PerlArray
    Dim i As Long
    Set rst = DB.retrieve("SELECT ItemNumber, DSable FROM vCheckForBeingDCItemsStore ORDER BY ItemNumber")
    While Not rst.EOF
        If rst("DSable") Then
            dsA.Push rst("ItemNumber").value
        Else
            storeA.Push rst("ItemNumber").value
        End If
        rst.MoveNext
    Wend
    rst.Close
    If dsA.Scalar <> 0 Then
        'If MsgBox("Discontinue the following items in the STORE (but will remain DROPSHIPPABLE at the WEB)?" & vbCrLf & vbCrLf & dsA.Join(vbCrLf), vbYesNo) = vbYes Then
        '    For i = 0 To dsA.Scalar - 1
        '        HandleDiscontinuation dsA.Elem(i)
        '    Next i
        'End If
        For i = 0 To dsA.Scalar - 1
            Load DiscontinueDSableDlg
            DiscontinueDSableDlg.ItemNumber = dsA.Elem(i)
            Mouse.TempClear True
            DiscontinueDSableDlg.Show MODAL
            reportHash.item("ds-" & DiscontinueDSableDlg.ReturnValue).Push dsA.Elem(i)
            Mouse.TempClear False
            'logic+unload is handled in the buttons, we're done
        Next i
    End If
    If storeA.Scalar <> 0 Then
        If MsgBox("Discontinue the following items in the STORE and at the WEBSITE?" & vbCrLf & vbCrLf & storeA.Join(vbCrLf), vbYesNo) = vbYes Then
            For i = 0 To storeA.Scalar - 1
                'SetDiscontinued storeA.Elem(i)
                HandleDiscontinuation storeA.Elem(i)
                reportHash.item("storeweb").Push storeA.Elem(i)
            Next i
        End If
    End If
    Set rst = DB.retrieve("SELECT ItemNumber FROM vCheckForBeingDCItemsWeb ORDER BY ItemNumber")
    While Not rst.EOF
        webA.Push rst("ItemNumber").value
        rst.MoveNext
    Wend
    rst.Close
    If webA.Scalar <> 0 Then
        If MsgBox("Discontinue the following items at the WEBSITE?" & vbCrLf & vbCrLf & webA.Join(vbCrLf), vbYesNo) = vbYes Then
            For i = 0 To webA.Scalar - 1
                'SetWebDiscontinued webA.Elem(i)
                HandleDiscontinuation webA.Elem(i)
                reportHash.item("web").Push webA.Elem(i)
            Next i
        End If
    End If
    'Set rst = DB.retrieve("SELECT ItemNumber FROM vCheckForBeingDCItemsBase ORDER BY ItemNumber")
    'While Not rst.EOF
    '    baseA.Push rst("ItemNumber").value
    '    rst.MoveNext
    'Wend
    'rst.Close
    'If baseA.Scalar <> 0 Then
    '    MsgBox "The following base items need to be discontinued, but I don't know how to do that:" & vbCrLf & vbCrLf & baseA.Join(vbCrLf)
    'End If
    'Set rst = Nothing
    
    Set rst = DB.retrieve("SELECT SalesKitNo, ComponentItemCode FROM vCheckForBeingDCItemsKits ORDER BY SalesKitNo")
    While Not rst.EOF
        If Not kitH.exists(CStr(rst("SalesKitNo"))) Then
            kitH.Add CStr(rst("SalesKitNo")), New PerlArray
        End If
        kitH.item(CStr(rst("SalesKitNo"))).Push CStr(rst("ComponentItemCode"))
        rst.MoveNext
    Wend
    rst.Close
    Dim kitReportString As String
    kitReportString = ""
    If kitH.count <> 0 Then
        Dim temp As PerlArray
        Set temp = New PerlArray
        Dim iter As Variant
        For Each iter In kitH.keys
            temp.Push CStr(iter) & " (" & kitH.item(CStr(iter)).Join(", ") & ")"
        Next iter
        kitReportString = temp.Join(vbCrLf)
        If MsgBox("Discontinue the following kits?" & vbCrLf & vbCrLf & kitReportString, vbYesNo) = vbYes Then
            For Each iter In kitH.keys
                updateInventoryMaster "ItemStatus", ItemStatusBitFlags.SetItemDCbyUs(CStr(iter)), CStr(iter), ""
                HandleDiscontinuation CStr(iter)
                reportHash.item("kit").Push CStr(iter)
            Next iter
        End If
    End If
    
    Set rst = Nothing
    
    Dim reportBody As String
    reportBody = ""
    If reportHash.item("ds-discontinue").Scalar <> 0 Then
        reportBody = reportBody & "Dropship items discontinued completely:" & vbCrLf & vbCrLf
        For i = 0 To reportHash.item("ds-discontinue").Scalar - 1
            reportBody = reportBody & reportHash.item("ds-discontinue").Elem(i) & vbCrLf
        Next i
        reportBody = reportBody & vbCrLf & vbCrLf
    End If
    If reportHash.item("ds-dropship").Scalar <> 0 Then
        reportBody = reportBody & "Dropship items discontinued stocking, but kept dropshippable:" & vbCrLf & vbCrLf
        For i = 0 To reportHash.item("ds-dropship").Scalar - 1
            reportBody = reportBody & reportHash.item("ds-dropship").Elem(i) & vbCrLf
        Next i
        reportBody = reportBody & vbCrLf & vbCrLf
    End If
    If reportHash.item("storeweb").Scalar <> 0 Then
        reportBody = reportBody & "Discontinued store and web:" & vbCrLf & vbCrLf
        For i = 0 To reportHash.item("storeweb").Scalar - 1
            reportBody = reportBody & reportHash.item("storeweb").Elem(i) & vbCrLf
        Next i
        reportBody = reportBody & vbCrLf & vbCrLf
    End If
    If reportHash.item("web").Scalar <> 0 Then
        reportBody = reportBody & "Discontinued on the web (avail limit > 0):" & vbCrLf & vbCrLf
        For i = 0 To reportHash.item("web").Scalar - 1
            reportBody = reportBody & reportHash.item("web").Elem(i) & vbCrLf
        Next i
        reportBody = reportBody & vbCrLf & vbCrLf
    End If
    If reportHash.item("kit").Scalar <> 0 Then
        reportBody = reportBody & "Kits discontinued because components were discontinued:" & vbCrLf & vbCrLf
        reportBody = reportBody & kitReportString & vbCrLf
        reportBody = reportBody & vbCrLf & vbCrLf
    End If
    
    If reportBody <> "" Then
        SendEmailTo "kirk@tools-plus.com", "discontinuation report", reportBody
    End If
    
    Mouse.Hourglass False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HandleDiscontinuation
' DateTime  : 11/21/2005 14:28
' Author    : briandonorfio
' Purpose   : Handles d/c and being d/c for an item, using the new itemstatus field.
'             Assumes the new status change was already handled (i.e., user clicked the
'             being d/c button, change to itemstatus field, then this function goes from
'             there.) If on a PO, then being d/c (for one-shot items, mostly), if qty is
'             <= the availability limit, it's set d/c on the web (price change, web
'             orderable), if = 0 then it's d/c at the store as well (price change,
'             unsetting the stock flag). if < 0 then it just does the web discontinued,
'             store waits until inventory is corrected.
'---------------------------------------------------------------------------------------
'
Public Sub HandleDiscontinuation(item As String)
    Mouse.Hourglass True
    Dim status As String
    status = DLookup("ItemStatus", "InventoryMaster", "ItemNumber='" & item & "'")
    If IsItemDC(status) Then
        Dim QtyOnHand As Long, QtyOnSO As Long, QtyOnPO As Long, availLimit As Long
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT QtyOnHand, QtyOnSO, QtyOnPO FROM vActualWhseQty WHERE ItemNumber='" & item & "'")
        QtyOnHand = rst("QtyOnHand")
        QtyOnSO = rst("QtyOnSO")
        QtyOnPO = rst("QtyOnPO")
        rst.Close
        Set rst = Nothing
        availLimit = DLookup("AvailabilityLimit", "PartNumbers", "ItemNumber='" & item & "'")
        If QtyOnPO > 0 Then
            SetBeingDiscontinued item
        ElseIf QtyOnHand = 0 Then
            If Not IsItemDSable(status) Then
                SetDiscontinued item
                updatePartNumbers "AvailabilityOverride", "Null", item, ""
            Else
                SetStoreDiscontinued item
            End If
            UpdateNotificationEmail "store disc", item
        'ElseIf QtyOnHand - QtyOnSO <= availLimit Then
        ElseIf QtyOnHand <= availLimit Then
            If Not IsItemDSable(status) Then
                SetWebDiscontinued item
                updatePartNumbers "AvailabilityOverride", "Null", item, ""
            End If
        Else
            SetBeingDiscontinued item
        End If
    End If
    Mouse.Hourglass False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetBeingDiscontinued
' DateTime  : 6/14/2005 11:05
' Author    : briandonorfio
' Purpose   : Sets an item to being discontinued.
'---------------------------------------------------------------------------------------
'
Public Sub SetBeingDiscontinued(item As String)
    Mouse.Hourglass True
    LogItemBeingDiscontinued item
    'dbExecute ("UPDATE InventoryMaster SET DiscontinuedFlag=1, Dropshippable=0 WHERE ItemNumber='" & item & "'")
    'dbExecute ("UPDATE PartNumbers SET Availability=" & AVAIL_BEING_DC_ID & " WHERE ItemNumber='" & item & "'")
    Mouse.Hourglass False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetWebDiscontinued
' DateTime  : 6/14/2005 11:05
' Author    : briandonorfio
' Purpose   : Marks an item as discontinued on the web site.
'---------------------------------------------------------------------------------------
'
Public Sub SetWebDiscontinued(item As String)
    Mouse.Hourglass True
    LogItemWebPriceChanged item, "0.00"
    LogItemDiscontinuedAtWeb item
    'dbExecute ("UPDATE PartNumbers SET Availability=" & AVAIL_DISCONTINUED_ID & ", WebOrderable=0, WebToBePublished=0 WHERE ItemNumber='" & item & "'")
    'dbExecute ("UPDATE InventoryMaster SET DiscontinuedFlag=1, Dropshippable=0, IBreakQty1=999999, IBreakQty2=0, IBreakQty3=0, IBreakQty4=0, IBreakQty5=0, IDiscountMarkupPriceRate1=0, IDiscountMarkupPriceRate2=0, IDiscountMarkupPriceRate3=0, IDiscountMarkupPriceRate4=0, IDiscountMarkupPriceRate5=0 WHERE ItemNumber='" & item & "'")
    DB.Execute "UPDATE PartNumbers SET WebOrderable=0, ShowRequestForm=0, ShowDCSpecs=1, WebToBePublished=0 WHERE ItemNumber='" & item & "'"
    DB.Execute "UPDATE InventoryMaster SET IBreakQty1=999999, IBreakQty2=0, IBreakQty3=0, IBreakQty4=0, IBreakQty5=0, IDiscountMarkupPriceRate1=0, IDiscountMarkupPriceRate2=0, IDiscountMarkupPriceRate3=0, IDiscountMarkupPriceRate4=0, IDiscountMarkupPriceRate5=0 WHERE ItemNumber='" & item & "'"
    Dim repl As String
    If repl <> "" Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT WebToBePublished, TBPubPriority FROM PartNumbers WHERE ItemNumber='" & repl & "'")
        If rst.EOF Then
            MsgBox "Replacement itemnumber " & qq(repl) & " does not exist for item " & qq(item) & "!"
        Else
            If rst("WebToBePublished") = 1 And rst("TBPubPriority") <> 0 Then
                DB.Execute "UPDATE PartNumbers SET TBPubPriority=0 WHERE ItemNumber='" & repl & "'"
            End If
        End If
        rst.Close
        Set rst = Nothing
    End If
    Mouse.Hourglass False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetStoreDiscontinued
' DateTime  : 3/14/2006 12:13
' Author    : briandonorfio
' Purpose   : Marks an item as discontinued in the store only. Probably only used for
'             dropships that we don't want to stock anymore, otherwise just use the
'             SetDiscontinued function.
'---------------------------------------------------------------------------------------
'
Public Sub SetStoreDiscontinued(item As String)
    Mouse.Hourglass True
    LogItemStorePriceChanged item, "88888.88"
    DB.Execute "UPDATE PartNumbers SET PrintSign=0 WHERE ItemNumber='" & item & "'"
    DB.Execute "UPDATE InventoryMaster SET ItemStatus=" & UnSetItemStocked(item) & ", StdPrice=88888.88, BreakQty1=999999, BreakQty2=0, BreakQty3=0, BreakQty4=0, BreakQty5=0, DiscountMarkupPriceRate1=88888.88, DiscountMarkupPriceRate2=0, DiscountMarkupPriceRate3=0, DiscountMarkupPriceRate4=0, DiscountMarkupPriceRate5=0 WHERE ItemNumber='" & item & "'"
    Mouse.Hourglass False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetDiscontinued
' DateTime  : 6/14/2005 11:05
' Author    : briandonorfio
' Purpose   : Marks an item as discontinued in the store, and at the web site.
'---------------------------------------------------------------------------------------
'
Public Sub SetDiscontinued(item As String)
    Mouse.Hourglass True
    LogItemStorePriceChanged item, "88888.88"
    LogItemWebPriceChanged item, "0.00"
    LogItemDiscontinuedInStoreAndWeb item
    DB.Execute "UPDATE PartNumbers SET PrintSign=0, WebOrderable=0, ShowRequestForm=0, ShowDCSpecs=1, WebToBePublished=0 WHERE ItemNumber='" & item & "'"
    DB.Execute "UPDATE InventoryMaster SET ItemStatus=" & UnSetItemStocked(item) & ", StdPrice=88888.88, BreakQty1=999999, BreakQty2=0, BreakQty3=0, BreakQty4=0, BreakQty5=0, DiscountMarkupPriceRate1=88888.88, DiscountMarkupPriceRate2=0, DiscountMarkupPriceRate3=0, DiscountMarkupPriceRate4=0, DiscountMarkupPriceRate5=0, IBreakQty1=999999, IBreakQty2=0, IBreakQty3=0, IBreakQty4=0, IBreakQty5=0, IDiscountMarkupPriceRate1=0, IDiscountMarkupPriceRate2=0, IDiscountMarkupPriceRate3=0, IDiscountMarkupPriceRate4=0, IDiscountMarkupPriceRate5=0 WHERE ItemNumber='" & item & "'"
    Dim repl As String
    repl = HasReplacementPart(item)
    If repl <> "" Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT WebToBePublished, TBPubPriority FROM PartNumbers WHERE ItemNumber='" & repl & "'")
        If rst.EOF Then
            MsgBox "Replacement itemnumber " & qq(repl) & " does not exist for item " & qq(item) & "!"
        Else
            If rst("WebToBePublished") And rst("TBPubPriority") <> 0 Then
                DB.Execute "UPDATE PartNumbers SET TBPubPriority=0 WHERE ItemNumber='" & repl & "'"
            End If
        End If
        rst.Close
        Set rst = Nothing
    End If
    Dim qoso As Long
    qoso = DLookup("QtyOnSO", "vActualWhseQty", "ItemNumber='" & item & "'")
    If qoso <> 0 Then
        SendEmailTo ORDER_PROCESSING_EMAIL, item & " discontinued with active sales orders", "Warning: item " & item & " has been discontinued with " & qoso & " on sales order!"
    End If
    Mouse.Hourglass False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetUnDiscontinued
' DateTime  : 6/14/2005 11:05
' Author    : briandonorfio
' Purpose   : Tries to un-discontinue an item. Re-marked as t/b published in signs, so
'             everything can be checked out to make sure it's right.
'---------------------------------------------------------------------------------------
'
Public Sub SetUnDiscontinued(item As String)
    Mouse.Hourglass True
    LogItemUnDiscontinued item
    'dbExecute ("UPDATE InventoryMaster SET DiscontinuedFlag=0 WHERE ItemNumber='" & item & "'")
    'dbExecute ("UPDATE PartNumbers SET PrintSign=1, WebToBePublished=1, WebOrderable=1, Availability=" & AVAIL_NONE_ID & " WHERE ItemNumber='" & item & "'")
    DB.Execute "UPDATE PartNumbers SET PrintSign=1, WebToBePublished=1, WebOrderable=1 WHERE ItemNumber='" & item & "'"
    Mouse.Hourglass False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetDropShippable
' DateTime  : 7/8/2005 15:48
' Author    : briandonorfio
' Purpose   : Flags an item as able to be dropshipped. No checking for already d/c or
'             anything.
'---------------------------------------------------------------------------------------
'
Public Sub SetDropShippable(item As String)
    Mouse.Hourglass True
    'updateInventoryMaster "DropShippable", 1, item, ""
    'updatePartNumbers "Availability", AVAIL_DROPSHIP_ID, item, ""
    updateInventoryMaster "ItemStatus", SetItemDSable(item), item, ""
    If CreateZZZClone(item) Then
        If IsFormLoaded("InventoryMaintenance") Then
            If InventoryMaintenance.ItemNumber = item Then
                InventoryMaintenance.RefreshForm
            End If
        End If
    End If
    Mouse.Hourglass False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetNotDropShippable
' DateTime  : 7/8/2005 15:48
' Author    : briandonorfio
' Purpose   : Unsets the dropship status flag, nothing else.
'---------------------------------------------------------------------------------------
'
Public Sub SetNotDropShippable(item As String)
    Mouse.Hourglass True
'    updateInventoryMaster "DropShippable", 0, item, ""
'    updatePartNumbers "Availability", AVAIL_NONE_ID, item, ""
    updateInventoryMaster "ItemStatus", UnSetItemDSable(item), item, ""
    Mouse.Hourglass False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetStoreOnly
' DateTime  : 7/8/2005 16:01
' Author    : briandonorfio
' Purpose   : Sets an item to be not for sale on the web. IOOS so the backorder form
'             doesn't show up (should also be taken care of on weboff, but double
'             checking it here), weborderable false.
'---------------------------------------------------------------------------------------
'
'Public Sub SetStoreOnly(item As String)
'    Mouse.Hourglass True
'    DB.Execute "UPDATE PartNumbers SET IgnoreOutOfStock=1, WebOrderable=0 WHERE ItemNumber='" & item & "'"
'    Mouse.Hourglass False
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetNotStoreOnly
' DateTime  : 7/8/2005 16:02
' Author    : briandonorfio
' Purpose   : Unsets an item so that it is available on the web.
'---------------------------------------------------------------------------------------
'
'Public Sub SetNotStoreOnly(item As String)
'    Mouse.Hourglass True
'    DB.Execute "UPDATE PartNumbers SET IgnoreOutOfStock=0, WebOrderable=1 WHERE ItemNumber='" & item & "'"
'    Mouse.Hourglass False
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : ShouldBeWebOrderable
' DateTime  : 3/6/2006 12:35
' Author    : briandonorfio
' Purpose   : Returns whether the item should be web orderable or not, according to the
'             current values in the database. Looks at the discontinued status based on
'             availability limit, and any availability overrides it has.
'---------------------------------------------------------------------------------------
'
Public Function ShouldBeWebOrderable(statuscode As String, QtyOnHand As Long, availLimit As Long, overrideOrderable As Boolean) As Boolean
    Mouse.Hourglass True
    If IsItemDC(statuscode) Then
        ShouldBeWebOrderable = Not CBool(QtyOnHand <= availLimit)
    Else
        ShouldBeWebOrderable = overrideOrderable
    End If
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : ShouldBeWebOrderableForItem
' DateTime  : 3/6/2006 12:37
' Author    : briandonorfio
' Purpose   : Wrapper for ShouldBeWebOrderable.
'---------------------------------------------------------------------------------------
'
Public Function ShouldBeWebOrderableForItem(item As String) As String
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM vOrderabilityInfo WHERE ItemNumber='" & item & "'")
    ShouldBeWebOrderableForItem = ShouldBeWebOrderable(rst("ItemStatus"), rst("QtyOnHand"), rst("AvailabilityLimit"), IIf(IsNull(rst("Orderable")), True, rst("Orderable")))
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetAvailabilityText
' DateTime  : 3/6/2006 12:37
' Author    : briandonorfio
' Purpose   : Returns the availability text, based on discontinued status, availability
'             overrides.
'---------------------------------------------------------------------------------------
'
Public Function GetAvailabilityText(statuscode As String, QtyOnHand As Long, QtyOnSO As Long, QtyOnPO As Long, QtyOrdered As Long, availLimit As Long, reqForm As Boolean, overrideString As String) As String
    Mouse.Hourglass True
    If overrideString <> "" Then
        GetAvailabilityText = overrideString
    Else
        If statuscode = 0 Then
            'this is "special order", what to put on the website?
            GetAvailabilityText = "Usually ships in 1 - 2 weeks"
        'ElseIf reqForm Then
        '    GetAvailabilityText = "Special order, call for availability"
        ElseIf IsItemDSable(statuscode) Then
            'regular d/s-able item, like MLW, with no special string
            'GetAvailabilityText = "Usually ships within 2 to 3 days"
            GetAvailabilityText = "Subject to manufacturer availability [more]"
        ElseIf Not IsItemStocked(statuscode) Then
            'no stock flag, so assume it's d/c without checking inventory amount
            GetAvailabilityText = "This item is not stocked or has been discontinued"
        ElseIf IsItemDC(statuscode) Then
            'uses availlimit, so we can't use IsItemCompletelyDC()
            'note: if we ever go back to qtyonso factoring in, need to change weboffload.getOutOfStockStatus()
            'If (QtyOnHand - QtyOnSO) <= availLimit And QtyOnPO = 0 Then
            If QtyOnHand <= availLimit And QtyOnPO = 0 And QtyOrdered = 0 Then
                'discontinued by us or them, none left for website
                GetAvailabilityText = "This item is not stocked or has been discontinued"
            Else
                'being d/c
                'GetAvailabilityText = "A stock item usually ships in 1 to 3 days"
                GetAvailabilityText = "Usually ships within 1 day"
            End If
        Else
            'regular stock item, regular string
            'GetAvailabilityText = "A stock item usually ships in 1 to 3 days"
            GetAvailabilityText = "Usually ships within 1 day"
        End If
    End If
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetAvailabilityTextForItem
' DateTime  : 3/6/2006 12:38
' Author    : briandonorfio
' Purpose   : Wrapper for GetAvailabilityText.
'---------------------------------------------------------------------------------------
'
Public Function GetAvailabilityTextForItem(item As String) As String
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM vOrderabilityInfo WHERE ItemNumber='" & item & "'")
    GetAvailabilityTextForItem = GetAvailabilityText(rst("ItemStatus"), rst("QtyOnHand"), rst("QtyOnSO"), rst("QtyOnPO"), rst("QtyOrdered"), rst("AvailabilityLimit"), rst("ShowRequestForm"), Nz(rst("AvailString")))
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : FillDCSpecs
' DateTime  : 3/6/2006 12:40
' Author    : briandonorfio
' Purpose   : Fills in the specs with the stock discontinued text.
'---------------------------------------------------------------------------------------
'
Public Function FillDCSpecs(item As String, specNumber As Long)
    Mouse.Hourglass True
    Select Case specNumber
        Case Is = 1
            FillDCSpecs = "No longer available..."
        Case Is = 2
            Dim replacePart As String, retval As String
            replacePart = HasReplacementPart(item)
            If replacePart = "" Then
                FillDCSpecs = ""
            Else
                'the rtml templates depend on this exact wording for the canonical link
                'functions. rtml regexp is:
                '    "^.+Consider\\s+the\\s+<a\\s+href=\"([-a-z0-9]+)\\.html\">.+$"
                'if this text changes from that match, then the template needs to be
                'changed as well, otherwise the canonical link tags will be massively
                'fucked up, and will probably break the page (inserting @ext-specs
                'html into the head tags, fun!)
                retval = "Consider the <a href=""LINE-PART.html""><strong>LINE-PART</strong></a>"
                retval = Replace(retval, "LINE-PART", CreateYahooID(replacePart), , 1)
                retval = Replace(retval, "LINE-PART", IIf(Left(replacePart, 3) <> Left(item, 3), TitleCase(DLookup("ManufFullNameWeb", "ProductLine", "ProductLine='" & Left(replacePart, 3) & "'")) & " ", "") & Mid(replacePart, 4))
                FillDCSpecs = retval
            End If
        Case Else
            FillDCSpecs = ""
    End Select
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : FillBaseItemSpecs
' DateTime  : 8/3/2006 10:55
' Author    : briandonorfio
' Purpose   : Fills in the specs with the stock text if it has a base item. Doesn't
'             really belong in this module, but it goes with FillDCSpecs, I guess.
'---------------------------------------------------------------------------------------
'
Public Function FillBaseItemSpecs(item As String, specNumber As Long)
    Mouse.Hourglass True
    Select Case specNumber
        Case Is = 1
            FillBaseItemSpecs = "SEE BASE ITEM"
        Case Else
            FillBaseItemSpecs = ""
    End Select
    Mouse.Hourglass False
End Function

Public Function FillTemplateItemSpecs(groupMaster As String, specNumber As Long)
    Mouse.Hourglass True
    Select Case specNumber
        Case Is = 1
            FillTemplateItemSpecs = "PART OF ITEM GROUP"
        Case Is = 2
            FillTemplateItemSpecs = "SEE " & groupMaster & " FOR SPECS"
        Case Is = 3
            FillTemplateItemSpecs = ""
        Case Is = 4
            FillTemplateItemSpecs = "NOTE: YOU STILL NEED TO FILL OUT DESC 1-4, ACC, ETC"
        Case Else
            FillTemplateItemSpecs = ""
    End Select
    Mouse.Hourglass False
End Function

Public Function HasReplacementPart(item As String) As String
    Dim replacePart As String
    replacePart = DLookup("ReplacedBy", "InventoryMaster", "ItemNumber='" & item & "'")
    If replacePart = "" Or Left(replacePart, 1) = "\" Or Left(replacePart, 1) = "/" Then
        replacePart = ""
    End If
    HasReplacementPart = replacePart
End Function
