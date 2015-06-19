Attribute VB_Name = "EBay"
Option Explicit

#Const IS_AUTO_IMPORT = False

Public Enum EBayStatus
    EBayStatusDown = 0
    EbayStatusUp = 1
End Enum

'Private Const EBAY_FVF_TIER_1_PCT As Long = 11
'Private Const EBAY_FVF_TIER_2_PCT As Long = 6
'Private Const EBAY_FVF_TIER_3_PCT As Long = 2
'Private Const EBAY_FVF_TIER_1_LIMIT As Long = 50
'Private Const EBAY_FVF_TIER_2_LIMIT As Long = 1000
'Private Const EBAY_FVF_DISCOUNT_PCT As Long = 20
Private Const EBAY_FVF_PCT As Variant = 9
Private Const EBAY_FVF_DISCOUNT_PCT As Long = 20
Private Const EBAY_PP_FLAT As Variant = 0.3
Private Const EBAY_PP_PCT As Variant = 1.9

Private ebaySellingLimitReachedFlag As Boolean
Private ebayAPICallLimitReachedFlag As Boolean
Private ebayLastUnhandledErrorMessage As String
Private ebayImagesMissing As PerlArray
Private ebayImagesSmall As PerlArray
Private ebayImagesText As PerlArray

Public Function EBaySellingLimitReached() As Boolean
    EBaySellingLimitReached = ebaySellingLimitReachedFlag
End Function

Public Function EBayAPICallLimitReached() As Boolean
    EBayAPICallLimitReached = ebayAPICallLimitReachedFlag
End Function

Public Function EBayGetLastUnhandledErrorMessage() As String
    EBayGetLastUnhandledErrorMessage = ebayLastUnhandledErrorMessage
End Function

Public Function EBayGetLastImageErrors() As String
    Dim parts As PerlArray
    Set parts = New PerlArray
    If Not ebayImagesMissing Is Nothing Then
        If ebayImagesMissing.Scalar <> 0 Then
            parts.Push CStr("MISSING IMAGES:" & vbCrLf & ebayImagesMissing.Join(vbCrLf))
        End If
    End If
    If Not ebayImagesSmall Is Nothing Then
        If ebayImagesSmall.Scalar <> 0 Then
            parts.Push CStr("IMAGE TOO SMALL:" & vbCrLf & ebayImagesSmall.Join(vbCrLf))
        End If
    End If
    If Not ebayImagesText Is Nothing Then
        If ebayImagesText.Scalar <> 0 Then
            parts.Push CStr("IMAGE HAS TEXT:" & vbCrLf & ebayImagesText.Join(vbCrLf))
        End If
    End If
    
    If parts.Scalar = 0 Then
        EBayGetLastImageErrors = ""
    Else
        EBayGetLastImageErrors = parts.Join(vbCrLf)
    End If
End Function

Public Function EBayGetProbableEBayCategoryIDFor(item As String) As String
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT TOP 1 EBayCategoryID, COUNT(*) AS TotalOccurrences " & _
                          "FROM PartNumbers " & _
                          "WHERE EBayCategoryID IS NOT NULL " & _
                          "  AND ItemNumber IN (SELECT ItemNumber " & _
                          "                     FROM PartNumberWebPaths " & _
                          "                     WHERE ItemNumber<>'" & item & "' " & _
                          "                       AND WebPathID IN (SELECT WebPathID " & _
                          "                                         FROM PartNumberWebPaths " & _
                          "                                         WHERE ItemNumber='" & item & "')) " & _
                          "  AND EBayCategoryID IN (SELECT CategoryID " & _
                          "                         FROM EBayCategory " & _
                          "                         WHERE CategoryType=0) " & _
                          "GROUP BY EBayCategoryID " & _
                          "ORDER BY TotalOccurrences DESC")
    If rst.EOF Then
        EBayGetProbableEBayCategoryIDFor = ""
    Else
        EBayGetProbableEBayCategoryIDFor = rst("EBayCategoryID")
    End If
    rst.Close
    Set rst = Nothing
End Function

Public Function EBayGetProbableStoreCategoryIDFor(item As String) As String
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT TOP 1 EBayStoreCategoryID, COUNT(*) AS TotalOccurrences " & _
                          "FROM PartNumbers " & _
                          "WHERE LEFT(ItemNumber,3)='" & Left(item, 3) & "' " & _
                          "  AND EBayStoreCategoryID IS NOT NULL " & _
                          "  AND EBayStoreCategoryID IN (SELECT CategoryID " & _
                          "                              FROM EBayCategory " & _
                          "                              WHERE CategoryType=1) " & _
                          "GROUP BY EBayStoreCategoryID " & _
                          "ORDER BY TotalOccurrences DESC")
    If rst.EOF Then
        EBayGetProbableStoreCategoryIDFor = ""
    Else
        EBayGetProbableStoreCategoryIDFor = rst("EBayStoreCategoryID")
    End If
    rst.Close
    Set rst = Nothing
End Function

Public Function EBayCalculateTotalSellingFees(itemSellPrice As String) As Variant
    EBayCalculateTotalSellingFees = EBayCalculatePayPalFlatFee(itemSellPrice) + _
                                    EBayCalculatePayPalPercentageFee(itemSellPrice) + _
                                    EBayCalculateFinalValueFee(itemSellPrice)
End Function

Public Function EBayCalculatePayPalFlatFee(itemSellPrice As String) As Variant
    EBayCalculatePayPalFlatFee = CDec(EBAY_PP_FLAT)
End Function

Public Function EBayCalculatePayPalPercentageFee(itemSellPrice As String) As Variant
    EBayCalculatePayPalPercentageFee = CDec(Math.Round(itemSellPrice * EBAY_PP_PCT / 100, 2))
End Function

Public Function EBayCalculateFinalValueFee(itemSellPrice As String) As Variant
    'similar function dbo.CalcEBayCommission(@price)
    Dim p As Variant
    p = CDec(Replace(Replace(itemSellPrice, "$", ""), ",", ""))
    Dim fee As Variant
'    fee = CDec(EBAY_FVF_TIER_1_PCT / 100 * minimumOfValues(p, EBAY_FVF_TIER_1_LIMIT))
'    If p > EBAY_FVF_TIER_1_LIMIT Then
'        fee = fee + CDec(EBAY_FVF_TIER_2_PCT / 100 * (minimumOfValues(EBAY_FVF_TIER_2_LIMIT, p) - EBAY_FVF_TIER_1_LIMIT))
'    End If
'    If p > EBAY_FVF_TIER_2_LIMIT Then
'        fee = fee + CDec(EBAY_FVF_TIER_3_PCT / 100 * (p - EBAY_FVF_TIER_2_LIMIT))
'    End If
'    fee = fee * (1 - EBAY_FVF_DISCOUNT_PCT / 100)
'    EBayCalculateFinalValueFee = CDec(Math.Round(fee, 2))
    fee = CDec(EBAY_FVF_PCT / 100 * p)
    If fee > 250 Then
        fee = 250
    End If
    EBayCalculateFinalValueFee = CDec(Math.Round(fee, 2))
End Function

Public Function EBayCalculateFinalValueDiscount(fvf As Variant) As Variant
    EBayCalculateFinalValueDiscount = CDec(CDec(fvf) * (EBAY_FVF_DISCOUNT_PCT / 100))
End Function

Public Function EBayGetAvailableQuantityFor(item As String, QtyOnHand As Long, VendorQtyOnHand As Long, QtyOnSO As Long, dsable As Boolean, dsOnly As Boolean, availLimit As Long, ebayReserveQty As Long) As Long
    If dsOnly Then
        EBayGetAvailableQuantityFor = VendorQtyOnHand - availLimit - ebayReserveQty
    ElseIf dsable Then
        EBayGetAvailableQuantityFor = QtyOnHand + VendorQtyOnHand - QtyOnSO - availLimit - ebayReserveQty
    Else
        EBayGetAvailableQuantityFor = QtyOnHand - QtyOnSO - availLimit - ebayReserveQty
    End If
End Function

Public Function EBayGetAvailableQuantityForKit(kit As String) As Long
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT K.ComponentItemCode, K.QuantityPerAssembly, Q.QtyOnHand, Q.QtyOnSO, P.AvailabilityLimit, P.EBayReserveQty, I.VendorQuantityOnHand, S.DSable, S.DSonly " & _
                          "FROM IM_SalesKitDetail AS K " & _
                          "       INNER JOIN vActualWhseQty AS Q ON K.ComponentItemCode=Q.ItemNumber " & _
                          "       INNER JOIN vItemStatusBreakdown AS S ON K.ComponentItemCode=S.ItemNumber " & _
                          "       INNER JOIN InventoryMaster AS I ON K.ComponentItemCode=I.ItemNumber " & _
                          "       INNER JOIN PartNumbers AS P ON K.ComponentItemCode=P.ItemNumber " & _
                          "WHERE K.SalesKitNo='" & kit & "'")
    Dim retval As Long
    retval = 999999
    While Not rst.EOF
        Dim thisQty As Long
        thisQty = EBayGetAvailableQuantityFor(rst("ComponentItemCode"), rst("QtyOnHand"), Nz(rst("VendorQuantityOnHand"), "0"), rst("QtyOnSO"), rst("DSable"), rst("DSonly"), rst("AvailabilityLimit"), rst("EBayReserveQty"))
        thisQty = Int(thisQty / CLng(rst("QuantityPerAssembly")))
        If thisQty < retval Then
            retval = thisQty
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    If retval = 999999 Then
        EBayGetAvailableQuantityForKit = 0
    Else
        EBayGetAvailableQuantityForKit = retval
    End If
End Function

Public Function EBayItemIDFor(item As String) As String
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT EBayItemID FROM PartNumbers WHERE ItemNumber='" & item & "'")
    EBayItemIDFor = Nz(rst("EBayItemID"))
    rst.Close
    Set rst = Nothing
End Function

Public Function EBayHasItemID(item As String) As Boolean
    Dim itemID As String
    itemID = EBayItemIDFor(item)
    EBayHasItemID = CBool(itemID <> "")
End Function

Public Function EBayIsItemPublishable(item As String) As String
    Dim retval As PerlArray
    Set retval = New PerlArray
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT EBayCategoryID FROM PartNumbers WHERE ItemNumber='" & item & "'")
    If Nz(rst("EBayCategoryID")) = "" Then
        retval.Push "Missing category!"
    End If
    rst.Close
    
    Set rst = DB.retrieve("SELECT EBayPrice FROM InventoryMaster WHERE ItemNumber='" & item & "'")
    If Nz(rst("EBayPrice"), "0") = 0 Then
        retval.Push "Missing price!"
    End If
    rst.Close
    
    Set rst = DB.retrieve("SELECT dbo.StatusIsDSonly(ItemStatus) AS DropshipOnly, ShipsFromZip FROM InventoryMaster INNER JOIN ProductLine ON InventoryMaster.ProductLine=ProductLine.ProductLine WHERE ItemNumber='" & item & "'")
    If rst("DropshipOnly") = True Then
        If Nz(rst("ShipsFromZip")) = "" Then
            retval.Push "Needs vendor ship-from zip code filled out!"
        End If
    End If
    rst.Close
    
    Set rst = DB.retrieve("SELECT ManufFullNameWeb FROM ProductLine WHERE ProductLine='" & Left(item, 3) & "'")
    If Nz(rst("ManufFullNameWeb"), "") = "" Then
        retval.Push "Needs a manufacturer web name! (poinv ""E"" button, 3rd name blank)"
    End If
    rst.Close
    
    Dim conflictingItem As String
    conflictingItem = CheckConflictingItem(item)
    If conflictingItem <> "" Then
        retval.Push "Name conflicts with item " & conflictingItem
    End If
    
    If retval.Scalar > 0 Then
        EBayIsItemPublishable = retval.Join(vbCrLf)
    Else
        EBayIsItemPublishable = "OK"
    End If
End Function

Public Function EBayCurrentStatus(item As String, Optional statusID As Long = -1) As EBayStatus
    If statusID = -1 Then
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT EBayStatusID FROM PartNumbers WHERE ItemNumber='" & item & "'")
        statusID = CLng(rst("EBayStatus"))
        rst.Close
        Set rst = Nothing
    End If
    
    Select Case statusID
        Case Is = 0
            EBayCurrentStatus = EBayStatus.EBayStatusDown
        Case Is = 1
            EBayCurrentStatus = EBayStatus.EbayStatusUp
    End Select
End Function

Public Function EBayIsItemQtyUpdateRequired(quantityAtEBay As Long, adjustedQOH As Long) As Boolean
    If quantityAtEBay = adjustedQOH Then
        'equal, so nothing to do
        EBayIsItemQtyUpdateRequired = False
    ElseIf adjustedQOH <= 0 And quantityAtEBay > 0 Then
        'no more available, but up on ebay, update
        EBayIsItemQtyUpdateRequired = True
    ElseIf adjustedQOH <= 0 Then
        'no more available, not up on ebay, nothing to do
        EBayIsItemQtyUpdateRequired = False
    'ElseIf adjustedQOH > 11 And quantityAtEBay > 11 Then
    '    'quantity shows as 10+ available, nothing to do even if large difference
    '    EBayIsItemQtyUpdateRequired = False
    Else
    '    'equal has already been done, so this is a <=10 (either us or them), so we want the exact qty to show, update
    ' always updating quantity to exact amount now
    ' TODO: this includes active ebay s/o in the adjusted quantity, when they've
    ' already been pulled out by ebay automatically. after we get history and
    ' current broken down by channel, only apply non-ebay active orders to this
        EBayIsItemQtyUpdateRequired = True
    End If
End Function

Private Function ebayUpdateToPublished(item As String, ebayItemID As String, ebayQty As Long, Optional currentStatus As EBayStatus) As Boolean
    'If IsMissing(currentStatus) Then
    '    currentStatus = EBayCurrentStatus(item)
    'End If
    
    DB.Execute "UPDATE PartNumbers SET EBayStatusID=" & EBayStatus.EbayStatusUp & ", EBayItemID='" & EscapeSQuotes(ebayItemID) & "', EBayAvailableQty=" & ebayQty & ", EBaySoldQty=0, EBayOffloadRequired=0 WHERE ItemNumber='" & item & "'"
    
    'this should mirror the itemid up to the VPS? selling something quick would cause problems otherwise
    #If IS_AUTO_IMPORT = False Then
    ExternalItemDBSync item
    #End If
    
    ebayUpdateToPublished = True
End Function

Private Function ebayUpdateToRevised(item As String, ebayItemID As String, ebayQty As Long, Optional currentStatus As EBayStatus) As Boolean
    'If IsMissing(currentStatus) Then
    '    currentStatus = EBayCurrentStatus(item)
    'End If
    
    DB.Execute "UPDATE PartNumbers SET EBayStatusID=" & EBayStatus.EbayStatusUp & ", EBayItemID='" & EscapeSQuotes(ebayItemID) & "', EBayAvailableQty=EBaySoldQty + " & ebayQty & ", EBayOffloadRequired=0 WHERE ItemNumber='" & item & "'"
    
    ebayUpdateToRevised = True
End Function

Private Function ebayUpdateToRelisted(item As String, ebayItemID As String, ebayQty As Long, Optional currentStatus As EBayStatus) As Boolean
    'If IsMissing(currentStatus) Then
    '    currentStatus = EBayCurrentStatus(item)
    'End If
    
    DB.Execute "UPDATE PartNumbers SET EBayStatusID=" & EBayStatus.EbayStatusUp & ", EBayItemID='" & EscapeSQuotes(ebayItemID) & "', EBayAvailableQty=" & ebayQty & ", EBaySoldQty=0, EBayOffloadRequired=0 WHERE ItemNumber='" & item & "'"
    
    'this should mirror the itemid up to the VPS? selling something quick would cause problems otherwise
    #If IS_AUTO_IMPORT = False Then
    ExternalItemDBSync item
    #End If
    
    ebayUpdateToRelisted = True
End Function

Private Function ebayUpdateToDelisted(item As String, Optional currentStatus As EBayStatus) As Boolean
    'If IsMissing(currentStatus) Then
    '    currentStatus = EBayCurrentStatus(item)
    'End If
    
    ' should this wipe out qty?
    ' let's keep it, in case we want the historical info for whatever reason (at least until it gets wiped when back in stock)
    DB.Execute "UPDATE PartNumbers SET EBayStatusID=" & EBayStatus.EBayStatusDown & ", EBayOffloadRequired=0, EBayUnlistedTime=GETDATE() WHERE ItemNumber='" & item & "'"
    
    ebayUpdateToDelisted = True
End Function

Public Function EBayMarkRevisionRequired(item As String, Optional currentStatus As EBayStatus) As Boolean
    'If IsMissing(currentStatus) Then
    '    currentStatus = EBayCurrentStatus(item)
    'End If
    
    DB.Execute "UPDATE PartNumbers SET EBayOffloadRequired=1 WHERE ItemNumber='" & item & "'"
    
    EBayMarkRevisionRequired = True
End Function

Public Function EBayMarkRevisionNotRequired(item As String, Optional currentStatus As EBayStatus) As Boolean
    'If IsMissing(currentStatus) Then
    '    currentStatus = EBayCurrentStatus(item)
    'End If
    
    DB.Execute "UPDATE PartNumbers SET EBayOffloadRequired=0 WHERE ItemNumber='" & item & "'"
    
    EBayMarkRevisionNotRequired = True
End Function

Public Function EBayOffloadItemsAll(Optional whereClause As String = "") As PerlArray
    Set ebayImagesMissing = Nothing
    Set ebayImagesMissing = New PerlArray
    Set ebayImagesSmall = Nothing
    Set ebayImagesSmall = New PerlArray
    Set ebayImagesText = Nothing
    Set ebayImagesText = New PerlArray
    Dim rst As ADODB.Recordset
    Dim query As String
    query = "SELECT * FROM vWebOffloadEBay" & IIf(whereClause <> "", " WHERE " & whereClause, "") & " ORDER BY ItemNumber DESC"
    Set rst = DB.retrieve(query)
    'MsgBox rst("EBayItemID")
    Dim dccheck As ADODB.Recordset
    Set dccheck = DB.retrieve("SELECT IsBeingDC,IsDC,IsDConWeb FROM vInventorySpreadsheet WHERE ItemNumber='" & rst("ItemNumber") & "'")
    Dim retval As PerlArray
    Set retval = New PerlArray
    While Not rst.EOF
        Dim qty As Long
        If rst("IsMasKit") Then
            qty = EBayGetAvailableQuantityForKit(rst("ItemNumber"))
        Else
            qty = EBayGetAvailableQuantityFor(rst("ItemNumber"), rst("QtyOnHand"), Nz(rst("VendorQuantityOnHand"), "0"), rst("QtyOnSO"), rst("Dropshippable"), rst("DropshipOnly"), rst("AvailabilityLimit"), rst("EBayReserveQty"))
        End If
        If CDec(rst("EBayPrice")) >= 1000 Then
            If qty > 20 Then
                qty = 20
            End If
        ElseIf CDec(rst("EBayPrice")) >= 500 Then
            If qty > 50 Then
                qty = 50
            End If
        ElseIf CDec(rst("EBayPrice")) >= 100 Then
            If qty > 250 Then
                qty = 250
            End If
        End If
        Dim thisitem As Dictionary
        Set thisitem = New Dictionary
        Dim changetype As String
        If rst("EBayPublished") = True Then
        'Tom's Code 12-31-2014 to get mpn in dictionary...
        'thisitem.Add "ManufFullNameWeb", CStr(rst("ManufFullNameWeb"))
        
            'we want it published, can be add, relist, or revise
            If rst("EBayStatusID") = EBayStatus.EBayStatusDown And qty > 0 Then
                'we want it up, it's down, we have it in stock: add or relist
                If IsNull(rst("EBayItemID")) Or IsNull(rst("EBayUnlistedTime")) Then
                    changetype = "Add"
                ElseIf DateDiff("d", rst("EBayUnlistedTime"), Now()) > 90 Then
                    changetype = "Add"
                Else
                
                    changetype = "Relist"
                    
                End If
            ElseIf rst("EBayStatusID") = EBayStatus.EBayStatusDown Then ' and qty<=0
                'we want it up, it's down, but zero on hand: use out of stock option
                changetype = "NOTHING"
                
                'changetype = "Relist"
            ElseIf rst("EBayStatusID") = EBayStatus.EbayStatusUp And qty > 0 Then
                'we want it up, it's up, we have it in stock: revise, i guess? no way to tell the specific change required
                changetype = "Revise"
            ElseIf rst("EBayStatusID") = EBayStatus.EbayStatusUp And qty <= 0 Then
                'we want it up, it's up, but zero on hand: end it
                'changetype = "End"
                'Tom's Code 11-26-2014 to allow for EBay's 'Out-Of-Stock' Option.
                
                If dccheck.RecordCount > 0 Then
                    If CLng(qty) <= 0 Then
                        If dccheck("IsDC") = "Y" Or dccheck("IsBeingDC") = "Y" Or dccheck("IsDConWeb") = "Y" Then
                            changetype = "End"
                        Else
                        
                            changetype = "Revise"
                            'MsgBox rst("Quantity")
                        End If
                        
                    End If
                End If
                
            Else
                'error
                MsgBox "ERROR: unhandled ebay status id " & rst("EBayStatusID") & " for item " & rst("ItemNumber")
                changetype = "NOTHING"
            End If
        ElseIf rst("EBayPublished") = False Then
            If rst("EBayStatusID") = EBayStatus.EbayStatusUp Then
                'we want it unpublished, it's up: end it
                'Tom's Code to modify for EBay's 'Out-Of-Stock' Option.
                'If qty <= 0 Then
                'changetype = "Revise"
                'Else
                changetype = "End"
                'End If
            Else
                'we want it unpublished, it's down: do nothing
                changetype = "NOTHING"
            End If
        End If
        
        If changetype = "NOTHING" Then
            EBayMarkRevisionNotRequired rst("ItemNumber")
        Else
            thisitem.Add "Action", CStr(changetype)
            thisitem.Add "ItemNumber", CStr(rst("ItemNumber"))
            
            'for references to an active item, item id is required
            Select Case changetype
                Case Is = "Revise", "Relist", "End"
                    If dccheck("IsDC") = "Y" Or dccheck("IsBeingDC") = "Y" Or dccheck("IsDConWeb") = "Y" Then
                            changetype = "End"
                        Else
                        
                        
                        If (IsNull(rst("EBayItemID")) = False) Then
                            thisitem.Add "ItemID", CStr(rst("EBayItemID")) 'crash on null, but that's bad anyway
                        End If
                        
                        End If
                    
            End Select
            'add fields for content changes
            Select Case changetype
                Case Is = "Add", "Revise", "Relist"
                    Dim itemTitle As String
                    itemTitle = Nz(rst("ManufFullNameWeb")) & " " & Nz(rst("Desc2")) & " " & Nz(rst("Desc1")) & " " & Nz(rst("Desc3")) & " " & Nz(rst("Desc4")) & IIf(rst("IsReconditioned"), "", " New")
                    itemTitle = TrimWhitespace(itemTitle, True, True, True, False)
                    If Len(itemTitle) > 80 Then
                        itemTitle = Nz(rst("ManufFullNameWeb")) & " " & Nz(rst("Desc2")) & " " & Nz(rst("Desc1")) & " " & Nz(rst("Desc3")) & IIf(rst("IsReconditioned"), "", " New")
                        itemTitle = TrimWhitespace(itemTitle, True, True, True, False)
                        If Len(itemTitle) > 80 Then
                            itemTitle = Nz(rst("ManufFullNameWeb")) & " " & Nz(rst("Desc2")) & " " & Nz(rst("Desc3")) & IIf(rst("IsReconditioned"), "", " New")
                            itemTitle = TrimWhitespace(itemTitle, True, True, True, False)
                        End If
                    End If
                    'error 21919128 - cannot change a category if item has been sold
                    'handle this error elsewhere somehow. we need category id for secondary cat...
                    'If changetype = "Add" Or changetype = "Relist" Or rst("EBaySoldQty") = 0 Then
                        thisitem.Add "Category", CStr(rst("EBayCategoryID"))
                        thisitem.Add "StoreCategory", CStr(Nz(rst("EBayStoreCategoryID"), "0"))
                    'End If
                    'If InStr(CStr(rst("ItemNumber")), "HAN") = True Then
                    'MsgBox "Item Number is " + CStr(rst("ItemNumber"))
                    'End If
                    
                    
                    'Tom's Code 4-21-2015 to hook up the multiple conditions that ebay has...
                        'thisitem.Add "ConditionID", EBayCategoryHashIDtoConditionID.item(CStr(rst("EBayCategoryID")))(CLng(IIf(CBool(rst("IsReconditioned")), 1, 0)))
                    Dim conditionID As ADODB.Recordset
                    Set conditionID = DB.retrieve("SELECT ItemConditionMaster.EBayConditionID FROM ItemConditionLines INNER JOIN ItemConditionMaster ON ItemConditionMaster.ConditionID=ItemConditionLines.ConditionID WHERE ItemNumber='" & CStr(rst("ItemNumber")) & "'")
                    If (conditionID.RecordCount > 0) Then
                        If IsNull(conditionID("EBayConditionID")) = True Then
                            MsgBox "Tell Tom a defined item condition has no ebay condition id associated with it. The culprit item is " & CStr(rst("ItemNumber"))
                        Else
                            thisitem.Add "ConditionID", conditionID("EBayConditionID")
                        End If
                    Else
                        'MsgBox "Tell Tom item " & CStr(rst("ItemNumber")) & " does not have a condition associated with it (how bizzare)"
                        thisitem.Add "ConditionID", "1000"
                    End If
                    
                    
                    
                    
                    'error 10039 - cannot change a title if item has been sold
                    If changetype = "Add" Or changetype = "Relist" Or rst("EBaySoldQty") = 0 Then
                        thisitem.Add "Title", itemTitle
                    End If
                    #If IS_AUTO_IMPORT = False Then
                    thisitem.Add "Description", WebOffloadFunctions.GenerateItemPageContentsForEBay(rst)
                    #End If
                    If changetype = "Revise" And CLng(qty) < 0 Then
                        thisitem.Add "Quantity", 0
                    Else
                        thisitem.Add "Quantity", CLng(qty)
                    End If
                    
                    
                    
                    
                    thisitem.Add "StartPrice", CDec(rst("EBayPrice"))
                    thisitem.Add "AllowBestOffer", CBool(rst("EBayAllowBestOffer"))
                    'thisitem.Add "BestOfferAutoDecline", CDec(rst("EBayBestOfferAutoDecline"))
                    thisitem.Add "BestOfferAutoAccept", CDec(rst("EBayBestOfferAutoAccept"))
                    
                    'Tom's Code 10-2-2014
                    'to make listprice = ebayprice + 91%
                    'to make ebay mapp = our mapp + 5%
                    'revision: if list price exists, use it instead of ebay price
                    If CDec(rst("ListPrice")) > CDec(rst("EBayPrice")) Then
                      thisitem.Add "ListPrice", CDec(rst("ListPrice"))
                    Else
                    thisitem.Add "ListPrice", CDec(rst("EBayPrice") * 1.91)
                    End If
                  
                     
                    
                    thisitem.Add "MAPP", CDec(rst("MAPP") * 1.05)
                                       
                    'thisitem.Add "ListPrice", CDec(rst("ListPrice"))
                    'thisitem.Add "ListPrice", CDec(0)
                    thisitem.Add "TaxExempt", CBool(rst("TaxExempt"))
                    thisitem.Add "ShippingCharge", CDec(rst("ShippingCharge"))
                    thisitem.Add "Brand", CStr(rst("ManufFullNameWeb"))
                    thisitem.Add "ModelNo", CStr(rst("Desc2"))
                    #If IS_AUTO_IMPORT = False Then
                    thisitem.Add "UPC", BarcodeFunctions.FindBestBarcodeFor(rst("ItemNumber"), rst("ItemID"))
                    #End If
                    'thisitem.Add "PicURLs", PictureDBFunctions.GetAllItemImageURLsFor(rst("ItemNumber"), False, False, True)
                    'thisitem.Add "PicURLs", PictureDBFunctions.GetAllItemImageURLsFor(rst("ItemNumber"), True, True, True) '2013-07-08: switch to authorized images
                    #If IS_AUTO_IMPORT = False Then
                    thisitem.Add "PicURL", PictureDBFunctions.GenerateItemPictureURL(rst("ItemNumber"), True, Array("noauth-nobg", "auth-bg"))
                    #End If
                    thisitem.Add "ShipFromPostalCode", CStr(rst("ShipFromPostalCode")) 'null->06712 handled in view
            End Select
            'type-specific changes
            Select Case changetype
                Case Is = "End"
                    thisitem.Add "EndCode", "NotAvailable"
            End Select
            'MsgBox thisitem.item("ItemID")
            'MsgBox thisitem.item("MAPP")
            retval.Push thisitem
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    Set EBayOffloadItemsAll = retval
End Function


Private Function minimumOfValues(ParamArray vals() As Variant) As Variant
    Dim retval As Variant
    retval = CDec(vals(0))
    Dim i As Long
    For i = 0 To UBound(vals)
        If CDec(vals(i)) < CDec(retval) Then
            retval = CDec(vals(i))
        End If
    Next i
    minimumOfValues = CDec(retval)
End Function


'--------------------
' API FUNCTION CALLS
'--------------------
Public Function EBayAPIAddFixedPriceItem(def As Dictionary) As Boolean
'EBayAPIAddFixedPriceItem = EBayAPIAddItem(def)
'Exit Function





    ebaySellingLimitReachedFlag = False
    ebayAPICallLimitReachedFlag = False
    
    Dim xmlin As String
    xmlin = ebayGenerateAddFixedPriceItemXML(def)
    Dim xmlout As String
    xmlout = ebayDoXMLRequest(xmlin, "AddFixedPriceItem")
    
    If xmlout = "" Then
        'error
        EBayAPIAddFixedPriceItem = False
        Exit Function
    End If
    
    
    
    'Dim filefree As Long
    'filefree = FreeFile
    'Open "c:\users\tomwestbrook\desktop\ebay-second-category.txt" For Output As #filefree
    'Print #filefree, xmlin & vbCrLf & vbCrLf & xmlout
    'Close #filefree
    
    Dim root As MSXML2.DOMDocument
    Set root = ebayParseXML(xmlin, xmlout, def.item("ItemNumber"), "AddFixedPriceItemResponse")
    
    If Not root Is Nothing Then
        'here's the actual work, just grab the id and update with qty
        Dim id As String
        id = root.selectSingleNode("/AddFixedPriceItemResponse/ItemID").Text
        ebayUpdateToPublished def.item("ItemNumber"), id, def.item("Quantity")
        EBayAPIAddFixedPriceItem = True
    Else
        'already emailed, nothing to do
        EBayAPIAddFixedPriceItem = False
    End If
End Function

Public Function EBayAPIRelistFixedPriceItem(def As Dictionary) As Boolean
'   EBayAPIRelistFixedPriceItem = EBayAPIRelistItem(def)
'   Exit Function
    
    
    ebaySellingLimitReachedFlag = False
    ebayAPICallLimitReachedFlag = False
    
    Dim xmlin As String
    xmlin = ebayGenerateRelistFixedPriceItemXML(def)
    Dim xmlout As String
    xmlout = ebayDoXMLRequest(xmlin, "RelistFixedPriceItem")
    
    'Dim filefree As Long
    'filefree = FreeFile
    'Open "c:\users\tomwestbrook\desktop\ebay-second-category.txt" For Output As #filefree
    'Print #filefree, xmlin & vbCrLf & vbCrLf & xmlout
    'Close #filefree
    'Dim filefree As Integer
    'filefree = FreeFile
    'Open "c:\users\tomwestbrook\desktop\ebayerror.txt" For Output As #filefree
    'Print #filefree, xmlin
    'Close #filefree
    
    If xmlout = "" Then
        'error
        EBayAPIRelistFixedPriceItem = False
        Exit Function
    End If
    
    Dim root As MSXML2.DOMDocument
    Set root = ebayParseXML(xmlin, xmlout, def.item("ItemNumber"), "RelistFixedPriceItemResponse")
    
    If Not root Is Nothing Then
        'grab id, update qty
        Dim id As String
        id = root.selectSingleNode("/RelistFixedPriceItemResponse/ItemID").Text
        ebayUpdateToRelisted def.item("ItemNumber"), id, def.item("Quantity")
        EBayAPIRelistFixedPriceItem = True
    Else
        'perhaps this item needed revising since we now use the ebay out of stock option...
        'EBayAPIRelistFixedPriceItem = EBayAPIReviseFixedPriceItem(def)
        
        EBayAPIRelistFixedPriceItem = False
    End If
End Function

Public Function EBayAPIReviseFixedPriceItem(def As Dictionary) As Boolean
   
 '   EBayAPIReviseFixedPriceItem = EBayAPIReviseItem(def)
 '   Exit Function
    
    
    ebaySellingLimitReachedFlag = False
    ebayAPICallLimitReachedFlag = False
    
    Dim xmlin As String
    'MsgBox def("ItemID")
    'MsgBox def.item("ItemID")
    'MsgBox def.item("MAPP")
    xmlin = ebayGenerateReviseFixedPriceItemXML(def)
    
    
    'xmlin = Replace(xmlin, "ReviseFixedPriceItem", "ReviseItem")
    'Dim filefree As Long
    'filefree = FreeFile
    'Open "c:\users\tomwestbrook\desktop\xmlIN.txt" For Output As #filefree
    '    Print #filefree, xmlin
    'Close #filefree
    
    
    
    Dim xmlout As String
    xmlout = ebayDoXMLRequest(xmlin, "ReviseFixedPriceItem")
    'xmlout = ebayDoXMLRequest(xmlin, "ReviseItem")
    
   'Dim filefree2 As Long
   ' filefree2 = FreeFile
   ' Open "c:\users\tomwestbrook\desktop\xmlOUT.txt" For Output As #filefree2
   '    Print #filefree2, xmlout & vbCrLf & vbCrLf & vbCrLf & xmlin
   ' Close #filefree2

    If xmlout = "" Then
        'error
        EBayAPIReviseFixedPriceItem = False
        Exit Function
    End If
    
    Dim root As MSXML2.DOMDocument
    

    
    
    Set root = ebayParseXML(xmlin, xmlout, def.item("ItemNumber"), "ReviseFixedPriceItemResponse")
    
    If Not root Is Nothing Then
        'make sure it's marked as up, then...?
        'qty update, item id stays the same
        ebayUpdateToRevised def.item("ItemNumber"), def.item("ItemID"), def.item("Quantity")
        EBayAPIReviseFixedPriceItem = True
    Else
        EBayAPIReviseFixedPriceItem = False
    End If
End Function

Public Function EBayAPIEndFixedPriceItem(def As Dictionary) As Boolean
'    EBayAPIEndFixedPriceItem = EBayAPIEndItem(def)
'    Exit Function
    
        
    ebaySellingLimitReachedFlag = False
    ebayAPICallLimitReachedFlag = False
    
    Dim xmlin As String
    xmlin = ebayGenerateEndFixedPriceItemXML(def)
    Dim xmlout As String
    xmlout = ebayDoXMLRequest(xmlin, "EndFixedPriceItem")
    
    If xmlout = "" Then
        'error
        EBayAPIEndFixedPriceItem = False
        Exit Function
    End If
    
    Dim root As MSXML2.DOMDocument
    Set root = ebayParseXML(xmlin, xmlout, def.item("ItemNumber"), "EndFixedPriceItemResponse")
    'if there is a db inconsistency with ebay (they don't always send the
    'notification that an item sold), then we might get an error response
    'of "auction has already ended", this is handled in the parse and
    'returned as a fake success code. if we ever do anything other than
    'just mark as delisted, this may need to be addressed.
    
    If Not root Is Nothing Then
        'just marking as down for now
        ebayUpdateToDelisted def.item("ItemNumber")
        EBayAPIEndFixedPriceItem = True
    Else
        EBayAPIEndFixedPriceItem = False
    End If
End Function

'---------------------------------------
' XML GENERATION AND REQ/RESP FUNCTIONS
'---------------------------------------
Private Function ebayParseXML(xmlin As String, xmlout As String, item As String, responseNode As String) As MSXML2.DOMDocument
    Dim parser As MSXML2.DOMDocument
    Set parser = New MSXML2.DOMDocument
    parser.async = False
    
    'Dim filefree As Long
    'filefree = FreeFile
    'Open "c:\users\tomwestbrook\desktop\wtf.txt" For Output As #filefree
    'Print #filefree, xmlout
    'Close #filefree
    
    Dim success As Boolean
    success = parser.loadXML(xmlout)
    If success = False Then
        SendEmailTo "brian@tools-plus.com", "ebay offload error - bad parse", xmlin & vbCrLf & vbCrLf & xmlout
        Set ebayParseXML = Nothing
        Exit Function
    End If
    
    Dim ack As String
    ack = parser.selectSingleNode("/" & responseNode & "/Ack").Text
    If ack = "Error" Then
        SendEmailTo "brian@tools-plus.com", "ebay offload error - api returned ack==error", xmlin & vbCrLf & vbCrLf & xmlout
        Set ebayParseXML = Nothing
        Exit Function
    End If
    
    ebayLastUnhandledErrorMessage = "<no error>"
    
    If ack = "Warning" Or ack = "Failure" Then
        Dim nodeList As MSXML2.IXMLDOMNodeList
        Set nodeList = parser.selectNodes("/" & responseNode & "/Errors")
        Dim node As MSXML2.IXMLDOMNode
        For Each node In nodeList
            Dim num As String
            num = node.selectSingleNode("ErrorCode").Text
            Dim epvaltext As String
            If node.selectSingleNode("ErrorParameters/Value") Is Nothing Then
                epvaltext = ""
            Else
                epvaltext = node.selectSingleNode("ErrorParameters/Value").Text
            End If
            Dim smtext As String
            If node.selectSingleNode("ShortMessage") Is Nothing Then
                smtext = ""
            Else
                smtext = node.selectSingleNode("ShortMessage").Text
            End If
            If num = "21916689" Then
                'ignore, "No product found for ProductListingDetails.<UPC> <885911212373>."
            ElseIf num = "21917236" Then
                'ignore, "Funds from your sales will be unavailable and show as pending in your PayPal account for a period of time."
            ElseIf num = "23015" Then
                'ignore, "If this item sells by a Best Offer, you will not be able to require immediate payment."
            ElseIf num = "21919002" Then
                'ignore, "Discounted pricing will not be applied to this item. The item does not qualify for discounted pricing."
                'what makes it not qualify?
            ElseIf num = "5116" Then
                'ignore, "The Item Specific &quot;10244&quot; was dropped. Please review your listing and update your Item Specifics if necessary."
            ElseIf num = "21919189" Then
                'ignore, You'll be able to list $XXX more this month.
            ElseIf responseNode = "EndFixedPriceItemResponse" And num = "1047" Then
                'ignore, "The auction has been closed."
                'XXX: this is a db sync issue that should probably be fixed
                'for now, we call it a success, and the calling function will
                'continue like normal (marking it down, which it is).
                ack = "Success"
            ElseIf num = "291" And responseNode = "EndFixedPriceItemResponse" Then
                'ignore, "Auction ended."
                'XXX: another sync issue, or something else? how are these ending?
                ack = "Success"
            ElseIf num = "291" And responseNode = "ReviseFixedPriceItemResponse" Then
                'XXX: sync issues, copying the stuff in EBayAPIEndFixedPriceItem,
                'then we can add with the next offload. other 291 errors (add/relist)
                'are unhandled, and will give errors. not sure how they would happen
                'since we wouldn't add/relist something we think is up.
                ebayUpdateToDelisted item
                ack = "Success"
            ElseIf num = "240" And CBool(InStr(epvaltext, "exceeded the dollar amount you can list")) Then
                'selling limit reached, that's bad
                ebaySellingLimitReachedFlag = True
                ebayLastUnhandledErrorMessage = "selling limit reached"
            ElseIf num = "518" And CBool(InStr(smtext, "Call usage limit has been reached")) Then
                'api call limit reached, that's really bad
                ebayAPICallLimitReachedFlag = True
                ebayLastUnhandledErrorMessage = "api call limit reached"
            Else
                'unhandled error, but we want to alert the offloader of the error
                If num = "3021" Then
                    'no email, "System temporarily unavailable, please try listing again."
                ElseIf num = "21919136" Or num = "106" Then
                    'no email, "Your listing must have at least 1 photo."
                    ebayImagesMissing.Push item
                ElseIf num = "21919137" Then
                    'no email, "Buyers love large photos that clearly show the item, so please upload high-resolution photos that are at least 500 pixels on the longest side."
                    ebayImagesSmall.Push item
                ElseIf num = "21919184" Then
                    'no email, "To meet our photo quality requirements and improve the chances of your item selling, upload a photo that has no added artwork or text."
                    ebayImagesText.Push item
                ElseIf num = "70" Then
                    'no email, "Title too long."
                ElseIf num = "10029" Or num = "21919028" Or num = "21916923" Then
                    'no email, "Title or subtitle cannot be revised for auction style listings that have at least one pending bid or are ending in less than twelve hours, and for fixed price listings that have at least one quantity sold or have a pending best offer."
                    'no email, "Portions of this listing cannot be revised if the item has bid or active Best Offers or is ending in 12 hours."
                    'no email, "Pictures cannot be edited because this listing has pending offers, or the maximum number of pictures has been reached."
                ElseIf num = "240" Then
                    '"The item cannot be listed or modified. The title and/or description may contain improper words, or the listing or seller may be in violation of eBay policy."
                    'there are a bunch of different reason in the <ErrorParameters>, but for now, we can
                    'just send this email off to kirk to handle. sorry kirk!
                    SendEmailTo "kirk@tools-plus.com", "ebay offload error", "--------------" & vbCrLf & "XML SENT:" & vbCrLf & "--------------" & vbCrLf & xmlin & vbCrLf & vbCrLf & vbCrLf & "--------------" & vbCrLf & "XML RECEIVED:" & vbCrLf & "--------------" & vbCrLf & xmlout
                Else
                    'other unhandled errors, going to kirk now.
                    SendEmailTo "kirk@tools-plus.com", "ebay offload error", "--------------" & vbCrLf & "XML SENT:" & vbCrLf & "--------------" & vbCrLf & xmlin & vbCrLf & vbCrLf & vbCrLf & "--------------" & vbCrLf & "XML RECEIVED:" & vbCrLf & "--------------" & vbCrLf & xmlout
                End If
                'looks like the first error in the xml is most important (other errors are because
                'of the previous error (e.g., invalid category -> autopay dropped)) so only grab
                'the first
                If ebayLastUnhandledErrorMessage = "<no error>" Then
                    If node.selectSingleNode("ShortMessage") Is Nothing Then
                        ebayLastUnhandledErrorMessage = "no error message?"
                    Else
                        ebayLastUnhandledErrorMessage = node.selectSingleNode("ShortMessage").Text
                    End If
                End If
            End If
        Next node
    End If
    
    If ack = "Failure" Then
        Set ebayParseXML = Nothing
    Else
        Set ebayParseXML = parser
    End If
End Function

Private Function ebayDoXMLRequest(xmlin As String, callName As String) As String

    
    Dim req As MSXML2.XMLHTTP
    Set req = New MSXML2.XMLHTTP
    req.Open "POST", "https://api.ebay.com/ws/api.dll", False
    req.setRequestHeader "content-type", "application/x-www-form-urlencoded"
    req.setRequestHeader "user-agent", "ToolsPlus EBay Interface v0.1"
    req.setRequestHeader "X-EBAY-API-COMPATIBILITY-LEVEL", "791"
    req.setRequestHeader "X-EBAY-API-DEV-NAME", "18b4e21d-d682-4afe-b8d4-bb14c5cf62ad"
    req.setRequestHeader "X-EBAY-API-APP-NAME", "ToolsPlu-e1bd-4921-ac6b-14497e74488a"
    req.setRequestHeader "X-EBAY-API-CERT-NAME", "e54d8580-281d-4a70-89ea-884c77e3704c"
    req.setRequestHeader "X-EBAY-API-SITEID", "0"
    req.setRequestHeader "X-EBAY-API-CALL-NAME", callName
    
    'Dim filefree As Integer
    'filefree = FreeFile
    'Open "c:\users\tomwestbrook\desktop\ebay-specifics.txt" For Output As #filefree
    'Print #filefree, xmlin
    'Close #filefree
    req.Send xmlin
    
    While req.ReadyState <> 4
        Sleep 100
        DoEvents
    Wend
    If req.status = 200 Then
        ebayDoXMLRequest = req.responseText
    ElseIf req.status = 503 And CBool(InStr(req.responseText, "This application is not currently available")) Then
        'no error email
        ebayDoXMLRequest = ""
    Else
        SendEmailTo "brian@tools-plus.com", "ebay offload error - bad xml request/response", req.StatusText & vbCrLf & vbCrLf & xmlin & vbCrLf & vbCrLf & req.responseText
        ebayDoXMLRequest = ""
    End If
    Set req = Nothing
End Function

Private Function ebayGenerateAddFixedPriceItemXML(def As Dictionary) As String
    Dim xmlstr As String
    xmlstr = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf & _
             "<AddFixedPriceItemRequest xmlns=""urn:ebay:apis:eBLBaseComponents"">" & vbCrLf & _
             ebayGenerateCredentialsXML() & _
             " <Item>" & vbCrLf
    
    'full item def
    xmlstr = xmlstr & ebayGenerateItemDetailXML(def, "Add")
    
    'done
    xmlstr = xmlstr & " </Item>" & vbCrLf & _
                      "</AddFixedPriceItemRequest>"
    
    ebayGenerateAddFixedPriceItemXML = xmlstr
End Function

Private Function ebayGenerateRelistFixedPriceItemXML(def As Dictionary) As String
    Dim xmlstr As String
    xmlstr = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf & _
             "<RelistFixedPriceItemRequest xmlns=""urn:ebay:apis:eBLBaseComponents"">" & vbCrLf & _
             ebayGenerateCredentialsXML() & _
             " <Item>" & vbCrLf
    'item ebay id
    xmlstr = xmlstr & ebayGenerateItemIDXML(def)
    
    'full item def
    xmlstr = xmlstr & ebayGenerateItemDetailXML(def, "Relist")
    
    'done
    xmlstr = xmlstr & " </Item>" & vbCrLf & _
                      "</RelistFixedPriceItemRequest>"
    
    ebayGenerateRelistFixedPriceItemXML = xmlstr
End Function

Private Function ebayGenerateReviseFixedPriceItemXML(def As Dictionary) As String
'MsgBox def("ItemID")
'MsgBox def.item("ItemID")
    Dim xmlstr As String
    xmlstr = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf & _
             "<ReviseFixedPriceItemRequest xmlns=""urn:ebay:apis:eBLBaseComponents"">" & vbCrLf & _
             ebayGenerateCredentialsXML() & _
             " <Item>" & vbCrLf
    'item ebay id
    xmlstr = xmlstr & ebayGenerateItemIDXML(def)
    
    'full item def
    xmlstr = xmlstr & ebayGenerateItemDetailXML(def, "Revise")
    
    'done
    xmlstr = xmlstr & " </Item>" & vbCrLf & _
                      "</ReviseFixedPriceItemRequest>"
    
    'Dim filefree As Integer
    'filefree = FreeFile
    'Open "c:\users\tomwestbrook\desktop\ebay-specifics3.txt" For Output As #filefree
    'Print #filefree, xmlstr
    'Close #filefree
    
    
    ebayGenerateReviseFixedPriceItemXML = xmlstr
End Function

Private Function ebayGenerateEndFixedPriceItemXML(def As Dictionary) As String
    Dim xmlstr As String
    xmlstr = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf & _
             "<EndFixedPriceItemRequest xmlns=""urn:ebay:apis:eBLBaseComponents"">" & vbCrLf & _
             ebayGenerateCredentialsXML()
    'item ebay id
    xmlstr = xmlstr & ebayGenerateItemIDXML(def, " ")
    'end code
    xmlstr = xmlstr & " <EndingReason>NotAvailable</EndingReason>" & vbCrLf
    
    xmlstr = xmlstr & "</EndFixedPriceItemRequest>" & vbCrLf
    
    ebayGenerateEndFixedPriceItemXML = xmlstr
End Function

Private Function ebayGenerateItemDetailXML(def As Dictionary, revisionType As String) As String
'MsgBox ebayGenerateItemSpecifics(def, revisionType)
    Dim xmlstr As String
    xmlstr = ebayGenerateProductDetailsXML(def, revisionType) & _
             ebayGenerateCategoryXML(def, revisionType) & _
             ebayGenerateConditionXML(def, revisionType) & _
             ebayGenerateItemSpecifics(def, revisionType) & _
             ebayGenerateTitleXML(def, revisionType) & _
             ebayGenerateDescriptionXML(def, revisionType) & _
             ebayGenerateQuantityXML(def, revisionType) & _
             ebayGeneratePriceXML(def, revisionType) & _
             ebayGenerateShippingXML(def, revisionType) & _
             ebayGeneratePictureXML(def, revisionType) & _
             ebayGenerateDurationXML(def, revisionType) & _
             ebayGenerateLocationXML(def, revisionType) & _
             ebayGeneratePaymentXML(def, revisionType) & _
             ebayGenerateReturnsXML(def, revisionType) & _
             ebayGenerateBuyerRequirementsXML(def, revisionType)
    ebayGenerateItemDetailXML = xmlstr
End Function

Private Function ebayGenerateCredentialsXML(Optional indent As String = " ") As String
    Dim xmlstr As String
    'xmlstr = indent & "<RequesterCredentials>" & vbCrLf & _
    '         indent & " <eBayAuthToken>AgAAAA**AQAAAA**aAAAAA**6nIIUA**nY+sHZ2PrBmdj6wVnY+sEZ2PrA2dj6AFkoCgDJOKogudj6x9nY+seQ**+4cBAA**AAMAAA**vywJnU8B6ReHO7tLCoiXHy/Ak8pbplIIfesmwW1AwuiOtXc3ZQYp7srOatwGmC3ILKxBFe/osozyrDIYsvKVpUMn32TDL43I6OyQN9WH5dCcS4Y1/0fiwtbdttyNbLbCCSfzOuRmt09zh9qd16zwf5ny+gxF/x4AnXa5IbzfRXeNQospUA8pRC5sBPhQea05JQSH/dGYXT6BTmQDDkK+7BxiSkzRraiRrp6tebrnRyhMQ96AFxO67arRP1tLokMMxoCWgvJvOo8vhTGlBnZSYeQCB7TnRoWPUZYAUntsZh4VQOXdFDRqsbNNfIm0EqOE7fDQWzYNmRMQWtZI3sjwhepF5eQvr4PxglfDKN8AF4FD6r4u4GsZ7DmGxy8QKrhnU7qWYQElIkxxc+iQDdVnrMRQ0WC2zYFX3Vx4zV2ipuiF2eVlhsxXSuodWP8ljR6jgYrpuKvF8IPE1xig1Gvj9lL8sXX2ChJjMmSajb7Rsp7iPYgrFXXm+OHVA2z7ZYBMbWdrH+Okl7G8thN0Y1mHHFPCGZFeq1nM8rDU9CukyT5o5PulOWxIWh462Eluu5iUzhuLYk3rMKU4EU69aZ6ajeZk8PVzEAKTZZem4ysrp4FHc21+HqCQWyatdeL9z85NLzo/k3B2dMyYgVHlfzVmmZiVnFVPgEq0kr24XbG3LhBJFF7EUMyS1t6FmOc4UzXn2B6P0wwOgP7r8IzmuJv+uQ4IVSnv6cWYg2slAOuqXI4HINrs7d1iRJAwg0KCfyr2</eBayAuthToken>" & vbCrLf & _
    '         indent & "</RequesterCredentials>" & vbCrLf
    xmlstr = indent & "<RequesterCredentials>" & vbCrLf & _
             indent & " <eBayAuthToken>AgAAAA**AQAAAA**aAAAAA**4l7QUg**nY+sHZ2PrBmdj6wVnY+sEZ2PrA2dj6AFkoCgDJOKogudj6x9nY+seQ**+4cBAA**AAMAAA**+9vd97XY1RzNQaC3Ny/35At4dmh29/7mJimJm7iynlxb9R5IPHYGl1/ekf0mmYQ/XivwI8VBQlkPWVFVi8w6AIXNpkZsCBUnsFYSmShD2oeRZO1UpMQnjKCv3VmYXMrb+LO9NZAHh0NwgdC97zVbS3fsJIMUll1O1ct4IcibXDQ+LunjRpiHW2j7kEILWR7j7e2nxLRNG7lSTNTI5WnvsXcs+CgilfGnli358i+KaWNN6lRDPOP8PrHaX1yjhJSQZ3ZC0mjf6oEyYEOl0O9Igmf0U4mfP980iRYWR9kInH8l2wEuKtyrGOx6zX4Pgrw//kWvOgxKng9rzE3X3+UtP4lTGNLGbSfQ0AmL1a+fQ0wmHBYSsi3r1a/PiifWEHZyoySEKw2mUgB9R3dUMQk+zTsRsB4HRDM3y6Ygapa1MAOYjqPTJCYhns+W2MPuXYtDcdKG46exHx2h1XjHOkbWby2QPlrDcOhXMdSPeE6c5y7KB0RPNp5VtHgkxZVsseWUXMZtowJ05Glv+cfhVMveJPgHkbNU7Ql0aW5hgthWTjD1xFTeF3v19BbnwVhvq5aC1uJuYEKc06YEb9qJ/AZ/Cj8X/4arrVzIBvY7ri0famvgvQAXuOo2DLA8ElVyL/c5XdRJrVfBJ58gvY0SCm/pQZdRmsPWnlGRQzJLxfSLwdXdlbjef8DxTu7+8FMB43ScCZGcI1lkgbuhDP4qC1ywZ/SYpO+N0/r81VSQEj8sYCLd1W+S3KQEybOtkKd2XKJa</eBayAuthToken>" & vbCrLf & _
             indent & "</RequesterCredentials>" & vbCrLf
    ebayGenerateCredentialsXML = xmlstr
End Function

Private Function ebayGenerateItemIDXML(def As Dictionary, Optional indent As String = "  ") As String
    Dim xmlstr As String
    'MsgBox "ID: " & def("ItemID")
    'MsgBox "Condition: " & def("ConditionID")
    
    If (IsEmpty(def.item("ItemID")) = True) Then
        'wtf man this is so reduntantly stupid. why do i have to do this
        Dim id As ADODB.Recordset
        Set id = DB.retrieve("SELECT EBayItemID FROM vWebOffloadEBay WHERE ItemNumber='" & def("ItemNumber") & "'")
        If id.RecordCount > 0 Then
            xmlstr = indent & "<ItemID>" & id("EBayItemID") & "</ItemID>" & vbCrLf
        Else
            xmlstr = indent & "<ItemID></ItemID>" & vbCrLf
        End If
        
    Else
    xmlstr = indent & "<ItemID>" & def.item("ItemID") & "</ItemID>" & vbCrLf
    End If
    ebayGenerateItemIDXML = xmlstr
End Function

Private Function ebayGenerateProductDetailsXML(def As Dictionary, revisionType As String, Optional indent As String = "  ") As String
    Dim xmlstr As String
    
    'EVERYTHING - BRAND/MPN, UPC, SKU
    ' includes ListIfNoProduct to turn error into warning
'    xmlstr = indent & "<ProductListingDetails>" & vbCrLf & _
'             indent & " <BrandMPN>" & vbCrLf & _
'             indent & "  <Brand>" & def.item("Brand") & "</Brand>" & vbCrLf & _
'             indent & "  <MPN>" & def.item("ModelNo") & "</MPN>" & vbCrLf & _
'             indent & " </BrandMPN>" & vbCrLf
'    If def.item("UPC") <> "" Then
'        Dim upcTag As String
'        upcTag = IIf(Len(def.item("UPC")) = 12, "UPC", "EAN")
'        xmlstr = xmlstr & indent & " <" & upcTag & ">" & def.item("UPC") & "</" & upcTag & ">" & vbCrLf
'    End If
'    xmlstr = xmlstr & indent & " <ListIfNoProduct>true</ListIfNoProduct>" & vbCrLf
'    xmlstr = xmlstr & indent & "</ProductListingDetails>" & vbCrLf
'    xmlstr = xmlstr & indent & "<InventoryTrackingMethod>SKU</InventoryTrackingMethod>" & vbCrLf & _
'                      indent & "<SKU>" & def.item("ItemNumber") & "</SKU>" & vbCrLf
    
    'BRAND/MPN, SKIP UPC, SKU
'    xmlstr = indent & "<ProductListingDetails>" & vbCrLf & _
'             indent & " <BrandMPN>" & vbCrLf & _
'             indent & "  <Brand>" & def.item("Brand") & "</Brand>" & vbCrLf & _
'             indent & "  <MPN>" & def.item("ModelNo") & "</MPN>" & vbCrLf & _
'             indent & " </BrandMPN>" & vbCrLf
'    xmlstr = xmlstr & indent & "</ProductListingDetails>" & vbCrLf
'    xmlstr = xmlstr & indent & "<InventoryTrackingMethod>SKU</InventoryTrackingMethod>" & vbCrLf & _
'                      indent & "<SKU>" & def.item("ItemNumber") & "</SKU>" & vbCrLf
    
    'SKIP BRAND/MPN, UPC, SKU
    If def.item("UPC") <> "" And (Len(def.item("UPC")) = 12 Or Len(def.item("UPC")) = 13) And InStr(def.item("UPC"), "TP") <= 0 Then
        Dim upcTag As String
        upcTag = IIf(Len(def.item("UPC")) = 12, "UPC", "EAN")
        xmlstr = indent & "<ProductListingDetails>" & vbCrLf & _
                 indent & " <" & upcTag & ">" & def.item("UPC") & "</" & upcTag & ">" & vbCrLf & _
                 indent & "</ProductListingDetails>" & vbCrLf
    Else
        xmlstr = xmlstr & indent & "<ProductListingDetails>" & vbCrLf & _
            indent & " <" & IIf(Len(def.item("UPC")) = 12, "UPC", "EAN") & ">Does Not Apply</" & IIf(Len(def.item("UPC")) = 12, "UPC", "EAN") & ">" & vbCrLf & _
            indent & "</ProductListingDetails>" & vbCrLf
    End If
    xmlstr = xmlstr & indent & "<InventoryTrackingMethod>SKU</InventoryTrackingMethod>" & vbCrLf & _
                      indent & "<SKU>" & def.item("ItemNumber") & "</SKU>" & vbCrLf
    
    'SKU ONLY
    'xmlstr = indent & "<InventoryTrackingMethod>SKU</InventoryTrackingMethod>" & vbCrLf & _
    '         indent & "<SKU>" & def.item("ItemNumber") & "</SKU>" & vbCrLf
    
    ebayGenerateProductDetailsXML = xmlstr
End Function

Private Function ebayGenerateCategoryXML(def As Dictionary, revisionType As String, Optional indent As String = "  ") As String
    Dim xmlstr As String
    Dim BusinessID As String
    Dim HomeID As String
    
    xmlstr = ""
    If def.exists("Category") Then
        
       'The below commented code should work but would require all listings to end and relist.
        'Dim SCat As ADODB.Recordset
        'Set SCat = DB.retrieve("SELECT CategoryID FROM EBayHomeGardenConversionID WHERE BusinessIndustrialID=" & def.item("Category") & " ORDER BY CategoryID DESC")
        'If SCat.RecordCount > 0 Then
        '    HomeID = SCat("CategoryID")
        'Else
        '    BusinessID = def.item("Category")
        'End If
        'If Len(HomeID) > 0 Then
        '    Dim bizID As ADODB.Recordset
        '    Set bizID = DB.retrieve("SELECT BusinessIndustrialID FROM EBayHomeGardenConversionID WHERE CategoryID=" & HomeID & " ORDER BY BusinessIndustrialID DESC")
        '    If (bizID.RecordCount > 0) Then
        '        BusinessID = bizID("BusinessIndustrialID")
        '    End If
        'Else
        '    Dim hID As ADODB.Recordset
        '    Set hID = DB.retrieve("SELECT CategoryID FROM EBayHomeGardenConversion WHERE BusinessIndustrialID=" & BusinessID & " ORDER BY CategoryID DESC")
        '    If (hID.RecordCount > 0) Then
        '        HomeID = hID("CategoryID")
        '    End If
        'End If
        'xmlstr = indent & "<PrimaryCategory><CategoryID>" & HomeID & "</CategoryID></PrimaryCategory>" & vbCrLf
        'If Len(BusinessID) > 0 Then
        '    xmlstr = xmlstr & indent & "<SecondaryCategory><CategoryID>" & BusinessID & "</CategoryID></SecondaryCategory>" & vbCrLf
        'End If
        
        
        
        Dim FirstCategory As String
        Dim SecondCategory As String
        Dim HoldAltCat As String
        Dim qry As ADODB.Recordset
        Set qry = DB.retrieve("SELECT EBayCategoryID, dbo.AlternativeCategoryID(EBayCategoryID) AS AlternateID FROM PartNumbers WHERE ItemNumber='" & def.item("ItemNumber") & "'")
        HoldAltCat = qry("AlternateID")
        If qry.RecordCount > 0 Then FirstCategory = qry("EBayCategoryID")
        
        Set qry = DB.retrieve("SELECT SecondaryCategoryID FROM EBaySecondaryCategoryOverride WHERE ItemNumber='" & def.item("ItemNumber") & "'")
        If qry.RecordCount > 0 Then
            SecondCategory = qry("SecondaryCategoryID")
        Else
            If Len(FirstCategory) > 0 Then
                SecondCategory = HoldAltCat
            End If
        End If
        
        
        xmlstr = xmlstr & indent & "<PrimaryCategory><CategoryID>" & FirstCategory & "</CategoryID></PrimaryCategory>" & vbCrLf
        If Len(SecondCategory) > 0 Then
            xmlstr = xmlstr & indent & "<SecondaryCategory><CategoryID>" & SecondCategory & "</CategoryID></SecondaryCategory>" & vbCrLf
        End If
        
     '   Dim FirstCategory As String
     '   Dim SecondCategory As String
     '   Dim qryTwo As ADODB.Recordset
     '
     '   Dim qryOne As ADODB.Recordset
     '   Set qryOne = DB.retrieve("SELECT CategoryID FROM EBayHomeGardenConversionID WHERE BusinessIndustrialID=" & def("Category") & " ORDER BY CategoryID DESC")
     '   If (qryOne.RecordCount > 0) Then
     '   'first cat is home & garden
     '   FirstCategory = qryOne("CategoryID")
    '        Set qryTwo = DB.retrieve("SELECT BusinessIndustrialID FROM EBayHomeGardenConversionID WHERE CategoryID=" & qryOne("CategoryID") & " ORDER BY BusinessIndustrialID DESC")
    '        If qryTwo.RecordCount > 0 Then
    '            SecondCategory = qryTwo("BusinessIndustrialID")
    '        End If
    '    Else
    '    'first cat is b & I
    '        FirstCategory = def("Category")
    '        Set qryTwo = DB.retrieve("SELECT CategoryID FROM EBayHomeGardenConversionID WHERE BusinessIndustrialID=" & def("Category"))
    '        If qryTwo.RecordCount > 0 Then
    '            SecondCategory = qryTwo("CategoryID")
    '        End If
    '    End If
    '
    '    xmlstr = xmlstr & indent & "<PrimaryCategory><CategoryID>" & FirstCategory & "</CategoryID></PrimaryCategory>" & vbCrLf
    '
    '    'this should enable second category
    '    If Len(SecondCategory) > 0 Then
    '        xmlstr = xmlstr & indent & "<SecondaryCategory><CategoryID>" & SecondCategory & "</CategoryID></SecondaryCategory>" & vbCrLf
    '    End If
        
        
        
        
        
        
        
        
        If def.item("StoreCategory") <> 0 Then
            xmlstr = xmlstr & indent & "<Storefront><StoreCategoryID>" & def.item("StoreCategory") & "</StoreCategoryID></Storefront>" & vbCrLf
        End If
    End If
    'MsgBox xmlstr
    
    ebayGenerateCategoryXML = xmlstr
End Function

Private Function ebayGenerateConditionXML(def As Dictionary, revisionType As String, Optional indent As String = "  ") As String
    Dim cid As ADODB.Recordset
    Dim xmlstr As String
    Set cid = DB.retrieve("SELECT ItemConditionMaster.EBayConditionID FROM ItemConditionLines INNER JOIN ItemConditionMaster ON ItemConditionMaster.ConditionID=ItemConditionLines.ConditionID WHERE ItemNumber='" & def("ItemNumber") & "'")
    If (cid.RecordCount > 0) Then
        xmlstr = indent & "<ConditionID>" & cid("EBayConditionID") & "</ConditionID>" & vbCrLf
    Else
        xmlstr = indent & "<ConditionID>1000</ConditionID>" & vbCrLf
    End If
    
    ebayGenerateConditionXML = xmlstr
End Function
'Private Function ebayGenerateConditionXML(def As Dictionary, revisionType As String, Optional indent As String = "  ") As String
'    Dim xmlstr As String
'    If (IsEmpty(def.item("ConditionID")) = True) Then
'        Dim cid As ADODB.Recordset
'        Set cid = DB.retrieve("SELECT IsReconditioned FROM vWebOffloadEBay WHERE ItemNumber='" & def("ItemNumber") & "'")
'        If (cid.RecordCount > 0) Then
'            xmlstr = indent & "<ConditionID>" & IIf(CStr(CInt(cid("IsReconditioned"))) = 0, "1000", "0") & "</ConditionID>" & vbCrLf
'        Else
'            xmlstr = indent & "<ConditionID></ConditionID>" & vbCrLf
'        End If
'
'    Else
'    xmlstr = indent & "<ConditionID>" & def.item("ConditionID") & "</ConditionID>" & vbCrLf
'    End If
'    ebayGenerateConditionXML = xmlstr
'End Function

Private Function ebayGenerateTitleXML(def As Dictionary, revisionType As String, Optional indent As String = "  ") As String
    Dim xmlstr As String
    xmlstr = ""
    If def.exists("Title") Then
        xmlstr = indent & "<Title>" & Utilities.escapeXMLEntities(def.item("Title")) & "</Title>" & vbCrLf
    End If
    ebayGenerateTitleXML = xmlstr
    
     'Tom's Code to hardcode a Subtitle into the Ebay Feed. It shows warranty if list price > $300
        'Tom's Code 9-3-2014... made sql table to allow items to be subtitled if they are in the list.
        Dim sqlResults As ADODB.Recordset
        Set sqlResults = DB.retrieve("SELECT ItemNumber, SubtitleID FROM SubtitledItems WHERE ItemNumber='" & def.item("ItemNumber") & "'")
        Dim sqlSalesResults As ADODB.Recordset
        Set sqlSalesResults = DB.retrieve("SELECT TotalSalesLast365Days FROM vInventorySpreadsheet WHERE ITEMNUMBER='" & def.item("ItemNumber") & "'")
        Dim subtitleItem As Boolean
        
        'Tom 's Code to get sales count for particular item, then give it subtitle if it sold more than 50 in past year.
        If sqlResults.RecordCount < 1 Then
            If sqlSalesResults.RecordCount > 0 Then
                If sqlSalesResults("TotalSalesLast365Days") > 50 Then
                    subtitleItem = True
                End If
            End If
        
        End If
        
        
        Dim SQLResults2 As ADODB.Recordset
        
        'MsgBox xmlstr
        
        If ((def.exists("StartPrice")) And (CDec(def.item("StartPrice")) > 150)) Or (sqlResults.RecordCount > 0) Or (subtitleItem = True) Then
       
            If (sqlResults.RecordCount <> 0) Then
                If (Len(sqlResults("SubtitleID")) <= 0) Then
                    'so we have records returned, but none have subtitle text, so let's default it..
                    'xmlstr = xmlstr & "  <SubTitle>99.9%fb-Auth Seller-Warranty-Free Ship-$$ Back Guar</SubTitle>" & vbCrLf
                    Set SQLResults2 = DB.retrieve("SELECT Subtitle FROM Subtitles WHERE ID=1")
                    'MsgBox "Dbug test Subtitle: " & SQLResults2("Subtitle")
                    xmlstr = xmlstr & "  <SubTitle>" & SQLResults2("Subtitle") & "</SubTitle>" & vbCrLf
                Else
                    'so there is a specific subtitle for this item.
                    Set SQLResults2 = DB.retrieve("SELECT Subtitle FROM Subtitles WHERE ID=" & sqlResults("SubtitleID"))
                    If SQLResults2.RecordCount > 0 Then
                        xmlstr = xmlstr & "  <SubTitle>" & SQLResults2("Subtitle") & "</SubTitle>" & vbCrLf
                    Else
                        'so there was a specific subtitle, but it had been deleted. reverting this item to default sub...
                        Set SQLResults2 = DB.retrieve("SELECT Subtitle FROM Subtitles WHERE ID=1")
                        xmlstr = xmlstr & "  <SubTitle>" & SQLResults2("Subtitle") & "</SubTitle>" & vbCrLf
                        DB.Execute "UPDATE SubtitledItems SET SubtitleID=1 WHERE ItemNumber='" & def("ItemNumber") & "'"
                    End If
                End If
            Else
                'there where no records of this item, so default the sub again...
                Set SQLResults2 = DB.retrieve("SELECT Subtitle FROM Subtitles WHERE ID=1")
                If SQLResults2.RecordCount > 0 Then
                xmlstr = xmlstr & "  <SubTitle>" & SQLResults2("Subtitle") & "</SubTitle>" & vbCrLf
                
                'xmlstr = xmlstr & "  <SubTitle>99.9%fb-Auth Seller-Warranty-Free Ship-$$ Back Guar</SubTitle>" & vbCrLf
                Else
                MsgBox "This is a little awkward, but I seemed to have lost the default subtitle in the sql database. Start screaming at Tom."
                
                End If
            End If
        
        Else
        
        'xmlstr = xmlstr & "  <DeleteField>Item.SubTitle</DeleteField>" & vbCrLf
        
        End If
        
        
        
       ebayGenerateTitleXML = xmlstr
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'MsgBox xmlstr
End Function

Private Function ebayGenerateDescriptionXML(def As Dictionary, revisionType As String, Optional indent As String = "  ") As String
    Dim xmlstr As String
    xmlstr = indent & "<Description>" & vbCrLf & _
                      "<![CDATA[" & vbCrLf & _
                      def.item("Description") & vbCrLf & _
                      "]]>" & vbCrLf & _
                      indent & "</Description>" & vbCrLf
    ebayGenerateDescriptionXML = xmlstr
End Function

Private Function ebayGenerateQuantityXML(def As Dictionary, revisionType As String, Optional indent As String = "  ") As String
    Dim xmlstr As String
    
    If (IsEmpty(def.item("Quantity")) = True) Then
    
        Dim qty As ADODB.Recordset
        Set qty = DB.retrieve("SELECT (QtyOnHand - QtyOnSO) as Qty FROM vWebOffloadEBay WHERE ItemNumber='" & def("ItemNumber") & "'")
        If qty.RecordCount > 0 Then
            xmlstr = indent & "<Quantity>" & CStr(qty("QTY").value) & "</Quantity>" & vbCrLf
            
        Else
            xmlstr = indent & "<Quantity>0</Quantity>" & vbCrLf
        
        End If
    
    Else
    If (CInt(def.item("Quantity")) < 0) Then
        xmlstr = indent & "<Quantity>0</Quantity>" & vbCrLf
    Else
        xmlstr = indent & "<Quantity>" & def.item("Quantity") & "</Quantity>" & vbCrLf
    End If
    End If
    ebayGenerateQuantityXML = xmlstr
End Function

Private Function ebayGeneratePriceXML(def As Dictionary, revisionType As String, Optional indent As String = "  ") As String
    
    'This function needs reworking. Ebay told us to remove a MAPP price from an item, set the MAPP to 0.
    'This function works by checking if the item's MAPP is higher, lower, non-existing, or equal to StartPrice
    'It writes either the (MAPP * 1.05) or ListPrice
    
    Dim xmlstr As String
        
   If (IsEmpty(def.item("StartPrice")) = True) Then
    Dim sp As ADODB.Recordset
    Set sp = DB.retrieve("SELECT EBayPrice FROM vWebOffloadEBay WHERE ItemNumber ='" & def("ItemNumber") & "'")
    If sp.RecordCount > 0 Then
    xmlstr = indent & "<StartPrice>" & sp("EBayPrice") & "</StartPrice>" & vbCrLf
    Else
    xmlstr = indent & "<StartPrice></StartPrice>" & vbCrLf
    End If
   Else
    xmlstr = indent & "<StartPrice>" & def.item("StartPrice") & "</StartPrice>" & vbCrLf
   End If
    'xmlstr = indent & "<StartPrice>" & (CDec(def.item("StartPrice")) + 10) & "</StartPrice>" & vbCrLf
    
    'temp line of code to test list price change
    'xmlstr = xmlstr & "  <ListPrice>" & CStr(CDec(def.item("ListPrice"))) & "</ListPrice>" & vbCrLf
    
    
    Dim hasMAPP As Boolean, hasList As Boolean
    
    hasMAPP = CDec(def.item("MAPP")) <> 0 And CDec(def.item("MAPP")) > CDec(def.item("StartPrice"))
    hasList = CDec(def.item("ListPrice")) <> 0 And CDec(def.item("ListPrice")) > CDec(def.item("StartPrice"))
    
    'If hasMAPP Then MsgBox "HAS A MAPP PRICE"
    'If hasList Then MsgBox "HAS A LIST PRICE"
    
    If hasMAPP Or hasList Then
        xmlstr = xmlstr & indent & "<DiscountPriceInfo>" & vbCrLf
    End If
    'MsgBox "MAPP = " & CStr(CDec(def.item("MAPP")))
        'MsgBox "List = " & CStr(CDec(def.item("ListPrice")))
        'MsgBox "Start = " & CStr(CDec(def.item("StartPrice")))
    If hasMAPP Then
    'Tom's Code that basically sets MAPP price to either listprice or mapp price depending on which one is a larger value than the others
    
    'MsgBox "HAS MAPP!"
        'MsgBox CStr((CDec(def.item("MAPP")) / 1.05)) & " == " & CStr(CDec(def.item("StartPrice")))
        
        
             If CDec(def.item("MAPP") / 1.05) > CDec(def.item("StartPrice")) Then
                'MsgBox "I AM HERE"
                'MsgBox "MAPP GREATER THAN LISTPRICE"
                xmlstr = xmlstr & indent & " <OriginalRetailPrice>" & def.item("MAPP") & "</OriginalRetailPrice>" & vbCrLf
                xmlstr = xmlstr & indent & " <MinimumAdvertisedPrice>" & def.item("MAPP") & "</MinimumAdvertisedPrice>" & vbCrLf & _
                          indent & " <MinimumAdvertisedPriceExposure>PreCheckout</MinimumAdvertisedPriceExposure>" & vbCrLf
             Else
                'MsgBox "MAPP LESS THAN LISTPRICE OR EQUAL TO"
                'MsgBox "I WAS HERE"
                xmlstr = xmlstr & indent & " <OriginalRetailPrice>" & def.item("ListPrice") & "</OriginalRetailPrice>" & vbCrLf
                xmlstr = xmlstr & indent & " <MinimumAdvertisedPrice>" & def.item("StartPrice") & "</MinimumAdvertisedPrice>" & vbCrLf & _
                          indent & " <MinimumAdvertisedPriceExposure>PreCheckout</MinimumAdvertisedPriceExposure>" & vbCrLf
            End If
        
    End If
    If hasList Then
        'xmlstr = xmlstr & indent & " <OriginalRetailPrice>" & def.item("ListPrice") & "</OriginalRetailPrice>" & vbCrLf
    End If
    If hasMAPP Or hasList Then
        xmlstr = xmlstr & indent & "</DiscountPriceInfo>" & vbCrLf
    End If
    If hasMAPP = False And (revisionType = "Relist" Or revisionType = "Revise") Then
        'xmlstr = xmlstr & indent & "<DeletedField>DiscountPriceInfo.MinimumAdvertisedPrice</DeletedField>" & vbCrLf
        'xmlstr = xmlstr & indent & "<DeletedField>DiscountPriceInfo.MinimumAdvertisedPriceExposure</DeletedField>" & vbCrLf
        
        'Tom's Code Test to see if we cant bypass ebay cache... (Works!)
        'MsgBox "List Price: " & def.item("ListPrice")
        'MsgBox "Start Price: " & def.item("StartPrice")
        
        'CDec(def.item("MAPP")) < CDec(def.item("StartPrice"))
        
        xmlstr = xmlstr & indent & "<DiscountPriceInfo>" & vbCrLf
        If hasList = True Then
        'MsgBox "Setting MAPP to list price"
        ' And CDec(def.item("ListPrice")) < CDec(def.item("MAPP"))
            
                xmlstr = xmlstr & indent & " <OriginalRetailPrice>" & def.item("ListPrice") & "</OriginalRetailPrice>" & vbCrLf
                xmlstr = xmlstr & indent & " <MinimumAdvertisedPrice>" & def.item("ListPrice") & "</MinimumAdvertisedPrice>" & vbCrLf
                xmlstr = xmlstr & indent & " <MinimumAdvertisedPriceExposure>PreCheckout</MinimumAdvertisedPriceExposure>" & vbCrLf
            
        Else
        'MsgBox "Setting MAPP to start price"
        xmlstr = xmlstr & indent & " <MinimumAdvertisedPrice>" & def.item("StartPrice") & "</MinimumAdvertisedPrice>" & vbCrLf
        xmlstr = xmlstr & indent & " <MinimumAdvertisedPriceExposure>None</MinimumAdvertisedPriceExposure>" & vbCrLf
        End If
        xmlstr = xmlstr & indent & " <MinimumAdvertisedPrice>" & def.item("StartPrice") & "</MinimumAdvertisedPrice>" & vbCrLf
        
        xmlstr = xmlstr & indent & "</DiscountPriceInfo>" & vbCrLf
        ' The idea is to set the MAPP price even tho we didn't want one to the selling price in hope that it removes
        ' the strikethrough, which it does
        
        
    End If
    xmlstr = xmlstr & indent & "<BestOfferDetails><BestOfferEnabled>" & CBool(IIf(def.item("AllowBestOffer"), "true", "false")) & "</BestOfferEnabled></BestOfferDetails>" & vbCrLf
    If def.item("AllowBestOffer") And (def.item("BestOfferAutoAccept") <> 0 And def.item("BestOfferAutoAccept") < def.item("StartPrice")) Then
        xmlstr = xmlstr & indent & "<ListingDetails>" & vbCrLf
        xmlstr = xmlstr & indent & " <MinimumBestOfferPrice>" & CDec(CDec(def.item("BestOfferAutoAccept")) - 0.01) & "</MinimumBestOfferPrice>" & vbCrLf
        xmlstr = xmlstr & indent & " <BestOfferAutoAcceptPrice>" & CDec(def.item("BestOfferAutoAccept")) & "</BestOfferAutoAcceptPrice>" & vbCrLf
        xmlstr = xmlstr & indent & "</ListingDetails>" & vbCrLf
    ElseIf revisionType = "Relist" Or revisionType = "Revise" Then
        If def.item("AllowBestOffer") = False Then
            xmlstr = xmlstr & indent & "<DeletedField>ListingDetails.MinimumBestOfferPrice</DeletedField>" & vbCrLf
            xmlstr = xmlstr & indent & "<DeletedField>ListingDetails.BestOfferAutoAcceptPrice</DeletedField>" & vbCrLf
        ElseIf def.item("BestOfferAutoAccept") = 0 Then
            xmlstr = xmlstr & indent & "<DeletedField>ListingDetails.MinimumBestOfferPrice</DeletedField>" & vbCrLf
            xmlstr = xmlstr & indent & "<DeletedField>ListingDetails.BestOfferAutoAcceptPrice</DeletedField>" & vbCrLf
        End If
    End If
    ebayGeneratePriceXML = xmlstr
    
    'Dim filefree As Integer
    'filefree = FreeFile
     
    'Open "C:\users\tomwestbrook\desktop\offloadB5.txt" For Output As #filefree
    'Print #filefree, xmlstr
    'Close #filefree
    
End Function

Private Function ebayGenerateShippingXML(def As Dictionary, revisionType As String, Optional indent As String = "  ") As String
    Dim xmlstr As String
    xmlstr = indent & "<ShipToLocations>US</ShipToLocations>" & vbCrLf & _
             indent & "<ShippingDetails>" & vbCrLf & _
             indent & " <SalesTax>" & vbCrLf & _
             indent & "  <SalesTaxPercent>" & IIf(def.item("TaxExempt") = False, "6.35", "0") & "</SalesTaxPercent>" & vbCrLf & _
             indent & "  <SalesTaxState>CT</SalesTaxState>" & vbCrLf & _
             indent & "  <ShippingIncludedInTax>true</ShippingIncludedInTax>" & vbCrLf & _
             indent & " </SalesTax>" & vbCrLf & _
             indent & " <ShippingType>Flat</ShippingType>" & vbCrLf & _
             indent & " <ShippingServiceOptions>" & vbCrLf & _
             indent & "  <ShippingService>ShippingMethodStandard</ShippingService>" & vbCrLf & _
             indent & "  <ShippingServiceCost>" & def.item("ShippingCharge") & "</ShippingServiceCost>" & vbCrLf & _
             indent & "  <ShippingServiceAdditionalCost>" & def.item("ShippingCharge") & "</ShippingServiceAdditionalCost>" & vbCrLf
    If def.item("ShippingCharge") = 0 Then
        xmlstr = xmlstr & indent & "  <FreeShipping>true</FreeShipping>" & vbCrLf
    End If
    xmlstr = xmlstr & indent & " </ShippingServiceOptions>" & vbCrLf
    'xmlstr = xmlstr & indent & " <ShippingServiceOptions>" & vbCrLf & _
    '                  indent & "  <ShippingService>Pickup</ShippingService>" & vbCrLf & _
    '                  indent & "  <ShippingServiceCost>0</ShippingServiceCost>" & vbCrLf & _
    '                  indent & "  <ShippingServiceAdditionalCost>0</ShippingServiceAdditionalCost>" & vbCrLf & _
    '                  indent & " </ShippingServiceOptions>" & vbCrLf
    xmlstr = xmlstr & indent & " <ExcludeShipToLocation>Alaska/Hawaii</ExcludeShipToLocation>" & vbCrLf & _
                      indent & " <ExcludeShipToLocation>US Protectorates</ExcludeShipToLocation>" & vbCrLf & _
                      indent & " <ExcludeShipToLocation>APO/FPO</ExcludeShipToLocation>" & vbCrLf & _
                      indent & " <ExcludeShipToLocation>Africa</ExcludeShipToLocation>" & vbCrLf & _
                      indent & " <ExcludeShipToLocation>Asia</ExcludeShipToLocation>" & vbCrLf & _
                      indent & " <ExcludeShipToLocation>Central America and Caribbean</ExcludeShipToLocation>" & vbCrLf & _
                      indent & " <ExcludeShipToLocation>Europe</ExcludeShipToLocation>" & vbCrLf & _
                      indent & " <ExcludeShipToLocation>Middle East</ExcludeShipToLocation>" & vbCrLf & _
                      indent & " <ExcludeShipToLocation>North America</ExcludeShipToLocation>" & vbCrLf & _
                      indent & " <ExcludeShipToLocation>Oceania</ExcludeShipToLocation>" & vbCrLf & _
                      indent & " <ExcludeShipToLocation>Southeast Asia</ExcludeShipToLocation>" & vbCrLf & _
                      indent & " <ExcludeShipToLocation>South America</ExcludeShipToLocation>" & vbCrLf & _
                      indent & " <ExcludeShipToLocation>PO Box</ExcludeShipToLocation>" & vbCrLf & _
                      indent & " <SellerExcludeShipToLocationsPreference>false</SellerExcludeShipToLocationsPreference>" & vbCrLf
    xmlstr = xmlstr & indent & "</ShippingDetails>" & vbCrLf & _
                      indent & "<DispatchTimeMax>0</DispatchTimeMax>" & vbCrLf
    ebayGenerateShippingXML = xmlstr
End Function

Private Function ebayGeneratePictureXML(def As Dictionary, revisionType As String, Optional indent As String = "  ") As String
    Dim xmlstr As String
    If def.exists("PicURL") Then
        If def.item("PicURL") = "" Then
            xmlstr = ""
        Else
            xmlstr = indent & "<PictureDetails><PictureURL>" & def.item("PicURL") & "</PictureURL></PictureDetails>" & vbCrLf
        End If
    ElseIf def.exists("PicURLs") Then
        If UBound(def.item("PicURLs")) = -1 Then
            xmlstr = ""
        Else
            xmlstr = indent & "<PictureDetails>" & vbCrLf
            Dim i As Long
            For i = 0 To UBound(def.item("PicURLs"))
                xmlstr = xmlstr & indent & " <PictureURL>" & def.item("PicURLs")(i) & "</PictureURL>" & vbCrLf
            Next i
            xmlstr = xmlstr & indent & "</PictureDetails>" & vbCrLf
        End If
    End If
    ebayGeneratePictureXML = xmlstr
End Function

Private Function ebayGenerateDurationXML(def As Dictionary, revisionType As String, Optional indent As String = "  ") As String
    Dim xmlstr As String
    xmlstr = indent & "<ListingDuration>GTC</ListingDuration>" & vbCrLf
    ebayGenerateDurationXML = xmlstr
End Function

Private Function ebayGenerateLocationXML(def As Dictionary, revisionType As String, Optional indent As String = "  ") As String
    Dim xmlstr As String
    xmlstr = indent & "<PostalCode>" & def.item("ShipFromPostalCode") & "</PostalCode>" & vbCrLf & _
             indent & "<Country>US</Country>" & vbCrLf
    ebayGenerateLocationXML = xmlstr
End Function

Private Function ebayGeneratePaymentXML(def As Dictionary, revisionType As String, Optional indent As String = "  ") As String
    Dim xmlstr As String
    xmlstr = indent & "<PaymentMethods>PayPal</PaymentMethods>" & vbCrLf & _
             indent & "<PayPalEmailAddress>eric@tools-plus.com</PayPalEmailAddress>" & vbCrLf & _
             indent & "<AutoPay>" & IIf(def.item("StartPrice") >= 10000, "false", "true") & "</AutoPay>" & vbCrLf & _
             indent & "<Currency>USD</Currency>" & vbCrLf
    ebayGeneratePaymentXML = xmlstr
End Function

Private Function ebayGenerateReturnsXML(def As Dictionary, revisionType As String, Optional indent As String = "  ") As String
    Dim xmlstr As String
    xmlstr = indent & "<ReturnPolicy>" & vbCrLf & _
             indent & " <RefundOption>MoneyBack</RefundOption>" & vbCrLf & _
             indent & " <ReturnsAcceptedOption>ReturnsAccepted</ReturnsAcceptedOption>" & vbCrLf & _
             indent & " <ReturnsWithinOption>Days_14</ReturnsWithinOption>" & vbCrLf & _
             indent & " <ShippingCostPaidByOption>Buyer</ShippingCostPaidByOption>" & vbCrLf & _
             indent & " <RestockingFeeValueOption>Percent_15</RestockingFeeValueOption>" & vbCrLf & _
             indent & "</ReturnPolicy>" & vbCrLf
    ebayGenerateReturnsXML = xmlstr
End Function

Private Function ebayGenerateBuyerRequirementsXML(def As Dictionary, revisionType As String, Optional indent As String = "  ") As String
    Dim xmlstr As String
    xmlstr = indent & "<BuyerRequirementDetails>" & vbCrLf & _
             indent & " <LinkedPayPalAccount>true</LinkedPayPalAccount>" & vbCrLf & _
             indent & " <ShipToRegistrationCountry>true</ShipToRegistrationCountry>" & vbCrLf & _
             indent & "</BuyerRequirementDetails>" & vbCrLf
    ebayGenerateBuyerRequirementsXML = xmlstr
End Function
Private Function ebayGenerateItemSpecifics(def As Dictionary, revisionType As String, Optional indent As String = "  ") As String

    
    
    Dim getASIN As ADODB.Recordset
    Set getASIN = DB.retrieve("SELECT AmazonASIN FROM UPCtoASIN WHERE ItemNumber='" & def.item("ItemNumber") & "'")
    Dim itemASIN As String
    If getASIN.RecordCount > 0 Then
        itemASIN = getASIN("AmazonASIN")
    Else
        itemASIN = ""
    End If
    
    
    Dim getPackageQty As ADODB.Recordset
    Dim totQty As Integer
    Set getPackageQty = DB.retrieve("SELECT ComponentID,Quantity FROM InventoryComponentMap WHERE ItemNumber='" & def.item("ItemNumber") & "'")
    If getPackageQty.RecordCount > 0 Then
        Do
            totQty = totQty + CDec(getPackageQty("Quantity"))
            getPackageQty.MoveNext
        Loop Until getPackageQty.EOF = True Or getPackageQty.BOF = True
               
        
        
    End If
    
    
    'MsgBox def.Item("Brand")
    'MsgBox def.Item("UPC")
    'MsgBox def.Item("ManufFullNameWeb")
    'MsgBox itemASIN
    'MsgBox def.Item("ShippingCharge")
   Dim maxSpecs As Integer
    maxSpecs = 15
    
    Dim xmlstr As String
    xmlstr = indent & "<ItemSpecifics>" & vbCrLf & _
             indent & " <NameValueList>" & vbCrLf & _
             indent & "  <Name>Brand</Name>" & vbCrLf & _
             indent & "  <Value>" & Replace(def.item("Brand"), "&", "and") & "</Value>" & vbCrLf & _
             indent & " </NameValueList>" & vbCrLf & _
             indent & " <NameValueList>" & vbCrLf & _
             indent & "  <Name>UPC</Name>" & vbCrLf & _
             indent & "  <Value>" & IIf(def.item("UPC") = "", "Does Not Apply", def.item("UPC")) & "</Value>" & vbCrLf & _
             indent & " </NameValueList>" & vbCrLf & _
             indent & " <NameValueList>" & vbCrLf & _
             indent & "  <Name>MPN</Name>" & vbCrLf & _
             indent & "  <Value>" & def.item("ModelNo") & "</Value>" & vbCrLf & _
             indent & " </NameValueList>" & vbCrLf & _
             indent & " <NameValueList>" & vbCrLf & _
             indent & "  <Name>Free Shipping</Name>" & vbCrLf & _
             indent & "  <Value>" & IIf(def.item("ShippingCharge") > 0, "No", "Yes") & "</Value>" & vbCrLf & _
             indent & " </NameValueList>" & vbCrLf
             maxSpecs = 11
             If itemASIN <> "" And itemASIN <> "UPCNOTFOUND" Then

                xmlstr = xmlstr & indent & " <NameValueList>" & vbCrLf & _
                                    indent & "  <Name>ASIN</Name>" & vbCrLf & _
                                    indent & "  <Value>" & itemASIN & "</Value>" & vbCrLf & _
                                    indent & " </NameValueList>" & vbCrLf
                maxSpecs = maxSpecs - 1
             End If
             
             
             If totQty > 0 Then

                xmlstr = xmlstr & indent & " <NameValueList>" & vbCrLf & _
                                    indent & "  <Name>Package Quantity</Name>" & vbCrLf & _
                                    indent & "  <Value>" & CStr(totQty) & "</Value>" & vbCrLf & _
                                    indent & " </NameValueList>" & vbCrLf
                maxSpecs = maxSpecs - 1
             End If
             
             
             'now for some dynamicness... lets pull techspecs, filter em, and put em up here too!
             'p.s. I HATE VB6

             Dim techSpecs As ADODB.Recordset
             Set techSpecs = DB.retrieve("SELECT TechSpecsHTML FROM PartNumbers WHERE ItemNumber='" & def.item("ItemNumber") & "'")
             If techSpecs.RecordCount > 0 Then
                Dim specArray() As String
                If IsNull(techSpecs("TechSpecsHTML")) = False Then
                If InStr(CStr(techSpecs("TechSpecsHTML")), "<li>") > 0 Then
                specArray = Split(techSpecs("TechSpecsHTML"), "<li>")
                Dim str As Variant
                For Each str In specArray
                    If maxSpecs > 0 Then
                        If InStr(str, "<ul>") <= 0 Then
                            'Dim keyFull As String
                            'keyFull = Split(Split(CStr(str), "<li>")(1), "<")(0)
                            Dim keyName As String
                            Dim keyValue As String
                            If InStr(CStr(str), ":") > 0 Then
                                maxSpecs = maxSpecs - 1
                                keyName = Split(CStr(str), ":")(0)
                                
                                keyValue = Split(CStr(str), ":")(1)
                                keyValue = Replace(keyValue, "</ul>", "")
                                keyValue = Replace(keyValue, vbCrLf, "")
                                keyValue = Replace(keyValue, "&mdash;", "-")
                                keyValue = Replace(keyValue, "&deg;", "º")
                                keyValue = Replace(keyValue, "&plusmn;", "+/-")
                                keyValue = Replace(keyValue, "</li>", "")
                                keyValue = Replace(keyValue, "<small><sup>", " ")
                                keyValue = Replace(keyValue, "</sup></small>", "")
                                keyValue = Replace(keyValue, "<small><sub>", "")
                                keyValue = Replace(keyValue, "</sub></small>", "")
                                keyValue = Replace(keyValue, "&reg;", "®")
                                keyValue = Replace(keyValue, "&copy;", "©")
                                keyValue = Replace(keyValue, "&", "and")
                                keyValue = Replace(keyValue, "<", "less than ")
                                keyValue = Replace(keyValue, ">", "greater than ")
                                keyValue = Replace(keyValue, "<br>", "")
                                keyValue = Trim(keyValue)
                                keyName = Replace(keyName, ",", "")
                                keyName = Replace(keyName, "</ul>", "")
                                keyName = Replace(keyName, vbCrLf, "")
                                keyName = Replace(keyName, "&mdash;", "-")
                                keyName = Replace(keyName, "&deg;", "º")
                                keyName = Replace(keyName, "&plusmn;", "+/-")
                                keyName = Replace(keyName, "</li>", "")
                                keyName = Replace(keyName, "<small><sup>", " ")
                                keyName = Replace(keyName, "</sup></small>", "")
                                keyName = Replace(keyName, "<small><sub>", "")
                                keyName = Replace(keyName, "</sub></small>", "")
                                keyName = Replace(keyName, "&reg;", "®")
                                keyName = Replace(keyName, "&copy;", "©")
                                keyName = Replace(keyName, "&", "and")
                                keyName = Replace(keyName, "<br>", "")

                                keyName = Trim(keyName)
                                
                                If Len(keyValue) <= 50 Then
                                xmlstr = xmlstr & indent & " <NameValueList>" & vbCrLf & _
                                        indent & "  <Name>" & keyName & "</Name>" & vbCrLf & _
                                        indent & "  <Value>" & keyValue & "</Value>" & vbCrLf & _
                                        indent & " </NameValueList>" & vbCrLf
                                End If
                            End If
                           
                        End If
                    End If
                Next str
                End If
                End If
             End If
             
             'Dim filefree As Integer
             'filefree = FreeFile
             'Open "c:\users\tomwestbrook\desktop\ebayspecificstest.txt" For Output As #filefree
             'Print #filefree, xmlstr
            ' Close #filefree
             
             xmlstr = xmlstr & indent & "</ItemSpecifics>" & vbCrLf

    
    ebayGenerateItemSpecifics = xmlstr

End Function

Public Function EBayGenerateHTMLTemplate(itemBodyContent As String, rst As ADODB.Recordset) As String
    'rst is a select * from vweboffloadebay
    Dim retval As String
    retval = "<!DOCTYPE html>" & vbCrLf & _
             "<html>" & vbCrLf
    retval = retval & _
             " <head>" & vbCrLf & _
             "  <meta charset=""UTF-8"">" & vbCrLf & _
             "  <title>" & TrimWhitespace(Nz(rst("ManufFullNameWeb")) & " " & Nz(rst("Desc2")) & " " & Nz(rst("Desc3"))) & "</title>" & vbCrLf & _
             "  <link rel=""stylesheet"" href=""http://lib.store.yahoo.net/lib/toolsplus/ebay-template.css"">" & vbCrLf & _
             " </head>" & vbCrLf
    retval = retval & _
             " <body>" & vbCrLf
    retval = retval & _
             "  <div id=""tp-container"">" & vbCrLf & _
             "   <div id=""tp-header"">" & vbCrLf
    If rst("EBayAllowBestOffer") = True Then
        If rst("EBayPrice") <= 75 Then
            retval = retval & "    <div id=""tp-best-offer-teaser""><img src=""http://p1.hostingprod.com/@tools-plus.com/images/make-offer-header.png"" alt=""Make an Offer/Suggestion: Offers for 2 or more are likely to be accepted""></div>" & vbCrLf
        End If
    End If
    retval = retval & _
             "    <img id=""tp-main-logo"" src=""http://p1.hostingprod.com/@tools-plus.com/ebay/tools-plus-outlet-571.png"" alt=""Tools Plus Outlet"">" & vbCrLf & _
             "    <img id=""tp-auth-logo"" src=""http://p1.hostingprod.com/@tools-plus.com/ebay/auth-logos/" & LCase(Left(rst("ItemNumber"), 3)) & ".jpg"" alt=""" & Nz(rst("ManufFullNameWeb")) & " Authorized Retailer"">" & vbCrLf & _
             "    <div style=""clear:both;""></div>" & vbCrLf & _
             "   </div>" & vbCrLf
    retval = retval & _
             "   <div id=""tp-top-nav"">" & vbCrLf & _
             "    <ul>" & vbCrLf & _
             "     <li><a href=""http://stores.ebay.com/toolsplusoutlet"">Store Home</a></li>" & vbCrLf & _
             "     <li>" & vbCrLf & _
             "      <a>Shop By Brands</a>" & vbCrLf & _
             "      <ul>" & vbCrLf & _
             "       <li><a href=""http://stores.ebay.com/Tools-Plus-Outlet/DeWalt-/_i.html?_fsub=3375286017"">DeWalt</a></li>" & vbCrLf & _
             "       <li><a href=""http://stores.ebay.com/Tools-Plus-Outlet/Milwaukee-/_i.html?_fsub=3375285017"">Milwaukee</a></li>" & vbCrLf & _
             "       <li><a href=""http://stores.ebay.com/Tools-Plus-Outlet/Makita-/_i.html?_fsub=3375287017"">Makita</a></li>" & vbCrLf & _
             "       <li><a href=""http://stores.ebay.com/Tools-Plus-Outlet/Bosch-/_i.html?_fsub=3375288017"">Bosch</a></li>" & vbCrLf & _
             "       <li><a href=""http://stores.ebay.com/Tools-Plus-Outlet/Hitachi-Tools-/_i.html?_fsub=38994860172"">Hitachi</a></li>" & vbCrLf & _
             "       <li><a href=""http://stores.ebay.com/Tools-Plus-Outlet/Bostitch-/_i.html?_fsub=3375289017"">Bostitch</a></li>" & vbCrLf & _
             "       <li><a href=""http://stores.ebay.com/Tools-Plus-Outlet/Porter-Cable-/_i.html?_fsub=3535203017"">Porter Cable</a></li>" & vbCrLf & _
             "       <li><a href=""http://stores.ebay.com/Tools-Plus-Outlet"">All Products</a></li>" & vbCrLf & _
             "      </ul>" & vbCrLf & _
             "     </li>" & vbCrLf & _
             "     <li><a href=""http://cgi3.ebay.com/ws/eBayISAPI.dll?ViewUserPage&amp;userid=tools-plus-outlet"">About Us</a></li>" & vbCrLf & _
             "     <li><a href=""http://feedback.ebay.com/ws/eBayISAPI.dll?ViewFeedback2&amp;userid=tools-plus-outlet&amp;ftab=AllFeedback"">Seller Feedback</a></li>" & vbCrLf & _
             "     <li><a href=""http://my.ebay.com/ws/eBayISAPI.dll?AcceptSavedSeller&amp;mode=0&amp;preference=0&amp;sellerid=tools-plus-outlet"">Add to Favorite Sellers</a></li>" & vbCrLf & _
             "    </ul>" & vbCrLf & _
             "   </div>" & vbCrLf
    retval = retval & _
             "   <div id=""tp-content"">" & vbCrLf & _
             itemBodyContent & vbCrLf & _
             "    <hr>" & vbCrLf & _
             "    <div id=""tp-policies"">" & vbCrLf & _
             "     " & IIf(rst("DropshipOnly"), "<img id=""tp-policy-dropship"" src=""http://p1.hostingprod.com/@tools-plus.com/ebay/dropship-store-467.png"" alt=""Dropship Item Policy"">", "<!-- stock item, skipping d/s policy -->") & vbCrLf & _
             "     <img class=""tp-policy"" src=""http://p1.hostingprod.com/@tools-plus.com/ebay/about-our-products-300.png"" alt=""About Our Products"">" & vbCrLf & _
             "     <img class=""tp-policy"" src=""http://p1.hostingprod.com/@tools-plus.com/ebay/feedback-300.png"" alt=""Feedback"">" & vbCrLf & _
             "     <img class=""tp-policy"" src=""http://p1.hostingprod.com/@tools-plus.com/ebay/shipping-300.png"" alt=""Shipping"">" & vbCrLf & _
             "     <img class=""tp-policy"" src=""http://p1.hostingprod.com/@tools-plus.com/ebay/product-images-300.png"" alt=""Product Images"">" & vbCrLf & _
             "     <img class=""tp-policy"" src=""http://p1.hostingprod.com/@tools-plus.com/ebay/availability-300.png"" alt=""Availability"">" & vbCrLf & _
             "     <img class=""tp-policy"" src=""http://p1.hostingprod.com/@tools-plus.com/ebay/before-buying-300.png"" alt=""Before Buying"">" & vbCrLf & _
             "     <img class=""tp-policy"" src=""http://p1.hostingprod.com/@tools-plus.com/ebay/returns-600.png"" alt=""Returns"">" & vbCrLf & _
             "    </div>" & vbCrLf & _
             "   </div>" & vbCrLf
    retval = retval & _
             "  </div>" & vbCrLf & _
             " </body>" & vbCrLf & _
             "</html>" & vbCrLf
    
    EBayGenerateHTMLTemplate = retval
End Function





















'tom's test code to determine that ebay wont let us revise titles/subtitles...
Public Function EBayAPIReviseItem(def As Dictionary) As Boolean
    ebaySellingLimitReachedFlag = False
    ebayAPICallLimitReachedFlag = False
    
    Dim xmlin As String
    xmlin = ebayGenerateReviseFixedPriceItemXML(def)
    
    'xmlin = Replace(xmlin, "ReviseFixedPriceItem", "ReviseItem")
    
    
    'xmlin = Replace(xmlin, "ReviseFixedPriceItem", "ReviseItem")
    'Dim filefree As Long
    'filefree = FreeFile
    'Open "c:\users\tomwestbrook\desktop\xmlIN.txt" For Output As #filefree
    '    Print #filefree, xmlin
    'Close #filefree
    
    
    
    Dim xmlout As String
    'xmlout = ebayDoXMLRequest(xmlin, "ReviseFixedPriceItem")
    xmlout = ebayDoXMLRequest(xmlin, "ReviseFixedPriceItem")
    
   '     Dim filefree2 As Long
   ' filefree2 = FreeFile
   ' Open "c:\users\tomwestbrook\desktop\xmlOUT.txt" For Output As #filefree2
   '     Print #filefree2, xmlout
   ' Close #filefree2

    If xmlout = "" Then
        'error
        EBayAPIReviseItem = False
        Exit Function
    End If
    
    Dim root As MSXML2.DOMDocument
    Set root = ebayParseXML(xmlin, xmlout, def.item("ItemNumber"), "ReviseItemResponse")
    
    If Not root Is Nothing Then
        'make sure it's marked as up, then...?
        'qty update, item id stays the same
        ebayUpdateToRevised def.item("ItemNumber"), def.item("ItemID"), def.item("Quantity")
        EBayAPIReviseItem = True
    Else
        EBayAPIReviseItem = False
    End If
End Function






Public Function EBayAPIAddItem(def As Dictionary) As Boolean
    ebaySellingLimitReachedFlag = False
    ebayAPICallLimitReachedFlag = False
    
    Dim xmlin As String
    xmlin = ebayGenerateAddFixedPriceItemXML(def)
    
    xmlin = Replace(xmlin, "AddFixedPriceItem", "AddItem")
    
    Dim xmlout As String
    xmlout = ebayDoXMLRequest(xmlin, "AddItem")
    
    'Dim filed As Long
    'filed = FreeFile
    'Open "c:\users\tomwestbrook\desktop\wtfman.txt" For Output As #filed
    'Print #filed, xmlout
    'Close #filed
    'Dim filefree As Long
    'filefree = FreeFile
    'Open "c:\users\tomwestbrook\desktop\ebay-second-category.txt" For Output As #filefree
    'Print #filefree, xmlin & vbCrLf & vbCrLf & xmlout
    'Close #filefree
    
    
    If xmlout = "" Then
        'error
        EBayAPIAddItem = False
        Exit Function
    End If
    
    Dim root As MSXML2.DOMDocument
    Set root = ebayParseXML(xmlin, xmlout, def.item("ItemNumber"), "AddItemResponse")
    
    If Not root Is Nothing Then
        'here's the actual work, just grab the id and update with qty
        Dim id As String
        id = root.selectSingleNode("/AddItemResponse/ItemID").Text
        ebayUpdateToPublished def.item("ItemNumber"), id, def.item("Quantity")
        EBayAPIAddItem = True
    Else
        'already emailed, nothing to do
        EBayAPIAddItem = False
    End If
End Function








Public Function EBayAPIEndItem(def As Dictionary) As Boolean
    ebaySellingLimitReachedFlag = False
    ebayAPICallLimitReachedFlag = False
    
    Dim xmlin As String
    xmlin = ebayGenerateEndFixedPriceItemXML(def)
    xmlin = Replace(xmlin, "EndFixedPriceItem", "EndItem")
    
    Dim xmlout As String
    xmlout = ebayDoXMLRequest(xmlin, "EndItem")
    
    If xmlout = "" Then
        'error
        EBayAPIEndItem = False
        Exit Function
    End If
    
    Dim root As MSXML2.DOMDocument
    Set root = ebayParseXML(xmlin, xmlout, def.item("ItemNumber"), "EndItemResponse")
    'if there is a db inconsistency with ebay (they don't always send the
    'notification that an item sold), then we might get an error response
    'of "auction has already ended", this is handled in the parse and
    'returned as a fake success code. if we ever do anything other than
    'just mark as delisted, this may need to be addressed.
    
    If Not root Is Nothing Then
        'just marking as down for now
        ebayUpdateToDelisted def.item("ItemNumber")
        EBayAPIEndItem = True
    Else
        EBayAPIEndItem = False
    End If
End Function




Public Function EBayAPIRelistItem(def As Dictionary) As Boolean
'   EBayAPIRelistFixedPriceItem = ebayapir
    EBayAPIRelistItem = EBayAPIAddItem(def)
    Exit Function
    
    
    
    
    
    ebaySellingLimitReachedFlag = False
    ebayAPICallLimitReachedFlag = False
    
    Dim xmlin As String
    xmlin = ebayGenerateRelistFixedPriceItemXML(def)
    Dim xmlout As String
    xmlout = ebayDoXMLRequest(xmlin, "RelistItem")
    
    If xmlout = "" Then
        'error
        EBayAPIRelistItem = False
        Exit Function
    End If
    
    Dim root As MSXML2.DOMDocument
    Set root = ebayParseXML(xmlin, xmlout, def.item("ItemNumber"), "RelistItemResponse")
    
    If Not root Is Nothing Then
        'grab id, update qty
        Dim id As String
        id = root.selectSingleNode("/RelistItemResponse/ItemID").Text
        ebayUpdateToRelisted def.item("ItemNumber"), id, def.item("Quantity")
        EBayAPIRelistItem = True
    Else
        'perhaps this item needed revising since we now use the ebay out of stock option...
        'EBayAPIRelistFixedPriceItem = EBayAPIReviseFixedPriceItem(def)
        
        EBayAPIRelistItem = False
    End If
End Function
