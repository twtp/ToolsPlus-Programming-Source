Attribute VB_Name = "JetDotCom"

Public Type JetItemInfo
     id As String 'required
    
    'vars for general info
     ProductTitle As String 'required
     JetBrowseNodeID As String
     SKU As String
     AmazonItemTypeKeyword As String
     CategoryPath As String
     SKU_Code As String 'required
     SKU_Type As String 'required
     ASIN As String
     MultipackQty As Integer 'required
     BrandName As String 'required
     Manufacturer As String 'required
     ManufacturerPartNumber As String
     ProductDescription As String
     Bullets() As String
     NumberUnitsForPricePerUnit As Single
     TypeOfUnitForPricePerUnit As String
     ShippingWeightPounds As Single
     PackageLengthInches As Single
     PackageWidthInches As Single
     PackageHeightInches As Single
     DisplayLengthInches As Single
     DisplayWidthInches As Single
     DisplayHeightInches As Single
     Prop65 As Boolean
     LegalDisclaimerDescription As String
     CPSIACautionaryStatements() As String
     CountryOfOrigin As String
     SafetyWarning As String
     StartSellingDate As String
     FulfillmentTime As Integer
     MSRP As Single
     MAPP As Currency
     MAPPImplementation As String
     ProductTaxCode As String
     NoReturnFeeAdjustment As String
     ExcludeFromFeeAdjustments As Boolean
     ShipsAlone As Boolean
     AttributeID() As Integer
     AttributeValue() As String
     AttributeValueUnit() As String
     
     'vars for photo...
     mainImageURL As String 'required
     SwatchImageURL As String
     AlternateImageSlotID() As String 'up to 8 slots 'required
     AlternateImageURL() As String 'required
     
     'inventory vars...
     FulfillmentNodeID() As String 'required
     FulfillmentNodeQuantity() As Single 'required
     
     'pricing vars...
     price As Currency 'required
     FulfillmentNodePrice() As Single 'required
     
     'relationship vars...
     Relationship As String 'required
     VariationRefinements As Single
     ChildrenSKUS() As String 'required
     
     'shipping exceptions
     ServiceLevel As String
     ShippingCarrier As String
     ShippingMethod As String
     OverrideType As String
     ShippingChargeAmount As Single
     IsShippingRestricted As Boolean
     IsShippingExclusive As Boolean
     
     
     
End Type
Dim JetIDToken As String
Const loginStr = "{""user"":""E69170C7462659AE0B026B0446F148984EEFE23E"",""pass"":""6tnEI4/Wv6GWITov0coeCWybi4RlSpJJedSPXomUdTRZ""}"
Const ScottRdNode = "40047d4d7c5f4ce6bd4fb5eb684588d8"

    
    




Public Sub JetOffloadAll()
RemoveOldItemsFromJet
If JetIDToken = "" Then SetJetCredentials

'for now if ebay wants update, so dont jet... for ease of coding atm..
'DB.Execute "UPDATE JetOffload SET OffloadRequired=1 WHERE ItemNumber IN (SELECT ItemNumber FROM vWebOffloadEBay WHERE EBayOffloadRequired=1)"


Dim jetOffloaders As ADODB.Recordset
Set jetOffloaders = DB.retrieve("SELECT ItemNumber FROM JetOffload WHERE OffloadRequired=1")
Dim FailedOffloads As String
Dim OffloadedCount As Integer
Dim recordCounter As Integer

If jetOffloaders.RecordCount > 0 Then
    Do
    
    recordCounter = recordCounter + 1
    
    '1.) Does item need offloading? Yes, the answer is already yes. So skip this step!
    '2.) Is this item supposed to be on Jet.com? Yeah, all set!

    WebOffload.lblStatusBar.Caption = "Jet.com: Listing " & jetOffloaders("ItemNumber") & " (" & CStr(recordCounter) & "/" & CStr(jetOffloaders.RecordCount) & ")"
    WebOffload.lblStatusBar.Refresh
    
    Dim success As Boolean
        success = OffloadItemToJet(jetOffloaders("ItemNumber"))
        DoEvents
        If success = False Then
            FailedOffloads = FailedOffloads & jetOffloaders("ItemNumber") & ", "
        Else
            OffloadedCount = OffloadedCount + 1
        End If
    

     jetOffloaders.MoveNext
     
   Loop Until jetOffloaders.EOF = True Or jetOffloaders.BOF = True
    If FailedOffloads <> "" Then
        MsgBox CStr(OffloadedCount) & " items where offloaded to Jet.com. However, these items failed to offload properly:" & vbCrLf & FailedOffloads
 
    Else
       MsgBox "The Jet.com offload completed sucessfully. " & CStr(OffloadedCount) & " items where offloaded."
    End If
 
Else
    MsgBox "No Items to offload at Jet.com!"
End If

End Sub

Public Function SetJetCredentials()

    
    Dim req As MSXML2.XMLHTTP
    Set req = New MSXML2.XMLHTTP
    req.Open "POST", "https://merchant-api.jet.com/api/Token", False
    'req.setRequestHeader "Host", "merchant-api.test.jet.com"
    req.setRequestHeader "Content-type", "application/json; charset=utf-8"
    req.setRequestHeader "Content-length", CStr(Len(loginStr))
    
    
    req.Send loginStr
    Dim response As String
    Dim tokenType As String
    Dim IDtoken As String
    
    response = req.responseText
    
    If InStr(response, "token_type") > 0 Then
        tokenType = Split(Split(response, "token_type"": """)(1), """")(0)
        If UCase(tokenType) = "BEARER" Then
           IDtoken = Split(Split(response, "id_token"": """)(1), """")(0)
           'MsgBox "ID TOKEN: " & IDtoken
           JetIDToken = IDtoken
           
        Else
            MsgBox "Looks like Jet.com changed their credential types. Thats just bloody terrific. Let Tom know."
            Exit Function
        End If
    
    End If
    
End Function
Public Function RemoveOldItemsFromJet() As Boolean
If JetIDToken = "" Then SetJetCredentials

Dim req As MSXML2.ServerXMLHTTP
Set req = New MSXML2.ServerXMLHTTP
req.Open "GET", "https://merchant-api.jet.com/api/merchant-skus"
req.setRequestHeader "Content-type", "application/json"
req.setRequestHeader "Authorization", "Bearer " & JetIDToken
req.Send
'req.Send ""


Dim response As String
response = req.responseText
If InStr(response, "[") > 0 Then
    response = Split(Split(req.responseText, "[")(1), "]")(0)
    Dim count As Integer
    count = UBound(Split(response, vbCrLf))
   Dim Results() As String
   If reponse <> "" Then
   Results = Split(response, vbCrLf)
   Dim xcount As Integer
   
   Do
   xcount = xcount + 1
   
    Results(xcount) = Split(Split(response, "-skus/")(xcount), """")(0)
    
   
   Loop Until xcount = count - 1
   
   End If
   
Else
'error perhaps?
End If
'xcount = 0
'Do
'xcount = xcount + 1
'MsgBox results(xcount)
'Loop Until xcount = UBound(results) - 1

'now that we have a list of all jet.com items, we must see if any are flagged ToBeDeleted
Dim deleters As ADODB.Recordset
Set deleters = DB.retrieve("SELECT ID, ItemNumber FROM JetOffload WHERE ToBeDeleted=1")
Do
If deleters.EOF = False And deleters.BOF = False Then
    Dim ItemNumber As String
    ItemNumber = deleters("ItemNumber")
    
    
    Dim delItem As MSXML2.XMLHTTP
    Set delItem = New MSXML2.XMLHTTP
    delItem.Open "PUT", "https://merchant-api.jet.com/api/merchant-skus/" & ItemNumber & "/inventory"
    delItem.setRequestHeader "Content-type", "application/json"
    delItem.setRequestHeader "Authorization", "Bearer " & JetIDToken
    delItem.Send "{""fulfillment_nodes"": []}"
    On Error Resume Next
    Do
    DoEvents
    Loop Until delItem.status > 200
    On Error GoTo 0
    DB.Execute "DELETE FROM JetOffload WHERE ID=" & deleters("ID")
    deleters.MoveNext
End If
Loop Until deleters.BOF = True Or deleters.EOF = True




End Function
Public Function OffloadItemToJet(ItemNumber As String) As Boolean
       
If JetIDToken = "" Then SetJetCredentials


Dim mainJSON As String
mainJSON = GetItemInfo(ItemNumber)
Dim imageJSON As String
imageJSON = GetImageInfo(ItemNumber)
Dim inventoryJSON As String
inventoryJSON = GetInventoryInfo(ItemNumber)
Dim priceJSON As String
priceJSON = GetPriceInfo(ItemNumber)
'Dim filefree As Long
'filefree = FreeFile
'Open "C:\users\tomwestbrook\desktop\JSON_OutputCheck.txt" For Output As #filefree
'Print #filefree, imageJSON
'Close #filefree

'Six steps to freedom!

'Step 1. Upload Main Details of Item to Jet.com. If error then abort.
Dim mainDetailsSuccess As Boolean
mainDetailsSuccess = OffloadMainJetData(mainJSON, ItemNumber)
If mainDetailsSuccess = False Then
    OffloadItemToJet = False
    Exit Function
End If

'Step 2. Upload Images of Item to Jet.com. If error then abort
If InStr(imageJSON, "http") > 0 Then
    Dim imageDetailsSuccess As Boolean
    imageDetailsSuccess = OffloadImageJetData(imageJSON, ItemNumber)
    If imageDetailsSuccess = False Then
    OffloadItemToJet = False
    Exit Function
    End If
Else
    OffloadItemToJet = False
    Exit Function
End If


'Step 3. Upload Inventory info to Jet.com. If error then abort
Dim inventorySuccess As Boolean
inventorySuccess = OffloadInventoryJetData(inventoryJSON, ItemNumber)
If inventorySuccess = False Then
    OffloadItemToJet = False
    Exit Function
End If
'Step 4. Upload Pricing info to Jet.com. If error then abort.
Dim priceSuccess As Boolean
inventorySuccess = OffloadPriceJetData(priceJSON, ItemNumber)
If inventorySuccess = False Then
    OffloadItemToJet = False
    Exit Function
End If
'Step 5. Upload Relationship info to Jet.com. If error then abort.
'variations and such. skipping this section for now as we probably wont use it in vb6 ever.
'Step 6. Upload Shipping Exception info to Jet.com. If error then abort.
'overrides default shipping data. skipping for now as i dont think we have this setup either.
'Step 7. All supposedly went ok, which means it's time to unflag the jet.com offload for this item! Well done.
Dim flagSuccess As Boolean
flagSuccess = DB.Execute("UPDATE JetOffload SET OffloadRequired=0 WHERE ItemNumber='" & ItemNumber & "'")
If flagSuccess = False Then
    OffloadItemToJet = False
    Exit Function
End If
    
OffloadItemToJet = True
End Function

Public Function OffloadMainJetData(JSON As String, ItemNumber As String) As Boolean
Dim itemPut As MSXML2.XMLHTTP
    Set itemPut = New MSXML2.XMLHTTP
    
    itemPut.Open "PUT", "https://merchant-api.jet.com/api/merchant-skus/" & ItemNumber, False
    itemPut.setRequestHeader "Content-type", "application/json; charset=utf-8"
    itemPut.setRequestHeader "Authorization", "Bearer " & JetIDToken
    itemPut.setRequestHeader "Content-length", CStr(Len(JSON))
    On Error Resume Next
    itemPut.Send JSON
    
    response = itemPut.responseText

    '201= created
    '204= updated
    '1223= updated (ie screws up the 204 and turns it into a 1223... so check for both just in case...
    
    If InStr(itemPut.status, "201") > 0 Or InStr(itemPut.status, "204") > 0 Or InStr(itemPut.status, "1223") > 0 Then
        OffloadMainJetData = True
    Else
        OffloadMainJetData = False
    End If
    On Error GoTo 0
End Function

Public Function OffloadImageJetData(JSON As String, ItemNumber As String) As Boolean
Dim itemPut As MSXML2.XMLHTTP
    Set itemPut = New MSXML2.XMLHTTP
    
    itemPut.Open "PUT", "https://merchant-api.jet.com/api/merchant-skus/" & ItemNumber & "/image", False
    itemPut.setRequestHeader "Content-type", "application/json; charset=utf-8"
    itemPut.setRequestHeader "Authorization", "Bearer " & JetIDToken
    itemPut.setRequestHeader "Content-length", CStr(Len(JSON))
    On Error Resume Next
    itemPut.Send JSON
    
    response = itemPut.responseText

    '201= created
    '204= updated
    '1223= updated (ie screws up the 204 and turns it into a 1223... so check for both just in case...
    'MsgBox "Image status for " & ItemNumber & ": " & itemPut.status
    If InStr(itemPut.status, "201") > 0 Or InStr(itemPut.status, "204") > 0 Or InStr(itemPut.status, "1223") > 0 Then
        OffloadImageJetData = True
    Else
        OffloadImageJetData = False
    End If


End Function
Public Function OffloadInventoryJetData(JSON As String, ItemNumber As String) As Boolean
Dim itemPut As MSXML2.XMLHTTP
    Set itemPut = New MSXML2.XMLHTTP
    
    itemPut.Open "PUT", "https://merchant-api.jet.com/api/merchant-skus/" & ItemNumber & "/inventory", False
    itemPut.setRequestHeader "Content-type", "application/json; charset=utf-8"
    itemPut.setRequestHeader "Authorization", "Bearer " & JetIDToken
    itemPut.setRequestHeader "Content-length", CStr(Len(JSON))
    On Error Resume Next
    itemPut.Send JSON
    
    response = itemPut.responseText

    '201= created
    '204= updated
    '1223= updated (ie screws up the 204 and turns it into a 1223... so check for both just in case...
    'MsgBox "Inventory status for " & ItemNumber & ": " & itemPut.status
    If InStr(itemPut.status, "201") > 0 Or InStr(itemPut.status, "204") > 0 Or InStr(itemPut.status, "1223") > 0 Then
        OffloadInventoryJetData = True
    Else
        OffloadInventoryJetData = False
    End If


End Function
Public Function OffloadPriceJetData(JSON As String, ItemNumber As String) As Boolean
Dim itemPut As MSXML2.XMLHTTP
    Set itemPut = New MSXML2.XMLHTTP
  
    itemPut.Open "PUT", "https://merchant-api.jet.com/api/merchant-skus/" & ItemNumber & "/price", False
    itemPut.setRequestHeader "Content-type", "application/json; charset=utf-8"
    itemPut.setRequestHeader "Authorization", "Bearer " & JetIDToken
    itemPut.setRequestHeader "Content-length", CStr(Len(JSON))
    On Error Resume Next
    itemPut.Send JSON
    
    response = itemPut.responseText

    '201= created
    '204= updated
    '1223= updated (ie screws up the 204 and turns it into a 1223... so check for both just in case...
    'MsgBox "Price status for " & ItemNumber & ": " & itemPut.status
    If InStr(itemPut.status, "201") > 0 Or InStr(itemPut.status, "204") > 0 Or InStr(itemPut.status, "1223") > 0 Then
        OffloadPriceJetData = True
    Else
        OffloadPriceJetData = False
    End If


End Function
Public Function GetPriceInfo(ItemNumber As String) As String
    Dim JetItem As JetItemInfo
    ReDim JetItem.FulfillmentNodeID(0) As String
    ReDim JetItem.FulfillmentNodePrice(0) As Single
    JetItem.FulfillmentNodeID(0) = ScottRdNode
    Dim priceData As ADODB.Recordset
    Set priceData = DB.retrieve("SELECT Price FROM JetOffload WHERE ItemNumber = '" & ItemNumber & "'")
    If priceData.RecordCount > 0 Then
        JetItem.FulfillmentNodePrice(0) = CDbl(priceData("Price"))
        JetItem.price = CDbl(priceData("Price"))
    End If
    GetPriceInfo = CreatePriceJSON(JetItem)

End Function

Public Function GetInventoryInfo(ItemNumber As String) As String
Dim JetItem As JetItemInfo

'need to pull nodes from jet.com in the future...
ReDim JetItem.FulfillmentNodeID(0 To 10) As String
ReDim JetItem.FulfillmentNodeQuantity(0 To 10) As Single
JetItem.FulfillmentNodeID(0) = ScottRdNode



Dim invData As ADODB.Recordset
Set invData = DB.retrieve("SELECT QuantityOnHand FROM vInventorySpreadsheet WHERE ITEMNUMBER='" & ItemNumber & "'")
If (invData.RecordCount > 0) Then
    JetItem.FulfillmentNodeQuantity(0) = CInt(invData("QuantityOnHand"))
    GetInventoryInfo = CreateInventoryJSON(JetItem)
Else
    GetInventoryInfo = ""
    Exit Function
End If



End Function

Public Function GetImageInfo(ItemNumber As String) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Phase 2: Get image info...
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim JetItem As JetItemInfo
Dim mainImageURL As String
mainImageURL = PictureDBFunctions.GenerateItemPictureURL(ItemNumber, False, Array("noauth-nobg"))
If mainImageURL <> "" Then JetItem.mainImageURL = mainImageURL

GetImageInfo = CreateImageJSON(JetItem)

End Function


Public Function GetItemInfo(ItemNumber As String) As String
'vars for general info
     'ProductTitle As String 'required
     'JetBrowseNodeID As String
     'SKU As String
     'AmazonItemTypeKeyword As String
     'CategoryPath As String
     'SKU_Code As String 'required
     'SKU_Type As String 'required
     'ASIN As String
     'MultipackQty As Integer 'required
     'BrandName As String 'required
     'Manufacturer As String 'required
     'ManufacturerPartNumber As String
     'ProductDescription As String
     'Bullets() As String
     'NumberUnitsForPricePerUnit As Single
     'TypeOfUnitForPricePerUnit As String
     'ShippingWeightPounds As Single
     'PackageLengthInches As Single
     'PackageWidthInches As Single
     'PackageHeightInches As Single
     'DisplayLengthInches As Single
     'DisplayWidthInches As Single
     'DisplayHeightInches As Single
     'Prop65 As Boolean
     'LegalDisclaimerDescription As String
     'CPSIACautionaryStatements() As String
     'CountryOfOrigin As String
     'SafetyWarning As String
     'StartSellingDate As String
     'FulfillmentTime As Integer
     'MSRP As Single
     'MAPP As Currency
     'MAPPImplementation As String
     'ProductTaxCode As String
     'NoReturnFeeAdjustment As String
     'ExcludeFromFeeAdjustments As Boolean
     'ShipsAlone As Boolean
     'AttributeID() As Integer
     'AttributeValue() As String
     'AttributeValueUnit() As String
     
'vars for photo...
     'mainImageURL As String 'required
     'SwatchImageURL As String
     'AlternateImageSlotID() As String 'up to 8 slots 'required
     'AlternateImageURL() As String 'required
    
'inventory vars...
     'FulfillmentNodeID() As String 'required
     'FulfillmentNodeQuantity() As Single 'required
     
'pricing vars...
     'Price As Currency 'required
     'FulfillmentNodePrice() As Single 'required
     
'relationship vars...
     'Relationship As String 'required
     'VariationRefinements As Single
     'ChildrenSKUS() As String 'required
     
'shipping exceptions
     'ServiceLevel As String
     'ShippingCarrier As String
     'ShippingMethod As String
     'OverrideType As String
     'ShippingChargeAmount As Single
     'IsShippingRestricted As Boolean
     'IsShippingExclusive As Boolean

Dim JetItem As JetItemInfo
JetItem.id = ItemNumber

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Phase 1: Get main info...
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim sqlData As ADODB.Recordset
Set sqlData = DB.retrieve("SELECT ItemDescription,ManufFullName,ManufFullNameWeb,MAPP,Desc1,Desc3,Desc2,Spec1HTML,Spec2HTML,Spec3HTML,Spec4HTML,Spec5HTML,Spec6HTML,Spec7HTML,Spec8HTML,EBayPrice FROM InventoryMaster,ProductLine,PartNumbers WHERE InventoryMaster.ItemNumber='" & ItemNumber & "' AND InventoryMaster.ProductLine=ProductLine.ProductLine AND InventoryMaster.ItemNumber=PartNumbers.ItemNumber")
If sqlData.RecordCount > 0 Then
If Len(sqlData("ItemDescription")) > 0 Then JetItem.ProductTitle = JSONFormatString(sqlData("ItemDescription"))
JetItem.JetBrowseNodeID = """b5035a1a17d342658760c2f402db425d"""
If Len(sqlData("ManufFullNameWeb")) > 0 Then JetItem.BrandName = JSONFormatString(sqlData("ManufFullNameWeb"))
If Len(sqlData("ManufFullName")) > 0 Then JetItem.Manufacturer = JSONFormatString(sqlData("ManufFullName"))
If Len(sqlData("Desc1")) > 0 Then JetItem.ProductDescription = JSONFormatString(sqlData("Desc1"))
If Len(sqlData("Desc3")) > 0 Then JetItem.ProductDescription = JetItem.ProductDescription & " " & JSONFormatString(sqlData("Desc3"))
If Len(sqlData("Desc2")) > 0 Then JetItem.ManufacturerPartNumber = JSONFormatString(sqlData("Desc2"))

Dim BulletArray(0 To 7) As String
If Len(sqlData("Spec1HTML")) > 0 Then BulletArray(0) = JSONFormatString(sqlData("Spec1HTML"))
If Len(sqlData("Spec2HTML")) > 0 Then BulletArray(1) = JSONFormatString(sqlData("Spec2HTML"))
If Len(sqlData("Spec3HTML")) > 0 Then BulletArray(2) = JSONFormatString(sqlData("Spec3HTML"))

If Len(sqlData("Spec4HTML")) > 0 Then BulletArray(3) = JSONFormatString(sqlData("Spec4HTML"))
If Len(sqlData("Spec5HTML")) > 0 Then BulletArray(4) = JSONFormatString(sqlData("Spec5HTML"))
If Len(sqlData("Spec6HTML")) > 0 Then BulletArray(5) = JSONFormatString(sqlData("Spec6HTML"))
If Len(sqlData("Spec7HTML")) > 0 Then BulletArray(6) = JSONFormatString(sqlData("Spec7HTML"))
If Len(sqlData("Spec8HTML")) > 0 Then BulletArray(7) = JSONFormatString(sqlData("Spec8HTML"))
JetItem.Bullets = BulletArray

If CInt(sqlData("MAPP")) > 0 Then
JetItem.MAPP = CDec(sqlData("MAPP"))
JetItem.MAPPImplementation = "type1"
End If

If Len(sqlData("EBayPrice")) > 0 Then JetItem.price = CDec(sqlData("EBayPrice"))
End If


Set sqlData = DB.retrieve("SELECT Barcode FROM InventoryComponentBarcodes,InventoryComponentMap WHERE ItemNumber='" & ItemNumber & "' AND InventoryComponentMap.ComponentID=InventoryComponentBarcodes.ComponentID")
If (sqlData.RecordCount > 0) Then
JetItem.SKU = ItemNumber
JetItem.SKU_Code = sqlData("Barcode")
JetItem.SKU_Type = "UPC"
End If

Dim mainImageURL As String
mainImageURL = PictureDBFunctions.GenerateItemPictureURL(JetItem.id, False, Array("noauth-nobg"))
JetItem.mainImageURL = mainImageURL

JetItem.ASIN = ""
JetItem.MultipackQty = 1
JetItem.AmazonItemTypeKeyword = ""
JetItem.CategoryPath = ""


'Get all components of this item...
Dim ComponentList As ADODB.Recordset
Set ComponentList = DB.retrieve("SELECT ComponentID, Quantity FROM InventoryComponentMap WHERE ItemNumber='" & ItemNumber & "'")
Dim components As String
Dim componentCounter As Integer
componentCounter = 0
Do
    If ComponentList.EOF = False And ComponentList.BOF = False Then
        If Len(ComponentList("ComponentID")) > 0 And Len(ComponentList("Quantity")) > 0 Then
            
            components = components & ComponentList("ComponentID") & "," & ComponentList("Quantity") & "+"
        End If
        componentCounter = componentCounter + 1
    End If
    ComponentList.MoveNext
Loop Until ComponentList.BOF = True Or ComponentList.EOF = True

'lets calc the weight, length, width, and height...
Dim weightTotal As Double
Dim lengthTotal As Double
Dim widthTotal As Double
Dim heightTotal As Double


Do
If componentCounter > 0 Then
    componentCounter = 0
    Dim comp As Variant
    Dim componentID As String
    Dim componentQTY As Integer
    For Each comp In Split(components, "+")
        If Len(comp) > 0 Then
            componentID = Split(comp, ",")(0)
            componentQTY = Split(comp, ",")(1)
            Dim componentData As ADODB.Recordset
            Set componentData = DB.retrieve("SELECT Length,Width,Height,Weight FROM InventoryComponents WHERE ID=" & componentID)
            If componentData.RecordCount > 0 Then
                weightTotal = weightTotal + CDbl(componentData("Weight"))
                lengthTotal = lengthTotal + CDbl(componentData("Length"))
                widthTotal = widthTotal + CDbl(componentData("Width"))
                heightTotal = heightTotal + CDbl(componentData("Height"))
            End If
        End If
    Next comp
End If
Loop Until componentCounter = ComponentList.RecordCount - 1

JetItem.ShippingWeightPounds = weightTotal + 1 'adding a pound for the box itself...
JetItem.PackageHeightInches = heightTotal
JetItem.PackageLengthInches = lengthTotal
JetItem.PackageWidthInches = widthTotal






GetItemInfo = CreateMainJSON(JetItem)


End Function
Public Function JSONFormatString(stringToConvert As String) As String
    Dim jsonFormat As String
    jsonFormat = stringToConvert
    
    jsonFormat = Replace(jsonFormat, "\", "\\")
    jsonFormat = Replace(jsonFormat, """", "\""")
    jsonFormat = Replace(jsonFormat, "/", "\/")
    jsonFormat = Replace(jsonFormat, vbCrLf, "\r\n")
    JSONFormatString = jsonFormat
    
End Function


Public Function CreateMainJSON(JetItem As JetItemInfo) As String

    Dim JSON_OUT As String

    '& ",""jet_browse_node_id"": " & JetItem.JetBrowseNodeID
    JSON_OUT = ""
    JSON_OUT = "{""product_title"": """ & JetItem.ProductTitle & """"
    JSON_OUT = JSON_OUT & ",""jet_retail_sku"": """ & JetItem.SKU & ""","
    JSON_OUT = JSON_OUT & """standard_product_codes"": [{""standard_product_code"": """ & JetItem.SKU_Code
    JSON_OUT = JSON_OUT & """,""standard_product_code_type"": """ & JetItem.SKU_Type & """}],"
    JSON_OUT = JSON_OUT & """multipack_quantity"": " & CStr(JetItem.MultipackQty)
    JSON_OUT = JSON_OUT & ",""brand"": """ & JetItem.BrandName
    JSON_OUT = JSON_OUT & """,""manufacturer"": """ & JetItem.Manufacturer & ""","
    JSON_OUT = JSON_OUT & """mfr_part_number"": """ & JetItem.ManufacturerPartNumber
    JSON_OUT = JSON_OUT & """," & """product_description"": """ & JetItem.ProductDescription & ""","
    
    If JetItem.Bullets(0) <> "" Then
         JSON_OUT = JSON_OUT & """bullets"": [""" & JetItem.Bullets(0) & """"
   
         If JetItem.Bullets(1) <> "" Then
         JSON_OUT = JSON_OUT & ",""" & JetItem.Bullets(1) & """"
         End If
         If JetItem.Bullets(2) <> "" Then
         JSON_OUT = JSON_OUT & ",""" & JetItem.Bullets(2) & """"
         End If
         If JetItem.Bullets(3) <> "" Then
         JSON_OUT = JSON_OUT & ",""" & JetItem.Bullets(3) & """"
         End If
         If JetItem.Bullets(4) <> "" Then
         JSON_OUT = JSON_OUT & ",""" & JetItem.Bullets(4) & """"
         End If
         'If JetItem.Bullets(5) <> "" Then
         'JSON_OUT = JSON_OUT & ",""" & JetItem.Bullets(5) & """"
         'End If
         'If JetItem.Bullets(6) <> "" Then
         'JSON_OUT = JSON_OUT & ",""" & JetItem.Bullets(6) & """"
         'End If
         'If JetItem.Bullets(7) <> "" Then
         'JSON_OUT = JSON_OUT & ",""" & JetItem.Bullets(7) & """"
         'End If
         JSON_OUT = JSON_OUT & "],"
    End If
    
    JSON_OUT = JSON_OUT & """number_units_for_price_per_unit"":1,"
    JSON_OUT = JSON_OUT & """type_of_unit_for_price_per_unit"":""each"","
    JSON_OUT = JSON_OUT & """shipping_weight_pounds"": " & CStr(JetItem.ShippingWeightPounds) & ","
    JSON_OUT = JSON_OUT & """package_length_inches"": " & CStr(JetItem.PackageLengthInches) & ","
    JSON_OUT = JSON_OUT & """package_width_inches"": " & CStr(JetItem.PackageWidthInches) & ","
    JSON_OUT = JSON_OUT & """package_height_inches"": " & CStr(JetItem.PackageHeightInches) & ","
    'JSON_OUT = JSON_OUT & """display_length_inches"": 0,"
    'JSON_OUT = JSON_OUT & """display_width_inches"": 0,"
    'JSON_OUT = JSON_OUT & """display_height_inches"": 0,"
    'JSON_OUT = JSON_OUT & """prop_65"":false,"
    'JSON_OUT = JSON_OUT & """legal_disclaimer_description"": ""Legal Statement Here"","

    'If JetItem.CPSIACautionaryStatements(0) <> "" Or JetItem.CPSIACautionaryStatements <> Null Then
    '    JSON_OUT = JSON_OUT & """cpsia_cautionary_statements"":["
    '        Dim cpsiaCount As Integer
    '        For cpsiaCount = 0 To UBound(JetItem.CPSIACautionaryStatements)
    '            JSON_OUT = JSON_OUT & """" & JetItem.CPSIACautionaryStatements & """"
    '            If cpsiaCount <> UBound(JetItem.CPSIACautionaryStatements) Then JSON_OUT = JSON_OUT & ","
    '        Next cpsiaCount
    '   JSON_OUT = JSON_OUT & "],"
    'End If

    'JSON_OUT = JSON_OUT & """country_of_origin"":""" & JetItem.CountryOfOrigin & ""","
    'JSON_OUT = JSON_OUT & """safety_warning"":""" & JetItem.SafetyWarning & ""","
    'JSON_OUT = JSON_OUT & """start_selling_date"": """ & JetItem.StartSellingDate & ""","
    'JSON_OUT = JSON_OUT & """fulfillment_time"":" & CStr(JetItem.FulfillmentTime) & ","
    'JSON_OUT = JSON_OUT & """msrp"":" & CStr(JetItem.MSRP) & ","
    If JetItem.MAPP > 0 Then
    JSON_OUT = JSON_OUT & """map_price"":" & CStr(JetItem.MAPP) & ","
    JSON_OUT = JSON_OUT & """map_implementation"":""" & JetItem.MAPPImplementation & ""","
    End If
    'JSON_OUT = JSON_OUT & """product_tax_code"":""" & JetItem.ProductTaxCode & ""","
    'JSON_OUT = JSON_OUT & """no_return_fee_adjustment"":" & JetItem.NoReturnFeeAdjustment & ","
    'JSON_OUT = JSON_OUT & """exclude_from_fee_adjustments"":" & JetItem.ExcludeFromFee & ","
    'JSON_OUT = JSON_OUT & """ships_alone"":" & JetItem.ShipsAlone & ","
    
    
    'add in attributes here...
    
    JSON_OUT = Mid(JSON_OUT, 1, Len(JSON_OUT) - 1)
    JSON_OUT = JSON_OUT & "}"
    CreateMainJSON = JSON_OUT

End Function
Public Function CreateImageJSON(JetItem As JetItemInfo) As String
    Dim JSON_OUT As String
    
    JSON_OUT = "{ ""main_image_url"": """ & JetItem.mainImageURL & """ }"
    
CreateImageJSON = JSON_OUT
End Function
Public Function CreateInventoryJSON(JetItem As JetItemInfo) As String
    'needs to support all fulfillment nodes in future...
    Dim JSON_OUT As String
    JSON_OUT = "{ ""fulfillment_nodes"": ["
    JSON_OUT = JSON_OUT & "{ ""fulfillment_node_id"": """ & JetItem.FulfillmentNodeID(0) & ""","
    JSON_OUT = JSON_OUT & """quantity"": " & JetItem.FulfillmentNodeQuantity(0) & "}"
    JSON_OUT = JSON_OUT & "]}"

    'MsgBox JSON_OUT
    CreateInventoryJSON = JSON_OUT
End Function
Public Function CreatePriceJSON(JetItem As JetItemInfo) As String
    Dim JSON_OUT As String
    'gotta make sure we have a fulfillment node...
    If Len(JetItem.FulfillmentNodeID(0)) > 0 And JetItem.FulfillmentNodePrice(0) > 0 Then
    
    JSON_OUT = "{""price"": " & CStr(JetItem.price) & ", ""fulfillment_nodes"": ["
    Dim X As Integer
    
    Do
    X = X + 1
    If JetItem.FulfillmentNodePrice(X - 1) > 0 Then
        JSON_OUT = JSON_OUT & "{ ""fulfillment_node_id"": """ & JetItem.FulfillmentNodeID(X - 1) & ""","
        JSON_OUT = JSON_OUT & """fulfillment_node_price"": " & JetItem.FulfillmentNodePrice(X - 1) & "},"
    End If
        
    Loop Until X > UBound(JetItem.FulfillmentNodePrice) - 1
    JSON_OUT = Mid(JSON_OUT, 1, Len(JSON_OUT) - 1)
    JSON_OUT = JSON_OUT & "]}"
    
    Else
    If JetItem.FulfillmentNodePrice(0) > 0 Then
    JSON_OUT = "{ ""price"": " & CStr(JetItem.FulfillmentNodePrice(0)) & "}"
    Else
        'dont offload this section as it would set the price to 0...
    CreatePriceJSON = "ERROR: Price set to zero."
    End If
    End If
    CreatePriceJSON = JSON_OUT
    'MsgBox JSON_OUT
End Function
