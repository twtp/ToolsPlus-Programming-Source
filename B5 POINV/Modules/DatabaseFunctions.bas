Attribute VB_Name = "DatabaseFunctions"
'---------------------------------------------------------------------------------------
' Module    : DatabaseFunctions
' DateTime  : 6/14/2005 10:33
' Author    : briandonorfio
' Purpose   : Everything to do with database connectivity above the DBConn level.
'             Mostly lookups and utility functions.
'
'             Dependencies:
'               - Microsoft ActiveX Data Objects 2.8 Library
'               - MouseHourglass + global Mouse object
'               - Utilities
'---------------------------------------------------------------------------------------

Option Compare Text
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : lookupPrimaryVendorNumber
' DateTime  : 6/14/2005 11:32
' Author    : briandonorfio
' Purpose   : Given a VendorName, returns associated VendorNumber, null string if not
'             found. VendorName is not guaranteed unique, only returns first.
'---------------------------------------------------------------------------------------
'
Public Function lookupPrimaryVendorNumber(VendorName As String) As String
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("EXEC spLookupVendorNumber '" & EscapeSQuotes(VendorName) & "'")
    If rst.EOF Then
        lookupPrimaryVendorNumber = ""
    Else
        lookupPrimaryVendorNumber = Nz(rst("VendorNo").value)
    End If
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : lookupLineCodePrimaryVendor
' DateTime  : 6/14/2005 11:33
' Author    : briandonorfio
' Purpose   : Returns the primary vendor for a line code, null string if not found.
'---------------------------------------------------------------------------------------
'
Public Function lookupLineCodePrimaryVendor(lc As String) As String
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("EXEC spLookupLineCodePrimaryVendor '" & lc & "'")
    If rst.EOF Then
        lookupLineCodePrimaryVendor = ""
    Else
        lookupLineCodePrimaryVendor = Nz(rst("PrimaryVendorNumber").value)
    End If
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : lookupPOTermsByDesc
' DateTime  : 6/14/2005 12:47
' Author    : briandonorfio
' Purpose   : Returns the PO's TermsCode for a terms description, null string if not
'             found.
'---------------------------------------------------------------------------------------
'
Public Function lookupPOTermsByDesc(Desc As String) As String
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("EXEC spLookupPOTermsByDesc '" & Desc & "'")
    If rst.EOF Then
        lookupPOTermsByDesc = ""
    Else
        lookupPOTermsByDesc = Nz(rst("TermsCode").value)
    End If
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : lookupPOTermDescByNum
' DateTime  : 6/14/2005 13:36
' Author    : briandonorfio
' Purpose   : Returns the description for a PO terms code, null string if not found.
'---------------------------------------------------------------------------------------
'
Public Function lookupPOTermDescByNum(num As String) As String
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("EXEC spLookupPOTermDescByCode '" & num & "'")
    If rst.EOF Then
        lookupPOTermDescByNum = ""
    Else
        lookupPOTermDescByNum = Nz(rst("Description").value)
    End If
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : lookupGroupSignID
' DateTime  : 8/24/2005 11:09
' Author    : briandonorfio
' Purpose   : Returns the ID for a given group sign name.
'---------------------------------------------------------------------------------------
'
'Public Function lookupGroupSignID(groupSignName As String) As String
'    Mouse.Hourglass True
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT ID FROM GroupSigns WHERE GroupName='" & EscapeSQuotes(groupSignName) & "'")
'    If rst.EOF Then
'        lookupGroupSignID = ""
'    Else
'        lookupGroupSignID = CStr(rst("ID"))
'    End If
'    rst.Close
'    Set rst = Nothing
'    Mouse.Hourglass False
'End Function

'---------------------------------------------------------------------------------------
' Procedure : updateInventoryMaster
' DateTime  : 6/14/2005 11:33
' Author    : briandonorfio
' Purpose   : Executes an UPDATE on InventoryMaster. Arguments are the field name, the
'             content for that field, the itemnumber PK, and the necessary delimiter for
'             the field to be updated.
'---------------------------------------------------------------------------------------
'
Public Function updateInventoryMaster(field As String, content As String, item As String, Optional delim As String = "") As Boolean
On Error GoTo errh
    Mouse.Hourglass True
    If Left$(content, 1) = "$" Then
        content = Replace(content, ",", "")
    End If
    DB.Execute "UPDATE InventoryMaster SET " & field & "=" & delim & EscapeSQuotes(content) & delim & " WHERE ItemNumber='" & item & "'"
    updateInventoryMaster = True
    Mouse.Hourglass False
    Exit Function

errh:
    MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    updateInventoryMaster = False
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : updateXSheetDetail
' DateTime  : 6/14/2005 11:34
' Author    : briandonorfio
' Purpose   : Executes an UPDATE on NewProdXSheetDetail. Arguments are the field name,
'             the content for that field, the itemnumber PK, and the necessary delimiter
'             for the field to be updated.
'---------------------------------------------------------------------------------------
'
Public Function updateXSheetDetail(field As String, content As String, item As String, Optional delim As String = "") As Boolean
On Error GoTo errh
    Mouse.Hourglass True
    DB.Execute "UPDATE NewProdXSheetDetail SET " & field & "=" & delim & EscapeSQuotes(content) & delim & " WHERE ItemNumber='" & item & "'"
    updateXSheetDetail = True
    Mouse.Hourglass False
    Exit Function

errh:
    MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    updateXSheetDetail = False
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : updateXSheetHeader
' DateTime  : 6/14/2005 11:35
' Author    : briandonorfio
' Purpose   : Executes an UPDATE on NewProdXSheetHeader. Arguments are the field name,
'             the content for that field, the vendornumber PK, and the necessary
'             delimiter for the field to be updated.
'---------------------------------------------------------------------------------------
'
Public Function updateXSheetHeader(field As String, content As String, VendorNumber As String, Optional delim As String = "") As Boolean
On Error GoTo errh
    Mouse.Hourglass True
    DB.Execute "UPDATE NewProdXSheetHeader SET " & field & "=" & delim & EscapeSQuotes(content) & delim & " WHERE VendorNumber='" & VendorNumber & "'"
    updateXSheetHeader = True
    Mouse.Hourglass False
    Exit Function

errh:
    MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    updateXSheetHeader = False
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : updatePurchaseOrders
' DateTime  : 6/14/2005 11:36
' Author    : briandonorfio
' Purpose   : Executes an UPDATE on PurchaseOrders. Arguments are the field name, the
'             content for that field, the purchase order ID PK, and the necessary
'             delimiter for the field to be updated.
'---------------------------------------------------------------------------------------
'
Public Function updatePurchaseOrders(field As String, content As String, POID As String, Optional delim As String = "") As Boolean
On Error GoTo errh
    Mouse.Hourglass True
    DB.Execute "UPDATE PurchaseOrders SET " & field & "=" & delim & EscapeSQuotes(content) & delim & " WHERE ID=" & POID
    updatePurchaseOrders = True
    Mouse.Hourglass False
    Exit Function

errh:
    MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    updatePurchaseOrders = False
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : updateProdLine
' DateTime  : 6/16/2005 14:53
' Author    : briandonorfio
' Purpose   : Executes an UPDATE on ProductLine. Arguments are the field name, the
'             content for that field, the LineCode/ProductLine, and the necessary
'             delimiter for the filed to be updated.
'---------------------------------------------------------------------------------------
'
Public Function updateProdLine(field As String, content As String, lc As String, Optional delim As String = "") As Boolean
On Error GoTo errh
    Mouse.Hourglass True
    If Left$(content, 1) = "$" Then
        content = Replace(content, ",", "")
    End If
    DB.Execute "UPDATE ProductLine SET " & field & "=" & delim & EscapeSQuotes(content) & delim & " WHERE ProductLine='" & lc & "'"
    updateProdLine = True
    Mouse.Hourglass False
    Exit Function

errh:
    MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    updateProdLine = False
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : updatePartNumbers
' DateTime  : 7/1/2005 12:15
' Author    : briandonorfio
' Purpose   : Executes an UPDATE on the PartNumbers table. Arguments are the field name,
'             content for the field, the itemnumber to update, and the delimiter, if
'             needed. Returns true on successful update, false otherwise.
'---------------------------------------------------------------------------------------
'
Public Function updatePartNumbers(field As String, content As String, item As String, Optional delim As String = "") As Boolean
On Error GoTo errh
    Mouse.Hourglass True
    If field = "RebatePrice" Then
        content = Replace(Replace(content, ",", ""), "$", "")
    End If
    DB.Execute "UPDATE PartNumbers SET " & field & "=" & delim & EscapeSQuotes(content) & delim & " WHERE ItemNumber='" & item & "'"
    updatePartNumbers = True
    Mouse.Hourglass False
    Exit Function

errh:
    MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    updatePartNumbers = False
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : DLookup
' DateTime  : 6/14/2005 10:56
' Author    : briandonorfio
' Purpose   : Similar to Access 2k's DLookup() function. Returns the first record
'             specified by the where clause, should be identifying by PK for best
'             results (returns the field from the first record, so any others get thrown
'             out). Where clause is a normal SQL clause, minus the WHERE.
'---------------------------------------------------------------------------------------
'
Public Function DLookup(field As String, Optional table As String = "", Optional whereClause As String = "") As Variant
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT TOP 1 " & field & IIf(table <> "", " FROM " & table, "") & IIf(whereClause <> "", " WHERE " & whereClause, ""))
    If rst.EOF Then
        DLookup = ""
    Else
        DLookup = Nz(rst(0))
    End If
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Function

Public Function GetSQLServerDate() As String
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT GETDATE()")
    Dim datestamp As String
    datestamp = rst(0)
    rst.Close
    GetSQLServerDate = datestamp
End Function

Public Function ModifyItemCost(itemID As Long, costTypeStr As String, costValue As String, Optional applicationModuleID As Long = -1)
    If applicationModuleID = -1 Then
        applicationModuleID = ApplicationModuleHashNameToID.item("PoInv - Inventory Maintenance")
    End If
    Dim cleanedCost As String
    cleanedCost = Replace(Replace(costValue, "$", ""), ",", "")
    
    Dim currentCost As String
    currentCost = DLookup("CostValue", "vItemCostCurrent", "ItemID=" & itemID & " AND ItemCostType='" & costTypeStr & "'")
    If currentCost <> "" Then
        If CDec(cleanedCost) = CDec(currentCost) Then
            'nothing to do
            ModifyItemCost = False
            Exit Function
        End If
    End If
    
    Dim sql As String
    'sql = "INSERT INTO ItemCost (ItemID, ItemCostTypeID, CostValue, PersonID, ApplicationModuleID) VALUES (@itemid, @typeid, @cost, @pid, @aid)"
    sql = "INSERT INTO ItemCost (ItemID, ItemCostType, CostValue, PersonID, ApplicationModuleID) VALUES (@itemid, @type, @cost, @pid, @aid)"
    'poor mans placeholders?
    sql = Replace(sql, "@itemid", itemID)
    'sql = Replace(sql, "@typeid", typeID)
    sql = Replace(sql, "@type", "'" & EscapeSQuotes(costTypeStr) & "'")
    sql = Replace(sql, "@cost", cleanedCost)
    sql = Replace(sql, "@pid", GetCurrentUserID())
    sql = Replace(sql, "@aid", applicationModuleID) ' also enum-able
    DB.Execute sql
    ModifyItemCost = True
End Function
