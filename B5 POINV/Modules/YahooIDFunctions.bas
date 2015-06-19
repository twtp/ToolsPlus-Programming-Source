Attribute VB_Name = "YahooIDFunctions"
Option Explicit

Private yahooItemNames As Dictionary

Public Function CheckConflictingItem(newitem As String) As String
    If yahooItemNames Is Nothing Then
        'similar function is also in import_orders.pl + dbo.CreateYahooIDs, change there if you change here
        Set yahooItemNames = New Dictionary
        Dim rst As ADODB.Recordset
        Dim itemID As String
        Set rst = DB.retrieve("SELECT ItemNumber, ManufFullNameWeb FROM PartNumbers INNER JOIN ProductLine ON LEFT(PartNumbers.ItemNumber,3)=ProductLine.ProductLine WHERE (WebPublished=1 OR EBayPublished=1) And WebPointered=0")
        While Not rst.EOF
            itemID = CreateYahooID(rst("ItemNumber"))
            'as long as items only get added with the proper sanity checks,
            'these will always be unique, so we don't have to worry here
            yahooItemNames.Add CStr(itemID), CStr(rst("ItemNumber"))
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
    End If
    
    Dim thisid As String
    thisid = CreateYahooID(newitem)
    If yahooItemNames.exists(thisid) Then
        If yahooItemNames.item(thisid) <> newitem Then
            CheckConflictingItem = yahooItemNames.item(thisid)
        Else
            CheckConflictingItem = ""
        End If
    Else
        CheckConflictingItem = ""
    End If
End Function

Public Function CreateYahooID(item As String) As String
    Dim newitem As String, manuf As String
    'if no web manuf, then keep the line code
    If Nz(ManufWebHashPLtoName.item(Left(item, 3))) = "" Then
        newitem = item
    Else
        manuf = ManufWebHashPLtoName.item(Left(item, 3))
        newitem = URLify(Replace(manuf, ".", "") & "-" & LCase(Mid(item, 4)))
        While InStr(newitem, "--")
            newitem = Replace(newitem, "--", "-")
        Wend
    End If
    CreateYahooID = newitem
End Function

Public Function FlushYahooIDCache() As Boolean
    'will regenerate on next query
    Set yahooItemNames = Nothing
End Function
