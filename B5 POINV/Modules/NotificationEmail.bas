Attribute VB_Name = "NotificationEmail"
'---------------------------------------------------------------------------------------
' Module    : NotificationEmail
' DateTime  : 8/30/2005 14:59
' Author    : briandonorfio
' Purpose   : Functions to handle email notifications about different item changes
'             (price, d/c, etc). Separated from the rest of the logic file so we don't
'             drag in the NotifyChanges+OutlookIntegration dependencies.
'
'             Dependencies:
'               - OutlookIntegration
'               - PerlArray + global NotifyChanges
'               - Microsoft Scripting Runtime
'---------------------------------------------------------------------------------------

Option Explicit

Public Sub UpdateNotificationEmail(changetype As String, item As String)
    '2013-08-06: not including web price anymore, adding store discontinued
    Select Case changetype
        'Case Is = "cost"
        '    NotifyChanges.Push Array(item, changeType, DLookup("StdCost", "InventoryMaster", "ItemNumber='" & item & "'"))
        Case Is = "store price"
            NotifyChanges.Push Array(item, changetype, DLookup("StdPrice", "InventoryMaster", "ItemNumber='" & item & "'"))
        'Case Is = "web price"
        '    NotifyChanges.Push Array(item, changetype, DLookup("IDiscountMarkupPriceRate1", "InventoryMaster", "ItemNumber='" & item & "'"))
        Case Is = "store disc"
            NotifyChanges.Push Array(item, changetype, "")
    End Select
End Sub

Public Sub SendNotifyEmail()
    Dim dcBody As String, cBody As String, spBody As String, wpBody As String, sdBody As String
    Dim i As Long
    Dim newprice As Double
    
    Dim discH As Dictionary, spH As Dictionary, wpH As Dictionary, sdH As Dictionary
    Set discH = New Dictionary
    Set spH = New Dictionary
    Set wpH = New Dictionary
    Set sdH = New Dictionary
    
    For i = NotifyChanges.Scalar - 1 To 0 Step -1   'reverse through the array, so we grab last update if item was
        Select Case NotifyChanges.Elem(i)(1)        'done twice, then save to a hash, so we don't duplicate
            'Case Is = "discontinued"
            '    If Not discH.exists(NotifyChanges.Elem(i)(0)) Then
            '        discH.Add NotifyChanges.Elem(i)(0), "1"
            '        'dlookup qoh -> being d/c or dc
            '            dcBody = dcBody & "<tr><td>" & NotifyChanges.Elem(i)(0) & "</td><td>" & DLookup("Desc3", "PartNumbers", "ItemNumber='" & NotifyChanges.Elem(i)(0) & "'") & "</td><td>" & DLookup("ReplacedBy", "InventoryMaster", "ItemNumber='" & NotifyChanges.Elem(i)(0) & "'") & "</td></tr>"
            '    End If
            'Case Is = "cost"
            '    cBody = cBody & "<tr><td>" & NotifyChanges.Elem(i)(0) & "</td><td>" & DLookup("Desc3HTML", "PartNumbers", "ItemNumber='" & NotifyChanges.Elem(i)(0) & "'") & "</td><td>" & NotifyChanges.Elem(i)(2) & "</td><td>" & DLookup("StdCost", "InventoryMaster", "ItemNumber='" & NotifyChanges.Elem(i)(0) & "'") & "</td></tr>"
            Case Is = "store price"
                If Not spH.exists(NotifyChanges.Elem(i)(0)) Then
                    spH.Add NotifyChanges.Elem(i)(0), "1"
                    newprice = DLookup("StdPrice", "InventoryMaster", "ItemNumber='" & NotifyChanges.Elem(i)(0) & "'")
                    Dim storeQty As String
                    storeQty = getStoreQuantityFor(CStr(NotifyChanges.Elem(i)(0)))
                    If newprice <> NotifyChanges.Elem(i)(2) Then
                        spBody = spBody & "<tr><td>" & NotifyChanges.Elem(i)(0) & "</td>" & _
                                              "<td>" & DLookup("Desc3", "PartNumbers", "ItemNumber='" & NotifyChanges.Elem(i)(0) & "'") & "</td>" & _
                                              "<td>" & FormatNumber(NotifyChanges.Elem(i)(2), 2, vbTrue, vbFalse, vbFalse) & "</td>" & _
                                              "<td>" & FormatNumber(newprice, 2, vbTrue, vbFalse, vbFalse) & "</td>" & _
                                              "<td>" & storeQty & "</td></tr>"
                    End If
                End If
            Case Is = "web price"
                If Not wpH.exists(NotifyChanges.Elem(i)(0)) Then
                    wpH.Add NotifyChanges.Elem(i)(0), "1"
                    newprice = DLookup("IDiscountMarkupPriceRate1", "InventoryMaster", "ItemNumber='" & NotifyChanges.Elem(i)(0) & "'")
                    If newprice <> NotifyChanges.Elem(i)(2) Then
                        wpBody = wpBody & "<tr><td>" & NotifyChanges.Elem(i)(0) & "</td>" & _
                                              "<td>" & DLookup("Desc3", "PartNumbers", "ItemNumber='" & NotifyChanges.Elem(i)(0) & "'") & "</td>" & _
                                              "<td>" & FormatNumber(NotifyChanges.Elem(i)(2), 2, vbTrue, vbFalse, vbFalse) & "</td>" & _
                                              "<td>" & FormatNumber(newprice, 2, vbTrue, vbFalse, vbFalse) & "</td></tr>"
                    End If
                End If
            Case Is = "store disc"
                If Not sdH.exists(NotifyChanges.Elem(i)(0)) Then
                    sdH.Add NotifyChanges.Elem(i)(0), "1"
                    Dim repl As String
                    repl = DLookup("ReplacedBy", "InventoryMaster", "ItemNumber='" & NotifyChanges.Elem(i)(0) & "'")
                    If repl = "" Then
                        repl = "&nbsp;"
                    End If
                    sdBody = sdBody & "<tr><td>" & NotifyChanges.Elem(i)(0) & "</td>" & _
                                          "<td>" & repl & "</td>"
                End If
        End Select
    Next i

    If dcBody <> "" Then
        dcBody = "<h2>The following items have been discontinued:</h2><table border=""1""><thead><tr><td>Item:</td><td>Description</td><td>Replacement:</td></tr></thead><tbody>" & dcBody & "</tbody></table>"
    End If
    'If cBody <> "" Then
    '    cBody = "<h2>The Cost has changed for the following:</h2><table border=""1""><thead><tr><td>Item:</td><td>Description</td><td>Old Cost:</td><td>New Cost:</td></tr></thead><tbody>" & cBody & "</tbody></table>"
    'End If
    If spBody <> "" Then
        spBody = "<h2>The In-Store price has changed for the following:</h2><table border=""1""><thead><tr><td>Item:</td><td>Description</td><td>Old Price:</td><td>New Price:</td><td>Store Quantity</td></tr></thead><tbody>" & spBody & "</tbody></table>"
    End If
    If wpBody <> "" Then
        wpBody = "<h2>The Internet price has changed for the following:</h2><table border=""1""><thead><tr><td>Item:</td><td>Description</td><td>Old Price:</td><td>New Price:</td></tr></thead><tbody>" & wpBody & "</tbody></table>"
    End If
    If sdBody <> "" Then
        sdBody = "<h2>The following items have been discontinued at the store:</h2><table border=""1""><thead><tr><td>Item:</td><td>Replacement</td></tr></thead><tbody>" & sdBody & "</tbody></table>"
    End If
    
    OpenEmailTo "ProductChangeUpdates", "Product Change Updates", , "<h1>Product Change Updates</h1>" & dcBody & cBody & spBody & wpBody & sdBody
End Sub

Private Function getStoreQuantityFor(item As String) As String
    Dim req As MSXML2.XMLHTTP
    Set req = New MSXML2.XMLHTTP
    req.Open "GET", "http://toolsplus04/whse/inventory.plex?item=" & item, False
    req.Send
    While req.ReadyState <> 4
        Sleep 100
    Wend
    If req.status = 200 Then
        getStoreQuantityFor = req.responseText
    Else
        getStoreQuantityFor = "???"
    End If
    Set req = Nothing
End Function

