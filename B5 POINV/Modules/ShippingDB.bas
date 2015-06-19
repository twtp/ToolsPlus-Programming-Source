Attribute VB_Name = "ShippingDB"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : DetermineBoxSize
' DateTime  : 8/2/2005 12:18
' Author    : briandonorfio
' Purpose   : Given an items dimensions, returns the box id for the smallest box it fits
'             into. Assumes weight and dimensions are correctly formatted, no error
'             handling, so everything should be checked before calling this.
'---------------------------------------------------------------------------------------
'
'Public Function DetermineBoxSize(Dimensions As String, Weight As String) As String
'    Mouse.Hourglass True
'    Dim rstBox As ADODB.Recordset
'    Set rstBox = DB.retrieve("SELECT ID, Length, Width, Height FROM BoxDimensions WHERE ID<>-1 ORDER BY Length, Width, Height")
'    Dim boxtype As String
'    boxtype = ""
'    Dim dimA() As String
'    Dim wtA() As String
'    If InStr(Dimensions, ";") Then
'        dimA = Split(Dimensions, "; ")
'        wtA = Split(Weight, "; ")
'        If UBound(dimA) <> UBound(wtA) Then
'            DetermineBoxSize = "-1"
'        Else
'            DetermineBoxSize = "999"
'        End If
'    ElseIf Dimensions = "0x0x0" Or Weight = "0" Then
'        DetermineBoxSize = "-1"
'    Else
'        dimA = Split(Dimensions, "x")
'        While Not rstBox.EOF
'            If CSng(dimA(0)) <= rstBox("Length") And CSng(dimA(1)) <= rstBox("Width") And CSng(dimA(2)) <= rstBox("Height") Then
'                boxtype = rstBox("ID")
'                rstBox.MoveLast
'            End If
'            rstBox.MoveNext
'        Wend
'        If boxtype <> "" Then
'            DetermineBoxSize = boxtype
'        Else
'            DetermineBoxSize = "999"
'        End If
'    End If
'    Mouse.Hourglass False
'End Function
Public Function DetermineBoxSize2(l As Double, w As Double, h As Double, wt As Double) As String
    Dim temp As Double
    If l < w Then
        temp = l
        l = w
        w = temp
    End If
    If w < h Then
        temp = w
        w = h
        h = temp
    End If
    If l < w Then
        temp = l
        l = w
        w = temp
    End If
    Dim rstBox As ADODB.Recordset
    Set rstBox = DB.retrieve("SELECT TOP 1 ID FROM BoxDimensions WHERE Deprecated=0 AND Length>=" & l & " AND Width>=" & w & " AND Height>=" & h & " AND (MaximumWeight IS NULL OR MaximumWeight>=" & wt & ") ORDER BY Length*Width*Height")
    Dim boxtype As String
    If rstBox.EOF Then
        boxtype = "-1"
    Else
        boxtype = rstBox("ID")
    End If
    rstBox.Close
    Set rstBox = Nothing
    DetermineBoxSize2 = boxtype
End Function

'---------------------------------------------------------------------------------------
' Procedure : ExternalItemDBSync
' DateTime  : 2/7/2008 10:53
' Author    : briandonorfio
' Purpose   : Synchronizes the external item database that we use for calculating the
'             shipping cost for a sales order and sending confirmation emails.
'---------------------------------------------------------------------------------------
'
'Public Function ExternalItemDBSync(Optional item As String = "") As Boolean
'    Dim errorcount As Long
'    Dim errorstring As String
'On Error GoTo errh
'    Dim iterArray As Variant
'    If item = "" Then
'        iterArray = Array("A%", "B%", "C%", "D%", "E%", "F%", "G%", "H%", "I%", "J%", "K%", "L%", "M%", "N%", "O%", "P%", "Q%", "R%", "S%", "T%", "U%", "V%", "W%")
'    Else
'        iterArray = Array(item)
'    End If
'    Dim iter As Variant
'    Dim rst As ADODB.Recordset
'    For Each iter In iterArray
'        If UBound(iterArray) <> 0 Then
'            Sleep 30000
'        End If
'        DoEvents
'        Dim xmlstr As String
'        xmlstr = ""
'        Set rst = DB.retrieve("SELECT ItemNumber, YahooID, Description, Dropship, StdCost, AvgCost, TruckFreight, TruckFreightCheap FROM vVPSItems WHERE ItemNumber" & IIf(CBool(InStr(iter, "%")), " LIKE ", "=") & "'" & iter & "'")
'        While Not rst.EOF
'            Dim currxml As String
'            currxml = "<Item>"
'            Dim i As Long
'            For i = 0 To rst.Fields.Count - 1
'                If rst(i).name = "Dropship" Or rst(i).name = "TruckFreight" Or rst(i).name = "TruckFreightCheap" Then
'                    currxml = currxml & "<" & rst(i).name & ">" & IIf(rst(i).value = "True", "1", "0") & "</" & rst(i).name & ">"
'                Else
'                    currxml = currxml & "<" & rst(i).name & ">" & escapeXMLEntities(removeUpperASCII(rst(i).value)) & "</" & rst(i).name & ">"
'                End If
'            Next i
'            currxml = currxml & "</Item>" & vbLf
'            xmlstr = xmlstr & currxml
'            rst.MoveNext
'        Wend
'
'        If xmlstr <> "" Then
'            xmlstr = "<Items>" & vbCrLf & xmlstr & "</Items>"
'            Dim currErrCount As Long, currDone As Boolean
'            currErrCount = 0
'            currDone = False
'            Dim req As MSXML2.XMLHTTP
'            While Not currDone And currErrCount < 5
'                Set req = New MSXML2.XMLHTTP
'                req.Open "POST", "http://64.22.97.22/internal/poinv/set_item_info.fcgi", False
'                req.send xmlstr
'                'Debug.Print iter & " " & req.status
'                While req.readyState <> 4
'                    'Debug.Print req.ReadyState
'                    Sleep 500
'                    DoEvents
'                Wend
'                If req.status = 200 Then
'                    currDone = True
'                Else
'                    'transient 500 error? try again
'                    currErrCount = currErrCount + 1
'                    Sleep 60000
'                End If
'            Wend
'            If Not currDone Then
'                errorcount = errorcount + 1
'                errorstring = errorstring & iter & ": " & req.responseText & vbCrLf
'            Else
'                errorstring = errorstring & iter & ": " & req.status & vbCrLf
'            End If
'        Else
'            If rst.RecordCount = 0 Then
'                'ok
'            Else
'                'hmm
'                errorcount = errorcount + 1
'                errorstring = errorstring & "No xml generated for " & qq(CStr(iter)) & vbCrLf
'            End If
'        End If
'
'        rst.Close
'    Next iter
'
'    If errorcount = 0 Then
'        ExternalItemDBSync = True
'    Else
'        SendEmailTo "brian@tools-plus.com", "Database Sync Report (Item)", errorcount & " error(s) occurred." & vbCrLf & vbCrLf & errorstring
'        ExternalItemDBSync = False
'    End If
'    Exit Function
'
'errh:
'    SendEmailTo "brian@tools-plus.com", "Database Sync Report (Item)", "unexpected error(s) occurred." & vbCrLf & vbCrLf & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & errorstring
'    Exit Function
'End Function

Public Function ExternalItemDBSync(item As String) As Boolean
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber, YahooID, EBayItemID, Description, Dropship, StdCost, AvgCost, InetPrice, ShippingType FROM vVPSItems WHERE ItemNumber='" & item & "'")
    If rst.EOF Then
        ExternalItemDBSync = False
        rst.Close
        Set rst = Nothing
        Exit Function
    End If
    
    Dim xmlin As String
    xmlin = "<Item>" & vbCrLf & _
            " <ItemNumber>" & rst("ItemNumber") & "</ItemNumber>" & vbCrLf & _
            " <YahooID>" & rst("YahooID") & "</YahooID>" & vbCrLf & _
            " <EBayItemID>" & Nz(rst("EBayItemID"), "") & "</EBayItemID>" & vbCrLf & _
            " <Description>" & escapeXMLEntities(removeUpperASCII(rst("Description"))) & "</Description>" & vbCrLf & _
            " <Dropship>" & IIf(rst("Dropship"), "1", "0") & "</Dropship>" & vbCrLf & _
            " <StdCost>" & rst("StdCost") & "</StdCost>" & vbCrLf & _
            " <AvgCost>" & rst("AvgCost") & "</AvgCost>" & vbCrLf & _
            " <InetPrice>" & rst("InetPrice") & "</InetPrice>" & vbCrLf & _
            " <ShippingType>" & rst("ShippingType") & "</ShippingType>" & vbCrLf & _
            "</Item>" & vbCrLf
    rst.Close
    Set rst = Nothing
    
    'Dim result As VbMsgBoxResult
    'result = MsgBox("Would you like to debug the XML output before attempting command?", vbYesNo + vbDefaultButton2, "Debug?")
    'If result = vbOK Or result = vbYes Then
    
    'Dim filefree As Long
    'filefree = FreeFile
    'Open "c:\users\tomwestbrook\desktop\poinv.txt" For Output As #filefree
    
    'Print #filefree, xmlin
    
    'Close #filefree
    
    
    'End If
    
    

    
    
    Dim currErrCount As Long, currDone As Boolean
    currErrCount = 0
    currDone = False
    Dim req As MSXML2.XMLHTTP
    While Not currDone And currErrCount < 10
        Set req = New MSXML2.XMLHTTP
        req.Open "POST", "http://www.toolsplusinternal.com/internal/poinv/set_item_info.pl", False
        req.Send xmlin
        'Debug.Print iter & " " & req.status
        While req.ReadyState <> 4
            'Debug.Print req.ReadyState
            Sleep 200
            DoEvents
        Wend
        If req.status = 200 Then
            currDone = True
            'Debug.Print req.responseText
        Else
            'transient 500 error? try again
            currErrCount = currErrCount + 1
            Sleep 20000
            DoEvents
        End If
    Wend
    If Not currDone Then
        Err.Raise "123", "ExternalItemDBSync", "Error sending data for " & item & ", error code was " & req.status
    End If
    
    ExternalItemDBSync = True
End Function

'---------------------------------------------------------------------------------------
' Procedure : ExternalComponentDBSync
' DateTime  : 2/7/2008 10:58
' Author    : briandonorfio
' Purpose   : Updates our external database with the weight and dimension information
'             for a given component or item.
'---------------------------------------------------------------------------------------
'
'Public Function ExternalComponentDBSync(Optional componentID As String = "", Optional item As String = "") As Boolean
Public Function ExternalComponentDBSync(item As String, Optional componentID As String = "") As Boolean
    'Dim errorcount As Long
    'Dim errorstring As String
On Error GoTo errh
    Dim iterArray As Variant, iterType As Variant
    'If componentID = "" And item = "" Then
    '    iterArray = Array("A%", "B%", "C%", "D%", "E%", "FEA%", "FES%", "FE[^AS]%", "F[^E]%", "G%", "H%", "I%", "J%", "K%", "L%", "M%", "N%", "O%", "P%", "Q%", "R%", "S%", "T%", "U%", "V%", "W%")
    '    iterType = Array("ItemNumber", " LIKE ", "'")
    'ElseIf item = "" Then
    '    iterArray = Array(componentID)
    '    iterType = Array("ComponentID", "=", "")
    'Else
    '    iterArray = Array(item)
    '    iterType = Array("ItemNumber", "=", "'")
    'End If
    'Dim iter As Variant
    Dim rst As ADODB.Recordset
    'For Each iter In iterArray
    '    If UBound(iterArray) <> 0 Then
    '        Sleep 30000
    '    End If
    '    DoEvents
        Dim xmlstr As String
        xmlstr = ""
    '    Set rst = DB.retrieve("SELECT ItemNumber, ComponentID, Quantity, Length, Width, Height, Weight, BoxLength, BoxWidth, BoxHeight, BoxWeight, PackingType, HazMat, NMFC FROM vVPSComponents WHERE " & iterType(0) & iterType(1) & iterType(2) & iter & iterType(2))
        Set rst = DB.retrieve("SELECT ItemNumber, " & _
                              "       ComponentID, " & _
                              "       Quantity, " & _
                              "       Length, " & _
                              "       Width, " & _
                              "       Height, " & _
                              "       Weight, " & _
                              "       COALESCE(BoxLength,0) AS BoxLength, " & _
                              "       COALESCE(BoxWidth,0) AS BoxWidth, " & _
                              "       COALESCE(BoxHeight,0) AS BoxHeight, " & _
                              "       COALESCE(BoxWeight,0) AS BoxWeight, " & _
                              "       PackingType, " & _
                              "       HazMat, " & _
                              "       COALESCE(NMFC,'') AS NMFC " & _
                              "FROM vVPSComponents " & _
                              "WHERE ItemNumber='" & item & "' " & _
                              IIf(componentID = "", "", "AND ComponentID=" & componentID))
        While Not rst.EOF
            Dim currxml As String
            currxml = "<Component>"
            Dim i As Long
            For i = 0 To rst.fields.count - 1
                If rst(i).name = "HazMat" Then
                    currxml = currxml & "<" & rst(i).name & ">" & IIf(rst(i).value = "True", "1", "0") & "</" & rst(i).name & ">"
    '            ElseIf IsNull(rst(i)) Then
    '                If rst(i).name = "BoxLength" Or rst(i).name = "BoxWidth" Or rst(i).name = "BoxHeight" Or rst(i).name = "BoxWeight" Then
    '                    currxml = currxml & "<" & rst(i).name & ">0</" & rst(i).name & ">"
    '                Else
    '                    currxml = currxml & "<" & rst(i).name & " />"
    '                End If
                Else
                    currxml = currxml & "<" & rst(i).name & ">" & escapeXMLEntities(removeUpperASCII(rst(i).value)) & "</" & rst(i).name & ">"
                End If
            Next i
            currxml = currxml & "</Component>" & vbLf
            xmlstr = xmlstr & currxml
            rst.MoveNext
        Wend
        
        If xmlstr <> "" Then
            xmlstr = "<Components>" & vbCrLf & xmlstr & "</Components>"
            Dim currErrCount As Long, currDone As Boolean
            currErrCount = 0
            currDone = False
            Dim req As MSXML2.XMLHTTP
            While Not currDone And currErrCount < 5
                Set req = New MSXML2.XMLHTTP
                req.Open "POST", "http://www.toolsplusinternal.com/internal/poinv/set_component_info.pl", False
                req.Send xmlstr
                'Dim filefree As Integer
                'filefree = FreeFile
                'Open "c:\users\tomwestbrook\desktop\component_xml.txt" For Output As #filefree
                'Print #filefree, xmlstr
                'Close #filefree
                
                'Debug.Print iter & " " & req.status
                While req.ReadyState <> 4
                    'Debug.Print req.ReadyState
                    Sleep 500
                    DoEvents
                Wend
                If req.status = 200 Then
                    currDone = True
                Else
                    'transient 500 error? try again
                    currErrCount = currErrCount + 1
                    Sleep 10000
                    DoEvents
                End If
            Wend
            If Not currDone Then
    '            errorcount = errorcount + 1
    '            errorstring = errorstring & iter & ": " & req.responseText & vbCrLf
    '        Else
    '            errorstring = errorstring & iter & ": " & req.status & vbCrLf
                Err.Raise "124", "ExternalComponentDBSync", "Error sending data for " & item & ", error code was " & req.status
            End If
        Else
            If rst.RecordCount = 0 Then
                'ok
            Else
                'hmm
    '            errorcount = errorcount + 1
    '            errorstring = errorstring & "No xml generated for " & qq(CStr(iter)) & vbCrLf
                Err.Raise "123", "ExternalComponentDBSync", "No xml generated for " & item
            End If
        End If
        
        rst.Close
    'Next iter
    '
    'If errorcount = 0 Then
    '    ExternalComponentDBSync = True
    'Else
    '    SendEmailTo "brian@tools-plus.com", "Database Sync Report (Component)", errorcount & " error(s) occurred." & vbCrLf & vbCrLf & errorstring
    '    ExternalComponentDBSync = False
    'End If
    Exit Function

errh:
    'SendEmailTo "brian@tools-plus.com", "Database Sync Report (Component)", "unexpected error(s) occurred." & vbCrLf & vbCrLf & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & errorstring
    SendEmailTo "brian@tools-plus.com", "Database Sync Report (Component)", "unexpected error(s) occurred." & vbCrLf & vbCrLf & Err.Number & vbCrLf & Err.Description
    Exit Function
End Function

Public Function ExternalComponentDBRemove(item As String, componentID As String) As Boolean
    Dim errorcount As Long
    Dim errorstring As String
On Error GoTo errh

    Dim xmlstr As String
    xmlstr = "<Components><Component>"
    xmlstr = xmlstr & "<ItemNumber>" & item & "</ItemNumber>"
    xmlstr = xmlstr & "<ComponentID>" & componentID & "</ComponentID>"
    xmlstr = xmlstr & "<Remove />"
    xmlstr = xmlstr & "</Component></Components>" & vbLf
    
    Dim currErrCount As Long, currDone As Boolean
    currErrCount = 0
    currDone = False
    Dim req As MSXML2.XMLHTTP
    While Not currDone And currErrCount < 5
        Set req = New MSXML2.XMLHTTP
        req.Open "POST", "http://www.toolsplusinternal.com/internal/poinv/set_component_info.fcgi", False
        req.Send xmlstr
        While req.ReadyState <> 4
            'Debug.Print req.ReadyState
            Sleep 500
            DoEvents
        Wend
        If req.status = 200 Then
            currDone = True
        Else
            'transient 500 error? try again
            currErrCount = currErrCount + 1
            Sleep 60000
        End If
    Wend
    
    If Not currDone Then
        errorcount = errorcount + 1
        errorstring = errorstring & "Remove" & item & "_" & componentID & ": " & req.responseText & vbCrLf
    Else
        errorstring = errorstring & "Remove" & item & "_" & componentID & ": " & req.status & vbCrLf
    End If
    
    If errorcount = 0 Then
        ExternalComponentDBRemove = True
    Else
        SendEmailTo "brian@tools-plus.com", "Database Sync Report (Component Remove)", errorcount & " error(s) occurred." & vbCrLf & vbCrLf & errorstring
        ExternalComponentDBRemove = False
    End If
    Exit Function

errh:
    SendEmailTo "brian@tools-plus.com", "Database Sync Report (Component Remove)", "unexpected error(s) occurred." & vbCrLf & vbCrLf & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & errorstring
    Exit Function
End Function


Public Function EstimateFreightForShipment(itemHash As Dictionary, ZipCode As String, Optional State As String = "") As Dictionary
    'itemHash should be itemnumber => qty
    'return value is service name => price, we only look for ups ground and usps priority/first-class mail
    
    Set EstimateFreightForShipment = Nothing 'assume the worst
    
    If State = "" Then
        State = DLookup("State", "LookupCityZipCode", "ZipCode='" & EscapeSQuotes(ZipCode) & "'")
        State = Replace(State, " ", "")
        If State = "" Then
            State = InputBox("Unable to determine which state this is, please enter state abbreviation here:")
        End If
        If State = "" Then
            Exit Function
        End If
    End If
    
    Dim i As Long
    i = 0
    Dim iter As Variant
    Dim query As String
    query = "type=item"
    For Each iter In itemHash.keys
        query = query & "&item" & i & "=" & CStr(iter) & "&qty" & i & "=" & CStr(CLng(itemHash.item(iter)))
    Next iter
    query = query & "&zipcode=" & ZipCode & "&state=" & State & "&country=US&residential=1&carrier0=UPS&carrier1=USPS"
   
    Dim req As MSXML2.XMLHTTP
    Set req = New MSXML2.XMLHTTP
    Dim content As String
    req.Open "GET", "http://toolsplus04/whse/freightquote.plex?" & query, False
    req.Send ""
    While req.ReadyState <> 4
        'Debug.Print req.ReadyState
        Sleep 100
        DoEvents
    Wend
    If req.status = 200 Then
        content = req.responseText
        
        
    Else
        MsgBox "Error: " & req.status & " " & req.StatusText & vbCrLf & vbCrLf & req.responseText
        Exit Function
    End If
    
    
    Dim parser As VBJSON.JSONParser
    Set parser = New VBJSON.JSONParser
    Dim retval As Dictionary
    Set retval = parser.Decode(content)
    If retval Is Nothing Then
        MsgBox "Could not parse JSON return value!"
        Exit Function
    ElseIf Not retval.exists("carriers") Then
        MsgBox "Expected 'carriers' param in JSON return value"
        Exit Function
    Else
        Set EstimateFreightForShipment = New Dictionary
        If retval.item("carriers").exists("UPS") Then
            If retval.item("carriers").item("UPS").exists("results") Then
                If retval.item("carriers").item("UPS").item("results").exists("03") Then
                    EstimateFreightForShipment.Add "UPS Ground", CDec(retval.item("carriers").item("UPS").item("results").item("03").item("negotiated_rate"))
                End If
            End If
        End If
        If retval.item("carriers").exists("USPS") Then
            If retval.item("carriers").item("USPS").exists("results") Then
                For Each iter In retval.item("carriers").item("USPS").item("results").keys
                    Dim svc As String
                    svc = retval.item("carriers").item("USPS").item("results").item(CStr(iter)).item("service")
                    'regex "Priority Mail \d{1}-Day"
                    If Len(svc) = 19 Then
                        If Left(svc, 14) = "Priority Mail " Then
                            If Right(svc, 4) = "-Day" Then
                                If Utilities.IsDigit(Mid(svc, 15, 1)) Then
                                    EstimateFreightForShipment.Add "USPS Priority Mail", CDec(retval.item("carriers").item("USPS").item("results").item(CStr(iter)).item("negotiated_rate"))
                                End If
                            End If
                        End If
                    ElseIf svc = "First-Class Mail Parcel" Then
                        EstimateFreightForShipment.Add "USPS First-Class Mail", CDec(retval.item("carriers").item("USPS").item("results").item(CStr(iter)).item("negotiated_rate"))
                    End If
                    
                Next iter
            End If
        End If
    End If
End Function



Public Sub reboxeverything()
    Set Mouse = New MouseHourglass
    Set DB = New DBConn.DatabaseSingleton

    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemID, ItemNumber FROM vItemStatusBreakdown WHERE DC=0 AND ItemNumber<'X%'")
    While Not rst.EOF
        Dim rst2 As ADODB.Recordset
        Set rst2 = DB.retrieve("SELECT InventoryComponents.ID, InventoryComponents.PackingType, InventoryComponents.ShippingBoxID, InventoryComponents.Length, InventoryComponents.Width, InventoryComponents.Height, InventoryComponents.Weight FROM InventoryComponentMap INNER JOIN InventoryComponents ON InventoryComponentMap.ComponentID=InventoryComponents.ID WHERE InventoryComponentMap.ItemID=" & rst("ItemID"))
        While Not rst2.EOF
            If rst2("PackingType") = 0 Then
                Dim newbox As String
                newbox = DetermineBoxSize2(CDbl(rst2("Length")), CDbl(rst2("Width")), CDbl(rst2("Height")), CDbl(rst2("Weight")))
                If newbox = rst2("ShippingBoxID") Then
                    'skip
                ElseIf newbox = -1 And IsNull(rst2("ShippingBoxID").value) Then
                    'skip
                Else
                    DB.Execute "UPDATE InventoryComponents SET ShippingBoxID=" & newbox & " WHERE ID=" & rst2("ID")
                    ExternalComponentDBSync rst("ItemNumber")
                    Debug.Print rst("ItemNumber")
                End If
            End If
            rst2.MoveNext
        Wend
        rst2.Close
        Set rst2 = Nothing

        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub
