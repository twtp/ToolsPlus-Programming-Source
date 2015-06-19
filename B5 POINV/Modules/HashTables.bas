Attribute VB_Name = "HashTables"
'---------------------------------------------------------------------------------------
' Module    : HashTables
' DateTime  : 7/12/2005 15:33
' Author    : briandonorfio
' Purpose   : Creates a few lookup hash tables so we don't have to hit the database all
'             the time for simple things.
'
'             Dependencies:
'               - Microsoft Scripting Runtime
'               - MouseHourglass + global Mouse object
'---------------------------------------------------------------------------------------

Option Explicit

Public WebsiteNameHash As Dictionary
Public WebsiteURLHash As Dictionary

Public ManufHashPLtoName As Dictionary
Public ManufWebHashPLtoName As Dictionary

Public AvailHashIDtoStr As Dictionary
Public AvailHashStrToID As Dictionary

Public VendorHash As Dictionary

Public SignHashIDtoName As Dictionary
Public SignHashNameToID As Dictionary

'Public GroupSignHashIDtoName As Dictionary
'Public GroupSignHashNametoID As Dictionary

Public ItemStatusHashIDtoStr As Dictionary

Public EBayCategoryHashIDtoName As Dictionary
Public EBayCategoryHashIDtoConditionID As Dictionary

'Public ItemCostTypeHashNameToID As Dictionary

Public ApplicationModuleHashNameToID As Dictionary



'---------------------------------------------------------------------------------------
' Procedure : CreateHashes
' DateTime  : 7/12/2005 15:34
' Author    : briandonorfio
' Purpose   : Function to create the lookups, can also recreate a single lookup.
'             Currently does the following:
'               ManufHashPLtoName->{ProductLine} = ManufFullNameCleaned
'               ManufWebHashPLtoName->{ProductLine} = ManufFullNameWeb
'               AvailHashIDtoStr->{AvailID} = AvailString
'               AvailHashStrToID->{AvailString} = AvailID
'               VendorHash->{VendorNumber} = VendorName
'               SignHashIDtoName->{SignID} = SignName
'               SignHashNameToID->{SignName} = SignID
'               ItemStatusHashIDtoStr->{StatusID} = StatusString
'               EBayCategoryHashIDtoName->{CategoryID}->[0] = CategoryType (0==ebay,1==store)
'               EBayCategoryHashIDtoName->{CategoryID}->[1] = CategoryName
'               EBayCategoryHashIDtoConditionID->{CategoryID}->[0] = ConditionNewID
'               EBayCategoryHashIDtoConditionID->{CategoryID}->[1] = ConditionReconID
'               ItemCostTypeHashNameToID->{Readable} = ID
'               ApplicationModuleHashNameToID->{Readable} = ID
'---------------------------------------------------------------------------------------
'
Public Sub CreateHashes(Optional recreateHash As String = "")
    Mouse.Hourglass True

    Dim rst As ADODB.Recordset
    
    If recreateHash = "" Or recreateHash = "WebsiteHash" Then
        Set WebsiteNameHash = Nothing
        Set WebsiteURLHash = Nothing
        Set WebsiteNameHash = New Dictionary
        Set WebsiteURLHash = New Dictionary
        Set rst = DB.retrieve("SELECT WebsiteID, URL, Description FROM Websites")
        While Not rst.EOF
            WebsiteNameHash.Add CLng(rst("WebsiteID").value), rst("Description").value
            WebsiteURLHash.Add CLng(rst("WebsiteID").value), rst("URL").value
            rst.MoveNext
        Wend
        rst.Close
    End If

    If recreateHash = "" Or recreateHash = "ManufHash" Then
        Set ManufHashPLtoName = Nothing
        Set ManufWebHashPLtoName = Nothing
        Set ManufHashPLtoName = New Dictionary
        Set ManufWebHashPLtoName = New Dictionary
        Set rst = DB.retrieve("SELECT ProductLine, ManufFullNameCleaned, ManufFullNameWeb FROM ProductLine")
        While Not rst.EOF
            ManufHashPLtoName.Add rst("ProductLine").value, rst("ManufFullNameCleaned").value
            ManufWebHashPLtoName.Add rst("ProductLine").value, rst("ManufFullNameWeb").value
            rst.MoveNext
        Wend
        rst.Close
    End If
    
    If recreateHash = "" Or recreateHash = "AvailHash" Then
        Set AvailHashIDtoStr = Nothing
        Set AvailHashStrToID = Nothing
        Set AvailHashIDtoStr = New Dictionary
        Set AvailHashStrToID = New Dictionary
        Set rst = DB.retrieve("SELECT ID, ShortDesc FROM AvailOverrides")
        While Not rst.EOF
            AvailHashIDtoStr.Add rst("ID").value, rst("ShortDesc").value
            AvailHashStrToID.Add rst("ShortDesc").value, rst("ID").value
            rst.MoveNext
        Wend
        rst.Close
    End If

    If recreateHash = "" Or recreateHash = "VendorHash" Then
        Set VendorHash = Nothing
        Set VendorHash = New Dictionary
        Set rst = DB.retrieve("SELECT VendorNo, VendorName FROM AP_Vendor")
        While Not rst.EOF
            VendorHash.Add rst("VendorNo").value, rst("VendorName").value
            rst.MoveNext
        Wend
        rst.Close
    End If
        
    If recreateHash = "" Or recreateHash = "SignHash" Then
        Set SignHashIDtoName = Nothing
        Set SignHashIDtoName = New Dictionary
        Set SignHashNameToID = Nothing
        Set SignHashNameToID = New Dictionary
        Set rst = DB.retrieve("SELECT ID, SignName FROM SignNames")
        While Not rst.EOF
            SignHashIDtoName.Add rst("ID").value, rst("SignName").value
            SignHashNameToID.Add rst("SignName").value, rst("ID").value
            rst.MoveNext
        Wend
        rst.Close
    End If
    
    'If recreateHash = "" Or recreateHash = "SignHash" Then
    '    Set GroupSignHashIDtoName = Nothing
    '    Set GroupSignHashIDtoName = New Dictionary
    '    Set GroupSignHashNametoID = Nothing
    '    Set GroupSignHashNametoID = New Dictionary
    '    Set rst = DB.retrieve("SELECT ID, GroupName FROM GroupSigns")
    '    While Not rst.EOF
    '        GroupSignHashIDtoName.Add rst("ID").value, rst("GroupName").value
    '        GroupSignHashNametoID.Add rst("GroupName").value, rst("ID").value
    '        rst.MoveNext
    '    Wend
    '    rst.Close
    'End If
    
    If recreateHash = "" Or recreateHash = "ItemStatusHash" Then
        Set ItemStatusHashIDtoStr = Nothing
        Set ItemStatusHashIDtoStr = New Dictionary
        Set rst = DB.retrieve("SELECT ID, Status FROM InventoryItemStockStatus")
        While Not rst.EOF
            ItemStatusHashIDtoStr.Add rst("ID").value, rst("Status").value
            rst.MoveNext
        Wend
        rst.Close
    End If
    
    If recreateHash = "" Or recreateHash = "EBayHash" Then
        Set EBayCategoryHashIDtoName = Nothing
        Set EBayCategoryHashIDtoName = New Dictionary
        Set EBayCategoryHashIDtoConditionID = Nothing
        Set EBayCategoryHashIDtoConditionID = New Dictionary
        Set rst = DB.retrieve("SELECT CategoryID, CategoryType, CategoryName, ConditionNewID, ConditionReconID FROM EBayCategory")
        While Not rst.EOF
            EBayCategoryHashIDtoName.Add CStr(rst("CategoryID")), Array(CLng(rst("CategoryType")), CStr(rst("CategoryName")))
            EBayCategoryHashIDtoConditionID.Add CStr(rst("CategoryID")), Array(CLng(rst("ConditionNewID")), CLng(rst("ConditionReconID")))
            rst.MoveNext
        Wend
        rst.Close
    End If
    
    'If recreateHash = "" Or recreateHash = "ItemCostType" Then
    '    Set ItemCostTypeHashNameToID = Nothing
    '    Set ItemCostTypeHashNameToID = New Dictionary
    '    Set rst = DB.retrieve("SELECT ID, Readable FROM ItemCostType")
    '    While Not rst.EOF
    '        ItemCostTypeHashNameToID.Add CStr(rst("Readable")), CLng(rst("ID"))
    '        rst.MoveNext
    '    Wend
    '    rst.Close
    'End If
    
    If recreateHash = "" Or recreateHash = "ApplicationModule" Then
        Set ApplicationModuleHashNameToID = Nothing
        Set ApplicationModuleHashNameToID = New Dictionary
        Set rst = DB.retrieve("SELECT ID, Readable FROM ApplicationModule")
        While Not rst.EOF
            ApplicationModuleHashNameToID.Add CStr(rst("Readable")), CLng(rst("ID"))
            rst.MoveNext
        Wend
        rst.Close
    End If
    
    Set rst = Nothing
    Mouse.Hourglass False
End Sub
