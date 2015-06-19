Attribute VB_Name = "AutoImportMain"
Option Explicit

Public DB As DBConn.DatabaseSingleton
Public MASDB As DBConn.DatabaseSingleton

Public Sub Main()

    Set Mouse = New MouseHourglass
    
    Set DB = New DBConn.DatabaseSingleton
    DB.CurrentDB = DBENG_MSSQL
    Set MASDB = New DBConn.DatabaseSingleton
    MASDB.CurrentDB = DBENG_MAS90

    If DEBUG_MODE Then
        DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Loading' )"
    End If
    
    Dim updateFreight As Boolean, updateSalesRank As Boolean
    Dim triadCodeTime As String, triadCodeMetric As String
    Dim threshold As Dictionary
    Dim shopFeeds As Boolean, affFeed As Boolean, vendorQtyUpdate As Boolean
    Dim forceOPCalc As Boolean
    
    updateFreight = True
    updateSalesRank = False
    triadCodeTime = ""
    triadCodeMetric = ""
    Set threshold = New Dictionary
    threshold.Add "A", "70"
    threshold.Add "B", "80"
    threshold.Add "C", "95"
    threshold.Add "D", "100"
    shopFeeds = False
    affFeed = False
    forceOPCalc = False
    
    Dim conf As Dictionary
    Set conf = ParseConfFile("s:\mastest\mas90-signs\conf_files\AutomaticImport.conf")
    If conf.exists("FreightTable") Then
        updateFreight = CBool(conf.item("FreightTable"))
    End If
    If conf.exists("SalesRanking") Then
        updateSalesRank = CBool(conf.item("SalesRanking"))
    End If
    If conf.exists("TriadCodeTime") Then
        triadCodeTime = conf.item("TriadCodeTime")
    End If
    If conf.exists("TriadCodeMetric") Then
        triadCodeMetric = conf.item("TriadCodeMetric")
    End If
    If conf.exists("ThresholdA") Then
        threshold.item("A") = CStr(conf.item("ThresholdA"))
    End If
    If conf.exists("ThresholdB") Then
        threshold.item("B") = CStr(conf.item("ThresholdB"))
    End If
    If conf.exists("ThresholdC") Then
        threshold.item("C") = CStr(conf.item("ThresholdC"))
    End If
    If conf.exists("ThresholdD") Then
        threshold.item("D") = CStr(conf.item("ThresholdD"))
    End If
    If conf.exists("ShoppingFeeds") Then
        shopFeeds = conf.item("ShoppingFeeds")
    End If
    If conf.exists("AffiliateFeed") Then
        affFeed = conf.item("AffiliateFeed")
    End If
    If conf.exists("VendorQtyUpdate") Then
        vendorQtyUpdate = conf.item("VendorQtyUpdate")
    End If
    If conf.exists("ForceOPCalc") Then
        forceOPCalc = conf.item("ForceOPCalc")
    End If
    Set conf = Nothing

    Load Mas200AutoImportStatus
    Mas200AutoImportStatus.Show
    
    If DEBUG_MODE Then
        DB.Execute "INSERT INTO DebugAutoImport ( Label ) VALUES ( 'Starting' )"
    End If
    
    Mas200AutoImportStatus.doImport updateFreight, updateSalesRank, threshold, triadCodeTime, triadCodeMetric, forceOPCalc
    Unload Mas200AutoImportStatus
    
    If vendorQtyUpdate Then
        Shell PERL & " " & REFRESH_VENDOR_QTYS, vbNormalFocus
    End If
    
    If shopFeeds Then
        Shell PERL & " " & SHOPPING_FEEDS, vbNormalFocus
    End If
    
    'ExternalItemDBSync
    'ExternalComponentDBSync
    Shell PERL & " " & "s:\mastest\mas90-signs\sync_vps_database.pl", vbNormalFocus
End Sub
