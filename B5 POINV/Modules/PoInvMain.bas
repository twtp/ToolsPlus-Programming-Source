Attribute VB_Name = "PoInvMain"
'---------------------------------------------------------------------------------------
' Module    : PoInvMain
' DateTime  : 4/26/2006 10:22
' Author    : briandonorfio
' Purpose   : "It all comes from here, the stench and the peril." -- Frodo
'---------------------------------------------------------------------------------------

Option Explicit

Public Sub Main()

On Error GoTo errh
    If LCase(Environ("Username")) = "administrator" Then
        MsgBox "Get the hell out of here!" & vbCrLf & vbCrLf & "More explanation: the admin login has special permissions in sql server that can really break the database."
        Exit Sub
    End If

    'publically available database connection, everything for db access goes through here
    Set DB = New DBConn.DatabaseSingleton
    Set MASDB = New DBConn.DatabaseSingleton
    'Set UPCDB = New DBConn.DatabaseSingleton
    'Set SHIPDB = New DBConn.DatabaseSingleton
    DB.CurrentDb = DBENG_MSSQL
    'tom's code to test timeout
    DB.TimeoutPeriod = 120
    
    MASDB.CurrentDb = DBENG_MAS90
    
    'UPCDB.CurrentDb = DBENG_BARCODE
    'SHIPDB.CurrentDb = DBENG_SHIPPING
    
    'If USE_ALT_DATABASE Then
    '    DB.Catalog = "tptest"
    'End If
    
On Error GoTo errhDB
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT TOP 1 ItemNumber FROM InventoryMaster")
    rst.Close
    Set rst = Nothing
On Error GoTo errh
    
    CURRENT_WEBSITE_ID = 0 'tp
    META_NOINDEX_MOLLY_ON = True
    INV_MAINT_FOLLIES_EDIT = False 'global follies toggle


    'publically available mouse object from globals.bas, sets/unsets the hourglass
    Set Mouse = New MouseHourglass
    Mouse.Hourglass True
    
    'KEYWORDS_SERVICES = Array("Overture", "AdWords")
    
    'finds the user's temp directory
    'this should already exist from the copy process?
    DESTINATION_DIR = Environ("TEMP") & "\poinv\"
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    If Not fso.FolderExists(DESTINATION_DIR) Then
        fso.CreateFolder Environ("TEMP") & "\poinv"
    End If
    
    'sets CUR_POS_OPT and WARN_ON_IMPORT
    ReadConfFile
    
    'If Not RunningInIDE Then
    '    POINV_SESSION_ID = LogInPoInv(Environ("UserName"))
    'End If
    
    Set NotifyChanges = New PerlArray
    
    CreateHashes
    
    'If Environ("UserName") = "briandonorfio" Or Environ("UserName") = "eric" Then
        USING_SECURITY = True
    'Else
    '    USING_SECURITY = False
    'End If
    
    'Load MainMenu
    'MainMenu.Show
    Load Menu
    Menu.Show
    
    'If Not USE_ALT_DATABASE Then
        CheckAutoImport
    'End If
    
    Mouse.Hourglass False
    
    Exit Sub
    
errh:
    If Err.Number = 429 Or Err.Number = 430 Then
        MsgBox "ActiveX error " & Err.Number & " occurred, this is usually a version problem with a DLL/OCX." & vbCrLf & vbCrLf & "Try running Start PO-INV/Signs as administrator."
    Else
        MsgBox "Fatal error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    End If
    Mouse.Hourglass False
    End
errhDB:
    MsgBox "Error connecting to database! Is the server running?" & vbCrLf & vbCrLf & "Things to try:" & vbCrLf & _
           " 1. Reboot your computer" & vbCrLf & _
           " 2. Delete %windir%\system32\DBConn.dll" & vbCrLf & _
           " 3. Reboot server (toolsplus06)" & vbCrLf & _
           " 4. (Re-)Install MDAC 2.8" & vbCrLf & _
           " 5. Go home" & vbCrLf & vbCrLf & _
           "Full error message: Error #" & Err.Number & vbCrLf & vbCrLf & Err.Description
    End
End Sub

Public Sub cleanup()
    If NotifyChanges.Scalar <> 0 Then
        If MsgBox("Send email about price changes, etc.?", vbYesNo) = vbYes Then
            SendNotifyEmail
        End If
    End If

    Set ManufHashPLtoName = Nothing
    Set AvailHashIDtoStr = Nothing
    Set AvailHashStrToID = Nothing
    Set VendorHash = Nothing
    Set SignHashIDtoName = Nothing
    Set SignHashNameToID = Nothing
    'Set GroupSignHashIDtoName = Nothing
    'Set GroupSignHashNametoID = Nothing
    Set Mouse = Nothing
End Sub

Public Sub ReadConfFile()
    Dim conf As Dictionary
    Set conf = New Dictionary
    conf.Add "WarnOnImport", "1"
    conf.Add "CursorPosition", "select"
    If FileExists(CONF_FILE) Then
        Set conf = ParseConfFile(CONF_FILE)
    Else
        MsgBox "Can't load config file " & qq(CONF_FILE) & ", maybe your network drives are down?" & vbCrLf & vbCrLf & "Continuing with defaults..."
    End If
    If conf.exists("CursorPosition") Then
        CUR_POS_OPT = conf.item("CursorPosition")
    End If
    WARN_ON_IMPORT = True
    If conf.exists("WarnOnImport") Then
        WARN_ON_IMPORT = CBool(conf.item("WarnOnImport"))
    End If
    Set conf = Nothing
End Sub
