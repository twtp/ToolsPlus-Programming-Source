Attribute VB_Name = "OfficeIntegration"
'---------------------------------------------------------------------------------------
' Module    : OfficeIntegration
' DateTime  : 2/2/2006 10:27
' Author    : briandonorfio
' Purpose   : Functions to interact w/ outlook, excel, access.
'
'             Dependencies:
'               - Microsoft Outlook 9.0 Object Library
'               - Microsoft Excel 9.0 Object Library
'               - Microsoft Access 9.0 Object Library
'               - Microsoft ActiveX Data Objects 2.8 Library
'               - Microsoft Scripting Runtime
'               - HashTables
'               - MouseHourglass + global Mouse object
'---------------------------------------------------------------------------------------

Option Explicit

Private accessApp As Access.Application

'---------------------------------------------------------------------------------------
' Procedure : OpenReportsDB
' DateTime  : 7/27/2005 13:23
' Author    : briandonorfio
' Purpose   : Checks if the reports database is already open, opens it if not already
'             running, otherwise switches to it. If file is not found, error message
'             displayed.
'---------------------------------------------------------------------------------------
'
Public Sub OpenReportsDB()
    If Not doesDBExist(REPORTS_DB) Then
        MsgBox "Unable to find " & DISTRIBUTION_DIR & REPORTS_DB & "!"
        Exit Sub
    End If
    Mouse.Hourglass True
    CopyOverIfNewer DISTRIBUTION_DIR & REPORTS_DB, DESTINATION_DIR & REPORTS_DB
    
    openDB REPORTS_DB
    
    accessApp.Visible = True
    
    Set accessApp = Nothing 'releases handle, otherwise closing access just makes it not visible
    
    Mouse.Hourglass False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : OpenPurchaseOrderReport
' DateTime  : 7/27/2005 13:23
' Author    : briandonorfio
' Purpose   : Checks if the PO database is already open, opens it if not found. Opens
'             the PO report, if anything else is ever added to the mdb, might want to
'             change this?
'---------------------------------------------------------------------------------------
'
Public Sub OpenPurchaseOrderReport(POID As String)
    If Not doesDBExist(PO_DB) Then
        MsgBox "Unable to find " & DISTRIBUTION_DIR & PO_DB & "!"
        Exit Sub
    End If
    Mouse.Hourglass True
    CopyOverIfNewer DISTRIBUTION_DIR & PO_DB, DESTINATION_DIR & PO_DB
    
    openDB PO_DB

    accessApp.DoCmd.OpenReport "RptPurchaseOrder", acViewPreview, , "ID=" & POID
    accessApp.DoCmd.Maximize
    accessApp.Visible = True
    
    Set accessApp = Nothing
    
    Mouse.Hourglass False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UpdateCachedSigns
' DateTime  : 7/27/2005 13:24
' Author    : briandonorfio
' Purpose   : Caches list of all reports (theoretically signs only) from databases
'             SIGNS_DB and MISC_DB. Any additions/deletions in these mdbs need this
'             update to be run for the app to know about the signs.
'               3 categories:
'                 Regular signs (from the SIGNS_DB mdb, not prefaced with "GROUP") go in
'                   the SignNames table.
'                 Group signs (also from SIGNS_DB, should start with "GROUP") and are
'                   stored in the GroupSigns table. The "GROUP" is removed when entered
'                   in the table.
'                 Miscellaneous signs are stored in the MISC_DB, since they're generally
'                   larger, so the split mdb is quicker to copy over in most cases.
'                   Prefaced with MISC, removed for the user, stored in MiscSigns table.
'---------------------------------------------------------------------------------------
'
'Public Function UpdateCachedSigns() As Boolean
'    Mouse.Hourglass True
'
'    updateCachedRegularSigns
'    updateCachedMiscSigns
'    CreateHashes "SignHash"
'
'    Mouse.Hourglass False
'
'End Function

'---------------------------------------------------------------------------------------
' Procedure : OpenSignInAccess
' DateTime  : 7/27/2005 13:25
' Author    : briandonorfio
' Purpose   : Opens specified sign, in given view mode (view modes are part of enum
'             acViewMode). whereClause specifies which items to show on sign, if none
'             given, then access defaults to every item. Should usually be something
'             similar to ItemNumber='" & itemNumber & "'" or for all with a given sign
'             SignName='" & SignName & "'". Check the AccessSignInformation query for
'             acceptable fields to use.
'
'             if keepHandle is true, then the access database will still be considered
'             under VB's control at the end of the function. Only use this if you're
'             opening multiple signs all in the same db, and remember to close it after.
'             Optional DiscountPct lets you print the signs with a discount applied,
'             this only changes the price, no other changes. DiscountPct=10 means
'             0.9*SellPrice
'---------------------------------------------------------------------------------------
'
Public Sub OpenSignInAccess(SignName As String, viewMode As Long, Optional whereClause As String = "", Optional keepHandle As Boolean = False, Optional overridePrintSignFlag As Boolean = False, Optional discountPct As Long = 0)
    'if we don't already have access open
    If accessApp Is Nothing Then
        Dim dbName As String
        If Left(SignName, 4) = "MISC" Then
            dbName = MISC_DB
        Else
            dbName = SIGNS_DB
        End If
        If Not doesDBExist(dbName) Then
            MsgBox "Unable to find " & DISTRIBUTION_DIR & dbName & "!"
            Exit Sub
        End If
        Mouse.Hourglass True
        CopyOverIfNewer DISTRIBUTION_DIR & dbName, DESTINATION_DIR & dbName
        openDB dbName
    Else
        Mouse.Hourglass True
    End If
    
On Error GoTo errh
    If doesReportExist(SignName) Then
        CloseReportIfOpen SignName
        accessApp.DoCmd.OpenReport SignName, acViewDesign, , IIf(overridePrintSignFlag, "", "PrintSign=-1 AND ") & whereClause
        If discountPct = 0 Then
            If accessApp.Reports(SignName).RecordSource <> "SignInformation" Then     'check for wrong recordsource
                accessApp.Reports(SignName).RecordSource = "SignInformation"
            End If
        Else
            If accessApp.Reports(SignName).RecordSource <> "ShowSignInfo" Then
                accessApp.Reports(SignName).RecordSource = "ShowSignInfo"
            End If
        End If
        accessApp.DoCmd.Close
        accessApp.DoCmd.OpenReport SignName, viewMode, , IIf(overridePrintSignFlag, "", "PrintSign=-1 AND ") & whereClause
        If viewMode = acViewNormal Then
            accessApp.DoCmd.Close
            accessApp.DoCmd.Quit
            Set accessApp = Nothing
        Else
            accessApp.Visible = True
        End If
    Else
        MsgBox "The sign you selected isn't in the Access database."
    End If
    
    If Not keepHandle Then
        Set accessApp = Nothing
    End If
    
    Mouse.Hourglass False
    Exit Sub
    
errh:
    If Err.Number = 2501 Then
        'do nothing
        Resume Next
    Else
        MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    End If
End Sub

Public Sub PrintSignByQOH(SignName As String, viewMode As Long, Optional whereClause As String = "")
    If Not doesDBExist(SIGNS_DB) Then
        MsgBox "Unable to find " & DISTRIBUTION_DIR & SIGNS_DB & "!"
        Exit Sub
    End If
    Mouse.Hourglass True
    CopyOverIfNewer DISTRIBUTION_DIR & SIGNS_DB, DESTINATION_DIR & SIGNS_DB
    
    openDB SIGNS_DB
On Error GoTo errh
    If doesReportExist(SignName) Then
        CloseReportIfOpen SignName
        'set up a temporary table for the report.
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT AccessSignInformation.*, vActualWhseQty.QtyOnHand AS QtyOnHand FROM AccessSignInformation INNER JOIN vActualWhseQty ON AccessSignInformation.ItemNumber=vActualWhseQty.ItemNumber WHERE " & Replace(whereClause, "=-1", "=1"))
        DB.Execute "TRUNCATE TABLE TempPrintSigns"
        While Not rst.EOF
            Dim i As Long
            For i = 0 To rst("QtyOnHand") - 1
                DB.Execute "INSERT INTO TempPrintSigns ( LineNum, ItemNumber, LineCode, PartNumber, ItemDescription, SignName, Desc1, Desc3, Spec1, Spec2, Spec3, Spec4, Spec5, Spec6, Spec7, Spec8, ExtendedSpecs, SellPrice, ListPrice, BreakQty2, BreakPrice2, BreakQty3, BreakPrice3, BreakQty4, BreakPrice4, BreakQty5, BreakPrice5, PriceChanged, PrintSign, LastChanged, TriadCode, YahooID ) VALUES ( " & _
                           i & ", '" & rst("ItemNumber") & "', '" & rst("LineCode") & "', '" & rst("PartNumber") & "', '" & EscapeSQuotes(rst("ItemDescription")) & "', '" & EscapeSQuotes(rst("SignName")) & "', '" & EscapeSQuotes(Nz(rst("Desc1"))) & "', '" & EscapeSQuotes(Nz(rst("Desc3"))) & "', '" & EscapeSQuotes(Nz(rst("Spec1"))) & "', '" & EscapeSQuotes(Nz(rst("Spec2"))) & "', '" & EscapeSQuotes(Nz(rst("Spec3"))) & "', '" & EscapeSQuotes(Nz(rst("Spec4"))) & "', '" & EscapeSQuotes(Nz(rst("Spec5"))) & "', '" & EscapeSQuotes(Nz(rst("Spec6"))) & "', '" & EscapeSQuotes(Nz(rst("Spec7"))) & _
                           "', '" & EscapeSQuotes(Nz(rst("Spec8"))) & "', '" & EscapeSQuotes(Nz(rst("ExtendedSpecs"))) & "', '" & rst("SellPrice") & "', '" & rst("ListPrice") & "', " & rst("BreakQty2") & ", '" & rst("BreakPrice2") & "', " & rst("BreakQty3") & ", '" & rst("BreakPrice3") & "', " & rst("BreakQty4") & ", '" & rst("BreakPrice4") & "', " & rst("BreakQty5") & ", '" & rst("BreakPrice5") & "', " & SQLBool(rst("PriceChanged")) & _
                           ", " & SQLBool(rst("PrintSign")) & ", '" & rst("LastChanged") & "', '" & EscapeSQuotes(Nz(rst("TriadCode"))) & "', '" & Nz(rst("YahooID")) & "' )"
            Next i
            rst.MoveNext
        Wend
        rst.Close
        'now open the report with the table as the recordsource
        accessApp.DoCmd.OpenReport SignName, acViewDesign, , whereClause        'open it, so we can change the recordsource
        If accessApp.Reports(SignName).RecordSource <> "SignsByQOH" Then
            accessApp.Reports(SignName).RecordSource = "SignsByQOH"             'change to the right one
        End If
        accessApp.DoCmd.Close                                                   'need to close, otherwise it doesn't set the where clause
        accessApp.DoCmd.OpenReport SignName, viewMode, , whereClause            'reopen it for preview/print
        If viewMode = acViewNormal Then                                         'don't change it back on close, just check next time?
            accessApp.DoCmd.Close
            accessApp.DoCmd.Quit
            Set accessApp = Nothing
        Else
            accessApp.Visible = True
        End If
    Else
        MsgBox "The sign you selected isn't in the Access database."
    End If
    
    Set accessApp = Nothing
    
    Mouse.Hourglass False
    Exit Sub

errh:
    If Err.Number = 2501 Then
        'do nothing
        Resume Next
    Else
        MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PreviewGroupSign
' DateTime  : 8/5/2005 16:31
' Author    : briandonorfio
' Purpose   : Opens a group sign for previewing
'---------------------------------------------------------------------------------------
'
Public Sub PreviewGroupSign(GroupSign As String)
    If Not doesDBExist(SIGNS_DB) Then
        MsgBox "Unable to find " & DISTRIBUTION_DIR & SIGNS_DB & "!"
        Exit Sub
    End If
    Mouse.Hourglass True
    CopyOverIfNewer DISTRIBUTION_DIR & SIGNS_DB, DESTINATION_DIR & SIGNS_DB
    
    openDB SIGNS_DB
    
    accessApp.DoCmd.OpenReport "GROUP " & GroupSign, acViewPreview, , "PrintSign=-1 AND GroupName='" & GroupSign & "'"
    accessApp.Visible = True
    
    Set accessApp = Nothing
    
    Mouse.Hourglass False
End Sub




'---------------------------------------------------------------------------------------
' Procedure : doesDBExist
' DateTime  : 7/27/2005 13:27
' Author    : briandonorfio
' Purpose   : Checks for a database file, returns true if file exists, false otherwise.
'---------------------------------------------------------------------------------------
'
Private Function doesDBExist(DB As String) As Boolean
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    doesDBExist = fso.FileExists(DISTRIBUTION_DIR & DB)
    Set fso = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : doesReportExist
' DateTime  : 7/27/2005 13:27
' Author    : briandonorfio
' Purpose   : Checks for a report in the currently open accessApp database, returns true
'             if report exists, false otherwise.
'---------------------------------------------------------------------------------------
'
Private Function doesReportExist(repName As String) As Boolean
On Error Resume Next
    Dim foo As String
    foo = accessApp.CurrentDb.Containers("Reports").Documents(repName).name
    If Err.Number = 3265 Or foo = "" Then
        doesReportExist = False
    Else
        doesReportExist = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : doesFormExist
' DateTime  : 8/31/2005 12:46
' Author    : briandonorfio
' Purpose   : Checks for a form, returns true if it exists, false otherwise.
'---------------------------------------------------------------------------------------
'
Private Function doesFormExist(formName As String) As Boolean
On Error Resume Next
    Dim foo As String
    foo = accessApp.CurrentDb.Containers("Forms").Documents(formName).name
    If Err.Number = 3265 Or foo = "" Then
        doesFormExist = False
    Else
        doesFormExist = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : openDB
' DateTime  : 8/5/2005 15:09
' Author    : briandonorfio
' Purpose   : Opens an Access Database if not already open. If open, switches to it.
'---------------------------------------------------------------------------------------
'
Private Sub openDB(DB As String, Optional from_dist As Boolean = False)
    If accessApp Is Nothing Then
        Set accessApp = New Access.Application
        accessApp.OpenCurrentDatabase (IIf(from_dist, DISTRIBUTION_DIR, DESTINATION_DIR) & DB)
    Else
        On Error GoTo errh
        If accessApp.CurrentProject.name <> DB Then
            Set accessApp = New Access.Application
            accessApp.OpenCurrentDatabase (IIf(from_dist, DISTRIBUTION_DIR, DESTINATION_DIR) & DB)
        End If
        On Error GoTo 0
    End If
    Exit Sub
    
errh:
    If Err.Number = 2467 Or Err.Number = 462 Then
        Set accessApp = Nothing 'closed by user, null the pointer and retry
        Set accessApp = New Access.Application
        accessApp.OpenCurrentDatabase (IIf(from_dist, DISTRIBUTION_DIR, DESTINATION_DIR) & DB)
        Err.Clear
        Resume
    Else
        MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : closeDB
' DateTime  : 8/5/2005 15:09
' Author    : briandonorfio
' Purpose   : Close the open Access Database. Shouldn't crash if not open, but no
'             message either.
'---------------------------------------------------------------------------------------
Public Sub closeDB()
    If Not accessApp Is Nothing Then
        accessApp.DoCmd.Close
        Set accessApp = Nothing
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : dropHandle
' DateTime  : 3/14/2006 08:33
' Author    : briandonorfio
' Purpose   : Just drops the handle, doesn't close the Access DB. After this is called,
'             the OS is supposed to clean up after it's closed by the user, not our
'             problem anymore.
'---------------------------------------------------------------------------------------
'
Public Sub dropHandle()
    If Not accessApp Is Nothing Then
        Set accessApp = Nothing
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CopyOverIfNewer
' DateTime  : 8/15/2005 15:50
' Author    : briandonorfio
' Purpose   : Checks date of last modification for destination file, overwrites with
'             from file if from is newer. Probably crashes if from file not found, test
'             with doesdbexist() first.
'---------------------------------------------------------------------------------------
'
Private Sub CopyOverIfNewer(from As String, dest As String)
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    If fso.FileExists(dest) Then
        If fso.GetFile(from).DateLastModified > fso.GetFile(dest).DateLastModified Then
            fso.CopyFile from, dest, True
        End If
    Else
        fso.CopyFile from, dest, True
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CloseReportIfOpen
' DateTime  : 8/17/2005 12:19
' Author    : briandonorfio
' Purpose   : Checks the currently open accessApp for a report with rptname. If it is
'             open, closes it, otherwise, no action taken.
'---------------------------------------------------------------------------------------
'
Private Sub CloseReportIfOpen(rptname As String)
    Dim rpt As Access.Report
    For Each rpt In accessApp.Reports
        If rpt.name = rptname Then
            accessApp.DoCmd.Close acReport, rpt.name, acSaveNo
        End If
    Next rpt
End Sub

'---------------------------------------------------------------------------------------
' Procedure : updateCachedMiscSigns
' DateTime  : 10/14/2005 11:35
' Author    : briandonorfio
' Purpose   : Refreshes the cached list in MiscSigns with the current list of reports in
'             the MISC_DB database.
'---------------------------------------------------------------------------------------
'
'Private Function updateCachedMiscSigns() As Boolean
'    If Not doesDBExist(MISC_DB) Then
'        MsgBox "Unable to find " & DISTRIBUTION_DIR & MISC_DB & "!"
'        updateCachedMiscSigns = False
'        Exit Function
'    End If
'    CopyOverIfNewer DISTRIBUTION_DIR & MISC_DB, DESTINATION_DIR & MISC_DB
'
'    openDB MISC_DB
'
'    'delete non-existent signs from the table
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT ID, MiscSignName AS ReportName FROM MiscSigns")
'    While Not rst.EOF
'        If Not doesReportExist("MISC " & rst("ReportName")) Then
'            DB.Execute "DELETE FROM MiscSigns WHERE ID=" & rst("ID")
'        End If
'        rst.MoveNext
'    Wend
'    rst.Close
'
'    'store the list of signs that exist, and we already know about
'    Dim tempSignH As Dictionary
'    Set tempSignH = New Dictionary
'    Set rst = DB.retrieve("SELECT ID, MiscSignName AS ReportName FROM MiscSigns")
'    While Not rst.EOF
'        tempSignH.Add "MISC " & rst("ReportName").value, rst("ID").value
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'
'    'loop through each report, if MISC* and we don't know about it, add it
'    Dim SIGN As Variant
'    For Each SIGN In accessApp.CurrentProject.AllReports
'        If Left(SIGN.name, 4) = "MISC" Then
'            If Not tempSignH.exists(SIGN.name) Then
'                DB.Execute "INSERT INTO MiscSigns ( MiscSignName ) VALUES ( '" & Replace(SIGN.name, "MISC ", "", , 1) & "' )"
'            End If
'        End If
'    Next SIGN
'
'    Set tempSignH = Nothing
'    closeDB
'End Function

'---------------------------------------------------------------------------------------
' Procedure : updateCachedRegularSigns
' DateTime  : 10/14/2005 11:36
' Author    : briandonorfio
' Purpose   : Refreshes both SignNames and GroupSigns with the current list of reports
'             in the SIGNS_DB database. Same idea/flow as updateCachedMiscSigns.
'---------------------------------------------------------------------------------------
'
'Private Function updateCachedRegularSigns() As Boolean
'    If Not doesDBExist(SIGNS_DB) Then
'        MsgBox "Unable to find " & DISTRIBUTION_DIR & SIGNS_DB & "!"
'        updateCachedRegularSigns = False
'        Exit Function
'    End If
'    CopyOverIfNewer DISTRIBUTION_DIR & SIGNS_DB, DESTINATION_DIR & SIGNS_DB
'
'    openDB SIGNS_DB
'
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT * FROM vUpdateCachedSigns")
'    Dim repName As String
'    While Not rst.EOF
'        repName = rst("ReportName")
'        If rst("TableName") = "GroupSigns" Then
'            repName = "GROUP " & repName
'        End If
'        If Not doesReportExist(repName) Then
'            DB.Execute "DELETE FROM " & rst("TableName") & " WHERE ID=" & rst("ID")
'        End If
'        rst.MoveNext
'    Wend
'    rst.Close
'
'    Dim tempSignH As Dictionary
'    Set tempSignH = New Dictionary
'    Set rst = DB.retrieve("SELECT * FROM vUpdateCachedSigns")
'    While Not rst.EOF
'        repName = rst("ReportName")
'        If rst("TableName") = "GroupSigns" Then
'            repName = "GROUP " & repName
'        End If
'        tempSignH.Add repName, rst("ID").value
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'
'    Dim SIGN As Variant
'    For Each SIGN In accessApp.CurrentProject.AllReports
'        If Not tempSignH.exists(SIGN.name) Then
'            If Left(SIGN.name, 5) = "GROUP" Then
'                DB.Execute "INSERT INTO GroupSigns ( GroupName ) VALUES ( '" & Replace(SIGN.name, "GROUP ", "", , 1) & "' )"
'            Else
'                DB.Execute "INSERT INTO SignNames ( SignName ) VALUES ( '" & SIGN.name & "' )"
'            End If
'        End If
'    Next SIGN
'
'    Set tempSignH = Nothing
'    closeDB
'End Function

'---------------------------------------------------------------------------------------
' Procedure : OpenEmailTo
' DateTime  : 2/2/2006 10:27
' Author    : briandonorfio
' Purpose   : Displays an email, addressed to whatever is given in the first argument.
'             Could be an outlook alias/group or a regular email address, so no error
'             checking. Optional arguments to specify the subject, and body or htmlbody.
'             Prefers htmlbody over body if both are given, doesn't send multipart
'             format emails. On any error, just returns false, otherwise returns true.
'---------------------------------------------------------------------------------------
'
Public Function OpenEmailTo(Optional Address As String = "", Optional subject As String = "", Optional body As String = "", Optional htmlbody As String = "") As Boolean
On Error GoTo errh
    Dim outlookapp As New Outlook.Application
    Set outlookapp = New Outlook.Application
    Dim Email As Outlook.MailItem
    Set Email = outlookapp.CreateItem(olMailItem)
    If Address <> "" Then
        Email.To = Address
    End If
    If subject <> "" Then
        Email.subject = subject
    End If
    If htmlbody <> "" Then
        Email.BodyFormat = OlBodyFormat.olFormatHTML
        Email.htmlbody = htmlbody
    ElseIf body <> "" Then
        Email.BodyFormat = OlBodyFormat.olFormatPlain
        Email.body = body
    End If
    Email.Display
    Set Email = Nothing
    Set outlookapp = Nothing
    OpenEmailTo = True
    Exit Function
    
errh:
    MsgBox "Unable to open Outlook: " & Err.Description
    Err.Clear
    Set Email = Nothing
    Set outlookapp = Nothing
    OpenEmailTo = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : AddOutlookContact
' DateTime  : 2/2/2006 11:01
' Author    : briandonorfio
' Purpose   : Adds a new contact to the Margie's Contacts public folder. Requires an
'             outlookContact struct (see globals.bas) with the information. fullName is
'             required, everything else is probably optional.
'---------------------------------------------------------------------------------------
'
Public Function AddOutlookContact(contactInfo As outlookContact)
On Error GoTo errh
    Dim outlookapp As Outlook.Application
    Set outlookapp = New Outlook.Application
    Dim contact As Outlook.ContactItem
    Dim folder As Outlook.MAPIFolder
    Set folder = outlookapp.GetNamespace("MAPI").Folders("Public Folders").Folders("All Public Folders").Folders("Margie's Contacts")
    
    Set contact = folder.items.Find("[fullName]=" & qq(contactInfo.fullName))
    If Not contact Is Nothing Then
        MsgBox "Contact already exists."
        contact.Display
    Else
        Set contact = folder.items.Add(olContactItem)
        contact.fullName = contactInfo.fullName
        contact.jobTitle = contactInfo.jobTitle
        contact.BusinessAddress = contactInfo.Address
        contact.BusinessAddressCity = contactInfo.City
        contact.BusinessAddressState = contactInfo.State
        contact.BusinessAddressPostalCode = contactInfo.ZipCode
        contact.BusinessTelephoneNumber = contactInfo.phoneNumber1 & IIf(contactInfo.phoneExtension1 <> "", " " & contactInfo.phoneExtension1, "")
        contact.Business2TelephoneNumber = contactInfo.phoneNumber2 & IIf(contactInfo.phoneExtension2 <> "", " " & contactInfo.phoneExtension2, "")
        contact.BusinessFaxNumber = contactInfo.faxNumber
        contact.MobileTelephoneNumber = contactInfo.cellPhoneNumber
        contact.Email1Address = contactInfo.emailAddress
        contact.WebPage = contactInfo.webAddress
        contact.body = contactInfo.body
        contact.Save
        MsgBox "Contact added. Check the ""File As"" in Outlook to make sure it shows up in the right place."
        contact.Display
    End If
    Set contact = Nothing
    Set folder = Nothing
    AddOutlookContact = True
    Exit Function
    
errh:
    MsgBox "Unable to add contact: " & Err.Description
    Err.Clear
    Set contact = Nothing
    Set folder = Nothing
    AddOutlookContact = False
End Function


'---------------------------------------------------------------------------------------
' Procedure : ListBoxAsXLS
' DateTime  : 4/12/2006 15:34
' Author    : briandonorfio
' Purpose   : Convert the contents of a listbox to an excel spreadsheet. The lines the
'             rows, and the tab-separations become cells. Optional argument to stop
'             partway through the listbox (line number as counted from zero). Returns
'             the excel application object.
'---------------------------------------------------------------------------------------
'
Public Function ListBoxAsXLS(list As ListBox, Optional stopAfterLine As Long = -1) As Excel.Application
    If stopAfterLine = -1 Then
        stopAfterLine = list.ListCount - 1
    End If
    Dim excelapp As Excel.Application
    Set excelapp = New Excel.Application
    excelapp.Workbooks.Add
    Dim line As Long, col As Long
    For line = 0 To stopAfterLine
        Dim cells As Variant
        cells = Split(list.list(line), vbTab)
        For col = 0 To UBound(cells)
            excelapp.ActiveWorkbook.ActiveSheet.Range(Chr(col + 65) & line + 1) = cells(col)
        Next col
    Next line
    excelapp.Visible = True
    Set ListBoxAsXLS = excelapp
    Set excelapp = Nothing
End Function
