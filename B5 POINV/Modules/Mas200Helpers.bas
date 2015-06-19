Attribute VB_Name = "Mas200Helpers"
Option Explicit

Public Sub MasImportClick()
If RunningThroughVPN Then
    MsgBox "You can't run this through the VPN!"
    Exit Sub
End If
'If USE_ALT_DATABASE Then
'    Exit Sub
'End If
    If MsgBox("Are you sure you want to EXPORT and then IMPORT new data?", vbYesNo) = vbYes Then
        MsgBox "THE IMPORT IS BROKEN." & vbCrLf & vbCrLf & "until it gets fixed, run an autoimport manually:" & vbCrLf & qq("s:\mastest\mas90-signs\a_dist\automaticimport.exe"): Exit Sub
        IMPORT_GOING = True
        Load MasImportExportStatus
        MasImportExportStatus.Show
        'MasImportExportMenu.Hide
        MasImportExportStatus.doImport
        'MasImportExportMenu.Show
        Unload MasImportExportStatus
        IMPORT_GOING = False
    End If
End Sub

Public Sub MasExportClick()
If RunningThroughVPN Then
    MsgBox "You can't run this through the VPN!"
    Exit Sub
End If
If LCase(Environ("ComputerName")) <> "toolsplus04" Then
    If vbNo = MsgBox("This should really only be run on toolsplus04, are you sure you want to export locally?", vbYesNo) Then
        Exit Sub
    End If
End If
'If USE_ALT_DATABASE Then
'    Exit Sub
'End If
    If MsgBox("Are you sure you want to EXPORT changes to MAS 200?", vbYesNo) = vbYes Then
        EXPORT_GOING = True                 'we don't need to warn outselves about the export, just set it manually
        Load MasImportExportStatus          'mutex and logs are handled in-class, so we just need to set these bools
        MasImportExportStatus.Show
        'MasImportExportMenu.Hide
        MasImportExportStatus.doExport
        'MasImportExportMenu.Show
        Unload MasImportExportStatus
        EXPORT_GOING = False                'clear the warning flag
    End If
End Sub

Public Sub MasRetryClick()
If RunningThroughVPN Then
    MsgBox "You can't run this through the VPN!"
    Exit Sub
End If
'If USE_ALT_DATABASE Then
'    Exit Sub
'End If
    If MsgBox("Are you sure you want to RETRY the previous EXPORT?", vbYesNo) = vbYes Then
        EXPORT_GOING = True
        Load MasImportExportStatus
        MasImportExportStatus.Show
        'MasImportExportMenu.Hide
        MasImportExportStatus.doRetryExport
        'MasImportExportMenu.Show
        Unload MasImportExportStatus
        EXPORT_GOING = False
    End If
End Sub

Public Sub MasUnlockClick()
'If USE_ALT_DATABASE Then
'    Exit Sub
'End If
    If MsgBox("If you're sure that the import isn't running (crashed on last run, etc.), click ok, otherwise check with the last person to run it.", vbOKCancel) = vbOK Then
        IMPORT_GOING = False
        EXPORT_GOING = False
        ClearProcessFlag "import"
        ClearProcessFlag "export"
        If InStr(Menu.lblGeneric1.Caption, "import") Then
            DB.Execute "EXEC spImportEnableTriggers"
        End If
        Menu.btnSomething(0).Enabled = True
        Menu.btnSomething(1).Enabled = True
        Menu.btnSomething(2).Enabled = True
        Menu.lblGeneric1.Visible = False
        Menu.lblGeneric2.Visible = False
        Menu.btnGeneric.Visible = False
    End If
End Sub

Public Function ItemExistsInMas(item As String) As Boolean
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = MASDB.retrieve("SELECT COUNT(*) AS Occurrences FROM CI_Item WHERE ItemCode='" & item & "'")
    ItemExistsInMas = CBool(rst(0) = 1)
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Function
