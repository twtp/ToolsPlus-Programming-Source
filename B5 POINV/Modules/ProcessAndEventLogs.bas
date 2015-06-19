Attribute VB_Name = "ProcessAndEventLogs"
'---------------------------------------------------------------------------------------
' Module    : ProcessAndEventLogs
' DateTime  : 3/9/2006 12:04
' Author    : briandonorfio
' Purpose   : Functions dealing with logging and mutex flagging.
'
'             Dependencies:
'               - Microsoft ActiveX Data Objects 2.8 Library
'               - DatabaseFunctions + DBConn
'               - MouseHourglass + global Mouse object
'               - Utilities
'---------------------------------------------------------------------------------------

Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : ClearProcessFlag
' DateTime  : 8/19/2005 16:01
' Author    : briandonorfio
' Purpose   : Basic mutex unlocking for processes like the mas200 import/export.
'---------------------------------------------------------------------------------------
'
Public Sub ClearProcessFlag(task As String)
    DB.Execute "UPDATE ProcessFlags SET UserName='NOTRUNNING' WHERE Task='" & task & "'"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetProcessFlag
' DateTime  : 8/19/2005 16:02
' Author    : briandonorfio
' Purpose   : Basic mutex locking for processes like the mas200 import/export. Needs a
'             task name, can also supply a username to run as, defaults to windows logon
'             name.
'---------------------------------------------------------------------------------------
'
Public Sub SetProcessFlag(task As String, Optional username As String = "")
    DB.Execute "UPDATE ProcessFlags SET UserName='" & IIf(username = "", Environ("UserName"), username) & "' WHERE Task='" & task & "'"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsProcessFlagSet
' DateTime  : 9/6/2005 14:48
' Author    : briandonorfio
' Purpose   : Returns true/false if the given process is running.
'---------------------------------------------------------------------------------------
'
Public Function IsProcessFlagSet(task As String) As Boolean
    If DLookup("UserName", "ProcessFlags", "Task='" & task & "'") = "NOTRUNNING" Then
        IsProcessFlagSet = False
    Else
        IsProcessFlagSet = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : ImportExportProcessFlags
' DateTime  : 9/13/2005 09:44
' Author    : briandonorfio
' Purpose   : Gets the import and export flags in one db access. Returns a boolean array
'             with true/false if the process is running. Import in array(0), export in
'             array(1).
'---------------------------------------------------------------------------------------
'
Public Function ImportExportProcessFlags() As Boolean()
    Dim retval(1) As Boolean
    Dim rst As ADODB.Recordset
'don't let this kill the program if network dies
On Error GoTo errh
    Set rst = DB.retrieve("SELECT Task, UserName FROM ProcessFlags WHERE Task='export' or Task='import'")
    If rst Is Nothing Then
        GoTo statusquo
    End If
On Error GoTo 0
    While Not rst.EOF
        If rst("Task") = "import" Then
            retval(0) = Not CBool(rst("UserName") = "NOTRUNNING")
        ElseIf rst("Task") = "export" Then
            retval(1) = Not CBool(rst("UserName") = "NOTRUNNING")
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    ImportExportProcessFlags = retval
    Exit Function
    
errh:
    If Err.Number = -2147467259 Then
        'communication error, probably intermittent network problems
        'maybe they're just idling, so leave it alone
    Else
        MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    End If
statusquo:
    retval(0) = IMPORT_GOING
    retval(1) = EXPORT_GOING
    ImportExportProcessFlags = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : EventLogEntry
' DateTime  : 9/9/2005 13:23
' Author    : briandonorfio
' Purpose   : Puts an entry into the event log. Only thing needed is a name for the
'             event. Returns the record ID, should be used for updating event later.
'---------------------------------------------------------------------------------------
'
Public Function EventLogEntry(eventName As String) As Long
    DB.Execute "INSERT INTO EventLog ( Event ) VALUES ( '" & eventName & "' )"
    EventLogEntry = DLookup("MAX(ID)", "EventLog", "Event='" & eventName & "'") 'assumes it's always the last one added
End Function

'---------------------------------------------------------------------------------------
' Procedure : EventLogUpdate
' DateTime  : 9/9/2005 13:24
' Author    : briandonorfio
' Purpose   : Updates an entry in the event log, given by the eventID.
'---------------------------------------------------------------------------------------
'
Public Sub EventLogUpdate(eventID As Long, success As Boolean)
    DB.Execute "UPDATE EventLog SET TimeFinished=GETDATE(), Success=" & SQLBool(success) & " WHERE ID=" & eventID
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CheckAutoImport
' DateTime  : 9/14/2005 10:20
' Author    : briandonorfio
' Purpose   : Checks for "autoimport" in the EventLog, messages about success/failure.
'             If not run last night, warning message. If failure last night, warning
'             message. Message if successful?
'---------------------------------------------------------------------------------------
'
Public Sub CheckAutoImport()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT TimeStarted, TimeFinished, Success FROM EventLog WHERE Event='autoimport' AND TimeStarted>'" & Format(Now(), "Short Date") & "'")
    If rst.EOF Then
        Select Case Environ("UserName")
            Case Is = "briandonorfio", "margiebrochu", "eric"
                MsgBox "Autoimport was not found in the Event Log! No import done this morning?"
            Case Else
                'do nothing
        End Select
    Else
        If rst("Success") Then
            'MsgBox "Autoimport successful!" & vbCrLf & vbCrLf & "Started: " & Format(rst("TimeStarted"), "Long Time") & vbCrLf & "Finished: " & Format(rst("TimeFinished"), "Long Time")
        Else
            Select Case Environ("UserName")
                Case Is = "briandonorfio", "margiebrochu", "eric"
                    MsgBox "Autoimport started but not finished! Either crashed, or just not finished yet?", vbCritical
                Case Else
                    'do nothing
            End Select
        End If
    End If
    rst.Close
    Set rst = Nothing
End Sub

''---------------------------------------------------------------------------------------
'' Procedure : LogInPoInv
'' DateTime  : 3/6/2006 12:40
'' Author    : briandonorfio
'' Purpose   : Marks the username as logged into po/inv. Returns the ID, so we can log
''             out later.
''---------------------------------------------------------------------------------------
''
'Public Function LogInPoInv(username As String) As Long
'    'this is probably the first db access for the program, so we'll
'    'error handle+die here, if something's wrong
'    On Error GoTo errh
'    
'    Dim computerName As String, ipAddr As String
'    computerName = Environ("ComputerName")
'    ipAddr = GetCurrentIPAddress()
'    If Environ("ClientName") <> "" And Environ("ClientName") <> "Console" Then
'        computerName = computerName & "/" & Environ("ClientName")
'        ipAddr = ipAddr & "/" & GetTerminalServicesClientIPAddress()
'    End If
'    
'    DB.Execute "INSERT INTO PoInvSessions ( UserName, ComputerName, IPAddress ) VALUES ( '" & username & "', '" & computerName & "', '" & ipAddr & " ' )"
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT @@IDENTITY")
'    LogInPoInv = rst(0)
'    rst.Close
'    Set rst = Nothing
'    Exit Function
'errh:
'    MsgBox "Error connecting to database! Is the server running?" & vbCrLf & vbCrLf & "Things to try:" & vbCrLf & _
'            " 1. Reboot your computer" & vbCrLf & _
'            " 2. Delete %windir%\system32\DBConn.dll" & vbCrLf & _
'            " 3. Reboot server (toolsplus06)" & vbCrLf & _
'            " 4. (Re-)Install MDAC 2.8" & vbCrLf & _
'            " 5. Go home" & vbCrLf & vbCrLf & _
'            "Full error message: Error #" & Err.Number & vbCrLf & vbCrLf & Err.Description
'    End
'End Function
'
''---------------------------------------------------------------------------------------
'' Procedure : LogOutPoInv
'' DateTime  : 3/6/2006 12:41
'' Author    : briandonorfio
'' Purpose   : Given an ID from LogInPoInv, updates the record with the out-time. If the
''             program crashed, then this never gets called, so it looks like they're in
''             there still
''---------------------------------------------------------------------------------------
''
'Public Sub LogOutPoInv(loginId As Long)
'    DB.Execute "UPDATE PoInvSessions SET TimeOut=GETDATE() WHERE ID=" & loginId
'End Sub
