VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatabaseSingleton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum DBEngines
    DBENG_MSSQL = 0
    DBENG_MAS90 = 1
    DBENG_BARCODE = 2
    DBENG_SHIPPING = 3
End Enum

Private Const MSSQL_CONN As String = "Provider=SQLNCLI;DataTypeCompatibility=80;Data Source=toolsplus06;Initial Catalog=**CATALOG**;Integrated Security=SSPI"
Private Const MSSQL_ALT_CONN As String = "Provider=sqloledb;Data Source=toolsplus06;Initial Catalog=**CATALOG**;Integrated Security=SSPI"
'Private Const MAS90_CONN As String = "Driver={Mas 90 32-Bit ODBC Driver};Directory=s:\mastest\mas200\version4\MAS90;Prefix=s:\mastest\mas200\version4\MAS90\soa;ViewDLL=s:\mastest\mas200\version4\MAS90\home;LogFile=\PVXODB32.LOG;SERVER=NotTheServer;UID=bd|TOO;PWD=brian"
'Private Const MAS90_CONN As String = "Driver={MAS 90 4.0 ODBC Driver};Directory=\\toolsplus04\databases\mastest\mas200\version4\MAS90;Prefix=\\toolsplus04\databases\mastest\mas200\version4\MAS90\SY,\\toolsplus04\databases\mastest\mas200\version4\MAS90\==;ViewDLL=\\toolsplus04\databases\mastest\mas200\version4\MAS90\Home;LogFile=\PVXODBC.LOG;CacheSize=4;DirtyReads=1;BurstMode=1;StripTrailingSpaces=1;SERVER=NotTheServer;UID=bd|TOO;PWD=brian"
Private Const MAS90_CONN As String = "Driver={MAS 90 4.0 ODBC Driver};Directory=s:\mas45server\MAS90;Prefix=s:\mas45server\MAS90\SY,s:\mas45server\MAS90\==;ViewDLL=s:\mas45server\MAS90\Home;LogFile=\PVXODBC.LOG;CacheSize=4;DirtyReads=1;BurstMode=1;StripTrailingSpaces=1;SERVER=NotTheServer;UID=bd|TOO;PWD=brian"
Private Const BARCODE_CONN As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=s:\mastest\BC Wedge\BCWedge.MDB"
Private Const SHIPPING_CONN As String = "Driver={MySQL ODBC 3.51 Driver};Server=cindy.asmallorange.com;Uid=esavelle_rw;Pwd=7539042;Database=esavelle_toolsplus"

Private WithEvents T As TimerClass
Attribute T.VB_VarHelpID = -1
Private conn As ADODB.Connection

Private cdb As DBEngines
Private timeoutLen As Long
Private catalogName As String

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Class_Initialize()
    cdb = DBEngines.DBENG_MSSQL
    catalogName = "toolsplus"
    Set T = New TimerClass
    T.Interval = 30000
    T.Enabled = False
End Sub

Private Sub Class_Terminate()
    Set T = Nothing
End Sub

Public Property Get CurrentDB() As DBEngines
    CurrentDB = cdb
End Property

Public Property Let CurrentDB(newcdb As DBEngines)
    If cdb <> newcdb Then
        cdb = newcdb
        killDBH
    End If
End Property

Public Property Get Catalog() As String
    If cdb = DBENG_MSSQL Then
        Catalog = catalogName
    Else
        Err.Raise 123, "DatabaseSingleton", "Catalog() only supported for MSSQL"
    End If
End Property

Public Property Let Catalog(newCatalogName As String)
    If cdb = DBENG_MSSQL Then
        If catalogName <> newCatalogName Then
            catalogName = newCatalogName
            killDBH
        End If
    Else
        Err.Raise 123, "DatabaseSingleton", "Catalog() only supported for MSSQL"
    End If
End Property

Public Property Get TimeoutPeriod() As Long
    TimeoutPeriod = timeoutLen
End Property

Public Property Let TimeoutPeriod(newTimeout As Long)
    If newTimeout <= 0 Then
        timeoutLen = 0
        T.Enabled = False
    Else
        timeoutLen = newTimeout
        T.Interval = timeoutLen
    End If
End Property

Public Property Get Handle() As ADODB.Connection
    Set Handle = getDBH()
End Property

Public Property Get ConnState() As ADODB.ObjectStateEnum
    ConnState = conn.State
End Property

Public Function execute(strsql As String, Optional suppressKeyViolations As Boolean = False, Optional handleTimeouts As Boolean = True, Optional noRaiseErrors As Boolean = False) As Boolean
On Error GoTo errh
start:
    getDBH().execute strsql
    execute = True
    Exit Function

errh:
    'these are MSSQL error codes, but whatever.
    Select Case Err.Number
        Case Is = -2147217871, -2147467259
            If handleTimeouts Then
                Err.Clear
                GoTo start
            Else
                Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
            End If
        Case Else
            If suppressKeyViolations And (Err.Number = -2147217873 Or Err.Number = -2147217873) Then 'pk or pdx violation
                'nothing
            Else
                'MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "SQL string follows: " & vbCrLf & strsql
                'If InStr(strsql, "DisableTriggers") Or InStr(strsql, "EnableTriggers") Then
                '    MsgBox "Error disabling or enabling triggers! THIS IS VERY BAD!"
                'End If
                If noRaiseErrors Then
                    'this is a workaround...for some reason the previous error handler isn't
                    'being called, so this raise ends up crashing. instead of raising an error
                    'we're just going to check the return value, and handle it that way.
                Else
                    Err.Raise Err.Number, "DBConn", Err.Description & vbCrLf & vbCrLf & "SQL: " & strsql
                End If
            End If
            Err.Clear
    End Select
    execute = False
End Function

Public Function retrieve(strsql As String, Optional handleTimeouts As Boolean = True) As ADODB.Recordset
On Error GoTo errh
start:
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open strsql, getDBH(), adOpenStatic, adLockBatchOptimistic
    rst.ActiveConnection = Nothing
    Set retrieve = rst
    Exit Function

errh:
    'these are MSSQL error codes, but whatever.
    Select Case Err.Number
        Case Is = -2147217871, -2147467259
            'If MsgBox("Network error...either the network is broken, or the server is busy." & vbCrLf & vbCrLf & "Try again?", vbYesNo) = vbYes Then
            '    Err.Clear
            '    GoTo start
            'End If
            If handleTimeouts Then
                Err.Clear
                GoTo start
            Else
                'is this how to propagate errors?
                Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
            End If
        Case Else
            'MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "SQL string follows: " & vbCrLf & strsql
            Err.Raise Err.Number, "DBConn", Err.Description & vbCrLf & vbCrLf & "SQL: " & strsql
    End Select
End Function

Private Function getDBH() As ADODB.Connection
'MsgBox "Called getDBH()"
    If conn Is Nothing Then
'MsgBox "conn is not initialized"
        Set conn = New ADODB.Connection
'MsgBox "initialized conn"
        Select Case cdb
            Case Is = DBEngines.DBENG_MSSQL
'MsgBox "mssql connection process started"
                'Select Case mssqlconn_picked
                '    Case Is = 0
                '        On Error Resume Next
                '        conn.Open MSSQL_CONN
                '        On Error GoTo 0
'MsgBox "attempted sqlncli connection, current error code is " & Err.Number
                '        mssqlconn_picked = 1
                '        If Err.Number <> 0 Then 'looking for 3706, but catch anything
'MsgBox "sqlncli failed?"
                '            Err.Clear
                            conn.Open Replace(MSSQL_ALT_CONN, "**CATALOG**", catalogName)
'MsgBox "attempted oledb connection, current error code is " & Err.Number
                '            mssqlconn_picked = 2
                '        End If
                '    Case Is = 1
                '        conn.Open MSSQL_CONN
                '    Case Is = 2
                '        conn.Open MSSQL_ALT_CONN
                'End Select
            Case Is = DBEngines.DBENG_MAS90
                conn.Open MAS90_CONN
            Case Is = DBEngines.DBENG_BARCODE
                conn.Open BARCODE_CONN
            Case Is = DBEngines.DBENG_SHIPPING
                conn.Open SHIPPING_CONN
        End Select
        'Debug.Print "opening connection"
    End If
    If timeoutLen <> 0 Then
        T.Enabled = True 'setting enabled when already enabled re-initializes
    End If
'MsgBox "returning conn, state is " & conn.State
    Set getDBH = conn
End Function

Private Sub killDBH()
    If Not conn Is Nothing Then
        'if other state (fetching/executing?), then ignore for another interval
        If conn.State = adStateOpen Then
            conn.Close
            Set conn = Nothing
            'Debug.Print "closing connection"
        End If
    End If
End Sub

Private Sub T_Timer()
    killDBH
End Sub



' requires VBJSON, Microsoft Scripting Runtime, Microsoft XML 6.0
Public Function MASRetrieveViaJSON(strsql As String) As ADODB.Recordset
    Dim json As String
    json = jsonQuery(strsql)
    If json = "" Then
        Err.Raise 1, "MASRetrieveViaJSON", "Error during JSON query"
    End If
    
    Dim parser As VBJSON.JSONParser
    Set parser = New VBJSON.JSONParser
    Dim retval As Dictionary
    Set retval = parser.Decode(json)
    If retval Is Nothing Then
        Err.Raise 2, "MASRetrieveViaJSON", "Error parsing JSON"
    End If
    If Not retval.Exists("columnInfo") Then
        Err.Raise 3, "MASRetrieveViaJSON", "Missing expected parameter 'columnInfo' in JSON response"
    End If
    If Not retval.Exists("rows") Then
        Err.Raise 4, "MASRetrieveViaJSON", "Missing expected parameter 'rows' in JSON response"
    End If
    
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    Dim iter As Variant, iter2 As Variant
    
    For Each iter In retval.Item("columnInfo")
        rst.Fields.Append iter.Item("columnName"), columnTypeIDToADOType(iter.Item("columnTypeID")), iter.Item("columnSize"), adFldIsNullable
    Next iter
    
    rst.Open
    
    For Each iter In retval.Item("rows")
        rst.AddNew
        For Each iter2 In iter.Keys
            If IsNull(iter.Item(CStr(iter2))) Then
                'nothing to do, default is null
            Else
                rst.Fields(CStr(iter2)).Value = iter.Item(CStr(iter2))
            End If
        Next iter2
        rst.Update
    Next iter
    
    If rst.RecordCount > 0 Then
        rst.MoveFirst
    End If
    
    Set MASRetrieveViaJSON = rst
End Function

Private Function jsonQuery(strsql As String) As String
    Dim req As MSXML2.XMLHTTP
    Set req = New MSXML2.XMLHTTP
    req.Open "GET", "http://toolsplus04/whse/arbitrary_sql.plex?" & strsql, False
    req.send ""
    While req.readyState <> 4
        'Debug.Print req.ReadyState
        Sleep 50
        DoEvents
    Wend
    If req.Status = 200 Then
        jsonQuery = req.responseText
    Else
        jsonQuery = ""
    End If
End Function

Private Function columnTypeIDToADOType(columnTypeID As Long) As ADODB.DataTypeEnum
    Select Case columnTypeID
        Case Is = -7
            columnTypeIDToADOType = adBinary
        Case Is = -1, 1, 12
            columnTypeIDToADOType = adVarChar
        Case Is = 3
            columnTypeIDToADOType = adDecimal
        Case Is = 9, 91
            columnTypeIDToADOType = adDate
        Case Else
            columnTypeIDToADOType = adVarChar
    End Select
End Function
