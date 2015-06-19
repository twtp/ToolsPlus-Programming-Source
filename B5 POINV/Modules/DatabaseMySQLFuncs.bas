Attribute VB_Name = "DatabaseMySQLFuncs"
Option Explicit

'Public Function openMySQLConn() As ADODB.Connection
'    Dim conn As ADODB.Connection
'    Set conn = New ADODB.Connection
'    conn.Open EXTERN_DB_DSN
'    Set openMySQLConn = conn
'End Function
'
'Public Sub closeMySQLConn(conn As ADODB.Connection)
'    conn.Close
'    Set conn = Nothing
'End Sub
'
'Public Function mysqlRetrieve(strsql As String, Optional conn As ADODB.Connection = Nothing) As ADODB.Recordset
' On Error GoTo errh
'    Mouse.Hourglass True
'start:
'    Dim closeconn As Boolean
'    If conn Is Nothing Then
'        closeconn = True
'        Set conn = New ADODB.Connection
'        conn.Open EXTERN_DB_DSN
'    Else
'        closeconn = False
'    End If
'    Dim rst As Recordset
'    Set rst = New ADODB.Recordset
'    rst.CursorLocation = adUseClient
'    rst.Open strsql, conn, adOpenStatic, adLockBatchOptimistic
'    rst.ActiveConnection = Nothing
'    If closeconn Then
'        conn.Close
'        Set conn = Nothing
'    End If
'    Set mysqlRetrieve = rst
'    Set rst = Nothing
'    Mouse.Hourglass False
'    Exit Function
'
'errh:
'    Select Case Err.Number
'        Case Is = -2147467259   'connection timeout (20sec), try again
'            Err.Clear
'            Resume
'        Case Else
'            MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.description & vbCrLf & vbCrLf & "SQL string follows: " & vbCrLf & strsql
'    End Select
'    Mouse.Hourglass False
'End Function
'
'Public Function mysqlExecute(strsql As String, Optional conn As ADODB.Connection = Nothing) As Boolean
'    Mouse.Hourglass True
'On Error GoTo errh
'    Dim closeconn As Boolean
'    If conn Is Nothing Then
'        closeconn = True
'        Set conn = New ADODB.Connection
'        conn.Open EXTERN_DB_DSN
'    End If
'    conn.Execute strsql
'    If closeconn Then
'        conn.Close
'        Set conn = Nothing
'    End If
'    mysqlExecute = True
'    Mouse.Hourglass False
'    Exit Function
'errh:
'    Select Case Err.Number
'        Case Is = -2147467259
'            Err.Clear
'            Resume
'        Case Else
'            MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.description & vbCrLf & vbCrLf & "SQL string follows: " & vbCrLf & strsql
'    End Select
'    Mouse.Hourglass False
'End Function
