VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MessageServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Listener As CSocketMaster
Attribute Listener.VB_VarHelpID = -1
Private WithEvents ConnPool As ConnectionPool
Attribute ConnPool.VB_VarHelpID = -1

Private partialMsg As String
Private msgLen As Long

Public Event ReceivedMessage(message As String, fromIP As String, fromConnID As String)
Public Event RelayedMessage(message As String, fromIP As String, fromConnID As String)
Public Event RPCQueryHandled(typeOfRPC As String, fromIP As String, fromConnID As String)
Public Event RPCQueryInvalid(typeOfRPC As String, fromIP As String, fromConnID As String)
Public Event PeerConnected(IP As String, connID As String)
Public Event PeerDisconnected(IP As String, connID As String)

'Public Function RelayMessage(message As String)
'    Dim i As Long
'    If toIP_or_ConnID = "" Then
'        'for each client, send message
'        For i = 1 To ConnPool.Count
'            ConnPool.Client(CStr(i)).SendData Len(message) & "|" & message
'        Next i
'    ElseIf InStr(toIP_or_ConnID, ".") Then
'        'to single ip, possibly multiple instances, linear search for now?
'        For i = 1 To ConnPool.Count
'            If ConnPool.Client(CStr(i)).IPAddress = toIP_or_ConnID Then
'                ConnPool.Client(CStr(i)).SendData Len(message) & "|" & message
'            End If
'        Next i
'    Else
'        'to a specific connection id
'        ConnPool.Client(toIP_or_ConnID).SendData Len(message) & "|" & message
'    End If
'End Function

Public Sub SendMessage(clientID As String, message As String, whichApp As String, destIP As String, origIP As String, origUser As String)
    Dim thismsg As String
    thismsg = whichApp & "|" & _
              destIP & "|" & _
              origIP & "|" & _
              origUser & "|" & _
              message
    thismsg = Len(thismsg) & "|" & thismsg
    ConnPool.Client(clientID).SendData thismsg
End Sub

Public Function EnumerateConnectionIDs(Optional IP As String = "") As Dictionary
    Dim retval As Dictionary
    Set retval = New Dictionary
    Dim i As Long
    For i = 1 To ConnPool.Count
        If IP = "" Or IP = ConnPool.Client(CStr(i)).IPAddress Then
            retval.Add CStr(ConnPool.Client(CStr(i)).Key), CStr(ConnPool.Client(CStr(i)).IPAddress)
        End If
    Next i
    Set EnumerateConnectionIDs = retval
    Set retval = Nothing
End Function

'Public Function EnumerateConnectionIPs() As String()
'    Dim retval() As String
'    Dim init As Boolean
'    init = False
'    Dim i As Long
'    For i = 1 To ConnPool.Count
'        If init Then
'            ReDim Preserve retval(UBound(retval) + 1) As String
'        Else
'            ReDim retval(0) As String
'        End If
'        retval(i) = ConnPool.Client(CStr(i)).IPAddress
'    Next i
'    Set EnumerateConnectionIPs = retval()
'End Function

Private Sub Class_Initialize()
    Set Listener = New CSocketMaster
    Listener.LocalPort = SERVER_PORT
On Error GoTo errh
    Listener.Listen
    'Debug.Print "Server started on " & Listener.LocalHostName & ":" & Listener.LocalPort
    
    Set ConnPool = New ConnectionPool
    'Debug.Print "Connection pool active, waiting for connections"
    Exit Sub
errh:
    If Err.Number = 10048 Then 'addr in use
        Listener.LocalPort = Listener.LocalPort + 1
        Err.Clear
        Resume
    Else
        MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    End If
End Sub

Private Sub Class_Terminate()
    Set ConnPool = Nothing
    'Debug.Print "Connection pool destroyed"
    Set Listener = Nothing
    'Debug.Print "Server stopped"
End Sub

Private Sub ConnPool_DataArrival(clientID As String, IP As String, data As String)
    If partialMsg = "" Then
        'Debug.Print "Processing new incoming message from " & ip
        Dim sep As Long, msg As String, newMsgLen As Long
        sep = InStr(data, "|")
        If sep <> 0 Then
            msg = Left(data, sep - 1)
            If IsNumeric(msg) Then
                newMsgLen = CLng(msg)
                msg = Mid(data, sep + 1)
                If Len(msg) = newMsgLen Then
                    'Debug.Print "Message received from " & ip & ", notifying app"
                    RaiseEvent ReceivedMessage(msg, IP, clientID)
                Else
                    'Debug.Print "Partial message received, waiting for next chunk"
                    msgLen = newMsgLen
                    partialMsg = msg
                End If
            Else
                Debug.Print "Invalid message received: bad length"
            End If
        Else
            Debug.Print "Invalid message received: no length found"
        End If
    Else
        'Debug.Print "Processing continued message block from " & ip
        partialMsg = partialMsg & data
        If Len(partialMsg) = msgLen Then
            'Debug.Print "Message received from " & ip & ", notifying app"
            RaiseEvent ReceivedMessage(partialMsg, IP, clientID)
        Else
            'Debug.Print "Partial message received, waiting for next chunk"
        End If
    End If
End Sub

Private Sub ConnPool_Disconnect(clientID As String, IP As String)
    PeerDisconnectNotification clientID
    ConnPool.Disconnect clientID
    'Debug.Print "Connection #" & clientID & " disconnected, removed from pool"
    RaiseEvent PeerDisconnected(IP, clientID)
End Sub

Private Sub ConnPool_PeerConnect(clientID As String, IP As String)
    'this gets handled by RPC
    'RaiseEvent PeerConnected(IP, clientID)
End Sub

Private Sub Listener_ConnectionRequest(ByVal requestID As Long)
    ConnPool.Add "#" & requestID
    'Debug.Print "Connection request #" & requestID & " added to connection pool"
    If ConnPool.Client("#" & requestID).AcceptConn(requestID) Then
        'Debug.Print "Connection request #" & requestID & " accepted"
    Else
        'Debug.Print "Connection request #" & requestID & " rejected!"
    End If
End Sub

Private Sub parseAndHandle(ByVal msg As String, connIP As String, clientID As String)
    Dim orig As String
    orig = msg
On Error GoTo errh
    Dim thisAppName As String, destIP As String, origIP As String, origUser As String
    Dim pos As Long
    
    pos = InStr(msg, "|")
    thisAppName = Left(msg, pos - 1)
    msg = Mid(msg, pos + 1)
    
    pos = InStr(msg, "|")
    destIP = Left(msg, pos - 1)
    msg = Mid(msg, pos + 1)
    
    pos = InStr(msg, "|")
    origIP = Left(msg, pos - 1) 'should be same as connIP
    msg = Mid(msg, pos + 1)
    
    pos = InStr(msg, "|")
    origUser = Left(msg, pos - 1)
    msg = Mid(msg, pos + 1)
    
    If thisAppName = "RPC" Then
        'remote procedure call, this means that the MessageServer is supposed to handle something
        Dim args As String
        pos = InStr(msg, "|")
        If pos = 0 Then
            'no args, do nothing
        Else
            'has args, split
            msg = Left(msg, pos - 1)
            args = Mid(msg, pos + 1)
        End If
        Select Case msg
            Case Is = "EnumeratePeers"
                EnumeratePeers "#" & clientID
                RaiseEvent RPCQueryHandled("EnumeratePeers", connIP, clientID)
            Case Is = "SetUserName"
                ConnPool.Client("#" & clientID).UserName = args
                RaiseEvent RPCQueryHandled("SetUserName", connIP, clientID)
            Case Is = "SetAppName"
                ConnPool.Client("#" & clientID).appName = args
                RaiseEvent RPCQueryHandled("SetAppName", connIP, clientID)
            Case Is = "LoginOK"
                NewPeerNotification "#" & clientID
                RaiseEvent RPCQueryHandled("LoginOK", connIP, clientID)
                RaiseEvent PeerConnected(connIP, clientID)
            Case Else
                'invalid RPC, do nothing
                RaiseEvent RPCQueryInvalid(msg, connIP, clientID)
        End Select
    Else
        'TODO
        'if broadcast msg, then relay to all
        'if to other, then relay
        'else ?
        If destIP = SERVER_IP Then
            RaiseEvent ReceivedMessage(msg, connIP, clientID)
        Else
            Dim i As Long
            For i = 1 To ConnPool.Count
                If destIP = "broadcast" Or ConnPool.Client(CStr(i)).IPAddress = destIP Then
                    ConnPool.Client(CStr(i)).SendData orig
                End If
            Next i
            RaiseEvent RelayedMessage(msg, connIP, clientID)
        End If
    End If
    Exit Sub
errh:
    Err.Clear
    Exit Sub
End Sub

Private Sub EnumeratePeers(clientID As String)
    Dim i As Long
    Dim tempH As Dictionary
    Dim un As String, IP As String
    
    Dim forWhichApp As String
    forWhichApp = ConnPool.Client(clientID).appName
    
    For i = 1 To ConnPool.Count
        If ConnPool.Client(CStr(i)).Key <> clientID Then
            If ConnPool.Client(CStr(i)).appName = forWhichApp Then
                un = ConnPool.Client(CStr(i)).UserName
                IP = ConnPool.Client(CStr(i)).IPAddress
                If tempH.Exists(CStr(un)) Then
                    tempH.Item(CStr(un)) = CStr(tempH.Item(CStr(un)) & "," & IP)
                Else
                    tempH.Add CStr(un), CStr(IP)
                End If
            End If
        End If
    Next i
    
    Dim thiskey As Variant, retval As String
    For Each thiskey In tempH.Keys
        retval = retval & "|" & thiskey & "," & tempH.Item(CStr(thiskey))
    Next thiskey
    If retval <> "" Then
        retval = Mid(retval, 2)
    End If
    
    'construct envelope and send
    SendMessage clientID, retval, "RPC-EnumeratePeers", ConnPool.Client(clientID).IPAddress, SERVER_IP, "admin"
End Sub

Private Sub NewPeerNotification(clientID As String)
    Dim un As String, IP As String, thisapp As String
    un = ConnPool.Client(clientID).UserName
    IP = ConnPool.Client(clientID).IPAddress
    thisapp = ConnPool.Client(clientID).appName
    
    Dim i As Long
    For i = 1 To ConnPool.Count
        If ConnPool.Client(CStr(i)).Key <> clientID Then
            If ConnPool.Client(CStr(i)).appName = thisapp Then
                SendMessage ConnPool.Client(CStr(i)).Key, _
                            un & "|" & IP, _
                            "RPC-NewPeer", _
                            ConnPool.Client(CStr(i)).IPAddress, _
                            SERVER_IP, _
                            "admin"
            End If
        End If
    Next i
End Sub

Private Sub PeerDisconnectNotification(clientID As String)
    Dim un As String, IP As String, thisapp As String
    un = ConnPool.Client(clientID).UserName
    IP = ConnPool.Client(clientID).IPAddress
    thisapp = ConnPool.Client(clientID).appName
    
    Dim i As Long
    For i = 1 To ConnPool.Count
        If ConnPool.Client(CStr(i)).Key <> clientID Then
            If ConnPool.Client(CStr(i)).appName = thisapp Then
                SendMessage ConnPool.Client(CStr(i)).Key, _
                            un & "|" & IP, _
                            "RPC-GonePeer", _
                            ConnPool.Client(CStr(i)).IPAddress, _
                            SERVER_IP, _
                            "admin"
            End If
        End If
    Next i
End Sub
