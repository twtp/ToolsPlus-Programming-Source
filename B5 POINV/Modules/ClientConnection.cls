VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClientConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents sock As CSocketMaster
Attribute sock.VB_VarHelpID = -1
'Public WithEvents outsock As CSocketMaster

Private id As String
Private parent As Long
Private ipaddr As String
Private fqdn As String
Private usern As String
Private appn As String

Private SendComplete As Boolean
Private sendSuccess As Boolean

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)

'Public Event PeerConnect(clientID As String, IP As String)
'Public Event Disconnect(clientID As String, IP As String)
'Public Event DataArrival(clientID As String, IP As String, data As String)

Public Property Get Key() As String
    Key = id
End Property

Public Property Let Key(newKey As String)
    'read/write, but this probably shouldn't be
    id = newKey
End Property

Public Property Let ParentPointer(newPointer As Long)
    'write-once, read-none
    If parent = 0 Then
        parent = newPointer
    End If
End Property

Public Property Get IPAddress() As String
    'read only
    IPAddress = ipaddr
End Property

Public Property Get Hostname() As String
    'read only
    Hostname = fqdn
End Property

Public Property Get UserName() As String
    UserName = usern
End Property

Public Property Let UserName(newUserName As String)
    usern = newUserName
End Property

Public Property Get appName() As String
    appName = appn
End Property

Public Property Let appName(newAppName As String)
    appn = newAppName
End Property

Public Function AcceptConn(requestID As Long) As Boolean
On Error GoTo errh
    Me.sock.Accept requestID
    ipaddr = Me.sock.RemoteHostIP
    fqdn = Me.sock.RemoteHost
    AcceptConn = True
    'RaiseEvent PeerConnect(id, ipaddr)
    callEvent clientevent_Connect, Array(ipaddr)
    Exit Function
errh:
    AcceptConn = False
    Err.Clear
End Function

'Public Function ConnectTo(IP As String, port As String) As Boolean
'    Me.outsock.Connect IP, port
'    ConnectTo = True
'End Function

Public Function SendData(msg As String) As Boolean
On Error GoTo errh
    SendComplete = False
    sendSuccess = False
    Me.sock.SendData msg
    While Not SendComplete
        DoEvents
    Wend
    SendData = sendSuccess
    Exit Function
errh:
    SendData = False
    Err.Clear
End Function

Private Sub Class_Initialize()
    Set sock = New CSocketMaster
End Sub

Private Sub sock_CloseSck()
    sock.CloseSck
    'RaiseEvent Disconnect(id, sock.RemoteHostIP)
    callEvent clientevent_Disconnect, Array(ipaddr)
End Sub

Private Sub sock_DataArrival(ByVal bytesTotal As Long)
    Dim sockData As String
    sock.GetData sockData
    'RaiseEvent DataArrival(id, sock.RemoteHostIP, sockData)
    callEvent clientevent_DataArrival, Array(ipaddr, sockData)
End Sub

Private Sub callEvent(eventid As EVENT_INDEXES, anyArgs As Variant)
    If IsValidPointer(parent) > -1 Then
        resolvePtr(parent).FireEvent eventid, id, ObjPtr(Me), anyArgs
    End If
End Sub

Private Function resolvePtr(thePointer As Long) As ConnectionPool
    Dim tmp As ConnectionPool
    CopyMemory tmp, thePointer, PTR_SIZE
    Set resolvePtr = tmp
    CopyMemory tmp, 0&, PTR_SIZE
End Function

Private Sub sock_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sendSuccess = False
    SendComplete = True
    sock_CloseSck
End Sub

Private Sub sock_sendComplete()
    sendSuccess = True
    SendComplete = True
End Sub

Private Sub sock_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    SendComplete = False
End Sub
