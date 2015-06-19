VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MessageClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ServerConn As CSocketMaster
Attribute ServerConn.VB_VarHelpID = -1

Private partialMsg As String
Private msgLen As Long

Private enumPeersDone As Boolean
Private enumPeersMsg As String

Public Event ReceivedMessage(message As String, fromIP As String, fromUser As String, appName As String)
Public Event PeerConnect(peerUserName As String, peerIP As String)
Public Event PeerDisconnect(peerUserName As String, peerIP As String)

Public Function SendMessage(message As String, appName As String, Optional toIP As String = "")
    'construct the envelope, same as parseAndNotify()
    Dim fullmsg As String
    fullmsg = appName & "|" & _
              IIf(toIP = "", "broadcast", toIP) & "|" & _
              ServerConn.LocalIP & "|" & _
              Environ("UserName") & "|" & _
              message
    fullmsg = Len(fullmsg) & "|" & fullmsg
    
    ServerConn.SendData fullmsg
End Function

Public Sub SetUserName(newUserName As String)
    Dim thismsg As String
    thismsg = "RPC" & "|" & _
              SERVER_IP & "|" & _
              ServerConn.LocalIP & "|" & _
              Environ("UserName") & "|" & _
              "SetUserName-" & newUserName
    thismsg = Len(thismsg) & "|" & thismsg
    ServerConn.SendData thismsg
End Sub

Public Sub SetAppName(newAppName As String)
    Dim thismsg As String
    thismsg = "RPC" & "|" & _
              SERVER_IP & "|" & _
              ServerConn.LocalIP & "|" & _
              Environ("UserName") & "|" & _
              "SetAppName-" & newAppName
    thismsg = Len(thismsg) & "|" & thismsg
    ServerConn.SendData thismsg
End Sub

Public Sub CompleteLogin()
    Dim thismsg As String
    thismsg = "RPC" & "|" & _
              SERVER_IP & "|" & _
              ServerConn.LocalIP & "|" & _
              Environ("UserName") & "|" & _
              "LoginOK"
    thismsg = Len(thismsg) & "|" & thismsg
    ServerConn.SendData thismsg
End Sub

Public Function EnumeratePeers(appName As String) As Dictionary
    enumPeersDone = False
    enumPeersMsg = ""
    Dim thismsg As String
    thismsg = "RPC" & "|" & _
              SERVER_IP & "|" & _
              ServerConn.LocalIP & "|" & _
              Environ("UserName") & "|" & _
              "EnumeratePeers"
    thismsg = Len(thismsg) & "|" & thismsg
    ServerConn.SendData thismsg
    
    While Not enumPeersDone
        DoEvents
    Wend
    'TODO
    'parse message, pipe + comma delimited
    'return dict
    
End Function

Private Sub Class_Initialize()
    Set ServerConn = New CSocketMaster
    ServerConn.Connect SERVER_IP, SERVER_PORT
End Sub

Private Sub Class_Terminate()
    Set ServerConn = Nothing
End Sub

Private Sub ServerConn_DataArrival(ByVal bytesTotal As Long)
    Dim data As String
    ServerConn.GetData data
    If partialMsg = "" Then
        'Debug.Print "Processing new incoming message"
        Dim sep As Long, msg As String, newMsgLen As Long
        sep = InStr(data, "|")
        If sep <> 0 Then
            msg = Left(data, sep - 1)
            If IsNumeric(msg) Then
                newMsgLen = CLng(msg)
                msg = Mid(data, sep + 1)
                If Len(msg) = newMsgLen Then
                    'Debug.Print "Message received, parsing now"
                    parseAndNotify msg
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
        'Debug.Print "Processing continued message block"
        partialMsg = partialMsg & data
        If Len(partialMsg) = msgLen Then
            'Debug.Print "Message received, parsing now"
            parseAndNotify msg
        Else
            'Debug.Print "Partial message received, waiting for next chunk"
        End If
    End If
End Sub

Private Sub parseAndNotify(ByVal msg As String)
    'assumes the length has already been removed from the front and validated,
    'now we just need to split the envelope apart. message handling itself should
    'be handled in another higher module
On Error GoTo errh
    Dim appName As String, destIP As String, origIP As String, origUser As String
    Dim pos As Long
    
    pos = InStr(msg, "|")
    appName = Left(msg, pos - 1)
    msg = Mid(msg, pos + 1)
    
    pos = InStr(msg, "|")
    destIP = Left(msg, pos - 1)
    msg = Mid(msg, pos + 1)
    
    pos = InStr(msg, "|")
    origIP = Left(msg, pos - 1)
    msg = Mid(msg, pos + 1)
    
    pos = InStr(msg, "|")
    origUser = Left(msg, pos - 1)
    msg = Mid(msg, pos + 1)
    
    If Left(appName, 4) = "RPC-" Then
        'ok, received return from remote procedure call, handle it
        Select Case Mid(appName, 5)
            Case Is = "EnumeratePeers"
                enumPeersMsg = msg
                enumPeersDone = True
            Case Is = "NewPeer"
                RaiseEvent PeerConnect(Left(msg, InStr(msg, "|") - 1), Mid(msg, InStr(msg, "|") + 1))
            Case Is = "GonePeer"
                RaiseEvent PeerDisconnect(Left(msg, InStr(msg, "|") - 1), Mid(msg, InStr(msg, "|") + 1))
            Case Else
                'probably should have handled this
        End Select
    Else
        RaiseEvent ReceivedMessage(msg, origIP, origUser, appName)
    End If
    Exit Sub
errh:
    Err.Clear
    Exit Sub
End Sub
