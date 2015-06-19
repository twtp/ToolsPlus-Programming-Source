VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConnectionPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private thisPool As Collection

Public Enum EVENT_INDEXES
    clientevent_Connect
    clientevent_Disconnect
    clientevent_DataArrival
End Enum

'these get fired through the pointer, really belong to individual connections
Public Event PeerConnect(clientID As String, IP As String)
Public Event Disconnect(clientID As String, IP As String)
Public Event DataArrival(clientID As String, IP As String, data As String)


Public Function Add(clientID As String) As ClientConnection
    Dim newMember As ClientConnection
    Set newMember = New ClientConnection
    newMember.Key = clientID
    newMember.ParentPointer = ObjPtr(Me)
    thisPool.Add newMember, clientID
    Set Add = newMember
    Set newMember = Nothing
End Function

Public Property Get Client(clientID As String) As ClientConnection
On Error GoTo errh
    If IsNumeric(clientID) Then
        Set Client = thisPool.item(CLng(clientID))
    Else
        Set Client = thisPool.item(clientID)
    End If
    Exit Property
errh:
    If Err.Number = 5 Then 'not exists, forget it
        Err.Clear
        Set Client = Nothing
    Else
        MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    End If
End Property

Public Property Get Count() As Long
    Count = thisPool.Count
End Property

Public Sub Disconnect(clientID As String)
On Error GoTo errh
    thisPool.Remove clientID
    Exit Sub
errh:
    If Err.Number = 5 Then 'not exists, forget it
        Err.Clear
    Else
        MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    End If
End Sub

Private Sub Class_Initialize()
    Set thisPool = New Collection
    PointerFunctions.AddPointer ObjPtr(Me)
End Sub

Private Sub Class_Terminate()
    Set thisPool = Nothing
End Sub

'Public Property Get NewEnum() As IUnknown
'    Set NewEnum = thisPool.[_NewEnum]
'End Property

Friend Sub FireEvent(ByVal eventid As EVENT_INDEXES, ByVal clientID As String, ByVal thePointer As Long, anyArgs As Variant)
    Dim i As Long, dx As Long
    For i = 1 To thisPool.Count
        If ObjPtr(thisPool.item(i)) = thePointer Then
            dx = i
            Exit For
        End If
    Next i
    
    If dx > 0 Then
        Select Case eventid
            Case clientevent_Connect
                'anyargs(0) = ip
                RaiseEvent PeerConnect(clientID, CStr(anyArgs(0)))
            Case clientevent_DataArrival
                'anyargs(0) = ip
                'anyargs(1) = the data arriving
                RaiseEvent DataArrival(clientID, CStr(anyArgs(0)), CStr(anyArgs(1)))
            Case clientevent_Disconnect
                'anyargs(0) = ip
                RaiseEvent Disconnect(clientID, CStr(anyArgs(0)))
        End Select
    End If
End Sub
