VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MettlerToledoScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type ScaleWeight
    weight As String
    units As String
    steady As Boolean
End Type

Public Event WeightIncoming(incoming As ScaleWeight)

Private WithEvents pollTimer As TimerClass
Attribute pollTimer.VB_VarHelpID = -1
Private comm As MSCommLib.MSComm

Private currentCommPort As String
Private currentBaudRate As String
Private currentParity As String
Private currentDataBit As String
Private currentStopBit As String

Public Property Get CommPort() As String
    CommPort = currentCommPort
End Property
Public Property Let CommPort(newCommPort As String)
    newCommPort = Replace(newCommPort, "COM", "")
    If IsNumeric(newCommPort) Then
        currentCommPort = newCommPort
    Else
        Err.Raise 1, , "Invalid COM specified"
    End If
End Property

Public Property Get BaudRate() As String
    BaudRate = currentBaudRate
End Property
Public Property Let BaudRate(newBaudRate As String)
    If IsNumeric(newBaudRate) Then
        currentBaudRate = newBaudRate
    Else
        Err.Raise 1, , "Invalid baud rate specified"
    End If
End Property

Public Property Get Parity() As String
    Parity = currentParity
End Property
Public Property Let Parity(newParity As String)
    newParity = UCase(Left(newParity, 1))
    If newParity = "O" Or newParity = "E" Or newParity = "N" Then
        currentParity = newParity
    Else
        Err.Raise 1, , "Invalid parity specified"
    End If
End Property

Public Property Get DataBit() As String
    DataBit = currentDataBit
End Property
Public Property Let DataBit(newDataBits As String)
    If IsNumeric(newDataBits) Then
        currentDataBit = newDataBits
    Else
        Err.Raise 1, , "Invalid data bits specified"
    End If
End Property

Public Property Get StopBit() As String
    StopBit = currentStopBit
End Property
Public Property Let StopBit(newStopBits As String)
    If IsNumeric(newStopBits) Then
        currentStopBit = newStopBits
    Else
        Err.Raise 1, , "Invalid stop bits specified"
    End If
End Property

Public Function OpenConn() As Boolean
On Error GoTo errh
    If comm Is Nothing Then
        Set comm = New MSCommLib.MSComm
        comm.CommPort = currentCommPort
        comm.Settings = getSettingsAsString
    End If
    If comm.PortOpen = False Then
        comm.PortOpen = True
    End If
    OpenConn = True
    Exit Function
errh:
    OpenConn = False
End Function

Public Function CloseConn() As Boolean
On Error GoTo errh
    If Not comm Is Nothing Then
        If comm.PortOpen = True Then
            comm.PortOpen = False
        End If
    End If
    CloseConn = True
    Exit Function
errh:
    CloseConn = False
End Function

Public Function GetCurrentValue() As ScaleWeight
    Me.OpenConn
    comm.Output = "H"
    Sleep 150
    Dim inputstr As String, retval As ScaleWeight
    inputstr = ""
    If comm.InBufferCount Then
        inputstr = comm.Input
    End If
    If inputstr = "" Then
        retval.weight = 0
        retval.units = "NA"
        retval.steady = False
    Else
        inputstr = Left(inputstr, 8)
        ' byte format
        '   good:
        '     1 - STX
        '     2 through 8 - weight
        '   bad:
        '     1 - STX
        '     2 - "?" literal
        '     3 - status byte
        '           0 - in motion
        '           1 - over capacity
        '           2 - under zero
        '           3 - outside zero capture range
        '           4 - center of zero
        '           5 - always on
        '           6 - always on
        '     4 through 8 - weight
        If Mid(inputstr, 2, 1) = "?" Then
            retval.weight = "--.--"
            retval.units = "N/A"
            retval.steady = False 'should this be expanded for more status flags?
        Else
            retval.steady = True
            retval.units = "N/A"
            retval.weight = Mid(inputstr, 2, 6)
        End If
        While Left(retval.weight, 1) = "0"
            retval.weight = Mid(retval.weight, 2)
        Wend
        If Left(retval.weight, 1) = "." Then
            retval.weight = "0" & retval.weight
        End If
        retval.weight = Trim(retval.weight)
    End If
    GetCurrentValue = retval
End Function

Public Function StartPolling() As Boolean
    pollTimer.Enabled = True
    StartPolling = True
End Function

Public Function StopPolling() As Boolean
    pollTimer.Enabled = False
    StopPolling = True
End Function

Private Sub Class_Initialize()
    Set comm = Nothing
    
    currentCommPort = "1"
    currentBaudRate = "9600"
    currentParity = "E"
    currentDataBit = "7"
    currentStopBit = "1"
    
    Set pollTimer = New TimerClass
    pollTimer.Interval = 300
    pollTimer.Enabled = False
End Sub

Private Sub Class_Terminate()
    Set pollTimer = Nothing
    If Not comm Is Nothing Then
        If comm.PortOpen = True Then
            comm.PortOpen = False
        End If
        Set comm = Nothing
    End If
End Sub

Private Sub pollTimer_Timer()
    RaiseEvent WeightIncoming(GetCurrentValue)
End Sub

Private Function getSettingsAsString() As String
    getSettingsAsString = currentBaudRate & "," & _
                          currentParity & "," & _
                          currentDataBit & "," & _
                          currentStopBit
End Function

