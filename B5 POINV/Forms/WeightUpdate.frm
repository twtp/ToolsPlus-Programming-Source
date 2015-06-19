VERSION 5.00
Begin VB.Form WeightUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6120
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox ConnectionState 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4410
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Connected"
      Top             =   1140
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton CloseButton 
      Caption         =   "Close"
      Height          =   405
      Left            =   2520
      TabIndex        =   10
      Top             =   3210
      Width           =   1485
   End
   Begin VB.ComboBox CommPort 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   120
      Width           =   1395
   End
   Begin VB.CommandButton ConnectScale 
      Caption         =   "Connect Scale"
      Height          =   525
      Left            =   4320
      TabIndex        =   8
      Top             =   540
      Width           =   1545
   End
   Begin VB.CommandButton SaveButton 
      Caption         =   "Save"
      Height          =   465
      Left            =   1170
      TabIndex        =   7
      Top             =   2160
      Width           =   1875
   End
   Begin VB.TextBox WeightNew 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   1680
      Width           =   1965
   End
   Begin VB.TextBox WeightCurrent 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1260
      Width           =   1965
   End
   Begin VB.ComboBox ItemNumber 
      Height          =   315
      Left            =   1170
      TabIndex        =   2
      Top             =   630
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "New Weight:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   1710
      Width           =   1545
   End
   Begin VB.Label Label3 
      Caption         =   "Current Weight:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   1290
      Width           =   1545
   End
   Begin VB.Label Label2 
      Caption         =   "Select Item:"
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   690
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Weight Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   3525
   End
End
Attribute VB_Name = "WeightUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents scaleConn As ScaleInterface.MettlerToledoScale
Attribute scaleConn.VB_VarHelpID = -1

Private Sub CloseButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    requeryCOMPorts
    requeryItems
    
    Set scaleConn = New ScaleInterface.MettlerToledoScale
    scaleConn.CommPort = Me.CommPort
    scaleConn.BaudRate = "9600"
    scaleConn.Parity = "E"
    scaleConn.DataBit = "7"
    scaleConn.StopBit = "1"
    
    If scaleConn.OpenConn Then
        Me.ConnectionState.Visible = True
        scaleConn.StartPolling
    Else
        Me.ConnectionState.Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    scaleConn.StopPolling
    scaleConn.CloseConn
    Set scaleConn = Nothing
End Sub

Private Sub ItemNumber_Click()
    ItemNumber_LostFocus
End Sub

Private Sub ItemNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.ItemNumber, KeyCode, Shift
End Sub

Private Sub ItemNumber_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.ItemNumber, KeyAscii
End Sub

Private Sub ItemNumber_LostFocus()
    AutoCompleteLostFocus Me.ItemNumber
    
    If Me.ItemNumber = "" Then
        Me.WeightCurrent = ""
    Else
        Me.WeightCurrent = DLookup("Weight", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'")
    End If
    
    setSaveButtonEnableStatus
End Sub

Private Sub requeryItems()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber FROM vNonDiscontinuedItems")
    Me.ItemNumber.Clear
    While Not rst.EOF
        Me.ItemNumber.AddItem rst("ItemNumber")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryCOMPorts()
    Me.CommPort.Clear
    Me.CommPort.AddItem "COM1"
    Me.CommPort.AddItem "COM2"
    Me.CommPort.AddItem "COM3"
    Me.CommPort.AddItem "COM4"
    Me.CommPort = "COM1"
End Sub

Private Sub SaveButton_Click()
    If Me.ItemNumber = "" Or Me.WeightNew = "" Or Not IsNumeric(Me.WeightNew) Then
        Exit Sub
    End If
    
    'Mouse.Hourglass True
    'scaleConn.StopPolling
    'If vbYes = MsgBox("Set weight for " & qq(Me.ItemNumber) & " to " & Me.WeightNew & "?", vbYesNo) Then
    '    updateInventoryMaster "Weight", Me.WeightNew, Me.ItemNumber, "'"
    '    updateInventoryMaster "BoxID", DetermineBoxSize(DLookup("Dimensions", "InventoryMaster", "ItemNumber='" & Me.ItemNumber & "'"), Me.WeightNew), Me.ItemNumber, "'"
    '    If IsNumeric(Me.WeightNew) Then
    '        If CSng(Me.WeightNew) > 150 Then
    '            If vbYes = MsgBox("Over 150 lbs! Should this be set to ""Truck Freight""?", vbYesNo) Then
    '                updateInventoryMaster "TruckFreight", "1", Me.ItemNumber
    '                updateInventoryMaster "TruckFreightCheap", "0", Me.ItemNumber
    '            End If
    '        End If
    '    End If
    '    SyncExternalComponentDB "", Me.ItemNumber
    '
    '    Me.WeightCurrent = ""
    '    Me.ItemNumber = ""
    'End If
    'scaleConn.StartPolling
    
    MsgBox "TODO: Broken!"
    
    Mouse.Hourglass False
End Sub

Private Sub scaleConn_WeightIncoming(incoming As ScaleInterface.ScaleWeight)
    Me.WeightNew = incoming.Weight
    setSaveButtonEnableStatus
End Sub

Private Sub setSaveButtonEnableStatus()
    If Me.ItemNumber = "" Then
        Me.SaveButton.Enabled = False
    ElseIf Not IsNumeric(Me.WeightNew) Then
        Me.SaveButton.Enabled = False
    Else
        Me.SaveButton.Enabled = True
    End If
End Sub
