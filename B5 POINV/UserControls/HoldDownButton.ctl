VERSION 5.00
Begin VB.UserControl HoldDownButton 
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3150
   ScaleHeight     =   1335
   ScaleWidth      =   3150
   Begin VB.Timer recurrenceTimer 
      Enabled         =   0   'False
      Left            =   1380
      Top             =   690
   End
   Begin VB.Timer initialTimer 
      Enabled         =   0   'False
      Left            =   870
      Top             =   690
   End
   Begin VB.CommandButton theButton 
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "HoldDownButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private stopNextClickEvent As Boolean

Public Event Click()

Public Property Get TimeBeforeHoldDown() As Long
Attribute TimeBeforeHoldDown.VB_ProcData.VB_Invoke_Property = ";Behavior"
    TimeBeforeHoldDown = UserControl.initialTimer.Interval
End Property
Public Property Let TimeBeforeHoldDown(newTime As Long)
    If CanPropertyChange("TimeBeforeHoldDown") Then
        UserControl.initialTimer.Interval = newTime
        PropertyChanged "TimeBeforeHoldDown"
    End If
End Property

Public Property Get TimeBetweenRepeats() As Long
Attribute TimeBetweenRepeats.VB_ProcData.VB_Invoke_Property = ";Behavior"
    TimeBetweenRepeats = UserControl.recurrenceTimer.Interval
End Property
Public Property Let TimeBetweenRepeats(newTime As Long)
    If CanPropertyChange("TimeBetweenRepeats") Then
        UserControl.recurrenceTimer.Interval = newTime
        PropertyChanged "TimeBetweenRepeats"
    End If
End Property

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Caption = UserControl.theButton.Caption
End Property
Public Property Let Caption(newCaption As String)
    If CanPropertyChange("Caption") Then
        UserControl.theButton.Caption = newCaption
        PropertyChanged "Caption"
    End If
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Font = UserControl.theButton.Font
End Property
Public Property Set Font(newFont As StdFont)
    If CanPropertyChange("Font") Then
        Set UserControl.theButton.Font = newFont
        PropertyChanged "Font"
    End If
End Property


Private Sub theButton_Click()
    If stopNextClickEvent Then
        stopNextClickEvent = False
    Else
        'Debug.Print "click: " & GetTickCount & ", raised event"
        RaiseEvent Click
    End If
End Sub

Private Sub theButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.initialTimer.Enabled = True
    'Debug.Print "mousedown: " & GetTickCount & ", initial timer enabled"
End Sub

Private Sub theButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.recurrenceTimer.Enabled = False
    UserControl.initialTimer.Enabled = False
    'Debug.Print "mouseup: " & GetTickCount & ", timers disabled"
End Sub

Private Sub theButton_LostFocus()
    If UserControl.recurrenceTimer.Enabled Then
        'Debug.Print "lostfocus, calling mouseup"
        theButton_MouseUp 0, 0, 0, 0
    End If
End Sub




Private Sub UserControl_Resize()
    UserControl.theButton.width = UserControl.width
    UserControl.theButton.Height = UserControl.Height
End Sub

Private Sub UserControl_GotFocus()
    UserControl.theButton.SetFocus
End Sub

Private Sub initialTimer_Timer()
    UserControl.initialTimer.Enabled = False
    UserControl.recurrenceTimer.Enabled = True
    'Debug.Print "initialTimer fired: " & GetTickCount & ", starting repeat"
    stopNextClickEvent = True
End Sub

Private Sub recurrenceTimer_Timer()
    UserControl.recurrenceTimer.Enabled = False
    'Debug.Print "recurrenceTimer fired: " & GetTickCount & ", raised event"
    RaiseEvent Click
    UserControl.recurrenceTimer.Enabled = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.initialTimer.Interval = PropBag.ReadProperty("TimeBeforeHoldDown", 500)
    UserControl.recurrenceTimer.Interval = PropBag.ReadProperty("TimeBetweenRepeats", 50)
    UserControl.theButton.Caption = PropBag.ReadProperty("Caption", "")
    UserControl.theButton.Font = PropBag.ReadProperty("Font", New StdFont)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "TimeBeforeHoldDown", UserControl.initialTimer.Interval
    PropBag.WriteProperty "TimeBetweenRepeats", UserControl.recurrenceTimer.Interval
    PropBag.WriteProperty "Caption", UserControl.theButton.Caption
    PropBag.WriteProperty "Font", UserControl.theButton.Font
End Sub
