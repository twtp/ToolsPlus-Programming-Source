VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl NumericUpDown 
   ClientHeight    =   1305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   ScaleHeight     =   1305
   ScaleWidth      =   2565
   Begin VB.Timer recurrenceTimer 
      Enabled         =   0   'False
      Left            =   1290
      Top             =   630
   End
   Begin VB.Timer initialTimer 
      Enabled         =   0   'False
      Left            =   480
      Top             =   600
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   2295
      TabIndex        =   1
      Top             =   0
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   2565
   End
End
Attribute VB_Name = "NumericUpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event ValueChanged()

Private stopChangeEvent As Boolean
Private propFormatString As String
Private propAllowDecimals As Boolean
Private propDefaultValue As Variant '/decimal

Private lastButton As Long

Public Property Get Value() As Variant
    If UserControl.Text1.Text = "" Then
        Value = CDec(0)
    Else
        Value = CDec(UserControl.Text1.Text)
    End If
End Property
Public Property Let Value(newval As Variant)
    If CanPropertyChange("Value") Then
        UserControl.Text1.Text = newval 'will call change() event to clear invalids
        PropertyChanged "Value"
    End If
End Property

Public Property Get Minimum() As Long
    Minimum = CLng(UserControl.UpDown1.Min)
End Property
Public Property Let Minimum(newval As Long)
    If CanPropertyChange("Minimum") Then
        UserControl.UpDown1.Min = newval
        PropertyChanged "Minimum"
    End If
End Property

Public Property Get Maximum() As Long
    Maximum = CLng(UserControl.UpDown1.Max)
End Property
Public Property Let Maximum(newval As Long)
    If CanPropertyChange("Maximum") Then
        UserControl.UpDown1.Max = newval
        PropertyChanged "Maximum"
    End If
End Property

Public Property Get Increment() As Long
    Increment = CLng(UserControl.UpDown1.Increment)
End Property
Public Property Let Increment(newval As Long)
    If CanPropertyChange("Increment") Then
        UserControl.UpDown1.Increment = newval
        PropertyChanged "Increment"
    End If
End Property

Public Property Get DisplayFormat() As String
    DisplayFormat = propFormatString
End Property
Public Property Let DisplayFormat(newval As String)
    If CanPropertyChange("DisplayFormat") Then
        propFormatString = newval 'any way to test a format string?
        PropertyChanged "DisplayFormat"
    End If
End Property

Public Property Get AllowDecimals() As Boolean
    AllowDecimals = propAllowDecimals
End Property
Public Property Let AllowDecimals(newval As Boolean)
    If CanPropertyChange("AllowDecimals") Then
        propAllowDecimals = newval
        PropertyChanged "AllowDecimals"
    End If
End Property

Public Property Get TimeBeforeHoldDown() As Long
    TimeBeforeHoldDown = UserControl.initialTimer.Interval
End Property
Public Property Let TimeBeforeHoldDown(newTime As Long)
    If CanPropertyChange("TimeBeforeHoldDown") Then
        UserControl.initialTimer.Interval = newTime
        PropertyChanged "TimeBeforeHoldDown"
    End If
End Property

Public Property Get TimeBetweenRepeats() As Long
    TimeBetweenRepeats = UserControl.recurrenceTimer.Interval
End Property
Public Property Let TimeBetweenRepeats(newTime As Long)
    If CanPropertyChange("TimeBetweenRepeats") Then
        UserControl.recurrenceTimer.Interval = newTime
        PropertyChanged "TimeBetweenRepeats"
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Text1.Enabled
End Property
Public Property Let Enabled(newEnabledness As Boolean)
    If CanPropertyChange("Enabled") Then
        UserControl.Text1.Enabled = newEnabledness
        UserControl.UpDown1.Enabled = newEnabledness
        PropertyChanged "Enabled"
    End If
End Property

Public Property Get BackColor() As Long
    BackColor = UserControl.Text1.BackColor
End Property
Public Property Let BackColor(newBackColor As Long)
    If CanPropertyChange("BackColor") Then
        UserControl.Text1.BackColor = newBackColor
        PropertyChanged "BackColor"
    End If
End Property

Public Property Get DefaultValue() As Variant
    DefaultValue = propDefaultValue
End Property
Public Property Let DefaultValue(newDefaultValue As Variant)
    If CanPropertyChange("DefaultValue") Then
        propDefaultValue = newDefaultValue
        PropertyChanged "DefaultValue"
    End If
End Property



'Private Sub Text1_Change()
'    If stopChangeEvent = True Then
'        Exit Sub
'    End If
'    'second line of defense, here we block r-click paste after the fact, or
'    'any other possible ways for rogue characters to get in
'    Dim i As Long
'    Dim newstr As String
'    newstr = ""
'    Dim haveDecimalPoint
'    haveDecimalPoint = Not propAllowDecimals 'if we're not allowing decimals, then just assume there's one there
'    For i = 1 To Len(UserControl.Text1)
'        Dim c As String
'        c = Mid(UserControl.Text1, i, 1)
'        If haveDecimalPoint = False And c = "." Then
'            newstr = newstr & c
'            haveDecimalPoint = True
'        ElseIf Asc(c) >= Asc("0") And Asc(c) <= Asc("9") Then
'            newstr = newstr & c
'        End If
'    Next i
'    If newstr = "" Then
'        newstr = "0"
'    End If
'    stopChangeEvent = True
'    UserControl.Text1 = Format(newstr, propFormatString)
'    RaiseEvent ValueChanged
'    stopChangeEvent = False
'End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    'this is the first line of defense against bad characters, stop anything from
    'a regular keypress. this also handles the problem of ctrl+v.
    If propAllowDecimals = True And KeyAscii = Asc(".") Then
        If CBool(InStr(UserControl.Text1.Text, ".") = 0) Then
            'ok
            RaiseEvent ValueChanged
        Else
            'already have a decimal point
            KeyAscii = 0
        End If
    ElseIf KeyAscii <> vbKeyBack And (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
        KeyAscii = 0
    ElseIf UserControl.Text1 = "" And KeyAscii = Asc("0") Then
        KeyAscii = 0
    Else
        'ok
        RaiseEvent ValueChanged
    End If
End Sub

Private Sub Text1_LostFocus()
    If stopChangeEvent = True Then
        Exit Sub
    End If
    'second line of defense, here we block r-click paste after the fact, or
    'any other possible ways for rogue characters to get in
    Dim i As Long
    Dim newstr As String
    newstr = ""
    Dim haveDecimalPoint
    haveDecimalPoint = Not propAllowDecimals 'if we're not allowing decimals, then just assume there's one there
    For i = 1 To Len(UserControl.Text1)
        Dim c As String
        c = Mid(UserControl.Text1, i, 1)
        If haveDecimalPoint = False And c = "." Then
            newstr = newstr & c
            haveDecimalPoint = True
        ElseIf Asc(c) >= Asc("0") And Asc(c) <= Asc("9") Then
            newstr = newstr & c
        End If
    Next i
    If newstr = "" Then
        newstr = propDefaultValue
    End If
    stopChangeEvent = True
    UserControl.Text1 = Format(newstr, propFormatString)
    RaiseEvent ValueChanged
    stopChangeEvent = False
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'apparently there's no way to cancel a click event? that should go here if there is
End Sub

Private Sub UpDown1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Debug.Print "mousedown"
    UpDown1_GenericClickHandler
    UserControl.initialTimer.Enabled = True
End Sub

Private Sub UpDown1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim newLastButton As Long
    newLastButton = IIf(Y < (UserControl.UpDown1.Height / 2), 1, -1)
    If newLastButton <> lastButton Then
        UserControl.recurrenceTimer.Enabled = False
        UserControl.initialTimer.Enabled = False
    End If
    lastButton = newLastButton
End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Debug.Print "mouseup"
    UserControl.recurrenceTimer.Enabled = False
    UserControl.initialTimer.Enabled = False
End Sub

Private Sub UpDown1_GenericClickHandler()
'Debug.Print "clickhandler running, lastButton is " & lastButton
    Dim currVal As Variant, nextVal As Variant
    currVal = CDec(UserControl.Text1)
    nextVal = CDec(currVal + lastButton * Increment)
    If nextVal > Maximum Then
        nextVal = Maximum
    ElseIf nextVal < Minimum Then
        nextVal = Minimum
    End If
    If currVal <> nextVal Then
        stopChangeEvent = True
        UserControl.Text1 = Format(nextVal, propFormatString)
        RaiseEvent ValueChanged
        stopChangeEvent = False
    End If
End Sub

Private Sub initialTimer_Timer()
'Debug.Print "initial fired"
    UserControl.initialTimer.Enabled = False
    UserControl.recurrenceTimer.Enabled = True
End Sub

Private Sub recurrenceTimer_Timer()
'Debug.Print "recurrence fired"
    UpDown1_GenericClickHandler
End Sub

Private Sub UserControl_Resize()
    UserControl.Text1.Width = UserControl.Width
    UserControl.Text1.Height = UserControl.Height
    UserControl.UpDown1.Width = 240
    UserControl.UpDown1.Height = UserControl.Height - 60
    UserControl.UpDown1.Left = UserControl.Width - UserControl.UpDown1.Width - 30
    UserControl.UpDown1.Top = 30
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Text1.Text = PropBag.ReadProperty("Value", 1)
    UserControl.UpDown1.Min = PropBag.ReadProperty("Minimum", 0)
    UserControl.UpDown1.Max = PropBag.ReadProperty("Maximum", 99999)
    UserControl.UpDown1.Increment = PropBag.ReadProperty("Increment", 1)
    propFormatString = PropBag.ReadProperty("DisplayFormat", "#0")
    propAllowDecimals = PropBag.ReadProperty("AllowDecimals", False)
    UserControl.initialTimer.Interval = PropBag.ReadProperty("TimeBeforeHoldDown", 500)
    UserControl.recurrenceTimer.Interval = PropBag.ReadProperty("TimeBetweenRepeats", 50)
    UserControl.Text1.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.UpDown1.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.Text1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    propDefaultValue = PropBag.ReadProperty("DefaultValue", 0)
    
    UserControl.Text1.Text = Format(UserControl.Text1.Text, propFormatString)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Value", UserControl.Text1.Text
    PropBag.WriteProperty "Minimum", UserControl.UpDown1.Min
    PropBag.WriteProperty "Maximum", UserControl.UpDown1.Max
    PropBag.WriteProperty "Increment", UserControl.UpDown1.Increment
    PropBag.WriteProperty "DisplayFormat", propFormatString
    PropBag.WriteProperty "AllowDecimals", propAllowDecimals
    PropBag.WriteProperty "TimeBeforeHoldDown", UserControl.initialTimer.Interval
    PropBag.WriteProperty "TimeBetweenRepeats", UserControl.recurrenceTimer.Interval
    PropBag.WriteProperty "Enabled", UserControl.Text1.Enabled
    PropBag.WriteProperty "BackColor", UserControl.Text1.BackColor
    PropBag.WriteProperty "DefaultValue", propDefaultValue
End Sub


