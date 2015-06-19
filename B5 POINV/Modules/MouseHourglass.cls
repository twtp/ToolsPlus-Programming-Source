VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MouseHourglass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : MouseHourglass
' DateTime  : 5/25/2006 10:47
' Author    : briandonorfio
' Purpose   : Simple semaphore based mouse pointer control. Probably easiest thing to do
'             is create a project-global object for it.
'
'             Dependencies:
'               - none
'---------------------------------------------------------------------------------------

Option Explicit

Private semaphore As Long

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' DateTime  : 5/25/2006 10:49
' Author    : briandonorfio
' Purpose   : CTOR, nothing special here
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    semaphore = 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Hourglass
' DateTime  : 5/25/2006 10:49
' Author    : briandonorfio
' Purpose   : Call with true to set the mouse to an hourglass, call with false to unset
'             (maybe).
'---------------------------------------------------------------------------------------
'
Public Sub Hourglass(OnOrOff As Boolean)
    If OnOrOff = True Then
        semaphore = semaphore + 1
        If semaphore = 1 Then
            Screen.MousePointer = vbHourglass
        End If
    Else
        semaphore = semaphore - 1
        If semaphore = 0 Then
            Screen.MousePointer = vbNormal
        ElseIf semaphore < 0 Then
            semaphore = 0
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TempClear
' DateTime  : 5/25/2006 10:50
' Author    : briandonorfio
' Purpose   : Temporarily unsets/resets the mouse pointer state. Probably useful for
'             a fake dialog box, or something
'---------------------------------------------------------------------------------------
'
Public Sub TempClear(OnOrOff As Boolean)
    Screen.MousePointer = IIf(OnOrOff, vbNormal, vbHourglass)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Reset
' DateTime  : 5/25/2006 10:51
' Author    : briandonorfio
' Purpose   : Resets the semaphore and mouse state. No reason why you should have to use
'             this, other than if you're doing something wrong
'---------------------------------------------------------------------------------------
'
Public Sub Reset()
    semaphore = 0
    Screen.MousePointer = vbNormal
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' DateTime  : 5/25/2006 10:51
' Author    : briandonorfio
' Purpose   : DTOR, just clean up the mouse pointer.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    Screen.MousePointer = vbNormal
End Sub
