Attribute VB_Name = "TabStops"
'---------------------------------------------------------------------------------------
' Module    : TabStops
' DateTime  : 6/14/2005 10:37
' Author    : briandonorfio
' Purpose   : Sets tab stops in a combobox.
'
'             Dependencies:
'               - none
'---------------------------------------------------------------------------------------

Option Explicit

Private Const LB_SETTABSTOPS = &H192

'---------------------------------------------------------------------------------------
' Procedure : SetListTabStops
' DateTime  : 6/14/2005 11:08
' Author    : briandonorfio
' Purpose   : Sets tab stops in a combobox. iListHandle should be the Me.ComboBox.hWnd.
'             The tabs array is an array of the tab stops, not sure what units it's in,
'             pixels maybe, or twips, but it's roughly 6 per character.
'---------------------------------------------------------------------------------------
'
Public Sub SetListTabStops(iListHandle As Long, tabs() As Long)
    SendMessage iListHandle, LB_SETTABSTOPS, UBound(tabs) + 1, tabs(0)
End Sub

