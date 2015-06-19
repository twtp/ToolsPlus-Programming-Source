Attribute VB_Name = "CursorPosition"
'---------------------------------------------------------------------------------------
' Module    : CursorPosition
' DateTime  : 10/28/2005 09:43
' Author    : briandonorfio
' Purpose   : Sets the cursor position for a textbox/combobox when called. Options are
'             beginning of field, end of field, select all, default. Requires a global
'             variable CUR_POS_OPT
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : SetCursorPosition
' DateTime  : 8/17/2005 12:38
' Author    : briandonorfio
' Purpose   : Sets the global variable CUR_POS_OPT. Should be one of "start", "end", or
'             "select", invalid values will just use the default action for the control.
'---------------------------------------------------------------------------------------
'
Public Sub SetCursorPosition(pos As String)
    CUR_POS_OPT = pos
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HandleTextBoxCursorPosition
' DateTime  : 8/17/2005 12:39
' Author    : briandonorfio
' Purpose   : Sets the given textbox's cursor location to the position given by the
'             global variable CUR_POS_OPT, unless where string is included, in which
'             case it will override CUR_POS_OPT. Should be one of the following strings:
'             "start", "end", "select", "default". If it's not, it will just do the
'             default, no error message.
'---------------------------------------------------------------------------------------
'
Public Sub HandleTextBoxCursorPosition(field As TextBox, Optional where As String = "")
    Select Case IIf(where = "", CUR_POS_OPT, where)
        Case Is = "start"
            field.SelStart = 0
        Case Is = "end"
            field.SelStart = Len(field.Text)
        Case Is = "select"
            field.SelStart = 0
            field.SelLength = Len(field.Text)
        Case Is = "default"
            'do nothing
    End Select
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HandleComboBoxCursorPosition
' DateTime  : 8/17/2005 12:42
' Author    : briandonorfio
' Purpose   : Same as above, but for combo boxes.
'---------------------------------------------------------------------------------------
'
Public Sub HandleComboBoxCursorPosition(field As ComboBox, Optional where As String = "")
    If field.Style = vbComboDropdownList Then
        'do nothing
    Else
        Select Case IIf(where = "", CUR_POS_OPT, where)
            Case Is = "start"
                field.SelStart = 0
            Case Is = "end"
                field.SelStart = Len(field.Text)
            Case Is = "select"
                field.SelStart = 0
                field.SelLength = Len(field.Text)
            Case Is = "default"
                'do nothing
        End Select
    End If
End Sub
