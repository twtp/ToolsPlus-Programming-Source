Attribute VB_Name = "AutoComplete"
'---------------------------------------------------------------------------------------
' Module    : AutoComplete
' DateTime  : 6/14/2005 10:31
' Author    : briandonorfio
' Purpose   : Provides autocompletion for comboboxes
'
'             Dependencies:
'               - none
'---------------------------------------------------------------------------------------

Option Explicit

Private Const CB_SETCURSEL = &H14E

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'---------------------------------------------------------------------------------------
' Procedure : AutoCompleteKeyDown
' DateTime  : 6/14/2005 11:36
' Author    : briandonorfio
' Purpose   : Should be called by the ComboBox_KeyDown event.
'---------------------------------------------------------------------------------------
'
Public Sub AutoCompleteKeyDown(ctl As ComboBox, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        ctl.Text = ""
        'KeyCode = 0
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AutoCompleteKeyPress
' DateTime  : 6/14/2005 11:37
' Author    : briandonorfio
' Purpose   : Should be called by the ComboBox_KeyPress event.
'---------------------------------------------------------------------------------------
'
Public Sub AutoCompleteKeyPress(ctl As ComboBox, KeyAscii As Integer)
    Dim SearchText As String, enteredText As String
    Dim Length As Long, Index As Long, counter As Long

    With ctl
        If .selStart > 0 Then
            enteredText = Left$(.Text, .selStart)
        End If

        Select Case KeyAscii
            Case vbKeyReturn
                If .ListIndex > -1 Then
                    .selStart = 0
                    .selLength = Len(.list(.ListIndex))
                    Exit Sub
                End If
            'Case vbKeyEscape, vbKeyDelete  'delete is handled through keydown, and has the same code as period, so can't go here
            Case vbKeyEscape
                .Text = ""
                KeyAscii = 0
                Exit Sub
            Case vbKeyBack
                If Len(enteredText) > 1 Then
                    SearchText = LCase$(Left$(enteredText, Len(enteredText) - 1))
                Else
                    enteredText = ""
                    'KeyAscii = 0
                    .Text = ""
                    Exit Sub
                End If
            Case Else
                SearchText = LCase$(enteredText & Chr(KeyAscii))
        End Select

        Index = -1
        Length = Len(SearchText)
        For counter = 0 To .ListCount - 1
            If LCase$(Left$(.list(counter), Length)) = SearchText Then
                Index = counter
                Exit For
            End If
        Next counter
        If Index > -1 Then
            '.ListIndex = index
            SetListCtlIndex ctl, Index
            .selStart = Len(SearchText)
            .selLength = Len(.list(Index)) - Len(SearchText)
        End If
    End With
    KeyAscii = 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AutoCompleteLostFocus
' DateTime  : 6/14/2005 11:37
' Author    : briandonorfio
' Purpose   : Should be called by the ComboBox_LostFocus event.
'---------------------------------------------------------------------------------------
'
Public Sub AutoCompleteLostFocus(ctl As ComboBox)
    ctl.selLength = 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetListCtlIndex
' DateTime  : 6/14/2005 11:37
' Author    : briandonorfio
' Purpose   : Changes the ComboBox.Index value without calling the *_Click() event.
'---------------------------------------------------------------------------------------
'
Private Sub SetListCtlIndex(ctl As ComboBox, ByVal Index As Long)
    SendMessage ctl.hwnd, CB_SETCURSEL, Index, 0&
End Sub


