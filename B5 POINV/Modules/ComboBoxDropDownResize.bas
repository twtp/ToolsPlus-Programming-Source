Attribute VB_Name = "ComboBoxDropDownResize"
'---------------------------------------------------------------------------------------
' Module    : ComboBoxDropDownResize
' DateTime  : 4/26/2006 10:23
' Author    : briandonorfio
' Purpose   : Resizes the drop-down part of a combobox to the visible width (ie, checks
'             the font) of the longest item in the list, or the regular width if they're
'             all shorter.
'---------------------------------------------------------------------------------------

Option Explicit

Private Const CB_SETDROPPEDWIDTH = &H160
Private Const SM_CXVSCROLL = 2
Private Const WM_GETFONT = &H31
Private Const WU_LOGPIXELSX = 88
Private Const WU_LOGPIXELSY = 90

Private Type SIZE
    cx As Long
    cy As Long
End Type

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Public Sub ExpandDropDownToFit(ctl As ComboBox)
    SendMessage ctl.hwnd, CB_SETDROPPEDWIDTH, GetDropDownWidth(ctl), ByVal 0
End Sub

Private Function GetDropDownWidth(ctl As ComboBox) As Long
    Dim width As Long
    width = GetLongestStringSize(ctl)               'longest string
    width = width + 2                               '1px borders on each side
    width = width + GetSystemMetrics(SM_CXVSCROLL)  'plus scrollbar width
    width = width + 4                               'fudge factor
    GetDropDownWidth = width
End Function

Private Function GetLongestStringSize(ctl As ComboBox) As Long
    Dim curlen As Long, maxlen As Long
    Dim i As Long
    For i = 0 To ctl.ListCount - 1
        curlen = GetTextSize(ctl, i)
        If curlen > maxlen Then
            maxlen = curlen
        End If
    Next i
    GetLongestStringSize = maxlen
End Function

Private Function GetTextSize(ctl As ComboBox, i As Long) As Long
    Dim uSize As SIZE
    Dim hDC As Long, result As Long, font As Long, oldfont As Long
    hDC = GetDC(ctl.hwnd)                                                       'get the control device context
    font = SendMessage(ctl.hwnd, WM_GETFONT, 0, ByVal 0)                        'get the font the control is using
    oldfont = SelectObject(hDC, font)                                           'set the device context to use that font
    result = GetTextExtentPoint32(hDC, ctl.list(i), Len(ctl.list(i)), uSize)    'get displayed length of string
    GetTextSize = uSize.cx
    SelectObject hDC, oldfont                                                   'restore font for control device context
    ReleaseDC ctl.hwnd, hDC                                                     'close device context
End Function

Private Function ConvertTwipsToPixels(twips As Long, axis As Long) As Long
   Dim hDC As Long
   Dim PixelsPerInch As Long
   Const TwipsPerInch = 1440
   hDC = GetDC(0)

   If axis = 0 Then 'x
      PixelsPerInch = GetDeviceCaps(hDC, WU_LOGPIXELSX)
   Else 'y
      PixelsPerInch = GetDeviceCaps(hDC, WU_LOGPIXELSY)
   End If
   hDC = ReleaseDC(0, hDC)
   ConvertTwipsToPixels = (twips / TwipsPerInch) * PixelsPerInch
End Function

