Attribute VB_Name = "Utilities"
'---------------------------------------------------------------------------------------
' Module    : Utilities
' DateTime  : 6/14/2005 10:38
' Author    : briandonorfio
' Purpose   : Miscellaneous utility functions, and functions that need to be called from
'             multiple forms.
'
'             Dependencies:
'               - none! well maybe globals.bas for function declarations
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Integer
Private Declare Function GetIpAddrTable Lib "IpHlpApi" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long

Private Const WTS_CURRENT_SERVER_HANDLE As Long = 0
Private Const WTS_CURRENT_SESSION As Long = -1
Private Const AF_INET As Long = 2

Private Enum WTS_INFO_CLASS
    WTSInitialProgram
    WTSApplicationName
    WTSWorkingDirectory
    WTSOEMId
    WTSSessionId
    WTSUserName
    WTSWinStationName
    WTSDomainName
    WTSConnectState
    WTSClientBuildNumber
    WTSClientName
    WTSClientDirectory
    WTSClientProductId
    WTSClientHardwareId
    WTSClientAddress
    WTSClientDisplay
End Enum

Private Type WTS_CLIENT_ADDRESS
    AddressFamily As Long
    addr(20) As Byte
End Type

Private Declare Sub RtlMoveMemory Lib "kernel32" (pDest As Any, ByVal pSource As Long, ByVal ByteLen As Long)
Private Declare Function WTSQuerySessionInformationA Lib "wtsapi32.dll" (ByVal hServer As Long, _
                                                                         ByVal SessionId As Long, _
                                                                         ByVal WTSInfoClass As WTS_INFO_CLASS, _
                                                                         ByRef ppBuffer As Long, _
                                                                         ByRef pBytes As Long _
                                                                        ) As Long

'for TabStops
Private Const LB_SETTABSTOPS = &H192
'for AutoComplete
Private Const CB_SETCURSEL = &H14E
'for SyncShell
Private Const STARTF_USESHOWWINDOW& = &H1
Private Const STARTF_USESTDHANDLES = &H100
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const SW_HIDE = 0
Private Const INFINITE = -1&

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
'for Clipboard
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long

Private Const GHND = &H42
Private Const CF_TEXT = 1
Private Const CF_BITMAP = 2
Private Const MAXSIZE = 4096
'for ComboBoxDropDownResize
Private Const CB_SETDROPPEDWIDTH = &H160
Private Const SM_CXVSCROLL = 2
Private Const WM_GETFONT = &H31
Private Const WU_LOGPIXELSX = 88
Private Const WU_LOGPIXELSY = 90

Private Type Size
    cx As Long
    cy As Long
End Type

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "GDI32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'for Base64
Private Const BASE64_CHARACTER_SET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
'for ScreenCap
Private Const RASTERCAPS As Integer = 38
Private Const RC_PALETTE As Integer = &H100
Private Const SIZEPALETTE As Integer = 104

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
End Type

Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type FLASHWINFO
  cbSize As Long
  hWnd As Long
  dwFlags As Long
  uCount As Long
  dwTimeout As Long
End Type

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function FlashWindowEx Lib "user32" (FWInfo As FLASHWINFO) As Boolean

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
'Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long


'---------------------------------------------------------------------------------------
' Procedure : EscapeSQuotes
' DateTime  : 6/14/2005 10:54
' Author    : briandonorfio
' Purpose   : Escapes single quotes for SQL strings.
'---------------------------------------------------------------------------------------
'
Public Function EscapeSQuotes(txt As String) As String
    EscapeSQuotes = Replace(txt, "'", "''")
End Function

'---------------------------------------------------------------------------------------
' Procedure : EscapeDQuotes
' DateTime  : 7/5/2005 17:09
' Author    : briandonorfio
' Purpose   : Escapes double quotes for CSV files.
'---------------------------------------------------------------------------------------
'
Public Function EscapeDQuotes(txt As String) As String
    EscapeDQuotes = Replace(txt, """", """""")
End Function

'---------------------------------------------------------------------------------------
' Procedure : qq
' DateTime  : 10/5/2005 10:11
' Author    : briandonorfio
' Purpose   : Quick quote, puts literal double quotes around a string. NOT LIKE PERL'S.
'---------------------------------------------------------------------------------------
'
Public Function qq(txt As String) As String
    qq = """" & txt & """"
End Function

'---------------------------------------------------------------------------------------
' Procedure : StripDQuotes
' DateTime  : 6/14/2005 10:54
' Author    : briandonorfio
' Purpose   : Removes double quotes.
'---------------------------------------------------------------------------------------
'
Public Function StripDQuotes(txt As String) As String
    StripDQuotes = Replace(txt, """", "")
End Function

'---------------------------------------------------------------------------------------
' Procedure : Nz
' DateTime  : 6/14/2005 10:55
' Author    : briandonorfio
' Purpose   : Similar to Access 2k's Nz() function, if value is null, returns the value
'             given in default or null string if not given, otherwise returns value.
'---------------------------------------------------------------------------------------
'
Public Function Nz(txt As Variant, Optional default As String = "") As String
    Nz = IIf(IsNull(txt), default, txt)
End Function

'---------------------------------------------------------------------------------------
' Procedure : SQLBool
' DateTime  : 6/14/2005 10:56
' Author    : briandonorfio
' Purpose   : Converts a true/false boolean into an SQL 1/0.
'---------------------------------------------------------------------------------------
'
Public Function SQLBool(value As Boolean) As Long
    SQLBool = IIf(value, 1, 0)
End Function

'---------------------------------------------------------------------------------------
' Procedure : ConvertFracToDec
' DateTime  : 6/14/2005 10:58
' Author    : briandonorfio
' Purpose   : Converts fractions to decimals. Fractions can be formatted like 2 1/2 or
'             2-1/2. Adds a leading 0 if less than 1.
'---------------------------------------------------------------------------------------
'
Public Function ConvertFracToDec(ByVal fraction As String) As String
    Dim fracarray As Variant, frac As Variant   'fracarray is the array split across \s or -, frac is the
    If InStr(fraction, "/") Then                'fraction part split across /
        If InStr(fraction, " ") Then
            fracarray = Split(fraction, " ")
            frac = Split(fracarray(1), "/")
            fraction = fracarray(0) + (frac(0) / frac(1))
        ElseIf InStr(fraction, "-") Then
            fracarray = Split(fraction, "-")
            frac = Split(fracarray(1), "/")
            fraction = fracarray(0) + (frac(0) / frac(1))
        Else
            frac = Split(fraction, "/")
            fraction = (frac(0) / frac(1))
        End If
    End If
    If Left$(fraction, 1) = "." Then
        fraction = "0" & fraction
    End If
    ConvertFracToDec = fraction
End Function

'---------------------------------------------------------------------------------------
' Procedure : ForceUppercase
' DateTime  : 8/29/2005 11:27
' Author    : briandonorfio
' Purpose   : Converts a given ascii character to its uppercase equivalent.
'---------------------------------------------------------------------------------------
'
Public Function ForceUppercase(KeyAscii As Integer) As Integer
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    ForceUppercase = KeyAscii
End Function

'---------------------------------------------------------------------------------------
' Procedure : LimitToNumbers
' DateTime  : 1/12/2006 10:30
' Author    : briandonorfio
' Purpose   : Clears non-numeric character codes
'---------------------------------------------------------------------------------------
'
Public Function LimitToNumbers(KeyAscii As Integer, Optional decimalOk As Boolean = False, Optional fractionOk As Boolean = False) As Integer
    If decimalOk And KeyAscii = Asc(".") Then
        LimitToNumbers = KeyAscii
    ElseIf fractionOk And (KeyAscii = Asc("-") Or KeyAscii = Asc("/") Or KeyAscii = Asc(" ")) Then
        LimitToNumbers = KeyAscii
    ElseIf KeyAscii <> vbKeyBack And (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
        LimitToNumbers = 0
    Else
        LimitToNumbers = KeyAscii
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : LimitToLetters
' DateTime  : 1/12/2006 10:31
' Author    : briandonorfio
' Purpose   : Clears non-alphabetical character codes
'---------------------------------------------------------------------------------------
'
Public Function LimitToLetters(KeyAscii As Integer) As Integer
    If KeyAscii <> vbKeyBack And (Asc(UCase(Chr(KeyAscii))) < Asc("A") Or Asc(UCase(Chr(KeyAscii))) > Asc("Z")) Then
        LimitToLetters = 0
    Else
        LimitToLetters = KeyAscii
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : LimitToCurrency
' DateTime  : 1/13/2006 10:19
' Author    : briandonorfio
' Purpose   : Clears non-currency character codes, so it accepts $,.0-9
'---------------------------------------------------------------------------------------
'
Public Function LimitToCurrency(KeyAscii As Integer) As Integer
    If KeyAscii = vbKeyBack Or KeyAscii = Asc("$") Or KeyAscii = Asc(".") Or KeyAscii = Asc(",") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
        LimitToCurrency = KeyAscii
    Else
        LimitToCurrency = 0
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : TitleCase
' DateTime  : 9/15/2005 09:30
' Author    : briandonorfio
' Purpose   : Converts a given string to Title Case.
'---------------------------------------------------------------------------------------
'
Public Function TitleCase(txt As String) As String
    Dim char As String, retval As String
    Dim nextUpper As Boolean
    nextUpper = True
    Dim i As Long
    For i = 1 To Len(txt)
        char = Mid(txt, i, 1)
        If nextUpper Then
            char = UCase(char)
        Else
            char = LCase(char)
        End If
        If char = " " Or char = "-" Or char = "." Then
            nextUpper = True
        Else
            nextUpper = False
        End If
        retval = retval & char
    Next i
    TitleCase = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : Enable
' DateTime  : 6/14/2005 10:59
' Author    : briandonorfio
' Purpose   : Enables a control (probably a text box or combobox) and sets the color to
'             the default white background. Similar to how Access 2k's enable/disable
'             works.
'---------------------------------------------------------------------------------------
'
Public Sub Enable(ctl As Control)
    If TypeOf ctl Is OptionButton Or TypeOf ctl Is CheckBox Then
        ctl.BackColor = &H8000000F
    Else
        ctl.BackColor = &H80000005
    End If
    ctl.Enabled = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Disable
' DateTime  : 6/14/2005 11:00
' Author    : briandonorfio
' Purpose   : Disables a control and changes the background to dark grey, similar to
'             how Access 2k's enable/disable works.
'---------------------------------------------------------------------------------------
'
Public Sub Disable(ctl As Control)
    ctl.BackColor = &H8000000F
    ctl.Enabled = False
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsFormLoaded
' DateTime  : 6/29/2005 16:33
' Author    : briandonorfio
' Purpose   : Returns true if the form is loaded, false otherwise. Doesn't check whether
'             the form is show/hide
'---------------------------------------------------------------------------------------
'
Public Function IsFormLoaded(formName As String) As Boolean
    'IsFormLoaded = False
    'Dim frm As Form
    'For Each frm In Forms
    '    If frm.name = formName Then
    '        IsFormLoaded = True
    '    End If
    'Next frm
    IsFormLoaded = Not CBool(GetFormByName(formName) Is Nothing)
End Function

'---------------------------------------------------------------------------------------
' Procedure : ResizeAndInsertPic
' DateTime  : 7/27/2005 13:31
' Author    : briandonorfio
' Purpose   : Sets image at given path to given image control. If path is invalid, no
'             picture displayed. Resizes to given width/height, while maintaining the
'             picture's aspect ratio.
'
'             Integrated w/ TPPictureCtl.ocx, no longer needed here.
'---------------------------------------------------------------------------------------
'
'Public Sub ResizeAndInsertPic(path As String, imgCtl As Image, maxHeight As Integer, maxWidth As Integer)
'On Error GoTo errh
'    If Dir$(path) = "" Then
'        imgCtl.Picture = LoadPicture("")
'        imgCtl.tag = ""
'    Else
'        Dim pic As Picture
'        Set pic = LoadPicture(path)
'        If (pic.Height / maxHeight) > (pic.width / maxWidth) Then
'            imgCtl.Height = maxHeight
'            imgCtl.width = pic.width / (pic.Height / maxHeight)
'        Else
'            imgCtl.width = maxWidth
'            imgCtl.Height = pic.Height / (pic.width / maxWidth)
'        End If
'        imgCtl.Picture = pic
'        imgCtl.tag = path
'    End If
'    Exit Sub
'
'errh:
'    If Err.Number = 481 Then
'        MsgBox "Invalid picture file. Open it in Corel, click Image->Convert To, and select RGB 24-bit color, and resave it."
'    ElseIf Err.Number = 52 Then
'        'bad file name or number, probably a missing or invalid directory (CON, etc.)
'    Else
'        MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
'    End If
'    imgCtl.Picture = LoadPicture("")
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : ROT13
' DateTime  : 8/25/2005 09:21
' Author    : briandonorfio
' Purpose   : For "encrypting" item numbers.
'---------------------------------------------------------------------------------------
'
Public Function ROT13(txt As String) As String
    txt = UCase(txt)
    Dim i As Long
    Dim retval As String, char As String
    For i = 0 To Len(txt) - 1
        char = Mid(txt, i + 1, 1)
        If Asc(char) >= 65 And Asc(char) <= 77 Then 'a-m
            char = Chr(Asc(char) + 13)
        ElseIf Asc(char) >= 78 And Asc(char) <= 90 Then 'n-z
            char = Chr(Asc(char) - 13)
        ElseIf Asc(char) >= 48 And Asc(char) <= 52 Then '0-4
            char = Chr(Asc(char) + 5)
        ElseIf Asc(char) >= 53 And Asc(char) <= 57 Then '5-9
            char = Chr(Asc(char) - 5)
        ElseIf char = "-" Then
            'do nothing
        End If
        retval = retval & char
    Next i
    ROT13 = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : OpenDefaultApp
' DateTime  : 7/27/2005 13:30
' Author    : briandonorfio
' Purpose   : Opens the default app for a file or url.
'---------------------------------------------------------------------------------------
'
Public Function OpenDefaultApp(fname As String) As String
    ShellExecute 0, vbNullString, fname, vbNullString, vbNullString, vbNormalFocus
End Function

'---------------------------------------------------------------------------------------
' Procedure : ListBoxMoveUp
' DateTime  : 8/24/2005 11:39
' Author    : briandonorfio
' Purpose   : Moves a selected line in a listbox up, don't know how it will handle
'             multi-select list boxes.
'---------------------------------------------------------------------------------------
'
Public Function ListBoxMoveUp(ctl As ListBox) As Boolean
    If ctl.SelCount = 1 And ctl.ListIndex <> 0 Then
        Dim temp As String
        temp = ctl.list(ctl.ListIndex - 1)
        ctl.list(ctl.ListIndex - 1) = ctl.Text
        ctl.list(ctl.ListIndex) = temp
        ctl.Selected(ctl.ListIndex - 1) = True
        ListBoxMoveUp = True
    Else
        ListBoxMoveUp = False
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : ListBoxMoveDown
' DateTime  : 8/24/2005 11:40
' Author    : briandonorfio
' Purpose   : Moves a selected line in a listbox down, don't know how it will handle
'             multi-select list boxes.
'---------------------------------------------------------------------------------------
'
Public Sub ListBoxMoveDown(ctl As ListBox)
    If ctl.SelCount = 1 And ctl.ListIndex <> ctl.ListCount - 1 Then
        Dim temp As String
        temp = ctl.list(ctl.ListIndex + 1)
        ctl.list(ctl.ListIndex + 1) = ctl.Text
        ctl.list(ctl.ListIndex) = temp
        ctl.Selected(ctl.ListIndex + 1) = True
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : URLify
' DateTime  : 12/22/2005 15:17
' Author    : briandonorfio
' Purpose   : Returns the (path name, item name?) text as the url-part that yahoo would
'             turn it into. Power Tool "Accessories" becomes power-tool--accessories-.
'---------------------------------------------------------------------------------------
'
Public Function URLify(txt As String) As String
    Dim thischar As String, retval As String
    Dim i As Long
    For i = 1 To Len(txt)
        thischar = Mid(txt, i, 1)
        If Not IsAlphaNumeric(thischar) And thischar <> "-" Then
            retval = retval & "-"
        Else
            retval = retval & LCase(thischar)
        End If
    Next i
    URLify = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : UnFrontpage
' DateTime  : 12/26/2005 09:32
' Author    : briandonorfio
' Purpose   : Removes CP1252-specific characters from html, replaces with the normal
'             character or html char escape seq.
'---------------------------------------------------------------------------------------
'
Public Function UnFrontpage(html As String) As String
    Dim retval As String
    retval = html
    retval = Replace(retval, Chr(130), "'")
    retval = Replace(retval, Chr(132), "&quot;")
    retval = Replace(retval, Chr(133), "...")
    retval = Replace(retval, Chr(145), "'")
    retval = Replace(retval, Chr(146), "'")
    retval = Replace(retval, Chr(147), "&quot;")
    retval = Replace(retval, Chr(148), "&quot;")
    retval = Replace(retval, Chr(150), "&mdash;") 'technically an en dash, but i like em more
    retval = Replace(retval, Chr(151), "&mdash;")
    retval = Replace(retval, Chr(153), "<small><sup>TM</sup></small>")
    retval = Replace(retval, Chr(160), " ")       'non-breaking space
    retval = Replace(retval, Chr(169), "<small><sup>&copy;</sup></small>")
    retval = Replace(retval, Chr(173), "-")
    retval = Replace(retval, Chr(174), "<small><sup>&reg;</sup></small>")
    retval = Replace(retval, Chr(176), "&deg;")
    retval = Replace(retval, Chr(177), "&plusmn;")
    retval = Replace(retval, Chr(181), "&micro;")
    retval = Replace(retval, Chr(186), "&deg;")
    retval = Replace(retval, Chr(188), "<small><sup>1</sup></small>/<small><sub>4</sub></small>")
    retval = Replace(retval, Chr(189), "<small><sup>1</sup></small>/<small><sub>2</sub></small>")
    retval = Replace(retval, Chr(190), "<small><sup>3</sup></small>/<small><sub>4</sub></small>")
    retval = Replace(retval, Chr(224), "&agrave;") 'do we need more foreign chars?
    retval = Replace(retval, Chr(225), "&aacute;")
    retval = Replace(retval, Chr(226), "&acirc;")
    retval = Replace(retval, Chr(227), "&atilde;")
    retval = Replace(retval, Chr(231), "&ccedil;")
    retval = Replace(retval, Chr(232), "&egrave;")
    retval = Replace(retval, Chr(233), "&eacute;")
    retval = Replace(retval, Chr(234), "&ecirc;")
    retval = Replace(retval, Chr(235), "&euml;")
    UnFrontpage = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : Chomp
' DateTime  : 1/6/2006 10:58
' Author    : briandonorfio
' Purpose   : similar to perl chomp. remove trailing line ending from a string, if it's
'             there. returns chomped string (not like perl). checks for all three line
'             endings: 0x0D0A, 0x0D, 0x0A.
'---------------------------------------------------------------------------------------
'
Public Function Chomp(txt As String) As String
    If Right(txt, 2) = vbCrLf Then
        Chomp = Left(txt, Len(txt) - 2)
    ElseIf Right(txt, 1) = vbCr Or Right(txt, 1) = vbLf Then
        Chomp = Left(txt, Len(txt) - 1)
    Else
        Chomp = txt
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : RunningInIDE
' DateTime  : 2/24/2006 16:34
' Author    : briandonorfio
' Purpose   : Returns true if the app is running in the VB IDE, false if it's a compiled
'             version.
'---------------------------------------------------------------------------------------
'
Public Function RunningInIDE() As Boolean
    On Error GoTo errh
    Debug.Print 1 / 0
    RunningInIDE = False
    Exit Function
errh:
    Err.Clear
    RunningInIDE = True
End Function

'---------------------------------------------------------------------------------------
' Procedure : RunningThroughVPN
' DateTime  : 10/1/2008 11:14
' Author    : briandonorfio
' Purpose   : Returns true if we're running through a VPN or other slow connection. This
'             is really just a guess based on the ip address. Anything 10.0.0.* is
'             considered close network-wise.
'---------------------------------------------------------------------------------------
'
Public Function RunningThroughVPN() As Boolean
    Dim ipAddr As String
    ipAddr = GetCurrentIPAddress()
    If Left(ipAddr, 8) = "10.0.50." Then
        RunningThroughVPN = False
    Else
        RunningThroughVPN = True
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : ListAsCommaSep
' DateTime  : 4/3/2006 09:26
' Author    : briandonorfio
' Purpose   : Returns the contents of list as a comma separated string, suitable for an
'             sql in clause or something. Optional delimiter for each element of list,
'             optional argument escape for whether or not we should double up any single
'             quotes (for sql escaping).
'---------------------------------------------------------------------------------------
'
Public Function ListAsCommaSep(list() As String, Optional delim As String = "", Optional escape As Boolean = False, Optional firstColOnly As Boolean = False) As String
    Dim i As Long, retval As String, thisline As String
    retval = ""
    For i = LBound(list) To UBound(list)
        If firstColOnly Then
            thisline = Left(list(i), InStr(list(i), vbTab) - 1)
        Else
            thisline = list(i)
        End If
        If escape Then
            retval = retval & delim & EscapeSQuotes(thisline) & delim & ", "
        Else
            retval = retval & delim & thisline & delim & ", "
        End If
    Next i
    ListAsCommaSep = Left(retval, Len(retval) - 2)
End Function

Public Function ListBoxLineColumn(list As ListBox, Optional line As Long = -1, Optional column As Long = 0) As String
    If line = -1 Then
        line = list.ListIndex
    End If
    Dim cols As Variant
    cols = Split(list.list(line), vbTab)
    ListBoxLineColumn = cols(column)
End Function

'---------------------------------------------------------------------------------------
' Procedure : ListBoxAsArray
' DateTime  : 4/11/2006 17:01
' Author    : BrianDonorfio
' Purpose   : Converts the items in a listbox to an array of strings. Optional parameter
'             to only do the selected item(s). You should have probably already declared
'             the array with something like redim foo(me.list.listcount-1) as string
'---------------------------------------------------------------------------------------
'
Public Function ListBoxAsArray(list As ListBox, Optional selOnly As Boolean = False, Optional colNumber As Long = -1) As String()
    ReDim newarray(IIf(selOnly, list.SelCount - 1, list.ListCount - 1)) As String
    Dim i As Long, j As Long
    For i = 0 To list.ListCount - 1
        If Not selOnly Or list.Selected(i) Then
            If colNumber = -1 Then
                newarray(j) = list.list(i)
            Else
                Dim temp As Variant
                temp = Split(list.list(i), vbTab)
                newarray(j) = temp(colNumber)
                'newarray(j) = Left(list.list(i), InStr(list.list(i), vbTab) - 1)
            End If
            j = j + 1
        End If
    Next i
    ListBoxAsArray = newarray
End Function

'---------------------------------------------------------------------------------------
' Procedure : ListBoxEnsureSelectionIsVisible
' Author    : briandonorfio
' Date      : 11/23/2011
' Purpose   : Automatically scroll so selected line is display. Only works with standard
'             font size (8pt)
'---------------------------------------------------------------------------------------
'
Public Sub ListBoxEnsureSelectionIsVisible(lbx As ListBox)
    If lbx.ListIndex = -1 Then
        Exit Sub
    End If
    If lbx.FontSize <> 8.25 Then
        Exit Sub
    End If
    
    Dim count As Long
    count = (lbx.Height - 60) / 195
    
    If lbx.ListIndex < lbx.TopIndex Then
        'scroll up
        lbx.TopIndex = lbx.ListIndex
    ElseIf lbx.ListIndex <= lbx.TopIndex + count - 1 Then
        'already visible
    Else
        'scroll down
        lbx.TopIndex = lbx.ListIndex - count + 1
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RSFieldsAsArray
' DateTime  : 12/6/2006 09:35
' Author    : BrianDonorfio
' Purpose   : Converts the contents of a recordset's fields collection to an array of
'             strings.
'---------------------------------------------------------------------------------------
'
Public Function RSFieldsAsArray(fieldColl As ADODB.fields) As String()
    Dim i As Long
    ReDim retval(fieldColl.count - 1) As String
    For i = 0 To fieldColl.count - 1
        retval(i) = CStr(Nz(fieldColl(i).value))
    Next i
    RSFieldsAsArray = retval
End Function

Public Function RSFieldsAsHash(fieldColl As ADODB.fields) As Dictionary
    Dim i As Long
    Dim retval As Dictionary
    Set retval = New Dictionary
    For i = 0 To fieldColl.count - 1
        retval.Add CStr(fieldColl(i).name), CStr(Nz(fieldColl(i).value))
    Next i
    Set RSFieldsAsHash = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : RSToArrayOfHashes
' DateTime  : 11/26/2008 14:10
' Author    : briandonorfio
' Purpose   : Converts a recordset to an array of Dictionary objects, similar to perl's
'             fetchrow_arrayref({})
'---------------------------------------------------------------------------------------
'
Public Function RSToArrayOfHashes(rst As ADODB.Recordset) As Variant
    Dim i As Long
    Dim retval As Variant
    If rst.RecordCount > 0 Then
        Dim initialPosition As Long
        ReDim retval(rst.RecordCount - 1) As Variant
        initialPosition = rst.AbsolutePosition
        rst.MoveFirst
        While Not rst.EOF
            Dim thisdict As Dictionary
            Set thisdict = New Dictionary
            Dim j As Long
            For j = 0 To rst.fields.count - 1
                thisdict.Add rst(j).name, Nz(rst(j).value)
            Next j
            Set retval(i) = thisdict
            rst.MoveNext
            i = i + 1
        Wend
        If rst.AbsolutePosition <> initialPosition Then
            rst.AbsolutePosition = initialPosition
        End If
    Else
        retval = Split("", "zerolengtharray")
    End If
    RSToArrayOfHashes = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : FormatForYahooURL
' DateTime  : 1/26/2009 15:05
' Author    : briandonorfio
' Purpose   : Converts a name into a url-able chunk.
'---------------------------------------------------------------------------------------
'
Public Function FormatForYahooURL(url As String) As String
    url = LCase(url)
    'Dim re As RegExp
    'Set re = New RegExp
    're.Pattern = "[^a-z0-9]"
    're.Global = True
    'FormatForYahooURL = re.Replace(url, "-")
    Dim i As Long
    For i = 1 To Len(url)
        If IsAlphaNumeric(Mid(url, i, 1)) Then
            'ok
        ElseIf Mid(url, i, 1) = "-" Then
            'ok
        Else
            url = Replace(url, Mid(url, i, 1), "-")
        End If
    Next i
    While Left(url, 1) = "-"
        url = Mid(url, 2)
    Wend
    While Right(url, 1) = "-"
        url = Left(url, Len(url) - 1)
    Wend
    While CBool(InStr(url, "--"))
        url = Replace(url, "--", "-")
    Wend
    FormatForYahooURL = url
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetControlValue
' DateTime  : 4/14/2006 13:11
' Author    : BrianDonorfio
' Purpose   : Given a form and control (and possibly index), this will set the default
'             method for that control to val. .Text for text/comboboxes, etc.
'---------------------------------------------------------------------------------------
'
Public Function SetControlValue(formName As String, controlName As String, val As String, Optional Index As Integer = -1) As Boolean
    Dim frm As Form, ctl As Variant
    Set frm = GetFormByName(formName)
    If frm Is Nothing Then
        SetControlValue = False
    Else
        For Each ctl In frm.Controls
            If ctl.name = controlName Then
                If Index <> -1 Then
                    If ctl.Index <> Index Then
                        GoTo nextctl
                    End If
                End If
                Select Case True
                    Case Is = TypeOf ctl Is TextBox, TypeOf ctl Is ComboBox
                        ctl.Text = val
                        SetControlValue = True
                    Case Is = TypeOf ctl Is CheckBox, TypeOf ctl Is OptionButton
                        ctl.value = SQLBool(CBool(val))
                    Case Is = TypeOf ctl Is Label, TypeOf ctl Is CommandButton
                        ctl.Caption = val
                        SetControlValue = True
                    Case Else
                        SetControlValue = False
                End Select
                Exit Function
nextctl:
            End If
        Next ctl
    End If
    SetControlValue = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : OptionListSelectedIndex
' DateTime  : 5/22/2006 16:11
' Author    : briandonorfio
' Purpose   : Returns the index of the selected item from an array (collection?) of
'             option controls, or -1 if none are selected
'---------------------------------------------------------------------------------------
'
Public Function OptionListSelectedIndex(optGroup As Object) As Integer
    Dim this As Variant
    OptionListSelectedIndex = -1
    For Each this In optGroup
        If this.value Then
            OptionListSelectedIndex = this.Index
            Exit Function
        End If
    Next this
End Function

'---------------------------------------------------------------------------------------
' Procedure : TrimWhitespace
' DateTime  : 5/30/2006 15:23
' Author    : briandonorfio
' Purpose   : Removes excess whitespace from a string. Optionally can only do one side
'             or the other, defaults to both. Optionally can remove whitespace
'             between words in a string, defaults to off. Optionally, can consider line
'             endings as whitespace to be trimmed/chopped (ends only,not in middle),
'             defaults to off.
'---------------------------------------------------------------------------------------
'
Public Function TrimWhitespace(txt As String, Optional doLeading As Boolean = True, Optional doTrailing As Boolean = True, Optional doInside As Boolean = False, Optional doCrLfs As Boolean = False) As String
    Dim retval As String
    retval = txt
    If doLeading Then
        While Left(retval, 1) = " " Or (doCrLfs And (Left(retval, 1) = vbCr Or Left(retval, 1) = vbLf))
            retval = Mid(retval, 2)
        Wend
    End If
    If doTrailing Then
        While Right(retval, 1) = " " Or (doCrLfs And (Right(retval, 1) = vbCr Or Right(retval, 1) = vbLf))
            retval = Left(retval, Len(retval) - 1)
        Wend
    End If
    If doInside Then
        If doCrLfs Then
            While CBool(InStr(retval, vbCrLf & vbCrLf))
                retval = Replace(retval, vbCrLf & vbCrLf, vbCrLf)
            Wend
        End If
        
        While CBool(InStr(retval, "  "))
            retval = Replace(retval, "  ", " ")
        Wend
        While CBool(InStr(retval, vbCrLf & " "))
            retval = Replace(retval, vbCrLf & " ", vbCrLf)
        Wend
        While CBool(InStr(retval, " " & vbCrLf))
            retval = Replace(retval, " " & vbCrLf, vbCrLf)
        Wend
        
        'Dim txt2 As String
        'txt2 = retval
        'retval = Left(txt2, 1)
        'Dim i As Long
        'For i = 2 To Len(txt2)
        '    If Mid(txt2, i, 1) <> " " Then
        '        retval = retval & Mid(txt2, i, 1)
        '    Else
        '        If Right(retval, 1) <> " " Then
        '            retval = retval & Mid(txt2, i, 1)
        '        End If
        '    End If
        'Next i
    End If
    TrimWhitespace = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : RemoveSelectedFrom
' DateTime  : 5/31/2006 12:17
' Author    : briandonorfio
' Purpose   : Removes all selected items from a multiselect listbox.
'---------------------------------------------------------------------------------------
'
Public Sub RemoveSelectedFrom(list As ListBox)
    Dim i As Long
    For i = list.ListCount - 1 To 0 Step -1
        If list.Selected(i) Then
            list.RemoveItem i
        End If
    Next i
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsPlural
' DateTime  : 6/5/2006 15:18
' Author    : briandonorfio
' Purpose   : Dumb plural function, but it works with the IsPlural function
'---------------------------------------------------------------------------------------
'
Public Function IsPlural(word As String) As Boolean
    IsPlural = CBool(Right(word, 1) = "s") And Not CBool(Right(word, 2) = "ss")
End Function

'---------------------------------------------------------------------------------------
' Procedure : Pluralize
' DateTime  : 6/1/2006 11:43
' Author    : briandonorfio
' Purpose   : Attempts to pluralize the last word in the given string. Maybe it even
'             works sometimes. See the perl version in generate_overture_upload.pl, it's
'             a lot more readable.
'---------------------------------------------------------------------------------------
'
Public Function Pluralize(txt As String, Optional ignoreCaps As Boolean = False) As String
    Dim lastword As String
    lastword = IIf(InStrRev(txt, " "), Mid(txt, InStrRev(txt, " ") + 1), txt)
    If Not ignoreCaps And IsProperNoun(txt) Then
        'do nothing
    ElseIf IsPlural(txt) Then
        'do nothing
    ElseIf Right(lastword, 2) = "ss" Then
        'changes to -es
        lastword = lastword & "es"
    ElseIf Right(lastword, 1) = "y" Then
        If IsVowel(Mid(lastword, Len(lastword) - 1, 1)) Then
            'add -s
            lastword = lastword & "s"
        Else
            'changes to -ies
            lastword = Left(lastword, Len(lastword) - 1) & "ies"
        End If
    ElseIf Right(lastword, 1) = "x" Then
        'add -es
        lastword = lastword & "es"
    ElseIf Right(lastword, 2) = "ch" Or Right(lastword, 2) = "sh" Then
        'add -es
        lastword = lastword & "es"
    ElseIf IsVowel(Mid(lastword, Len(lastword) - 1, 1)) And Right(lastword, 1) = "f" Then
        'changes to -ves
        lastword = Left(lastword, Len(lastword) - 1) & "ves"
    ElseIf IsVowel(Mid(lastword, Len(lastword) - 2, 1)) And Right(lastword, 2) = "fe" Then
        'changes to -ves
        lastword = Left(lastword, Len(lastword) - 2) & "ves"
    Else
        lastword = lastword & "s"
    End If
    Pluralize = IIf(InStrRev(txt, " "), Left(txt, InStrRev(txt, " ")), "") & lastword
End Function

'---------------------------------------------------------------------------------------
' Procedure : DePluralize
' DateTime  : 6/8/2006 12:37
' Author    : BrianDonorfio
' Purpose   : Reverse of the above function.
'---------------------------------------------------------------------------------------
'
Public Function DePluralize(txt As String, Optional ignoreCaps As Boolean = False) As String
    Dim lastword As String, retval  As String
    lastword = IIf(InStrRev(txt, " "), Mid(txt, InStrRev(txt, " ") + 1), txt)
    If IsProperNoun(lastword) Then
        retval = txt
    ElseIf Not IsPlural(lastword) Then
        retval = txt
    Else
        If Right(lastword, 4) = "iers" Then
            'pliers
            retval = txt
        ElseIf Right(lastword, 3) = "ies" Then
            'batteries -> battery
            lastword = Left(lastword, Len(lastword) - 3) & "y"
        ElseIf Right(lastword, 3) = "ves" Then
            'knives -> knife
            lastword = Left(lastword, Len(lastword) - 3) & "f"
            If IsVowel(Mid(lastword, Len(lastword) - 1, 1)) Then
                lastword = lastword & "e"
            End If
        ElseIf Right(lastword, 2) = "es" Then
            If IsVowel(Mid(lastword, Len(lastword) - 3, 1)) Then
                'hoses
                lastword = Left(lastword, Len(lastword) - 1)
            ElseIf Not IsVowel(Mid(lastword, Len(lastword) - 2, 1)) Then
                'lathes
                lastword = Left(lastword, Len(lastword) - 1)
            Else
                lastword = Left(lastword, Len(lastword) - 2)
            End If
        Else
            lastword = Left(lastword, Len(lastword) - 1)
        End If
        If retval = "" Then
            retval = Left(txt, InStrRev(txt, " ")) & lastword
        End If
    End If
    DePluralize = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsProperNoun
' DateTime  : 6/5/2006 15:19
' Author    : briandonorfio
' Purpose   : if capitalized or has numbers/special chars in it, then proper noun
'---------------------------------------------------------------------------------------
'
Public Function IsProperNoun(word As String) As Boolean
    If CBool(UCase(Left(word, 1)) = Left(word, 1)) Then
        IsProperNoun = True
    Else
        'Dim re As RegExp
        'Set re = New RegExp
        're.Pattern = "^[!A-Z]$"
        're.IgnoreCase = True
        'If re.Test(word) Then
        '    IsProperNoun = True
        'Else
        '    IsProperNoun = False
        'End If
        Dim i As Long, char As Long
        For i = 1 To Len(word)
            char = Asc(LCase(Mid(word, i, 1)))
            If char >= 97 And char <= 122 Then
                'letters, ok
            Else
                IsProperNoun = True
                Exit Function
            End If
        Next i
        IsProperNoun = False
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsVowel
' DateTime  : 6/1/2006 11:44
' Author    : briandonorfio
' Purpose   : Returns true if the given string is a single vowel, false otherwise. "Y"
'             is not considered a vowel. Maybe it should be.
'---------------------------------------------------------------------------------------
'
Public Function IsVowel(letter As String) As Boolean
    Select Case LCase(letter)
        Case Is = "a", "e", "i", "o", "u"
            IsVowel = True
        Case Else
            IsVowel = False
    End Select
End Function

'---------------------------------------------------------------------------------------
' Procedure : InchMarksToAbbrev
' DateTime  : 6/1/2006 12:27
' Author    : briandonorfio
' Purpose   : Changes inch marks (defined as a double quote or two single quotes coming
'             after a numeric character) in the given string to "inch"
'---------------------------------------------------------------------------------------
'
Public Function InchMarksToAbbrev(txt As String) As String
    Dim i As Long, retval As String
    i = 2
    retval = Left(txt, 1)
    While i < Len(txt) + 1
        If IsNumeric(Mid(txt, i - 1, 1)) And Mid(txt, i, 1) = """" Then
            retval = retval & "inch"
        ElseIf IsNumeric(Mid(txt, i - 1, 1)) And Mid(txt, i, 1) = "'" And Mid(txt, i + 1, 1) = "'" Then
            retval = retval & "inch"
            i = i + 1
        Else
            retval = retval & Mid(txt, i, 1)
        End If
        i = i + 1
    Wend
    InchMarksToAbbrev = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : FootMarksToAbbrev
' DateTime  : 6/1/2006 12:34
' Author    : briandonorfio
' Purpose   : Same as InchMarksToAbbrev(), just changes a single quote to "ft". You
'             should definitely run the inch one first, otherwise you'll end up with
'             descriptions like "10ft' miter saw"
'---------------------------------------------------------------------------------------
'
Public Function FootMarksToAbbrev(txt As String) As String
    Dim i As Long, retval As String
    i = 2
    retval = Left(txt, 1)
    While i < Len(txt) + 1
        If IsNumeric(Mid(txt, i - 1, 1)) And Mid(txt, i, 1) = "'" Then
            retval = retval & "ft"
        Else
            retval = retval & Mid(txt, i, 1)
        End If
        i = i + 1
    Wend
    FootMarksToAbbrev = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : tr
' DateTime  : 6/5/2006 13:26
' Author    : briandonorfio
' Purpose   : Similar to perl's tr/// function, transliterates the searchstr set of
'             characters in a to the replacestr set in thestring. doesn't handle ranges
'             though, you'll have to spell it out.
'---------------------------------------------------------------------------------------
'
Public Function tr(thestring As String, searchstr As String, replacestr As String) As String
    If Len(searchstr) <> Len(replacestr) Then
        Err.Raise ErrorCodes.TransliterateInvalidArguments, "Utilities", "tr called w/ invalid search+replace lengths"
    End If
    
    Dim i As Long, retval As String, pos As String
    For i = 1 To Len(thestring)
        pos = InStr(searchstr, Mid(thestring, i, 1))
        If pos = 0 Then
            retval = retval & Mid(thestring, i, 1)
        Else
            retval = retval & Mid(replacestr, pos, 1)
        End If
    Next i
    tr = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : Download
' DateTime  : 6/12/2006 12:00
' Author    : briandonorfio
' Purpose   : Downloads a file from url and saves it in file, overwriting if necessary.
'             No progress indication, locks up until it finishes.
'---------------------------------------------------------------------------------------
'
Public Function Download(url As String, file As String) As Boolean
    Dim retval As Long
    retval = URLDownloadToFile(0, url, file, 0, 0)
    Download = CBool(retval = 0) 'S_OK
End Function

'---------------------------------------------------------------------------------------
' Procedure : SendEmailTo
' DateTime  : 2/7/2008 10:34
' Author    : briandonorfio
' Purpose   : Sends a text email to a single recipient. Nothing displayed at all, so
'             this might be good for diagnostic emails, etc. Any errors in the script
'             parameters are handled as error messages to the admin address specified in
'             the script.
'---------------------------------------------------------------------------------------
'
Public Function SendEmailTo(Address As String, subject As String, body As String)
On Error GoTo errh
    'Dim se As vbSendMail.clsSendMail
    'Set se = New vbSendMail.clsSendMail
    Dim se As Object
    Set se = CreateObject("vbSendMail.clsSendMail")
    se.SMTPHost = "toolsplus06"
    se.From = "no-reply@tools-plus.com"
    se.Recipient = Address
    se.subject = subject
    se.Message = body
    se.Send
    Set se = Nothing
    SendEmailTo = True
    Exit Function
    
errh:
    MsgBox "Unable to send email: " & Err.Description
    Err.Clear
    SendEmailTo = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetCurrentIPAddress
' DateTime  : 3/25/2008 12:50
' Author    : briandonorfio
' Purpose   : Returns the first IP address in the win32 address mapping table
'---------------------------------------------------------------------------------------
'
Public Function GetCurrentIPAddress() As String
    Dim buf(511) As Byte
    Dim bufsize As Long
    bufsize = 512
    Dim rc As Long
    rc = GetIpAddrTable(buf(0), bufsize, 1)
    If rc = 0 Then
        If buf(0) = 0 Then
            GetCurrentIPAddress = "0.0.0.0"
        Else
            GetCurrentIPAddress = buf(4) & "." & buf(5) & "." & buf(6) & "." & buf(7)
        End If
    Else
        GetCurrentIPAddress = "0.0.0.0"
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetTerminalServicesClientIPAddress
' DateTime  : 3/25/2008 12:50
' Author    : briandonorfio
' Purpose   : Returns the first IP address in the terminal services IP address table.
'---------------------------------------------------------------------------------------
'
Public Function GetTerminalServicesClientIPAddress() As String
    Dim addr As WTS_CLIENT_ADDRESS
    Dim buf As Long
    Dim lenBytes As Long
    
    Dim qsi As Long
    qsi = WTSQuerySessionInformationA(WTS_CURRENT_SERVER_HANDLE, _
                                      WTS_CURRENT_SESSION, _
                                      WTSClientAddress, _
                                      buf, _
                                      lenBytes _
                                     )
    RtlMoveMemory addr, buf, lenBytes
    If qsi = 0 Then
        GetTerminalServicesClientIPAddress = ""
    Else
        GetTerminalServicesClientIPAddress = addr.addr(2) & "." & addr.addr(3) & "." & addr.addr(4) & "." & addr.addr(5)
    End If
End Function

Public Function KeyCodeIsAlpha(KeyAscii As Integer) As Boolean
    If (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Then
        KeyCodeIsAlpha = True
    Else
        KeyCodeIsAlpha = False
    End If
End Function

Public Function KeyCodeIsNumeric(KeyAscii As Integer) As Boolean
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
        KeyCodeIsNumeric = True
    Else
        KeyCodeIsNumeric = False
    End If
End Function

Public Function Similarity(one As String, two As String, Optional compareMethod As VbCompareMethod = vbTextCompare) As Single
    Dim i As Long, j As Long
    Dim pos As Long
    pos = 1
    Dim matchCount As Long
    For i = 1 To Len(two)
        For j = pos To Len(one)
            If StrComp(Mid(one, j, 1), Mid(two, i, 1), compareMethod) = 0 Then
                matchCount = matchCount + 1
                pos = j + 1
                Exit For
            End If
        Next j
    Next i
    Similarity = (matchCount / IIf(Len(one) > Len(two), Len(one), Len(two))) * 100
End Function

Public Function IsAlphaNumeric(char As String) As Boolean
    If Asc(char) >= Asc("a") And Asc(char) <= Asc("z") Then
        IsAlphaNumeric = True
    ElseIf Asc(char) >= Asc("A") And Asc(char) <= Asc("Z") Then
        IsAlphaNumeric = True
    ElseIf Asc(char) >= Asc("0") And Asc(char) <= Asc("9") Then
        IsAlphaNumeric = True
    Else
        IsAlphaNumeric = False
    End If
End Function

Public Function IsIntegeric(num As String) As Boolean
    Dim i As Long
    Dim retval As Boolean
    retval = True
    For i = 1 To Len(num)
        If IsDigit(Mid(num, i, 1)) = False Then
            retval = False
            Exit For
        End If
    Next i
    IsIntegeric = retval
End Function

Public Function IsDigit(num As String) As Boolean
    IsDigit = False
    If Len(num) = 1 Then
        If Asc("0") <= Asc(num) And Asc(num) <= Asc("9") Then
            IsDigit = True
        End If
    End If
End Function

Public Function EndsWithHRule(ByVal txt As String) As Boolean
    txt = TrimWhitespace(txt, False, True, False, True)
    If LCase(Right(txt, 4)) = "<hr>" Or LCase(Right(txt, 5)) = "<hr/>" Or LCase(Right(txt, 6)) = "<hr />" Then
        EndsWithHRule = True
    Else
        EndsWithHRule = False
    End If
End Function

Public Function FullStateName(abbrev As String) As String
    Select Case LCase(abbrev)
        Case Is = "al": FullStateName = "Alabama"
        Case Is = "ak": FullStateName = "Alaska"
        Case Is = "az": FullStateName = "Arizona"
        Case Is = "ar": FullStateName = "Arkansas"
        Case Is = "ca": FullStateName = "California"
        Case Is = "co": FullStateName = "Colorado"
        Case Is = "ct": FullStateName = "Connecticut"
        Case Is = "de": FullStateName = "Delaware"
        Case Is = "dc": FullStateName = "DC"
        Case Is = "fl": FullStateName = "Florida"
        Case Is = "ga": FullStateName = "Georgia"
        Case Is = "hi": FullStateName = "Hawaii"
        Case Is = "id": FullStateName = "Idaho"
        Case Is = "il": FullStateName = "Illinois"
        Case Is = "in": FullStateName = "Indiana"
        Case Is = "ia": FullStateName = "Iowa"
        Case Is = "ks": FullStateName = "Kansas"
        Case Is = "ky": FullStateName = "Kentucky"
        Case Is = "la": FullStateName = "Louisiana"
        Case Is = "me": FullStateName = "Maine"
        Case Is = "md": FullStateName = "Maryland"
        Case Is = "ma": FullStateName = "Massachusetts"
        Case Is = "mi": FullStateName = "Michigan"
        Case Is = "mn": FullStateName = "Minnesota"
        Case Is = "ms": FullStateName = "Mississippi"
        Case Is = "mo": FullStateName = "Missouri"
        Case Is = "mt": FullStateName = "Montana"
        Case Is = "ne": FullStateName = "Nebraska"
        Case Is = "nv": FullStateName = "Nevada"
        Case Is = "nh": FullStateName = "New Hampshire"
        Case Is = "nj": FullStateName = "New Jersey"
        Case Is = "nm": FullStateName = "New Mexico"
        Case Is = "ny": FullStateName = "New York"
        Case Is = "nc": FullStateName = "North Carolina"
        Case Is = "nd": FullStateName = "North Dakota"
        Case Is = "oh": FullStateName = "Ohio"
        Case Is = "ok": FullStateName = "Oklahoma"
        Case Is = "or": FullStateName = "Oregon"
        Case Is = "pa": FullStateName = "Pennsylvania"
        Case Is = "ri": FullStateName = "Rhode Island"
        Case Is = "sc": FullStateName = "South Carolina"
        Case Is = "sd": FullStateName = "South Dakota"
        Case Is = "tn": FullStateName = "Tennessee"
        Case Is = "tx": FullStateName = "Texas"
        Case Is = "ut": FullStateName = "Utah"
        Case Is = "vt": FullStateName = "Vermont"
        Case Is = "va": FullStateName = "Virginia"
        Case Is = "wa": FullStateName = "Washington"
        Case Is = "wv": FullStateName = "West Virginia"
        Case Is = "wi": FullStateName = "Wisconsin"
        Case Is = "wy": FullStateName = "Wyoming"
        Case Else: FullStateName = ""
    End Select
End Function

Public Function IsFormMinimized(formName As String) As Boolean
    Dim frm As Form
    Set frm = GetFormByName(formName)
    If frm Is Nothing Then
        IsFormMinimized = False
    Else
        IsFormMinimized = CBool(IsIconic(frm.hWnd))
    End If
End Function

Public Function GetFormByName(formName As String) As Form
    Dim frm As Form
    For Each frm In Forms
        If frm.name = formName Then
            Set GetFormByName = frm
            Exit Function
        End If
    Next frm
    Set GetFormByName = Nothing
End Function

Public Function UnMinimizeForm(formName As String) As Boolean
    Dim frm As Form
    Set frm = GetFormByName(formName)
    If frm Is Nothing Then
        UnMinimizeForm = False
    Else
        frm.WindowState = vbNormal
        UnMinimizeForm = True
    End If
End Function

Public Function FocusOnForm(formName As String) As Boolean
    Dim frm As Form
    Set frm = GetFormByName(formName)
    If frm Is Nothing Then
        FocusOnForm = False
    Else
        If IsFormMinimized(formName) Then
            UnMinimizeForm formName
        End If
        frm.SetFocus
        FocusOnForm = True
    End If
End Function

'Public Function LoadAndShowForm(formName As String) As Boolean
'    Dim frm As Form
'    Set frm = Forms.Add(formName)
'    frm.Show
'End Function

Public Function URLEncode(strToEncode As String) As String
    Dim i As Long, retval As String, char As String
    For i = 1 To Len(strToEncode)
        char = Mid(strToEncode, i, 1)
        If IsAlphaNumeric(char) Or char = "-" Or char = "_" Or _
           char = "." Or char = "!" Or char = "~" Or char = "*" Or _
           char = "'" Or char = "(" Or char = ")" Then
            'ok
            retval = retval & char
        Else
            'encode
            retval = retval & "%" & Hex(Asc(char))
        End If
    Next i
    URLEncode = retval
End Function

Public Function URLDecode(strToDecode As String) As String
    Dim i As Long, retval As String, char As String
    For i = 1 To Len(strToDecode)
        char = Mid(strToDecode, i, 1)
        If char <> "%" Then
            'ok
            retval = retval & char
        Else
            'decode
            If Len(Mid(strToDecode, i + 1, 2)) <> 2 Then
                'bad escape combination, ignore
                i = i + Len(Mid(strToDecode, i + 1, 2))
            Else
                retval = retval & Chr("&H" & Mid(strToDecode, i + 1, 2))
                i = i + 2
            End If
        End If
    Next i
    URLDecode = retval
End Function

Public Sub EnumerateEnvironmentVariables()
    Dim i As Long, thisenv As String
    i = 1
    Do
        thisenv = Environ(i)
        If thisenv <> "" Then
            Debug.Print thisenv
        Else
            Exit Do
        End If
        i = i + 1
    Loop
End Sub

Public Function OptionGroupHasSelection(optGroup As Object) As Boolean
    Dim retval As Boolean
    Dim i As Long
    For i = 0 To optGroup.UBound
        If optGroup(i).value Then
            retval = True
            Exit For
        End If
    Next i
    OptionGroupHasSelection = retval
End Function

Public Function ProcessTextBoxPaste(ctl As TextBox) As String
    Dim txt As String, selStart As Long, selEnd As Long
    txt = ctl.Text
    selStart = ctl.selStart
    selEnd = ctl.selStart + ctl.SelLength
    
    Dim newText As String
    newText = Mid(txt, 1, selStart) & ClipBoard_GetData() & Mid(txt, selEnd + 1)
    
    ProcessTextBoxPaste = newText
End Function

Public Function IndentTextParagraph(txt As String, Optional indentCount As Long = 4, Optional charsPerLine As Long = 70)
    'Dim Lines As Variant
    'Lines = Split(txt, vbCrLf)
    'Dim i As Long
    'For i = 0 To UBound(Lines)
    '    Lines(i) = indent & Lines(i)
    'Next i
    'IndentTextParagraph = Join(Lines, vbCrLf)
    Dim retval As PerlArray
    Set retval = New PerlArray
    
    Dim linesArray As Variant
    linesArray = Split(TrimWhitespace(txt, True, True, True, True), vbCrLf)
    
    Dim i As Long
    For i = 0 To UBound(linesArray)
        If linesArray(i) <> "" Then
            If i <> 0 Then
                retval.Push Space(indentCount)
            End If
            
            Dim wordsArray As Variant
            wordsArray = Split(linesArray(i), " ")
            
            Dim currentLine As String
            currentLine = Space(indentCount) & wordsArray(0)
            
            Dim j As Long
            For j = 1 To UBound(wordsArray)
                If Len(currentLine & " " & wordsArray(j)) <= charsPerLine Then
                    currentLine = currentLine & " " & wordsArray(j)
                Else
                    retval.Push currentLine
                    currentLine = Space(indentCount) & wordsArray(j)
                End If
            Next j
            
            retval.Push currentLine
        End If
    Next i
    IndentTextParagraph = retval.Join(vbCrLf)
End Function

Public Function ColumnIndexToLetter(dx As Long) As String
    Dim First As Long
    First = dx \ 26
    Dim second As Long
    second = dx Mod 26
    If second = 0 Then
        second = 26
        First = First - 1
    End If
    ColumnIndexToLetter = IIf(First = 0, "", Chr(64 + First)) & IIf(second = 0, "Z", Chr(64 + second))
End Function

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

Public Sub SetListTabs(lb As ListBox, tabs As Variant, Optional additive As Boolean = True)
    ReDim realtabs(UBound(tabs)) As Long
    Dim i As Long
    For i = 0 To UBound(tabs)
        realtabs(i) = tabs(i)
        If additive = True And i > 0 Then
            realtabs(i) = realtabs(i) + realtabs(i - 1)
        End If
    Next i
    Call SetListTabStops(lb.hWnd, realtabs)
End Sub

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
                    .SelLength = Len(.list(.ListIndex))
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
            .SelLength = Len(.list(Index)) - Len(SearchText)
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
    ctl.SelLength = 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetListCtlIndex
' DateTime  : 6/14/2005 11:37
' Author    : briandonorfio
' Purpose   : Changes the ComboBox.Index value without calling the *_Click() event.
'---------------------------------------------------------------------------------------
'
Private Sub SetListCtlIndex(ctl As ComboBox, ByVal Index As Long)
    SendMessage ctl.hWnd, CB_SETCURSEL, Index, 0&
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ShellWait
' DateTime  : 6/14/2005 11:10
' Author    : briandonorfio
' Purpose   : Starts a synchronous shell. Usage similar to the normal shell command.
'---------------------------------------------------------------------------------------
'
Public Function ShellWait(PathName As String, Optional WindowStyle As VbAppWinStyle = vbNormalFocus) As Boolean
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim Ret As Long
    ' Initialize the STARTUPINFO structure:
    With start
        .cb = Len(start)
        'If Not IsMissing(WindowStyle) Then
            .dwFlags = STARTF_USESHOWWINDOW
            .wShowWindow = WindowStyle
        'End If
    End With
    ' Start the shelled application:
    Ret& = CreateProcessA(0&, PathName, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    ' Wait for the shelled application to finish:
    Ret& = WaitForSingleObject(proc.hProcess, INFINITE)
    Ret& = CloseHandle(proc.hProcess)
    ShellWait = True
End Function

'---------------------------------------------------------------------------------------
' Procedure : ShellCapture
' Author    : briandonorfio
' Date      : 1/19/2012
' Purpose   : Starts a synchronous shell, and returns STDOUT/STDERR
'---------------------------------------------------------------------------------------
'
Public Function ShellCapture(cmd As String) As String
    Dim piperead As Long, pipewrite As Long
    Dim sa As SECURITY_ATTRIBUTES
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1
    If 0 = CreatePipe(piperead, pipewrite, sa, 0) Then
        ShellCapture = ""
        Exit Function
    End If
    
    Dim si As STARTUPINFO
    si.cb = Len(si)
    si.dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
    si.wShowWindow = SW_HIDE
    si.hStdOutput = pipewrite
    si.hStdError = pipewrite
    
    Dim pi As PROCESS_INFORMATION
    
    If CreateProcessA(0&, cmd, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, si, pi) Then
        CloseHandle pipewrite
        CloseHandle pi.hThread
        pipewrite = 0
        
        Const bufsize As Long = 10240
        Dim out(10240) As Byte
        Dim bytesRead As Long
        Dim retval As String
        retval = ""
        Do
            If ReadFile(piperead, out(0), bufsize, bytesRead, ByVal 0&) = 0 Then
                Exit Do
            End If
            retval = retval & Left$(StrConv(out(), vbUnicode), bytesRead)
        Loop
        CloseHandle pi.hProcess
        
        ShellCapture = retval
    Else
        ShellCapture = ""
    End If
    
    CloseHandle piperead
    CloseHandle pipewrite
End Function

Public Function ClipBoard_SetImage(pic As Picture)
    Dim hClipMemory As Long, X As Long

    ' Open the Clipboard to copy data to.
    If OpenClipboard(0&) = 0 Then
        MsgBox "Could not open the Clipboard. Copy aborted."
        Exit Function
    End If

    ' Clear the Clipboard.
    X = EmptyClipboard()

    ' Copy the data to the Clipboard.
    hClipMemory = SetClipboardData(CF_BITMAP, pic.handle)

OutOfHere2:

    If CloseClipboard() = 0 Then
        MsgBox "Could not close Clipboard."
    End If
End Function

Public Function ClipBoard_SetData(MyString As String)
    Dim hGlobalMemory As Long, lpGlobalMemory As Long
    Dim hClipMemory As Long, X As Long

    ' Allocate moveable global memory.
    '-------------------------------------------
    hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

    ' Lock the block to get a far pointer
    ' to this memory.
    lpGlobalMemory = GlobalLock(hGlobalMemory)

    ' Copy the string to this global memory.
    lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

    ' Unlock the memory.
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        MsgBox "Could not unlock memory location. Copy aborted."
        GoTo OutOfHere2
    End If

    ' Open the Clipboard to copy data to.
    If OpenClipboard(0&) = 0 Then
        MsgBox "Could not open the Clipboard. Copy aborted."
        Exit Function
    End If

    ' Clear the Clipboard.
    X = EmptyClipboard()

    ' Copy the data to the Clipboard.
    hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:

    If CloseClipboard() = 0 Then
        MsgBox "Could not close Clipboard."
    End If

End Function

Public Function ClipBoard_GetData() As String
On Error GoTo errh
   Dim hClipMemory As Long
   Dim lpClipMemory As Long
   Dim MyString As String
   Dim retval As Long

   If OpenClipboard(0&) = 0 Then
      MsgBox "Cannot open Clipboard. Another app. may have it open"
      Exit Function
   End If

   ' Obtain the handle to the global memory
   ' block that is referencing the text.
   hClipMemory = GetClipboardData(CF_TEXT)
   If IsNull(hClipMemory) Then
      MsgBox "Could not allocate memory"
      GoTo OutOfHere
   End If

   ' Lock Clipboard memory so we can reference
   ' the actual data string.
   lpClipMemory = GlobalLock(hClipMemory)

   If Not IsNull(lpClipMemory) Then
      MyString = Space$(MAXSIZE)
      retval = lstrcpy(MyString, lpClipMemory)
      retval = GlobalUnlock(hClipMemory)

      ' Peel off the null terminating character.
      MyString = Mid(MyString, 1, InStr(1, MyString, Chr$(0), 0) - 1)
   Else
      MsgBox "Could not lock memory to copy string from."
   End If

OutOfHere:

   retval = CloseClipboard()
   ClipBoard_GetData = MyString
   Exit Function
   
errh:
    MsgBox "Error " & Err.Number & " accessing clipboard!"
    ClipBoard_GetData = ""
End Function

Public Sub ExpandDropDownToFit(ctl As ComboBox)
    SendMessage ctl.hWnd, CB_SETDROPPEDWIDTH, ComboBoxDropDownResize_GetDropDownWidth(ctl), ByVal 0
End Sub

Private Function ComboBoxDropDownResize_GetDropDownWidth(ctl As ComboBox) As Long
    Dim width As Long
    width = ComboBoxDropDownResize_GetLongestStringSize(ctl) 'longest string
    width = width + 2                                        '1px borders on each side
    width = width + GetSystemMetrics(SM_CXVSCROLL)           'plus scrollbar width
    width = width + 4                                        'fudge factor
    ComboBoxDropDownResize_GetDropDownWidth = width
End Function

Private Function ComboBoxDropDownResize_GetLongestStringSize(ctl As ComboBox) As Long
    Dim curlen As Long, maxlen As Long
    Dim i As Long
    For i = 0 To ctl.ListCount - 1
        curlen = ComboBoxDropDownResize_GetTextSize(ctl, i)
        If curlen > maxlen Then
            maxlen = curlen
        End If
    Next i
    ComboBoxDropDownResize_GetLongestStringSize = maxlen
End Function

Private Function ComboBoxDropDownResize_GetTextSize(ctl As ComboBox, i As Long) As Long
    Dim uSize As Size
    Dim hDC As Long, result As Long, Font As Long, oldfont As Long
    hDC = GetDC(ctl.hWnd)                                                       'get the control device context
    Font = SendMessage(ctl.hWnd, WM_GETFONT, 0, ByVal 0)                        'get the font the control is using
    oldfont = SelectObject(hDC, Font)                                           'set the device context to use that font
    result = GetTextExtentPoint32(hDC, ctl.list(i), Len(ctl.list(i)), uSize)    'get displayed length of string
    ComboBoxDropDownResize_GetTextSize = uSize.cx
    SelectObject hDC, oldfont                                                   'restore font for control device context
    ReleaseDC ctl.hWnd, hDC                                                     'close device context
End Function

Private Function ComboBoxDropDownResize_ConvertTwipsToPixels(twips As Long, axis As Long) As Long
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
   ComboBoxDropDownResize_ConvertTwipsToPixels = (twips / TwipsPerInch) * PixelsPerInch
End Function

Public Function Base64Encode(data As String, Optional doCrLfs As Boolean = True) As String
    Dim retval As String
    retval = ""
    Dim i As Long
    For i = 1 To Len(data) Step 3
        Dim group As Long
        group = &H10000 * Asc(Mid(data, i + 0, 1))
        If Len(data) >= i + 1 Then
            group = group + &H100 * Asc(Mid(data, i + 1, 1))
            If Len(data) >= i + 2 Then
                group = group + Asc(Mid(data, i + 2, 1))
            End If
        End If
        
        Dim octal As String
        octal = Oct(group)
        octal = String(8 - Len(octal), "0") & octal
        
        Dim part As String
        part = Mid(BASE64_CHARACTER_SET, CLng("&o" & Mid(octal, 1, 2)) + 1, 1) & _
               Mid(BASE64_CHARACTER_SET, CLng("&o" & Mid(octal, 3, 2)) + 1, 1) & _
               Mid(BASE64_CHARACTER_SET, CLng("&o" & Mid(octal, 5, 2)) + 1, 1) & _
               Mid(BASE64_CHARACTER_SET, CLng("&o" & Mid(octal, 7, 2)) + 1, 1)
        
        retval = retval & part
        
        If doCrLfs Then
            If (i + 2) Mod 57 = 0 Then
                retval = retval & vbCrLf
            End If
        End If
    Next i
    
    Select Case Len(data) Mod 3
        Case 1:
            retval = Left(retval, Len(retval) - 2) + "=="
        Case 2:
            retval = Left(retval, Len(retval) - 1) + "="
    End Select
    
    Base64Encode = retval
End Function

Public Function Base64Decode(base64 As String) As String
    base64 = Replace(base64, vbCrLf, "")
    If Len(base64) Mod 4 <> 0 Then
        Base64Decode = ""
        Exit Function
    End If
    
    Dim retval As String
    retval = ""
    Dim i As Long
    For i = 1 To Len(base64) Step 4
        Dim group As Long
        group = 0
        Dim numBytes As Long
        numBytes = 3
        Dim j As Long
        For j = 0 To 3 Step 1
            Dim thischar As String
            thischar = Mid(base64, i + j, 1)
            
            Dim thisData As String
            If thischar = "=" Then
                thisData = 0
                numBytes = numBytes - 1
            Else
                thisData = InStr(1, BASE64_CHARACTER_SET, thischar, vbBinaryCompare) - 1
            End If
            If thisData = -1 Then
                Base64Decode = ""
                Exit Function
            End If
            
            group = 64 * group + thisData
        Next j
        
        Dim hexGroup As String
        hexGroup = Hex(group)
        hexGroup = String(6 - Len(hexGroup), "0") & hexGroup
        
        Dim part As String
        part = Chr(CByte("&H" & Mid(hexGroup, 1, 2))) & _
               Chr(CByte("&H" & Mid(hexGroup, 3, 2))) & _
               Chr(CByte("&H" & Mid(hexGroup, 5, 2)))
        
        retval = retval & Left(part, numBytes)
    Next i
    
    Base64Decode = retval
End Function


Public Function CaptureActiveWindow() As Picture
    Dim hWnd As Long
    hWnd = GetForegroundWindow()
    
    Set CaptureActiveWindow = CaptureSpecifiedWindow(hWnd)
End Function

Public Function CaptureSpecifiedWindow(hWnd As Long) As Picture
    Dim r As Long, windowRect As RECT
    r = GetWindowRect(hWnd, windowRect)
    
    Set CaptureSpecifiedWindow = screenCapture(hWnd, windowRect.Right - windowRect.Left, windowRect.Bottom - windowRect.Top)
End Function

Public Function CaptureVB6Form(f As Form) As Picture
    Set CaptureVB6Form = screenCapture(f.hWnd, f.ScaleX(f.width, vbTwips, vbPixels), f.ScaleY(f.Height, vbTwips, vbPixels))
End Function

Private Function screenCapture(hWnd As Long, widthSrc As Long, heightSrc As Long) As Picture
    Dim r As Long
    
    Dim hDCSrc As Long
    hDCSrc = GetWindowDC(hWnd) ' Get device context for entire window.
    Dim hDCMemory As Long
    hDCMemory = CreateCompatibleDC(hDCSrc) ' Create a memory device context for the copy process
    Dim hBmp As Long
    hBmp = CreateCompatibleBitmap(hDCSrc, widthSrc, heightSrc) ' Create a bitmap and place it in the memory DC.
    Dim hBmpPrev As Long
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    
    Dim rasterCapsScrn As Long
    rasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster capabilities.
    Dim hasPaletteScrn As Long
    hasPaletteScrn = rasterCapsScrn And RC_PALETTE       ' Palette support.
    Dim paletteSizeScrn As Long
    paletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of palette.
    
    Dim hPal As Long, hPalPrev As Long, logPal As LOGPALETTE
    If hasPaletteScrn And paletteSizeScrn = 256 Then
        ' If the screen has a palette make a copy and realize it.
        ' Create a copy of the system palette.
        logPal.palVersion = &H300
        logPal.palNumEntries = 256
        r = GetSystemPaletteEntries(hDCSrc, 0, 256, logPal.palPalEntry(0))
        hPal = CreatePalette(logPal)
        ' Select the new palette into the memory DC and realize it.
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        r = RealizePalette(hDCMemory)
    End If
    
    ' Copy the on-screen image into the memory DC.
    r = BitBlt(hDCMemory, 0, 0, widthSrc, heightSrc, hDCSrc, 0, 0, vbSrcCopy)
    
    ' Remove the new copy of the  on-screen image.
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    
    ' If the screen has a palette get back the palette that was selected in previously.
    If hasPaletteScrn And (paletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    
    ' Release the device context resources back to the system.
    r = DeleteDC(hDCMemory)
    r = ReleaseDC(hWnd, hDCSrc)
    
    ' Call CreateBitmapPicture to create a picture object from the
    ' bitmap and palette handles. Then return the resulting picture
    ' object.
    Set screenCapture = createBitmapPicture(hBmp, hPal)
End Function

Private Function createBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim r As Long
    
    ' Fill in with IDispatch Interface ID.
    Dim IID_IDispatch As GUID
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    ' Fill Pic with necessary parts.
    Dim pic As PicBmp
    With pic
        .Size = Len(pic)          ' Length of structure.
        .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
        .hBmp = hBmp              ' Handle to bitmap.
        .hPal = hPal              ' Handle to palette (may be null).
    End With
    
    ' Create Picture object.
    Dim IPic As IPicture ' IPicture requires a reference to "Standard OLE Types."
    r = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
    
    ' Return the new Picture object.
    Set createBitmapPicture = IPic
End Function

Public Sub FlashWindow(hWnd As Long, Optional NumberOfFlashes As Long = 5)
    If Not APIFunctionPresent("FlashWindowEx", "user32") Then
        Exit Sub
    End If
    Dim retval As Boolean
    Dim udtFWInfo As FLASHWINFO
    With udtFWInfo
       .cbSize = 20
       .hWnd = hWnd
       .dwFlags = 2 'FLASHW_TRAY
       .uCount = NumberOfFlashes
       .dwTimeout = 0
    End With
    retval = FlashWindowEx(udtFWInfo)
End Sub

Private Function APIFunctionPresent(ByVal FunctionName As String, ByVal DllName As String) As Boolean
    Dim handle As Long, addr  As Long
    handle = LoadLibrary(DllName)
    If handle <> 0 Then
        addr = GetProcAddress(handle, FunctionName)
        FreeLibrary handle
    End If
    APIFunctionPresent = (addr <> 0)
End Function

Public Function escapeXMLEntities(ByVal txt As String) As String
    txt = Replace(txt, "&", "&amp;")
    txt = Replace(txt, """", "&quot;")
    txt = Replace(txt, "<", "&lt;")
    txt = Replace(txt, ">", "&gt;")
    escapeXMLEntities = txt
End Function

Public Function removeUpperASCII(str As String) As String
    Dim i As Long
    Dim retval As String
    For i = 1 To Len(str)
        If Asc(Mid(str, i, 1)) < 127 Then
            retval = retval & Mid(str, i, 1)
        End If
    Next i
    removeUpperASCII = retval
End Function

'---------------------------------------------------------------------------------------
' Procedure : StripHTML
' DateTime  : 7/5/2005 17:10
' Author    : briandonorfio
' Purpose   : Removes html tags from text. Optional tagsOnly argument will strip tags,
'             but leave html entities.
'---------------------------------------------------------------------------------------
'
Public Function StripHTML(ByVal txt As String, Optional tagsOnly As Boolean = False) As String
    Dim startPos As Long, endpos As Long, tag As String, taglen As Long
    txt = Replace(txt, vbCrLf, "")
    'first, scan for any comments
    startPos = InStr(txt, "<!--")
    endpos = InStr(txt, "-->")
    taglen = 3
    If endpos = 0 Then
        endpos = InStr(txt, "--!>")
        taglen = 4
    End If
    While startPos <> 0 And endpos <> 0
        tag = Mid(txt, startPos, endpos - startPos + taglen)
        txt = Replace(txt, tag, "")
        startPos = InStr(txt, "<!--")
        endpos = InStr(txt, "-->")
        taglen = 3
        If endpos = 0 Then
            endpos = InStr(txt, "--!>")
            taglen = 4
        End If
    Wend
    'now do the other tags
    startPos = InStr(txt, "<")
    endpos = InStr(txt, ">")
    While startPos <> 0 And endpos <> 0 And endpos > startPos
        tag = Mid(txt, startPos, endpos - startPos + 1)
        Select Case LCase(tag)
            Case Is = "<li>"
                txt = Replace(txt, tag, vbCrLf & "  " & Chr(149) & " ") '149 = bullet
            Case Is = "<p>"
                txt = Replace(txt, tag, vbCrLf & vbCrLf)
            Case Is = "<br>", "<br />"
                txt = Replace(txt, tag, vbCrLf)
            Case Is = "<h1>", "<h2>", "<h3>", "<h4>", "<h5>"
                txt = Replace(txt, tag, vbCrLf)
            Case Else
                txt = Replace(txt, tag, "")
        End Select
        startPos = InStr(txt, "<")
        endpos = InStr(txt, ">")
    Wend
    If Not tagsOnly Then
        txt = Replace(txt, "&nbsp;", " ")
        txt = Replace(txt, "&amp;", "&")
        txt = Replace(txt, "&quot;", """")
        txt = Replace(txt, "&lt;", "<")
        txt = Replace(txt, "&gt;", ">")
        txt = Replace(txt, "&deg;", Chr(186))   '186 = degree symbol
        txt = Replace(txt, "&plusmn;", "+/-")
        txt = Replace(txt, "&reg;", "")
        txt = Replace(txt, "&trade;", "")
    End If
    
    While InStr(txt, vbCrLf) = 1            'remove leading cr's
        txt = Replace(txt, vbCrLf, "", , 1)
    Wend
    StripHTML = txt
End Function

Public Function EscapeHTMLEntities(ByVal txt As String) As String
    txt = Replace(txt, "&", "&amp;")
    txt = Replace(txt, """", "&quot;")
    txt = Replace(txt, "<", "&lt;")
    txt = Replace(txt, ">", "&gt;")
    txt = Replace(txt, Chr(186), "&deg;")
    EscapeHTMLEntities = txt
End Function

'---------------------------------------------------------------------------------------
' Procedure : text2html
' DateTime  : 8/1/2006 15:33
' Author    : briandonorfio
' Purpose   : Changes text to a slightly HTMLish version of the text. Converts vbcrlf-
'             delimited text to paragraphs, 4+ dashes on a line to hrules, quoted
'             full paragraphs to blockquotes. Given a regex, it will replace all found
'             instances with an image tag. First capture should be the alignment for
'             the image, second capture should be the file number in BlogImages.
'---------------------------------------------------------------------------------------
'
Public Function Text2html(ByVal txt As String, Optional picRE As String = "") As String
    txt = Replace(txt, "&", "&amp;")
    txt = Replace(txt, "<", "&lt;")
    txt = Replace(txt, ">", "&gt;")
    txt = Replace(txt, """", "&quot;")
    While Left(txt, 2) = vbCrLf
        txt = Mid(txt, 3)
    Wend
    While Right(txt, 2) = vbCrLf
        txt = Left(txt, Len(txt) - 2)
    Wend
    Dim Lines As Variant
    Lines = Split(txt, vbCrLf)
    'Dim re As RegExp, m As MatchCollection
    'Set re = New RegExp
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    Dim i As Long, j As Long, html As String, InUL As Boolean
    For i = 0 To UBound(Lines)
        re.Pattern = "^\s*\*\s*(.+)\s*$"
        If re.Test(Lines(i)) Then
            Set m = re.Execute(Lines(i))
            If Not InUL Then
                html = html & "<ul>" & vbCrLf
                InUL = True
            End If
            html = html & "<li>" & m.item(0).SubMatches(0) & "</li>" & vbCrLf
        Else
            If InUL Then
                html = html & "</ul>" & vbCrLf
                InUL = False
            End If
            re.Pattern = "^\s*[-]{4,}\s*$"
            If re.Test(Lines(i)) Then
                html = html & "<hr />" & vbCrLf
            Else
                re.Pattern = "^\s*&quot;(.+)&quot;\s*$"
                If re.Test(Lines(i)) Then
                    Set m = re.Execute(Lines(i))
                    html = html & "<blockquote><p>" & m.item(0).SubMatches(0) & "</p></blockquote>" & vbCrLf
                Else
                    Lines(i) = Trim(Lines(i))
                    'If picRE <> "" Then
                    '    re.Pattern = picRE
                    '    If re.Test(Lines(i)) Then
                    '        Set m = re.Execute(Lines(i))
                    '        For j = 0 To m.count - 1
                    '            html = html & "<img src=""" & BLOG_PIC_DIR & m.item(j).SubMatches(1) & ".jpg"" float=""" & m.item(j).SubMatches(0) & """ />" & vbCrLf
                    '        Next j
                    '    End If
                    '    Lines(i) = re.Replace(Lines(i), "")
                    'End If
                    If Lines(i) <> "" Then
                        html = html & "<p>" & Lines(i) & "</p>" & vbCrLf
                    End If
                End If
            End If
        End If
    Next i
    If InUL Then
        html = html & "</ul>" & vbCrLf
    End If
    Text2html = html
End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidateHTML
' DateTime  : 9/30/2008 12:58
' Author    : briandonorfio
' Purpose   : Validates the given html, returns a variant array of errors, or a zero-
'             length array if valid. This needs to be a complete html page, including
'             doctype.
'---------------------------------------------------------------------------------------
'
Public Function ValidateHTML(html As String) As Variant
    Dim nsgmls As String, infile As String, options As String
    nsgmls = "s:\mastest\mas90-signs\utilities\nsgmls\bin\nsgmls.exe"
    infile = DESTINATION_DIR & "validate.html"
    options = "-E 100 " & _
              "-s " & _
              "-c " & qq("s:\mastest\mas90-signs\utilities\nsgmls\pubtext\html4.cat") & " " & _
              "-f " & qq(DESTINATION_DIR & "validateresults.txt")
    Open infile For Output As #1
    Print #1, html
    Close #1
    
    ShellWait nsgmls & " " & options & " " & infile, vbHide
    
    Dim out As String
    out = ""
    
    Dim buf As String
    Open DESTINATION_DIR & "validateresults.txt" For Input As #1
    While Not EOF(1)
        Line Input #1, buf
        If buf <> "" Then
            buf = Replace(buf, nsgmls & ":" & infile & ":", "", 1, 1, vbTextCompare)
            out = out & buf & vbCrLf
        End If
    Wend
    Close #1
    
    out = TrimWhitespace(out, True, True, False, True)
    ValidateHTML = Split(out, vbCrLf)
End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidateHTMLFragment
' DateTime  : 9/30/2008 12:58
' Author    : briandonorfio
' Purpose   : Same as previous function, except this will add a standard html body
'             around the given fragment. Doctype is html-401-loose, same as Yahoo.
'---------------------------------------------------------------------------------------
'
Public Function ValidateHTMLFragment(frag As String, Optional strict As Boolean = False) As Variant
    Dim html As String
    html = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01" & IIf(strict, "", " Transitional") & "//EN"">" & vbCrLf & _
           "<html>" & vbCrLf & _
           " <head><title>validate</title></head>" & _
           " <body>" & vbCrLf & _
              frag & vbCrLf & _
           " </body>" & vbCrLf & _
           "</html>"
    ValidateHTMLFragment = ValidateHTML(html)
End Function

Public Function ChangeRelativeURLsToAbsolute(html As String, base As String) As String
    If Left(base, 7) <> "http://" Then
        base = "http://" & base
    End If
    If Right(base, 1) <> "/" Then
        base = base & "/"
    End If
    
    'Dim re1 As VBScript_RegExp_55.RegExp, matches As VBScript_RegExp_55.MatchCollection
    'Set re1 = New VBScript_RegExp_55.RegExp
    Dim re1 As Object, matches As Object
    Set re1 = CreateObject("VBScript.RegExp")
    re1.Pattern = "\bhref=""([^""]+)"""
    re1.Global = True
    Set matches = re1.Execute(html)
    
    Dim re2 As Object
    Set re2 = CreateObject("VBScript.RegExp")
    re2.Pattern = "^[a-z]+://"
    
    If matches.count > 0 Then
        Dim thisMatch As Object
        For Each thisMatch In matches
            Dim url As String
            url = thisMatch.SubMatches(0)
            If re2.Test(url) Then
                'absolute
            ElseIf Left(url, 7) = "mailto:" Then
                'ignore
            Else
                html = Replace(html, """" & url & """", """" & base & IIf(Left(url, 1) = "/", Mid(url, 2), url) & """", 1, 1)
            End If
        Next thisMatch
    End If
    
    ChangeRelativeURLsToAbsolute = html
End Function

Public Function GetItemIDByItemCode(item As String) As Long
    Dim temp As String
    temp = DLookup("ItemID", "InventoryMaster", "ItemNumber='" & item & "'")
    If temp = "" Then
        GetItemIDByItemCode = -1
    Else
        GetItemIDByItemCode = CLng(temp)
    End If
End Function

