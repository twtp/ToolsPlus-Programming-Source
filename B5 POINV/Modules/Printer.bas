Attribute VB_Name = "Printer"
'---------------------------------------------------------------------------------------
' Module    : Printer
' DateTime  : 8/17/2005 13:06
' Author    : briandonorfio
' Purpose   : Functions for getting/setting printers.
'
'             Dependencies:
'               - none
'---------------------------------------------------------------------------------------

Option Explicit

Private Const HWND_BROADCAST As Long = &HFFFF&
Private Const WM_WININICHANGE As Long = &H1A

Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Private Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDataType As String
End Type

Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal level As Long, pDocInfo As DOCINFO) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

'---------------------------------------------------------------------------------------
' Procedure : getDefaultPrinter
' DateTime  : 8/17/2005 13:06
' Author    : briandonorfio
' Purpose   : Returns the name of the computer's default printer, or null string if no
'             default or no printer.
'---------------------------------------------------------------------------------------
'
Public Function getDefaultPrinter() As String
    Dim printer As String
    Dim buf As Long
    printer = String(255, Chr(0))
    buf = GetProfileString("Windows", "Device", "", printer, 255)
    If buf > 0 Then
        getDefaultPrinter = Left(printer, InStr(printer, ",") - 1)
    Else
        getDefaultPrinter = ""
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : setDefaultPrinter
' DateTime  : 8/17/2005 13:06
' Author    : briandonorfio
' Purpose   : Sets the default printer to the printer name given, probably determined
'             with one of the getters. This change is permanent, so if it's supposed
'             to be temporary, change it back after.
'---------------------------------------------------------------------------------------
'
Public Function setDefaultPrinter(printer As String) As Boolean
    Dim str As String, deviceline As String
    Dim buf As Long
    str = String(1024, " ")
    buf = GetProfileString("PrinterPorts", printer, "", str, 1024)
    If buf > 0 Then
        deviceline = printer & "," & Left(str, InStr(str, Chr(0)) - 1) & "," & Mid(str, InStr(str, Chr(0)) + 1)
        
        If WriteProfileString("windows", "Device", deviceline) And SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows") Then
            setDefaultPrinter = True
        Else
            setDefaultPrinter = False
        End If
    Else
        setDefaultPrinter = False
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : getPrinters
' DateTime  : 8/17/2005 13:06
' Author    : briandonorfio
' Purpose   : Returns an array of the printers available to the computer, or an array
'             with its only element a null string if no printers.
'---------------------------------------------------------------------------------------
'
Public Function getPrinters() As String()
    Dim printers As String
    Dim buf As Long
    printers = String(2048, " ")
    buf = GetProfileString("PrinterPorts", vbNullString, "", printers, 2048)
    If buf > 0 Then
        getPrinters = Split(Left(printers, InStr(printers, Chr(0) & Chr(0)) - 1), Chr(0))
    Else
        getPrinters = Array("")
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : PrintEPL2Data
' DateTime  : 3/25/2008 16:42
' Author    : briandonorfio
' Purpose   : Sends RAW print commands to the specified printer. This is probably for
'             EPL2 print codes for the Zebra thermal printers, but may eventually be for
'             some other functions, independent of the EPL2 name.
'
'             EPL2 print structure:
'
'               EPL2       - header
'               q456       - page width in dots, 203 DPI, so this is 2-1/4 inches
'               Q152,24+0  - page length in dots, spacer length+butterfly
'               S4         - speed
'               UN         - disable errors
'               WN         - disable windows mode
'               ZT         - print direction top
'               N          - start a page
'               ...page format goes here
'               P1         - print 1 copy of given page
'
'               ASCII string format example:
'                 A80,120,0,3,1,1,N,"text"
'                   - x offset
'                   - y offset
'                   - rotation
'                   - font type
'                   - horizontal stretch
'                   - vertical stretch
'                   - normal reverse
'                   - data to print
'
'               Barcode format example:
'                 B80,120,0,1,2,4,150,N,"012345678901"
'                   - x offset
'                   - y offset
'                   - rotation
'                   - barcode type/format
'                   - narrow bar width
'                   - wide bar width
'                   - barcode height
'                   - print readable area
'                   - barcode data to print
'
'---------------------------------------------------------------------------------------
'
Public Function PrintEPL2Data(printerName As String, epl2 As String, printJobName As String) As Boolean
    Dim di As DOCINFO
    di.pDocName = printJobName
    di.pDataType = "RAW"
    
    Dim retval As Boolean
    retval = False
    
    Dim hPrinter As Long, lpcWritten As Long
    
    Dim epl2bytes() As Byte
    epl2bytes = StrConv(epl2, vbFromUnicode)
    
    If OpenPrinter(printerName, hPrinter, 0) Then
        If StartDocPrinter(hPrinter, 1, di) Then
            If StartPagePrinter(hPrinter) Then
                retval = CBool(WritePrinter(hPrinter, epl2bytes(0), UBound(epl2bytes) + 1, lpcWritten))
                EndPagePrinter hPrinter
            End If
            EndDocPrinter hPrinter
        End If
        ClosePrinter hPrinter
    End If
    
    If retval = False Then
        Dim errnum As Long
        errnum = GetLastError()
        Dim buf As String
        buf = Space(200)
        FormatMessage &H1000, ByVal 0&, errnum, &H0, buf, 200, ByVal 0&
        MsgBox buf
    End If
    
    PrintEPL2Data = retval
End Function
