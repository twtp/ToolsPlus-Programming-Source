Attribute VB_Name = "PriceCompareCSharp"
Option Explicit
     Const WM_GETTEXTLENGTH = &HE
     Const WM_GETTEXT = &HD
     Const WM_SETTEXT = &HC

      
Public Type PriceCompareInfo
    ItemName As String
    ResultsCount As Integer
    CompetitorName(0 To 50) As String
    CompetitorPrice(0 To 50) As String
    CompetitorURL(0 To 50) As String
    
End Type


Private Type UUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type
     
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWndParent As Long, _
    ByVal hWndChildAfter As Long, _
    ByVal lpszClassName As String, _
    ByVal lpszWindowName As String _
) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, ByVal msg As Long, _
    wParam As Any, lParam As Any) As Long
    
    
Declare Function EnumChildWindows Lib "user32.dll" (ByVal hWndParent As _
Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

      Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
         (ByVal hwnd As Long, ByVal lpClassName As String, _
         ByVal nMaxCount As Long) As Long
         
Private Declare Function ObjectFromLresult Lib "oleacc" ( _
   ByVal lResult As Long, _
   riid As UUID, _
   ByVal wParam As Long, _
   ppvObject As Any) As Long
   
Private Declare Function RegisterWindowMessage Lib "user32" _
   Alias "RegisterWindowMessageA" ( _
   ByVal lpString As String) As Long
   
Private Declare Function SendMessageTimeout Lib "user32" _
   Alias "SendMessageTimeoutA" ( _
   ByVal hwnd As Long, _
   ByVal msg As Long, _
   ByVal wParam As Long, _
   lParam As Any, _
   ByVal fuFlags As Long, _
   ByVal uTimeout As Long, _
   lpdwResult As Long) As Long

Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long


Private Const SMTO_ABORTIFHUNG = &H2

Public FollowMeMainHandle As Long
         
Public EBayIEHandle As Long
Public GoogleIEHandle As Long
Public YahooIEHandle As Long

Public EBayStatusHandle As Long
Public GoogleStatusHandle As Long
Public YahooStatusHandle As Long




         



Public Sub SetPriceComparerItem(ProductName As String, InputHandle As Long)
    
    
    Dim returnedText As String
    Dim textLength As Long
    textLength = GetWindowTextLength(InputHandle)
    Dim returnLength As Long
    returnedText = Space(textLength + 1)
    returnLength = GetWindowText(InputHandle, returnedText, textLength + 1)
    returnedText = Left(returnedText, returnLength)
    'MsgBox returnedText
    Dim ugh As Long
    
    ugh = SendMessage(InputHandle&, WM_SETTEXT, 0&, ByVal ProductName & "<FIND>")
    
    
    
End Sub
Public Function IsPriceComparerOpen() As Long
    Dim frm_hWnd As Long
    frm_hWnd = FindWindow(vbNullString, "Follow Me Pricing")
    If frm_hWnd > 0 Then
        FollowMeMainHandle = frm_hWnd
        Dim txt_hWnd As Long
        txt_hWnd = FindWindowEx(frm_hWnd, 0, vbNullString, vbNullString)
        txt_hWnd = FindWindowEx(frm_hWnd, txt_hWnd, vbNullString, vbNullString)
        txt_hWnd = FindWindowEx(frm_hWnd, txt_hWnd, vbNullString, vbNullString)
        txt_hWnd = FindWindowEx(frm_hWnd, txt_hWnd, vbNullString, vbNullString)
        txt_hWnd = FindWindowEx(frm_hWnd, txt_hWnd, vbNullString, vbNullString)
        txt_hWnd = FindWindowEx(frm_hWnd, txt_hWnd, vbNullString, vbNullString)
        IsPriceComparerOpen = txt_hWnd
    Else
    'one more check. Its possible we have "consumed" this form inside the main menu of poinv... so...
        frm_hWnd = 0
        frm_hWnd = FindWindow("ThunderRT6Main", "PoInv")
        If frm_hWnd > 0 Then
            txt_hWnd = 0
            frm_hWnd = FindWindowEx(frm_hWnd, 0, vbNullString, "Follow Me Pricing")
            FollowMeMainHandle = frm_hWnd
            txt_hWnd = FindWindowEx(frm_hWnd, txt_hWnd, vbNullString, vbNullString)
            txt_hWnd = FindWindowEx(frm_hWnd, txt_hWnd, vbNullString, vbNullString)
            txt_hWnd = FindWindowEx(frm_hWnd, txt_hWnd, vbNullString, vbNullString)
            txt_hWnd = FindWindowEx(frm_hWnd, txt_hWnd, vbNullString, vbNullString)
            txt_hWnd = FindWindowEx(frm_hWnd, txt_hWnd, vbNullString, vbNullString)
            txt_hWnd = FindWindowEx(frm_hWnd, txt_hWnd, vbNullString, vbNullString)
            IsPriceComparerOpen = txt_hWnd
        Else
        
            IsPriceComparerOpen = 0
        End If
    End If
    
    'MsgBox CStr(txt_hWnd)
    'Dim txt_hWnd As Long
    'txt_hWnd = FindWindowEx(frm_hWnd, 0, "WindowsForms10.EDIT.app.0.b7ab7b_r11_ad1", vbNullString)
    'If txt_hWnd = 0 Then
    'txt_hWnd = FindWindowEx(frm_hWnd, 0, "WindowsForms10.EDIT.app.0.bf7d44_r11_ad1", vbNullString)
    'End If
    'If txt_hWnd = 0 Then
    'txt_hWnd = FindWindowEx(frm_hWnd, 0, "WindowsForms10.EDIT.app.0.2e65425_r11_ad1", vbNullString)
    'End If
    'If txt_hWnd = 0 Then
    'txt_hWnd = FindWindowEx(frm_hWnd, 0, "WindowsForms10.EDIT.app.0.2bf8098_r11_ad1", vbNullString)
    'End If
    'If txt_hWnd = 0 Then
    'txt_hWnd = FindWindowEx(frm_hWnd, 0, "WindowsForms10.EDIT.app.0.b7ab7b_r11_ad1", vbNullString)
    'End If
    'IsPriceComparerOpen = txt_hWnd
        'MsgBox txt_hWnd


End Function
Public Function GetPriceComparerEBayHandle() As Long
Dim frm_hWnd As Long
Dim ebay_hWnd As Long
frm_hWnd = FindWindow(vbNullString, "Follow Me Pricing")
If frm_hWnd > 0 Then
    
    ebay_hWnd = FindWindowEx(frm_hWnd, 0, vbNullString, vbNullString)
    ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
    ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
    ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
    ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
    ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
    ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
    ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
    'ebay_hWnd = EnumChildProc(ebay_hWnd, 0)
    'If ebay_hWnd > 0 Then
     ebay_hWnd = FindWindowEx(ebay_hWnd, 0, vbNullString, vbNullString)
     ebay_hWnd = FindWindowEx(ebay_hWnd, 0, vbNullString, vbNullString)
     ebay_hWnd = FindWindowEx(ebay_hWnd, 0, vbNullString, vbNullString)
     ebay_hWnd = FindWindowEx(ebay_hWnd, 0, vbNullString, vbNullString)
     ebay_hWnd = FindWindowEx(ebay_hWnd, 0, vbNullString, vbNullString)
     
    
    If IsIEServer(ebay_hWnd) = True Then
        'MsgBox "here it is..."
        EBayIEHandle = ebay_hWnd
        'MsgBox GetHTMLDocument(ebay_hWnd)
    Else
        GetPriceComparerEBayHandle = 0
        MsgBox "Could not find a handle in Follow Me Pricing"
    End If

Else

    'check for consumed form...
    frm_hWnd = FindWindow("ThunderRT6Main", "PoInv")
    frm_hWnd = FindWindowEx(frm_hWnd, 0, vbNullString, "Follow Me Pricing")
    If (frm_hWnd > 0) Then
        ebay_hWnd = FindWindowEx(frm_hWnd, 0, vbNullString, vbNullString)
        ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
        ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
        ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
        ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
        ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
        ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
        ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
        'ebay_hWnd = EnumChildProc(ebay_hWnd, 0)
        'If ebay_hWnd > 0 Then
         ebay_hWnd = FindWindowEx(ebay_hWnd, 0, vbNullString, vbNullString)
         ebay_hWnd = FindWindowEx(ebay_hWnd, 0, vbNullString, vbNullString)
         ebay_hWnd = FindWindowEx(ebay_hWnd, 0, vbNullString, vbNullString)
         ebay_hWnd = FindWindowEx(ebay_hWnd, 0, vbNullString, vbNullString)
         ebay_hWnd = FindWindowEx(ebay_hWnd, 0, vbNullString, vbNullString)
         
        
        If IsIEServer(ebay_hWnd) = True Then
            'MsgBox "here it is..."
            EBayIEHandle = ebay_hWnd
            'MsgBox GetHTMLDocument(ebay_hWnd)
        Else
            GetPriceComparerEBayHandle = 0
            MsgBox "Could not find a handle in Follow Me Pricing"
        End If
    Else
        MsgBox "Could not find a handle in Follow Me Pricing"
        GetPriceComparerEBayHandle = 0
        Exit Function
    End If
End If
End Function

Public Function GetPriceComparerGoogleHandle() As Long
Dim frm_hWnd As Long
Dim google_hWnd As Long
Dim tabs As Long
frm_hWnd = FindWindow(vbNullString, "Follow Me Pricing")

If frm_hWnd > 0 Then
    
    google_hWnd = FindWindowEx(frm_hWnd, 0, vbNullString, vbNullString)
    google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    'google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    'google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    'ebay_hWnd = EnumChildProc(ebay_hWnd, 0)
    'If ebay_hWnd > 0 Then
    
    tabs = google_hWnd
     tabs = FindWindowEx(google_hWnd, 0, vbNullString, vbNullString)

     tabs = FindWindowEx(google_hWnd, tabs, vbNullString, vbNullString)
     tabs = FindWindowEx(google_hWnd, tabs, vbNullString, vbNullString)
     google_hWnd = FindWindowEx(google_hWnd, tabs, vbNullString, vbNullString)
     google_hWnd = FindWindowEx(google_hWnd, 0, vbNullString, vbNullString)
     google_hWnd = FindWindowEx(google_hWnd, 0, vbNullString, vbNullString)
     google_hWnd = FindWindowEx(google_hWnd, 0, vbNullString, vbNullString)
     google_hWnd = FindWindowEx(google_hWnd, 0, vbNullString, vbNullString)
     
    
    If IsIEServer(google_hWnd) = True Then
        'MsgBox CStr(google_hWnd)
        GoogleIEHandle = google_hWnd
        'MsgBox GetHTMLDocument(ebay_hWnd)
    Else
        GetPriceComparerGoogleHandle = 0
        MsgBox "Could not find a handle in Follow Me Pricing"
    End If

Else
    'check google handle for being consumed...
    frm_hWnd = FindWindow("ThunderRT6Main", "PoInv")
    frm_hWnd = FindWindowEx(frm_hWnd, 0, vbNullString, "Follow Me Pricing")
    
    If frm_hWnd > 0 Then
    
    google_hWnd = FindWindowEx(frm_hWnd, 0, vbNullString, vbNullString)
    google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    'google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    'google_hWnd = FindWindowEx(frm_hWnd, google_hWnd, vbNullString, vbNullString)
    'ebay_hWnd = EnumChildProc(ebay_hWnd, 0)
    'If ebay_hWnd > 0 Then
    
    tabs = google_hWnd
     tabs = FindWindowEx(google_hWnd, 0, vbNullString, vbNullString)

     tabs = FindWindowEx(google_hWnd, tabs, vbNullString, vbNullString)
     tabs = FindWindowEx(google_hWnd, tabs, vbNullString, vbNullString)
     google_hWnd = FindWindowEx(google_hWnd, tabs, vbNullString, vbNullString)
     google_hWnd = FindWindowEx(google_hWnd, 0, vbNullString, vbNullString)
     google_hWnd = FindWindowEx(google_hWnd, 0, vbNullString, vbNullString)
     google_hWnd = FindWindowEx(google_hWnd, 0, vbNullString, vbNullString)
     google_hWnd = FindWindowEx(google_hWnd, 0, vbNullString, vbNullString)
     
    
    If IsIEServer(google_hWnd) = True Then
        'MsgBox CStr(google_hWnd)
        GoogleIEHandle = google_hWnd
        'MsgBox GetHTMLDocument(ebay_hWnd)
    Else
        GetPriceComparerGoogleHandle = 0
        MsgBox "Could not find a handle in Follow Me Pricing"
    End If

Else

    MsgBox "Could not find a handle in Follow Me Pricing"
    GetPriceComparerGoogleHandle = 0
    Exit Function
    End If
End If
End Function

Public Function GetPriceComparerStatusHandles()
Dim frm_hWnd As Long
Dim ebay_hWnd As Long
frm_hWnd = FindWindow(vbNullString, "Follow Me Pricing")
If frm_hWnd > 0 Then
    
    ebay_hWnd = FindWindowEx(frm_hWnd, 0, vbNullString, vbNullString)
    YahooStatusHandle = ebay_hWnd
    
    
    ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
    GoogleStatusHandle = ebay_hWnd
    
    
    ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
    
    EBayStatusHandle = ebay_hWnd
    


Else

    'perhaps the form was consumed...
    frm_hWnd = 0
    frm_hWnd = FindWindow("ThunderRT6Main", "PoInv")
    If (frm_hWnd > 0) Then
        frm_hWnd = FindWindowEx(frm_hWnd, 0, vbNullString, "Follow Me Pricing")
        ebay_hWnd = FindWindowEx(frm_hWnd, 0, vbNullString, vbNullString)
        YahooStatusHandle = ebay_hWnd
    
    
        ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
        GoogleStatusHandle = ebay_hWnd
        
        
        ebay_hWnd = FindWindowEx(frm_hWnd, ebay_hWnd, vbNullString, vbNullString)
        
        EBayStatusHandle = ebay_hWnd
    Else

        MsgBox "Could not find a handle in Follow Me Pricing"
        Exit Function
    End If
End If


End Function




Function IEDOMFromhWnd(ByVal hwnd As Long) As HTMLDocument
Dim IID_IHTMLDocument As UUID
Dim hWndChild As Long
Dim lRes As Long
Dim lMsg As Long
Dim hr As Long

   If hwnd <> 0 Then
      
      'If Not IsIEServerWindow(hWnd) Then
      
         ' Find a child IE server window
      '   EnumChildWindows hWnd, AddressOf EnumChildProc, hWnd
         
      'End If
      
      If hwnd <> 0 Then
            
         ' Register the message
         lMsg = RegisterWindowMessage("WM_HTML_GETOBJECT")
            
         ' Get the object pointer
         Call SendMessageTimeout(hwnd, lMsg, 0, 0, _
                 SMTO_ABORTIFHUNG, 1000, lRes)

         If lRes Then
               
            ' Initialize the interface ID
            With IID_IHTMLDocument
               .Data1 = &H626FC520
               .Data2 = &HA41E
               .Data3 = &H11CF
               .Data4(0) = &HA7
               .Data4(1) = &H31
               .Data4(2) = &H0
               .Data4(3) = &HA0
               .Data4(4) = &HC9
               .Data4(5) = &H8
               .Data4(6) = &H26
               .Data4(7) = &H37
            End With
               
            ' Get the object from lRes
            hr = ObjectFromLresult(lRes, IID_IHTMLDocument, 0, IEDOMFromhWnd)
               
         End If

      End If
      
   End If

End Function


Public Function GetObjectText(hwnd As Long) As String
    Dim buffer As String * 255
    Dim retval As Long
    retval = GetWindowText(hwnd, buffer, 255)
    GetObjectText = StripNulls(buffer)
End Function
Public Function SetObjectText(hwnd As Long, caption As String) As Long
    Dim retval As Long
    
    retval = SetWindowText(hwnd, caption)
    SetObjectText = retval

End Function


Function EnumChildProc(ByVal lhWnd As Long, ByVal lParam As Long) _
         As Long
         Dim retval As Long
         Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
         Dim WinClass As String, WinTitle As String
         'Dim WinRect As RECT
         Dim WinWidth As Long, WinHeight As Long

         retval = GetClassName(lhWnd, WinClassBuf, 255)
         WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
         retval = GetWindowText(lhWnd, WinTitleBuf, 255)
         WinTitle = StripNulls(WinTitleBuf)
         'ChildCount = ChildCount + 1
         ' see the Windows Class and Title for each Child Window enumerated
         Debug.Print "   Child Class = "; WinClass; ", Title = "; WinTitle
         ' You can find any type of Window by searching for its WinClass
         If InStr(WinClass, "WindowsForms10.SysTabControl32.app") Then   ' TextBox Window
            'RetVal = GetWindowRect(lhWnd, WinRect)  ' get current size
            'WinWidth = WinRect.Right - WinRect.Left ' keep current width
            'WinHeight = (WinRect.Bottom - WinRect.Top) * 2 ' double height
            'RetVal = MoveWindow(lhWnd, 0, 0, WinWidth, WinHeight, True)
            EnumChildProc = lhWnd
         Else
            EnumChildProc = 0
         End If
End Function
Function IsIEServer(ByVal lhWnd As Long) As Boolean
         Dim retval As Long
         Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
         Dim WinClass As String, WinTitle As String
         'Dim WinRect As RECT
         Dim WinWidth As Long, WinHeight As Long

         retval = GetClassName(lhWnd, WinClassBuf, 255)
         WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
         retval = GetWindowText(lhWnd, WinTitleBuf, 255)
         WinTitle = StripNulls(WinTitleBuf)
         'ChildCount = ChildCount + 1
         ' see the Windows Class and Title for each Child Window enumerated
         Debug.Print "   Child Class = "; WinClass; ", Title = "; WinTitle
         ' You can find any type of Window by searching for its WinClass
         If InStr(WinClass, "Internet Explorer_Server") Then   ' TextBox Window
            'RetVal = GetWindowRect(lhWnd, WinRect)  ' get current size
            'WinWidth = WinRect.Right - WinRect.Left ' keep current width
            'WinHeight = (WinRect.Bottom - WinRect.Top) * 2 ' double height
            'RetVal = MoveWindow(lhWnd, 0, 0, WinWidth, WinHeight, True)
           IsIEServer = True
         Else
           IsIEServer = False
         End If
End Function
      
      
       Public Function StripNulls(OriginalStr As String) As String
         ' This removes the extra Nulls so String comparisons will work
         If (InStr(OriginalStr, Chr(0)) > 0) Then
            OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
         End If
         StripNulls = OriginalStr
      End Function
      
            
Public Function GetEBayComparisons(html As String) As PriceCompareInfo
    Dim tmpInfo As PriceCompareInfo
    
    Dim Result() As String
    Dim ResultsCount As Integer
    Result = Split(html, "<!--{SELLER}-->")
    ResultsCount = UBound(Result)
    'MsgBox Chr(60)
    tmpInfo.ResultsCount = ResultsCount
    
    If ResultsCount <= 0 Then
    'GetEBayComparisons = vbNull
    Exit Function
    End If
    
    'ReDim tmpInfo.CompetitorName(0 To ResultsCount - 1)
    'ReDim tmpInfo.CompetitorPrice(0 To ResultsCount - 1)
    'ReDim tmpInfo.CompetitorURL(0 To ResultsCount - 1)
    Dim xcount As Integer
    xcount = 0
    Do
        Dim cPrice As String
        cPrice = Split(Split(html, "<!--{PRICE}-->")(xcount + 1), "<!--{/PRICE}-->")(0)
        
        If InStr(cPrice, "(") Then
            cPrice = "$" & Split(Split(Split(cPrice, "(")(1), "$")(1), "<")(0)
        End If
        
        tmpInfo.CompetitorName(xcount) = Trim(Split(Split(Split(html, "<!--{SELLER}-->")(xcount + 1), "<!--{/SELLER}-->")(0), "<")(0))
        tmpInfo.CompetitorPrice(xcount) = cPrice
        tmpInfo.CompetitorURL(xcount) = Split(Split(html, "<!--{ITEMURL}-->")(xcount + 1), "<!--{/ITEMURL}-->")(0)
        xcount = xcount + 1
    Loop Until xcount >= ResultsCount
GetEBayComparisons = tmpInfo

End Function


Public Function GetGoogleComparisons(html As String) As PriceCompareInfo
    Dim tmpInfo As PriceCompareInfo
    
    Dim Result() As String
    Dim ResultsCount As Integer
    Result = Split(html, "<!--{SELLER}-->")
    ResultsCount = UBound(Result)
    'MsgBox Chr(60)
    tmpInfo.ResultsCount = ResultsCount
    
    If ResultsCount <= 0 Then
    'GetEBayComparisons = vbNull
    Exit Function
    End If
    
    'ReDim tmpInfo.CompetitorName(0 To ResultsCount - 1)
    'ReDim tmpInfo.CompetitorPrice(0 To ResultsCount - 1)
    'ReDim tmpInfo.CompetitorURL(0 To ResultsCount - 1)
    Dim xcount As Integer
    xcount = 0
    Do
        Dim cPrice As String
        cPrice = Split(Split(html, "<!--{PRICE}-->")(xcount + 1), "<!--{/PRICE}-->")(0)
        
        If InStr(cPrice, "(") Then
            cPrice = "$" & Split(Split(Split(cPrice, "(")(1), "$")(1), "<")(0)
        End If
        
        tmpInfo.CompetitorName(xcount) = Trim(Split(Split(Split(html, "<!--{SELLER}-->")(xcount + 1), "<!--{/SELLER}-->")(0), "<")(0))
        tmpInfo.CompetitorPrice(xcount) = cPrice
        tmpInfo.CompetitorURL(xcount) = Split(Split(html, "<!--{ITEMURL}-->")(xcount + 1), "<!--{/ITEMURL}-->")(0)
        xcount = xcount + 1
    Loop Until xcount >= ResultsCount
GetGoogleComparisons = tmpInfo

End Function
      
      
      
      
      
      
      
      
      
      
      
      
      
      

