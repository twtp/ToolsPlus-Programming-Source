VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormTooltips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private tip_hwnd As Long

Private Const WS_EX_TOPMOST As Long = &H8&
Private Const CW_USEDEFAULT As Long = &H80000000
Private Const HWND_TOPMOST As Long = -1

Private Const TTF_IDISHWND As Long = &H1
Private Const TTF_SUBCLASS As Long = &H10

Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOSIZE As Long = &H1

Private Const WS_POPUP As Long = &H80000000
Private Const TTS_BALLOON As Long = &H40

Private Const WM_USER As Long = &H400
Private Const TTM_ACTIVATE As Long = (WM_USER + 1)
Private Const TTM_SETDELAYTIME As Long = (WM_USER + 3)
Private Const TTM_ADDTOOL As Long = (WM_USER + 4)
Private Const TTM_DELTOOL As Long = (WM_USER + 5)
Private Const TTM_GETTOOLINFO As Long = (WM_USER + 8)
Private Const TTM_UPDATETIPTEXT As Long = (WM_USER + 12)
Private Const TTM_SETTIPBKCOLOR As Long = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR As Long = (WM_USER + 20)
Private Const TTM_SETMAXTIPWIDTH As Long = (WM_USER + 24)

Private Const TTDT_AUTOPOP As Long = 2

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type TOOLINFO
  cbSize As Long
  uFlags As Long
  hWnd As Long
  uid As Long
  rc As RECT
  hinst As Long
  lpszText As String
  lParam As Long
End Type

Private Declare Function InitCommonControls Lib "comctl32" () As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Function addToolTip(ctlHwnd As Long, formHwnd As Long, txt As String) As Boolean
    If tip_hwnd = 0 Then
        tip_hwnd = create_tooltip(formHwnd)
        If tip_hwnd = 0 Then
            addToolTip = False
        End If
    End If
    
    'add_tooltip tip_hwnd, ctlHwnd, txt
    
    If tip_hwnd <> 0 Then
        Dim ti As TOOLINFO
        ti.cbSize = Len(ti)
        ti.hWnd = ctlHwnd
        ti.uid = ctlHwnd
        ti.lpszText = txt
        ti.uFlags = TTF_SUBCLASS Or TTF_IDISHWND
        ti.hinst = app.hInstance
        If tooltip_exists(ctlHwnd) Then
            SendMessage tip_hwnd, TTM_UPDATETIPTEXT, 0&, ti
        Else
            Dim rc As RECT
            GetClientRect ctlHwnd, rc
            ti.rc = rc
            SendMessage tip_hwnd, TTM_ADDTOOL, 0&, ti
        End If
        addToolTip = True
    End If
End Function

Public Function AddToolTipToCtl(ctl As VB.Control, txt As String) As Boolean
    AddToolTipToCtl = addToolTip(ctl.hWnd, ctl.Parent.hWnd, txt)
End Function

'Public Function EditToolTip(newTxt As String, ctlHwnd As Long) As Boolean
'    If tip_hwnd <> 0 Then
'        Dim ti As TOOLINFO
'        ti.cbSize = Len(ti)
'        ti.hwnd = ctlHwnd
'        ti.uid = ctlHwnd
'        ti.lpszText = newTxt
'        SendMessage tip_hwnd, TTM_UPDATETIPTEXT, 0&, ti
'        EditToolTip = True
'    Else
'        EditToolTip = False
'    End If
'End Function

Private Function removeToolTip(ctlHwnd As Long) As Boolean
    If tip_hwnd <> 0 Then
        Dim ti As TOOLINFO
        ti.cbSize = Len(ti)
        ti.hWnd = ctlHwnd
        ti.uid = ctlHwnd
        SendMessage tip_hwnd, TTM_DELTOOL, 0&, ti
        removeToolTip = True
    Else
        removeToolTip = False
    End If
End Function

Public Function RemoveToolTipFromCtl(ctl As VB.Control) As Boolean
    RemoveToolTipFromCtl = removeToolTip(ctl.hWnd)
End Function

Private Sub Class_Initialize()
    InitCommonControls
End Sub

Private Sub Class_Terminate()
    tip_hwnd = destroy_tooltip(tip_hwnd)
End Sub







Private Function tooltip_exists(ctlHwnd As Long) As Boolean
    'tooltip_exists = False
    Dim ti As TOOLINFO
    ti.cbSize = Len(ti)
    ti.hWnd = ctlHwnd
    ti.uid = ctlHwnd
    SendMessage tip_hwnd, TTM_GETTOOLINFO, 0&, ti
    
    tooltip_exists = Not CBool(ti.hinst = 0)
End Function

Private Function create_tooltip(formHwnd As Long) As Long
    Dim newHwnd As Long
    newHwnd = CreateWindowEx(WS_EX_TOPMOST, _
                             "tooltips_class32", _
                             vbNullString, _
                             WS_POPUP Or TTS_BALLOON, _
                             CW_USEDEFAULT, CW_USEDEFAULT, _
                             CW_USEDEFAULT, CW_USEDEFAULT, _
                             formHwnd, _
                             0&, _
                             app.hInstance, _
                             ByVal 0& _
                            )
    If newHwnd <> 0 Then
        SetWindowPos newHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
        
        SendMessage newHwnd, TTM_SETTIPBKCOLOR, &HE1FFFF, ByVal 0&
        SendMessage newHwnd, TTM_SETTIPTEXTCOLOR, vbBlack, ByVal 0&
        'SendMessage newHwnd, TTM_SETMAXTIPWIDTH, 0, ByVal 0&
        SendMessage newHwnd, TTM_SETMAXTIPWIDTH, 0, ByVal 350&
        SendMessage newHwnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, ByVal 32767&
        
        SendMessage newHwnd, TTM_ACTIVATE, 1, ByVal 0&
    End If
    
    create_tooltip = newHwnd
End Function

Private Function destroy_tooltip(hWnd As Long) As Long
    If hWnd <> 0 Then
        If DestroyWindow(hWnd) Then
            destroy_tooltip = 0
        Else
            destroy_tooltip = hWnd
        End If
    End If
End Function


