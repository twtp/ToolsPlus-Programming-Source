Attribute VB_Name = "ScreenCap"
Option Explicit

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

Private Declare Function GetForegroundWindow Lib "USER32" () As Long
Private Declare Function GetWindowRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function GetWindowDC Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "USER32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long


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

