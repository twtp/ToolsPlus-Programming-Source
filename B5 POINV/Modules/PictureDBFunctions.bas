Attribute VB_Name = "PictureDBFunctions"
Option Explicit

' auth-bg
Private Const ITEM_PIC_DIR_A_B As String = "h:\Item Images\auth-bg\"
Private Const ITEM_PIC_URL_A_B As String = "http://p1.hostingprod.com/@tools-plus.com/itemimage/authlogo/"
' auth-nobg
Private Const ITEM_PIC_DIR_A_N As String = "h:\Item Images\auth-nobg\"
Private Const ITEM_PIC_URL_A_N As String = "http://p1.hostingprod.com/@tools-plus.com/itemimage/auth-nobg/"
' noauth-bg
Private Const ITEM_PIC_DIR_N_B As String = "h:\Item Images\noauth-bg\"
Private Const ITEM_PIC_URL_N_B As String = "http://p1.hostingprod.com/@tools-plus.com/itemimage/noauth-bg/"
' noauth-nobg
Private Const ITEM_PIC_DIR_N_N As String = "h:\Item Images\noauth-nobg\"
Private Const ITEM_PIC_URL_N_N As String = "http://p1.hostingprod.com/@tools-plus.com/itemimage/noauth-nobg/"

' alternates
'Private Const ITEM_PIC_DIR_ALT As String = "h:\FINAL WEB ALTERNATE\"
'Private Const ITEM_PIC_URL_ALT As String = "http://p1.hostingprod.com/@tools-plus.com/itemimage/alternate/"

' misc
Private Const MFR_LOGO_DIR As String = "h:\Mfg Logos\"
Private Const SECTION_PIC_DIR As String = "h:\section images\"

Private Function picDbNameToDirectory(dbName)
    Select Case dbName
        Case "auth-bg"
            picDbNameToDirectory = ITEM_PIC_DIR_A_B
        Case "auth-nobg"
            picDbNameToDirectory = ITEM_PIC_DIR_A_N
        Case "noauth-bg"
            picDbNameToDirectory = ITEM_PIC_DIR_N_B
        Case "noauth-nobg"
            picDbNameToDirectory = ITEM_PIC_DIR_N_N
        Case Else
            Err.Raise 123, "PictureDBFunctions", "invalid db name"
    End Select
End Function

Private Function picDbNameToURL(dbName)
    Select Case dbName
        Case "auth-bg"
            picDbNameToURL = ITEM_PIC_URL_A_B
        Case "auth-nobg"
            picDbNameToURL = ITEM_PIC_URL_A_N
        Case "noauth-bg"
            picDbNameToURL = ITEM_PIC_URL_N_B
        Case "noauth-nobg"
            picDbNameToURL = ITEM_PIC_URL_N_N
        Case Else
            Err.Raise 123, "PictureDBFunctions", "invalid db name"
    End Select
End Function

Private Function buildPictureFilePath(item As String, dbName As String)
    buildPictureFilePath = picDbNameToDirectory(dbName) & Left(item, 3) & "\" & CreateYahooID(item) & ".jpg"
End Function

Private Function buildPictureURL(item As String, dbName As String)
    buildPictureURL = picDbNameToURL(dbName) & CreateYahooID(item) & ".jpg"
End Function

Private Function pictureExists(item As String, dbName As String) As Boolean
    pictureExists = FileOperations.FileExists(buildPictureFilePath(item, dbName))
End Function

Public Function GenerateItemPicturePath(item As String, dbNameArray As Variant) As String
    Dim path As String, iter As Variant
    path = ""
    For Each iter In dbNameArray
        If pictureExists(item, CStr(iter)) Then
            path = buildPictureFilePath(item, CStr(iter))
            Exit For
        End If
    Next iter
    GenerateItemPicturePath = path
End Function
Public Function GenerateItemPicturePathAny(item As String) As String
    GenerateItemPicturePathAny = GenerateItemPicturePath(item, Array("auth-bg", "noauth-nobg", "auth-nobg"))
End Function


Public Function GenerateItemPictureURL(item As String, includeModifiedTimestamp As Boolean, dbNameArray As Variant) As String
    Dim path As String, url As String, iter As Variant
    path = ""
    url = ""
    For Each iter In dbNameArray
        If pictureExists(item, CStr(iter)) Then
            path = buildPictureFilePath(item, CStr(iter))
            url = buildPictureURL(item, CStr(iter))
            Exit For
        End If
    Next iter
    
    If path <> "" And includeModifiedTimestamp Then
        url = url & "?" & FileOperations.GetFileModifiedDate(path)
    End If
    
    GenerateItemPictureURL = url
End Function

Public Function GenerateLogoPicturePath(item As String) As String
    'will potentially return a non-existent path
    GenerateLogoPicturePath = MFR_LOGO_DIR & Left(item, 3) & ".jpg"
End Function

Public Function GenerateSectionPicturePath(fullURL As String) As String
    'requires the full "power-tools-electric-routers" style url piece
    'will potentially return a non-existent path
    GenerateSectionPicturePath = SECTION_PIC_DIR & fullURL & ".jpg"
End Function

'Public Function GetAllItemImageURLsFor(item As String, includeAuthorizedLogo As Boolean, allowFallbackToNonAuthorized As Boolean, includeModifiedTimestamp As Boolean) As String()
'    'primary
'    ReDim retval(0 To 11) As String
'    retval(0) = PictureDBFunctions.GenerateItemPictureURL(item, includeAuthorizedLogo, allowFallbackToNonAuthorized, includeModifiedTimestamp)
'    'then alternates
'    Dim yid As String
'    yid = CreateYahooID(item)
'    Dim files As Variant
'    files = FileOperations.GetFolderContents(ITEM_PIC_DIR_ALT & Left(item, 3), yid & "-alt*.jpg", True, False)
'    Dim re As RegExp
'    Set re = New RegExp
'    re.IgnoreCase = True
'    re.Pattern = "^" & yid & "-alt(\d+)" & "\.jpg$"
'    Dim file As Variant
'
'    For Each file In files
'        Dim matches As MatchCollection
'        Set matches = re.Execute(file)
'        If matches.count = 1 Then
'            'file is good
'            If matches.item(0).SubMatches.count = 1 Then
'                Dim mtime As String
'                mtime = ""
'                If includeModifiedTimestamp Then
'                    mtime = "?" & FileOperations.GetFileModifiedDate(ITEM_PIC_DIR_ALT & Left(item, 3) & "\" & file)
'                End If
'                If matches.item(0).SubMatches(0) > UBound(retval) Then
'                    ReDim Preserve retval(UBound(retval) * 2) As String
'                End If
'                retval(matches.item(0).SubMatches(0)) = LCase(ITEM_PIC_URL_ALT & file & mtime)
'            End If
'        End If
'    Next file
'
'    'now clear up the empty spots in the array, if any
'    Dim i As Long, j As Long
'    i = 0
'    j = 0
'    For i = 0 To UBound(retval)
'        If retval(i) = "" Then
'            'skip
'        ElseIf i <> j Then
'            retval(j) = retval(i)
'            retval(i) = ""
'            j = j + 1
'        Else
'            'nothing to do
'            j = j + 1
'        End If
'    Next i
'
'    'and trim and return
'    If j = 0 Then
'        GetAllItemImageURLsFor = Split("", "zerolengtharray")
'    Else
'        ReDim Preserve retval(0 To j - 1) As String
'        GetAllItemImageURLsFor = retval
'    End If
'End Function


'Public Const ITEM_PIC_DIR As String = "h:\final store pictures\"
''Public Const WEB_PIC_DIR As String = "h:\final web graphics\"
'Public Const WEB_PIC_DIR As String = "\\toolsplus04\company\final web graphics\"
'Public Const LOGO_PIC_DIR As String = "h:\logo's\"
'Public Const SECTION_PIC_DIR As String = "h:\section images\"
'Public Const BLOG_PIC_DIR As String = "h:\blog pictures\"
'
''---------------------------------------------------------------------------------------
'' Procedure : GenerateItemPicPath
'' DateTime  : 7/27/2005 13:30
'' Author    : briandonorfio
'' Purpose   : Creates item picture path for a given itemnumber.
''---------------------------------------------------------------------------------------
''
'Public Function GenerateItemPicPath(item As String, Optional webpic As Boolean = False, Optional forEBay As Boolean = False) As String
'    'GenerateItemPicPath = IIf(webpic, WEB_PIC_DIR, ITEM_PIC_DIR) & Left(item, 3) & "\" & CreateYahooID(item) & ".jpg"
'    GenerateItemPicPath = WEB_PIC_DIR & Left(item, 3) & "\" & CreateYahooID(item) & IIf(forEBay, "-ebay", "") & ".jpg"
'End Function
'
''---------------------------------------------------------------------------------------
'' Procedure : GenerateItemPicURL
'' Author    : briandonorfio
'' Date      : 8/24/2012
'' Purpose   : Creates a URL to a hosted image for the given item, or empty string when
''             no image exists.
''---------------------------------------------------------------------------------------
''
'Public Function GenerateItemPicURL(item As String) As String
'    'similar to feeds logic, if local then assume also remote on hostingprod
'    'saved as yahooid.jpg, but uploaded as item.jpg, yup
'    If FileExists(GenerateItemPicPath(item, True)) Then
'        GenerateItemPicURL = "http://p1.hostingprod.com/@tools-plus.com/googleimg/" & LCase(item) & ".jpg"
'    Else
'        GenerateItemPicURL = ""
'    End If
'End Function
'
'Public Function GenerateEBayItemPicURL(item As String) As String
'    If FileExists(GenerateItemPicPath(item, True, True)) Then
'        GenerateEBayItemPicURL = "http://p1.hostingprod.com/@tools-plus.com/ebayimg/" & LCase(item) & ".jpg"
'    Else
'        GenerateEBayItemPicURL = GenerateItemPicURL(item)
'    End If
'End Function
'
'Public Function ItemNumberFromPicURL(url As String) As String
'    ItemNumberFromPicURL = UCase(Replace(Replace(url, "http://p1.hostingprod.com/@tools-plus.com/googleimg/", ""), ".jpg", ""))
'End Function
'
''---------------------------------------------------------------------------------------
'' Procedure : GenerateLogoPicPath
'' DateTime  : 7/27/2005 13:30
'' Author    : briandonorfio
'' Purpose   : Creates line code logo path for a given itemnumber.
''---------------------------------------------------------------------------------------
''
'Public Function GenerateLogoPicPath(item As String) As String
'    GenerateLogoPicPath = LOGO_PIC_DIR & Left(item, 3) & ".jpg"
'End Function
'
''---------------------------------------------------------------------------------------
'' Procedure : GenerateReconLogoPicPath
'' DateTime  : 1/9/2009 10:19
'' Author    : briandonorfio
'' Purpose   : Reconditioned image for a product line, I guess you can also just send an
''             itemnumber to this and it'll work.
''---------------------------------------------------------------------------------------
''
'Public Function GenerateReconLogoPicPath(PL As String) As String
'    GenerateReconLogoPicPath = LOGO_PIC_DIR & Left(PL, 3) & "-reconditioned.jpg"
'End Function
'
''---------------------------------------------------------------------------------------
'' Procedure : GenerateSectionImagePath
'' DateTime  : 1/16/2009 11:20
'' Author    : briandonorfio
'' Purpose   : Representative image for a section page. Argument should be the full
''             urlidentifier for the path and parents
''---------------------------------------------------------------------------------------
''
'Public Function GenerateSectionImagePath(fullurlid As String) As String
'    GenerateSectionImagePath = SECTION_PIC_DIR & fullurlid & ".jpg"
'End Function
