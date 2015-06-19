Attribute VB_Name = "WebPathCache"
Option Explicit

'basic path info, indexed by id#
'  - id
'  - websiteid
'  - name (w/o prefix or suffix)
'  - url chunk
'  - url full
'  - parent id
'  - path
'  - prefix
'  - suffix
'  - children


Private WPCache As Dictionary 'hash of hashes, first key is id, 2nd key is attribute
Private WPRoot As Dictionary  'hash of perlarrays of 2-element vb arrays, first key is websiteid, then list of 2-element pairs id & ismanuf

Public Function WebPathCache_Flush() As Boolean
    Set WPCache = Nothing
    Set WPRoot = Nothing
    WebPathCache_Flush = True
End Function

Public Function WebPathCache_Exists(id As String) As Boolean
    If WPCache Is Nothing Then
        Set WPCache = New Dictionary
    End If
    
    WebPathCache_Exists = CBool(WPCache.exists(id))
End Function

Public Function WebPathCache_Add(id As String) As Boolean
    If WPCache Is Nothing Then
        Set WPCache = New Dictionary
    End If
    
    If WPCache.exists(id) Then
        WPCache.Remove id
    End If
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT WebsiteID, WebPathName, URLIdentifier, ParentID, ChildPrefixText, ChildSuffixText, IsManufacturerPage, PublishStatus FROM WebPaths WHERE ID=" & id)
    WPCache.Add id, New Dictionary
    WPCache.item(id).Add "id", CStr(id)
    WPCache.item(id).Add "websiteid", CLng(rst("WebsiteID"))
    WPCache.item(id).Add "name", CStr(rst("WebPathName"))
    WPCache.item(id).Add "url", CStr(rst("URLIdentifier"))
    WPCache.item(id).Add "parent", CStr(Nz(rst("ParentID"), "-1"))
    WPCache.item(id).Add "prefix", CStr(Nz(rst("ChildPrefixText"), ""))
    WPCache.item(id).Add "suffix", CStr(Nz(rst("ChildSuffixText"), ""))
    WPCache.item(id).Add "ismanuf", CBool(rst("IsManufacturerPage"))
    WPCache.item(id).Add "published", CLng(rst("PublishStatus"))
    addRootEntryIfNeeded id, CStr(Nz(rst("ParentID"), "-1")), CBool(rst("IsManufacturerPage"))
    addChildren id
    addParentChildRef id
    
    WebPathCache_Add = True
End Function

Public Function WebPathCache_Remove(id As String) As Boolean
    If WPCache Is Nothing Then
        Set WPCache = New Dictionary
    End If
    
    If WPCache.exists(id) Then
        Dim pid As String
        pid = WPCache.item(id).item("parent")
        If WPCache.exists(pid) Then
            If WPCache.item(pid).item("children").exists(id) Then
                WPCache.item(pid).item("children").Remove id
            End If
        End If
        WPCache.Remove id
    End If
    
    WebPathCache_Remove = True
End Function

Public Function WebPathCache_GetName(id As String, Optional includePrefix As Boolean = True, Optional includeSuffix As Boolean = True) As String
    Dim pre As String, suf As String
    pre = ""
    suf = ""
    If includePrefix Or includeSuffix Then
        Dim pid As String
        pid = WebPathCache_GetParent(id)
        If pid <> "-1" Then
            If includePrefix Then
                pre = WebPathCache_GetPrefix(pid)
            End If
            If includeSuffix Then
                suf = WebPathCache_GetSuffix(pid)
            End If
        End If
    End If
    WebPathCache_GetName = IIf(pre = "", "", pre & " ") & CStr(getProp(id, "name")) & IIf(suf = "", "", " " & suf)
End Function

Public Function WebPathCache_GetURLPart(id As String) As String
    WebPathCache_GetURLPart = CStr(getProp(id, "url"))
End Function

Public Function WebPathCache_GetURLFull(id As String) As String
    Dim retval As String
    retval = WebPathCache_GetURLPart(id)
    Dim pid As String
    pid = WebPathCache_GetParent(id)
    While pid <> "-1"
        retval = WebPathCache_GetURLPart(pid) & "-" & retval
        pid = WebPathCache_GetParent(pid)
    Wend
    WebPathCache_GetURLFull = retval
End Function

Public Function WebPathCache_GetParent(id As String) As String
    WebPathCache_GetParent = CStr(getProp(id, "parent"))
End Function

Public Function WebPathCache_GetPath(id As String) As String
    Dim pid As String, retval As String
    pid = WebPathCache_GetParent(id)
    If pid = "-1" Then
        If WebPathCache_IsManufacturerPage(id) Then
            retval = "search-by-manufacturer"
        Else
            retval = ""
        End If
    Else
        retval = WebPathCache_GetURLPart(pid)
        pid = WebPathCache_GetParent(pid)
        While pid <> "-1"
            retval = WebPathCache_GetURLPart(pid) & ":" & retval
            pid = WebPathCache_GetParent(pid)
        Wend
    End If
    WebPathCache_GetPath = retval
End Function

Public Function WebPathCache_GetPrefix(id As String) As String
    WebPathCache_GetPrefix = CStr(getProp(id, "prefix"))
End Function

Public Function WebPathCache_GetSuffix(id As String) As String
    WebPathCache_GetSuffix = CStr(getProp(id, "suffix"))
End Function

Public Function WebPathCache_GetPublishStatus(id As String) As Long
    WebPathCache_GetPublishStatus = CLng(getProp(id, "published"))
End Function

Public Function WebPathCache_GetChildList(id As String, Optional includeDeleted As Boolean = False) As Variant
    Dim allKeys As Variant
    allKeys = getProp(id, "children")
    If includeDeleted Then
        WebPathCache_GetChildList = allKeys
    Else
        Dim i As Long
        Dim retval As Variant
        retval = Split("", "zerolengtharray")
        For i = 0 To UBound(allKeys)
            If WebPathCache_GetPublishStatus(CStr(allKeys(i))) = 2 Then
                'skip
            Else
                ReDim Preserve retval(UBound(retval) + 1)
                retval(UBound(retval)) = CStr(allKeys(i))
            End If
        Next i
        WebPathCache_GetChildList = retval
    End If
End Function

Public Function WebPathCache_IsManufacturerPage(id As String) As Boolean
    WebPathCache_IsManufacturerPage = CBool(getProp(id, "ismanuf"))
End Function

Public Function WebPathCache_IsInManufacturerTree(id As String) As Boolean
    Dim pid As String
    pid = id
    While WebPathCache_GetParent(pid) <> "-1"
        pid = WebPathCache_GetParent(pid)
    Wend
    WebPathCache_IsInManufacturerTree = WebPathCache_IsManufacturerPage(pid)
End Function

Public Function WebPathCache_GetRoots(websiteID As Long, includeCategories As Boolean, includeManufs As Boolean, Optional includeDeleted As Boolean = False) As Variant
    'conditional population, won't run if not required for this site
    populateRoots websiteID
    
    Dim retval As Variant
    retval = Split("", "zerolengtharray")
    If Not includeCategories And Not includeManufs Then
        'wtf? just skip processing then
    Else
        Dim i As Long
        For i = 0 To WPRoot.item(websiteID).Scalar - 1
            If (includeCategories And Not WPRoot.item(websiteID).Elem(i)(1)) Or (includeManufs And WPRoot.item(websiteID).Elem(i)(1)) Then
                If includeDeleted Or WebPathCache_GetPublishStatus(CStr(WPRoot.item(websiteID).Elem(i)(0))) <> 2 Then
                    ReDim Preserve retval(UBound(retval) + 1)
                    retval(UBound(retval)) = CStr(WPRoot.item(websiteID).Elem(i)(0))
                End If
            End If
        Next i
    End If
    
    WebPathCache_GetRoots = retval
End Function

Private Function populateRoots(websiteID) As Boolean
    'WPRoot might be null, if this is the first website root
    'we're populating. otherwise, we might have already asked
    'for a different website, and this one needs to be filled.
    'if we're asking for the same site again, just return
    If WPRoot Is Nothing Then
        Set WPRoot = New Dictionary
    End If
    If WPRoot.exists(websiteID) Then
        'why are we here?
        populateRoots = True
        Exit Function
    End If
    
    WPRoot.Add websiteID, New PerlArray
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, IsManufacturerPage FROM WebPaths WHERE WebsiteID=" & websiteID & " AND ParentID IS NULL")
    While Not rst.EOF
        WPRoot.item(websiteID).Push Array(CStr(rst("ID")), CBool(rst("IsManufacturerPage")))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    populateRoots = True
End Function

Private Function getProp(id As String, prop As String) As Variant
    If WPCache Is Nothing Then
        Set WPCache = New Dictionary
    End If
    
    If Not WPCache.exists(id) Then
        WebPathCache_Add id
    End If
    
    If TypeName(WPCache.item(id).item(prop)) = "Dictionary" Then
        getProp = WPCache.item(id).item(prop).keys
    Else
        getProp = WPCache.item(id).item(prop)
    End If
End Function

Private Function addChildren(id As String) As Boolean
    WPCache.item(id).Add "children", New Dictionary
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID FROM WebPaths WHERE ParentID=" & id)
    While Not rst.EOF
        WPCache.item(id).item("children").Add CStr(rst("ID")), "1"
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    addChildren = True
End Function

Private Function addParentChildRef(id As String) As Boolean
    Dim pid As String
    pid = WPCache.item(id).item("parent")
    If WPCache.exists(pid) Then
        If Not WPCache.item(pid).item("children").exists(id) Then
            WPCache.item(pid).item("children").Add id, "1"
        End If
    End If
    
    addParentChildRef = True
End Function

Private Function removeParentChildRef(id As String) As Boolean
    Dim pid As String
    pid = WPCache.item(id).item("parent")
    If WPCache.exists(pid) Then
        If WPCache.item(pid).item("children").exists(id) Then
            WPCache.item(pid).item("children").Remove id
        End If
    End If
    
    removeParentChildRef = True
End Function

Private Function addRootEntryIfNeeded(id As String, pid As String, ismanuf As Boolean)
    If pid = "-1" Then
        If Not WPRoot Is Nothing Then
            Dim websiteID As Long
            websiteID = getProp(id, "websiteid")
            Dim i As Long, found As Boolean
            found = False
            If Not WPRoot.exists(websiteID) Then
                WPRoot.Add websiteID, New PerlArray
            End If
            For i = 0 To WPRoot.item(websiteID).Scalar - 1
                If WPRoot.item(websiteID).Elem(i)(0) = id Then
                    found = True
                    Exit For
                End If
            Next i
            If Not found Then
                WPRoot.item(websiteID).Push Array(id, ismanuf)
            End If
        End If
    End If
End Function
