Attribute VB_Name = "WebOffloadFunctions"
Option Explicit

Private groupedItemsH As Dictionary 'hashes group subitem to group master
Private templateItemsH As Dictionary 'hashes group master to array(null/rst, group id, group url, show thumbs, show avail)

Private specialsItemsH As Dictionary 'hashes itemnumber to array(section image/url stanza, flat percent discount, rebate amount, rebate instant, specials tab, info tab, above tab)

Public Function UpdateWebOffloadChangedQuantity()
    If IsFormLoaded("WebOffload") Then
        SetControlValue "WebOffload", "lblWebOffloadCount", "(" & DLookup("COUNT(*)", "vWebOffload", "WebsiteID=" & WebOffload.btnWebsiteSelection.tag & " AND LastChanged>=LastOffload") & ")"
        SetControlValue "WebOffload", "lblWebOffloadSectionCount", "(" & DLookup("COUNT(*)", "WebPaths", "WebsiteID=" & WebOffload.btnWebsiteSelection.tag & " AND PublishStatus=1 AND LastChanged>=LastOffload") & ")"
    Else
        Load WebOffload
        WebOffload.Show
    End If
End Function

Public Function GetHTMLContactInfoFor(sectionid As String) As String
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, ManufFullNameCleaned, CompanyAddress1, CompanyAddress2, CompanyCity, CompanyState, CompanyZip, CompanyCountry FROM ProductLine WHERE WebSectionID=" & sectionid)
    
    Dim contactinfohtml As String
    If rst.EOF Then
        contactinfohtml = ""
    ElseIf Nz(rst("CompanyAddress1")) = "" Or Nz(rst("CompanyCity")) = "" Then
        contactinfohtml = ""
    Else
        contactinfohtml = "<p>" & rst("ManufFullNameCleaned") & "</p>" & vbCrLf
        contactinfohtml = contactinfohtml & "<p>" & rst("CompanyAddress1") & "</p>" & vbCrLf
        If Nz(rst("CompanyAddress2")) <> "" Then
            contactinfohtml = contactinfohtml & "<p>" & rst("CompanyAddress2") & "</p>" & vbCrLf
        End If
        contactinfohtml = contactinfohtml & "<p>" & rst("CompanyCity")
        If Nz(rst("CompanyState")) <> "" Then
            contactinfohtml = contactinfohtml & ", " & rst("CompanyState")
        End If
        If Nz(rst("CompanyZip")) <> "" Then
            contactinfohtml = contactinfohtml & " " & rst("CompanyZip")
        End If
        If Nz(rst("CompanyCountry")) <> "" And _
           Nz(rst("CompanyCountry")) <> "US" And _
           Nz(rst("CompanyCountry")) <> "USA" Then
            contactinfohtml = contactinfohtml & " " & rst("CompanyCountry")
        End If
        contactinfohtml = contactinfohtml & "</p>" & vbCrLf
        Dim contactrst As ADODB.Recordset
        Set contactrst = DB.retrieve("SELECT Phone, Fax, Email, URL FROM ProductLineContactInfo WHERE ProductLineID=" & rst("ID") & " AND InfoTypeID=0")
        If contactrst.EOF Then
            'do nothing
        Else
            If Nz(contactrst("Phone")) <> "" Then
                contactinfohtml = contactinfohtml & "<p>Phone: " & contactrst("Phone") & "</p>" & vbCrLf
            End If
            If Nz(contactrst("Fax")) <> "" Then
                contactinfohtml = contactinfohtml & "<p>Fax: " & contactrst("Fax") & "</p>" & vbCrLf
            End If
            'If Nz(contactrst("Email")) <> "" Then
            '    contactinfohtml = contactinfohtml & "<p><a href=""mailto:" & contactrst("Email") & """>" & contactrst("Email") & "</a></p>" & vbCrLf
            'End If
            If Nz(contactrst("URL")) <> "" Then
                Dim urlwithhttp As String
                urlwithhttp = IIf(Left(LCase(contactrst("URL")), 7) = "http://", LCase(contactrst("URL")), "http://" & LCase(contactrst("URL")))
                Dim urlwithouthttp As String
                urlwithouthttp = Replace(urlwithhttp, "http://", "")
                'drop the slash on the end, if it ends in a .com domain, do we need something better here?
                'we can't just drop the slash completely, because if there is a directory structure, then
                'example.com/foo and example.com/foo/ can be different resources (probably not though?)
                If Right(urlwithouthttp, 5) = ".com/" Then
                    urlwithouthttp = Left(urlwithouthttp, Len(urlwithouthttp) - 1)
                End If
                'contactinfohtml = contactinfohtml & "<p><a href=""" & urlwithhttp & """ target=""_blank"">" & urlwithouthttp & "</a></p>" & vbCrLf
                contactinfohtml = contactinfohtml & "<p>" & urlwithouthttp & "</p>" & vbCrLf
            End If
        End If
        contactrst.Close
        Set contactrst = Nothing
    End If
    rst.Close
    Set rst = Nothing
    
    GetHTMLContactInfoFor = contactinfohtml
End Function

'--------------------------------------------------------------------------------------
' Procedure : GetItemAccessoriesAsString
' DateTime  : 7/1/2005 16:03
' Author    : briandonorfio
' Purpose   : Returns a space-separated list of the item's accessories, null string if
'             item not found or no accessories. Will limit the number of accessories to
'             a good display amount for the item-grid on the site (anything <=4, or even
'             numbers up to a max of 12)
'---------------------------------------------------------------------------------------
'
Public Function GetItemAccessoriesAsString(item As String) As String
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT AccessoryYahooID, AccessoryItemNumber FROM vPartNumberAccessoriesGood WHERE ItemNumber='" & item & "'")
    If rst.EOF Then
        GetItemAccessoriesAsString = ""
    Else
        ReDim allAccs(rst.RecordCount - 1) As String
        Dim i As Long
        For i = 0 To rst.RecordCount - 1
            allAccs(i) = rst("AccessoryYahooID")
            rst.MoveNext
        Next i
        rst.Close
        Set rst = Nothing
        
        Dim max As Long
        If UBound(allAccs) > 15 Then 'allow up to 16
            max = 15
        Else
            max = UBound(allAccs)
        End If
        
        Randomize
        
        ReDim selectedAccs(max) As String
        Dim j As Long
        For i = 0 To max
            Dim randnum As Long
            randnum = Int(UBound(allAccs)) * Rnd
            selectedAccs(i) = allAccs(randnum)
            For j = randnum + 1 To UBound(allAccs)
                allAccs(j - 1) = allAccs(j)
            Next j
            If UBound(allAccs) <> 0 Then
                ReDim Preserve allAccs(UBound(allAccs) - 1) As String
            End If
        Next i
        
        QuickSort selectedAccs
        
        GetItemAccessoriesAsString = Join(selectedAccs, " ")
    End If
    
    Mouse.Hourglass False
End Function

'---------------------------------------------------------------------------------------
' Procedure : GenerateYahooMetaKeywords
' DateTime  : 3/6/2006 12:38
' Author    : briandonorfio
' Purpose   : Creates a list of meta-keywords for the webpage, based on the path
'             hierarchy.
'---------------------------------------------------------------------------------------
'
Public Function GenerateYahooMetaKeywords(item As String, Optional groupPageID As String = "-1") As String
    Dim rst As ADODB.Recordset
    
    Dim kwds As PerlArray
    Set kwds = New PerlArray
    
    Set rst = DB.retrieve("SELECT WebPathID FROM PartNumberWebPaths WHERE ItemNumber='" & item & "'")
    If rst.RecordCount = 0 Then
        'grouped subitems don't need paths, but since canonical-linked to parent, no meta is ok too
        GenerateYahooMetaKeywords = ""
        Exit Function
    End If
    ReDim paths(rst.RecordCount - 1) As String
    Dim i As Long
    For i = 0 To rst.RecordCount - 1
        paths(i) = CStr(rst("WebPathID"))
        rst.MoveNext
    Next i
    rst.Close
    Dim done As Boolean
    done = False
    Dim which As Long
    which = 0
    While Not done
        If paths(which) <> "-1" Then
            kwds.Push WebPathCache_GetName(paths(which), True, True)
            paths(which) = WebPathCache_GetParent(paths(which))
        End If
        
        which = which + 1
        If which > UBound(paths) Then
            which = 0
        End If
        
        done = True
        For i = 0 To UBound(paths)
            If paths(i) <> "-1" Then
                done = False
                Exit For
            End If
        Next i
    Wend

    Set rst = DB.retrieve("SELECT ManufFullNameWeb, Desc1, Desc2, Desc3 FROM PartNumbers INNER JOIN ProductLine ON SUBSTRING(PartNumbers.ItemNumber,1,3)=ProductLine.ProductLine WHERE PartNumbers.ItemNumber='" & item & "'")
    kwds.UnShift Nz(rst("ManufFullNameWeb")) & " " & rst("Desc3")
    kwds.UnShift Nz(rst("ManufFullNameWeb")) & " " & rst("Desc2") & " " & rst("Desc1") & " " & rst("Desc3")
    rst.Close
    
    For i = 0 To kwds.Scalar - 1
        kwds.Elem(i) = StripHTML(kwds.Elem(i), True)
        kwds.Elem(i) = Replace(kwds.Elem(i), ",", "")
        kwds.Elem(i) = EscapeHTMLEntities(kwds.Elem(i))
        kwds.Elem(i) = TrimWhitespace(kwds.Elem(i), True, True, True)
    Next i
    
    Dim retval As String
    For i = 0 To kwds.Scalar - 1
        If kwds.Elem(i) <> "" Then
            '512 is also hard coded as length in the web offload function, changes should be
            'made in both places
            If Len(retval & kwds.Elem(i)) > 512 Then
                'skip, all done
                Exit For
            Else
                retval = retval & kwds.Elem(i) & ", "
                Dim j As Long
                For j = i + 1 To kwds.Scalar - 1
                    If kwds.Elem(j) = kwds.Elem(i) Then
                        kwds.Elem(j) = ""
                    End If
                Next j
            End If
        End If
    Next i
    
    GenerateYahooMetaKeywords = Left(retval, Len(retval) - 2)
End Function

Public Function GenerateIncludesTextForCaption(item As String) As String
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT IncludeText FROM PartNumberIncludes WHERE ItemNumber='" & item & "' ORDER BY SortOrder")
    Dim retval As String
    While Not rst.EOF
        retval = IIf(retval = "", "", retval & ", ") & rst("IncludeText")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    If retval = "" Then
        GenerateIncludesTextForCaption = ""
    Else
        GenerateIncludesTextForCaption = "includes " & retval
    End If
End Function

Public Function GenerateHighlightsTextForCaption(rstItem As ADODB.Recordset) As String
    Dim html As String
    html = "<ul>"
    Dim i As Long
    For i = 1 To 8
        Dim h As String
        h = GenerateHighlightByPosition(rstItem, i)
        If h <> "" Then
            html = html & "<li>" & h & "</li>"
        End If
    Next i
    html = html & "</ul>"
    GenerateHighlightsTextForCaption = Utilities.StripHTML(html)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GenerateCrossSellText
' DateTime  : 9/26/2006 15:50
' Author    : briandonorfio
' Purpose   : Returns the div for the cross sell text on the website.
'---------------------------------------------------------------------------------------
'
Public Function GenerateCrossSellText(item As String, Optional xhtml As Boolean = False) As String
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ShowText, URL FROM PartNumberCrossSell WHERE ItemNumber='" & item & "' ORDER BY SortOrder")
    Dim retval As String
    While Not rst.EOF
        retval = retval & rst("URL") & vbCrLf & rst("ShowText") & vbCrLf
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    GenerateCrossSellText = Chomp(retval)
End Function

Public Function GenerateKeywordAlternatesFromPathing(item As String) As String
    Dim retval As PerlArray
    Set retval = New PerlArray
    Dim rst As ADODB.Recordset
    If GroupItemCache_IsGroupedSubItem(item) Then
        Set rst = DB.retrieve("SELECT WebPathID FROM PartNumberWebPaths WHERE ItemNumber='" & GroupItemCache_GetMasterItemNumberForSubItem(item) & "'")
    Else
        Set rst = DB.retrieve("SELECT WebPathID FROM PartNumberWebPaths WHERE ItemNumber='" & item & "'")
    End If
    While Not rst.EOF
        If WebPathCache.WebPathCache_IsInManufacturerTree(rst("WebPathID")) Then
            Dim kwdstr As String
            kwdstr = WebPathCache.WebPathCache_GetName(rst("WebPathID")) & "," & DLookup("COALESCE(MetaKeywords,'')", "WebPaths", "ID=" & rst("WebPathID"))
            'Dim pid As String
            'pid = WebPathCache.WebPathCache_GetParent(rst("WebPathID"))
            'If pid <> "-1" Then
            '    kwdstr = kwdstr & "," & WebPathCache.WebPathCache_GetName(pid) & "," & DLookup("COALESCE(MetaKeywords,'')", "WebPaths", "ID=" & pid)
            'End If
            Dim kwdA As Variant
            kwdA = Split(kwdstr, ",")
            Dim i As Long
            For i = LBound(kwdA) To UBound(kwdA)
                kwdA(i) = TrimWhitespace(CStr(kwdA(i)), True, True, True, True)
                If kwdA(i) <> "" Then
                    If retval.FindString(CStr(kwdA(i)), vbTextCompare) = -1 Then
                        retval.Push CStr(kwdA(i))
                    End If
                End If
            Next i
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    If retval.Scalar = 0 Then
        GenerateKeywordAlternatesFromPathing = ""
    Else
        GenerateKeywordAlternatesFromPathing = retval.Join(", ")
    End If
End Function



Public Function GenerateSectionPageSpecsList(rstItem As ADODB.Recordset) As String
    'same recordset switch as the full contents when dealing with a group sub-item
    Dim rst As ADODB.Recordset
    If GroupItemCache_IsGroupedSubItem(rstItem("ItemNumber")) Then
        'need to switch recordsets to the group
        Set rst = GroupItemCache_GetGroupRecordsetBySubItemNumber(rstItem("ItemNumber"))
    Else
        'current recordset is okay
        Set rst = rstItem
    End If
    
    Dim i As Long
    Dim Spec(1 To 8) As String
    If rst("ShowDCSpecs") Then
        For i = 1 To 8
            Spec(i) = Trim(FillDCSpecs(rst("ItemNumber"), i))
        Next i
    Else
        For i = 1 To 8
            Spec(i) = Trim(Nz(rst("Spec" & i & "HTML")))
        Next i
    End If
    
    Dim partspecs As String
    partspecs = ""
    For i = 1 To 3
        If Spec(i) <> "" Then
            partspecs = partspecs & " <li>" & Spec(i) & "</li>" & vbCrLf
        End If
    Next i
    If partspecs <> "" Then
        partspecs = "<ul class=""specs-short-list"">" & vbCrLf & partspecs & "</ul>"
    End If
    
    GenerateSectionPageSpecsList = partspecs
End Function

Public Function GenerateHighlightByPosition(rstItem As ADODB.Recordset, pos As Long) As String
    'parameter is the recordset for the item or group master item (basically the thing we want the specs from)
    
    'set up group information
    '  reset recordset to use when grouped
    '  merge d/c specs flag
    Dim rst As ADODB.Recordset
    If GroupItemCache_IsGroupedSubItem(rstItem("ItemNumber")) Then
        'need to switch recordsets to the group
        Set rst = GroupItemCache_GetGroupRecordsetBySubItemNumber(rstItem("ItemNumber"))
    Else
        'current recordset is okay
        Set rst = rstItem
    End If
    Dim useDCSpecs As Boolean
    useDCSpecs = CBool(rst("ShowDCSpecs") Or rstItem("ShowDCSpecs")) 'item d/c or subitem d/c

    If useDCSpecs Then
        GenerateHighlightByPosition = CStr(Trim(FillDCSpecs(rst("ItemNumber"), pos)))
    Else
        Dim highlights As PerlArray
        Set highlights = New PerlArray
        Dim i As Long
        For i = 1 To 8
            highlights.Push CStr(Trim(Nz(rst("Spec" & i & "HTML"))))
        Next i
        If pos <= highlights.Scalar Then
            GenerateHighlightByPosition = highlights.Elem(pos - 1)
        Else
            GenerateHighlightByPosition = ""
        End If
    End If
End Function

Public Function GenerateItemPageContents(rstItem As ADODB.Recordset) As String
    'parameter is the recordset for the item or group master item (basically the thing we want the specs from)
    
    'set up group information
    '  reset recordset to use when grouped
    '  grab misc info required
    '  merge d/c specs flag
    Dim rst As ADODB.Recordset
    If GroupItemCache_IsGroupedSubItem(rstItem("ItemNumber")) Then
        'need to switch recordsets to the group
        Set rst = GroupItemCache_GetGroupRecordsetBySubItemNumber(rstItem("ItemNumber"))
    Else
        'current recordset is okay
        Set rst = rstItem
    End If
    Dim groupSubItemNumber As String, groupMasterItemID As Long, groupMasterTemplateURL As String
    If GroupItemCache_IsGroupedSubItem(rstItem("ItemNumber")) Then
        'subitems get these filled
        groupSubItemNumber = rstItem("ItemNumber")
        groupMasterItemID = GroupItemCache_GetGroupDatabaseIDByMasterItemNumber(rst("ItemNumber"))
        groupMasterTemplateURL = GroupItemCache_GetGroupYahooIDByMasterItemNumber(rst("ItemNumber"))
    Else
        'regular items and group masters get nothing
        groupSubItemNumber = ""
        groupMasterItemID = -1
        groupMasterTemplateURL = ""
    End If
    Dim useDCSpecs As Boolean
    useDCSpecs = CBool(rst("ShowDCSpecs") Or rstItem("ShowDCSpecs")) 'item d/c or subitem d/c
    
    
    Dim i As Long, j As Long, iter As Variant
    
    'NON-TABBED TOP REGION
    Dim aboveTabHTML As String
    aboveTabHTML = ""
    If Not useDCSpecs Then
        aboveTabHTML = generateAboveTabsHTML(rst("ItemNumber"), True)
    End If
    
    'WRITEUP
    Dim specsWriteup As String
    specsWriteup = ""
    If Not useDCSpecs Then
        specsWriteup = generateWriteupHTML(Nz(rst("WriteupHTML")), True)
    End If
    
    'HIGHLIGHTS
    Dim specsHighlights As String
    specsHighlights = generateHighlightsHTML( _
        rst:=rst, _
        useDCSpecs:=useDCSpecs, _
        groupSubItemNumber:=groupSubItemNumber, _
        groupMasterItemID:=groupMasterItemID, _
        groupMasterTemplateURL:=groupMasterTemplateURL, _
        includeGroupCategoryLink:=True, _
        includeBestOfferText:=False, _
        wrapInDiv:=True _
      )
    
    'NOTES
    Dim specsNotes As String
    specsNotes = ""
    If Not useDCSpecs Then
        specsNotes = generateNotesHTML(rst("ItemNumber"), Nz(rst("OverrideNotesTabText")), Nz(rst("NotesHTML")), True)
    End If
    
    'SPECIALS
    Dim specsSpecials As String
    specsSpecials = ""
    If Not useDCSpecs Then
        specsSpecials = generateSpecialsHTML(rst("ItemNumber"), True)
    End If
    
    'INCLUDES
    Dim specsIncludes As String
    specsIncludes = ""
    If Not useDCSpecs Then
        If Left(rst("ItemNumber"), 3) = "GFT" Then
            'gift certificates don't get includes
        Else
            specsIncludes = generateIncludesHTML(rst("ItemNumber"), Nz(rst("Desc2")), False, False, True)
        End If
    End If
    
    'FEATURES
    Dim specsFeatures As String
    specsFeatures = ""
    If Not useDCSpecs Then
        specsFeatures = generateGenericSpecBlockHTML(Nz(rst("FeaturesHTML")), True, "features")
    End If
    
    'SPECIFICATIONS
    Dim specsTechspecs As String
    specsTechspecs = ""
    If Not useDCSpecs Then
        specsTechspecs = generateGenericSpecBlockHTML(Nz(rst("TechSpecsHTML")), True, "tech")
    End If
    
    'MEDIA
    Dim specsMedia As String
    specsMedia = ""
    If Not useDCSpecs Then
        specsMedia = generateGenericSpecBlockHTML(Nz(rst("MediaHTML")), True, "media")
    End If
    
    'FAQ - was "additional" before
    Dim specsFAQ As String
    specsFAQ = ""
    If Not useDCSpecs Then
        specsFAQ = generateGenericSpecBlockHTML(Nz(rst("ExtendedSpecsHTML")), True, "faq")
    End If
    
    'TESTIMONIALS
    Dim specsTestimonials As String
    specsTestimonials = ""
    'If Not useDCSpecs Then
    '    specsTestimonials = generateCustomerQuoteText(rst("ItemNumber"))
    'End If
    
    'REVIEWS - tools plus site only
    Dim specsReviews As String
    specsReviews = ""
    If rst("WebsiteID") = 0 And Not useDCSpecs Then
        specsReviews = generateGenericSpecBlockHTML("%%%%REVIEWS%%%%", True, "reviews")
    End If
    
    'QUESTIONS AND ANSWERS - tools plus site only
    Dim specsQA As String
    specsQA = ""
    '2012-08-27 no more q&a tab for now
    'If rst("WebsiteID") = 0 And Not useDCSpecs Then
    '    specsQA = generateGenericSpecBlockHTML("%%%%QA%%%%", True, "q-and-a")
    'End If
    
    
    
    'ALL THE TAB CONTENTS ARE SET UP, BUILD IT
    
    'create arrays for tabbed sections, this is the order that everything appears in
    Dim tabsArray As Variant, namesArray As Variant
    tabsArray = Array(specsHighlights, specsWriteup, specsNotes, specsSpecials, specsIncludes, specsFeatures, specsTechspecs, specsMedia, specsFAQ, specsTestimonials, specsReviews, specsQA)
    namesArray = Array(IIf(useDCSpecs, "Status", "Highlights"), "Description", "Notes", "Specials", "Includes", "Features", "Specs", "Media", "FAQ", "Testimonials", "Reviews", "Q &amp; A")
    
    Dim hasTabs As Boolean
    hasTabs = False
    
    'build the tab control bit now, this doesn't really matter where it happens,
    'since we add it in conditionally
    Dim tabControl As String
    tabControl = "<p id=""tabs-expand""><a href=""#"" onclick=""TP.Tabs.unzip(); return false;""><img src=""http://p1.hostingprod.com/@tools-plus.com/images/tabs-expand.gif"" alt="""">Expand all tabs</a></p>" & vbCrLf
    tabControl = tabControl & "<p id=""tabs-condense""><a href=""#"" onclick=""TP.Tabs.zip(); return false;""><img src=""http://p1.hostingprod.com/@tools-plus.com/images/tabs-condense.gif"" alt="""">Collapse into tabs</a></p>" & vbCrLf
    
    Dim extspec As String
    extspec = ""
    
    'add non-tabbed sections, separated by an <hr>
    If aboveTabHTML <> "" Then
        extspec = extspec & aboveTabHTML & "<hr>" & vbCrLf
    End If
    
    'add placeholder for controls, may or may not be added
    extspec = extspec & Chr(0) & "CONTROL_PLACEHOLDER" & Chr(0)
    
    'build the tabstrip
    'the tabstrip doesn't get vbCrLf's and spacing like everything else,
    'because otherwise Firefox and maybe some other browsers put a space
    'between the elements, causing them to look like they're too far apart.
    Dim tabStrip As String
    tabStrip = ""
    i = 0
    j = 0
    For i = 0 To UBound(tabsArray)
        If tabsArray(i) <> "" Then
            Dim classStr As String
            Select Case namesArray(i)
                Case Is = "Specials"
                    classStr = " class=""highlight-tab"""
                Case Is = "Notes"
                    classStr = " class=""warn-tab"""
                Case Else
                    classStr = ""
            End Select
            tabStrip = tabStrip & "<li" & classStr & "><a href=""#"" onclick=""TP.Tabs.set_active(" & j & "); return false;"">" & namesArray(i) & "</a></li>"
            j = j + 1
        End If
    Next i
    If tabStrip <> "" Then
        extspec = extspec & "<ul id=""tabstrip"" style=""display: none;"">" & tabStrip & "</ul>" & vbCrLf
    End If
    
    'now start the tab part
    extspec = extspec & "<div id=""tabbable-region"">" & vbCrLf
    
    'add each tab, separated by an <hr> (taken out by JS when zipped)
    For Each iter In tabsArray
        If iter <> "" Then
            hasTabs = True
            extspec = extspec & iter & "<hr>" & vbCrLf
        End If
    Next iter
    'remove last hr
    If Right(extspec, 6) = "<hr>" & vbCrLf Then
        extspec = Left(extspec, Len(extspec) - 6)
    End If
    
    extspec = extspec & "</div>" & vbCrLf
    
    'add the controls, if there were any populated tabs
    If hasTabs Then
        extspec = Replace(extspec, Chr(0) & "CONTROL_PLACEHOLDER" & Chr(0), tabControl)
    Else
        extspec = Replace(extspec, Chr(0) & "CONTROL_PLACEHOLDER" & Chr(0), "")
    End If
    
    'add the acronyms in
    'replaceWithAcronyms abbrevList, extspec
    
    'global switch from regular <hr>s to styled <div>s
    extspec = Replace(extspec, "<hr>", "<div class=""spacer-decor""></div>")
    extspec = Replace(extspec, "<hr/>", "<div class=""spacer-decor""></div>")
    extspec = Replace(extspec, "<hr />", "<div class=""spacer-decor""></div>")
    
    extspec = ensurePercentEscapesAreOnOwnLine(extspec)
    
    GenerateItemPageContents = extspec
End Function

Public Function GenerateItemPageContentsForEBay(rstItem As ADODB.Recordset) As String
    'parameter is the recordset for the item or group master item (basically the thing we want the specs from)
    
    'set up group information
    '  reset recordset to use when grouped
    '  grab misc info required
    Dim rst As ADODB.Recordset
    If GroupItemCache_IsGroupedSubItem(rstItem("ItemNumber")) Then
        'need to switch recordsets to the group
        Set rst = GroupItemCache_GetGroupRecordsetBySubItemNumber(rstItem("ItemNumber"))
    Else
        'current recordset is okay
        Set rst = rstItem
    End If
    Dim groupSubItemNumber As String, groupMasterItemID As Long, groupMasterTemplateURL As String
    If GroupItemCache_IsGroupedSubItem(rstItem("ItemNumber")) Then
        'subitems get these filled
        groupSubItemNumber = rstItem("ItemNumber")
        groupMasterItemID = GroupItemCache_GetGroupDatabaseIDByMasterItemNumber(rst("ItemNumber"))
        groupMasterTemplateURL = GroupItemCache_GetGroupYahooIDByMasterItemNumber(rst("ItemNumber"))
    Else
        'regular items and group masters get nothing
        groupSubItemNumber = ""
        groupMasterItemID = -1
        groupMasterTemplateURL = ""
    End If
    
    'no specials, no media, no addl
    
    Dim extspec As PerlArray
    Set extspec = New PerlArray
    
    'WRITEUP
    Dim specsWriteup As String
    specsWriteup = generateWriteupHTML(Nz(rst("WriteupHTML")), False)
    If specsWriteup <> "" Then
        extspec.Push specsWriteup
    End If
    
    'HIGHLIGHTS
    Dim specsHighlights As String
    specsHighlights = generateHighlightsHTML( _
        rst:=rst, _
        useDCSpecs:=False, _
        groupSubItemNumber:=groupSubItemNumber, _
        groupMasterItemID:=groupMasterItemID, _
        groupMasterTemplateURL:=groupMasterTemplateURL, _
        includeGroupCategoryLink:=False, _
        includeBestOfferText:=rstItem("EBayAllowBestOffer"), _
        wrapInDiv:=False, _
        indent:="    " _
      )
    If specsHighlights <> "" Then
        extspec.Push specsHighlights
    End If
    
    'SPECIFICATIONS
    Dim specsTechspecs As String
    specsTechspecs = generateGenericSpecBlockHTML(Nz(rst("TechSpecsHTML")), False, "")
    If specsTechspecs <> "" Then
        extspec.Push specsTechspecs
    End If
    
    'FEATURES
    Dim specsFeatures As String
    specsFeatures = generateGenericSpecBlockHTML(Nz(rst("FeaturesHTML")), False, "")
    If specsFeatures <> "" Then
        extspec.Push specsFeatures
    End If
    
    'NOTES
    Dim specsNotes As String
    specsNotes = generateGenericSpecBlockHTML(Nz(rst("NotesHTML")), False, "")
    If specsNotes <> "" Then
        extspec.Push specsNotes
    End If
    
    'INCLUDES
    Dim specsIncludes As String
    specsIncludes = generateIncludesHTML(rstItem("ItemNumber"), Nz(rstItem("Desc2")), True, False, False)
    If specsIncludes <> "" Then
        extspec.Push specsIncludes
    End If
    
    'join, remove newlines, and return
    If extspec.Scalar = 0 Then
        GenerateItemPageContentsForEBay = ""
    Else
'        'add the footer here. if for some reason we try uploading an item with no
'        'specs (D-ADW9050-2), ebay will error out with "empty desc".
'        Dim ebayFooter As String
'        ebayFooter = ""
'        If rstItem("DropshipOnly") Then
'            ebayFooter = ebayFooter & "<div style=""margin-top:1em; text-align:center;""><img src=""http://p1.hostingprod.com/@tools-plus.com/ebay/ebay-store-table-template-dropship.png"" alt=""""></div>"
'        End If
'        ebayFooter = ebayFooter & "<div style=""margin-top:1em; text-align: center;""><img src=""http://p1.hostingprod.com/@tools-plus.com/ebay/ebay-store-table-template.png"" alt=""""></div>"'
'
'        extspec.Push ebayFooter
'
'        Dim retval As String
'        retval = Replace(extspec.Join("<hr>"), vbCrLf, "")
'        retval = HTMLProcessing.ChangeRelativeURLsToAbsolute(retval, CStr(WebsiteURLHash.item(CLng(rst("WebsiteID")))))

        Dim retval As String
        retval = EBay.EBayGenerateHTMLTemplate(extspec.Join("<hr>"), rstItem)
        retval = Replace(retval, vbCrLf, "")
        retval = Utilities.ChangeRelativeURLsToAbsolute(retval, CStr(WebsiteURLHash.item(CLng(rst("WebsiteID")))))
        
        GenerateItemPageContentsForEBay = retval
    End If
End Function



Private Function generateAboveTabsHTML(item As String, wrapInDiv As Boolean) As String
    Dim aboveTabParts As PerlArray
    Set aboveTabParts = New PerlArray
    If SpecialsCache_GetTextAboveTab(item) <> "" Then
        aboveTabParts.Push SpecialsCache_GetTextAboveTab(item)
    End If
    If aboveTabParts.Scalar = 0 Then
        generateAboveTabsHTML = ""
    ElseIf wrapInDiv = True Then
        generateAboveTabsHTML = "<div id=""item-specs-abovetab"">" & vbCrLf & _
                                aboveTabParts.Join("<hr>" & vbCrLf) & vbCrLf & _
                                "</div>"
    Else
        generateAboveTabsHTML = aboveTabParts.Join("<hr>" & vbCrLf) & vbCrLf
    End If
End Function

Private Function generateWriteupHTML(writeupHTML As String, wrapInDiv As Boolean) As String
    Dim wuParts As PerlArray
    Set wuParts = New PerlArray
    
    If writeupHTML <> "" Then
        wuParts.Push writeupHTML
    End If
    If wuParts.Scalar = 0 Then
        generateWriteupHTML = ""
    ElseIf wrapInDiv Then
        generateWriteupHTML = "<div id=""item-specs-writeup"" class=""tabbable"">" & vbCrLf & _
                              wuParts.Join("<hr>" & vbCrLf) & vbCrLf & _
                              "</div>"
    Else
        generateWriteupHTML = wuParts.Join("<hr>" & vbCrLf) & vbCrLf
    End If
End Function

Private Function generateHighlightsHTML(rst As ADODB.Recordset, useDCSpecs As Boolean, groupSubItemNumber As String, groupMasterItemID As Long, groupMasterTemplateURL As String, includeGroupCategoryLink As Boolean, includeBestOfferText As Boolean, wrapInDiv As Boolean, Optional indent As String = "") As String
    Dim i As Long
    Dim highlights As PerlArray
    Set highlights = New PerlArray
    
    'set up the basic specs in an array
    If useDCSpecs Then
        For i = 1 To 8
            highlights.Push CStr(Trim(FillDCSpecs(rst("ItemNumber"), i)))
        Next i
    Else
        For i = 1 To 8
            highlights.Push CStr(Trim(Nz(rst("Spec" & i & "HTML"))))
        Next i
    End If
    
    'on group items, add the column data "key: value" style to the end, as well as a link back to the group
    If groupSubItemNumber <> "" Then
        If Not useDCSpecs Then
            'push the column info
            'this is the dumbest query ever, but it works, just returns column name + value in consecutive rows
            Dim rstGI As ADODB.Recordset
            Set rstGI = DB.retrieve("SELECT ColumnValue FROM PartNumberGroupLines WHERE (ItemNumber='" & groupSubItemNumber & "' OR ItemNumber='%COLUMN%') AND GroupID=" & groupMasterItemID & " ORDER BY ColumnIndex, ItemNumber")
            While Not rstGI.EOF
                Dim thisSpecName As String, thisSpecValue As String
                thisSpecName = rstGI("ColumnValue")
                rstGI.MoveNext
                thisSpecValue = rstGI("ColumnValue")
                rstGI.MoveNext
                If thisSpecValue = "" Then
                    'skip, these are usually additional info columns that are occasionally blank
                Else
                    highlights.Push thisSpecName & ": " & thisSpecValue
                End If
            Wend
            rstGI.Close
            Set rstGI = Nothing
        End If
        If includeGroupCategoryLink Then
            'highlights.Push "See our other <a href=" & qq(groupMasterTemplateURL & ".html") & ">" & Nz(rst("ManufFullNameWeb")) & " " & Nz(rst("Desc1")) & " " & Nz(rst("Desc3")) & "</a>"
            highlights.Push "See our <a href=" & qq(groupMasterTemplateURL & ".html") & ">" & Nz(rst("ManufFullNameWeb")) & " " & Nz(rst("Desc1")) & " " & Nz(rst("Desc3")) & "</a>"
        End If
    End If
    
    If includeBestOfferText Then
        highlights.Push "<strong>Suggestion:</strong> Make an offer for 2 or more on lower priced items"
    End If
    
    Dim specsHighlights As String
    specsHighlights = ""
    For i = 0 To highlights.Scalar
        If highlights.Elem(i) <> "" Then
            specsHighlights = specsHighlights & indent & IIf(wrapInDiv, " ", "") & " <li>" & highlights.Elem(i) & "</li>" & vbCrLf
        End If
    Next i
    If specsHighlights = "" Then
        generateHighlightsHTML = "" 'weird!
    Else
        specsHighlights = IIf(useDCSpecs, "", indent & IIf(wrapInDiv, " ", "") & "<h3>Highlights:</h3>" & vbCrLf) & _
                          indent & IIf(wrapInDiv, " ", "") & "<ul>" & vbCrLf & _
                          specsHighlights & _
                          indent & IIf(wrapInDiv, " ", "") & "</ul>"
        If wrapInDiv Then
            generateHighlightsHTML = indent & "<div id=""item-specs-highlights"" class=""tabbable"">" & vbCrLf & _
                                     specsHighlights & vbCrLf & _
                                     indent & "</div>" & vbCrLf
        Else
            generateHighlightsHTML = specsHighlights
        End If
    End If
End Function

Private Function generateNotesHTML(item As String, availOverrideNotes As String, notesHTML As String, wrapInDiv As Boolean) As String
    Dim notesParts As PerlArray
    Set notesParts = New PerlArray
    
    ' 2012-05-15 truck freight and common carrier info moved to pit
    'If rst("ShippingType") = 2 Then
    '    notesParts.Push "<div id=""truck-freight-disclaimer"">" & vbCrLf & _
    '                    "%%%%COMMONCARRIER%%%%" & vbCrLf & _
    '                    "</div>" & vbCrLf
    'ElseIf rst("ShippingType") = 1 Then
    '    notesParts.Push "<div id=""truck-freight-disclaimer"">" & vbCrLf & _
    '                    "%%%%TRUCKFREIGHT%%%%" & vbCrLf & _
    '                    "</div>" & vbCrLf
    'End If
    ' 2009-12-24 recon disclaimer is moved to popup in pit
    'If rst("IsReconditioned") = "Yes" Then
    '    notesParts.Push "<div id=""reconditioned-disclaimer"">" & vbCrLf & _
    '                    "%%%%RECON%%%%" & vbCrLf & _
    '                    "</div>" & vbCrLf
    'End If
    If SpecialsCache_GetTextInfoTab(item) <> "" Then
        notesParts.Push "<div class=""notes-specials"">" & vbCrLf & SpecialsCache_GetTextInfoTab(item) & vbCrLf & "</div>" & vbCrLf
    End If
    If availOverrideNotes <> "" Then
        notesParts.Push "<div class=""notes-availability"">" & vbCrLf & availOverrideNotes & vbCrLf & "</div>" & vbCrLf
    End If
    If notesHTML <> "" Then
        notesParts.Push notesHTML
    End If
    'Dim notesGiftPart As String
    'notesGiftPart = getGiftWarningNotePart(item)
    'If notesGiftPart <> "" Then
    '    notesParts.Push notesGiftPart
    'End If
    
    If notesParts.Scalar = 0 Then
        generateNotesHTML = ""
    ElseIf wrapInDiv Then
        generateNotesHTML = "<div id=""item-specs-notes"" class=""tabbable"">" & vbCrLf & _
                            notesParts.Join("<hr>" & vbCrLf) & vbCrLf & _
                            "</div>"
    Else
        generateNotesHTML = notesParts.Join("<hr>" & vbCrLf) & vbCrLf
    End If
End Function

Private Function generateSpecialsHTML(item As String, wrapInDiv As Boolean) As String
    If SpecialsCache_GetTextSpecialsTab(item) = "" Then
        generateSpecialsHTML = ""
    ElseIf wrapInDiv Then
        generateSpecialsHTML = "<div id=""item-specs-specials"" class=""tabbable"">" & vbCrLf & _
                               SpecialsCache_GetTextSpecialsTab(item) & vbCrLf & _
                               "</div>" & vbCrLf
    Else
        generateSpecialsHTML = SpecialsCache_GetTextSpecialsTab(item) & vbCrLf
    End If
End Function

Private Function generateIncludesHTML(item As String, desc2 As String, stripIncludeLineHTML As Boolean, onlyWhenMoreThanOneLine As Boolean, wrapInDiv As Boolean) As String
    'description 2 should always be filled, if not then this breaks, so
    'we'll check for it right here
    If desc2 = "" Then
        generateIncludesHTML = ""
        Exit Function
    End If
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT IncludeText FROM PartNumberIncludes WHERE ItemNumber='" & item & "' ORDER BY SortOrder")
    Dim retval As String
    retval = ""
    If rst.RecordCount >= 2 Or onlyWhenMoreThanOneLine = False Then
        While Not rst.EOF
            If Not stripIncludeLineHTML Then
                retval = retval & "  <li>" & rst("IncludeText") & "</li>" & vbCrLf
            Else
                retval = retval & "  <li>" & StripHTML(rst("IncludeText"), True) & "</li>" & vbCrLf
            End If
            rst.MoveNext
        Wend
    End If
    rst.Close
    Set rst = Nothing
    If retval = "" Then
        generateIncludesHTML = ""
    Else
        retval = " <h3>The " & desc2 & " includes:</h3>" & vbCrLf & _
                 " <ul>" & vbCrLf & _
                 retval & _
                 " </ul>" & vbCrLf
        If wrapInDiv Then
            generateIncludesHTML = "<div id=""item-specs-includes"" class=""tabbable"">" & vbCrLf & _
                                   retval & _
                                   "</div>" & vbCrLf
        Else
            generateIncludesHTML = retval
        End If
    End If
End Function

Private Function generateGenericSpecBlockHTML(blockHTML As String, wrapInDiv As Boolean, wrapID As String) As String
    If blockHTML = "" Then
        generateGenericSpecBlockHTML = ""
    ElseIf wrapInDiv Then
        generateGenericSpecBlockHTML = "<div id=""item-specs-" & wrapID & """ class=""tabbable"">" & vbCrLf & blockHTML & vbCrLf & "</div>" & vbCrLf
    Else
        generateGenericSpecBlockHTML = blockHTML
    End If
End Function

Private Function ensurePercentEscapesAreOnOwnLine(extspec As String) As String
    'this is fucking disgusting..probably easier with a regex
    'since the rtml template parses a line at a time, we need to make sure that
    'any %%%%FOO%%%% escapes are on a line by themselves, and not in the middle
    'of anything. run through the input, checking through for %%%% starters.
    'make sure that the prior character is a cr/lf. if it isn't, then add one
    'and make sure that a space precedes it, since the template doesn't handle
    'that correctly either (i should fix that). then find the ending %%%%, and
    'repeat.
    Dim curptr As Long
    curptr = InStr(1, extspec, "%%%%")
    While curptr <> 0
        If curptr = 1 Then
            curptr = 2
        Else
            Select Case Mid(extspec, curptr - 1, 1)
                Case Is = " "
                    If curptr > 2 Then
                        extspec = Left(extspec, curptr - 2) & "&nbsp;" & vbCrLf & Mid(extspec, curptr)
                        curptr = curptr + 8
                    Else
                        extspec = Left(extspec, curptr - 1) & "&nbsp;" & vbCrLf & Mid(extspec, curptr)
                        curptr = curptr + 8
                    End If
                Case Is = vbLf
                    'do nothing
                    curptr = curptr + 1
                Case Else
                    extspec = Left(extspec, curptr - 1) & vbCrLf & Mid(extspec, curptr)
                    curptr = curptr + 3
            End Select
        End If
        
        curptr = InStr(curptr, extspec, "%%%%") ' set to end
        If curptr <> 0 Then
            If curptr + 4 > Len(extspec) Then
                'do nothing
            Else
                'do something else
                Select Case Mid(extspec, curptr + 4, 1)
                    Case Is = " "
                        extspec = Left(extspec, curptr + 3) & vbCrLf & "&nbsp;" & Mid(extspec, curptr + 5)
                    Case Is = vbCr
                        'do nothing
                    Case Else
                        extspec = Left(extspec, curptr + 3) & vbCrLf & Mid(extspec, curptr + 4)
                End Select
            End If
            
            curptr = InStr(curptr + 1, extspec, "%%%%")
        Else
            'error, oops
        End If
    Wend
    
    ensurePercentEscapesAreOnOwnLine = extspec
End Function



Public Function GroupItemCache_Initialize() As Boolean
    If Not groupedItemsH Is Nothing Then
        GroupItemCache_Clear
    End If
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, ItemNumber, GroupPageName, ItemNumberPointer, ShowItemThumbnails, ShowAvailabilityColumn FROM vWebOffloadActiveGroupedItems")
    Set groupedItemsH = New Dictionary
    Set templateItemsH = New Dictionary
    
    While Not rst.EOF
        groupedItemsH.Add CStr(rst("ItemNumber")), CStr(rst("ItemNumberPointer"))
        If Not templateItemsH.exists(CStr(rst("ItemNumberPointer"))) Then
            templateItemsH.Add CStr(rst("ItemNumberPointer")), Array(Nothing, CLng(rst("ID")), CStr(rst("GroupPageName")), CBool(rst("ShowItemThumbnails")), CBool(rst("ShowAvailabilityColumn")))
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    GroupItemCache_Initialize = True
End Function

Public Function GroupItemCache_Clear() As Boolean
    Set groupedItemsH = Nothing
    If Not templateItemsH Is Nothing Then
        Dim groupiter As Variant
        For Each groupiter In templateItemsH.keys
            If Not templateItemsH.item(CStr(groupiter))(0) Is Nothing Then
                templateItemsH.item(CStr(groupiter))(0).Close
            End If
        Next groupiter
        Set templateItemsH = Nothing
    End If
    GroupItemCache_Clear = True
End Function

Public Function GroupItemCache_IsGroupedSubItem(item As String) As Boolean
    If groupedItemsH Is Nothing Then
        GroupItemCache_Initialize
    End If
    GroupItemCache_IsGroupedSubItem = groupedItemsH.exists(item)
End Function

Public Function GroupItemCache_IsGroupMasterItem(item As String) As Boolean
    If groupedItemsH Is Nothing Then
        GroupItemCache_Initialize
    End If
    GroupItemCache_IsGroupMasterItem = templateItemsH.exists(item)
End Function

Public Function GroupItemCache_GetMasterItemNumberForSubItem(item As String) As String
    If groupedItemsH Is Nothing Then
        GroupItemCache_Initialize
    End If
    If groupedItemsH.exists(item) Then
        GroupItemCache_GetMasterItemNumberForSubItem = groupedItemsH.item(item)
    Else
        GroupItemCache_GetMasterItemNumberForSubItem = ""
    End If
End Function

Public Function GroupItemCache_GetGroupRecordsetByMasterItemNumber(item As String) As ADODB.Recordset
    If groupedItemsH Is Nothing Then
        GroupItemCache_Initialize
    End If
    If templateItemsH.exists(item) Then
        Dim tempA As Variant
        tempA = templateItemsH.item(item)
        If templateItemsH.item(item)(0) Is Nothing Then
            Set tempA(0) = DB.retrieve("SELECT * FROM vWebOffload WHERE ItemNumber='" & item & "'")
            templateItemsH.item(item) = tempA
        End If
        Set GroupItemCache_GetGroupRecordsetByMasterItemNumber = tempA(0)
    Else
        Set GroupItemCache_GetGroupRecordsetByMasterItemNumber = Nothing
    End If
End Function

Public Function GroupItemCache_GetGroupDatabaseIDByMasterItemNumber(item As String) As Long
    If groupedItemsH Is Nothing Then
        GroupItemCache_Initialize
    End If
    If templateItemsH.exists(item) Then
        GroupItemCache_GetGroupDatabaseIDByMasterItemNumber = templateItemsH.item(item)(1)
    Else
        GroupItemCache_GetGroupDatabaseIDByMasterItemNumber = -1
    End If
End Function

Public Function GroupItemCache_GetGroupYahooIDByMasterItemNumber(item As String) As String
    If groupedItemsH Is Nothing Then
        GroupItemCache_Initialize
    End If
    If templateItemsH.exists(item) Then
        GroupItemCache_GetGroupYahooIDByMasterItemNumber = templateItemsH.item(item)(2)
    Else
        GroupItemCache_GetGroupYahooIDByMasterItemNumber = ""
    End If
End Function

Public Function GroupItemCache_GetGroupThumbnailSettingByMasterItemNumber(item As String) As Boolean
    If groupedItemsH Is Nothing Then
        GroupItemCache_Initialize
    End If
    If templateItemsH.exists(item) Then
        GroupItemCache_GetGroupThumbnailSettingByMasterItemNumber = templateItemsH.item(item)(3)
    Else
        GroupItemCache_GetGroupThumbnailSettingByMasterItemNumber = ""
    End If
End Function

Public Function GroupItemCache_GetGroupAvailabilityColumnSettingByMasterItemNumber(item As String) As Boolean
    If groupedItemsH Is Nothing Then
        GroupItemCache_Initialize
    End If
    If templateItemsH.exists(item) Then
        GroupItemCache_GetGroupAvailabilityColumnSettingByMasterItemNumber = templateItemsH.item(item)(4)
    Else
        GroupItemCache_GetGroupAvailabilityColumnSettingByMasterItemNumber = ""
    End If
End Function

Public Function GroupItemCache_GetGroupRecordsetBySubItemNumber(item As String) As ADODB.Recordset
    Set GroupItemCache_GetGroupRecordsetBySubItemNumber = GroupItemCache_GetGroupRecordsetByMasterItemNumber(GroupItemCache_GetMasterItemNumberForSubItem(item))
End Function

Public Function GroupItemCache_GetGroupDatabaseIDBySubItemNumber(item As String) As Long
    GroupItemCache_GetGroupDatabaseIDBySubItemNumber = GroupItemCache_GetGroupDatabaseIDByMasterItemNumber(GroupItemCache_GetMasterItemNumberForSubItem(item))
End Function

Public Function GroupItemCache_GetGroupYahooIDBySubItemNumber(item As String) As String
    GroupItemCache_GetGroupYahooIDBySubItemNumber = GroupItemCache_GetGroupYahooIDByMasterItemNumber(GroupItemCache_GetMasterItemNumberForSubItem(item))
End Function

Public Function GroupItemCache_GetGroupThumbnailSettingBySubItemNumber(item As String) As Boolean
    GroupItemCache_GetGroupThumbnailSettingBySubItemNumber = GroupItemCache_GetGroupThumbnailSettingByMasterItemNumber(GroupItemCache_GetMasterItemNumberForSubItem(item))
End Function

Public Function GroupItemCache_GetGroupAvailabilityColumnSettingBySubItemNumber(item As String) As Boolean
    GroupItemCache_GetGroupAvailabilityColumnSettingBySubItemNumber = GroupItemCache_GetGroupAvailabilityColumnSettingByMasterItemNumber(GroupItemCache_GetMasterItemNumberForSubItem(item))
End Function




Public Function SpecialsCache_Initialize(Optional singleItem As String = "") As Boolean
    If Not specialsItemsH Is Nothing Then
        SpecialsCache_Clear
    End If
    
    Dim iter As Variant, i As Long
    
    Dim warnings As PerlArray
    Set warnings = New PerlArray
    
    Dim warnedAbout As Dictionary
    Set warnedAbout = New Dictionary
    
    Dim thisArray As Variant
    
    Dim tempSectionImgUrl As Dictionary
    Set tempSectionImgUrl = New Dictionary
    Dim tempSectionImgUrlEntry As PerlArray
    
    Set specialsItemsH = New Dictionary
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber, WebsiteID, SpecialName, SpecialText, SectionImage, SectionURL, FlatPercentDiscount, RebateAmount, InstantRebate, ShowOnTab FROM vSpecialsActive" & IIf(singleItem = "", "", " WHERE ItemNumber='" & singleItem & "'"))
    
    While Not rst.EOF
        'if there's a section url and no image, warn
        If Nz(rst("SectionImage")) = "" And Nz(rst("SectionURL")) <> "" Then
            If Not warnedAbout.exists("sectionurlnoimg-" & CStr(rst("SpecialName"))) Then
                warnings.Push "Special " & qq(rst("SpecialName")) & " has a section url, but no section image!"
                warnedAbout.Add "sectionurlnoimg-" & CStr(rst("SpecialName")), True
            End If
        End If
        'if we're showing on a tab and have no text, warn
        If rst("ShowOnTab") <> 0 And Nz(rst("SpecialText")) = "" Then
            If Not warnedAbout.exists("textempty-" & CStr(rst("SpecialName"))) Then
                warnings.Push "Special " & qq(rst("SpecialName")) & " has no text!"
                warnedAbout.Add "textempty-" & CStr(rst("SpecialName")), True
            End If
        End If
        
        If specialsItemsH.exists(CStr(rst("ItemNumber"))) Then
            thisArray = specialsItemsH.item(CStr(rst("ItemNumber")))
        Else
            'initialize with defaults
            'defaults are repeated later on here when checking for expired specials
            'and for each getter when item doesn't exist
            thisArray = Array("", 0, 0, False, "", "", "")
        End If
        
        If Nz(rst("SectionImage")) <> "" Then
            ''check for multiple section images, and warn
            ''TODO: the rtml for generating this is a mess, and would be much easier just uploading html here, and
            ''then we could do multiple very easily. same as tabs, basically.
            'If thisArray(0) <> "" Then
            '    If Not warnedAbout.exists("sectionimagemult-" & CStr(rst("ItemNumber"))) Then
            '        warnings.Push "Item " & qq(rst("ItemNumber")) & " has multiple section images! Only using the first one!"
            '        warnedAbout.Add "sectionimagemult-" & CStr(rst("ItemNumber")), True
            '    End If
            'Else
            '    thisArray(0) = CStr(rst("SectionImage"))
            '    thisArray(1) = Nz(rst("SectionURL"))
            'End If
            If tempSectionImgUrl.exists(CStr(rst("ItemNumber"))) Then
                Set tempSectionImgUrlEntry = tempSectionImgUrl.item(CStr(rst("ItemNumber")))
                tempSectionImgUrlEntry.Push Array(CStr(Nz(rst("SectionImage"))), CStr(Nz(rst("SectionURL"))))
                Set tempSectionImgUrl.item(CStr(rst("ItemNumber"))) = tempSectionImgUrlEntry
            Else
                Set tempSectionImgUrlEntry = New PerlArray
                tempSectionImgUrlEntry.Push Array(CStr(Nz(rst("SectionImage"))), CStr(Nz(rst("SectionURL"))))
                tempSectionImgUrl.Add CStr(rst("ItemNumber")), tempSectionImgUrlEntry
            End If
        End If
        
        If rst("FlatPercentDiscount") <> 0 Then
            'check for multiple sale discounts. this is really bad! (but probably never going to happen)
            If thisArray(1) <> 0 Then
                If Not warnedAbout.exists("salemult-" & CStr(rst("ItemNumber"))) Then
                    warnings.Push "Item " & qq(rst("ItemNumber")) & " has multiple % sale discounts! Only using the first one!"
                    warnedAbout.Add "salemult-" & CStr(rst("ItemNumber")), True
                End If
            Else
                thisArray(1) = CLng(rst("FlatPercentDiscount"))
            End If
        End If
        
        If rst("RebateAmount") <> 0 Then
            'check for multiple rebates. we could probably do multiple if they're all the same type (instant/mailin)?
            If thisArray(2) <> 0 Then
                If Not warnedAbout.exists("rebatemult-" & CStr(rst("ItemNumber"))) Then
                    warnings.Push "Item " & qq(rst("ItemNumber")) & " has multiple rebates! Only using the first one!"
                    warnedAbout.Add "rebatemult-" & CStr(rst("ItemNumber")), True
                End If
            Else
                thisArray(2) = CDec(rst("RebateAmount"))
                thisArray(3) = CBool(rst("InstantRebate"))
            End If
        End If
        
        'map radio button/db entry to hash entry
        Dim tabIndex As Long
        Select Case rst("ShowOnTab")
            Case Is = 0
                tabIndex = -1 'do not show
            Case Is = 1
                tabIndex = 4  'specials tab
            Case Is = 2
                tabIndex = 5  'info tab
            Case Is = 3
                tabIndex = 6  'above tab control
        End Select
        
        'add text, splitting multiple with hr if required
        If tabIndex <> -1 And Nz(rst("SpecialText")) <> "" Then
            If thisArray(tabIndex) = "" Then
                thisArray(tabIndex) = rst("SpecialText")
            Else
                thisArray(tabIndex) = thisArray(tabIndex) & vbCrLf & "<hr>" & vbCrLf & rst("SpecialText")
            End If
        End If
        
        'and then re-add the array to the hash
        If specialsItemsH.exists(CStr(rst("ItemNumber"))) Then
            specialsItemsH.item(CStr(rst("ItemNumber"))) = thisArray
        Else
            specialsItemsH.Add CStr(rst("ItemNumber")), thisArray
        End If
        
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    For Each iter In tempSectionImgUrl.keys
        Dim stanza As String
        stanza = ""
        For i = 0 To tempSectionImgUrl.item(CStr(iter)).Scalar - 1
            Dim thisImg As String, thisUrl As String
            thisImg = tempSectionImgUrl.item(CStr(iter)).Elem(i)(0)
            thisUrl = tempSectionImgUrl.item(CStr(iter)).Elem(i)(1)
            If thisUrl = "" Then
                thisUrl = "DEFAULT"
            End If
            stanza = stanza & thisImg & "~" & thisUrl & vbCrLf
        Next i
        
        thisArray = specialsItemsH.item(CStr(iter))
        thisArray(0) = stanza
        specialsItemsH.item(CStr(iter)) = thisArray
    Next iter
    
    Set rst = DB.retrieve("SELECT ItemNumber, SpecialText FROM vSpecialsExpired WHERE SpecialText IS NOT NULL" & IIf(singleItem = "", "", " AND ItemNumber='" & singleItem & "'"))
    While Not rst.EOF
        If specialsItemsH.exists(CStr(rst("ItemNumber"))) Then
            thisArray = specialsItemsH.item(CStr(rst("ItemNumber")))
        Else
            thisArray = Array("", 0, 0, False, "", "", "")
        End If
        
        If thisArray(5) = "" Then 'info tab
            thisArray(5) = rst("SpecialText")
        Else
            thisArray(5) = thisArray(5) & vbCrLf & "<hr>" & vbCrLf & rst("SpecialText")
        End If
        
        If specialsItemsH.exists(CStr(rst("ItemNumber"))) Then
            specialsItemsH.item(CStr(rst("ItemNumber"))) = thisArray
        Else
            specialsItemsH.Add CStr(rst("ItemNumber")), thisArray
        End If
        
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    If warnings.Scalar > 0 Then
        MsgBox warnings.Join(vbCrLf)
    End If
    Set warnings = Nothing
    Set warnedAbout = Nothing
    
    SpecialsCache_Initialize = True
End Function

Public Function SpecialsCache_Clear() As Boolean
    If Not specialsItemsH Is Nothing Then
        Set specialsItemsH = Nothing
    End If
End Function

'Public Function SpecialsCache_GetSectionImage(item As String) As String
'    If specialsItemsH Is Nothing Then
'        SpecialsCache_Initialize
'    End If
'
'    If specialsItemsH.exists(item) Then
'        SpecialsCache_GetSectionImage = CStr(specialsItemsH.item(item)(0))
'    Else
'        SpecialsCache_GetSectionImage = ""
'    End If
'End Function
'
'Public Function SpecialsCache_GetSectionURL(item As String) As String
'    If specialsItemsH Is Nothing Then
'        SpecialsCache_Initialize
'    End If
'
'    If specialsItemsH.exists(item) Then
'        SpecialsCache_GetSectionURL = CStr(specialsItemsH.item(item)(1))
'    Else
'        SpecialsCache_GetSectionURL = ""
'    End If
'End Function

Public Function SpecialsCache_GetSectionSpecialStanza(item As String) As String
    If specialsItemsH Is Nothing Then
        SpecialsCache_Initialize
    End If
    
    If specialsItemsH.exists(item) Then
        SpecialsCache_GetSectionSpecialStanza = CStr(specialsItemsH.item(item)(0))
    Else
        SpecialsCache_GetSectionSpecialStanza = ""
    End If
End Function

Public Function SpecialsCache_GetFlatPercentDiscount(item As String) As Long
    If specialsItemsH Is Nothing Then
        SpecialsCache_Initialize
    End If
    
    If specialsItemsH.exists(item) Then
        SpecialsCache_GetFlatPercentDiscount = CLng(specialsItemsH.item(item)(1))
    Else
        SpecialsCache_GetFlatPercentDiscount = 0
    End If
End Function

Public Function SpecialsCache_GetRebateAmount(item As String) As Variant
    If specialsItemsH Is Nothing Then
        SpecialsCache_Initialize
    End If
    
    If specialsItemsH.exists(item) Then
        SpecialsCache_GetRebateAmount = CDec(specialsItemsH.item(item)(2))
    Else
        SpecialsCache_GetRebateAmount = 0
    End If
End Function

Public Function SpecialsCache_GetRebateInstantFlag(item As String) As Boolean
    If specialsItemsH Is Nothing Then
        SpecialsCache_Initialize
    End If
    
    If specialsItemsH.exists(item) Then
        SpecialsCache_GetRebateInstantFlag = CBool(specialsItemsH.item(item)(3))
    Else
        SpecialsCache_GetRebateInstantFlag = False 'null?
    End If
End Function

Public Function SpecialsCache_GetTextSpecialsTab(item As String) As String
    If specialsItemsH Is Nothing Then
        SpecialsCache_Initialize
    End If
    
    If specialsItemsH.exists(item) Then
        SpecialsCache_GetTextSpecialsTab = CStr(specialsItemsH.item(item)(4))
    Else
        SpecialsCache_GetTextSpecialsTab = ""
    End If
End Function

Public Function SpecialsCache_GetTextInfoTab(item As String) As String
    If specialsItemsH Is Nothing Then
        SpecialsCache_Initialize
    End If
    
    If specialsItemsH.exists(item) Then
        SpecialsCache_GetTextInfoTab = CStr(specialsItemsH.item(item)(5))
    Else
        SpecialsCache_GetTextInfoTab = ""
    End If
End Function

Public Function SpecialsCache_GetTextAboveTab(item As String) As String
    If specialsItemsH Is Nothing Then
        SpecialsCache_Initialize
    End If
    
    If specialsItemsH.exists(item) Then
        SpecialsCache_GetTextAboveTab = CStr(specialsItemsH.item(item)(6))
    Else
        SpecialsCache_GetTextAboveTab = ""
    End If
End Function







'Private Function getGiftWarningNotePart(item As String, itemID as String) As String
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT SUM(GiftNotConcealableWarning) FROM InventoryComponents WHERE ID IN (SELECT ComponentID FROM InventoryComponentMap WHERE ItemID=" & itemID & ")")
'    If Nz(rst(0), "0") = "0" Then
'        getGiftWarningNotePart = ""
'    Else
'        getGiftWarningNotePart = "<p>If purchasing this item as a gift, please note: This item may ship in its original packaging and can not be re-wrapped or concealed.</p>"
'    End If
'    rst.Close
'    Set rst = Nothing
'End Function

'Private Sub replaceWithAcronyms(abbrevList As Dictionary, ParamArray listoffields())
'    Dim i As Long, j As Long, k As Long
'    Dim thisabbrev As String, thisfield As String, dienow As Boolean
'
'    Dim onlyReplFirstInstance As Boolean
'    onlyReplFirstInstance = True
'
'    'Dim listoffields(0) As Variant
'    'listoffields(0) = "foo (ft-lbs) bar"
'    'Set abbrevList = New Dictionary
'    'abbrevList.Add "ft-lbs", "Foot-Pounds, a measure of torque"
'
'    Dim re As RegExp
'    Dim foundabbrev As MatchCollection
'    Set re = New RegExp
'    re.IgnoreCase = False
'    're.Global = True
'    re.Global = False
'
'    For i = LBound(abbrevList.keys) To UBound(abbrevList.keys)
'        thisabbrev = abbrevList.keys(i)
'        re.Pattern = thisabbrev 'already includes boundaries from cacheAcronyms fn
'        For j = LBound(listoffields) To UBound(listoffields)
'            dienow = False
'            thisfield = listoffields(j)
'            If re.Test(thisfield) Then
'                Set foundabbrev = re.Execute(thisfield)
'                'For k = 0 To foundabbrev.count - 1
'                k = 0
'                    'check for acronyms inside acronyms
'                    'if it finds one, it should probably continue and see if there's any more later in the line
'                    'no idea how to do that though, re.replace doesn't have other parameters
'                    Dim openpos As Long, closepos As Long
'                    If foundabbrev.item(k).FirstIndex = 0 Then
'                        openpos = 0
'                    Else
'                        openpos = InStrRev(thisfield, "<span class=""info"">", foundabbrev.item(k).FirstIndex)
'                    End If
'                    If openpos <> 0 Then
'                        closepos = InStrRev(thisfield, "</span>", foundabbrev.item(k).FirstIndex)
'                        If closepos <> 0 And closepos > openpos Then
'                            'ok, the acronym closes before this
'                        Else
'                            dienow = True
'                        End If
'                    End If
'                    'now check for if we're in a tag (i.e., "div id=something", don't replace the id)
'                    If foundabbrev.item(k).FirstIndex = 0 Then
'                        openpos = 0
'                    Else
'                        openpos = InStrRev(thisfield, "<", foundabbrev.item(k).FirstIndex)
'                    End If
'                    If openpos <> 0 Then
'                        closepos = InStrRev(thisfield, ">", foundabbrev.item(k).FirstIndex)
'                        If closepos <> 0 And closepos > openpos Then
'                            'ok
'                        Else
'                            dienow = True
'                        End If
'                    End If
'                'Next k
'                If Not dienow Then
'                    If foundabbrev.item(0).SubMatches.count > 0 Then
'                        listoffields(j) = re.Replace(thisfield, "<span class=""info"">" & Left(foundabbrev.item(0), Len(foundabbrev.item(0)) - 1) & "<span> " & abbrevList.item(thisabbrev) & "</span></span>" & foundabbrev.item(0).SubMatches(0))
'                    Else
'                        listoffields(j) = re.Replace(thisfield, "<span class=""info"">" & foundabbrev.item(0) & "<span> " & abbrevList.item(thisabbrev) & "</span></span>")
'                    End If
'                    If onlyReplFirstInstance Then
'                        j = UBound(listoffields)
'                    End If
'                End If
'            Else
'                'no match
'            End If
'        Next j
'    Next i
'End Sub

'Public Function GetItemGroupList(groupPageID As String) As String
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT dbo.CreateYahooID(ItemNumber) AS YahooID, Description FROM PartNumberGroupItems WHERE GroupMasterID=" & groupPageID & " ORDER BY SortOrder")
'    Dim retval As String
'    retval = ""
'    While Not rst.EOF
'        retval = IIf(retval = "", "", retval & vbCrLf) & rst("YahooID") & vbCrLf & rst("Description")
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'
'    GetItemGroupList = retval
'End Function

'Public Function GetRBCategoryContentsAsString(item As String) As String
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT dbo.CreateYahooID(Accessory) AS Accessory FROM PartNumberAccessories WHERE ItemNumber='" & item & "'")
'    Dim retval As String
'    While Not rst.EOF
'        retval = retval & " " & rst("Accessory")
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'    GetRBCategoryContentsAsString = Trim(retval)
'End Function

'Public Function GetItemSimilarItemsAsString(item As String) As String
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT dbo.CreateYahooID(Accessory) AS Accessory FROM PartNumberAccessories WHERE ItemNumber='" & item & "' AND DirectAccessory=1")
'    Dim retval As String
'    While Not rst.EOF
'        retval = retval & " " & rst("Accessory")
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'    GetItemSimilarItemsAsString = Trim(retval)
'End Function

'Public Function WebPath1sAreOK() As Boolean
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT ItemNumber FROM vWebOffload WHERE BaseItem IS NULL AND ItemNumber NOT IN (SELECT ItemNumber FROM PartNumberWebPaths)")
'    If Not rst.EOF Then
'        Open "c:\badrecs.lst" For Output As #1
'        Print #1, Now() & vbCrLf & "The following items do not have any web path 1's and are marked published:"
'        While Not rst.EOF
'            Print #1, rst("ItemNumber")
'            rst.MoveNext
'        Wend
'        Close #1
'        MsgBox "Found records marked for publish without any web paths!"
'        Shell "notepad c:\badrecs.lst", vbNormalFocus
'        WebPath1sAreOK = False
'    Else
'        WebPath1sAreOK = True
'    End If
'    rst.Close
'    Set rst = Nothing
'End Function

'Public Function ManufNamesAreOK() As Boolean
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT ProductLine FROM ProductLine WHERE ProductLine IN (SELECT ProductLine FROM vWebOffload) AND (COALESCE(ManufFullNameCleaned,'')='' OR COALESCE(ManufFullNameWeb,'')='')")
'    If Not rst.EOF Then
'        Dim msg As String
'        msg = "The following product lines are missing a manufacturer name cleaned or web (or both):" & vbCrLf
'        While Not rst.EOF
'            msg = msg & vbCrLf & rst("ProductLine")
'            rst.MoveNext
'        Wend
'        msg = msg & vbCrLf & vbCrLf & "You need to fix this before offloading!"
'        MsgBox msg
'        ManufNamesAreOK = False
'    Else
'        ManufNamesAreOK = True
'    End If
'    rst.Close
'    Set rst = Nothing
'End Function

'Private Function generateCustomerQuoteText(item As String, Optional xhtml As Boolean = False) As String
'    Dim rst As ADODB.Recordset
'    'Set rst = DB.retrieve("SELECT TOP 1 FirstName, City, State, Country, Quote, URL, Desc2, Desc3 FROM CustomerQuoteItems INNER JOIN CustomerQuotes ON CustomerQuoteItems.QuoteID=CustomerQuotes.ID INNER JOIN PartNumbers ON CustomerQuoteItems.ItemNumber=PartNumbers.ItemNumber WHERE CustomerQuoteItems.ItemNumber='" & item & "' AND CustomerQuoteItems.RunFrom<>-2 AND CustomerQuoteItems.RunTo<>-2 ORDER BY NEWID()")
'    Set rst = DB.retrieve("SELECT FirstName, City, State, Country, Quote, URL, Desc2, Desc3 FROM CustomerQuoteItems INNER JOIN CustomerQuotes ON CustomerQuoteItems.QuoteID=CustomerQuotes.ID INNER JOIN PartNumbers ON CustomerQuoteItems.ItemNumber=PartNumbers.ItemNumber WHERE CustomerQuoteItems.ItemNumber='" & item & "' AND CustomerQuotes.RunFrom<>-2 AND CustomerQuotes.RunTo<>-2 ORDER BY NEWID()")
'    If rst.EOF Then
'        generateCustomerQuoteText = ""
'    Else
'        Dim retval As String
'        retval = ""
'        While Not rst.EOF
'            retval = retval & " <p>" & IIf(Nz(rst("URL")) = "", "", "<a href=""" & Nz(rst("URL")) & """>") & escapeHTMLEntities(rst("FirstName")) & IIf(Nz(rst("URL")) = "", "", "</a>") & ", from " & rst("City")
'            If Nz(rst("State")) <> "" Then
'                retval = retval & ", " & rst("State")
'            End If
'            If Nz(rst("Country")) <> "" And Nz(rst("Country")) <> "US" Then
'                retval = retval & ", " & rst("Country")
'            End If
'            retval = retval & " bought the " & rst("Desc2") & " " & rst("Desc3") & ". Here's what was said about it:</p>" & vbCrLf
'            retval = retval & " <blockquote style=""margin-left: 1em;"">"
'            Dim paras As Variant
'            paras = Split(rst("Quote"), vbCrLf)
'            Dim i As Long
'            For i = 0 To UBound(paras)
'                If paras(i) <> "" Then
'                    retval = retval & "<p>" & escapeHTMLEntities(paras(i)) & "</p>"
'                End If
'            Next i
'            retval = retval & "</blockquote>" & vbCrLf
'            rst.MoveNext
'            If Not rst.EOF Then
'                retval = retval & " " & IIf(xhtml, "<hr />", "<hr>") & vbCrLf
'            End If
'        Wend
'        generateCustomerQuoteText = "<div id=""item-specs-testimonial"" class=""tabbable"">" & vbCrLf & retval & "</div>" & vbCrLf '& IIf(xhtml, "<hr />", "<hr>") & vbCrLf
'    End If
'    rst.Close
'    Set rst = Nothing
'End Function
