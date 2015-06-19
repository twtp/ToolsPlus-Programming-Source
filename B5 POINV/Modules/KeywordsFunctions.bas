Attribute VB_Name = "KeywordsFunctions"
Option Explicit

Public Function KwdGetServiceIDFromName(svcName As String) As String
    KwdGetServiceIDFromName = DLookup("ID", "KeywordsEngines", "Engine='" & EscapeSQuotes(svcName) & "'")
End Function

Public Function GetKwdStatusForPhrase(phrase As String, svc As String) As String
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Active, Seasonal, Pending, Rejected, DoNotUse FROM vKeywords WHERE SearchPhrase='" & EscapeSQuotes(phrase) & "' AND Engine='" & svc & "'")
    GetKwdStatusForPhrase = GetKwdStatus(phrase, rst("Active"), rst("Seasonal"), rst("Pending"), rst("Rejected"), rst("DoNotUse"))
    rst.Close
    Set rst = Nothing
End Function

Public Function GetKwdStatus(phrase As String, active As Boolean, seasonal As Boolean, pending As Boolean, rejected As Boolean, donotuse As Long) As String
    'don't change these return values, keywordsphrasedetails.fillForm requires these (maybe more)
    'similar function in perl script to handle clearing export flags, should be changed if this does.
    If active And Not pending And Not rejected And Not CBool(donotuse) Then
        If seasonal Then
            Dim rst As ADODB.Recordset
            Set rst = DB.retrieve("SELECT '1' FROM KeywordsSeasonalActiveMonths WHERE SearchPhrase='" & EscapeSQuotes(phrase) & "' AND ActiveMonth=" & Format(Date, "m"))
            If rst.EOF Then
                GetKwdStatus = "Inactive (Seasonal)"
            Else
                GetKwdStatus = "Active (Seasonal)"
            End If
            rst.Close
            Set rst = Nothing
        Else
            GetKwdStatus = "Active"
        End If
    ElseIf active And pending And Not rejected And Not CBool(donotuse) Then
        GetKwdStatus = "Active (pending)"
    ElseIf Not active And pending And Not rejected And Not CBool(donotuse) Then
        GetKwdStatus = "Pending"
    ElseIf Not active And Not pending And rejected And Not CBool(donotuse) Then
        GetKwdStatus = "Rejected"
    ElseIf Not active And Not pending And Not rejected And CBool(donotuse) Then
        If donotuse = 1 Then
            GetKwdStatus = "Deleted"
        ElseIf donotuse = 2 Then
            GetKwdStatus = "To Be Deleted"
        Else
            GetKwdStatus = "INVALID STATUS"
        End If
    ElseIf Not active And Not pending And Not rejected And Not CBool(donotuse) Then
        GetKwdStatus = "Inactive"
    Else
        GetKwdStatus = "INVALID STATUS"
    End If
End Function

Public Function KwdCheckGoodToGo(phrase As String, svc As String, Optional noBidOK As Boolean = False) As String
    Dim retval As String
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT LandingPage, TrackLink, NewBid, CurrentBid FROM vKeywords WHERE SearchPhrase='" & EscapeSQuotes(phrase) & "' AND Engine='" & svc & "'")
    If Nz(rst("LandingPage")) = "" Then
        retval = retval & "Missing landing page" & vbCrLf
    End If
    'If Nz(rst("TrackLink")) = "" Then
    '    retval = retval & "Missing " & svc & " tracklink" & vbCrLf
    'End If
    If rst("NewBid") = 0 And rst("CurrentBid") = 0 And Not noBidOK Then
        retval = retval & "Missing bid" & vbCrLf
    End If
    rst.Close
    Set rst = Nothing
    If retval = "" Then
        KwdCheckGoodToGo = "OK"
    Else
        KwdCheckGoodToGo = Left(retval, Len(retval) - 2)
    End If
End Function

Public Sub SetKwdStatusActive(phrase As String, svcid As String, Optional noPending As Boolean = False)
    DB.Execute "UPDATE KeywordsAttributes SET NewlyActive=1, Active=1, Pending=" & IIf(noPending, 0, 1) & ", Rejected=0, DoNotUse=0, ReasonCode=-1 WHERE SearchPhrase='" & EscapeSQuotes(phrase) & "' AND EngineID='" & svcid & "'"
End Sub

Public Sub SetKwdStatusInactive(phrase As String, svcid As String)
    DB.Execute "UPDATE KeywordsAttributes SET NewlyActive=0, DoNotUse=0, Active=0, Rejected=0, Pending=0 WHERE SearchPhrase='" & EscapeSQuotes(phrase) & "' AND EngineID='" & svcid & "'"
End Sub

Public Sub SetKwdStatusPending(phrase As String, svcid As String)
    DB.Execute "UPDATE KeywordsAttributes SET NewlyActive=0, Pending=1, Active=0, Rejected=0, DoNotUse=0, ReasonCode=-1 WHERE SearchPhrase='" & EscapeSQuotes(phrase) & "' AND EngineID='" & svcid & "'"
End Sub

Public Sub SetKwdStatusToBeDeleted(phrase As String, svcid As String)
    DB.Execute "UPDATE KeywordsAttributes SET NewlyActive=0, DoNotUse=2, Active=0, Rejected=0, Pending=0 WHERE SearchPhrase='" & EscapeSQuotes(phrase) & "' AND EngineID='" & svcid & "'"
End Sub

Public Sub SetKwdStatusDeleted(phrase As String, svcid As String, Optional reasonCode As Long = -1)
    'Debug.Print "deleting keyword " & phrase
    DB.Execute "UPDATE KeywordsAttributes SET NewlyActive=0, DoNotUse=1, Active=0, Rejected=0, Pending=0" & IIf(reasonCode = -1, "", ", ReasonCode=" & reasonCode) & " WHERE SearchPhrase='" & EscapeSQuotes(phrase) & "' AND EngineID='" & svcid & "'"
End Sub

Public Sub SetBid(phrase As String, svcid As String, Optional bidVal As Single = -1)
    DB.Execute "UPDATE KeywordsAttributes SET NewBid=" & IIf(bidVal = -1, "SuggestedBid", bidVal) & " WHERE SearchPhrase='" & EscapeSQuotes(phrase) & "' AND EngineID='" & svcid & "'"
End Sub

Public Sub SetKwdWarnStatus(phrase As String, svcid As String, newstatus As String)
    DB.Execute "UPDATE KeywordsAttributes SET WarnStatus=" & newstatus & " WHERE SearchPhrase='" & EscapeSQuotes(phrase) & "' AND EngineID='" & svcid & "'"
End Sub

'Public Sub SetTrackLink(phrase As String, svc As String, tl As String)
'    If tl <> "e - Ya" And tl <> "savell" And tl <> "annot+" Then 'bad page load => usually one of these
'        dbExecute "UPDATE KeywordsAttributes SET TrackLinkChanged=1, TrackLink='" & tl & "' WHERE SearchPhrase='" & EscapeSQuotes(phrase) & "' AND Service='" & svc & "'"
'    End If
'End Sub
'
'Public Sub ClearTrackLink(phrase As String, svc As String)
'    dbExecute "UPDATE KeywordsAttributes SET TrackLink=NULL, TrackLinkChanged=1 WHERE SearchPhrase='" & EscapeSQuotes(phrase) & "' AND Service='" & svc & "'"
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : AddItemKeyword
' DateTime  : 4/21/2006 13:07
' Author    : briandonorfio
' Purpose   : Adds an entry in the KeywordsPhrases table for the item (manuf name + the
'             part number), unless the phrase already exists.
'
'             12/14/06 - now two entries, one is a phrase-based entry (desc2), one is
'             the itemnumber itself. The phrase entry is used by keywords engines, and
'             we can have multiple versions of it, if we change desc2 and re-add the
'             keyword. The item# version is used only by the shopping engines. This
'             should make the cost/rev calcs a bit faster. The item# version is also
'             added as part of the cleanup thing, so we don't have to worry about having
'             missing ones during the Feeds.pm script.
'---------------------------------------------------------------------------------------
'
Public Sub AddItemKeyword(item As String, Optional hideErrors As Boolean = False, Optional onoff As Boolean = True)
    Dim fullString As String
    fullString = DLookup("ManufFullNameWeb + ' ' + Desc2", "PartNumbers INNER JOIN ProductLine ON LEFT(PartNumbers.ItemNumber,3)=ProductLine.ProductLine", "ItemNumber='" & item & "'")
    Dim page As String
    page = LCase(item) & ".html"
    fullString = TrimWhitespace(fullString, True, True, True)
    DB.Execute "INSERT INTO KeywordsPhrases ( SearchPhrase, IsItem, ItemLink, LandingPage, Source ) VALUES ( '" & EscapeSQuotes(fullString) & "', 0, '" & item & "', '" & page & "', 'AddItemKeywords' )", hideErrors
    DB.Execute "INSERT INTO KeywordsPhrases ( SearchPhrase, IsItem, ItemLink, LandingPage, Source ) VALUES ( '" & item & "', 1, '" & item & "', '" & page & "', 'AddItemKeywords' )", hideErrors
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, ProductsOnly FROM KeywordsEngines")
    While Not rst.EOF
        Dim phr As String
        phr = IIf(rst("ProductsOnly"), item, EscapeSQuotes(fullString))
        Dim actualOnOff As String
        actualOnOff = IIf(onoff, "1, 0", "0, 1")
        '2012-02-27 amazon like normal engines
        ''2011-08-22 amazon only on explicitly
        'If rst("ID") = 8 Then
        '    actualOnOff = "0"
        'End If
        '2014-02-05 update with onoff value when already exists
        Dim rst2 As ADODB.Recordset
        Set rst2 = DB.retrieve("SELECT COUNT(*) FROM KeywordsAttributes WHERE SearchPhrase='" & phr & "' AND EngineID=" & rst("ID"))
        If rst2(0) = 0 Then
            DB.Execute "INSERT INTO KeywordsAttributes ( SearchPhrase, EngineID, Pending, DoNotUse ) VALUES ( '" & phr & "', '" & rst("ID") & "', " & IIf(onoff, "1, 0", "0, 1") & " )", True  'definitely hide errors here
        Else
            DB.Execute "UPDATE KeywordsAttributes SET Pending=" & Iif(onoff, "1", "0") & ", DoNotUse=" & Iif(onoff, "0", "1") & " WHERE  SearchPhrase='" & phr & "' AND EngineID=" & rst("ID")
        End If
        rst2.Close
        Set rst2 = Nothing
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Public Sub ExportDataToAdwords()
    Dim path As String
    path = "s:\mastest\mas90-signs\keywords_scripts\"
    Dim rst As ADODB.Recordset
    
    Set rst = DB.retrieve("SELECT COUNT(*) AS NumKwds FROM vKeywordsAdwordsToBeDeleted")
    If rst("NumKwds") = 0 Then
        MsgBox "No keywords to delete, skipping this step"
    Else
        ShellWait PERL & " " & path & "adwords_delete_keywords.pl", vbNormalFocus
    End If
    rst.Close
    
    Set rst = DB.retrieve("SELECT COUNT(*) AS NumKwds FROM vKeywordsAdWordsNewNeedUpload")
    If rst("NumKwds") = 0 Then
        MsgBox "No new keywords to upload, skipping this step"
    Else
        ShellWait PERL & " " & path & "adwords_upload_new_keywords.pl"
    End If
    rst.Close
    
    Set rst = DB.retrieve("SELECT COUNT(*) AS NumKwds FROM vKeywordsAdwordsChangedNeedUpload")
    If rst("NumKwds") = 0 Then
        MsgBox "No keywords to update, skipping this step"
    Else
        ShellWait PERL & " " & path & "adwords_update_existing_keywords.pl"
    End If
    rst.Close
    
    Set rst = Nothing
    
    MsgBox "AdWords sync complete!"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : FormatKeyword
' DateTime  : 6/8/2006 13:01
' Author    : BrianDonorfio
' Purpose   : Formats a keyword to what overture will take. Everything should probably
'             be depluralized when going up, since they randomly depluralize some of the
'             keywords.
'---------------------------------------------------------------------------------------
'
'Public Function FormatKeyword(phrase As String) As String
'    Dim retval As String
'    retval = phrase
'    retval = Replace(retval, "-", " ")
'    retval = Replace(retval, "+", " ")
'    retval = Replace(retval, "(", " ")
'    retval = Replace(retval, ")", " ")
'    retval = Replace(retval, "/", " ")
'    retval = Replace(retval, "\", " ")
'    retval = Replace(retval, "&", " ")
'    'retval = DePluralize(retval)
'    retval = TrimWhitespace(retval, True, True, True)
'    FormatKeyword = retval
'End Function
