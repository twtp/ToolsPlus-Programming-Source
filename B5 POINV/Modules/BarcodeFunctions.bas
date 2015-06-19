Attribute VB_Name = "BarcodeFunctions"
Option Explicit

Public Function AddNewBarcodeFor(cid As Long, Optional bc As String = "") As String
    If bc = "" Then
        bc = InputBox("Enter new barcode:" & vbCrLf & vbCrLf & "Note: if this new barcode is already associated with an item, you need to remove it from that item first!")
    End If
    bc = Replace(bc, " ", "")
    If bc <> "" Then
        If Left(bc, 2) = "TP" Then
            MsgBox "This is an internal TP barcode, you shouldn't be assigning this!"
        ElseIf Not IsNumeric(bc) Then
            If vbCancel = MsgBox(qq(bc) & " doesn't look like a barcode!" & vbCrLf & vbCrLf & "If this is a part #, then your BC-Wedge is grabbing the barcode, so you'll need to type it in manually.", vbOKCancel) Then
                bc = ""
            End If
        End If
    End If
    If bc = "" Then
        AddNewBarcodeFor = ""
    Else
        Dim oldcid As String
        oldcid = DLookup("ComponentID", "InventoryComponentBarcodes", "Barcode='" & EscapeSQuotes(bc) & "'")
        If oldcid = "" Then
            DB.Execute "INSERT INTO InventoryComponentBarcodes ( Barcode, ComponentID ) VALUES ( '" & EscapeSQuotes(bc) & "', '" & cid & "' )"
            AddNewBarcodeFor = bc
        ElseIf oldcid = cid Then 'same thing, so just silently fail
            AddNewBarcodeFor = ""
        ElseIf vbYes = MsgBox("This barcode " & bc & " is already associated with the component " & DLookup("Component", "InventoryComponents", "ID=" & oldcid) & "!" & vbCrLf & vbCrLf & "Are you sure you want to change it to this component?", vbYesNo + vbDefaultButton2) Then
            DB.Execute "UPDATE InventoryComponentBarcodes SET ComponentID='" & cid & "' WHERE Barcode='" & EscapeSQuotes(bc) & "'"
            AddNewBarcodeFor = bc
        Else
            AddNewBarcodeFor = ""
        End If
    End If
End Function

Public Function AddNewBarcodeForItem(item As String, itemID As String, Optional bc As String = "") As String
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID FROM InventoryComponents WHERE ID IN (SELECT ComponentID FROM InventoryComponentMap WHERE ItemID=" & itemID & ")")
    If rst.RecordCount = 1 Then
        AddNewBarcodeForItem = AddNewBarcodeFor(rst("ID"), bc)
    Else
        MsgBox "Can't add a barcode, since this is a multi-box item! Use the Barcode Form!"
        AddNewBarcodeForItem = ""
    End If
End Function

Public Function EditExistingBarcode(oldbarcode As String, isMasterPack As Boolean) As String
    Dim newbc As String
    Dim mpqty As String
    If isMasterPack = True Then
        mpqty = CStr(Split(Split(oldbarcode, "(")(1), ")")(0))
        oldbarcode = CStr(Split(oldbarcode, ") ")(1))
    End If
    newbc = InputBox("Replace barcode " & oldbarcode & " with:" & vbCrLf & vbCrLf & "Note: if this new barcode is already associated with an item, you need to remove it from that item first!", "Edit barcode", oldbarcode)
    If isMasterPack = True Then

        
reenter:
        Dim tmpqty As String
        tmpqty = InputBox("Enter Master Pack Qty (leave blank to not change):", "Edit Master Pack Qty", "")
        If (Len(tmpqty) > 0) Then mpqty = tmpqty
        
        If IsNumeric(mpqty) = False And mpqty <> "" Then
            MsgBox "Whoa, not expecting that for a master pack value... re-enter it please."
            GoTo reenter
        End If
    End If
    If newbc <> "" And (newbc <> oldbarcode Or isMasterPack = True) Then
        Dim cid As String
        cid = DLookup("ComponentID", "vBarcodesComponents", "Barcode='" & EscapeSQuotes(newbc) & "'")
        If cid = "" Or isMasterPack = True Then
            DB.Execute "UPDATE InventoryComponentBarcodes SET Barcode='" & EscapeSQuotes(newbc) & "' WHERE Barcode='" & EscapeSQuotes(oldbarcode) & "'"
            If isMasterPack = True And mpqty <> "" Then
                DB.Execute "UPDATE InventoryComponentBarcodes SET Quantity=" & mpqty & " WHERE Barcode='" & EscapeSQuotes(oldbarcode) & "'"
            End If
            If isMasterPack = True Then
                EditExistingBarcode = "(" & mpqty & ") " + newbc
            Else
                EditExistingBarcode = newbc
            End If
        Else
            MsgBox "This barcode is already associated with the component " & DLookup("Component", "InventoryComponents", "ID=" & cid) & "!" & vbCrLf & vbCrLf & "Remove it from there first, then you can add it here"
            EditExistingBarcode = ""
        End If
    End If
End Function

Public Function DeleteExistingBarcode(Barcode As String) As Boolean
    If vbYes = MsgBox("Are you sure you want to delete this barcode for this component?", vbYesNo) Then
        DB.Execute "DELETE FROM InventoryComponentBarcodes WHERE Barcode='" & EscapeSQuotes(Barcode) & "'"
        DeleteExistingBarcode = True
    Else
        DeleteExistingBarcode = False
    End If
End Function

Public Function FindBestBarcodeFor(item As String, Optional itemID As String = "") As String
    If itemID = "" Then
        itemID = Utilities.GetItemIDByItemCode(item)
    End If
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Barcode " & _
                          "FROM InventoryComponentBarcodes " & _
                          "WHERE Internal=0 " & _
                          "  AND Quantity=1 " & _
                          "  AND (LEN(Barcode)=12 OR LEN(Barcode)=13)" & _
                          "  AND ISNUMERIC(Barcode)=1 " & _
                          "  AND ComponentID IN (SELECT ComponentID FROM InventoryComponentMap WHERE ItemID=" & itemID & ") " & _
                          "ORDER BY LEN(Barcode) ASC")
    If rst.RecordCount <> 0 Then
        FindBestBarcodeFor = rst("Barcode")
    Else
        FindBestBarcodeFor = ""
    End If
    rst.Close
    Set rst = Nothing
End Function
