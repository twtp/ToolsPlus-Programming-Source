Attribute VB_Name = "ItemStatusBitFlags"
'---------------------------------------------------------------------------------------
' Module    : ItemStatusBitFlags
' DateTime  : 11/15/2005 10:42
' Author    : briandonorfio
' Purpose   : Mid- and low-level handlers for the ItemStatus field. Field is set up as
'             a bitmask, with stock, dropship, d/c by us, d/c by mfr as the bits, left
'             to right. Some bits will affect others, d/c by mfr assumes that it's no
'             longer dropshippable.
'                 stock    dropship     d/c us    d/c mfr
'                   1         1            0        0       basic stock+d/s item
'                   1         0            1        0       stock item, being d/c by us
'                   1         0            0        1       d/c by mfr but we still have
'                                                           stock. might have been
'                                                           d/s-able, but not anymore
'
'             Dependencies:
'               - Microsoft ActiveX Data Objects 2.8 Library
'               - DatabaseFunctions + DBConn
'---------------------------------------------------------------------------------------

Option Explicit

'these are also hard-coded into SQL queries
Private Const STOCKED_MASK As Long = &H8
Private Const DS_MASK As Long = &H4
Private Const DC_US_MASK As Long = &H2
Private Const DC_MFR_MASK As Long = &H1

Public Function IsItemCompletelyDC(itemNumberORstatusCode As String) As Boolean
    If IsItemDC(itemNumberORstatusCode) Then
        If IsItemStocked(itemNumberORstatusCode) Then
            IsItemCompletelyDC = False
        Else
            IsItemCompletelyDC = True
        End If
    Else
        IsItemCompletelyDC = False
    End If
End Function

Public Function IsItemBeingDC(itemNumberORstatusCode As String) As Boolean
    If IsItemDC(itemNumberORstatusCode) Then
        If IsItemStocked(itemNumberORstatusCode) Then
            IsItemBeingDC = True
        Else
            IsItemBeingDC = False
        End If
    Else
        IsItemBeingDC = False
    End If
End Function

Public Function IsItemDC(itemNumberORstatusCode As String) As Boolean
    If IsItemDCbyMfr(itemNumberORstatusCode) Then
        IsItemDC = True
    ElseIf IsItemDCbyUs(itemNumberORstatusCode) Then
        IsItemDC = True
    Else
        IsItemDC = False
    End If
End Function

Public Function IsItemDSonly(itemNumberORstatusCode As String) As Boolean
    IsItemDSonly = IsItemDSable(itemNumberORstatusCode) And Not IsItemStocked(itemNumberORstatusCode)
End Function




'---------------------------------------------------------------------------------------
' Procedure : IsItemStocked
' DateTime  : 11/15/2005 12:22
' Author    : briandonorfio
' Purpose   : Returns a boolean indicating whether the item is currently considered a
'             stock item. Can give just a status code, and will interpret from there, or
'             given an itemnumber, will look up the status code. If item not found,
'             returns false.
'---------------------------------------------------------------------------------------
'
Public Function IsItemStocked(itemNumberORstatusCode As String) As Boolean
    IsItemStocked = queryStatus(itemNumberORstatusCode, STOCKED_MASK)
End Function

Public Function IsItemDSable(itemNumberORstatusCode As String) As Boolean
    IsItemDSable = queryStatus(itemNumberORstatusCode, DS_MASK)
End Function

Public Function IsItemDCbyUs(itemNumberORstatusCode As String) As Boolean
    IsItemDCbyUs = queryStatus(itemNumberORstatusCode, DC_US_MASK)
End Function

Public Function IsItemDCbyMfr(itemNumberORstatusCode As String) As Boolean
    IsItemDCbyMfr = queryStatus(itemNumberORstatusCode, DC_MFR_MASK)
End Function

'---------------------------------------------------------------------------------------
' Procedure : UnSetItemStocked
' DateTime  : 11/15/2005 12:27
' Author    : briandonorfio
' Purpose   : Returns a status code with the stock flag not set. If item is not found,
'             returns a 0 status code (unknown status).
'---------------------------------------------------------------------------------------
'
Public Function UnSetItemStocked(itemNumberORstatusCode As String) As Long
    UnSetItemStocked = unsetStatus(itemNumberORstatusCode, STOCKED_MASK)
End Function

Public Function UnSetItemDSable(itemNumberORstatusCode As String) As Long
    UnSetItemDSable = unsetStatus(itemNumberORstatusCode, DS_MASK)
End Function

Public Function UnSetItemDCbyUs(itemNumberORstatusCode As String) As Long
    UnSetItemDCbyUs = unsetStatus(itemNumberORstatusCode, DC_US_MASK)
End Function

Public Function UnSetItemDCbyMfr(itemNumberORstatusCode As String) As Long
    UnSetItemDCbyMfr = unsetStatus(itemNumberORstatusCode, DC_MFR_MASK)
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetItemStocked
' DateTime  : 11/15/2005 12:24
' Author    : briandonorfio
' Purpose   : Returns a status code with the stock flag for an item or status code. If
'             already set, no change. If item not found, returns a new status code, with
'             just stock set.
'---------------------------------------------------------------------------------------
'
Public Function SetItemStocked(itemNumberORstatusCode As String) As Long
    Dim scode As String
    scode = getStatusCode(itemNumberORstatusCode)
    SetItemStocked = scode Or STOCKED_MASK
End Function

Public Function SetItemDSable(itemNumberORstatusCode As String) As Long
    Dim scode As String
    scode = getStatusCode(itemNumberORstatusCode)
    SetItemDSable = scode Or DS_MASK
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetItemDCbyUs
' DateTime  : 11/15/2005 13:08
' Author    : briandonorfio
' Purpose   : Returns a status code with the discontinued by us flag set. If qoh=0, then
'             it will also unset the stock flag (making the item fully d/c). If item not
'             found, then it will return a status code with just the d/c flag set.
'---------------------------------------------------------------------------------------
'
Public Function SetItemDCbyUs(itemNumberORstatusCode As String, Optional qoh As Long = 9999) As Long
    Dim scode As String
    scode = getStatusCode(itemNumberORstatusCode)
    If qoh = 0 Then
        SetItemDCbyUs = (UnSetItemStocked(scode)) Or DC_US_MASK
    Else
        SetItemDCbyUs = scode Or DC_US_MASK
    End If
End Function

Public Function SetItemDCbyMfr(itemNumberORstatusCode As String, Optional qoh As Long = 9999) As Long
    Dim scode As String
    scode = getStatusCode(itemNumberORstatusCode)
    If qoh = 0 Then
        SetItemDCbyMfr = UnSetItemDSable(UnSetItemStocked(scode)) Or DC_MFR_MASK
    Else
        SetItemDCbyMfr = UnSetItemDSable(scode) Or DC_MFR_MASK
    End If
End Function





Private Function queryStatus(itemNumberORstatusCode As String, mask As Long) As Boolean
    Dim scode As String
    scode = getStatusCode(itemNumberORstatusCode)
    queryStatus = scode And mask
End Function

Private Function unsetStatus(itemNumberORstatusCode As String, mask As Long) As Long
    Dim scode As String
    scode = getStatusCode(itemNumberORstatusCode)
    unsetStatus = IIf(scode And mask, scode Xor mask, scode)
End Function

Private Function getStatusCode(itemNumberORstatusCode As String) As String
    If IsNumeric(itemNumberORstatusCode) Then
        getStatusCode = itemNumberORstatusCode
    Else
        Dim scode As String
        scode = DLookup("ItemStatus", "InventoryMaster", "ItemNumber='" & itemNumberORstatusCode & "'")
        If scode = "" Then  'hopefully this never happens
            scode = "0"
        End If
        getStatusCode = scode
    End If
End Function



'Public Sub setds()
'    Dim rst As ADODB.Recordset
'    Set rst = db.retrieve("SELECT InventoryMaster.ItemNumber, ItemStatus, DiscontinuedFlag, Dropshippable, QtyOnHand FROM InventoryMaster INNER JOIN vActualWhseQty ON InventoryMaster.ItemNumber=vActualWhseQty.ItemNumber WHERE InventoryMaster.ItemNumber<'xx%'")
'    While Not rst.EOF
'        If rst("Dropshippable") Then
'            updateInventoryMaster "ItemStatus", SetItemDSable(rst("ItemStatus")), rst("ItemNumber"), ""
'        End If
'        If rst("DiscontinuedFlag") Then
'            updateInventoryMaster "ItemStatus", SetItemDCbyUs(rst("ItemStatus"), rst("QtyOnHand")), rst("ItemNumber"), ""
'        End If
'        rst.MoveNext
'    Wend
'
'    dbExecute "UPDATE PartNumbers SET AvailabilityOverride=3 WHERE Availability=4"
'    dbExecute "UPDATE PartNumbers SET AvailabilityOverride=4 WHERE Availability=5"
'    dbExecute "UPDATE PartNumbers SET AvailabilityOverride=5 WHERE Availability=6"
'    dbExecute "UPDATE PartNumbers SET AvailabilityOverride=2 WHERE Availability=7"
'    dbExecute "UPDATE PartNumbers SET AvailabilityOverride=6 WHERE Availability=10"
'    dbExecute "UPDATE PartNumbers SET AvailabilityOverride=7 WHERE Availability=13"
'End Sub
