Attribute VB_Name = "ChangeLogging"
Option Explicit

Public Sub LogItemStorePriceChanged(item As String, ByVal newprice As String)
    newprice = Replace(Replace(newprice, "$", ""), ",", "")
    Dim oldprice As String
    oldprice = DLookup("DiscountMarkupPriceRate1", "InventoryMaster", "ItemNumber='" & item & "'")
    If Format(oldprice, "Currency") <> Format(newprice, "Currency") Then
        DB.Execute "INSERT " & _
                  "INTO InventoryPriceLog ( ItemNumber, ChangedBy, IsStorePrice, OldPrice, NewPrice ) " & _
                  "SELECT ItemNumber, '" & Environ("UserName") & "', 1, DiscountMarkupPriceRate1, " & newprice & " " & _
                  "FROM InventoryMaster " & _
                  "WHERE ItemNumber='" & item & "'"
    End If
End Sub

Public Sub LogItemWebPriceChanged(item As String, ByVal newprice As String)
    newprice = Replace(Replace(newprice, "$", ""), ",", "")
    Dim oldprice As String
    oldprice = DLookup("IDiscountMarkupPriceRate1", "InventoryMaster", "ItemNumber='" & item & "'")
    If Format(oldprice, "Currency") <> Format(newprice, "Currency") Then
        DB.Execute "INSERT " & _
                  "INTO InventoryPriceLog ( ItemNumber, ChangedBy, IsStorePrice, OldPrice, NewPrice ) " & _
                  "SELECT ItemNumber, '" & Environ("UserName") & "', 0, IDiscountMarkupPriceRate1, " & newprice & " " & _
                  "FROM InventoryMaster " & _
                  "WHERE ItemNumber='" & item & "'"
    End If
End Sub

Public Sub LogItemEBayPriceChanged(item As String, ByVal newprice As String)
    newprice = Replace(Replace(newprice, "$", ""), ",", "")
    Dim oldprice As String
    oldprice = DLookup("EBayPrice", "InventoryMaster", "ItemNumber='" & item & "'")
    If Format(oldprice, "Currency") <> Format(newprice, "Currency") Then
        DB.Execute "INSERT " & _
                  "INTO InventoryPriceLog ( ItemNumber, ChangedBy, IsStorePrice, OldPrice, NewPrice ) " & _
                  "SELECT ItemNumber, '" & Environ("UserName") & "', 2, EBayPrice, " & newprice & " " & _
                  "FROM InventoryMaster " & _
                  "WHERE ItemNumber='" & item & "'"
    End If
End Sub

Public Sub LogItemCostChanged(item As String, ByVal NewCost As String)
    NewCost = Replace(Replace(NewCost, "$", ""), ",", "")
    Dim oldcost As String
    oldcost = DLookup("StdCost", "InventoryMaster", "ItemNumber='" & item & "'")
    If Format(oldcost, "Currency") <> Format(NewCost, "Currency") Then
        DB.Execute "INSERT " & _
                  "INTO InventoryCostLog ( ItemNumber, ChangedBy, IsRegularCost, OldCost, NewCost ) " & _
                  "SELECT ItemNumber, '" & Environ("UserName") & "', 1, StdCost, " & NewCost & " " & _
                  "FROM InventoryMaster " & _
                  "WHERE ItemNumber='" & item & "'"
    End If
End Sub

Public Sub LogItemTNCChanged(item As String, ByVal NewCost As String)
    NewCost = Replace(Replace(NewCost, "$", ""), ",", "")
    Dim oldcost As String
    oldcost = DLookup("TNC", "InventoryMaster", "ItemNumber='" & item & "'")
    If Format(oldcost, "Currency") <> Format(NewCost, "Currency") Then
        DB.Execute "INSERT " & _
                  "INTO InventoryCostLog ( ItemNumber, ChangedBy, IsRegularCost, OldCost, NewCost ) " & _
                  "SELECT ItemNumber, '" & Environ("UserName") & "', 0, TNC, " & NewCost & " " & _
                  "FROM InventoryMaster " & _
                  "WHERE ItemNumber='" & item & "'"
    End If
End Sub

Public Sub LogItemCostSwapped(item As String)
    DB.Execute "INSERT " & _
              "INTO InventoryCostLog ( ItemNumber, ChangedBy, IsSwap, OldCost, NewCost ) " & _
              "SELECT ItemNumber, '" & Environ("UserName") & "', 1, StdCost, TNC " & _
              "FROM InventoryMaster " & _
              "WHERE ItemNumber='" & item & "'"
End Sub

Public Sub LogItemMAPPChanged(item As String, newmapp As String, ignoreFlag As Boolean)
    If ignoreFlag Then
        'if ignoreFlag is set, then the newmapp field is the on/off value
        'otherwise, behave as normal
        DB.Execute "INSERT " & _
                   "INTO InventoryMAPPLog ( ItemNumber, ChangedBy, OldMAPP, NewMAPP, IgnoreFlag ) " & _
                   "VALUES ( '" & item & "', '" & Environ("UserName") & "', 0, " & newmapp & ", 1 )"
    Else
        newmapp = Replace(Replace(newmapp, "$", ""), ",", "")
        Dim oldmapp As String
        oldmapp = DLookup("MAPP", "InventoryMaster", "ItemNumber='" & item & "'")
        If Format(oldmapp, "Currency") <> Format(newmapp, "Currency") Then
            DB.Execute "INSERT " & _
                      "INTO InventoryMAPPLog ( ItemNumber, ChangedBy, OldMAPP, NewMAPP, IgnoreFlag ) " & _
                      "SELECT ItemNumber, '" & Environ("UserName") & "', MAPP, " & newmapp & ", 0 " & _
                      "FROM InventoryMaster " & _
                      "WHERE ItemNumber='" & item & "'"
        End If
    End If
End Sub

Public Sub LogItemBeingDiscontinued(item As String)
    DB.Execute "INSERT INTO InventoryDiscontinuedLog ( ItemNumber, ChangedBy, Status ) VALUES ( '" & item & "', '" & Environ("UserName") & "', 'being' )"
End Sub

Public Sub LogItemDiscontinuedAtWeb(item As String)
    DB.Execute "INSERT INTO InventoryDiscontinuedLog ( ItemNumber, ChangedBy, Status ) VALUES ( '" & item & "', '" & Environ("UserName") & "', 'web' )"
End Sub

Public Sub LogItemDiscontinuedInStoreAndWeb(item As String)
    DB.Execute "INSERT INTO InventoryDiscontinuedLog ( ItemNumber, ChangedBy, Status ) VALUES ( '" & item & "', '" & Environ("UserName") & "', 'store' )"
End Sub

Public Sub LogItemUnDiscontinued(item As String)
    DB.Execute "INSERT INTO InventoryDiscontinuedLog ( ItemNumber, ChangedBy, Status ) VALUES ( '" & item & "', '" & Environ("UserName") & "', 'un' )"
End Sub
