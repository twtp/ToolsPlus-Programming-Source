Attribute VB_Name = "SecurityPermissions"
Option Explicit

Private currentPersonID As Long
Private permissionsHash As Dictionary

Public Function GetCurrentUserID() As String
    If permissionsHash Is Nothing Then
        ReloadPermissionsHash
    End If
    GetCurrentUserID = currentPersonID
End Function

Public Function GetCurrentUserNickname() As String
    GetCurrentUserNickname = DLookup("ShortName", "Users", "ID=" & currentPersonID)
End Function

Public Function GetUserNickname(userID As Long) As String
    GetUserNickname = DLookup("ShortName", "Users", "ID=" & userID)
End Function


Public Function HasPermissionsTo(module As String) As Boolean
    If USING_SECURITY Then
        If permissionsHash Is Nothing Then
            ReloadPermissionsHash
        End If
        If permissionsHash.exists(LCase(module)) Then
            HasPermissionsTo = permissionsHash.item(LCase(module)) 'always true, but maybe add an extra allowed/disallowed bit later?
        Else
            HasPermissionsTo = False
        End If
    Else
        HasPermissionsTo = True
    End If
End Function

Public Function IsSupervisorForWhse(whseid As Long) As Boolean
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT COUNT(*) FROM WarehouseSupervisors WHERE WarehouseID=" & whseid & " AND UserID=" & GetCurrentUserID)
    If rst(0) = 0 Then
        IsSupervisorForWhse = False
    Else
        IsSupervisorForWhse = True
    End If
    rst.Close
    Set rst = Nothing
End Function

Public Function ReloadPermissionsHash()
    If Not permissionsHash Is Nothing Then
        Set permissionsHash = Nothing
    End If
    Set permissionsHash = New Dictionary
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM vPermissions WHERE NTUsername='" & EscapeSQuotes(Environ("UserName")) & "'")
    If rst.EOF Then
        rst.Close
        MsgBox "User not registered in po/inv, Brian or Kirk can fix this"
        Set rst = DB.retrieve("SELECT * FROM vPermissions WHERE NTUsername='DEFAULT'")
    End If
    If Not rst.EOF Then
        currentPersonID = rst("PersonID")
    End If
    While Not rst.EOF
        permissionsHash.Add LCase(rst("ModuleName")), True
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Function
