Attribute VB_Name = "RefreshQtysMain"
Option Explicit

Public DB As DBConn.DatabaseSingleton
Public MASDB As DBConn.DatabaseSingleton

Public Sub Main()
    Set Mouse = New MouseHourglass
    If MsgBox("Refresh quantities now?", vbYesNo) = vbYes Then
        Set DB = New DBConn.DatabaseSingleton
        DB.CurrentDB = DBENG_MSSQL
        
        If IsProcessFlagSet("RefreshQtys") Then
            MsgBox "Refresh qtys is already running by another user"
        Else
            Set MASDB = New DBConn.DatabaseSingleton
            MASDB.CurrentDB = DBENG_MAS90
            
            SetProcessFlag "RefreshQtys"
            Load RefreshQtysStandalone
            RefreshQtysStandalone.Show
            RefreshQtysStandalone.doIt
            Unload RefreshQtysStandalone
            ClearProcessFlag "RefreshQtys"
            
            Set MASDB = Nothing
        End If
        Set DB = Nothing
    End If
    Set Mouse = Nothing
End Sub
