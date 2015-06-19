VERSION 5.00
Begin VB.Form WarehouseLocationMaintenance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Warehouse / Location Maintenance"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   8445
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox MASWhseCode 
      Enabled         =   0   'False
      Height          =   1230
      Left            =   7020
      TabIndex        =   23
      Top             =   1050
      Width           =   1245
   End
   Begin VB.ListBox FillingWarehouse 
      Enabled         =   0   'False
      Height          =   1230
      Left            =   4740
      TabIndex        =   21
      Top             =   1050
      Width           =   1995
   End
   Begin VB.ListBox NonSupervisors 
      Height          =   2205
      Left            =   6750
      Sorted          =   -1  'True
      TabIndex        =   18
      Top             =   2700
      Width           =   1545
   End
   Begin VB.CommandButton SupervisorMoveRight 
      Caption         =   "=>"
      Height          =   405
      Left            =   6330
      TabIndex        =   17
      Top             =   3810
      Width           =   405
   End
   Begin VB.CommandButton SupervisorMoveLeft 
      Caption         =   "<="
      Height          =   405
      Left            =   6330
      TabIndex        =   16
      Top             =   3360
      Width           =   405
   End
   Begin VB.ListBox Supervisors 
      Height          =   2205
      Left            =   4740
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   2700
      Width           =   1545
   End
   Begin VB.CommandButton DeleteVehicle 
      Caption         =   "Remove Vehicle"
      Enabled         =   0   'False
      Height          =   525
      Left            =   2880
      TabIndex        =   14
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton DeleteLocation 
      Caption         =   "Remove Sub-Location"
      Enabled         =   0   'False
      Height          =   525
      Left            =   2790
      TabIndex        =   13
      Top             =   3900
      Width           =   1635
   End
   Begin VB.CommandButton DeleteWarehouse 
      Caption         =   "Remove Warehouse"
      Enabled         =   0   'False
      Height          =   525
      Left            =   390
      TabIndex        =   12
      Top             =   3870
      Width           =   1695
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Close"
      Height          =   465
      Left            =   7260
      TabIndex        =   11
      Top             =   30
      Width           =   1185
   End
   Begin VB.CommandButton AddVehicle 
      Caption         =   "Add Vehicle"
      Height          =   525
      Left            =   2880
      TabIndex        =   10
      Top             =   5820
      Width           =   1215
   End
   Begin VB.CommandButton AddLocation 
      Caption         =   "Add Sub-Location"
      Enabled         =   0   'False
      Height          =   525
      Left            =   2790
      TabIndex        =   9
      Top             =   3330
      Width           =   1635
   End
   Begin VB.CommandButton AddWarehouse 
      Caption         =   "Add Warehouse"
      Height          =   525
      Left            =   390
      TabIndex        =   8
      Top             =   3300
      Width           =   1695
   End
   Begin VB.ListBox VehicleList 
      Height          =   1035
      Left            =   210
      TabIndex        =   6
      Top             =   5850
      Width           =   2565
   End
   Begin VB.ListBox LocationList 
      Height          =   2205
      Left            =   2700
      TabIndex        =   3
      Top             =   1050
      Width           =   1785
   End
   Begin VB.ListBox WarehouseList 
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   1050
      Width           =   2265
   End
   Begin VB.Label generalLabel 
      Caption         =   "MAS Whse Code:"
      Height          =   225
      Index           =   7
      Left            =   7050
      TabIndex        =   22
      Top             =   810
      Width           =   1335
   End
   Begin VB.Label generalLabel 
      Caption         =   "Default Filling Warehouse:"
      Height          =   225
      Index           =   6
      Left            =   4830
      TabIndex        =   20
      Top             =   810
      Width           =   1935
   End
   Begin VB.Label generalLabel 
      Caption         =   "Warehouse Supervisors:"
      Height          =   225
      Index           =   5
      Left            =   4830
      TabIndex        =   19
      Top             =   2490
      Width           =   1845
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8460
      Y1              =   5100
      Y2              =   5100
   End
   Begin VB.Label generalLabel 
      Caption         =   "Vehicles:"
      Height          =   225
      Index           =   4
      Left            =   300
      TabIndex        =   7
      Top             =   5640
      Width           =   1605
   End
   Begin VB.Label generalLabel 
      Caption         =   "Vehicle Maintenance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   90
      TabIndex        =   5
      Top             =   5190
      Width           =   4455
   End
   Begin VB.Label generalLabel 
      Caption         =   "Sub-Locations:"
      Height          =   225
      Index           =   2
      Left            =   2790
      TabIndex        =   4
      Top             =   840
      Width           =   1305
   End
   Begin VB.Label generalLabel 
      Caption         =   "Warehouses:"
      Height          =   225
      Index           =   1
      Left            =   210
      TabIndex        =   2
      Top             =   840
      Width           =   1605
   End
   Begin VB.Label generalLabel 
      Caption         =   "Warehouse / Location Maintenance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   4455
   End
End
Attribute VB_Name = "WarehouseLocationMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private stopevents As Boolean

Private Sub Form_Load()
    stopevents = False
    ReDim tabs(0) As Long
    tabs(0) = 300
    SetListTabStops Me.WarehouseList.hwnd, tabs
    SetListTabStops Me.LocationList.hwnd, tabs
    SetListTabStops Me.VehicleList.hwnd, tabs
    SetListTabStops Me.Supervisors.hwnd, tabs
    SetListTabStops Me.NonSupervisors.hwnd, tabs
    SetListTabStops Me.FillingWarehouse.hwnd, tabs
    requeryWarehouses
    requeryWhseCodes
    requeryVehicles
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub WarehouseList_Click()
    Me.DeleteWarehouse.Enabled = True
    Me.AddLocation.Enabled = True
    Me.FillingWarehouse.Enabled = True
    Me.MASWhseCode.Enabled = True
    requeryLocations getWarehouseID()
    requerySupervisors getWarehouseID()
    setFillingWarehouse getWarehouseID()
    setMASWhseCode getWarehouseID()
End Sub

Private Sub WarehouseList_DblClick()
    Dim newname As String
    newname = InputBox("Rename " & qq(getWarehouseName()) & " to:", "Rename", getWarehouseName())
    If newname <> "" Then
        If newname = getWarehouseName() Then
            'do nothing
        Else
            DB.Execute "UPDATE Warehouses SET Name='" & EscapeSQuotes(newname) & "' WHERE ID=" & getWarehouseID()
            requeryWarehouses
        End If
    End If
End Sub

Private Sub LocationList_Click()
    Me.DeleteLocation.Enabled = True
End Sub

Private Sub LocationList_DblClick()
    Dim lid As String, wid As String, lname As String
    lid = getLocationID()
    wid = getWarehouseID()
    lname = getLocationName()
    Dim newname As String
    newname = InputBox("Rename " & qq(lname) & " to:", "Rename", lname)
    If newname <> "" Then
        DB.Execute "UPDATE WarehouseLocations SET Name='" & EscapeSQuotes(newname) & "' WHERE ID=" & lid
        requeryLocations wid
    End If
End Sub

Private Sub Supervisors_DblClick()
    SupervisorMoveRight_Click
End Sub

Private Sub NonSupervisors_DblClick()
    SupervisorMoveLeft_Click
End Sub

Private Sub VehicleList_Click()
    Me.DeleteVehicle.Enabled = True
End Sub

Private Sub VehicleList_DblClick()
    Dim vid As String, vname As String
    vid = Mid(Me.VehicleList, InStr(Me.VehicleList, vbTab) + 1)
    vname = Left(Me.VehicleList, InStr(Me.LocationList, vbTab) - 1)
    Dim newname As String
    newname = InputBox("Rename " & qq(vname) & " to:", "Rename", vname)
    If newname <> "" Then
        DB.Execute "UPDATE WarehouseTrucks SET Name='" & EscapeSQuotes(newname) & "' WHERE ID=" & vid
        requeryVehicles
    End If
End Sub

Private Sub AddWarehouse_Click()
    Dim newname As String, newcode As String
    newname = InputBox("Enter the name for the new warehouse:")
    If newname <> "" Then
retrycode:
        newcode = InputBox("Enter the associated MAS Warehouse Code:")
        If DLookup("WhseCode", "IMC_WarehouseCode", "WhseCode='" & EscapeSQuotes(newcode) & "'") = "" Then
            MsgBox "MAS Warehouse Code not found!"
            GoTo retrycode
        End If
        
        If newcode <> "" Then
            DB.Execute "INSERT INTO Warehouses ( Name, MasWhseCode ) VALUES ( '" & EscapeSQuotes(newname) & "', '" & EscapeSQuotes(newcode) & "' )"
            requeryWarehouses
        End If
    End If
End Sub

Private Sub DeleteWarehouse_Click()
    If Me.WarehouseList.SelCount = 0 Then
        MsgBox "No warehouse selected. This isn't supposed to happen."
    Else
        Dim wid As String, wname As String
        wid = getWarehouseID()
        wname = getWarehouseName()
        If vbYes = MsgBox("Are you sure you want to delete the warehouse " & qq(wname) & "? This will also delete all sub-locations.", vbYesNo) Then
            DB.Execute "DELETE FROM Warehouses WHERE ID=" & wid
            DB.Execute "DELETE FROM WarehouseLocations WHERE WarehouseID=" & wid
            requeryWarehouses
        End If
    End If
End Sub

Private Sub AddLocation_Click()
    If Me.WarehouseList.SelCount = 0 Then
        MsgBox "You haven't selected a warehouse, this isn't supposed to happen."
    Else
        Dim newname As String
        newname = InputBox("Enter the name for the new sub-location:")
        If newname <> "" Then
            Dim wid As String
            wid = getWarehouseID()
            DB.Execute "INSERT INTO WarehouseLocations ( WarehouseID, Name ) VALUES ( " & wid & ", '" & EscapeSQuotes(newname) & "' )"
            requeryLocations wid
        End If
    End If
End Sub

Private Sub DeleteLocation_Click()
    If Me.LocationList.SelCount = 0 Then
        MsgBox "No location selected. This isn't supposed to happen."
    Else
        Dim lid As String, wid As String, lname As String
        lid = getLocationID()
        wid = getWarehouseID()
        lname = getLocationName()
        If vbYes = MsgBox("Are you sure you want to delete the sub-location " & qq(lname) & "?", vbYesNo) Then
            DB.Execute "DELETE FROM WarehouseLocations WHERE ID=" & lid
            requeryLocations wid
        End If
    End If
End Sub

Private Sub SupervisorMoveLeft_Click()
    If Me.NonSupervisors.ListIndex <> -1 Then
        DB.Execute "INSERT INTO WarehouseSupervisors ( WarehouseID, UserID ) VALUES ( " & getWarehouseID() & ", " & getNonSupervisorID() & " )"
        Me.Supervisors.AddItem Me.NonSupervisors
        Me.NonSupervisors.RemoveItem Me.NonSupervisors.ListIndex
    End If
End Sub

Private Sub SupervisorMoveRight_Click()
    If Me.Supervisors.ListIndex <> -1 Then
        DB.Execute "DELETE FROM WarehouseSupervisors WHERE WarehouseID=" & getWarehouseID() & " AND UserID=" & getSupervisorID()
        Me.NonSupervisors.AddItem Me.Supervisors
        Me.Supervisors.RemoveItem Me.Supervisors.ListIndex
    End If
End Sub

Private Sub FillingWarehouse_Click()
    If Not stopevents Then
        Dim fwid As String
        fwid = Mid(Me.FillingWarehouse, InStr(Me.FillingWarehouse, vbTab) + 1)
        DB.Execute "UPDATE Warehouses SET DefaultFillWhse=" & fwid & " WHERE ID=" & getWarehouseID()
    End If
End Sub

Private Sub MASWhseCode_Click()
    If Not stopevents Then
        If Me.MASWhseCode.ListIndex <> -1 Then
            DB.Execute "UPDATE Warehouses SET MasWhseCode='" & Me.MASWhseCode & "'"
        End If
    End If
End Sub

Private Sub setFillingWarehouse(whseid As String)
    Dim dfltFill As String
    dfltFill = DLookup("DefaultFillWhse", "Warehouses", "ID=" & whseid)
    If dfltFill = "" Then
        dfltFill = "NULL"
    End If
    Dim i As Long
    For i = 0 To Me.FillingWarehouse.ListCount - 1
        If Mid(Me.FillingWarehouse.list(i), InStr(Me.FillingWarehouse.list(i), vbTab) + 1) = dfltFill Then
            stopevents = True
            Me.FillingWarehouse.Selected(i) = True
            stopevents = False
            Exit For
        End If
    Next i
End Sub

Private Sub setMASWhseCode(whseid As String)
    Dim mascode As String
    mascode = DLookup("MasWhseCode", "Warehouses", "ID=" & whseid)
    Dim i As Long
    For i = 0 To Me.MASWhseCode.ListCount - 1
        If Me.MASWhseCode.list(i) = mascode Then
            stopevents = True
            Me.MASWhseCode.Selected(i) = True
            stopevents = False
            Exit For
        End If
    Next i
End Sub

Private Sub AddVehicle_Click()
    Dim newname As String
    newname = InputBox("Enter the name for the new vehicle:")
    If newname <> "" Then
        DB.Execute "INSERT INTO WarehouseTrucks ( Name ) VALUES ( '" & EscapeSQuotes(newname) & "' )"
        requeryVehicles
    End If
End Sub

Private Sub DeleteVehicle_Click()
    If Me.VehicleList.SelCount = 0 Then
        MsgBox "No vehicle selected. This isn't supposed to happen."
    Else
        Dim vid As String, vname As String
        vid = Mid(Me.VehicleList, InStr(Me.VehicleList, vbTab) + 1)
        vname = Left(Me.VehicleList, InStr(Me.VehicleList, vbTab) - 1)
        If vbYes = MsgBox("Are you sure you want to delete the vehicle " & qq(vname) & "?", vbYesNo) Then
            DB.Execute "DELETE FROM WarehouseTrucks WHERE ID=" & vid
            requeryVehicles
        End If
    End If
End Sub

Private Sub requeryWarehouses()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Name, ID FROM Warehouses ORDER BY Name")
    Me.WarehouseList.Clear
    Me.LocationList.Clear
    Me.FillingWarehouse.Clear
    Me.FillingWarehouse.AddItem "<none>" & vbTab & "NULL"
    Me.DeleteWarehouse.Enabled = False
    Me.AddLocation.Enabled = False
    Me.DeleteLocation.Enabled = False
    While Not rst.EOF
        Me.WarehouseList.AddItem rst("Name") & vbTab & rst("ID")
        Me.FillingWarehouse.AddItem rst("Name") & vbTab & rst("ID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryVehicles()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Name, ID FROM WarehouseTrucks ORDER BY Name")
    Me.VehicleList.Clear
    Me.DeleteVehicle.Enabled = False
    While Not rst.EOF
        Me.VehicleList.AddItem rst("Name") & vbTab & rst("ID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryLocations(whseid As String)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Name, ID FROM WarehouseLocations WHERE WarehouseID=" & whseid & " ORDER BY Name")
    Me.LocationList.Clear
    Me.DeleteLocation.Enabled = False
    While Not rst.EOF
        Me.LocationList.AddItem rst("Name") & vbTab & rst("ID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requerySupervisors(whseid As String)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Users.NTUsername, Users.ID, WarehouseSupervisors.WarehouseID FROM Users LEFT OUTER JOIN WarehouseSupervisors ON Users.ID=WarehouseSupervisors.UserID AND (WarehouseSupervisors.WarehouseID=" & whseid & " OR WarehouseSupervisors.WarehouseID IS NULL)")
    Me.Supervisors.Clear
    Me.NonSupervisors.Clear
    While Not rst.EOF
        If Nz(rst("WarehouseID")) = "" Then
            Me.NonSupervisors.AddItem rst("NTUsername") & vbTab & rst("ID")
        Else
            Me.Supervisors.AddItem rst("NTUsername") & vbTab & rst("ID")
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryWhseCodes()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT WhseCode FROM IMC_WarehouseCode")
    Me.MASWhseCode.Clear
    While Not rst.EOF
        Me.MASWhseCode.AddItem rst("WhseCode")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Function getWarehouseID() As String
    getWarehouseID = Mid(Me.WarehouseList, InStr(Me.WarehouseList, vbTab) + 1)
End Function

Private Function getWarehouseName() As String
    getWarehouseName = Left(Me.WarehouseList, InStr(Me.WarehouseList, vbTab) - 1)
End Function

Private Function getLocationID() As String
    getLocationID = Mid(Me.LocationList, InStr(Me.LocationList, vbTab) + 1)
End Function

Private Function getLocationName() As String
    getLocationName = Left(Me.LocationList, InStr(Me.LocationList, vbTab) - 1)
End Function

Private Function getSupervisorID() As String
    getSupervisorID = Mid(Me.Supervisors, InStr(Me.Supervisors, vbTab) + 1)
End Function

Private Function getNonSupervisorID() As String
    getNonSupervisorID = Mid(Me.NonSupervisors, InStr(Me.NonSupervisors, vbTab) + 1)
End Function
