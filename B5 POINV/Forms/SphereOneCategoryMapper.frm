VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form SphereOneCategoryMapper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sphere One Category Mapper"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Sphere One Categories:"
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   6855
      Begin MSComctlLib.ListView sphereLVI 
         Height          =   1935
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3413
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Master Categories:"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6855
      Begin VB.ListBox masterCatLst 
         Height          =   645
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   6375
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Select a master category to obtain list of sphere one categories that connect to the selected master."
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "SphereOneCategoryMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ItemNumber As String

Public Sub DoIt()
If Me.Visible = True Then
If Len(ItemNumber) > 0 Then
    GetMasterCategoryData
    'Me.Visible = True
Else

End If
End If
End Sub
Public Sub SetItemNumber(itmNumber As String)
ItemNumber = itmNumber
DoIt
End Sub

Private Sub GetMasterCategoryData()
    Me.masterCatLst.Clear
    Dim masterData As ADODB.Recordset
    Set masterData = DB.retrieve("SELECT WebPathID, Name FROM PartNumberWebPaths INNER JOIN MasterCategories ON WebPathID=ID WHERE ItemNumber='" & ItemNumber & "'")
    If masterData.RecordCount > 0 Then
        Do
        Me.masterCatLst.AddItem "{" & masterData("WebPathID") & "} " + masterData("Name")
        masterData.MoveNext
        Loop Until masterData.EOF = True Or masterData.BOF = True
    End If
    
End Sub
Private Sub GetSphereCategoryData()
    Me.sphereLVI.ListItems.Clear
    Me.sphereLVI.ColumnHeaders.Clear
    Me.sphereLVI.View = lvwReport
    Me.sphereLVI.ColumnHeaders.Add , , "Category", sphereLVI.width - (sphereLVI.width / 12)
    Me.sphereLVI.ColumnHeaders.Add , , "ID", sphereLVI.width / 12
    Me.sphereLVI.ColumnHeaders.Add , , "Connected", 1
    
    
    Dim sphereResults As ADODB.Recordset
    Set sphereResults = DB.retrieve("SELECT SphereOneCategories.ID,SphereOneCategories.Name FROM SphereOneCategories")
    If (sphereResults.RecordCount > 0) Then
        Do
            Dim tmpItm As ListItem
            Set li = Me.sphereLVI.ListItems.Add(, , sphereResults("Name"))
            li.SubItems(1) = sphereResults("ID")
            sphereResults.MoveNext
        Loop Until sphereResults.EOF = True Or sphereResults.BOF = True
    End If
    
    
    
End Sub

Private Sub GetLinkedMasterToSphere(MasterID As String)
sphereLVI.sorted = False
    Dim sphereResults As ADODB.Recordset
    Set sphereResults = DB.retrieve("SELECT SphereOneCategories.ID AS SphereID,SphereOneCategories.Name AS SphereName,MasterCategories.ID AS MasterID, MasterCategories.Name AS MasterName FROM SphereOneCategories INNER JOIN MasterCategoryConnectors ON SphereOneCategories.ID = MasterCategoryConnectors.ConnectorID AND MasterCategoryConnectors.ConnectorType=4 AND MasterCategoryConnectors.MasterID=" & CStr(MasterID) & " INNER JOIN MasterCategories ON MasterCategories.ID=MasterCategoryConnectors.MasterID")
    If sphereResults.RecordCount > 0 Then
        Do
            For i = 1 To Me.sphereLVI.ListItems.count
                Dim sphereItm As String
                sphereItm = sphereLVI.ListItems(i).SubItems(1)
                Dim sphereTest As String
                sphereTest = sphereResults("SphereID")
                If sphereItm = sphereTest Then
                    sphereLVI.ListItems(i).ForeColor = vbGreen
                    sphereLVI.ListItems(i).SubItems(2) = "0" & sphereLVI.ListItems(i).Text
                Else
                    If (sphereLVI.ListItems(i).ForeColor <> vbGreen) Then
                        sphereLVI.ListItems(i).ForeColor = vbRed
                        sphereLVI.ListItems(i).SubItems(2) = "1" & sphereLVI.ListItems(i).Text
                    End If
                End If
            Next i
            sphereResults.MoveNext
        Loop Until sphereResults.EOF = True Or sphereResults.BOF = True
    Else
    For j = 1 To Me.sphereLVI.ListItems.count
        sphereLVI.ListItems(j).ForeColor = vbRed
    Next j
    End If
    
sphereLVI.SortKey = 2
sphereLVI.SortOrder = lvwAscending
sphereLVI.sorted = True
    
End Sub


Private Sub masterCatLst_Click()
    Dim index As Integer
    index = masterCatLst.ListIndex
    Dim MasterID As String
    MasterID = Split(Split(masterCatLst.list(index), "{")(1), "}")(0)
    GetSphereCategoryData
    GetLinkedMasterToSphere MasterID
End Sub

Private Sub sphereLVI_DblClick()
    Dim index As Integer
    index = sphereLVI.SelectedItem.index
    Dim ToLinkIt As Boolean
    If sphereLVI.ListItems(index).ForeColor = vbGreen Then
        ToLinkIt = False
    Else
        ToLinkIt = True
    End If
    
    
    Dim SphereID As String
    Dim MasterID As String
    SphereID = sphereLVI.ListItems(index).SubItems(1)
    MasterID = Split(Split(CStr(masterCatLst.list(masterCatLst.ListIndex)), "{")(1), "}")(0)
    
    Dim updateSphere As ADODB.Recordset
    
    
    If ToLinkIt = True Then
        Set updateSphere = DB.retrieve("INSERT INTO MasterCategoryConnectors (MasterID,ConnectorType,ConnectorID) VALUES(" & MasterID & ",4," & SphereID & ")")
        sphereLVI.ListItems(index).ForeColor = vbGreen
        sphereLVI.ListItems(index).SubItems(2) = "0"
    Else
        Set updateSphere = DB.retrieve("DELETE FROM MasterCategoryConnectors WHERE MasterID=" & MasterID & " AND ConnectorType=4 AND ConnectorID=" & SphereID)
        sphereLVI.ListItems(index).ForeColor = vbRed
        sphereLVI.ListItems(index).SubItems(2) = "1"
    End If
    
    
End Sub
