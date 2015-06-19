VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#3.5#0"; "TPControls.ocx"
Begin VB.Form AddWebPaths2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Edit Categories"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnWebsiteSelection 
      Caption         =   "Website: Tools Plus"
      Height          =   405
      Left            =   5250
      TabIndex        =   64
      Tag             =   "0"
      Top             =   30
      Width           =   1935
   End
   Begin VB.ComboBox SortStyleItems 
      Height          =   315
      ItemData        =   "AddWebPaths2.frx":0000
      Left            =   11160
      List            =   "AddWebPaths2.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   63
      Top             =   8010
      Width           =   3195
   End
   Begin VB.CommandButton btnClearCache 
      Caption         =   "CLEAR CACHE"
      Height          =   405
      Left            =   9900
      TabIndex        =   61
      Top             =   30
      Width           =   1575
   End
   Begin VB.CheckBox chkKirksExtraField 
      Caption         =   "Kirk's Checkbox!!!!!"
      Height          =   345
      Left            =   11790
      TabIndex        =   60
      Top             =   8520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton btnConditional2Validate 
      Caption         =   "V"
      Height          =   255
      Left            =   10230
      TabIndex        =   59
      Top             =   7590
      Width           =   285
   End
   Begin VB.CommandButton btnConditional1Validate 
      Caption         =   "V"
      Height          =   255
      Left            =   7560
      TabIndex        =   58
      Top             =   7590
      Width           =   285
   End
   Begin VB.TextBox Conditional2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   8340
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   57
      Top             =   7830
      Width           =   2715
   End
   Begin VB.TextBox Conditional1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   5580
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   55
      Top             =   7830
      Width           =   2715
   End
   Begin VB.ComboBox MetaTitleStyle 
      Height          =   315
      ItemData        =   "AddWebPaths2.frx":0004
      Left            =   6690
      List            =   "AddWebPaths2.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   9090
      Width           =   2025
   End
   Begin VB.CheckBox chkPreview 
      Caption         =   "Preview"
      Height          =   255
      Left            =   9990
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   4020
      Width           =   1035
   End
   Begin VB.CommandButton btnCaptionBottomValidate 
      Caption         =   "V"
      Height          =   255
      Left            =   7560
      TabIndex        =   50
      Top             =   5850
      Width           =   285
   End
   Begin VB.CommandButton btnCaptionTopValidate 
      Caption         =   "V"
      Height          =   255
      Left            =   7560
      TabIndex        =   49
      Top             =   4050
      Width           =   285
   End
   Begin VB.TextBox MetaTitle 
      Height          =   285
      Left            =   8760
      TabIndex        =   47
      Top             =   9090
      Width           =   4665
   End
   Begin VB.TextBox ChildSuffixText 
      Height          =   285
      Left            =   1320
      TabIndex        =   44
      Top             =   5520
      Width           =   4215
   End
   Begin VB.TextBox MetaDescription 
      Height          =   975
      Left            =   11130
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   43
      Top             =   5580
      Width           =   3705
   End
   Begin VB.TextBox ChildPrefixText 
      Height          =   285
      Left            =   1320
      TabIndex        =   41
      Top             =   5220
      Width           =   4215
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   495
      Left            =   13080
      TabIndex        =   39
      Top             =   0
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Publish Status:"
      Height          =   555
      Left            =   210
      TabIndex        =   31
      Top             =   3990
      Width           =   4605
      Begin VB.OptionButton opPublishStatus 
         Caption         =   "Deleted"
         Height          =   255
         Index           =   2
         Left            =   3450
         TabIndex        =   38
         Top             =   240
         Width           =   915
      End
      Begin VB.OptionButton opPublishStatus 
         Caption         =   "Published"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   37
         Top             =   240
         Width           =   1035
      End
      Begin VB.OptionButton opPublishStatus 
         Caption         =   "To Be Published"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   36
         Top             =   240
         Width           =   1485
      End
   End
   Begin VB.ComboBox LayoutStyleItems 
      Height          =   315
      Left            =   11160
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   7440
      Width           =   3195
   End
   Begin VB.ComboBox LayoutStyleSections 
      Height          =   315
      Left            =   11160
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   6870
      Width           =   3195
   End
   Begin VB.TextBox MetaKeywords 
      Height          =   975
      Left            =   11130
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   4290
      Width           =   3705
   End
   Begin VB.CommandButton btnDeleteAndPromote 
      Caption         =   "DELETE && PROMOTE"
      Height          =   405
      Left            =   7890
      TabIndex        =   24
      Top             =   30
      Width           =   1815
   End
   Begin VB.TextBox LoadedGUIIndex 
      Height          =   315
      Left            =   4200
      TabIndex        =   23
      Text            =   "GUIIndex"
      Top             =   90
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox LoadedPathID 
      Height          =   315
      Left            =   3270
      TabIndex        =   22
      Text            =   "ID"
      Top             =   90
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox CaptionBottom 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   5580
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   6090
      Width           =   5475
   End
   Begin VB.TextBox CaptionTop 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   5580
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   4290
      Width           =   5475
   End
   Begin VB.TextBox SectionImagePath 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Top             =   6150
      Width           =   3555
   End
   Begin TPControls.Picture SectionImage 
      Height          =   2100
      Left            =   1950
      TabIndex        =   16
      Top             =   6480
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   3704
   End
   Begin VB.ListBox ItemList 
      Height          =   2400
      ItemData        =   "AddWebPaths2.frx":0008
      Left            =   120
      List            =   "AddWebPaths2.frx":000A
      TabIndex        =   14
      Top             =   6150
      Width           =   1755
   End
   Begin VB.TextBox URLName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Top             =   4920
      Width           =   4215
   End
   Begin VB.TextBox PathName 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   4620
      Width           =   4215
   End
   Begin VB.TextBox ItemCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   255
      Left            =   660
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   5880
      Width           =   735
   End
   Begin VB.Frame PathSliderFrame 
      Caption         =   "Select Path:"
      Height          =   3405
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   14790
      Begin VB.Frame Frame2 
         Caption         =   "Path 1 Filter:"
         Height          =   495
         Left            =   2610
         TabIndex        =   32
         Top             =   120
         Width           =   3915
         Begin VB.OptionButton opPath1Filter 
            Caption         =   "Brands"
            Height          =   225
            Index           =   2
            Left            =   2430
            TabIndex        =   35
            Top             =   210
            Width           =   1365
         End
         Begin VB.OptionButton opPath1Filter 
            Caption         =   "Categories"
            Height          =   225
            Index           =   1
            Left            =   1020
            TabIndex        =   34
            Top             =   210
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton opPath1Filter 
            Caption         =   "All"
            Height          =   225
            Index           =   0
            Left            =   180
            TabIndex        =   33
            Top             =   210
            Width           =   585
         End
      End
      Begin VB.CommandButton PathSliderDeleteButton 
         Caption         =   "Delete Path (and kids!)"
         Height          =   315
         Index           =   0
         Left            =   1500
         TabIndex        =   7
         Top             =   2970
         Width           =   1995
      End
      Begin VB.CommandButton PathSliderAddButton 
         Caption         =   "Add New Path"
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   2970
         Width           =   1305
      End
      Begin VB.CommandButton PathSliderDeeperButton 
         Caption         =   ">> See Deeper Paths >>"
         Height          =   345
         Left            =   12540
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton PathSliderPreviousButton 
         Caption         =   "<< See Previous Paths <<"
         Height          =   345
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.ListBox PathSliderList 
         Height          =   2010
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   900
         Width           =   3585
      End
      Begin VB.Label PathSliderLevelLabel 
         Caption         =   "PATH LEVEL #"
         Height          =   225
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   660
         Width           =   1845
      End
   End
   Begin SHDocVwCtl.WebBrowser Preview 
      Height          =   345
      Left            =   5640
      TabIndex        =   52
      Top             =   4320
      Visible         =   0   'False
      Width           =   9195
      ExtentX         =   16219
      ExtentY         =   609
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label18 
      Caption         =   "Sort Style for Items:"
      Height          =   225
      Left            =   11310
      TabIndex        =   62
      Top             =   7800
      Width           =   1845
   End
   Begin VB.Label Label17 
      Caption         =   "Conditional 2: (HTML)"
      Height          =   225
      Left            =   8340
      TabIndex        =   56
      Top             =   7620
      Width           =   1845
   End
   Begin VB.Label Label16 
      Caption         =   "Conditional 1: (HTML)"
      Height          =   225
      Left            =   5670
      TabIndex        =   54
      Top             =   7620
      Width           =   1845
   End
   Begin VB.Label Label7 
      Caption         =   "Bottom Caption (HTML):"
      Height          =   225
      Left            =   5670
      TabIndex        =   19
      Top             =   5880
      Width           =   1845
   End
   Begin VB.Label Label15 
      Caption         =   "at Tools Plus"
      Height          =   225
      Left            =   13470
      TabIndex        =   48
      Top             =   9150
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "Meta Title Style:"
      Height          =   225
      Left            =   5460
      TabIndex        =   46
      Top             =   9150
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Suffix Kids w/:"
      Height          =   225
      Left            =   150
      TabIndex        =   45
      Top             =   5550
      Width           =   1125
   End
   Begin VB.Label Label12 
      Caption         =   "Meta Description:"
      Height          =   225
      Left            =   11250
      TabIndex        =   42
      Top             =   5370
      Width           =   1845
   End
   Begin VB.Label Label11 
      Caption         =   "Prefix Kids w/:"
      Height          =   225
      Left            =   150
      TabIndex        =   40
      Top             =   5250
      Width           =   1125
   End
   Begin VB.Label Label10 
      Caption         =   "Layout Style for Items:"
      Height          =   225
      Left            =   11310
      TabIndex        =   29
      Top             =   7230
      Width           =   1845
   End
   Begin VB.Label Label9 
      Caption         =   "Layout Style for Sub-sections:"
      Height          =   225
      Left            =   11280
      TabIndex        =   27
      Top             =   6660
      Width           =   2385
   End
   Begin VB.Label Label8 
      Caption         =   "Meta Keywords (comma-sep):"
      Height          =   225
      Left            =   11250
      TabIndex        =   25
      Top             =   4080
      Width           =   2355
   End
   Begin VB.Label Label6 
      Caption         =   "Top Caption (HTML):"
      Height          =   225
      Left            =   5700
      TabIndex        =   18
      Top             =   4080
      Width           =   1845
   End
   Begin VB.Label Label5 
      Caption         =   "Section Image:"
      Height          =   225
      Left            =   1950
      TabIndex        =   15
      Top             =   5910
      Width           =   1155
   End
   Begin VB.Label Label4 
      Caption         =   "URL Name:"
      Height          =   225
      Left            =   150
      TabIndex        =   12
      Top             =   4950
      Width           =   1125
   End
   Begin VB.Label Label3 
      Caption         =   "Section Name:"
      Height          =   225
      Left            =   150
      TabIndex        =   10
      Top             =   4650
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "Items:"
      Height          =   225
      Left            =   150
      TabIndex        =   8
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Add/Edit Categories"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   3075
   End
End
Attribute VB_Name = "AddWebPaths2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PATHSLIDER_NUM_CONTROLS As Long = 5 'actually 6, including 0
Private Const PATHSLIDER_NUM_VISIBLE As Long = 4
Private Const PATHSLIDER_HORIZ_SPACE As Long = 90

Private Const MSG_NO_PATHS As String = "[no subpaths for selected path]"

'Private pathsById As Dictionary
'Private pathsByParent As Dictionary

Private changed As Boolean
Private whichCtl As String

Private lsSections As Long
Private lsItems As Long
Private ssItems As Long

Private fillingForm As Boolean

Private checkBoxReadOnly As Boolean


Private Sub btnClearCache_Click()
    If vbYes = MsgBox("Flush web path cache?", vbYesNo) Then
        WebPathCache_Flush
        clearPathInfoControls
        loadPaths "-1", 0
    End If
End Sub

Private Sub btnWebsiteSelection_Click()
    Dim newsite As Long
    newsite = CLng(Me.btnWebsiteSelection.tag) + 1
    If Not WebsiteNameHash.exists(newsite) Then
        newsite = 0
    End If
    Me.btnWebsiteSelection.tag = newsite
    Me.btnWebsiteSelection.Caption = "Website: " & WebsiteNameHash.item(newsite)
    'reload form
    clearPathInfoControls
    loadPaths "-1", 0
End Sub

Private Sub chkKirksExtraField_Click()
    If fillingForm Then
        Exit Sub
    End If
    If checkBoxReadOnly Then
        fillingForm = True
        Me.chkKirksExtraField = SQLBool(Not CBool(Me.chkKirksExtraField))
        fillingForm = False
        Exit Sub
    End If
    DB.Execute "UPDATE WebPaths SET KirksExtraField=" & Me.chkKirksExtraField & " WHERE ID=" & Me.LoadedPathID
End Sub


'=================================================================

Private Sub btnDeleteAndPromote_Click()
    If Me.LoadedGUIIndex = 0 Then
        MsgBox "can't delete&promote path1s!"
        Exit Sub
    End If
    
    Dim pathToDel As String
    pathToDel = Me.LoadedPathID
    If vbYes = MsgBox("Are you sure you want to remove " & qq(Me.PathSliderList(Me.LoadedGUIIndex)) & " and move all sub-items and sub-sections up in the hierarchy?", vbYesNo + vbDefaultButton2) Then
        
        Mouse.Hourglass True
        
        Dim newParentID As String
        'newParentID = pathsById.item(pathToDel).item("parent")
        newParentID = WebPathCache_GetParent(pathToDel)
        
        'move deleted path's items to parent
        'first remove any that duplicate, so we don't get a key violation
        DB.Execute "DELETE FROM PartNumberWebPaths WHERE PartNumberWebPaths.WebPathID=" & pathToDel & " AND EXISTS(SELECT * FROM PartNumberWebPaths AS PNWP2 WHERE PartNumberWebPaths.ItemNumber=PNWP2.ItemNumber AND PNWP2.WebPathID=" & newParentID & ")"
        DB.Execute "UPDATE PartNumberWebPaths SET WebPathID=" & newParentID & " WHERE WebPathID=" & pathToDel
        
        'fix path levels (not necessary, but useful maybe)
        fixPathLevels pathToDel
        
        'change parent pointers on children
        DB.Execute "UPDATE WebPaths SET ParentID=" & newParentID & " WHERE ParentID=" & pathToDel
        
        'then delete
        DB.Execute "DELETE FROM PartNumberWebPaths WHERE WebPathID=" & Me.LoadedPathID
        DB.Execute "DELETE FROM WebPaths WHERE ID=" & Me.LoadedPathID
        
        'how do we sync the cache here? flush for now
        WebPathCache_Flush
        PathSliderList_Click Me.LoadedGUIIndex - 1
        
        Mouse.Hourglass False
    End If
End Sub

Private Sub fixPathLevels(parentid As String)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID FROM WebPaths WHERE ParentID=" & parentid)
    If Not rst.EOF Then
        While Not rst.EOF
            fixPathLevels rst("ID")
            rst.MoveNext
        Wend
        DB.Execute "UPDATE WebPaths SET PathLevel=PathLevel-1 WHERE ParentID=" & parentid
    End If
    rst.Close
    Set rst = Nothing
End Sub









'=================================================================
Private Sub Form_Load()
    
    If GetCurrentUserNickname() = "Kirk" Then
        Me.chkKirksExtraField.Visible = True
        'Me.LoadedPathID.Visible = True
    End If
    
    Me.btnWebsiteSelection.tag = CURRENT_WEBSITE_ID
    Me.btnWebsiteSelection.Caption = "Website: " & WebsiteNameHash.item(CURRENT_WEBSITE_ID)
    
    Mouse.Hourglass True
    changed = False
    sliderGUIInitialize
    'loadPathCache
    If ImageCache Is Nothing Then
        Set ImageCache = New Dictionary
    End If
    Me.SectionImage.SetCachePointer ImageCache
    
    Me.Preview.Height = 4725
    
    Me.LayoutStyleSections.AddItem "Bubbles 3-column"
    Me.LayoutStyleSections.AddItem "Router bit table"
    Me.LayoutStyleSections.AddItem "Alphabetical list"
    Me.LayoutStyleSections.AddItem "Multi-section (coastal style)"
    ExpandDropDownToFit Me.LayoutStyleSections
    Me.LayoutStyleItems.AddItem "Standard rectangle"
    Me.LayoutStyleItems.AddItem "Group table"
    Me.LayoutStyleItems.AddItem "Grid"
    Me.LayoutStyleItems.AddItem "Standard rectangle (orderable)"
    ExpandDropDownToFit Me.LayoutStyleItems
    Me.SortStyleItems.AddItem "By Name"
    Me.SortStyleItems.AddItem "By Price"
    ExpandDropDownToFit Me.SortStyleItems
    Me.MetaTitleStyle.AddItem "Section Name Only"
    Me.MetaTitleStyle.AddItem "Section Name w/ 2 Sub-sections"
    Me.MetaTitleStyle.AddItem "Override w/"
    ExpandDropDownToFit Me.MetaTitleStyle
    
    clearPathInfoControls
    loadPaths "-1", 0
    
    Mouse.Hourglass False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If changed Then
        handleChange
    End If
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub loadPaths(parentid As String, guiIndex As Long, Optional keepSelection As Boolean = False)
    Mouse.Hourglass True
    
    Dim origSelection As Long
    origSelection = -1
    If keepSelection Then
        origSelection = Me.PathSliderList(guiIndex).ListIndex
    End If
    
    Dim i As Long
    Me.PathSliderList(guiIndex).Clear
    
    Dim childList As Variant
    If parentid = "-1" Then
        childList = WebPathCache_GetRoots(Me.btnWebsiteSelection.tag, Me.opPath1Filter(0).value Or Me.opPath1Filter(1).value, Me.opPath1Filter(0).value Or Me.opPath1Filter(2).value, True)
    Else
        childList = WebPathCache_GetChildList(parentid, True)
    End If
    
    If UBound(childList) <> -1 Then
        ReDim subpaths(UBound(childList)) As String
        For i = 0 To UBound(subpaths)
            subpaths(i) = WebPathCache_GetName(CStr(childList(i))) & vbTab & CStr(childList(i))
            If WebPathCache_GetPublishStatus(CStr(childList(i))) = 2 Then
                subpaths(i) = "[deleted] " & subpaths(i)
            End If
        Next i
        
        QuickSort subpaths
        
        For i = 0 To UBound(subpaths)
            Me.PathSliderList(guiIndex).AddItem subpaths(i)
        Next i
        
        Me.PathSliderList(guiIndex).Enabled = True
        Me.PathSliderAddButton(guiIndex).Enabled = True
        Me.PathSliderDeleteButton(guiIndex).Enabled = False
        
        If keepSelection And origSelection <> -1 Then
            If Me.PathSliderList(guiIndex).ListCount >= origSelection Then
                Me.PathSliderList(guiIndex).Selected(origSelection) = True
            End If
        End If
    Else
        Me.PathSliderList(guiIndex).AddItem MSG_NO_PATHS
        Me.PathSliderList(guiIndex).Enabled = False
        Me.PathSliderAddButton(guiIndex).Enabled = True
    End If
    
    For i = guiIndex + 1 To PATHSLIDER_NUM_CONTROLS
        Me.PathSliderList(i).Clear
        Me.PathSliderList(i).Enabled = False
        Me.PathSliderAddButton(i).Enabled = False
        Me.PathSliderDeleteButton(i).Enabled = False
    Next i
    
    Mouse.Hourglass False
End Sub

Private Sub loadPathInfo(pathid As String)
    Mouse.Hourglass True
    fillingForm = True
    
    Dim i As Long
    
    Me.PathName = WebPathCache_GetName(pathid, False, False)
    Me.URLName = getFullURLIdentifier(pathid)
    Me.ChildPrefixText = WebPathCache_GetPrefix(pathid)
    Me.ChildSuffixText = WebPathCache_GetSuffix(pathid)
    Me.SectionImagePath = GenerateSectionPicturePath(Me.URLName)
    Me.SectionImage.DisplayImage Me.SectionImagePath
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT PublishStatus, ChildPrefixText, ChildSuffixText, CaptionTop, CaptionBottom, PageLayoutConditionalArea1, PageLayoutConditionalArea2, MetaKeywords, MetaDescription, MetaTitleStyle, MetaTitle, PageLayoutSections, PageLayoutItems, SortStyleItems, KirksExtraField FROM WebPaths WHERE ID=" & pathid)
    If rst.EOF Then
        MsgBox "wtf? can't get info!"
        For i = 0 To Me.opPublishStatus.UBound
            Me.opPublishStatus(i).value = False
        Next i
        Me.CaptionTop = ""
        Me.CaptionBottom = ""
        Me.Conditional1 = ""
        Me.Conditional2 = ""
        Me.MetaKeywords = ""
        Me.MetaDescription = ""
        Me.MetaTitleStyle.ListIndex = -1
        Me.MetaTitle = ""
        Me.LayoutStyleSections.ListIndex = -1
        lsSections = -1
        'Me.btnSectionLayoutExtras.Visible = False
        Me.LayoutStyleItems.ListIndex = -1
        lsItems = -1
        Me.SortStyleItems.ListIndex = -1
        ssItems = -1
        Me.chkKirksExtraField = 0
    Else
        Me.opPublishStatus(rst("PublishStatus")).value = True
        Me.CaptionTop = Nz(rst("CaptionTop"))
        Me.CaptionBottom = Nz(rst("CaptionBottom"))
        Me.Conditional1 = Nz(rst("PageLayoutConditionalArea1"))
        Me.Conditional2 = Nz(rst("PageLayoutConditionalArea2"))
        Me.MetaKeywords = Nz(rst("MetaKeywords"))
        Me.MetaDescription = Nz(rst("MetaDescription"))
        Me.MetaTitleStyle.ListIndex = rst("MetaTitleStyle")
        Me.MetaTitle = Nz(rst("MetaTitle"))
        Me.LayoutStyleSections.ListIndex = rst("PageLayoutSections")
        lsSections = rst("PageLayoutSections")
        Me.LayoutStyleItems.ListIndex = rst("PageLayoutItems")
        lsItems = rst("PageLayoutItems")
        Me.SortStyleItems.ListIndex = rst("SortStyleItems")
        ssItems = rst("SortStyleItems")
        Me.chkKirksExtraField = SQLBool(rst("KirksExtraField"))
    End If
    rst.Close
    
    Set rst = DB.retrieve("SELECT ItemNumber FROM PartNumberWebPaths WHERE WebPathID=" & pathid)
    Me.ItemCount = rst.RecordCount
    Me.ItemList.Clear
    While Not rst.EOF
        Me.ItemList.AddItem rst("ItemNumber")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    For i = 0 To Me.opPublishStatus.UBound
        Me.opPublishStatus(i).Enabled = True
    Next i
    Enable Me.PathName
    Enable Me.ChildPrefixText
    Enable Me.ChildSuffixText
    Enable Me.CaptionTop
    Enable Me.CaptionBottom
    'TODO: conditional1/2
    Enable Me.MetaKeywords
    Enable Me.MetaDescription
    Enable Me.MetaTitleStyle
    If Me.MetaTitleStyle.ListIndex = 2 Then
        Enable Me.MetaTitle
    Else
        Disable Me.MetaTitle
    End If
    Me.LayoutStyleSections.Enabled = True
    Me.LayoutStyleItems.Enabled = True
    Me.SortStyleItems.Enabled = True
    checkBoxReadOnly = False
    Me.ItemList.Enabled = True
    Disable Me.Conditional1
    Disable Me.Conditional2
    Select Case Me.LayoutStyleSections
        Case Is = "Bubbles 3-column"
            'Me.btnSectionLayoutExtras.Visible = False
        Case Is = "Router bit table"
            'Me.btnSectionLayoutExtras.Visible = True
            Enable Me.Conditional1
        Case Is = "Alphabetical list"
            'Me.btnSectionLayoutExtras.Visible = False
        Case Is = "Multi-section (coastal style)"
            'Me.btnSectionLayoutExtras.Visible = False
        Case Else
            MsgBox "Layout style not included in switch!"
            'Me.btnSectionLayoutExtras.Visible = False
    End Select
    Disable Me.Conditional2
    Select Case Me.LayoutStyleItems
        Case Is = "Standard rectangle"
            'nothing
        Case Is = "Group table"
            Enable Me.Conditional2
        Case Is = "Grid"
            'nothing
        Case Is = "Standard rectangle (orderable)"
            'nothing
        Case Else
            MsgBox "Layout style not included in switch!"
    End Select
    
    If Me.chkPreview Then
        loadPreviewWebPage
    End If
    
    fillingForm = False
    Mouse.Hourglass False
End Sub

Private Sub clearPathInfoControls()
    fillingForm = True
    Dim i As Long
    For i = 0 To Me.opPublishStatus.UBound
        Me.opPublishStatus(i).value = False
        Me.opPublishStatus(i).Enabled = False
    Next i
    Me.PathName = ""
    Disable Me.PathName
    Me.ChildPrefixText = ""
    Disable Me.ChildPrefixText
    Me.ChildSuffixText = ""
    Disable Me.ChildSuffixText
    Me.URLName = ""
    Me.ItemCount = ""
    Me.ItemList.Clear
    Me.ItemList.Enabled = False
    Me.SectionImagePath = ""
    Me.SectionImage.DisplayImage ""
    Me.CaptionTop = ""
    Disable Me.CaptionTop
    Me.CaptionBottom = ""
    Disable Me.CaptionBottom
    Me.Conditional1 = ""
    Disable Me.Conditional1
    Me.Conditional2 = ""
    Disable Me.Conditional2
    Me.MetaKeywords = ""
    Disable Me.MetaKeywords
    Me.MetaDescription = ""
    Disable Me.MetaDescription
    Me.MetaTitleStyle.ListIndex = -1
    Me.MetaTitleStyle.Enabled = False
    Me.MetaTitle = ""
    Disable Me.MetaTitle
    Me.LayoutStyleSections.ListIndex = -1
    Me.LayoutStyleSections.Enabled = False
    'Me.btnSectionLayoutExtras.Visible = False
    Me.LayoutStyleItems.ListIndex = -1
    Me.LayoutStyleItems.Enabled = False
    Me.SortStyleItems.ListIndex = -1
    Me.SortStyleItems.Enabled = False
    Me.chkKirksExtraField = 0
    checkBoxReadOnly = True
    fillingForm = False
End Sub

Private Sub sliderGUIInitialize()
    Dim initial_left As Long
    initial_left = Me.PathSliderList(0).Left
    Dim i As Long
    For i = 0 To PATHSLIDER_NUM_CONTROLS
        If i <> 0 Then
            Load Me.PathSliderList(i)
            Load Me.PathSliderAddButton(i)
            Load Me.PathSliderDeleteButton(i)
            Load Me.PathSliderLevelLabel(i)
        End If
        
        Me.PathSliderList(i).Top = Me.PathSliderList(0).Top
        Me.PathSliderList(i).Left = initial_left + (i * (Me.PathSliderList(0).width + PATHSLIDER_HORIZ_SPACE))
        Me.PathSliderList(i).Visible = True
        
        Me.PathSliderAddButton(i).Top = Me.PathSliderList(i).Top + 2070
        Me.PathSliderAddButton(i).Left = Me.PathSliderList(i).Left + 60
        Me.PathSliderAddButton(i).Visible = True
        
        Me.PathSliderDeleteButton(i).Top = Me.PathSliderList(i).Top + 2070
        Me.PathSliderDeleteButton(i).Left = Me.PathSliderList(i).Left + 1410
        Me.PathSliderDeleteButton(i).Visible = True
        
        Me.PathSliderLevelLabel(i).Caption = "Path " & CStr(i + 1)
        Me.PathSliderLevelLabel(i).Top = Me.PathSliderList(i).Top - 240
        Me.PathSliderLevelLabel(i).Left = Me.PathSliderList(i).Left + 150
        Me.PathSliderLevelLabel(i).Visible = True
    Next i
    
    Dim newwidth As Long
    newwidth = (initial_left * 2) + _
               (PATHSLIDER_NUM_VISIBLE * Me.PathSliderList(0).width) + _
               ((PATHSLIDER_NUM_VISIBLE - 1) * PATHSLIDER_HORIZ_SPACE)
    If Me.PathSliderFrame.width <> newwidth Then
        MsgBox "The width for the frame is incorrect. Should be " & newwidth & ", tell Brian to fix this!"
        Me.PathSliderFrame.width = newwidth
        Me.width = newwidth + (Me.PathSliderFrame.Left * 4)
    End If
    
    Me.PathSliderPreviousButton.Enabled = False
    Me.PathSliderDeeperButton.Enabled = True
    
    For i = 0 To PATHSLIDER_NUM_CONTROLS
        SetListTabs Me.PathSliderList(i), Array(300)
    Next i
End Sub

Private Sub sliderGUIMove(direction As Long)
    direction = direction * -1
    Me.PathSliderPreviousButton.Enabled = False
    Me.PathSliderDeeperButton.Enabled = False
    Dim i As Long, j As Long, offset As Long
    offset = 150
    Dim moveAmt As Long, movePer As Long, offby As Long
    moveAmt = PATHSLIDER_HORIZ_SPACE + Me.PathSliderList(0).width
    movePer = CLng(moveAmt / offset)
    offby = moveAmt - (movePer * offset)
    If direction < 0 Then
        offset = offset * -1
        offby = offby * -1
    End If
    For i = 0 To movePer - 1
        For j = 0 To PATHSLIDER_NUM_CONTROLS
            Me.PathSliderList(j).Left = Me.PathSliderList(j).Left + offset
            Me.PathSliderAddButton(j).Left = Me.PathSliderAddButton(j).Left + offset
            Me.PathSliderDeleteButton(j).Left = Me.PathSliderDeleteButton(j).Left + offset
            Me.PathSliderLevelLabel(j).Left = Me.PathSliderLevelLabel(j).Left + offset
        Next j
        Sleep 1
        DoEvents
    Next i
    For j = 0 To PATHSLIDER_NUM_CONTROLS
        Me.PathSliderList(j).Left = Me.PathSliderList(j).Left + offby
        Me.PathSliderAddButton(j).Left = Me.PathSliderAddButton(j).Left + offby
        Me.PathSliderDeleteButton(j).Left = Me.PathSliderDeleteButton(j).Left + offby
        Me.PathSliderLevelLabel(j).Left = Me.PathSliderLevelLabel(j).Left + offby
    Next j
    DoEvents
    
    Me.PathSliderPreviousButton.Enabled = CBool(Me.PathSliderList(0).Left <= 0)
    Me.PathSliderDeeperButton.Enabled = CBool(Me.PathSliderList(Me.PathSliderList.UBound).Left >= Me.PathSliderFrame.width)
End Sub

Private Sub opPath1Filter_Click(index As Integer)
    clearPathInfoControls
    loadPaths "-1", 0
End Sub

Private Sub PathSliderAddButton_Click(index As Integer)
    Dim parentid As String, previousURL As String, previousName As String
    parentid = getSelectedPathID(index - 1)
    previousURL = getFullURLIdentifier(parentid)
    previousName = getFullPathName(parentid)
    
    Dim newpathname As String, newPathURL As String
    newpathname = InputBox(IIf(previousName = "", "Add top-level path", "Add path under " & previousName), "Add Path", "")
    If newpathname = "" Then
        Exit Sub
    End If
    
    newPathURL = newpathname
    newPathURL = FormatForYahooURL(newPathURL)
    
    Dim urlIsOk As Boolean
    urlIsOk = False
    While Not urlIsOk
        Dim maxPartLen As Long
        maxPartLen = 100 - Len(IIf(previousURL = "", "", previousURL & "-"))
        
        If maxPartLen > 32 Then
            '32 is the field size for the urlidentifier
            maxPartLen = 32
        End If
        
        newPathURL = InputBox("Enter URL (max " & maxPartLen & " chars) " & IIf(previousURL = "", "", " (will prepend with " & previousURL & ")") & ":", "Edit URL", newPathURL)
        If newPathURL = "" Then
            Exit Sub
        End If
        
        newPathURL = FormatForYahooURL(newPathURL)
        
        If Len(newPathURL) > maxPartLen Then
            MsgBox "URL is too long! Currently " & Len(newPathURL) & " chars, must be " & maxPartLen & " or less!"
        ElseIf "0" <> DLookup("COUNT(*)", "WebPaths", "WebsiteID=" & Me.btnWebsiteSelection.tag & " AND ParentID" & IIf(parentid = "-1", " IS NULL", "=" & parentid) & " AND URLIdentifier='" & EscapeSQuotes(newPathURL) & "'") Then
            MsgBox "That URL is already in use for this parent!"
        Else
            If vbYes = MsgBox("Use this URL? You can't edit it after this!" & vbCrLf & vbCrLf & IIf(previousURL = "", "", previousURL & "-") & newPathURL, vbYesNo) Then
                urlIsOk = True
            End If
        End If
    Wend
    
    Dim ismanuf As String, linkLC As String
    ismanuf = "0"
    linkLC = ""
    If previousURL = "" Then
        If vbYes = MsgBox("Is this a manufacturer path?", vbYesNo) Then
            ismanuf = "1"
            Dim lcIsOk As Boolean
            lcIsOk = False
            While Not lcIsOk
                linkLC = InputBox("Enter the 3-letter line code to link to:", "Enter Line Code", newPathURL)
                linkLC = DLookup("ID", "ProductLine", "ProductLine='" & EscapeSQuotes(linkLC) & "'")
                If linkLC = "" Then
                    MsgBox "Error: couldn't find that line code!"
                Else
                    lcIsOk = True
                End If
            Wend
        End If
    End If
    
    If vbYes = MsgBox("Are you sure you want to create this new path?" & vbCrLf & vbCrLf & "Name: " & newpathname & vbCrLf & vbCrLf & "Full URL: " & IIf(previousURL = "", "", previousURL & "-") & newPathURL, vbYesNo) Then
        DB.Execute "INSERT INTO WebPaths ( WebsiteID, URLIdentifier, WebPathName, PathLevel, ParentID, IsManufacturerPage) VALUES ( " & Me.btnWebsiteSelection.tag & ", '" & EscapeSQuotes(newPathURL) & "', '" & EscapeSQuotes(newpathname) & "', " & CStr(index + 1) & ", " & IIf(parentid = "-1", "NULL", parentid) & ", " & ismanuf & " )"
        Dim newid As String
        newid = DLookup("@@IDENTITY", "WebPaths")
        If linkLC <> "" Then
            DB.Execute "UPDATE ProductLine SET WebSectionID=" & newid & " WHERE ID=" & linkLC
        End If
        WebPathCache_Add newid
        clearPathInfoControls
        loadPaths parentid, CLng(index)
    End If
End Sub

Private Sub PathSliderDeleteButton_Click(index As Integer)
    If Me.PathSliderList(index).ListIndex = -1 Then
        'how did that happen?
        MsgBox "wtf? you shouldn't have been able to click that!"
    Else
        If vbYes = MsgBox("Are you sure you want to delete " & Me.PathSliderList(index) & ", all subpaths, and all associations with items?", vbYesNo + vbDefaultButton2) Then
            Dim i As Long, j As Long
            Dim initialID As String
            initialID = getSelectedPathID(CLng(index))
            Dim idsToDel As PerlArray
            Set idsToDel = New PerlArray
            idsToDel.Push initialID
            While idsToDel.Scalar > 0
                Dim curr As String
                curr = idsToDel.Shift
                Dim childList As Variant
                childList = WebPathCache_GetChildList(curr, True)
                For i = 0 To UBound(childList)
                    idsToDel.Push CStr(childList(i))
                Next i
                DB.Execute "DELETE FROM PartNumberWebPaths WHERE WebPathID=" & curr
                DB.Execute "DELETE FROM WebPaths WHERE ID=" & curr
                WebPathCache_Remove curr
            Wend
            
            clearPathInfoControls
            loadPaths getSelectedPathID(index - 1), CLng(index)
        End If
    End If
End Sub

Private Sub PathSliderList_Click(index As Integer)
    If changed Then
        handleChange
    End If
    
    If Me.PathSliderList(index) = MSG_NO_PATHS Then
        Me.PathSliderList(index).ListIndex = -1
    Else
        Dim pathid As String
        pathid = getSelectedPathID(CLng(index))
        Me.LoadedPathID = pathid
        Me.LoadedGUIIndex = index
        clearPathInfoControls
        loadPaths pathid, CLng(index) + 1
        loadPathInfo pathid
        Me.PathSliderDeleteButton(index).Enabled = True
    End If
    
    If Me.PathSliderList(index).Left > Me.PathSliderFrame.width - (Me.PathSliderList(index).width + PATHSLIDER_HORIZ_SPACE * 2) Then
        PathSliderDeeperButton_Click
    End If
    
    If Me.PathSliderList(index).Left < 200 And index <> 0 Then
        PathSliderPreviousButton_Click
    End If
End Sub

Private Sub PathSliderPreviousButton_Click()
    sliderGUIMove -1
End Sub

Private Sub PathSliderDeeperButton_Click()
    sliderGUIMove 1
End Sub

Private Function getSelectedPathID(guidx As Long) As String
    If guidx = -1 Then
        'not sure if this shortcut is a good idea...
        getSelectedPathID = "-1"
    Else
        getSelectedPathID = Mid(Me.PathSliderList(guidx), InStr(Me.PathSliderList(guidx), vbTab) + 1)
    End If
End Function

Private Function getFullURLIdentifier(pathid As String) As String
    If pathid = "-1" Then
        getFullURLIdentifier = ""
    Else
        Dim fullURL As String, pid As String
        fullURL = WebPathCache_GetURLPart(pathid)
        pid = WebPathCache_GetParent(pathid)
        While pid <> "-1"
            fullURL = WebPathCache_GetURLPart(pid) & "-" & fullURL
            pid = WebPathCache_GetParent(pid)
        Wend
        getFullURLIdentifier = fullURL
    End If
End Function

Private Function getFullPathName(pathid As String) As String
    If pathid = "-1" Then
        getFullPathName = ""
    Else
        Dim fullName As String, pid As String
        fullName = WebPathCache_GetName(pathid)
        pid = WebPathCache_GetParent(pathid)
        While pid <> "-1"
            fullName = WebPathCache_GetName(pid) & "-" & fullName
            pid = WebPathCache_GetParent(pid)
        Wend
        getFullPathName = fullName
    End If
End Function

Private Sub opPublishStatus_Click(index As Integer)
    If Not fillingForm Then
        If index = -1 Then
            MsgBox "WHOA! NOT GOING TO INSERT A BAD VALUE! TELL BRIAN!"
        Else
            DB.Execute "UPDATE WebPaths SET PublishStatus=" & index & " WHERE ID=" & Me.LoadedPathID
        End If
    End If
End Sub

'Private Sub btnSectionLayoutExtras_Click()
'    If Me.LayoutStyleSections.Enabled = False Then
'        MsgBox "This button should have been invisible. Oops."
'    Else
'        Select Case Me.LayoutStyleSections
'            Case Is = "Bubbles 3-column"
'                MsgBox "This button should have been invisible. Oops."
'            Case Is = "Router bit table"
'                Load AddWebPathsRouterBitTableEx
'                AddWebPathsRouterBitTableEx.Show MODAL
'            Case Is = "Alphabetical list"
'                MsgBox "This button should have been invisible. Oops."
'            Case Is = "Multi-section (coastal style)"
'                MsgBox "This button should have been invisible. Oops."
'            Case Else
'                MsgBox "This style wasn't added to the switch!"
'        End Select
'    End If
'End Sub

Private Sub PathName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.PathName.selStart = 0
                Me.PathName.SelLength = Len(Me.PathName)
                KeyCode = 0
                Shift = 0
            End If
        Case Is = vbKeyReturn
            PathName_LostFocus
        Case Is = vbKeyDelete
            PathName_KeyPress KeyCode
    End Select
End Sub

Private Sub PathName_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "PathName"
End Sub

Private Sub PathName_LostFocus()
    If changed Then
        Me.PathName = TrimWhitespace(Me.PathName, True, True, True)
        If Len(Me.PathName) > 64 Then
            Me.PathName = Left(Me.PathName, 64)
            MsgBox "Path name text is too long! Truncating!"
        End If
        DB.Execute "UPDATE WebPaths SET WebPathName='" & EscapeSQuotes(Me.PathName) & "' WHERE ID=" & Me.LoadedPathID
        changed = False
        'this event can be called after the PathSliderList has changed to a different
        'index, so we need to bsearch through everything to find it, can't just use
        'the .ListIndex
        Dim pos As Long
        pos = InListPosition(WebPathCache_GetName(Me.LoadedPathID) & vbTab & Me.LoadedPathID, Me.PathSliderList(Me.LoadedGUIIndex))
        WebPathCache_Remove Me.LoadedPathID
        WebPathCache_Add Me.LoadedPathID
        If pos = -1 Then
            MsgBox "Couldn't find the section in the listbox! You should probably close and reopen, because things are messed up!"
        Else
            Me.PathSliderList(Me.LoadedGUIIndex).list(pos) = WebPathCache_GetName(Me.LoadedPathID) & vbTab & Me.LoadedPathID
        End If
    End If
End Sub

Private Sub ChildPrefixText_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.ChildPrefixText.selStart = 0
                Me.ChildPrefixText.SelLength = Len(Me.ChildPrefixText)
                KeyCode = 0
                Shift = 0
            End If
        Case Is = vbKeyReturn
            ChildPrefixText_LostFocus
        Case Is = vbKeyDelete
            ChildPrefixText_KeyPress KeyCode
    End Select
End Sub

Private Sub ChildPrefixText_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "ChildPrefixText"
End Sub

Private Sub ChildPrefixText_LostFocus()
    If changed Then
        Me.ChildPrefixText = TrimWhitespace(Me.ChildPrefixText, True, True, True)
        If Len(Me.ChildPrefixText) > 32 Then
            Me.ChildPrefixText = Left(Me.ChildPrefixText, 32)
            MsgBox "Prefix text is too long! Truncating!"
        End If
        DB.Execute "UPDATE WebPaths SET ChildPrefixText='" & EscapeSQuotes(Me.ChildPrefixText) & "' WHERE ID=" & Me.LoadedPathID
        changed = False
        'update the cache, lazy way
        WebPathCache_Remove Me.LoadedPathID
        WebPathCache_Add Me.LoadedPathID
        'now need to update the kids in the PathSlider
        loadPaths Me.LoadedPathID, Me.LoadedGUIIndex + 1, True
    End If
End Sub

Private Sub ChildSuffixText_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.ChildSuffixText.selStart = 0
                Me.ChildSuffixText.SelLength = Len(Me.ChildSuffixText)
                KeyCode = 0
                Shift = 0
            End If
        Case Is = vbKeyReturn
            ChildSuffixText_LostFocus
        Case Is = vbKeyDelete
            ChildSuffixText_KeyPress KeyCode
    End Select
End Sub

Private Sub ChildSuffixText_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "ChildSuffixText"
End Sub

Private Sub ChildSuffixText_LostFocus()
    If changed Then
        Me.ChildSuffixText = TrimWhitespace(Me.ChildSuffixText, True, True, True)
        If Len(Me.ChildSuffixText) > 32 Then
            Me.ChildSuffixText = Left(Me.ChildSuffixText, 32)
            MsgBox "Suffix text is too long! Truncating!"
        End If
        DB.Execute "UPDATE WebPaths SET ChildSuffixText='" & EscapeSQuotes(Me.ChildSuffixText) & "' WHERE ID=" & Me.LoadedPathID
        changed = False
        'update the cache, lazy way
        WebPathCache_Remove Me.LoadedPathID
        WebPathCache_Add Me.LoadedPathID
        'now need to update the kids in the PathSlider
        loadPaths Me.LoadedPathID, Me.LoadedGUIIndex + 1, True
    End If
End Sub

Private Sub CaptionTop_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.CaptionTop.selStart = 0
                Me.CaptionTop.SelLength = Len(Me.CaptionTop)
                KeyCode = 0
                Shift = 0
            End If
        Case Is = vbKeyReturn
            CaptionTop_LostFocus
        Case Is = vbKeyDelete
            CaptionTop_KeyPress KeyCode
    End Select
End Sub

Private Sub CaptionTop_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "CaptionTop"
End Sub

Private Sub CaptionTop_LostFocus()
    If changed Then
        If Me.CaptionTop <> "" And Not CBool(InStr(Me.CaptionTop, "<")) Then
            'MsgBox "This doesn't look like HTML!"
            ''but save anyway?
            Me.CaptionTop = "<p>" & Me.CaptionTop & "</p>"
        End If
        Me.CaptionTop = UnFrontpage(Me.CaptionTop)
        DB.Execute "UPDATE WebPaths SET CaptionTop='" & EscapeSQuotes(Me.CaptionTop) & "' WHERE ID=" & Me.LoadedPathID
        changed = False
    End If
End Sub

Private Sub CaptionBottom_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.CaptionBottom.selStart = 0
                Me.CaptionBottom.SelLength = Len(Me.CaptionBottom)
                KeyCode = 0
                Shift = 0
            End If
        Case Is = vbKeyReturn
            CaptionBottom_LostFocus
        Case Is = vbKeyDelete
            CaptionBottom_KeyPress KeyCode
    End Select
End Sub

Private Sub CaptionBottom_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "CaptionBottom"
End Sub

Private Sub CaptionBottom_LostFocus()
    If changed Then
        If Me.CaptionBottom <> "" And Not CBool(InStr(Me.CaptionBottom, "<")) Then
            'MsgBox "This doesn't look like HTML!"
            ''but save anyway?
            Me.CaptionBottom = "<p>" & Me.CaptionBottom & "</p>"
        End If
        Me.CaptionBottom = UnFrontpage(Me.CaptionBottom)
        DB.Execute "UPDATE WebPaths SET CaptionBottom='" & EscapeSQuotes(Me.CaptionBottom) & "' WHERE ID=" & Me.LoadedPathID
        changed = False
    End If
End Sub

Private Sub Conditional1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.Conditional1.selStart = 0
                Me.Conditional1.SelLength = Len(Me.Conditional1)
                KeyCode = 0
                Shift = 0
            End If
        Case Is = vbKeyReturn
            Conditional1_LostFocus
        Case Is = vbKeyDelete
            Conditional1_KeyPress KeyCode
    End Select
End Sub

Private Sub Conditional1_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "Conditional1"
End Sub

Private Sub Conditional1_LostFocus()
    If changed Then
        If Me.Conditional1 <> "" And Not CBool(InStr(Me.Conditional1, "<")) Then
            MsgBox "This doesn't look like HTML!"
            'but save anyway?
        End If
        Me.Conditional1 = UnFrontpage(Me.Conditional1)
        DB.Execute "UPDATE WebPaths SET PageLayoutConditionalArea1='" & EscapeSQuotes(Me.Conditional1) & "' WHERE ID=" & Me.LoadedPathID
        changed = False
    End If
End Sub

Private Sub Conditional2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.Conditional2.selStart = 0
                Me.Conditional2.SelLength = Len(Me.Conditional2)
                KeyCode = 0
                Shift = 0
            End If
        Case Is = vbKeyReturn
            Conditional2_LostFocus
        Case Is = vbKeyDelete
            Conditional2_KeyPress KeyCode
    End Select
End Sub

Private Sub Conditional2_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "Conditional2"
End Sub

Private Sub Conditional2_LostFocus()
    If changed Then
        If Me.Conditional2 <> "" And Not CBool(InStr(Me.Conditional2, "<")) Then
            MsgBox "This doesn't look like HTML!"
            'but save anyway?
        End If
        Me.Conditional2 = UnFrontpage(Me.Conditional2)
        DB.Execute "UPDATE WebPaths SET PageLayoutConditionalArea2='" & EscapeSQuotes(Me.Conditional2) & "' WHERE ID=" & Me.LoadedPathID
        changed = False
    End If
End Sub

Private Sub MetaKeywords_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.MetaKeywords.selStart = 0
                Me.MetaKeywords.SelLength = Len(Me.MetaKeywords)
                KeyCode = 0
                Shift = 0
            End If
        Case Is = vbKeyReturn
            MetaKeywords_LostFocus
        Case Is = vbKeyDelete
            MetaKeywords_KeyPress KeyCode
    End Select
End Sub

Private Sub MetaKeywords_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "MetaKeywords"
End Sub

Private Sub MetaKeywords_LostFocus()
    If changed Then
        DB.Execute "UPDATE WebPaths SET MetaKeywords='" & EscapeSQuotes(Me.MetaKeywords) & "' WHERE ID=" & Me.LoadedPathID
        changed = False
    End If
End Sub

Private Sub MetaDescription_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.MetaDescription.selStart = 0
                Me.MetaDescription.SelLength = Len(Me.MetaDescription)
                KeyCode = 0
                Shift = 0
            End If
        Case Is = vbKeyReturn
            MetaDescription_LostFocus
        Case Is = vbKeyDelete
            MetaDescription_KeyPress KeyCode
    End Select
End Sub

Private Sub MetaDescription_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "MetaDescription"
End Sub

Private Sub MetaDescription_LostFocus()
    If changed Then
        DB.Execute "UPDATE WebPaths SET MetaDescription='" & EscapeSQuotes(Me.MetaDescription) & "' WHERE ID=" & Me.LoadedPathID
        changed = False
    End If
End Sub

Private Sub MetaTitleStyle_Click()
    If Not fillingForm Then
        If Me.MetaTitleStyle.ListIndex = -1 Then
            MsgBox "WHOA! NOT GOING TO INSERT A BAD VALUE! TELL BRIAN!"
        Else
            DB.Execute "UPDATE WebPaths SET MetaTitleStyle=" & Me.MetaTitleStyle.ListIndex & " WHERE ID=" & Me.LoadedPathID
        End If
        Select Case Me.MetaTitleStyle
            Case Is = "Section Name Only"
                If Me.MetaTitle <> "" Then
                    Me.MetaTitle = ""
                    DB.Execute "UPDATE WebPaths SET MetaTitle='" & EscapeSQuotes(Me.MetaTitle) & "' WHERE ID=" & Me.LoadedPathID
                    Disable Me.MetaTitle
                End If
            Case Is = "Section Name w/ 2 Sub-sections"
                If Me.MetaTitle <> "" Then
                    Me.MetaTitle = ""
                    DB.Execute "UPDATE WebPaths SET MetaTitle='" & EscapeSQuotes(Me.MetaTitle) & "' WHERE ID=" & Me.LoadedPathID
                    Disable Me.MetaTitle
                End If
            Case Is = "Override w/"
                Enable Me.MetaTitle
                Me.MetaTitle.SetFocus
            Case Else
                MsgBox "Meta Title style not included in switch!"
        End Select
    End If
End Sub

Private Sub MetaTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyA
            If Shift = vbCtrlMask Then
                Me.CaptionTop.selStart = 0
                Me.CaptionTop.SelLength = Len(Me.MetaTitle)
                KeyCode = 0
                Shift = 0
            End If
        Case Is = vbKeyReturn
            MetaTitle_LostFocus
        Case Is = vbKeyDelete
            MetaTitle_KeyPress KeyCode
    End Select
End Sub

Private Sub MetaTitle_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "MetaTitle"
End Sub

Private Sub MetaTitle_LostFocus()
    If changed Then
        Me.MetaTitle = UnFrontpage(Me.MetaTitle)
        Me.MetaTitle = TrimWhitespace(Me.MetaTitle)
        If Len(Me.MetaTitle) > 128 Then
            MsgBox "Over 128 characters! Truncating!"
            Me.MetaTitle = Left(Me.MetaTitle, 128)
        End If
        If Len(Me.MetaTitle) > 80 Then
            MsgBox "Title is over 80 characters! This is ok, but might be bad for SEO."
        End If
        DB.Execute "UPDATE WebPaths SET MetaTitle='" & EscapeSQuotes(Me.MetaTitle) & "' WHERE ID=" & Me.LoadedPathID
        changed = False
    End If
End Sub

Private Sub handleChange()
    If changed = True Then
        Select Case whichCtl
            Case Is = "PathName"
                PathName_LostFocus
            Case Is = "ChildPrefixText"
                ChildPrefixText_LostFocus
            Case Is = "ChildSuffixText"
                ChildSuffixText_LostFocus
            Case Is = "CaptionTop"
                CaptionTop_LostFocus
            Case Is = "CaptionBottom"
                CaptionBottom_LostFocus
            Case Is = "Conditional1"
                Conditional1_LostFocus
            Case Is = "Conditional2"
                Conditional2_LostFocus
            Case Is = "MetaKeywords"
                MetaKeywords_LostFocus
            Case Is = "MetaDescription"
                MetaDescription_LostFocus
            Case Is = "MetaTitle"
                MetaTitle_LostFocus
            Case Else
                MsgBox "Control for change event " & qq(whichCtl) & " not in switch! Losing changes!"
        End Select
    End If
End Sub

Private Sub LayoutStyleSections_Click()
    If Not fillingForm Then
        If Me.LayoutStyleSections.ListIndex = -1 Then
            MsgBox "WHOA! NOT GOING TO INSERT A BAD VALUE! TELL BRIAN!"
        ElseIf Me.LayoutStyleSections = "Router bit table" Then
            'router bits are deprecated, group items preferred
            MsgBox "Fuck you, I'm not doing router bit specific pages anymore. Use a group table....Kirk."
            Me.LayoutStyleSections.ListIndex = lsSections
        ElseIf Me.LayoutStyleSections.ListIndex = lsSections Then
            'no change
        Else
            DB.Execute "UPDATE WebPaths SET PageLayoutSections=" & Me.LayoutStyleSections.ListIndex & " WHERE ID=" & Me.LoadedPathID
            Me.Conditional1 = ""
            Conditional1_LostFocus
            Disable Me.Conditional1
            Select Case Me.LayoutStyleSections
                Case Is = "Bubbles 3-column"
                    'Me.btnSectionLayoutExtras.Visible = False
                Case Is = "Router bit table"
                    'Me.btnSectionLayoutExtras.Visible = True
                    Enable Me.Conditional1
                Case Is = "Alphabetical list"
                    'Me.btnSectionLayoutExtras.Visible = False
                Case Is = "Multi-section (coastal style)"
                    'Me.btnSectionLayoutExtras.Visible = False
                Case Else
                    MsgBox "Layout style not included in switch!"
                    'Me.btnSectionLayoutExtras.Visible = False
            End Select
        End If
    End If
End Sub

Private Sub LayoutStyleItems_Click()
    If Not fillingForm Then
        If Me.LayoutStyleItems.ListIndex = -1 Then
            MsgBox "WHOA! NOT GOING TO INSERT A BAD VALUE! TELL BRIAN!"
        ElseIf Me.LayoutStyleItems = "Group table" Then
            MsgBox "Don't use this anymore! Set it up as an item group instead."
            Me.LayoutStyleItems.ListIndex = lsItems
        ElseIf Me.LayoutStyleItems.ListIndex = lsItems Then
            'no change
        Else
            DB.Execute "UPDATE WebPaths SET PageLayoutItems=" & Me.LayoutStyleItems.ListIndex & " WHERE ID=" & Me.LoadedPathID
            Me.Conditional2 = ""
            Conditional2_LostFocus
            Disable Me.Conditional2
            Select Case Me.LayoutStyleItems
                Case Is = "Standard rectangle"
                    'nothing
                Case Is = "Group table"
                    Enable Me.Conditional2
                Case Is = "Grid"
                    'nothing
                Case Is = "Standard rectangle (orderable)"
                    'nothing
                Case Else
                    MsgBox "Layout style not included in switch!"
            End Select
        End If
    End If
End Sub

Private Sub SortStyleItems_Click()
    If Not fillingForm Then
        If Me.SortStyleItems.ListIndex = -1 Then
            MsgBox "WHOA! NOT GOING TO INSERT A BAD VALUE! TELL BRIAN!"
        ElseIf Me.SortStyleItems.ListIndex = ssItems Then
            'no change
        Else
            DB.Execute "UPDATE WebPaths SET SortStyleItems=" & Me.SortStyleItems.ListIndex & " WHERE ID=" & Me.LoadedPathID
        End If
    End If
End Sub

Private Sub btnCaptionTopValidate_Click()
    If Len(Me.CaptionTop) = 0 Then
        MsgBox "This field is empty!"
    Else
        Dim errors As Variant
        Mouse.Hourglass True
        errors = ValidateHTMLFragment(Me.CaptionTop)
        Mouse.Hourglass False
        If UBound(errors) = -1 Then
            MsgBox "HTML is ok!"
        Else
            MsgBox "Invalid HTML!" & vbCrLf & vbCrLf & Join(errors, vbCrLf)
        End If
    End If
End Sub

Private Sub btnCaptionBottomValidate_Click()
    If Len(Me.CaptionBottom) = 0 Then
        MsgBox "This field is empty!"
    Else
        Dim errors As Variant
        Mouse.Hourglass True
        errors = ValidateHTMLFragment(Me.CaptionBottom)
        Mouse.Hourglass False
        If UBound(errors) = -1 Then
            MsgBox "HTML is ok!"
        Else
            MsgBox "Invalid HTML!" & vbCrLf & vbCrLf & Join(errors, vbCrLf)
        End If
    End If
End Sub

Private Sub btnConditional1Validate_Click()
    If Len(Me.Conditional1) = 0 Then
        MsgBox "This field is empty!"
    Else
        Dim errors As Variant
        Mouse.Hourglass True
        errors = ValidateHTMLFragment(Me.Conditional1)
        Mouse.Hourglass False
        If UBound(errors) = -1 Then
            MsgBox "HTML is ok!"
        Else
            MsgBox "Invalid HTML!" & vbCrLf & vbCrLf & Join(errors, vbCrLf)
        End If
    End If
End Sub

Private Sub btnConditional2Validate_Click()
    If Len(Me.Conditional2) = 0 Then
        MsgBox "This field is empty!"
    Else
        Dim errors As Variant
        Mouse.Hourglass True
        errors = ValidateHTMLFragment(Me.Conditional2)
        Mouse.Hourglass False
        If UBound(errors) = -1 Then
            MsgBox "HTML is ok!"
        Else
            MsgBox "Invalid HTML!" & vbCrLf & vbCrLf & Join(errors, vbCrLf)
        End If
    End If
End Sub

Private Sub chkPreview_Click()
    If Not fillingForm Then
        If Me.chkPreview Then
            loadPreviewWebPage
            Me.CaptionTop.Visible = False
            Me.CaptionBottom.Visible = False
            Me.Label6.Visible = False
            Me.Label7.Visible = False
            Me.btnCaptionTopValidate.Visible = False
            Me.btnCaptionBottomValidate.Visible = False
            Me.MetaKeywords.Visible = False
            Me.MetaDescription.Visible = False
            Me.LayoutStyleSections.Visible = False
            Me.LayoutStyleItems.Visible = False
            Me.SortStyleItems.Visible = False
            Me.Label8.Visible = False
            Me.Label12.Visible = False
            Me.Label9.Visible = False
            Me.Label10.Visible = False
            Me.Label18.Visible = False
            Me.Conditional1.Visible = False
            Me.Conditional2.Visible = False
            Me.btnConditional1Validate.Visible = False
            Me.btnConditional2Validate.Visible = False
            Me.Label16.Visible = False
            Me.Label17.Visible = False
            Me.chkKirksExtraField.Visible = False
            Me.Preview.Visible = True
        Else
            Me.CaptionTop.Visible = True
            Me.CaptionBottom.Visible = True
            Me.Label6.Visible = True
            Me.Label7.Visible = True
            Me.btnCaptionTopValidate.Visible = True
            Me.btnCaptionBottomValidate.Visible = True
            Me.MetaKeywords.Visible = True
            Me.MetaDescription.Visible = True
            Me.LayoutStyleSections.Visible = True
            Me.LayoutStyleItems.Visible = True
            Me.SortStyleItems.Visible = True
            Me.Label8.Visible = True
            Me.Label12.Visible = True
            Me.Label9.Visible = True
            Me.Label10.Visible = True
            Me.Label18.Visible = True
            Me.Conditional1.Visible = True
            Me.Conditional2.Visible = True
            Me.btnConditional1Validate.Visible = True
            Me.btnConditional2Validate.Visible = True
            Me.Label16.Visible = True
            Me.Label17.Visible = True
            Me.chkKirksExtraField.Visible = CBool(GetCurrentUserNickname = "Kirk")
            Me.Preview.Visible = False
        End If
    End If
End Sub

Private Sub loadPreviewWebPage()
    Open DESTINATION_DIR & "caption.html" For Output As #1
    Print #1, "<html>" & vbCrLf & _
              " <head>" & vbCrLf & _
              "  <base href=""http://" & WebsiteURLHash.item(CLng(Me.btnWebsiteSelection.tag)) & "/"" target=""_blank"">" & vbCrLf & _
              "  <link rel=""stylesheet"" type=""text/css"" href=""http://p1.hostingprod.com/@tools-plus.com/solidcactus/toolsplus-styles.css"">" & vbCrLf & _
              "  <script type=""text/javascript"" src=""http://ajax.googleapis.com/ajax/libs/jquery/1.3.2/jquery.min.js""></script>" & vbCrLf & _
              "  <script type=""text/javascript"" src=""http://p1.hostingprod.com/@tools-plus.com/solidcactus/toolsplus.js""></script>" & vbCrLf & _
              " </head>" & vbCrLf & _
              " <body>"
    Print #1, "  <div style=""text-align: left;"">" 'takes the place of #overall
    Print #1, "  <div style=""width: 737px; background: white;"">" 'takes of place of #container
    Print #1, "  <div class=""content"" style=""overflow: hidden;"" width=""737"">"
    
    Print #1, "   <p><a href=""" & DESTINATION_DIR & "caption.html"">[open preview in new window]</a></p>"
    
    If Len(Me.CaptionTop) <> 0 Then
        Print #1, "   <div id=""FAKE-TOP-ID"" class=""text-caption-area"">" & vbCrLf
        Print #1, Me.CaptionTop & vbCrLf
        Print #1, "  </div>" & vbCrLf
        Print #1, "   <script type=""text/javascript"">TP.DisplayUtils.TruncateSection(""FAKE-TOP-ID"",100,200);</script>"
        Print #1, "   <div class=""spacer-decor""></div>"
    End If
    
    Print #1, "   <p>[sections and/or items go here]</p>" & vbCrLf
    
    If Len(Me.CaptionBottom) <> 0 Then
        Print #1, "   <div class=""spacer-decor""></div>"
        Print #1, "   <div id=""FAKE-BOTTOM-ID"" class=""text-caption-area"">" & vbCrLf
        Print #1, Me.CaptionBottom & vbCrLf
        Print #1, "  </div>" & vbCrLf
    End If
    
    Print #1, "   <p>[end of page content]</p>" & vbCrLf
    
    Print #1, "  </div>"
    Print #1, "  </div>"
    Print #1, "  </div>"
    Print #1, " </body>"
    Print #1, "</html>"
    Close #1
    Me.Preview.Navigate2 DESTINATION_DIR & "caption.html"
End Sub
