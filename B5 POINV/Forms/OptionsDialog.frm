VERSION 5.00
Begin VB.Form OptionsDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Miscellaneous Options"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Caption         =   "Meta No-Index"
      Height          =   765
      Left            =   3600
      TabIndex        =   19
      Top             =   1410
      Width           =   2265
      Begin VB.CommandButton btnDisableMetaNoIndexGuard 
         Caption         =   "Disable Meta-NoIndex Confirmation"
         Height          =   435
         Left            =   90
         TabIndex        =   20
         Top             =   240
         Width           =   2025
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Default Website"
      Height          =   1425
      Left            =   3600
      TabIndex        =   17
      Top             =   2250
      Width           =   2295
      Begin VB.ListBox WebsiteList 
         Height          =   1035
         ItemData        =   "OptionsDialog.frx":0000
         Left            =   120
         List            =   "OptionsDialog.frx":0002
         TabIndex        =   18
         Top             =   270
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Follies"
      Height          =   765
      Left            =   3600
      TabIndex        =   14
      Top             =   570
      Width           =   2265
      Begin VB.CommandButton btnFolliesEditable 
         Caption         =   "Follies Editable"
         Height          =   345
         Left            =   90
         TabIndex        =   15
         Top             =   240
         Width           =   2025
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Flush Caches"
      Height          =   2115
      Left            =   150
      TabIndex        =   7
      Top             =   2310
      Width           =   3345
      Begin VB.CommandButton btnFlushCache 
         Caption         =   "Yahoo IDs"
         Height          =   375
         Index           =   6
         Left            =   1890
         TabIndex        =   16
         Top             =   1140
         Width           =   1185
      End
      Begin VB.CommandButton btnFlushCache 
         Caption         =   "Items Status"
         Height          =   375
         Index           =   5
         Left            =   1890
         TabIndex        =   13
         Top             =   720
         Width           =   1185
      End
      Begin VB.CommandButton btnFlushCache 
         Caption         =   "Signs"
         Height          =   375
         Index           =   4
         Left            =   1890
         TabIndex        =   12
         Top             =   300
         Width           =   1185
      End
      Begin VB.CommandButton btnFlushCache 
         Caption         =   "Vendors"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   1185
      End
      Begin VB.CommandButton btnFlushCache 
         Caption         =   "Availability"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1140
         Width           =   1185
      End
      Begin VB.CommandButton btnFlushCache 
         Caption         =   "Manufacturers"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1185
      End
      Begin VB.CommandButton btnFlushCache 
         Caption         =   "Web Paths"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Set Cursor Position"
      Height          =   1635
      Left            =   150
      TabIndex        =   2
      Top             =   570
      Width           =   3345
      Begin VB.OptionButton opStart 
         Caption         =   "go to the BEGINNING of the field"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   600
         Width           =   2895
      End
      Begin VB.OptionButton opEnd 
         Caption         =   "go to the END of the field"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   900
         Width           =   2865
      End
      Begin VB.OptionButton opSelect 
         Caption         =   "SELECT the entire field"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   1200
         Width           =   2865
      End
      Begin VB.Label Label2 
         Caption         =   "Moving between items, the cursor should:"
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   270
         Width           =   3165
      End
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   2430
      TabIndex        =   1
      Top             =   4530
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Miscellaneous Options"
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
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   3075
   End
End
Attribute VB_Name = "OptionsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnDisableMetaNoIndexGuard_Click()
    If META_NOINDEX_MOLLY_ON = True Then
        META_NOINDEX_MOLLY_ON = False
        MsgBox "No more confirmation box when editing noindex state!"
    Else
        META_NOINDEX_MOLLY_ON = True
        MsgBox "Confirmation box added back in!"
    End If
End Sub

Private Sub btnFlushCache_Click(index As Integer)
    Select Case index
        Case Is = 0
            WebPathCache_Flush
        Case Is = 1
            HashTables.CreateHashes "ManufHash"
        Case Is = 2
            HashTables.CreateHashes "AvailHash"
        Case Is = 3
            HashTables.CreateHashes "VendorHash"
        Case Is = 4
            HashTables.CreateHashes "SignHash"
        Case Is = 5
            HashTables.CreateHashes "ItemStatusHash"
        Case Is = 6
            YahooIDFunctions.FlushYahooIDCache
        Case Else
            MsgBox "OOPS!"
    End Select
End Sub

Private Sub btnFolliesEditable_Click()
    If INV_MAINT_FOLLIES_EDIT = False Then
        INV_MAINT_FOLLIES_EDIT = True
        MsgBox "Follies editable!"
    Else
        MsgBox "Follies are already editable, maybe we should be able to turn this off here?"
    End If
End Sub

Private Sub btnOK_Click()
    If Me.opStart = True Then
        SetCursorPosition "start"
    ElseIf Me.opEnd = True Then
        SetCursorPosition "end"
    ElseIf Me.opSelect = True Then
        SetCursorPosition "select"
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Select Case CUR_POS_OPT
        Case Is = "start"
            Me.opStart = True
        Case Is = "end"
            Me.opEnd = True
        Case Is = "select"
            Me.opSelect = True
        Case Is = "default"
            Me.opStart = True
    End Select
    
    Dim tabs(0) As Long
    tabs(0) = 300
    SetListTabStops Me.WebsiteList.hWnd, tabs
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT WebsiteID, URL, Description FROM Websites ORDER BY WebsiteID")
    Me.WebsiteList.Clear
    While Not rst.EOF
        Me.WebsiteList.AddItem rst("Description") & vbTab & rst("WebsiteID")
        If CURRENT_WEBSITE_ID = CLng(rst("WebsiteID")) Then
            Me.WebsiteList.Selected(Me.WebsiteList.ListCount - 1) = True
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub WebsiteList_Click()
    CURRENT_WEBSITE_ID = CLng(ListBoxLineColumn(Me.WebsiteList, Me.WebsiteList.ListIndex, 1))
End Sub
