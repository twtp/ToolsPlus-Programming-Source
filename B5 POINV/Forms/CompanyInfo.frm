VERSION 5.00
Begin VB.Form CompanyInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Information"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10965
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnDeleteCompany 
      Caption         =   "Delete Company"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4890
      TabIndex        =   91
      Top             =   90
      Width           =   1485
   End
   Begin VB.Frame Frame7 
      Caption         =   "Vendors:"
      Height          =   1695
      Left            =   60
      TabIndex        =   85
      Top             =   4050
      Width           =   3585
      Begin VB.CommandButton btnEditVendor 
         Caption         =   "Edit"
         Height          =   285
         Left            =   2850
         TabIndex        =   90
         Top             =   1170
         Width           =   645
      End
      Begin VB.CommandButton btnDeleteVendor 
         Caption         =   "Del"
         Height          =   285
         Left            =   2850
         TabIndex        =   89
         Top             =   870
         Width           =   645
      End
      Begin VB.CommandButton btnAddVendor 
         Caption         =   "Add"
         Height          =   285
         Left            =   2850
         TabIndex        =   88
         Top             =   570
         Width           =   645
      End
      Begin VB.ListBox VendorList 
         Height          =   1035
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   87
         Top             =   510
         Width           =   2715
      End
      Begin VB.Label generalLabel 
         Caption         =   "Current List of Vendors:"
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   86
         Top             =   270
         Width           =   1725
      End
   End
   Begin VB.TextBox ReturnArea 
      Height          =   285
      Left            =   9060
      TabIndex        =   84
      Text            =   "RetArea"
      Top             =   90
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CheckBox chkSphereOne 
      Alignment       =   1  'Right Justify
      Caption         =   "Sphere One"
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   3510
      Width           =   1245
   End
   Begin VB.TextBox CompanyID 
      Height          =   285
      Left            =   6630
      TabIndex        =   79
      Text            =   "CID"
      Top             =   90
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox DefaultWarrantyID 
      Height          =   285
      Left            =   7710
      TabIndex        =   78
      Text            =   "DfltW"
      Top             =   90
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox WarrantyID 
      Height          =   285
      Left            =   7170
      TabIndex        =   72
      Text            =   "WID"
      Top             =   90
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame5 
      Caption         =   "Warranty Information:"
      Height          =   5715
      Left            =   7050
      TabIndex        =   52
      Top             =   540
      Width           =   3885
      Begin VB.CommandButton btnDeleteWarranty 
         Caption         =   "Del"
         Height          =   285
         Left            =   3240
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   570
         Width           =   585
      End
      Begin VB.CheckBox chkLifetimeGuarantee 
         Caption         =   "Lifetime"
         Height          =   225
         Left            =   1860
         TabIndex        =   62
         Top             =   3930
         Width           =   855
      End
      Begin VB.CheckBox chkLifetimeWarranty 
         Caption         =   "Lifetime"
         Height          =   225
         Left            =   720
         TabIndex        =   57
         Top             =   1290
         Width           =   855
      End
      Begin VB.CheckBox chkDefaultWarranty 
         Caption         =   "Default Warranty"
         Height          =   495
         Left            =   2520
         TabIndex        =   59
         Top             =   1320
         Width           =   1005
      End
      Begin VB.CommandButton btnOpenURL 
         Caption         =   "WWW"
         Height          =   255
         Index           =   4
         Left            =   30
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   1890
         Width           =   675
      End
      Begin VB.TextBox GuaranteeInfo 
         Height          =   1155
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   64
         Top             =   4500
         Width           =   3705
      End
      Begin VB.TextBox GuaranteeLength 
         Height          =   285
         Left            =   1860
         TabIndex        =   63
         Top             =   4170
         Width           =   765
      End
      Begin VB.TextBox WarrantyURL 
         Height          =   285
         Left            =   690
         TabIndex        =   60
         Top             =   1890
         Width           =   3105
      End
      Begin VB.TextBox WarrantyLength 
         Height          =   285
         Left            =   720
         TabIndex        =   58
         Top             =   1530
         Width           =   765
      End
      Begin VB.TextBox WarrantyInfo 
         Height          =   1335
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   61
         Top             =   2400
         Width           =   3735
      End
      Begin VB.TextBox WarrantyName 
         Height          =   285
         Left            =   690
         TabIndex        =   56
         Top             =   960
         Width           =   3105
      End
      Begin VB.CommandButton btnAddNewWarranty 
         Caption         =   "Add"
         Height          =   285
         Left            =   3240
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   270
         Width           =   585
      End
      Begin VB.ListBox WarrantyList 
         Height          =   645
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   53
         Top             =   240
         Width           =   3105
      End
      Begin VB.Line Line2 
         X1              =   60
         X2              =   3840
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label generalLabel 
         Caption         =   "Guarantee:"
         Height          =   255
         Index           =   28
         Left            =   150
         TabIndex        =   71
         Top             =   4290
         Width           =   825
      End
      Begin VB.Label generalLabel 
         Caption         =   "days"
         Height          =   255
         Index           =   27
         Left            =   2670
         TabIndex        =   70
         Top             =   4200
         Width           =   645
      End
      Begin VB.Label generalLabel 
         Caption         =   "Satisfaction Guarantee:"
         Height          =   255
         Index           =   26
         Left            =   120
         TabIndex        =   69
         Top             =   3930
         Width           =   1695
      End
      Begin VB.Label generalLabel 
         Caption         =   "months"
         Height          =   255
         Index           =   24
         Left            =   1530
         TabIndex        =   68
         Top             =   1560
         Width           =   645
      End
      Begin VB.Label generalLabel 
         Caption         =   "Length:"
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   67
         Top             =   1290
         Width           =   645
      End
      Begin VB.Label generalLabel 
         Caption         =   "Warranty:"
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   66
         Top             =   2190
         Width           =   765
      End
      Begin VB.Label generalLabel 
         Caption         =   "Name:"
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   65
         Top             =   990
         Width           =   555
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Rebate Contact Info:"
      Height          =   1245
      Left            =   3750
      TabIndex        =   49
      Top             =   4770
      Width           =   3195
      Begin VB.TextBox PhoneExt 
         Height          =   285
         Index           =   3
         Left            =   2550
         TabIndex        =   23
         Top             =   270
         Width           =   555
      End
      Begin VB.CommandButton btnOpenURL 
         Caption         =   "WWW"
         Height          =   255
         Index           =   3
         Left            =   30
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   570
         Width           =   675
      End
      Begin VB.TextBox PhoneContact 
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   22
         Top             =   270
         Width           =   1515
      End
      Begin VB.TextBox URLContact 
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   24
         Top             =   570
         Width           =   2385
      End
      Begin VB.TextBox EmailContact 
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   25
         Top             =   870
         Width           =   2385
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Ext:"
         Height          =   255
         Index           =   17
         Left            =   2250
         TabIndex        =   83
         Top             =   300
         Width           =   285
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Phone:"
         Height          =   255
         Index           =   18
         Left            =   90
         TabIndex        =   51
         Top             =   300
         Width           =   555
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Email:"
         Height          =   255
         Index           =   16
         Left            =   90
         TabIndex        =   50
         Top             =   900
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parts Contact Info:"
      Height          =   1245
      Left            =   3750
      TabIndex        =   46
      Top             =   3450
      Width           =   3195
      Begin VB.TextBox PhoneExt 
         Height          =   285
         Index           =   2
         Left            =   2550
         TabIndex        =   19
         Top             =   270
         Width           =   555
      End
      Begin VB.CommandButton btnOpenURL 
         Caption         =   "WWW"
         Height          =   255
         Index           =   2
         Left            =   30
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   570
         Width           =   675
      End
      Begin VB.TextBox PhoneContact 
         Height          =   285
         Index           =   2
         Left            =   720
         TabIndex        =   18
         Top             =   270
         Width           =   1515
      End
      Begin VB.TextBox URLContact 
         Height          =   285
         Index           =   2
         Left            =   720
         TabIndex        =   20
         Top             =   570
         Width           =   2385
      End
      Begin VB.TextBox EmailContact 
         Height          =   285
         Index           =   2
         Left            =   720
         TabIndex        =   21
         Top             =   870
         Width           =   2385
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Ext:"
         Height          =   255
         Index           =   14
         Left            =   2250
         TabIndex        =   82
         Top             =   300
         Width           =   285
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Phone:"
         Height          =   255
         Index           =   15
         Left            =   90
         TabIndex        =   48
         Top             =   300
         Width           =   555
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Email:"
         Height          =   255
         Index           =   13
         Left            =   90
         TabIndex        =   47
         Top             =   900
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Service Contact Info:"
      Height          =   1245
      Left            =   3750
      TabIndex        =   43
      Top             =   2130
      Width           =   3195
      Begin VB.TextBox PhoneExt 
         Height          =   285
         Index           =   1
         Left            =   2550
         TabIndex        =   15
         Top             =   270
         Width           =   555
      End
      Begin VB.CommandButton btnOpenURL 
         Caption         =   "WWW"
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   570
         Width           =   675
      End
      Begin VB.TextBox PhoneContact 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   14
         Top             =   270
         Width           =   1515
      End
      Begin VB.TextBox URLContact 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   16
         Top             =   570
         Width           =   2385
      End
      Begin VB.TextBox EmailContact 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   17
         Top             =   870
         Width           =   2385
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Ext:"
         Height          =   255
         Index           =   11
         Left            =   2250
         TabIndex        =   81
         Top             =   300
         Width           =   285
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Phone:"
         Height          =   255
         Index           =   12
         Left            =   90
         TabIndex        =   45
         Top             =   300
         Width           =   555
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Email:"
         Height          =   255
         Index           =   10
         Left            =   90
         TabIndex        =   44
         Top             =   900
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "General Contact Info:"
      Height          =   1545
      Left            =   3750
      TabIndex        =   40
      Top             =   540
      Width           =   3195
      Begin VB.TextBox FaxContact 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   93
         Top             =   1170
         Width           =   2385
      End
      Begin VB.TextBox PhoneExt 
         Height          =   285
         Index           =   0
         Left            =   2550
         TabIndex        =   11
         Top             =   270
         Width           =   555
      End
      Begin VB.CommandButton btnOpenURL 
         Caption         =   "WWW"
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   570
         Width           =   675
      End
      Begin VB.TextBox EmailContact 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   13
         Top             =   870
         Width           =   2385
      End
      Begin VB.TextBox URLContact 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   12
         Top             =   570
         Width           =   2385
      End
      Begin VB.TextBox PhoneContact 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   10
         Top             =   270
         Width           =   1515
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Fax:"
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   92
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Ext:"
         Height          =   255
         Index           =   8
         Left            =   2250
         TabIndex        =   80
         Top             =   300
         Width           =   285
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Email:"
         Height          =   255
         Index           =   9
         Left            =   90
         TabIndex        =   42
         Top             =   900
         Width           =   555
      End
      Begin VB.Label generalLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Phone:"
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   41
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.ComboBox ParentCompany 
      Height          =   315
      Left            =   900
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3090
      Width           =   2745
   End
   Begin VB.TextBox Country 
      Height          =   285
      Left            =   900
      TabIndex        =   7
      Top             =   2700
      Width           =   2715
   End
   Begin VB.TextBox ZipCode 
      Height          =   285
      Left            =   900
      TabIndex        =   6
      Top             =   2400
      Width           =   2715
   End
   Begin VB.TextBox State 
      Height          =   285
      Left            =   900
      TabIndex        =   5
      Top             =   2100
      Width           =   2715
   End
   Begin VB.TextBox City 
      Height          =   285
      Left            =   900
      TabIndex        =   4
      Top             =   1800
      Width           =   2715
   End
   Begin VB.TextBox Address 
      Height          =   285
      Index           =   1
      Left            =   900
      TabIndex        =   3
      Top             =   1500
      Width           =   2715
   End
   Begin VB.TextBox Address 
      Height          =   285
      Index           =   0
      Left            =   900
      TabIndex        =   2
      Top             =   1200
      Width           =   2715
   End
   Begin VB.CommandButton btnAddCompany 
      Caption         =   "Add New Company"
      Height          =   315
      Left            =   3150
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   90
      Width           =   1725
   End
   Begin VB.ComboBox CompanySelector 
      Height          =   315
      Left            =   210
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   630
      Width           =   3405
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   9810
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   0
      Width           =   1155
   End
   Begin VB.CommandButton btnFirstItem 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   30
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6390
      Width           =   315
   End
   Begin VB.CommandButton btnLastItem 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2490
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6390
      Width           =   315
   End
   Begin VB.TextBox txtPosition 
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6390
      Width           =   795
   End
   Begin VB.HScrollBar barPosition 
      Height          =   285
      LargeChange     =   10
      Left            =   360
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6390
      Width           =   2115
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Parent Co:"
      Height          =   255
      Index           =   6
      Left            =   60
      TabIndex        =   39
      Top             =   3120
      Width           =   795
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Country:"
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   38
      Top             =   2730
      Width           =   795
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip Code:"
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   37
      Top             =   2430
      Width           =   795
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "State:"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   36
      Top             =   2130
      Width           =   795
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "City:"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   35
      Top             =   1830
      Width           =   795
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Address:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   34
      Top             =   1230
      Width           =   795
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   10920
      X2              =   0
      Y1              =   6330
      Y2              =   6330
   End
   Begin VB.Label lblMaxAmt 
      Caption         =   "of MAX"
      Height          =   195
      Left            =   3720
      TabIndex        =   31
      Top             =   6450
      Width           =   855
   End
   Begin VB.Label lblStatusBar 
      Height          =   255
      Left            =   4740
      TabIndex        =   30
      Top             =   6450
      Width           =   3855
   End
   Begin VB.Label generalLabel 
      Caption         =   "Company Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2685
   End
End
Attribute VB_Name = "CompanyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ItemList() As String
Private changed As Boolean
Private whichCtl As String
Private fillingForm As Boolean


Private Sub barPosition_Change()
    Mouse.Hourglass True
    If changed = True Then
        Select Case whichCtl
            Case Is = "Address0"
                Address_LostFocus 0
            Case Is = "Address1"
                Address_LostFocus 1
            Case Is = "City"
                City_LostFocus
            Case Is = "State"
                State_LostFocus
            Case Is = "ZipCode"
                ZipCode_LostFocus
            Case Is = "Country"
                Country_LostFocus
            Case Is = "PhoneContact0"
                PhoneContact_LostFocus 0
            Case Is = "PhoneContact1"
                PhoneContact_LostFocus 1
            Case Is = "PhoneContact2"
                PhoneContact_LostFocus 2
            Case Is = "PhoneContact3"
                PhoneContact_LostFocus 3
            Case Is = "PhoneExt0"
                PhoneExt_LostFocus 0
            Case Is = "PhoneExt1"
                PhoneExt_LostFocus 1
            Case Is = "PhoneExt2"
                PhoneExt_LostFocus 2
            Case Is = "PhoneExt3"
                PhoneExt_LostFocus 3
            Case Is = "URLContact0"
                URLContact_LostFocus 0
            Case Is = "URLContact1"
                URLContact_LostFocus 1
            Case Is = "URLContact2"
                URLContact_LostFocus 2
            Case Is = "URLContact3"
                URLContact_LostFocus 3
            Case Is = "EmailContact0"
                EmailContact_LostFocus 0
            Case Is = "EmailContact1"
                EmailContact_LostFocus 1
            Case Is = "EmailContact2"
                EmailContact_LostFocus 2
            Case Is = "EmailContact3"
                EmailContact_LostFocus 3
            Case Is = "FaxContact0"
                FaxContact_LostFocus 0
            Case Is = "FaxContact1"
                FaxContact_LostFocus 1
            Case Is = "FaxContact2"
                FaxContact_LostFocus 2
            Case Is = "FaxContact3"
                FaxContact_LostFocus 3
            Case Is = "WarrantyName"
                WarrantyName_LostFocus
            Case Is = "WarrantyLength"
                WarrantyLength_LostFocus
            Case Is = "WarrantyURL"
                WarrantyURL_LostFocus
            Case Is = "WarrantyInfo"
                WarrantyInfo_LostFocus
            Case Is = "GuaranteeLength"
                GuaranteeLength_LostFocus
            Case Is = "GuaranteeInfo"
                GuaranteeInfo_LostFocus
            Case Else
                MsgBox "Could not determine which control you were updating, so not saving."
        End Select
    End If
    If ItemList(Me.barPosition.value) = "NORECORDS" Then
        MsgBox "No records found."
    Else
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT * FROM ProductCompanies WHERE CompanyName='" & EscapeSQuotes(ItemList(Me.barPosition.value)) & "'")
        If rst.EOF Then
            MsgBox "Record " & ItemList(Me.barPosition.value) & " deleted by another user? requerying..."
            Dim item As String
            If Me.barPosition.value = Me.barPosition.Min Then
                item = ItemList(Me.barPosition.value + 1)
            Else
                item = ItemList(Me.barPosition.value - 1)
            End If
            requeryForm
            Me.barPosition.value = bsearch(item, ItemList)
            Mouse.Hourglass False
            Exit Sub
        End If
        fillForm rst
        rst.Close
        Set rst = Nothing
        Me.txtPosition = Me.barPosition.value + 1
    End If
    Mouse.Hourglass False
End Sub

Private Sub requeryForm()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT CompanyName FROM ProductCompanies ORDER BY CompanyName")
    Me.CompanySelector.Clear
    Me.ParentCompany.Clear
    Me.ParentCompany.AddItem "<none>"
    Dim i As Long
    ReDim ItemList(rst.RecordCount - 1) As String
    For i = 0 To rst.RecordCount - 1
        ItemList(i) = rst("CompanyName")
        Me.CompanySelector.AddItem rst("CompanyName")
        Me.ParentCompany.AddItem rst("CompanyName")
        rst.MoveNext
    Next i
    rst.Close
    Set rst = Nothing
    Me.txtPosition = "0"
    If Me.barPosition.value = 0 Then
        barPosition_Change
    Else
        Me.barPosition.value = 0
    End If
    Me.lblMaxAmt.caption = "of " & UBound(ItemList) + 1
    Me.barPosition.Min = 0
    Me.barPosition.max = UBound(ItemList)
    ExpandDropDownToFit Me.CompanySelector
    ExpandDropDownToFit Me.ParentCompany
End Sub

Private Sub fillForm(company As ADODB.Recordset)
    fillingForm = True
    Me.CompanyID = company("ID")
    Me.CompanySelector = company("CompanyName")
    Me.Address(0) = Nz(company("Address1"))
    Me.Address(1) = Nz(company("Address2"))
    Me.City = Nz(company("City"))
    Me.State = Nz(company("State"))
    Me.ZipCode = Nz(company("ZipCode"))
    Me.Country = Nz(company("Country"))
    Me.DefaultWarrantyID = Nz(company("DefaultWarrantyID"))
    If Nz(company("ParentCompanyID")) = "" Then
        Me.ParentCompany = "<none>"
    Else
        Me.ParentCompany = DLookup("CompanyName", "ProductCompanies", "ID=" & company("ParentCompanyID"))
    End If
    Me.chkSphereOne = SQLBool(company("IsSphereOne"))
    
    fillVendorInfo company("ID")
    fillContactInfo company("ID")
    fillWarrantyInfo company("ID")
    
    fillingForm = False
End Sub

Private Sub fillVendorInfo(idnum As String)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT VendorNumber, VendorName FROM ProductCompanyVendors WHERE CompanyID=" & idnum)
    Me.VendorList.Clear
    While Not rst.EOF
        Me.VendorList.AddItem rst("VendorNumber") & vbTab & rst("VendorName")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Me.btnDeleteVendor.Enabled = False
    Me.btnEditVendor.Enabled = False
End Sub

Private Sub fillContactInfo(idnum As String)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Method, Type, ContactInfo FROM ProductCompanyContactInfo WHERE CompanyID=" & idnum)
    Dim i As Long
    For i = 0 To 3
        Me.PhoneContact(i) = ""
        Me.PhoneExt(i) = ""
        Me.URLContact(i) = ""
        Me.EmailContact(i) = ""
        If i = 0 Then
            Me.FaxContact(i) = ""
        End If
    Next i
    While Not rst.EOF
        Select Case rst("Method")
            Case Is = "Phone"
                'Me.PhoneContact(rst("Type")) = rst("ContactInfo")
                If InStr(rst("ContactInfo"), " x") Then
                    Me.PhoneContact(rst("Type")) = Left(rst("ContactInfo"), InStr(rst("ContactInfo"), " x") - 1)
                    Me.PhoneExt(rst("Type")) = Mid(rst("ContactInfo"), InStr(rst("ContactInfo"), " x") + 2)
                Else
                    Me.PhoneContact(rst("Type")) = rst("ContactInfo")
                End If
            Case Is = "URL"
                Me.URLContact(rst("Type")) = rst("ContactInfo")
            Case Is = "Email"
                Me.EmailContact(rst("Type")) = rst("ContactInfo")
            Case Is = "Fax"
                Me.FaxContact(rst("Type")) = rst("ContactInfo")
        End Select
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub fillWarrantyInfo(idnum As String)
    'Dim rst As ADODB.Recordset
    'Set rst = DB.retrieve("SELECT WarrantyName, WarrantyID FROM ProductCompanyWarranties WHERE CompanyID=" & idnum)
    'Me.WarrantyList.Clear
    'While Not rst.EOF
    '    Me.WarrantyList.AddItem rst("WarrantyName") & vbTab & rst("WarrantyID")
    '    rst.MoveNext
    'Wend
    'rst.Close
    'Set rst = Nothing
    Me.WarrantyName = ""
    Me.WarrantyLength = ""
    Me.chkLifetimeWarranty = vbUnchecked
    Me.chkDefaultWarranty = vbUnchecked
    Me.WarrantyURL = ""
    Me.WarrantyInfo = ""
    Me.GuaranteeLength = ""
    Me.chkLifetimeGuarantee = vbUnchecked
    Me.GuaranteeInfo = ""
    Disable Me.WarrantyName
    Disable Me.WarrantyLength
    Disable Me.chkLifetimeWarranty
    Disable Me.chkDefaultWarranty
    Disable Me.WarrantyURL
    Disable Me.WarrantyInfo
    Disable Me.GuaranteeLength
    Disable Me.chkLifetimeGuarantee
    Disable Me.GuaranteeInfo
    
    Me.btnAddNewWarranty.Enabled = False
    Me.btnDeleteWarranty.Enabled = False
End Sub

Private Sub WarrantyList_Click()
    Dim rst As ADODB.Recordset
    Dim wid As String
    wid = Mid(Me.WarrantyList, InStr(Me.WarrantyList, vbTab) + 1)
    Set rst = DB.retrieve("SELECT WarrantyName, WarrantyInfo, WarrantyURL, WarrantyLengthMonths, SatisfactionGuaranteeInfo, SatisfactionGuaranteeLengthDays FROM ProductCompanyWarranties WHERE WarrantyID=" & wid)
    fillingForm = True
    If Not rst.EOF Then
        Me.WarrantyID = wid
        Me.WarrantyName = Nz(rst("WarrantyName"))
        Enable Me.chkLifetimeWarranty
        If Nz(rst("WarrantyLengthMonths")) = "-1" Then
            Me.chkLifetimeWarranty = vbChecked
            Me.WarrantyLength = ""
            Disable Me.WarrantyLength
        Else
            Me.chkLifetimeWarranty = vbUnchecked
            Me.WarrantyLength = Nz(rst("WarrantyLengthMonths"))
            Enable Me.WarrantyLength
        End If
        Me.chkDefaultWarranty = SQLBool(Me.DefaultWarrantyID = wid)
        Me.WarrantyURL = Nz(rst("WarrantyURL"))
        Me.WarrantyInfo = Nz(rst("WarrantyInfo"))
        Enable Me.chkLifetimeGuarantee
        If Nz(rst("SatisfactionGuaranteeLengthDays")) = "-1" Then
            Me.chkLifetimeGuarantee = vbChecked
            Me.GuaranteeLength = ""
            Disable Me.GuaranteeLength
        Else
            Me.chkLifetimeGuarantee = vbUnchecked
            Me.GuaranteeLength = Nz(rst("SatisfactionGuaranteeLengthDays"))
            Enable Me.GuaranteeLength
        End If
        Me.GuaranteeLength = Nz(rst("SatisfactionGuaranteeLengthDays"))
        Me.GuaranteeInfo = Nz(rst("SatisfactionGuaranteeInfo"))
        Enable Me.WarrantyName
        'Enable Me.WarrantyLength
        Enable Me.chkDefaultWarranty
        Enable Me.WarrantyURL
        Enable Me.WarrantyInfo
        'Enable Me.GuaranteeLength
        Enable Me.GuaranteeInfo
    Else
        MsgBox "Can't find this warranty, wtf?"
        Disable Me.WarrantyName
        Disable Me.WarrantyLength
        Disable Me.chkLifetimeWarranty
        Disable Me.chkDefaultWarranty
        Disable Me.WarrantyURL
        Disable Me.WarrantyInfo
        Disable Me.GuaranteeLength
        Disable Me.chkLifetimeGuarantee
        Disable Me.GuaranteeInfo
    End If
    fillingForm = False
    rst.Close
    Set rst = Nothing
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnFirstItem_Click()
    Me.barPosition.value = Me.barPosition.Min
End Sub

Private Sub btnLastItem_Click()
    Me.barPosition.value = Me.barPosition.max
End Sub

Private Sub Form_Load()
    Dim tabs(0) As Long
    tabs(0) = 300
    SetListTabStops Me.WarrantyList.hwnd, tabs
    requeryForm
End Sub

Private Sub btnOpenURL_Click(Index As Integer)
    Dim url As String
    url = IIf(Index = 4, Me.WarrantyURL, Me.URLContact(Index))
    If url <> "" Then
        OpenDefaultApp url
    End If
End Sub




'====================== ADDRESS & BASIC INFO ==================

Private Sub btnAddCompany_Click()
    Dim newName As String
    newName = InputBox("Enter Company Name:")
    If newName <> "" Then
        If "" <> DLookup("CompanyName", "ProductCompanies", "CompanyName='" & EscapeSQuotes(newName) & "'") Then
            MsgBox "That company already exists"
        Else
            DB.Execute "INSERT INTO ProductCompanies ( CompanyName ) VALUES ( '" & EscapeSQuotes(newName) & "' )"
            Me.CompanySelector.AddItem newName
            Me.ParentCompany.AddItem newName
            requeryForm
            Me.barPosition.value = bsearch(newName, ItemList)
        End If
    End If
End Sub

Private Sub btnDeleteCompany_Click()
    If vbYes = MsgBox("Are you sure you want to delete this company?" & vbCrLf & vbCrLf & "Any vendors, product lines, and items will be orphaned!", vbYesNo + vbDefaultButton2) Then
        DB.Execute "DELETE FROM ProductCompanies WHERE ID=" & Me.CompanyID
        requeryForm
    Else
        MsgBox "Good idea."
    End If
End Sub

Private Sub CompanySelector_Click()
    If Not fillingForm Then
        Me.barPosition.value = bsearch(Me.CompanySelector, ItemList)
    End If
End Sub

Private Sub Address_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            Address_LostFocus Index
        Case Is = vbKeyDelete
            Address_KeyPress Index, KeyCode
    End Select
End Sub

Private Sub Address_KeyPress(Index As Integer, KeyAscii As Integer)
    changed = True
    whichCtl = "Address" & Index
End Sub

Private Sub Address_LostFocus(Index As Integer)
    If changed = True Then
        If Len(Me.Address(Index)) > 50 Then
            MsgBox "Address too long! Must be under 50 chars."
            Me.Address(Index).SetFocus
        Else
            DB.Execute "UPDATE ProductCompanies SET Address" & Index + 1 & "='" & EscapeSQuotes(Me.Address(Index)) & "' WHERE ID=" & Me.CompanyID
            changed = False
        End If
    End If
End Sub

Private Sub City_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            City_LostFocus
        Case Is = vbKeyDelete
            City_KeyPress KeyCode
    End Select
End Sub

Private Sub City_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "City"
End Sub

Private Sub City_LostFocus()
    If changed = True Then
        If Len(Me.City) > 50 Then
            MsgBox "City too long! Must be under 50 chars."
            Me.City.SetFocus
        Else
            DB.Execute "UPDATE ProductCompanies SET City='" & EscapeSQuotes(Me.City) & "' WHERE ID=" & Me.CompanyID
            changed = False
        End If
    End If
End Sub

Private Sub State_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            State_LostFocus
        Case Is = vbKeyDelete
            State_KeyPress KeyCode
    End Select
End Sub

Private Sub State_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "State"
End Sub

Private Sub State_LostFocus()
    If changed = True Then
        If Len(Me.State) > 16 Then
            MsgBox "State too long! Must be under 50 chars."
            Me.State.SetFocus
        Else
            DB.Execute "UPDATE ProductCompanies SET State='" & EscapeSQuotes(Me.State) & "' WHERE ID=" & Me.CompanyID
            changed = False
        End If
    End If
End Sub

Private Sub ZipCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            ZipCode_LostFocus
        Case Is = vbKeyDelete
            ZipCode_KeyPress KeyCode
    End Select
End Sub

Private Sub ZipCode_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "ZipCode"
End Sub

Private Sub ZipCode_LostFocus()
    If changed = True Then
        If Len(Me.ZipCode) > 10 Then
            MsgBox "Zipcode too long! Must be under 50 chars."
            Me.ZipCode.SetFocus
        Else
            DB.Execute "UPDATE ProductCompanies SET ZipCode='" & EscapeSQuotes(Me.ZipCode) & "' WHERE ID=" & Me.CompanyID
            changed = False
        End If
    End If
End Sub

Private Sub Country_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            Country_LostFocus
        Case Is = vbKeyDelete
            Country_KeyPress KeyCode
    End Select
End Sub

Private Sub Country_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "Country"
End Sub

Private Sub Country_LostFocus()
    If changed = True Then
        If Len(Me.Country) > 50 Then
            MsgBox "Country too long! Must be under 50 chars."
            Me.Country.SetFocus
        Else
            DB.Execute "UPDATE ProductCompanies SET Country='" & EscapeSQuotes(Me.Country) & "' WHERE ID=" & Me.CompanyID
            changed = False
        End If
    End If
End Sub

Private Sub ParentCompany_Click()
    If Not fillingForm Then
        If Me.ParentCompany = Me.CompanySelector Then
            Me.ParentCompany = "<none>"
        ElseIf Me.ParentCompany = "<none>" Then
            DB.Execute "UPDATE ProductCompanies SET ParentCompanyID=NULL WHERE ID=" & Me.CompanyID
        Else
            DB.Execute "UPDATE ProductCompanies SET ParentCompanyID=" & DLookup("ID", "ProductCompanies", "CompanyName='" & EscapeSQuotes(Me.ParentCompany) & "'") & " WHERE ID=" & Me.CompanyID
        End If
    End If
End Sub

Private Sub chkSphereOne_Click()
    If Not fillingForm Then
        DB.Execute "UPDATE ProductCompanies SET IsSphereOne=" & SQLBool(Me.chkSphereOne) & " WHERE ID=" & Me.CompanyID
    End If
End Sub



'====================== CHILD VENDORS ==================

Private Sub VendorList_Click()
    'Me.btnDeleteVendor.Enabled = True
    Me.btnEditVendor.Enabled = True
End Sub

Private Sub btnAddVendor_Click()
    Dim vnum As String
    vnum = InputBox("Enter new vendor number:")
    If Len(vnum) > 7 Then
        MsgBox "Vendor number needs to be 7 characters or fewer!"
    ElseIf 1 = DLookup("COUNT(*)", "ProductCompanyVendors", "VendorNumber='" & EscapeSQuotes(vnum) & "'") Then
        MsgBox "This vendor number is already in use!"
    Else
        DB.Execute "INSERT INTO ProductCompanyVendors ( CompanyID, VendorNumber ) VALUES ( " & Me.CompanyID & ", '" & UCase(vnum) & "' )"
        Me.VendorList.AddItem UCase(vnum) & vbTab & ""
        Dim i As Long
        For i = 0 To Me.VendorList.ListCount - 1
            If Me.VendorList.list(i) = UCase(vnum) & vbTab & "" Then
                Me.VendorList.Selected(i) = True
                btnEditVendor_Click
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub btnDeleteVendor_Click()
    If Me.VendorList.ListIndex < 0 Then
        MsgBox "This button should have been disabled, how did you get here? Tell Brian."
    ElseIf vbYes = MsgBox("Are you sure you want to delete this vendor?" & vbCrLf & vbCrLf & "Any product lines and items for it will be orphaned!", vbYesNo + vbDefaultButton2) Then
        Dim vnum As String
        vnum = Left(Me.VendorList, InStr(Me.VendorList, vbTab) - 1)
        DB.Execute "DELETE FROM ProductCompanyVendors WHERE VendorNumber='" & vnum & "'"
        Me.VendorList.RemoveItem Me.VendorList.ListIndex
        If Me.VendorList.ListIndex < 0 Then
            Me.btnDeleteVendor.Enabled = False
            Me.btnEditVendor.Enabled = False
        End If
    Else
        MsgBox "Good idea."
    End If
End Sub

Private Sub btnEditVendor_Click()
    Dim vnum As String
    vnum = Left(Me.VendorList, InStr(Me.VendorList, vbTab) - 1)
    Load VendorInfo
    VendorInfo.LoadVendor vnum
    VendorInfo.Show
End Sub



'====================== CONTACT INFO ==================

Private Sub PhoneContact_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Shift And vbCtrlMask Then 'bitwise and
        If KeyCode = vbKeyV Then
            Dim clip As String
            clip = ClipBoard_GetData
            Dim i As Long
            Dim thiskey As Integer
            Dim cpos As Long
            cpos = Me.PhoneContact(Index).selStart
            If Me.PhoneContact(Index).selLength <> 0 Then
                Me.PhoneContact(Index) = Left(Me.PhoneContact(Index), Me.PhoneContact(Index).selStart) & Mid(Me.PhoneContact(Index), Me.PhoneContact(Index).selStart + Me.PhoneContact(Index).selLength + 1)
            End If
            For i = 0 To Len(clip) - 1
                thiskey = Asc(Mid(clip, i + 1, 1))
                'pass by ref, thiskey comes back same or 0 if invalid char
                PhoneContact_KeyPress Index, thiskey
                If thiskey <> 0 Then
                    Me.PhoneContact(Index) = Left(Me.PhoneContact(Index), cpos) & Chr(thiskey) & Mid(Me.PhoneContact(Index), cpos + 1)
                    cpos = cpos + 1
                End If
            Next i
            Me.PhoneContact(Index).selStart = cpos
            PhoneContact_LostFocus Index
        ElseIf KeyCode = vbKeyC Then
            If Me.PhoneContact(Index).SelText <> "" Then
                ClipBoard_SetData Me.PhoneContact(Index).SelText
            End If
        End If
    End If
    Select Case KeyCode
        Case Is = vbKeyReturn
            PhoneContact_LostFocus Index
        Case Is = vbKeyDelete
            PhoneContact_KeyPress Index, KeyCode
    End Select
End Sub

Private Sub PhoneContact_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = Asc("-") Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = Asc(" ") Then
        'ok
    ElseIf KeyAscii = Asc("+") Then
        'ok, foreign
    ElseIf KeyAscii = vbKeyReturn Or KeyAscii = vbKeyDelete Then
        'ok, control
    Else
        KeyAscii = LimitToNumbers(KeyAscii)
    End If
    If KeyAscii <> 0 Then
        changed = True
        whichCtl = "PhoneContact" & Index
    End If
End Sub

Private Sub PhoneContact_LostFocus(Index As Integer)
    If changed = True Then
        Dim fmtnumber As String
        fmtnumber = Replace(Replace(Replace(Replace(Me.PhoneContact(Index), " ", ""), ")", ""), "(", ""), "-", "")
        If Len(fmtnumber) = 11 And Left(fmtnumber, 1) = "1" Then
            fmtnumber = Mid(fmtnumber, 2)
        End If
        
        If Len(fmtnumber) = 10 And Not CBool(InStr(fmtnumber, "+")) Then
            'ok, use this
            Me.PhoneContact(Index) = Left(fmtnumber, 3) & "-" & Mid(fmtnumber, 4, 3) & "-" & Right(fmtnumber, 4)
        Else
            'forget it, just save the un-formatted number
        End If
        
        If Me.PhoneContact(Index) = "" And Me.PhoneExt(Index) = "" Then
            DB.Execute "DELETE FROM ProductCompanyContactInfo WHERE CompanyID=" & Me.CompanyID & " AND Method='Phone' AND Type=" & Index
            fmtnumber = ""
        ElseIf Me.PhoneExt(Index) = "" Then
            fmtnumber = Me.PhoneContact(Index)
        ElseIf Me.PhoneContact(Index) = "" Then
            'delete this too
            DB.Execute "DELETE FROM ProductCompanyContactInfo WHERE CompanyID=" & Me.CompanyID & " AND Method='Phone' AND Type=" & Index
            fmtnumber = ""
        Else
            fmtnumber = Me.PhoneContact(Index) & " x" & Me.PhoneExt(Index)
        End If
        changed = False
        If Len(fmtnumber) > 50 Then
            changed = True
            MsgBox "Phone number too long. Phone number + extension must be under 50 chars."
            Me.PhoneContact(Index).SetFocus
        ElseIf fmtnumber <> "" Then
            If "0" = DLookup("COUNT(*)", "ProductCompanyContactInfo", "CompanyID=" & Me.CompanyID & " AND Method='Phone' AND Type=" & Index) Then
                DB.Execute "INSERT INTO ProductCompanyContactInfo ( CompanyID, Method, Type, ContactInfo ) VALUES ( " & Me.CompanyID & ", 'Phone', " & Index & ", '" & EscapeSQuotes(fmtnumber) & "' )"
            Else
                DB.Execute "UPDATE ProductCompanyContactInfo SET ContactInfo='" & EscapeSQuotes(fmtnumber) & "' WHERE CompanyID=" & Me.CompanyID & " AND Method='Phone' AND Type=" & Index
            End If
        End If
    End If
End Sub

Private Sub PhoneExt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            PhoneExt_LostFocus Index
        Case Is = vbKeyDelete
            PhoneExt_KeyPress Index, KeyCode
    End Select
End Sub

Private Sub PhoneExt_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
    If KeyAscii <> 0 Then
        changed = True
        whichCtl = "PhoneExt" & Index
    End If
End Sub

Private Sub PhoneExt_LostFocus(Index As Integer)
    If changed = True Then
        Dim phonestr As String
        If Me.PhoneExt(Index) = "" And Me.PhoneContact(Index) = "" Then
            DB.Execute "DELETE FROM ProductCompanyContactInfo WHERE CompanyID=" & Me.CompanyID & " AND Method='Phone' AND Type=" & Index
            phonestr = ""
        ElseIf Me.PhoneExt(Index) = "" Then
            phonestr = Me.PhoneContact(Index)
        ElseIf Me.PhoneContact(Index) = "" Then
            'extension but no phone, so we'll not save anything, if they add the phone in
            'next, then it will save both
            phonestr = ""
        Else
            phonestr = Me.PhoneContact(Index) & " x" & Me.PhoneExt(Index)
        End If
        changed = False
        If Len(phonestr) > 50 Then
            changed = True
            MsgBox "Phone number too long. Phone number + extension must be under 50 chars."
            Me.PhoneExt(Index).SetFocus
        ElseIf phonestr <> "" Then
            If "0" = DLookup("COUNT(*)", "ProductCompanyContactInfo", "CompanyID=" & Me.CompanyID & " AND Method='Phone' AND Type=" & Index) Then
                DB.Execute "INSERT INTO ProductCompanyContactInfo ( CompanyID, Method, Type, ContactInfo ) VALUES ( " & Me.CompanyID & ", 'Phone', " & Index & ", '" & EscapeSQuotes(phonestr) & "' )"
            Else
                DB.Execute "UPDATE ProductCompanyContactInfo SET ContactInfo='" & EscapeSQuotes(phonestr) & "' WHERE CompanyID=" & Me.CompanyID & " AND Method='Phone' AND Type=" & Index
            End If
        End If
    End If
End Sub

Private Sub URLContact_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            URLContact_LostFocus Index
        Case Is = vbKeyDelete
            URLContact_KeyPress Index, KeyCode
    End Select
End Sub

Private Sub URLContact_KeyPress(Index As Integer, KeyAscii As Integer)
    changed = True
    whichCtl = "URLContact" & Index
End Sub

Private Sub URLContact_LostFocus(Index As Integer)
    If changed = True Then
        If Len(Me.URLContact(Index)) > 128 Then
            MsgBox "Contact URL too long! Must be under 128 chars"
            Me.URLContact(Index).SetFocus
        Else
            If Me.URLContact(Index) = "" Then
                DB.Execute "DELETE FROM ProductCompanyContactInfo WHERE CompanyID=" & Me.CompanyID & " AND Method='URL' AND Type=" & Index
            ElseIf "0" = DLookup("COUNT(*)", "ProductCompanyContactInfo", "CompanyID=" & Me.CompanyID & " AND Method='URL' AND Type=" & Index) Then
                DB.Execute "INSERT INTO ProductCompanyContactInfo ( CompanyID, Method, Type, ContactInfo ) VALUES ( " & Me.CompanyID & ", 'URL', " & Index & ", '" & EscapeSQuotes(Me.URLContact(Index)) & "' )"
            Else
                DB.Execute "UPDATE ProductCompanyContactInfo SET ContactInfo='" & EscapeSQuotes(Me.URLContact(Index)) & "' WHERE CompanyID=" & Me.CompanyID & " AND Method='URL' AND Type=" & Index
            End If
            changed = False
        End If
    End If
End Sub

Private Sub EmailContact_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            EmailContact_LostFocus Index
        Case Is = vbKeyDelete
            EmailContact_KeyPress Index, KeyCode
    End Select
End Sub

Private Sub EmailContact_KeyPress(Index As Integer, KeyAscii As Integer)
    changed = True
    whichCtl = "EmailContact" & Index
End Sub

Private Sub EmailContact_LostFocus(Index As Integer)
    If changed = True Then
        If Len(Me.EmailContact(Index)) > 128 Then
            MsgBox "Contact email too long! Must be under 128 chars."
            Me.EmailContact(Index).SetFocus
        Else
            If Me.EmailContact(Index) = "" Then
                DB.Execute "DELETE FROM ProductCompanyContactInfo WHERE CompanyID=" & Me.CompanyID & " AND Method='Email' AND Type=" & Index
            ElseIf "0" = DLookup("COUNT(*)", "ProductCompanyContactInfo", "CompanyID=" & Me.CompanyID & " AND Method='Email' AND Type=" & Index) Then
                DB.Execute "INSERT INTO ProductCompanyContactInfo ( CompanyID, Method, Type, ContactInfo ) VALUES ( " & Me.CompanyID & ", 'Email', " & Index & ", '" & EscapeSQuotes(Me.EmailContact(Index)) & "' )"
            Else
                DB.Execute "UPDATE ProductCompanyContactInfo SET ContactInfo='" & EscapeSQuotes(Me.EmailContact(Index)) & "' WHERE CompanyID=" & Me.CompanyID & " AND Method='Email' AND Type=" & Index
            End If
            changed = False
        End If
    End If
End Sub

Private Sub FaxContact_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            FaxContact_LostFocus Index
        Case Is = vbKeyDelete
            FaxContact_KeyPress Index, KeyCode
    End Select
End Sub

Private Sub FaxContact_KeyPress(Index As Integer, KeyAscii As Integer)
    changed = True
    whichCtl = "FaxContact" & Index
End Sub

Private Sub FaxContact_LostFocus(Index As Integer)
    If changed = True Then
        If Len(Me.FaxContact(Index)) > 128 Then
            MsgBox "Contact fax too long! Must be under 128 chars."
            Me.FaxContact(Index).SetFocus
        Else
            If Me.FaxContact(Index) = "" Then
                DB.Execute "DELETE FROM ProductCompanyContactInfo WHERE CompanyID=" & Me.CompanyID & " AND Method='Fax' AND Type=" & Index
            ElseIf "0" = DLookup("COUNT(*)", "ProductCompanyContactInfo", "CompanyID=" & Me.CompanyID & " AND Method='Fax' AND Type=" & Index) Then
                DB.Execute "INSERT INTO ProductCompanyContactInfo ( CompanyID, Method, Type, ContactInfo ) VALUES ( " & Me.CompanyID & ", 'Fax', " & Index & ", '" & EscapeSQuotes(Me.FaxContact(Index)) & "' )"
            Else
                DB.Execute "UPDATE ProductCompanyContactInfo SET ContactInfo='" & EscapeSQuotes(Me.FaxContact(Index)) & "' WHERE CompanyID=" & Me.CompanyID & " AND Method='Fax' AND Type=" & Index
            End If
            changed = False
        End If
    End If
End Sub




'====================== WARRANTY INFO ==================

Private Sub btnAddNewWarranty_Click()
    Dim wname As String
    wname = InputBox("Enter Warranty Name:")
    If wname <> "" Then
        DB.Execute "INSERT INTO ProductCompanyWarranties ( CompanyID, WarrantyName ) VALUES ( " & Me.CompanyID & ", '" & EscapeSQuotes(wname) & "' )"
        Dim newid As String
        newid = DLookup("@@IDENTITY", "ProductCompanyWarranties")
        Me.WarrantyList.AddItem wname & vbTab & newid
        Me.WarrantyList = wname & vbTab & newid
    End If
End Sub

Private Sub btnDeleteWarranty_Click()
    If Me.WarrantyList <> "" Then
        If vbYes = MsgBox("Are you sure you want to delete this warranty?", vbYesNo) Then
            DB.Execute "DELETE FROM ProductCompanyWarranties WHERE WarrantyID=" & Me.WarrantyID
            If Me.WarrantyID = DLookup("DefaultWarrantyID", "ProductCompanies", "ID=" & Me.CompanyID) Then
                DB.Execute "UPDATE ProductCompanies SET DefaultWarrantyID=NULL WHERE ID=" & Me.CompanyID
                Me.DefaultWarrantyID = ""
            End If
            fillingForm = True
            fillWarrantyInfo Me.CompanyID
            fillingForm = False
        End If
    End If
End Sub

Private Sub WarrantyName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            WarrantyName_LostFocus
        Case Is = vbKeyDelete
            WarrantyName_KeyPress KeyCode
    End Select
End Sub

Private Sub WarrantyName_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "WarrantyName"
End Sub

Private Sub WarrantyName_LostFocus()
    If changed = True Then
        If Len(Me.WarrantyName) > 50 Then
            MsgBox "Warranty name too long! Must be under 50 chars."
            Me.WarrantyName.SetFocus
        Else
            Dim oldname As String
            oldname = DLookup("WarrantyName", "ProductCompanyWarranties", "WarrantyID=" & Me.WarrantyID)
            DB.Execute "UPDATE ProductCompanyWarranties SET WarrantyName='" & EscapeSQuotes(Me.WarrantyName) & "' WHERE WarrantyID=" & Me.WarrantyID
            Me.WarrantyList.list(Me.WarrantyList.ListIndex) = Me.WarrantyName & vbTab & Me.WarrantyID
            changed = False
        End If
    End If
End Sub

Private Sub WarrantyLength_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            WarrantyLength_LostFocus
        Case Is = vbKeyDelete
            WarrantyLength_KeyPress KeyCode
    End Select
End Sub

Private Sub WarrantyLength_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
    If KeyAscii <> 0 Then
        changed = True
        whichCtl = "WarrantyLength"
    End If
End Sub

Private Sub WarrantyLength_LostFocus()
    If changed = True Then
        Dim w As String
        If Me.WarrantyLength = "" Then
            w = "NULL"
        Else
            w = Me.WarrantyLength
        End If
        DB.Execute "UPDATE ProductCompanyWarranties SET WarrantyLengthMonths=" & w & " WHERE WarrantyID=" & Me.WarrantyID
        changed = False
    End If
End Sub

Private Sub chkDefaultWarranty_Click()
    If Not fillingForm Then
        DB.Execute "UPDATE ProductCompanies SET DefaultWarrantyID=" & IIf(Me.chkDefaultWarranty, Me.WarrantyID, "NULL") & " WHERE ID=" & Me.CompanyID
        Me.DefaultWarrantyID = IIf(Me.chkDefaultWarranty, Me.WarrantyID, "")
    End If
End Sub

Private Sub WarrantyURL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            WarrantyURL_LostFocus
        Case Is = vbKeyDelete
            WarrantyURL_KeyPress KeyCode
    End Select
End Sub

Private Sub WarrantyURL_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "WarrantyURL"
End Sub

Private Sub WarrantyURL_LostFocus()
    If changed = True Then
        If Len(Me.WarrantyURL) > 512 Then
            MsgBox "Warranty URL too long! Must be under 512 chars."
            Me.WarrantyURL.SetFocus
        Else
            DB.Execute "UPDATE ProductCompanyWarranties SET WarrantyURL='" & EscapeSQuotes(Me.WarrantyURL) & "' WHERE WarrantyID=" & Me.WarrantyID
            changed = False
        End If
    End If
End Sub

Private Sub WarrantyInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            WarrantyInfo_LostFocus
        Case Is = vbKeyDelete
            WarrantyInfo_KeyPress KeyCode
    End Select
End Sub

Private Sub WarrantyInfo_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "WarrantyInfo"
End Sub

Private Sub WarrantyInfo_LostFocus()
    If changed = True Then
        DB.Execute "UPDATE ProductCompanyWarranties SET WarrantyInfo='" & EscapeSQuotes(Me.WarrantyInfo) & "' WHERE WarrantyID=" & Me.WarrantyID
        changed = False
    End If
End Sub

Private Sub GuaranteeLength_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            GuaranteeLength_LostFocus
        Case Is = vbKeyDelete
            GuaranteeLength_KeyPress KeyCode
    End Select
End Sub

Private Sub GuaranteeLength_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
    If KeyAscii <> 0 Then
        changed = True
        whichCtl = "GuaranteeLength"
    End If
End Sub

Private Sub GuaranteeLength_LostFocus()
    If changed = True Then
        Dim g As String
        If Me.GuaranteeLength = "" Then
            g = "NULL"
        Else
            g = Me.GuaranteeLength
        End If
        DB.Execute "UPDATE ProductCompanyWarranties SET SatisfactionGuaranteeLengthDays=" & g & " WHERE WarrantyID=" & Me.WarrantyID
        changed = False
    End If
End Sub

Private Sub GuaranteeInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            GuaranteeInfo_LostFocus
        Case Is = vbKeyDelete
            GuaranteeInfo_KeyPress KeyCode
    End Select
End Sub

Private Sub GuaranteeInfo_KeyPress(KeyAscii As Integer)
    changed = True
    whichCtl = "GuaranteeInfo"
End Sub

Private Sub GuaranteeInfo_LostFocus()
    If changed = True Then
        DB.Execute "UPDATE ProductCompanyWarranties SET SatisfactionGuaranteeInfo='" & EscapeSQuotes(Me.GuaranteeInfo) & "' WHERE WarrantyID=" & Me.WarrantyID
        changed = False
    End If
End Sub

Private Sub chkLifetimeWarranty_Click()
    If Not fillingForm Then
        If Me.chkLifetimeWarranty Then
            DB.Execute "UPDATE ProductCompanyWarranties SET WarrantyLengthMonths=-1 WHERE WarrantyID=" & Me.WarrantyID
            Me.WarrantyLength = ""
            Disable Me.WarrantyLength
        Else
            DB.Execute "UPDATE ProductCompanyWarranties SET WarrantyLengthMonths=NULL WHERE WarrantyID=" & Me.WarrantyID
            Enable Me.WarrantyLength
        End If
    End If
End Sub

Private Sub chkLifetimeGuarantee_Click()
    If Not fillingForm Then
        If Me.chkLifetimeGuarantee Then
            DB.Execute "UPDATE ProductCompanyWarranties SET SatisfactionGuaranteeLengthDays=-1 WHERE WarrantyID=" & Me.WarrantyID
            Me.GuaranteeLength = ""
            Disable Me.GuaranteeLength
        Else
            DB.Execute "UPDATE ProductCompanyWarranties SET SatisfactionGuaranteeLengthDays=NULL WHERE WarrantyID=" & Me.WarrantyID
            Enable Me.GuaranteeLength
        End If
    End If
End Sub

