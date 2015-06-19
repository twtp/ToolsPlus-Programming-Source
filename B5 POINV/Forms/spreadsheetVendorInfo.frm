VERSION 5.00
Begin VB.Form VendorInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendor Information"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7350
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox CoopVendor 
      Alignment       =   1  'Right Justify
      Caption         =   "Coop Vendor:"
      Height          =   255
      Left            =   300
      TabIndex        =   35
      Top             =   2100
      Width           =   1275
   End
   Begin VB.TextBox AccountNumber 
      Height          =   285
      Left            =   1380
      TabIndex        =   33
      Top             =   1710
      Width           =   2685
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   405
      Left            =   2730
      TabIndex        =   9
      Top             =   4290
      Width           =   1905
   End
   Begin VB.Frame Frame2 
      Caption         =   "Vendor Product Lines:"
      Height          =   1605
      Left            =   120
      TabIndex        =   8
      Top             =   2580
      Width           =   3945
      Begin VB.CommandButton btnEditProductLine 
         Caption         =   "Edit"
         Height          =   285
         Left            =   3030
         TabIndex        =   32
         Top             =   1080
         Width           =   795
      End
      Begin VB.CommandButton btnDeleteProductLine 
         Caption         =   "Del"
         Height          =   285
         Left            =   3030
         TabIndex        =   31
         Top             =   780
         Width           =   795
      End
      Begin VB.CommandButton btnAddProductLine 
         Caption         =   "Add"
         Height          =   285
         Left            =   3030
         TabIndex        =   30
         Top             =   480
         Width           =   795
      End
      Begin VB.ListBox ProductLineList 
         Height          =   1230
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   270
         Width           =   2865
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Company Information:"
      Height          =   3675
      Left            =   4140
      TabIndex        =   7
      Top             =   60
      Width           =   3075
      Begin VB.CommandButton btnCompanyEmail 
         Caption         =   "Email:"
         Height          =   255
         Left            =   150
         TabIndex        =   28
         Top             =   3180
         Width           =   675
      End
      Begin VB.CommandButton btnCompanyWWW 
         Caption         =   "WWW:"
         Height          =   255
         Left            =   150
         TabIndex        =   27
         Top             =   2910
         Width           =   675
      End
      Begin VB.TextBox CompanyEmail 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   3180
         Width           =   2115
      End
      Begin VB.TextBox CompanyURL 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2880
         Width           =   2115
      End
      Begin VB.TextBox CompanyPhone 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2580
         Width           =   2115
      End
      Begin VB.TextBox CompanyCountry 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2160
         Width           =   2115
      End
      Begin VB.TextBox CompanyZipCode 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1860
         Width           =   2115
      End
      Begin VB.TextBox CompanyState 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1560
         Width           =   2115
      End
      Begin VB.TextBox CompanyCity 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1260
         Width           =   2115
      End
      Begin VB.TextBox CompanyAddress2 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   960
         Width           =   2115
      End
      Begin VB.TextBox CompanyAddress1 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   660
         Width           =   2115
      End
      Begin VB.TextBox CompanyName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   2115
      End
      Begin VB.Label genericLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Phone:"
         Height          =   255
         Index           =   9
         Left            =   90
         TabIndex        =   23
         Top             =   2610
         Width           =   705
      End
      Begin VB.Label genericLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Country:"
         Height          =   255
         Index           =   8
         Left            =   90
         TabIndex        =   15
         Top             =   2190
         Width           =   705
      End
      Begin VB.Label genericLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Zip Code:"
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   14
         Top             =   1890
         Width           =   705
      End
      Begin VB.Label genericLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "State:"
         Height          =   255
         Index           =   6
         Left            =   90
         TabIndex        =   13
         Top             =   1590
         Width           =   705
      End
      Begin VB.Label genericLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "City:"
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   12
         Top             =   1290
         Width           =   705
      End
      Begin VB.Label genericLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Address:"
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   11
         Top             =   690
         Width           =   705
      End
      Begin VB.Label genericLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   10
         Top             =   390
         Width           =   705
      End
   End
   Begin VB.ComboBox PurchaseTerms 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1350
      Width           =   2685
   End
   Begin VB.TextBox VendorName 
      Height          =   285
      Left            =   1380
      TabIndex        =   5
      Top             =   1020
      Width           =   2685
   End
   Begin VB.TextBox VendorNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "ERROR"
      Top             =   690
      Width           =   1395
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Account #:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   34
      Top             =   1740
      Width           =   1215
   End
   Begin VB.Label genericLabel 
      Caption         =   "Purchase Terms:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label genericLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Vendor Name:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1050
      Width           =   1215
   End
   Begin VB.Label genericLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Vendor Number:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label generalLabel 
      Caption         =   "Vendor Information"
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
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   3045
   End
End
Attribute VB_Name = "VendorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private changed As Boolean
Private fillingform As Boolean


Public Sub LoadVendor(vnum As String)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT CompanyID, VendorName, TermsCode, IsCoopVendor, AccountNumber FROM ProductCompanyVendors WHERE VendorNumber='" & vnum & "'")
    If rst.EOF Then
        'error, what now?
    Else
        fillingform = True
        Me.VendorNumber = vnum
        Me.VendorName = Nz(rst("VendorName"))
        If Nz(rst("TermsCode")) = "" Then
            Me.PurchaseTerms = "<none>"
        Else
            Me.PurchaseTerms = lookupPOTermDescByNum(rst("TermsCode"))
        End If
        Me.CoopVendor = SQLBool(rst("IsCoopVendor"))
        Me.AccountNumber = Nz(rst("AccountNumber"))
        fillProductLineInfo
        fillCompanyInfo rst("CompanyID")
        fillingform = False
    End If
End Sub

Private Sub btnAddProductLine_Click()
    Dim PL As String
    PL = InputBox("Enter new product line:")
    If validateProductLine(PL) Then
        If PLExists(PL) Then
            MsgBox "This product line code is already in use!"
        Else
            DB.Execute "INSERT INTO ProductLine ( ProductLine, PrimaryVendorNumber ) VALUES ( '" & EscapeSQuotes(UCase(PL)) & "', '" & EscapeSQuotes(Me.VendorNumber) & "' )"
            Me.ProductLineList.AddItem UCase(PL) & vbTab & ""
            Dim i As Long
            For i = 0 To Me.ProductLineList.ListCount - 1
                If Me.ProductLineList.list(i) = UCase(PL) & vbTab & "" Then
                    Me.ProductLineList.Selected(i) = True
                    btnEditProductLine_Click
                    Exit For
                End If
            Next i
        End If
    Else
        MsgBox "The product line must be 3 characters!"
    End If
End Sub

Private Sub btnDeleteProductLine_Click()
    If Me.ProductLineList.ListIndex < 0 Then
        MsgBox "This button should have been disabled, how did you get here? Tell Brian."
    ElseIf vbYes = MsgBox("Are you sure you want to delete this product line?" & vbCrLf & vbCrLf & "Any items for it will be orphaned!", vbYesNo + vbDefaultButton2) Then
        Dim PL As String
        PL = Left(Me.ProductLineList, 3)
        DB.Execute "DELETE FROM ProductLine WHERE ProductLine='" & EscapeSQuotes(PL) & "'"
        Me.ProductLineList.RemoveItem Me.ProductLineList.ListIndex
        If Me.ProductLineList.ListIndex < 0 Then
            Me.btnDeleteProductLine.Enabled = False
            Me.btnEditProductLine.Enabled = False
        End If
    Else
        MsgBox "Good idea."
    End If
End Sub

Private Sub btnEditProductLine_Click()
    Dim PL As String
    PL = Left(Me.ProductLineList, 3)
    Load AddLineCode
    AddLineCode.LoadProductLine PL
    AddLineCode.Show
End Sub

Private Sub btnCompanyEmail_Click()
    If Me.CompanyEmail <> "" Then
        OpenEmailTo Me.CompanyEmail
    End If
End Sub

Private Sub btnCompanyWWW_Click()
    If Me.CompanyURL <> "" Then
        OpenDefaultApp Me.CompanyURL
    End If
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    requeryPurchaseTerms
    'requeryCoopVendors
End Sub

Private Sub requeryPurchaseTerms()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT TermsCodeDesc FROM AP_TermsCode ORDER BY TermsCode")
    Me.PurchaseTerms.Clear
    Me.PurchaseTerms.AddItem "<none>"
    While Not rst.EOF
        Me.PurchaseTerms.AddItem rst("TermsCodeDesc")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub fillProductLineInfo()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ProductLine, ManufFullNameCleaned FROM ProductLine WHERE PrimaryVendorNumber='" & EscapeSQuotes(Me.VendorNumber) & "'")
    Me.ProductLineList.Clear
    While Not rst.EOF
        Me.ProductLineList.AddItem rst("ProductLine") & vbTab & rst("ManufFullNameCleaned")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Me.btnDeleteProductLine.Enabled = False
    Me.btnEditProductLine.Enabled = False
End Sub

Private Sub fillCompanyInfo(cid As String)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT CompanyName, Address1, Address2, City, State, ZipCode, Country FROM ProductCompanies WHERE ID=" & cid)
    If rst.EOF Then
        Me.CompanyName = ""
        Me.CompanyAddress1 = ""
        Me.CompanyAddress2 = ""
        Me.CompanyCity = ""
        Me.CompanyState = ""
        Me.CompanyZipCode = ""
        Me.CompanyCountry = ""
    Else
        Me.CompanyName = Nz(rst("CompanyName"))
        Me.CompanyAddress1 = Nz(rst("Address1"))
        Me.CompanyAddress2 = Nz(rst("Address2"))
        Me.CompanyCity = Nz(rst("City"))
        Me.CompanyState = Nz(rst("State"))
        Me.CompanyZipCode = Nz(rst("ZipCode"))
        Me.CompanyCountry = Nz(rst("Country"))
    End If
    rst.Close
    
    Set rst = DB.retrieve("SELECT Method, ContactInfo FROM ProductCompanyContactInfo WHERE CompanyID=" & cid & " AND Type=0")
    Me.CompanyPhone = ""
    Me.CompanyURL = ""
    Me.CompanyEmail = ""
    While Not rst.EOF
        Select Case rst("Method")
            Case Is = "Phone"
                Me.CompanyPhone = Nz(rst("ContactInfo"))
            Case Is = "URL"
                Me.CompanyURL = Nz(rst("ContactInfo"))
            Case Is = "Email"
                Me.CompanyEmail = Nz(rst("ContactInfo"))
        End Select
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub


Private Sub ProductLineList_Click()
    'me.btnDeleteProductLine.Enabled = True
    Me.btnEditProductLine.Enabled = True
End Sub

Private Sub VendorName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            VendorName_LostFocus
        Case Is = vbKeyDelete
            VendorName_KeyPress KeyCode
    End Select
End Sub

Private Sub VendorName_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub VendorName_LostFocus()
    If changed Then
        If Len(Me.VendorName) > 30 Then
            Me.VendorName = Left(Me.VendorName, 30)
        End If
        DB.Execute "UPDATE ProductCompanyVendors SET VendorName='" & EscapeSQuotes(Me.VendorName) & "' WHERE VendorNumber='" & Me.VendorNumber & "'"
        changed = False
    End If
End Sub

Private Sub PurchaseTerms_Click()
    DB.Execute "UPDATE ProductCompanyVendors SET TermsCode='" & lookupPOTermsByDesc(Me.PurchaseTerms) & "' WHERE VendorNumber='" & Me.VendorNumber & "'"
End Sub

Private Sub CoopVendor_Click()
    If Not fillingform Then
        DB.Execute "UPDATE ProductCompanyVendors SET Coop=" & IIf(Me.CoopVendor, "1", "0") & " WHERE VendorNumber=" & Me.VendorNumber
    End If
End Sub

'Private Sub requeryCoopVendors()
'    Me.CoopVendor.Clear
'    Me.CoopVendor.AddItem "<none>"
'    Me.CoopVendor.AddItem "Truserv (TRUS01)"
'End Sub

Private Sub AccountNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            AccountNumber_LostFocus
        Case Is = vbKeyDelete
            AccountNumber_KeyPress KeyCode
    End Select
End Sub

Private Sub AccountNumber_KeyPress(KeyAscii As Integer)
    changed = True
End Sub

Private Sub AccountNumber_LostFocus()
    If changed = True Then
        If Len(Me.AccountNumber) > 20 Then
            MsgBox "Account number too long! Must be under 50 chars."
            Me.AccountNumber.SetFocus
        Else
            DB.Execute "UPDATE ProductCompanyVendors SET AccountNumber='" & EscapeSQuotes(Me.AccountNumber) & "' WHERE VendorNumber=" & Me.VendorNumber
            changed = False
        End If
    End If
End Sub
