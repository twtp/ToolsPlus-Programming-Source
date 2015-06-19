VERSION 5.00
Begin VB.Form AddVendor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Vendor"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4080
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2370
      TabIndex        =   25
      Top             =   4110
      Width           =   1275
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   570
      TabIndex        =   24
      Top             =   4110
      Width           =   1275
   End
   Begin VB.ComboBox Terms 
      Height          =   315
      Left            =   690
      TabIndex        =   23
      Top             =   3450
      Width           =   3285
   End
   Begin VB.TextBox VendorNumber 
      Height          =   285
      Left            =   1380
      TabIndex        =   19
      Top             =   630
      Width           =   1155
   End
   Begin VB.TextBox Website 
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Email 
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   2730
      Width           =   2295
   End
   Begin VB.TextBox Fax 
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   2460
      Width           =   2295
   End
   Begin VB.TextBox Phone 
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   2190
      Width           =   2295
   End
   Begin VB.CommandButton btnCopy 
      Caption         =   "Copy From LC"
      Height          =   705
      Left            =   3300
      TabIndex        =   10
      Top             =   1080
      Width           =   585
   End
   Begin VB.TextBox ZipCode 
      Height          =   255
      Left            =   2310
      TabIndex        =   4
      Top             =   1860
      Width           =   825
   End
   Begin VB.TextBox State 
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1860
      Width           =   555
   End
   Begin VB.TextBox City 
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1590
      Width           =   2295
   End
   Begin VB.TextBox Address 
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox VendorName 
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   1050
      Width           =   2295
   End
   Begin VB.Label generalLabel 
      Caption         =   "Terms:"
      Height          =   225
      Index           =   11
      Left            =   60
      TabIndex        =   22
      Top             =   3510
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Add Vendor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   10
      Left            =   30
      TabIndex        =   21
      Top             =   30
      Width           =   3285
   End
   Begin VB.Label generalLabel 
      Caption         =   "Vendor Number:"
      Height          =   225
      Index           =   9
      Left            =   210
      TabIndex        =   20
      Top             =   660
      Width           =   1155
   End
   Begin VB.Label generalLabel 
      Caption         =   "Website:"
      Height          =   225
      Index           =   8
      Left            =   210
      TabIndex        =   14
      Top             =   3030
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Email:"
      Height          =   225
      Index           =   7
      Left            =   210
      TabIndex        =   13
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Fax:"
      Height          =   225
      Index           =   6
      Left            =   210
      TabIndex        =   12
      Top             =   2490
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Phone:"
      Height          =   225
      Index           =   5
      Left            =   210
      TabIndex        =   11
      Top             =   2220
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "ZipCode:"
      Height          =   225
      Index           =   4
      Left            =   1620
      TabIndex        =   9
      Top             =   1890
      Width           =   705
   End
   Begin VB.Label generalLabel 
      Caption         =   "State:"
      Height          =   225
      Index           =   3
      Left            =   210
      TabIndex        =   8
      Top             =   1890
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "City:"
      Height          =   225
      Index           =   2
      Left            =   210
      TabIndex        =   7
      Top             =   1620
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Address:"
      Height          =   225
      Index           =   1
      Left            =   210
      TabIndex        =   6
      Top             =   1350
      Width           =   615
   End
   Begin VB.Label generalLabel 
      Caption         =   "Name:"
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
End
Attribute VB_Name = "AddVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload AddVendor
End Sub

Private Sub btnCopy_Click()
    'Me.VendorName = AddLineCode.CompanyName
    Me.VendorName = AddLineCode.ManufFullName
    Me.Address = AddLineCode.CompanyAddress1
    Me.City = AddLineCode.CompanyCity
    Me.State = AddLineCode.CompanyState
    Me.ZipCode = AddLineCode.CompanyZipCode
    Me.Phone = AddLineCode.CompanyPhone
    Me.Fax = AddLineCode.CompanyFax
    Me.Email = AddLineCode.CompanyEmail
    Me.Website = AddLineCode.CompanyURL
End Sub

Private Sub btnOK_Click()
    If Me.VendorNumber = "" Then
        MsgBox "You need a vendor number to continue."
        Exit Sub
    End If
    Dim termscode As String
    termscode = DLookup("TermsCode", "AP_TermsCode", "TermsCodeDesc='" & EscapeSQuotes(Me.Terms) & "'")
    If termscode = "" Then
        termscode = "00"
    End If
    Dim line As String
    line = qq(EscapeDQuotes(Me.VendorNumber)) & "," & _
           qq(EscapeDQuotes(Me.VendorName)) & "," & _
           qq(EscapeDQuotes(Me.Address)) & "," & _
           qq(EscapeDQuotes(Me.City)) & "," & _
           qq(EscapeDQuotes(Me.State)) & "," & _
           qq(EscapeDQuotes(Me.ZipCode)) & "," & _
           qq(EscapeDQuotes(Me.Phone)) & "," & _
           qq(EscapeDQuotes(Me.Fax)) & "," & _
           qq(EscapeDQuotes(Me.Email)) & "," & _
           qq(EscapeDQuotes(Me.Website)) & "," & _
           termscode
    Open "s:\mastest\mas90-signs\export\vendoradd.csv" For Output As #1
    Print #1, line
    Close #1
    ShellWait "s:\mastest\mas90-signs\mas200_import_vendor.bat"
    MsgBox "Click OK when Mas200 is complete."
    Unload AddVendor
End Sub

Private Sub Form_Load()
    requeryTerms
    Me.Terms.selStart = 0
    Me.Terms.selLength = 0
    'Me.VendorNumber.SetFocus
End Sub

Private Sub requeryTerms()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT TermsCodeDesc FROM AP_TermsCode ORDER BY TermsCode")
    Me.Terms.Clear
    While Not rst.EOF
        Me.Terms.AddItem rst("TermsCodeDesc")
        rst.MoveNext
    Wend
    rst.MoveFirst
    Me.Terms = rst("TermsCodeDesc")
    rst.Close
    Set rst = Nothing
    ExpandDropDownToFit Me.Terms
End Sub

Private Sub Terms_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.Terms, KeyCode, Shift
End Sub

Private Sub Terms_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.Terms, KeyAscii
End Sub

Private Sub Terms_LostFocus()
    AutoCompleteLostFocus Me.Terms
    If Me.Terms = "" Then
        Me.Terms = Me.Terms.list(0)
    End If
End Sub

Private Sub VendorNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Then
        'ok
    ElseIf KeyCodeIsAlpha(KeyAscii) Or KeyCodeIsNumeric(KeyAscii) Or KeyAscii = Asc("_") Then
        If Len(Me.VendorNumber) < 7 Then
            KeyAscii = ForceUppercase(KeyAscii)
        Else
            KeyAscii = 0
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub VendorNumber_LostFocus()
    If Me.VendorNumber <> "" Then
        If DLookup("VendorNo", "AP_Vendor", "VendorNo='" & Me.VendorNumber & "'") <> "" Then
            MsgBox "Vendor already exists!"
            Me.VendorNumber = ""
        End If
    End If
End Sub
