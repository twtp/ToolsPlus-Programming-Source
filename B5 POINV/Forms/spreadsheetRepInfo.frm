VERSION 5.00
Begin VB.Form RepInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Rep Info - New Products Spreadsheet Import"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnAddContact 
      Caption         =   "Add Outlook Contact"
      Height          =   1065
      Left            =   3390
      TabIndex        =   25
      Top             =   3510
      Width           =   765
   End
   Begin VB.TextBox SalesRepPhoneExt2 
      Height          =   315
      Left            =   3420
      TabIndex        =   17
      Top             =   2910
      Width           =   705
   End
   Begin VB.TextBox SalesRepPhoneExt 
      Height          =   315
      Left            =   3420
      TabIndex        =   13
      Top             =   2550
      Width           =   705
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "OK"
      Height          =   405
      Left            =   1440
      TabIndex        =   26
      Top             =   4590
      Width           =   1305
   End
   Begin VB.TextBox SalesRepEmail 
      Height          =   315
      Left            =   1290
      TabIndex        =   24
      Top             =   3990
      Width           =   1575
   End
   Begin VB.TextBox SalesRepFax 
      Height          =   315
      Left            =   1290
      TabIndex        =   22
      Top             =   3630
      Width           =   1545
   End
   Begin VB.TextBox SalesRepCell 
      Height          =   315
      Left            =   1290
      TabIndex        =   19
      Top             =   3270
      Width           =   1575
   End
   Begin VB.TextBox SalesRepPhone2 
      Height          =   315
      Left            =   1290
      TabIndex        =   16
      Top             =   2910
      Width           =   1575
   End
   Begin VB.TextBox SalesRepPhone 
      Height          =   315
      Left            =   1290
      TabIndex        =   12
      Top             =   2550
      Width           =   1575
   End
   Begin VB.TextBox SalesRepZip 
      Height          =   315
      Left            =   2730
      TabIndex        =   9
      Top             =   1740
      Width           =   885
   End
   Begin VB.TextBox SalesRepState 
      Height          =   315
      Left            =   1260
      TabIndex        =   8
      Top             =   1740
      Width           =   555
   End
   Begin VB.TextBox SalesRepCity 
      Height          =   315
      Left            =   1290
      TabIndex        =   5
      Top             =   1380
      Width           =   2295
   End
   Begin VB.TextBox SalesRepAddress 
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Top             =   1020
      Width           =   2325
   End
   Begin VB.TextBox SalesRepName 
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   660
      Width           =   2355
   End
   Begin VB.Label generalLabel 
      Caption         =   "Email:"
      Height          =   315
      Index           =   10
      Left            =   360
      TabIndex        =   23
      Top             =   3990
      Width           =   885
   End
   Begin VB.Label generalLabel 
      Caption         =   "Fax:"
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   21
      Top             =   3660
      Width           =   915
   End
   Begin VB.Label generalLabel 
      Caption         =   "Cell:"
      Height          =   285
      Index           =   8
      Left            =   360
      TabIndex        =   20
      Top             =   3300
      Width           =   945
   End
   Begin VB.Label generalLabel 
      Caption         =   "Ext:"
      Height          =   285
      Index           =   12
      Left            =   3000
      TabIndex        =   18
      Top             =   2940
      Width           =   405
   End
   Begin VB.Label generalLabel 
      Caption         =   "Ext:"
      Height          =   285
      Index           =   11
      Left            =   3030
      TabIndex        =   14
      Top             =   2580
      Width           =   405
   End
   Begin VB.Label generalLabel 
      Caption         =   "Phone #2:"
      Height          =   285
      Index           =   7
      Left            =   330
      TabIndex        =   15
      Top             =   2910
      Width           =   945
   End
   Begin VB.Label generalLabel 
      Caption         =   "Phone #1:"
      Height          =   285
      Index           =   6
      Left            =   330
      TabIndex        =   11
      Top             =   2550
      Width           =   885
   End
   Begin VB.Label generalLabel 
      Caption         =   "Sales Rep Contact Information"
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
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4035
   End
   Begin VB.Label generalLabel 
      Caption         =   "Zip Code:"
      Height          =   255
      Index           =   5
      Left            =   1920
      TabIndex        =   10
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label generalLabel 
      Caption         =   "State:"
      Height          =   255
      Index           =   4
      Left            =   540
      TabIndex        =   7
      Top             =   1770
      Width           =   675
   End
   Begin VB.Label generalLabel 
      Caption         =   "City:"
      Height          =   285
      Index           =   3
      Left            =   540
      TabIndex        =   6
      Top             =   1410
      Width           =   675
   End
   Begin VB.Label generalLabel 
      Caption         =   "Address:"
      Height          =   285
      Index           =   2
      Left            =   540
      TabIndex        =   4
      Top             =   1050
      Width           =   675
   End
   Begin VB.Label generalLabel 
      Caption         =   "Name:"
      Height          =   285
      Index           =   1
      Left            =   540
      TabIndex        =   2
      Top             =   690
      Width           =   675
   End
End
Attribute VB_Name = "RepInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private changed As Boolean

Private Sub btnAddContact_Click()
    Dim contact As outlookContact
    contact.fullName = Me.SalesRepName
    contact.Address = Me.SalesRepAddress
    contact.City = Me.SalesRepCity
    contact.State = Me.SalesRepState
    contact.ZipCode = Me.SalesRepZip
    contact.phoneNumber1 = Me.SalesRepPhone
    contact.phoneExtension1 = Me.SalesRepPhoneExt
    contact.phoneNumber2 = Me.SalesRepPhone2
    contact.phoneExtension2 = Me.SalesRepPhoneExt2
    contact.faxNumber = Me.SalesRepFax
    contact.emailAddress = Me.SalesRepEmail
    AddOutlookContact contact
End Sub

'Private Sub btnAddContact_Click()
'    Dim contact As outlook.ContactItem
'    Dim folder As outlook.MAPIFolder
'    Set folder = outlook.Application.GetNamespace("MAPI").Folders("Public Folders").Folders("All Public Folders").Folders("Margie's Contacts")
'
'    Set contact = folder.Items.Find("[fullName]=" & qq(Me.SalesRepName))
'    If contact Is Nothing Then
'        Set contact = folder.Items.Add(olContactItem)
'        contact.fullName = Me.SalesRepName
'        contact.jobTitle = "Sales Rep."
'        contact.businessAddress = Me.SalesRepAddress
'        contact.businessAddressCity = Me.SalesRepCity
'        contact.businessAddressState = Me.SalesRepState
'        contact.businessAddressPostalCode = Me.SalesRepZip
'        If Me.SalesRepPhoneExt = "" Then
'            contact.BusinessTelephoneNumber = Me.SalesRepPhone
'        Else
'            contact.BusinessTelephoneNumber = Me.SalesRepPhone & " x" & Me.SalesRepPhoneExt
'        End If
'        If Me.SalesRepPhoneExt2 = "" Then
'            contact.Business2TelephoneNumber = Me.SalesRepPhone2
'        Else
'            contact.Business2TelephoneNumber = Me.SalesRepPhone2 & " x" & Me.SalesRepPhoneExt2
'        End If
'        contact.MobileTelephoneNumber = Me.SalesRepCell
'        contact.BusinessFaxNumber = Me.SalesRepFax
'        contact.Email1Address = Me.SalesRepEmail
'        contact.Save
'        MsgBox "Contact added. Check the ""File As""."
'        contact.Display
'    Else
'        MsgBox "Contact already found."
'        contact.Display
'    End If
'
'End Sub

Private Sub btnExit_Click()
    Unload RepInfo
End Sub

Private Sub Form_Load()
    Dim rst As ADODB.Recordset
    Set rst = db.retrieve("SELECT * FROM NewProdXSheetHeader WHERE VendorNumber='" & NewProdDetailForm.VendorNumber & "'")
    Me.SalesRepName = Nz(rst("SalesRepName"))
    Me.SalesRepAddress = Nz(rst("SalesRepAddress"))
    Me.SalesRepCity = Nz(rst("SalesRepCity"))
    Me.SalesRepState = Nz(rst("SalesRepState"))
    Me.SalesRepZip = Nz(rst("SalesRepZip"))
    Me.SalesRepPhone = Nz(rst("SalesRepPhone"))
    Me.SalesRepPhoneExt = Nz(rst("SalesRepPhoneExt"))
    Me.SalesRepPhone2 = Nz(rst("SalesRepPhone2"))
    Me.SalesRepPhoneExt2 = Nz(rst("SalesRepPhone2"))
    Me.SalesRepCell = Nz(rst("SalesRepCell"))
    Me.SalesRepFax = Nz(rst("SalesRepFax"))
    Me.SalesRepEmail = Nz(rst("SalesRepEmail"))
    rst.Close
    Set rst = Nothing
End Sub




'--------------------------------
' db updating
'--------------------------------
Private Sub SalesRepName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SalesRepName_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub SalesRepName_LostFocus()
    If changed = True Then
        updateXSheetHeader "SalesRepName", Me.SalesRepName, NewProdDetailForm.VendorNumber, "'"
        changed = False
    End If
End Sub

Private Sub SalesRepAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SalesRepAddress_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub SalesRepAddress_LostFocus()
    If changed = True Then
        updateXSheetHeader "SalesRepAddress", Me.SalesRepAddress, NewProdDetailForm.VendorNumber, "'"
        changed = False
    End If
End Sub

Private Sub SalesRepCity_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SalesRepCity_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub SalesRepCity_LostFocus()
    If changed = True Then
        updateXSheetHeader "SalesRepCity", Me.SalesRepCity, NewProdDetailForm.VendorNumber, "'"
        changed = False
    End If
End Sub

Private Sub SalesRepState_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SalesRepState_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub SalesRepState_LostFocus()
    If changed = True Then
        updateXSheetHeader "SalesRepState", Me.SalesRepState, NewProdDetailForm.VendorNumber, "'"
        changed = False
    End If
End Sub

Private Sub SalesRepZip_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SalesRepZip_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub SalesRepZip_LostFocus()
    If changed = True Then
        updateXSheetHeader "SalesRepZip", Me.SalesRepZip, NewProdDetailForm.VendorNumber, "'"
        changed = False
    End If
End Sub

Private Sub SalesRepPhone_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SalesRepPhone_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub SalesRepPhone_LostFocus()
    If changed = True Then
        updateXSheetHeader "SalesRepPhone", Me.SalesRepPhone, NewProdDetailForm.VendorNumber, "'"
        changed = False
    End If
End Sub

Private Sub SalesRepPhoneExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SalesRepPhoneExt_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub SalesRepPhoneExt_LostFocus()
    If changed = True Then
        updateXSheetHeader "SalesRepPhoneExt", Me.SalesRepPhoneExt, NewProdDetailForm.VendorNumber, "'"
        changed = False
    End If
End Sub

Private Sub SalesRepPhone2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SalesRepPhone2_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub SalesRepPhone2_LostFocus()
    If changed = True Then
        updateXSheetHeader "SalesRepPhone2", Me.SalesRepPhone2, NewProdDetailForm.VendorNumber, "'"
        changed = False
    End If
End Sub

Private Sub SalesRepPhoneExt2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SalesRepPhoneExt2_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub SalesRepPhoneExt2_LostFocus()
    If changed = True Then
        updateXSheetHeader "SalesRepPhoneExt2", Me.SalesRepPhoneExt2, NewProdDetailForm.VendorNumber, "'"
        changed = False
    End If
End Sub

Private Sub SalesRepCell_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SalesRepCell_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub SalesRepCell_LostFocus()
    If changed = True Then
        updateXSheetHeader "SalesRepCell", Me.SalesRepCell, NewProdDetailForm.VendorNumber, "'"
        changed = False
    End If
End Sub

Private Sub SalesRepFax_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SalesRepFax_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub SalesRepFax_LostFocus()
    If changed = True Then
        updateXSheetHeader "SalesRepFax", Me.SalesRepFax, NewProdDetailForm.VendorNumber, "'"
        changed = False
    End If
End Sub

Private Sub SalesRepEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SalesRepEmail_LostFocus
    Else
        changed = True
    End If
End Sub

Private Sub SalesRepEmail_LostFocus()
    If changed = True Then
        updateXSheetHeader "SalesRepEmail", Me.SalesRepEmail, NewProdDetailForm.VendorNumber, "'"
        changed = False
    End If
End Sub
