VERSION 5.00
Begin VB.Form NewProdDetailForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detail View - New Products Spreadsheet Import"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbWebsite 
      Height          =   315
      Left            =   6450
      Style           =   2  'Dropdown List
      TabIndex        =   77
      Top             =   630
      Width           =   1785
   End
   Begin VB.CommandButton btnCopyInfoToMas 
      Caption         =   "Copy Vendor/Rep Info to MAS 200"
      Height          =   675
      Left            =   4860
      TabIndex        =   76
      Top             =   420
      Width           =   1425
   End
   Begin VB.CommandButton btnCopyInfoToPoInv 
      Caption         =   "Copy Vendor/Rep Info to PO/Inv"
      Height          =   675
      Left            =   3420
      TabIndex        =   75
      Top             =   420
      Width           =   1425
   End
   Begin VB.CommandButton btnDeleteFilter 
      Caption         =   "Delete All In Filter"
      Height          =   375
      Left            =   8790
      TabIndex        =   74
      Top             =   5760
      Width           =   1605
   End
   Begin VB.CommandButton btnDeleteThis 
      Caption         =   "Delete This Item"
      Height          =   345
      Left            =   8790
      TabIndex        =   73
      Top             =   5400
      Width           =   1605
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
      Left            =   3210
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   6600
      Width           =   315
   End
   Begin VB.CommandButton btnNextLC 
      Caption         =   "uu"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2700
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   6600
      Width           =   495
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
      Left            =   0
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   6600
      Width           =   315
   End
   Begin VB.CommandButton btnPrevLC 
      Caption         =   "tt"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   330
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   6600
      Width           =   495
   End
   Begin VB.ComboBox cmbFilters 
      Height          =   315
      Left            =   6450
      TabIndex        =   68
      Top             =   120
      Width           =   1755
   End
   Begin VB.TextBox Spec 
      Height          =   285
      Index           =   7
      Left            =   720
      TabIndex        =   67
      Top             =   5550
      Width           =   3495
   End
   Begin VB.TextBox Spec 
      Height          =   285
      Index           =   6
      Left            =   720
      TabIndex        =   66
      Top             =   5250
      Width           =   3495
   End
   Begin VB.TextBox Spec 
      Height          =   285
      Index           =   5
      Left            =   720
      TabIndex        =   65
      Top             =   4950
      Width           =   3495
   End
   Begin VB.TextBox Spec 
      Height          =   285
      Index           =   4
      Left            =   720
      TabIndex        =   64
      Top             =   4650
      Width           =   3495
   End
   Begin VB.TextBox Spec 
      Height          =   285
      Index           =   3
      Left            =   720
      TabIndex        =   63
      Top             =   4350
      Width           =   3495
   End
   Begin VB.TextBox Spec 
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   62
      Top             =   4050
      Width           =   3495
   End
   Begin VB.TextBox Spec 
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   61
      Top             =   3750
      Width           =   3495
   End
   Begin VB.TextBox Spec 
      Height          =   285
      Index           =   8
      Left            =   720
      TabIndex        =   60
      Top             =   5850
      Width           =   3495
   End
   Begin VB.CommandButton btnAddShipClass 
      Caption         =   "A"
      Height          =   255
      Left            =   6450
      TabIndex        =   58
      Top             =   2940
      Width           =   255
   End
   Begin VB.CommandButton btnHideExported 
      Caption         =   "Hide Exported"
      Height          =   465
      Left            =   7710
      TabIndex        =   56
      Top             =   4260
      Width           =   1485
   End
   Begin VB.CommandButton btnExportFilter 
      Caption         =   "Export All In Filter"
      Height          =   465
      Left            =   7710
      TabIndex        =   55
      Top             =   3780
      Width           =   1485
   End
   Begin VB.CommandButton btnExportThis 
      Caption         =   "Export This Item"
      Height          =   375
      Left            =   7710
      TabIndex        =   54
      Top             =   3390
      Width           =   1485
   End
   Begin VB.TextBox Components 
      Height          =   945
      Left            =   7320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   53
      Top             =   2250
      Width           =   2805
   End
   Begin VB.TextBox DropShipTerms 
      Height          =   285
      Left            =   8550
      TabIndex        =   50
      Top             =   1590
      Width           =   1545
   End
   Begin VB.CheckBox DropShippable 
      Height          =   285
      Left            =   8580
      TabIndex        =   49
      Top             =   1260
      Width           =   285
   End
   Begin VB.ComboBox ShippingClass 
      Height          =   315
      Left            =   5580
      TabIndex        =   47
      Top             =   2910
      Width           =   825
   End
   Begin VB.TextBox ExtendedSpecs 
      Height          =   1515
      Left            =   4410
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   44
      Top             =   4800
      Width           =   2775
   End
   Begin VB.TextBox QtyToOrder 
      Height          =   285
      Left            =   5580
      TabIndex        =   43
      Top             =   4110
      Width           =   825
   End
   Begin VB.TextBox StdPack 
      Height          =   285
      Left            =   5580
      TabIndex        =   41
      Top             =   3810
      Width           =   825
   End
   Begin VB.TextBox Dimensions 
      Height          =   285
      Left            =   5580
      TabIndex        =   39
      Top             =   3510
      Width           =   1395
   End
   Begin VB.TextBox Weight 
      Height          =   285
      Left            =   5580
      TabIndex        =   37
      Top             =   3210
      Width           =   915
   End
   Begin VB.TextBox WebPrice 
      Height          =   285
      Left            =   5610
      TabIndex        =   35
      Top             =   2490
      Width           =   885
   End
   Begin VB.TextBox StorePrice 
      Height          =   285
      Left            =   5610
      TabIndex        =   33
      Top             =   2190
      Width           =   885
   End
   Begin VB.TextBox ListPrice 
      Height          =   285
      Left            =   5610
      TabIndex        =   31
      Top             =   1890
      Width           =   885
   End
   Begin VB.TextBox RegularCost 
      Height          =   285
      Left            =   5610
      TabIndex        =   29
      Top             =   1590
      Width           =   885
   End
   Begin VB.TextBox PromoCost 
      Height          =   285
      Left            =   5610
      TabIndex        =   27
      Top             =   1290
      Width           =   885
   End
   Begin VB.TextBox UPCCode 
      Height          =   315
      Left            =   1110
      TabIndex        =   17
      Top             =   3030
      Width           =   1665
   End
   Begin VB.TextBox EPN 
      Height          =   885
      Left            =   1110
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   2070
      Width           =   2775
   End
   Begin VB.TextBox ItemDescription 
      Height          =   285
      Left            =   1110
      TabIndex        =   13
      Top             =   1710
      Width           =   2295
   End
   Begin VB.TextBox ItemNumber 
      Height          =   285
      Left            =   1110
      TabIndex        =   10
      Top             =   1350
      Width           =   1335
   End
   Begin VB.TextBox txtPosition 
      Height          =   285
      Left            =   3600
      TabIndex        =   8
      Top             =   6570
      Width           =   735
   End
   Begin VB.HScrollBar barPosition 
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Top             =   6600
      Width           =   1845
   End
   Begin VB.CommandButton btnViewRepInfo 
      Caption         =   "View Sales Rep Info"
      Height          =   345
      Left            =   1500
      TabIndex        =   6
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton btnViewVendorInfo 
      Caption         =   "View Vendor Info"
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1485
   End
   Begin VB.ComboBox VendorName 
      Height          =   315
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   810
      Width           =   2115
   End
   Begin VB.TextBox VendorNumber 
      Height          =   285
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1185
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   9780
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblExported 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   2520
      TabIndex        =   59
      Top             =   1320
      Width           =   1785
   End
   Begin VB.Label lblStatusBar 
      Height          =   285
      Left            =   7260
      TabIndex        =   57
      Top             =   6600
      Width           =   3345
   End
   Begin VB.Label generalLabel 
      Caption         =   "Components:"
      Height          =   255
      Index           =   27
      Left            =   7320
      TabIndex        =   52
      Top             =   2010
      Width           =   945
   End
   Begin VB.Label generalLabel 
      Caption         =   "DS Terms:"
      Height          =   255
      Index           =   26
      Left            =   7320
      TabIndex        =   51
      Top             =   1620
      Width           =   1245
   End
   Begin VB.Label generalLabel 
      Caption         =   "Drop Shippable:"
      Height          =   255
      Index           =   25
      Left            =   7320
      TabIndex        =   48
      Top             =   1290
      Width           =   1245
   End
   Begin VB.Label generalLabel 
      Caption         =   "Shipping Class:"
      Height          =   255
      Index           =   19
      Left            =   4440
      TabIndex        =   46
      Top             =   2910
      Width           =   1155
   End
   Begin VB.Label generalLabel 
      Caption         =   "Extended Specs:"
      Height          =   255
      Index           =   24
      Left            =   4410
      TabIndex        =   45
      Top             =   4560
      Width           =   1515
   End
   Begin VB.Label generalLabel 
      Caption         =   "Qty To Order:"
      Height          =   255
      Index           =   23
      Left            =   4440
      TabIndex        =   42
      Top             =   4140
      Width           =   1155
   End
   Begin VB.Label generalLabel 
      Caption         =   "Std Pack:"
      Height          =   255
      Index           =   22
      Left            =   4440
      TabIndex        =   40
      Top             =   3840
      Width           =   1155
   End
   Begin VB.Label generalLabel 
      Caption         =   "Dimensions:"
      Height          =   255
      Index           =   21
      Left            =   4440
      TabIndex        =   38
      Top             =   3540
      Width           =   1155
   End
   Begin VB.Label generalLabel 
      Caption         =   "Weight:"
      Height          =   255
      Index           =   20
      Left            =   4440
      TabIndex        =   36
      Top             =   3240
      Width           =   1155
   End
   Begin VB.Label generalLabel 
      Caption         =   "Web Price:"
      Height          =   255
      Index           =   18
      Left            =   4470
      TabIndex        =   34
      Top             =   2520
      Width           =   1155
   End
   Begin VB.Label generalLabel 
      Caption         =   "Store Price:"
      Height          =   255
      Index           =   17
      Left            =   4470
      TabIndex        =   32
      Top             =   2220
      Width           =   1155
   End
   Begin VB.Label generalLabel 
      Caption         =   "List Price:"
      Height          =   255
      Index           =   16
      Left            =   4470
      TabIndex        =   30
      Top             =   1920
      Width           =   1155
   End
   Begin VB.Label generalLabel 
      Caption         =   "Regular Cost:"
      Height          =   255
      Index           =   15
      Left            =   4470
      TabIndex        =   28
      Top             =   1620
      Width           =   1155
   End
   Begin VB.Label generalLabel 
      Caption         =   "Promo Cost:"
      Height          =   255
      Index           =   14
      Left            =   4470
      TabIndex        =   26
      Top             =   1320
      Width           =   1155
   End
   Begin VB.Label generalLabel 
      Caption         =   "Spec8:"
      Height          =   225
      Index           =   13
      Left            =   60
      TabIndex        =   25
      Top             =   5880
      Width           =   645
   End
   Begin VB.Label generalLabel 
      Caption         =   "Spec7:"
      Height          =   225
      Index           =   12
      Left            =   60
      TabIndex        =   24
      Top             =   5580
      Width           =   645
   End
   Begin VB.Label generalLabel 
      Caption         =   "Spec6:"
      Height          =   225
      Index           =   11
      Left            =   60
      TabIndex        =   23
      Top             =   5280
      Width           =   645
   End
   Begin VB.Label generalLabel 
      Caption         =   "Spec5:"
      Height          =   225
      Index           =   10
      Left            =   60
      TabIndex        =   22
      Top             =   4980
      Width           =   645
   End
   Begin VB.Label generalLabel 
      Caption         =   "Spec4:"
      Height          =   225
      Index           =   9
      Left            =   60
      TabIndex        =   21
      Top             =   4680
      Width           =   645
   End
   Begin VB.Label generalLabel 
      Caption         =   "Spec3:"
      Height          =   225
      Index           =   8
      Left            =   60
      TabIndex        =   20
      Top             =   4380
      Width           =   645
   End
   Begin VB.Label generalLabel 
      Caption         =   "Spec2:"
      Height          =   225
      Index           =   7
      Left            =   60
      TabIndex        =   19
      Top             =   4080
      Width           =   645
   End
   Begin VB.Label generalLabel 
      Caption         =   "Spec1:"
      Height          =   225
      Index           =   6
      Left            =   60
      TabIndex        =   18
      Top             =   3780
      Width           =   645
   End
   Begin VB.Label generalLabel 
      Caption         =   "UPC:"
      Height          =   255
      Index           =   5
      Left            =   60
      TabIndex        =   16
      Top             =   3060
      Width           =   735
   End
   Begin VB.Label generalLabel 
      Caption         =   "EPN:"
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   14
      Top             =   2100
      Width           =   1065
   End
   Begin VB.Label generalLabel 
      Caption         =   "Description:"
      Height          =   285
      Index           =   3
      Left            =   60
      TabIndex        =   12
      Top             =   1740
      Width           =   1065
   End
   Begin VB.Label generalLabel 
      Caption         =   "Item Number:"
      Height          =   285
      Index           =   2
      Left            =   60
      TabIndex        =   11
      Top             =   1350
      Width           =   1035
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   10620
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label lblMaxAmt 
      Caption         =   "of MAX"
      Height          =   225
      Left            =   4380
      TabIndex        =   9
      Top             =   6630
      Width           =   1065
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10620
      Y1              =   6510
      Y2              =   6510
   End
   Begin VB.Label generalLabel 
      Caption         =   "Vendor Name:"
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label generalLabel 
      Caption         =   "Vendor:"
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   1065
   End
End
Attribute VB_Name = "NewProdDetailForm"
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
    If changed = True Then
        Select Case whichCtl
            Case Is = "ItemNumber"
                ItemNumber_LostFocus
            Case Is = "ItemDescription"
                ItemDescription_LostFocus
            Case Is = "EPN"
                EPN_LostFocus
            Case Is = "UPCCode"
                UPCCode_LostFocus
            Case Is = "Spec1"
                Spec_LostFocus 1
            Case Is = "Spec2"
                Spec_LostFocus 2
            Case Is = "Spec3"
                Spec_LostFocus 3
            Case Is = "Spec4"
                Spec_LostFocus 4
            Case Is = "Spec5"
                Spec_LostFocus 5
            Case Is = "Spec6"
                Spec_LostFocus 6
            Case Is = "Spec7"
                Spec_LostFocus 7
            Case Is = "Spec8"
                Spec_LostFocus 8
            Case Is = "PromoCost"
                PromoCost_LostFocus
            Case Is = "RegularCost"
                RegularCost_LostFocus
            Case Is = "ListPrice"
                ListPrice_LostFocus
            Case Is = "StorePrice"
                StorePrice_LostFocus
            Case Is = "WebPrice"
                WebPrice_LostFocus
            Case Is = "Weight"
                Weight_LostFocus
            Case Is = "Dimensions"
                Dimensions_LostFocus
            Case Is = "StdPack"
                StdPack_LostFocus
            Case Is = "QtyToOrder"
                QtyToOrder_LostFocus
            Case Is = "ExtendedSpecs"
                ExtendedSpecs_LostFocus
            Case Is = "DropShipTerms"
                DropShipTerms_LostFocus
            Case Is = "Components"
                Components_LostFocus
            Case Else
                MsgBox "Couldn't determine which field you were updating, so not saving..."
        End Select
    End If
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("exec spNewProdHeaderAndDetailsByItem '" & ItemList(Me.barPosition.value) & "'")
    If rst.EOF Then
        clearForm
    Else
        fillForm rst
    End If
    rst.Close
    Set rst = Nothing
    Me.txtPosition = Me.barPosition.value + 1
    changed = False
End Sub

Private Sub btnAddShipClass_Click()
    MsgBox "goes nowhere, does nothing"
    'Load AddShipClass
    'AddShipClass.Show MODAL
    'requeryShipClass
End Sub

Private Sub btnCopyInfoToMas_Click()
    Dim sql As String
    sql = "SELECT VendorNumber, VendorName, VendorAddress, VendorCity, VendorState, VendorZip, VendorPhone, VendorFax, VendorEmail, VendorURL FROM NewProdXSheetHeader WHERE VendorNumber='" & Me.VendorNumber & "'"
    ExportAsCSV sql, "s:\mastest\mas90-signs\export\vendor_update.csv"
    ShellWait "s:\mastest\mas90-signs\export\mas200_import_vendor_update.bat", vbHide
End Sub

Private Sub btnCopyInfoToPoInv_Click()
    'Dim rst As ADODB.Recordset
    'Set rst = DB.retrieve("SELECT VendorName, VendorAddress, VendorCity, VendorState, VendorZip, VendorPhone, VendorFax, VendorEmail, VendorURL FROM NewProdXSheetHeader WHERE VendorNumber='" & Me.VendorNumber & "'")
    'DB.Execute "UPDATE ProductLine SET CompanyName='" & rst("VendorName") & "', CompanyAddress='" & rst("VendorAddress") & "', CompanyCity='" & rst("VendorCity") & "', CompanyState='" & rst("VendorState") & "', CompanyZip='" & rst("VendorZip") & "', CompanyPhone1='" & rst("VendorPhone") & "', CompanyFax='" & rst("VendorFax") & "', CompanyEmail='" & rst("VendorEmail") & "', CompanyURL='" & rst("VendorURL") & "' WHERE ProductLine='" & Left(Me.ItemNumber, 3) & "'"
    'rst.Close
    'Set rst = DB.retrieve("SELECT SalesRepName, SalesRepAddress, SalesRepCity, SalesRepState, SalesRepZip, SalesRepPhone + CASE WHEN SalesRepPhoneExt IS NULL OR SalesRepPhoneExt='' THEN '' ELSE ' x' + SalesRepPhoneExt END AS SalesRepPhone, SalesRepPhone2 + CASE WHEN SalesRepPhoneExt2 IS NULL OR SalesRepPhoneExt2='' THEN '' ELSE ' x' + SalesRepPhoneExt2 END AS SalesRepPhone2, SalesRepCell, SalesRepFax, SalesRepEmail FROM NewProdXSheetHeader WHERE VendorNumber='" & Me.VendorNumber & "'")
    'DB.Execute "UPDATE ProductLine SET SalesRepName='" & rst("SalesRepName") & "', SalesRepAddress='" & rst("SalesRepAddress") & "', SalesRepCity='" & rst("SalesRepCity") & "', SalesRepState='" & rst("SalesRepState") & "', SalesRepZip='" & rst("SalesRepZip") & "', SalesRepPhone1='" & rst("SalesRepPhone") & "', SalesRepPhone2='" & rst("SalesRepPhone2") & "', SalesRepCell='" & rst("SalesRepCell") & "', SalesRepFax='" & rst("SalesRepFax") & "', SalesRepEmail='" & rst("SalesRepEmail") & "' WHERE ProductLine='" & Left(Me.ItemNumber, 3) & "'"
    'rst.Close
    'Set rst = Nothing
    MsgBox "Don't use this anymore! This should be updated to use the new rep info thingy."
End Sub

Private Sub btnDeleteFilter_Click()
    If MsgBox("Delete all items in the current filter?", vbYesNo) = vbYes Then
        Dim item As Variant
        For Each item In ItemList
            DB.Execute "DELETE FROM NewProdXSheetDetail WHERE ItemNumber='" & CStr(item) & "'"
        Next item
    End If
End Sub

Private Sub btnDeleteThis_Click()
    If MsgBox("Delete this item?", vbYesNo) = vbYes Then
        DB.Execute "DELETE FROM NewProdXSheetDetail WHERE ItemNumber='" & Me.ItemNumber & "'"
    End If
End Sub

Private Sub btnExit_Click()
    Unload NewProdDetailForm
End Sub

Private Sub btnExportFilter_Click()
    Mouse.Hourglass True
    Dim websiteID As Long
    websiteID = getSelectedWebsiteID()
    Dim item As Variant, i As Long, txt As String
    i = 1
    txt = ""
    For Each item In ItemList
        Me.lblStatusBar.Caption = "Offloading " & CStr(item) & " (" & i & " of " & UBound(ItemList) + 1 & ")"
        DoEvents
        If Not offload_single_item(CStr(item), websiteID) Then
            txt = txt & CStr(item) & vbCrLf
        Else
            DB.Execute "UPDATE NewProdXSheetDetail SET Exported=1 WHERE ItemNumber='" & CStr(item) & "'"
        End If
        i = i + 1
    Next item
    Me.lblStatusBar.Caption = ""
    Mouse.Hourglass False
    If txt <> "" Then
        MsgBox "Errors occurred. The following did not export properly." & vbCrLf & vbCrLf & txt
    End If
End Sub

Private Sub btnExportThis_Click()
    Mouse.Hourglass True
    Dim websiteID As Long
    websiteID = getSelectedWebsiteID()
    Dim item As String
    item = Me.ItemNumber
    If Not offload_single_item(item, websiteID) Then
        MsgBox "Error offloading " & item
    Else
        DB.Execute ("UPDATE NewProdXSheetDetail SET Exported=1 WHERE ItemNumber='" & item & "'")
        handleExportedText 1
    End If
    Mouse.Hourglass False
End Sub

Private Sub btnFirstItem_Click()
    Me.barPosition.value = Me.barPosition.Min
End Sub

Private Sub btnHideExported_Click()
    If Me.btnHideExported.Caption = "Hide Exported" Then
        requeryForm "NotExported"
        Me.btnHideExported.Caption = "Unhide Exported"
    ElseIf Me.btnHideExported.Caption = "Unhide Exported" Then
        requeryForm "UnhideExported"
        Me.btnHideExported.Caption = "Hide Exported"
    Else
        requeryForm "UnhideExported"
        Me.btnHideExported.Caption = "Unhide Exported"
    End If
    Me.barPosition.value = 0
    barPosition_Change
End Sub

Private Sub btnLastItem_Click()
    Me.barPosition.value = Me.barPosition.max
End Sub

Private Sub btnNextLC_Click()
    Dim oldLC As String
    Dim pos As Long
    Dim found As Boolean
    oldLC = Left(Me.ItemNumber, 3)
    pos = Me.barPosition.value
    found = False
    While Not found
        If pos < Me.barPosition.max Then
            If oldLC = Left$(ItemList(pos), 3) Then
                pos = pos + 1
            Else
                found = True
            End If
        Else
            found = True
        End If
    Wend
    Me.barPosition.value = pos
End Sub

Private Sub btnPrevLC_Click()
    Mouse.Hourglass True
    Dim oldLC As String
    Dim pos As Long
    Dim found As Boolean
    oldLC = Left(Me.ItemNumber, 3)
    pos = Me.barPosition.value
    found = False
    While Not found
        If pos > Me.barPosition.Min Then
            If oldLC = Left$(ItemList(pos), 3) Then
                pos = pos - 1
            Else
                found = True
            End If
        Else
            found = True
        End If
    Wend
    Dim newLC As String
    newLC = Left$(ItemList(pos), 3)
    found = False
    While Not found
        If pos > Me.barPosition.Min Then
            If newLC = Left$(ItemList(pos), 3) Then
                pos = pos - 1
            Else
                pos = pos + 1
                found = True
            End If
        Else
            found = True
        End If
    Wend
    Me.barPosition.value = pos
    Mouse.Hourglass False
End Sub

Private Sub btnViewRepInfo_Click()
    Load RepInfo
    RepInfo.Show MODAL
End Sub

Private Sub btnViewVendorInfo_Click()
    Load VendorInfo
    VendorInfo.Show MODAL
End Sub

Private Sub cmbFilters_Click()
    cmbFilters_LostFocus
End Sub

Private Sub cmbFilters_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.cmbFilters, KeyCode, Shift
End Sub

Private Sub cmbFilters_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.cmbFilters, KeyAscii
    If KeyAscii = vbKeyReturn Then
        cmbFilters_LostFocus
    End If
End Sub

Private Sub cmbFilters_LostFocus()
    AutoCompleteLostFocus Me.cmbFilters
    If Me.cmbFilters = "" Then
        Me.cmbFilters = "All"
    End If
    requeryForm Me.cmbFilters
    Me.barPosition.value = 0
    barPosition_Change
    Me.btnHideExported.Caption = "Hide Exported"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageUp Then
        If Me.barPosition <> Me.barPosition.Min Then
            Me.barPosition = Me.barPosition - 1
        End If
    ElseIf KeyCode = vbKeyPageDown Then
        If Me.barPosition <> Me.barPosition.max Then
            Me.barPosition = Me.barPosition + 1
        End If
    End If
End Sub

Private Sub Form_Load()
    requeryVendors
    requeryShipClass
    requeryWebsites
    requeryForm "All"
    Me.barPosition.value = 0
    barPosition_Change
    Me.cmbFilters.AddItem "All"
    Me.cmbFilters.AddItem "Only This Vendor"
    Me.cmbFilters = "All"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'NewProdMenuForm.Show
End Sub




'-----------------------------------------------------
' form related functions, but not handlers
'-----------------------------------------------------
Private Sub requeryVendors()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT DISTINCT AP.VendorNo, AP.VendorName FROM ProductLine AS PL, AP_Vendor AS AP WHERE PL.PrimaryVendorNumber=AP.VendorNo ORDER BY AP.VendorName")
    Me.VendorName.Clear
    While Not rst.EOF
        Me.VendorName.AddItem (rst("VendorName"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryShipClass()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT * FROM ShipClass ORDER BY ShipClassName")
    Me.ShippingClass.Clear
    While Not rst.EOF
        Me.ShippingClass.AddItem (rst("ShipClassName"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryWebsites()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT WebsiteID, Description FROM Websites ORDER BY WebsiteID")
    Me.cmbWebsite.Clear
    While Not rst.EOF
        Me.cmbWebsite.AddItem rst("Description")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    Me.cmbWebsite = Me.cmbWebsite.list(0)
End Sub

Private Function getSelectedWebsiteID() As Long
    getSelectedWebsiteID = CLng(DLookup("WebsiteID", "Websites", "Description='" & EscapeSQuotes(Me.cmbWebsite) & "'"))
End Function

Private Sub requeryForm(filter As String)
    Dim rst As ADODB.Recordset
    Select Case filter
        Case Is = "All"
            Set rst = DB.retrieve("SELECT h.*, d.* FROM NewProdXSheetHeader AS h, NewProdXSheetDetail AS d WHERE h.VendorNumber=d.VendorNumber")
        Case Is = "Only This Vendor"
            Set rst = DB.retrieve("SELECT h.*, d.* FROM NewProdXSheetHeader AS h, NewProdXSheetDetail AS d WHERE h.VendorNumber=d.VendorNumber AND h.VendorNumber='" & Me.VendorNumber & "'")
        Case Is = "NotExported"
            If Me.cmbFilters = "All" Then
                Set rst = DB.retrieve("SELECT h.*, d.* FROM NewProdXSheetHeader AS h, NewProdXSheetDetail AS d WHERE h.VendorNumber=d.VendorNumber AND d.Exported=0")
            ElseIf Me.cmbFilters = "Only This Vendor" Then
                Set rst = DB.retrieve("SELECT h.*, d.* FROM NewProdXSheetHeader AS h, NewProdXSheetDetail AS d WHERE h.VendorNumber=d.VendorNumber AND d.Exported=0 AND h.VendorNumber='" & Me.VendorNumber & "'")
            End If
        Case Is = "UnhideExported"
            If Me.cmbFilters = "All" Then
                Set rst = DB.retrieve("SELECT h.*, d.* FROM NewProdXSheetHeader AS h, NewProdXSheetDetail AS d WHERE h.VendorNumber=d.VendorNumber")
            ElseIf Me.cmbFilters = "Only This Vendor" Then
                Set rst = DB.retrieve("SELECT h.*, d.* FROM NewProdXSheetHeader AS h, NewProdXSheetDetail AS d WHERE h.VendorNumber=d.VendorNumber AND h.VendorNumber='" & Me.VendorNumber & "'")
            End If
        Case Else
            MsgBox "Unknown filter type, default to ""All""."
            Set rst = DB.retrieve("SELECT h.*, d.* FROM NewProdXSheetHeader AS h, NewProdXSheetDetail AS d WHERE h.VendorNumber=d.VendorNumber")
    End Select
    If rst.RecordCount > 0 Then
        ReDim ItemList(rst.RecordCount - 1) As String
        Dim i As Long
        For i = 0 To rst.RecordCount - 1
            ItemList(i) = rst("ItemNumber")
            rst.MoveNext
        Next i
        Me.lblMaxAmt.Caption = "of " & rst.RecordCount
        Me.barPosition.Min = 0
        Me.barPosition.max = rst.RecordCount - 1
    Else
        ReDim ItemList(0) As String
        ItemList(0) = "NORECORDS"
        Me.lblMaxAmt.Caption = "of 0"
        Me.barPosition.Min = 0
        Me.barPosition.max = 0
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub fillForm(item As ADODB.Recordset)
    Me.VendorNumber = item("VendorNumber")
    Me.VendorName = item("VendorName")
    Me.ItemNumber = item("ItemNumber")
    Me.ItemDescription = item("ItemDescription")
    Me.EPN = Nz(item("EPN"))
    Me.UPCCode = Nz(item("UPC"))
    Dim i As Long
    For i = 1 To 8
        Me.Spec(i) = Nz(item("Spec" & i))
    Next i
    Me.PromoCost = Format(item("PromoCost"), "Currency")
    Me.RegularCost = Format(item("RegularCost"), "Currency")
    Me.ListPrice = Format(item("ListPrice"), "Currency")
    Me.StorePrice = Format(item("StorePrice"), "Currency")
    Me.WebPrice = Format(item("WebPrice"), "Currency")
    Me.ShippingClass = item("ShippingClass")
    Me.Weight = item("Weight")
    Me.Dimensions = item("Dimensions")
    Me.StdPack = item("StdPack")
    Me.QtyToOrder = item("QtyToOrder")
    Me.ExtendedSpecs = Nz(item("ExtendedSpecs"))
    fillingForm = True
    Me.DropShippable = SQLBool(item("DropShippable"))
    fillingForm = False
    Me.DropShipTerms = Nz(item("DropShipTerms"))
    Me.components = Nz(item("Components"))
    handleExportedText item("Exported")
End Sub

Private Sub clearForm()
    Me.VendorNumber = ""
    Me.VendorName = ""
    Me.ItemNumber = ""
    Me.ItemDescription = ""
    Me.EPN = ""
    Me.UPCCode = ""
    Dim i As Long
    For i = 1 To 8
        Me.Spec(i) = ""
    Next i
    Me.PromoCost = ""
    Me.RegularCost = ""
    Me.ListPrice = ""
    Me.StorePrice = ""
    Me.WebPrice = ""
    Me.ShippingClass = ""
    Me.Weight = ""
    Me.Dimensions = ""
    Me.StdPack = ""
    Me.QtyToOrder = ""
    Me.ExtendedSpecs = ""
    Me.DropShippable = 0
    Me.DropShipTerms = ""
    Me.components = ""
    handleExportedText False
End Sub

Private Sub handleExportedText(exported As Boolean)
    Me.lblExported.Caption = IIf(exported, "Exported!", "")
End Sub




'----------------------------------------------------
' error checking and updates to the database
'----------------------------------------------------
Private Sub ItemNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        ItemNumber_LostFocus
    Else
        changed = True
        whichCtl = "ItemNumber"
    End If
End Sub

Private Sub ItemNumber_LostFocus()
    If changed = True Then
        Me.ItemNumber = UCase$(Me.ItemNumber)
        If Not PLExists(Left$(Me.ItemNumber, 3)) Then
            MsgBox "Could not find line code. Create a new line code first."
        ElseIf Not validateItemNumber(Me.ItemNumber) Then
            MsgBox "Invalid length (between 4 and 27) or characters (only alphanumeric and -)."
        Else
            Dim newitem As String
            newitem = Me.ItemNumber
            updateXSheetDetail "ItemNumber", newitem, ItemList(Me.barPosition.value), "'"
            changed = False
            requeryForm "All"
            Me.barPosition.value = bsearch(newitem, ItemList)
        End If
    End If
End Sub

Private Sub ItemDescription_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        ItemDescription_LostFocus
    Else
        changed = True
        whichCtl = "ItemDescription"
    End If
End Sub

Private Sub ItemDescription_LostFocus()
    If changed = True Then
        Me.ItemDescription = Replace(Me.ItemDescription, """", "''")
        Me.ItemDescription = Replace(Me.ItemDescription, ",", " ")
        If validateItemDescription(Me.ItemDescription) Then
            updateXSheetDetail "ItemDescription", Me.ItemDescription, Me.ItemNumber, "'"
        Else
            MsgBox "Item description is invalid, must be less than 255 chars, no "" or ,"
            changed = False
            Me.ItemDescription.SetFocus
        End If
    End If
End Sub

Private Sub EPN_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        EPN_LostFocus
    Else
        changed = True
        whichCtl = "EPN"
    End If
End Sub

Private Sub EPN_LostFocus()
    If changed = True Then
        If validateEPN(Me.EPN) Then
            updateXSheetDetail "EPN", Me.EPN, Me.ItemNumber, "'"
            changed = False
        Else
            MsgBox "Invalid EPN, must be less than 512 chars."
            Me.EPN.SetFocus
        End If
    End If
End Sub

Private Sub UPCCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        UPCCode_LostFocus
    Else
        changed = True
        whichCtl = "UPCCode"
    End If
End Sub

Private Sub UPCCode_LostFocus()
    If changed = True Then
        Me.UPCCode = Replace(Me.UPCCode, "-", "")
        If validateUPC(Me.UPCCode) Then
            Me.UPCCode = Format(Me.UPCCode, "General Number")
            updateXSheetDetail "UPC", Me.UPCCode, Me.ItemNumber, "'"
            changed = False
        Else
            MsgBox "Invalid UPC. Must be numeric, <=12 chars."
            Me.UPCCode.SetFocus
        End If
    End If
End Sub

Private Sub Spec_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Spec_LostFocus Index
    Else
        changed = True
        whichCtl = "Spec" & Index
    End If
End Sub

Private Sub Spec_LostFocus(Index As Integer)
    If changed = True Then
        If validateSpecs(Me.Spec(Index)) Then
            updateXSheetDetail "Spec" & Index, Me.Spec(Index), Me.ItemNumber, "'"
            changed = False
        Else
            MsgBox "Invalid specs, too long."
            Me.Spec(Index).SetFocus
        End If
    End If
End Sub

Private Sub PromoCost_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        PromoCost_LostFocus
    Else
        changed = True
        whichCtl = "PromoCost"
    End If
End Sub

Private Sub PromoCost_LostFocus()
    If changed = True Then
        If validateCurrency(Me.PromoCost) Then
            Me.PromoCost = Format(Me.PromoCost, "Currency")
            updateXSheetDetail "PromoCost", Me.PromoCost, Me.ItemNumber, ""
            changed = False
        Else
            MsgBox "Invalid promo cost, must be a number."
            Me.PromoCost.SetFocus
        End If
    End If
End Sub

Private Sub RegularCost_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        RegularCost_LostFocus
    Else
        changed = True
        whichCtl = "RegularCost"
    End If
End Sub

Private Sub RegularCost_LostFocus()
    If changed = True Then
        If validateCurrency(Me.RegularCost) Then
            Me.RegularCost = Format(Me.RegularCost, "Currency")
            updateXSheetDetail "RegularCost", Me.RegularCost, Me.ItemNumber, ""
            changed = False
        Else
            MsgBox "Invalid regular cost, must be a number."
            Me.RegularCost.SetFocus
        End If
    End If
End Sub

Private Sub ListPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        ListPrice_LostFocus
    Else
        changed = True
        whichCtl = "ListPrice"
    End If
End Sub

Private Sub ListPrice_LostFocus()
    If changed = True Then
        If validateCurrency(Me.ListPrice) Then
            Me.ListPrice = Format(Me.ListPrice, "Currency")
            updateXSheetDetail "ListPrice", Me.ListPrice, Me.ItemNumber, ""
            changed = False
        Else
            MsgBox "Invalid list price, must be a number."
            Me.ListPrice.SetFocus
        End If
    End If
End Sub

Private Sub StorePrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        StorePrice_LostFocus
    Else
        changed = True
        whichCtl = "StorePrice"
    End If
End Sub

Private Sub StorePrice_LostFocus()
    If changed = True Then
        If validateCurrency(Me.StorePrice) Then
            Me.StorePrice = Format(Me.StorePrice, "Currency")
            updateXSheetDetail "StorePrice", Me.StorePrice, Me.ItemNumber, ""
            changed = False
        Else
            MsgBox "Invalid store price, must be a number."
            Me.StorePrice.SetFocus
        End If
    End If
End Sub

Private Sub WebPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        WebPrice_LostFocus
    Else
        changed = True
        whichCtl = "WebPrice"
    End If
End Sub

Private Sub WebPrice_LostFocus()
    If changed = True Then
        If validateCurrency(Me.WebPrice) Then
            Me.WebPrice = Format(Me.WebPrice, "Currency")
            updateXSheetDetail "WebPrice", Me.WebPrice, Me.ItemNumber, ""
            changed = False
        Else
            MsgBox "Invalid web price, must be a number."
            Me.WebPrice.SetFocus
        End If
    End If
End Sub

Private Sub ShippingClass_Click()
    ShippingClass_LostFocus
End Sub

Private Sub ShippingClass_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.ShippingClass, KeyCode, Shift
End Sub

Private Sub ShippingClass_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.ShippingClass, KeyAscii
    If KeyAscii = vbKeyReturn Then
        ShippingClass_LostFocus
    End If
End Sub

Private Sub ShippingClass_LostFocus()
    AutoCompleteLostFocus Me.ShippingClass
    updateXSheetDetail "ShippingClass", Me.ShippingClass, Me.ItemNumber, "'"
End Sub

Private Sub Weight_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Weight_LostFocus
    Else
        changed = True
        whichCtl = "Weight"
    End If
End Sub

Private Sub Weight_LostFocus()
    If changed = True Then
        Me.Weight = validateWeight(Me.Weight)
        If Me.Weight <> "" Then
            updateXSheetDetail "Weight", Me.Weight, Me.ItemNumber, "'"
            changed = False
        Else
            MsgBox "Unable to parse weight, try again?"
            Me.Weight.SetFocus
        End If
    End If
End Sub

Private Sub Dimensions_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dimensions_LostFocus
    Else
        changed = True
        whichCtl = "Dimensions"
    End If
End Sub

Private Sub Dimensions_LostFocus()
    If changed = True Then
        Dim dims As String
        dims = validateDimensions(Me.Dimensions)
        If Left(dims, 5) <> "ERROR" Then
            Me.Dimensions = dims
            updateXSheetDetail "Dimensions", Me.Dimensions, Me.ItemNumber, "'"
            changed = False
        ElseIf dims = "ERROR:numeric" Then
            MsgBox "Invalid dimensions, dimensions must be numeric."
            Me.Dimensions.SetFocus
        ElseIf dims = "ERROR:dimensions" Then
            MsgBox "Invalid dimensions, must have 3 dimensions formatted like: 1x2x3"
            Me.Dimensions.SetFocus
        End If
    End If
End Sub

Private Sub StdPack_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        StdPack_LostFocus
    Else
        changed = True
        whichCtl = "StdPack"
    End If
End Sub

Private Sub StdPack_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.StdPack) Then
            Me.StdPack = Format(Me.StdPack, "General Number")
            updateXSheetDetail "StdPack", Me.StdPack, Me.ItemNumber, ""
            changed = False
        Else
            MsgBox "Invalid std pack, must be an integer."
            Me.StdPack.SetFocus
        End If
    End If
End Sub

Private Sub QtyToOrder_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        QtyToOrder_LostFocus
    Else
        changed = True
        whichCtl = "QtyToOrder"
    End If
End Sub

Private Sub QtyToOrder_LostFocus()
    If changed = True Then
        If validateGeneralInteger(Me.QtyToOrder) Then
            Me.QtyToOrder = Format(Me.QtyToOrder, "General Number")
            updateXSheetDetail "QtyToOrder", Me.QtyToOrder, Me.ItemNumber, ""
            changed = False
        Else
            MsgBox "Invalid qty to order, must be an integer."
            Me.QtyToOrder.SetFocus
        End If
    End If
End Sub

Private Sub ExtendedSpecs_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        ExtendedSpecs_LostFocus
    Else
        changed = True
        whichCtl = "ExtendedSpecs"
    End If
End Sub

Private Sub ExtendedSpecs_LostFocus()
    If changed = True Then
        updateXSheetDetail "ExtendedSpecs", Me.ExtendedSpecs, Me.ItemNumber, "'"
        changed = False
    End If
End Sub

Private Sub DropShippable_Click()
    If Not fillingForm Then
        updateXSheetDetail "DropShippable", SQLBool(Me.DropShippable), Me.ItemNumber, "'"
    End If
End Sub

Private Sub DropShipTerms_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        DropShipTerms_LostFocus
    Else
        changed = True
        whichCtl = "DropShipTerms"
    End If
End Sub

Private Sub DropShipTerms_LostFocus()
    If changed = True Then
        If validateDropShipTerms(Me.DropShipTerms) Then
            updateXSheetDetail "DropShipTerms", Me.DropShipTerms, Me.ItemNumber, "'"
            changed = False
        Else
            MsgBox "Invalid drop ship terms, too long."
            Me.DropShipTerms.SetFocus
        End If
    End If
End Sub

Private Sub Components_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Components_LostFocus
    Else
        changed = True
        whichCtl = "Components"
    End If
End Sub

Private Sub Components_LostFocus()
    If changed = True Then
        If validateComponents(Me.components) Then
            updateXSheetDetail "Components", Me.components, Me.ItemNumber, "'"
            changed = False
        Else
            MsgBox "Invalid components, too long."
            Me.components.SetFocus
        End If
    End If
End Sub
