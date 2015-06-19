VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#2.1#0"; "TPControls.ocx"
Begin VB.Form PrintSigns 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Signs"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   7935
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "TRY THE NEW SIGNS!"
      Height          =   1245
      Left            =   5970
      TabIndex        =   34
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CheckBox hideDCItems 
      Caption         =   "don't show old discontinued items"
      Height          =   405
      Left            =   5640
      TabIndex        =   22
      Top             =   6780
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.ComboBox ItemNumber 
      Height          =   315
      Left            =   1770
      TabIndex        =   19
      Top             =   6810
      Width           =   3375
   End
   Begin VB.TextBox LineCode 
      Height          =   285
      Left            =   1770
      TabIndex        =   18
      Top             =   6480
      Width           =   1395
   End
   Begin VB.TextBox DateTo 
      Height          =   285
      Left            =   1770
      TabIndex        =   17
      Top             =   6150
      Width           =   1395
   End
   Begin VB.TextBox DateFrom 
      Height          =   285
      Left            =   1770
      TabIndex        =   16
      Top             =   5820
      Width           =   1395
   End
   Begin VB.ComboBox SignName 
      Height          =   315
      Left            =   1770
      TabIndex        =   15
      Top             =   5460
      Width           =   3375
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Print or Preview:"
      Height          =   1215
      Index           =   1
      Left            =   5640
      TabIndex        =   25
      Top             =   2070
      Width           =   1845
      Begin VB.OptionButton opPrint 
         Caption         =   "Preview"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   21
         Top             =   720
         Width           =   1515
      End
      Begin VB.OptionButton opPrint 
         Caption         =   "Print"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   20
         Top             =   420
         Width           =   1515
      End
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   6180
      TabIndex        =   24
      Top             =   1230
      Width           =   1395
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   405
      Left            =   6180
      TabIndex        =   23
      Top             =   750
      Width           =   1395
   End
   Begin VB.Frame generalFrame 
      Caption         =   "Select Signs To Print:"
      Height          =   4695
      Index           =   0
      Left            =   390
      TabIndex        =   0
      Top             =   630
      Width           =   4515
      Begin TPControls.DateDropdown DateDropdown1 
         Height          =   315
         Left            =   2310
         TabIndex        =   33
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
      End
      Begin VB.OptionButton opSign 
         Caption         =   "Print Price Changes Since"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   420
         Width           =   2325
      End
      Begin VB.ComboBox TriadCode 
         Height          =   315
         Left            =   2310
         TabIndex        =   11
         Top             =   3540
         Width           =   855
      End
      Begin VB.CommandButton btnResetCodes 
         Caption         =   "Reset"
         Height          =   255
         Left            =   3240
         TabIndex        =   12
         Top             =   3570
         Width           =   735
      End
      Begin VB.OptionButton opSign 
         Caption         =   "Print Signs by Triad Code:"
         Height          =   255
         Index           =   8
         Left            =   150
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3570
         Width           =   4095
      End
      Begin VB.OptionButton opSign 
         Caption         =   "Print Misc. Sign"
         Height          =   255
         Index           =   10
         Left            =   150
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   4320
         Width           =   4095
      End
      Begin VB.OptionButton opSign 
         Caption         =   "Print Group Sign"
         Height          =   255
         Index           =   9
         Left            =   150
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4020
         Width           =   4095
      End
      Begin VB.CommandButton btnResetFlags 
         Caption         =   "Reset"
         Height          =   255
         Left            =   2130
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton opSign 
         Caption         =   "Print Specified Sign (by line code)"
         Height          =   255
         Index           =   7
         Left            =   150
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3120
         Width           =   4095
      End
      Begin VB.OptionButton opSign 
         Caption         =   "Print Specified Sign (any part number)"
         Height          =   255
         Index           =   6
         Left            =   150
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2820
         Width           =   4095
      End
      Begin VB.OptionButton opSign 
         Caption         =   "Print Single Sign (by part number)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   150
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2370
         Width           =   4095
      End
      Begin VB.OptionButton opSign 
         Caption         =   "Print Signs For Specified Line Code (group by sign)"
         Height          =   255
         Index           =   4
         Left            =   150
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1920
         Width           =   4095
      End
      Begin VB.OptionButton opSign 
         Caption         =   "Print Signs By Date Last Updated (group by sign)"
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1470
         Width           =   4095
      End
      Begin VB.OptionButton opSign 
         Caption         =   "Print All Price Changes By Line Code"
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1020
         Width           =   4095
      End
      Begin VB.OptionButton opSign 
         Caption         =   "Print All Price Changes"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   4095
      End
   End
   Begin VB.Label lblDC 
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
      Left            =   5280
      TabIndex        =   32
      Top             =   6390
      Width           =   2415
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Item Number:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   31
      Top             =   6840
      Width           =   1485
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Line Code:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   30
      Top             =   6510
      Width           =   1485
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "To Date Updated:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   29
      Top             =   6180
      Width           =   1485
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "From Date Updated:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   28
      Top             =   5850
      Width           =   1485
   End
   Begin VB.Label generalLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Sign Name:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   27
      Top             =   5490
      Width           =   1485
   End
   Begin VB.Label generalLabel 
      Alignment       =   2  'Center
      Caption         =   "Print Signs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2490
      TabIndex        =   26
      Top             =   30
      Width           =   2265
   End
End
Attribute VB_Name = "PrintSigns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DO_PRINT As Long = 0
Private Const DO_PREVIEW As Long = 1

Private Const CHANGED_SINCE As Long = 0
Private Const ALL As Long = 1
Private Const ALL_LC As Long = 2
Private Const BY_DATE As Long = 3
Private Const BY_LINE As Long = 4
Private Const SINGLE_SIGN As Long = 5
Private Const SPEC_SIGN As Long = 6
Private Const SPEC_SIGN_LC As Long = 7
Private Const TRIAD_CODE As Long = 8
Private Const GROUP_SIGN As Long = 9
Private Const MISC_SIGN As Long = 10

Private currentSignType As Long

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Mouse.Hourglass True
    DB.Execute "DELETE FROM SignPrintMessageSelected"
    Dim viewMode As Long
    If Me.opPrint(DO_PRINT) Then
        viewMode = acViewNormal
    ElseIf Me.opPrint(DO_PREVIEW) Then
        viewMode = acViewPreview
    End If
    
    'sanity checks
    Dim die As Boolean, infoSign As Boolean, overridePrintFlag As Boolean
    Dim rst As ADODB.Recordset
    die = False
    Select Case True
        Case Is = Me.opSign(CHANGED_SINCE)
            Set rst = DB.retrieve("SELECT DISTINCT SignID FROM PartNumbers WHERE PrintSign=1 AND ItemNumber IN (SELECT ItemNumber FROM InventoryPriceLog WHERE IsStorePrice=1 AND TimeChanged>'" & Me.DateDropdown1.value & "')") ' & Me.priceChangeDate & "')")
            If rst.EOF Then
                MsgBox "No items to print"
                die = True
            End If
        Case Is = Me.opSign(ALL)
            Set rst = DB.retrieve("SELECT DISTINCT SignID FROM PartNumbers WHERE PriceChanged=1 AND PrintSign=1")
            If rst.EOF Then
                MsgBox "No items with price changes found."
                die = True
            End If
        Case Is = Me.opSign(ALL_LC)
            If Me.LineCode = "" Then
                MsgBox "Enter a line code."
                die = True
            Else
                Set rst = DB.retrieve("SELECT DISTINCT SignID FROM PartNumbers, InventoryMaster WHERE PartNumbers.ItemNumber=InventoryMaster.ItemNumber AND InventoryMaster.ProductLine='" & Me.LineCode & "' AND PriceChanged=1 AND PrintSign=1")
                If rst.EOF Then
                    MsgBox "No items with price changes found."
                    die = True
                End If
            End If
        Case Is = Me.opSign(BY_DATE)
            If Me.DateFrom = "" Or Me.DateTo = "" Then
                MsgBox "Enter a from date and to date."
                die = True
            Else
                Set rst = DB.retrieve("SELECT DISTINCT SignID FROM PartNumbers WHERE LastUpdated BETWEEN '" & Me.DateFrom & "' AND '" & Me.DateTo & "' AND PrintSign=1")
                If rst.EOF Then
                    MsgBox "No items found between these dates."
                    die = True
                End If
            End If
        Case Is = Me.opSign(BY_LINE)
            If Me.LineCode = "" Then
                MsgBox "Enter a line code."
                die = True
            Else
                Set rst = DB.retrieve("SELECT DISTINCT SignID FROM PartNumbers, InventoryMaster WHERE PartNumbers.ItemNumber=InventoryMaster.ItemNumber AND InventoryMaster.ProductLine='" & Me.LineCode & "' AND PrintSign=1")
                If rst.EOF Then
                    MsgBox "No items in line code."
                    die = True
                End If
            End If
        Case Is = Me.opSign(SINGLE_SIGN)
            If Me.ItemNumber = "" Then
                MsgBox "Select an item."
                die = True
            Else
                Set rst = DB.retrieve("SELECT PrintSign FROM PartNumbers WHERE ItemNumber='" & Me.ItemNumber & "'")
                If rst.EOF Then
                    MsgBox "Couldn't find part number? It was just there a minute ago..."
                    die = True
                Else
                    If rst("PrintSign") = 0 Then
                        If MsgBox("This item doesn't need to sign. To print one anyway, click yes.", vbYesNo) = vbNo Then
                            die = True
                        Else
                            overridePrintFlag = True
                        End If
                    End If
                End If
                If die = False Then
                    Set rst = DB.retrieve("SELECT DISTINCT SignID FROM PartNumbers WHERE ItemNumber='" & Me.ItemNumber & "'")
                End If
            End If
        Case Is = Me.opSign(SPEC_SIGN)
            If Me.SignName = "" Then
                MsgBox "Select a sign."
                die = True
            Else
                Set rst = DB.retrieve("SELECT DISTINCT SignID FROM PartNumbers WHERE SignID=(SELECT ID FROM SignNames WHERE SignName='" & Me.SignName & "') AND PrintSign=1")
                If rst.EOF Then
                    MsgBox "This sign is not associated with any part numbers."
                    die = True
                End If
            End If
        Case Is = Me.opSign(SPEC_SIGN_LC)
            If Me.SignName = "" Or Me.LineCode = "" Then
                MsgBox "Select a sign and line code."
                die = True
            Else
                Set rst = DB.retrieve("SELECT DISTINCT SignID FROM PartNumbers, InventoryMaster WHERE PartNumbers.ItemNumber=InventoryMaster.ItemNumber AND SignID=(SELECT SignID FROM SignNames WHERE SignName='" & Me.SignName & "') AND InventoryMaster.ProductLine='" & Me.LineCode & "' AND PrintSign=1")
                If rst.EOF Then
                    MsgBox "No items in line code with this sign."
                    die = True
                End If
            End If
        Case Is = Me.opSign(TRIAD_CODE)
            If Me.TriadCode = "" Then
                MsgBox "Enter the Triad Code."
                die = True
            End If
            Set rst = DB.retrieve("SELECT DISTINCT SignID FROM PartNumbers WHERE TriadCode='" & Me.TriadCode & "'")
            If rst.EOF Then
                MsgBox "No items found with this Triad Code."
                die = True
            End If
        Case Is = Me.opSign(GROUP_SIGN)
            If Me.SignName = "" Then
                MsgBox "Select a group sign."
                die = True
            End If
        Case Is = Me.opSign(MISC_SIGN)
            If Me.SignName = "" Then
                MsgBox "Select a miscellaneous sign."
                die = True
            End If
        Case Else
            MsgBox "No option selected?"
            die = True
    End Select
    If die Then
        If Not rst Is Nothing Then
            rst.Close
            Set rst = Nothing
        End If
        Mouse.Hourglass False
        Exit Sub
    End If

    Dim whereClause As String
    'set up info for which items to grab
    'rst is already open from sanity checks
    Select Case True
        Case Is = Me.opSign(ALL)
            whereClause = "PriceChanged=-1 AND PrintSign=-1"
        Case Is = Me.opSign(ALL_LC)
            whereClause = "LineCode='" & Me.LineCode & "' AND PriceChanged=-1 AND PrintSign=-1"
        Case Is = Me.opSign(BY_DATE)
            whereClause = "LastChanged BETWEEN '" & Me.DateFrom & "' AND '" & Me.DateTo & "' AND PrintSign=-1"
        Case Is = Me.opSign(BY_LINE)
            whereClause = "LineCode='" & Me.LineCode & "' AND PrintSign=-1"
        Case Is = Me.opSign(SINGLE_SIGN)
            whereClause = "ItemNumber='" & Me.ItemNumber & IIf(overridePrintFlag, "'", "' AND PrintSign=-1")
        Case Is = Me.opSign(SPEC_SIGN)
            whereClause = "SignName='" & Me.SignName & "' AND PrintSign=-1"
        Case Is = Me.opSign(SPEC_SIGN_LC)
            whereClause = "SignName='" & Me.SignName & "' AND LineCode='" & Me.LineCode & "' AND PrintSign=-1"
        Case Is = Me.opSign(TRIAD_CODE)
            whereClause = "TriadCode='" & Me.TriadCode & "'"
        Case Is = Me.opSign(CHANGED_SINCE)
            whereClause = "DateLastPriceChange>#" & Me.DateDropdown1.value & "# AND PrintSign=-1" 'Me.priceChangeDate & "# AND PrintSign=-1"
    End Select
    
    If Me.opSign(GROUP_SIGN) Then
        OpenSignInAccess "GROUP " & Me.SignName, viewMode, ""
    ElseIf Me.opSign(MISC_SIGN) Then
        OpenSignInAccess "MISC " & Me.SignName, viewMode, ""
    Else
        Dim multisign As Boolean
        multisign = rst.RecordCount > 1 And Me.opPrint(DO_PREVIEW)
        While Not rst.EOF
            OpenSignInAccess SignHashIDtoName.item(rst("SignID").value), viewMode, "SignName='" & SignHashIDtoName.item(rst("SignID").value) & "' AND " & whereClause, multisign, overridePrintFlag
            rst.MoveNext
        Wend
        If multisign Then
            AccessIntegration.dropHandle
        End If
        rst.Close
        Set rst = Nothing
    End If
    
    'dbExecute "UPDATE PartNumbers SET PriceChanged=0 WHERE " & whereClause
    
    Mouse.Hourglass False
End Sub

Private Sub btnResetCodes_Click()
    If MsgBox("Reset ""Triad Code"" for all items?", vbYesNo) = vbYes Then
        DB.Execute "UPDATE PartNumbers SET TriadCode=Null WHERE TriadCode IS NOT NULL"
    End If
End Sub

Private Sub btnResetFlags_Click()
    If MsgBox("Reset all ""to be printed"" flags?", vbYesNo) = vbYes Then
        DB.Execute "UPDATE PartNumbers SET PriceChanged=0 WHERE PriceChanged=1"
    End If
End Sub

Private Sub Command1_Click()
    'Shell "s:\mastest\mas90-signs\signs\print_signs.bat", vbNormalFocus
    'Open Environ("TEMP") & "\poinv\signshelp.txt" For Output As #1
    'Print #1, "Signs is loading now, it takes a while....I'll see if I can get it faster." & vbCrLf & vbCrLf
    'Print #1, "Select the group of signs you want to print, and hit the ""Add"" button. Then, hit the ""Print"" or ""Preview"" buttons." & vbCrLf & vbCrLf
    'Print #1, "You can also load up a bunch of items all at once, just keep adding things to the list before printing." & vbCrLf & vbCrLf
    'Print #1, "Most of the different signs should be working, the ones that I know aren't are the CARHARTT and HARTFORD ones. I'm going to talk to Jeanne at some point to see what exactly she wants to do with them, so until then you'll have to use the old Signs program to print them." & vbCrLf & vbCrLf
    'Print #1, "Any problems or suggestions, just let me know."
    'Close #1
    'OpenDefaultApp Environ("TEMP") & "\poinv\signshelp.txt"
    Shell WPERL & " " & "s:\mastest\mas90-signs\signs\signs.pl", vbNormalFocus
End Sub

Private Sub DateDropdown1_GotFocus()
    Me.opSign(CHANGED_SINCE) = True
End Sub

Private Sub Form_Load()
    Me.DateDropdown1.value = Date
    requerySignName
    requeryItemNumber
    requeryTriadCode
    Me.opSign(ALL) = True
    Me.opPrint(DO_PREVIEW) = True
    Disable Me.SignName
    Disable Me.DateFrom
    Disable Me.DateTo
    Disable Me.LineCode
    Disable Me.ItemNumber
End Sub

Private Sub hideDCItems_Click()
    requeryItemNumber
End Sub

Private Sub ItemNumber_Click()
    ItemNumber_LostFocus
End Sub

Private Sub ItemNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.ItemNumber, KeyCode, Shift
End Sub

Private Sub ItemNumber_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.ItemNumber, KeyAscii
    If KeyAscii = vbKeyReturn Then
        ItemNumber_LostFocus
    End If
End Sub

Private Sub ItemNumber_LostFocus()
    AutoCompleteLostFocus Me.ItemNumber
    If Me.ItemNumber <> "" Then
        Me.SignName = SignHashIDtoName.item(CInt(DLookup("SignID", "PartNumbers", "ItemNumber='" & Me.ItemNumber & "'")))
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("exec spPrintSignItemInfo '" & Me.ItemNumber & "'")
        Me.lblDC.caption = rst("DiscLabel")
        rst.Close
        Set rst = Nothing
    End If
End Sub

Private Sub opSign_Click(Index As Integer)
    Disable Me.SignName
    Disable Me.DateFrom
    Disable Me.DateTo
    Disable Me.LineCode
    Disable Me.ItemNumber
    Select Case Index
        Case Is = CHANGED_SINCE
            'do nothing
        Case Is = ALL
            'do nothing
        Case Is = ALL_LC
            Enable Me.LineCode
            Me.LineCode.SetFocus
        Case Is = BY_DATE
            Enable Me.DateFrom
            Enable Me.DateTo
            Me.DateFrom.SetFocus
        Case Is = BY_LINE
            Enable Me.LineCode
            Me.LineCode.SetFocus
        Case Is = SINGLE_SIGN
            Enable Me.ItemNumber
            Enable Me.SignName
            Me.ItemNumber.SetFocus
        Case Is = SPEC_SIGN
            Enable Me.SignName
            Me.SignName.SetFocus
        Case Is = SPEC_SIGN_LC
            Enable Me.SignName
            Enable Me.LineCode
            Me.SignName.SetFocus
        Case Is = GROUP_SIGN
            Enable Me.SignName
        Case Is = MISC_SIGN
            Enable Me.SignName
    End Select
    requeryIfNeeded Index
End Sub

'Private Sub priceChangeDate_GotFocus()
'    Me.opSign(CHANGED_SINCE) = True
'End Sub
'
'Private Sub priceChangeDate_LostFocus()
'    If Me.priceChangeDate = "" Then
'        Me.priceChangeDate = Date
'    ElseIf Not IsDate(Me.priceChangeDate) Then
'        MsgBox "Invalid date"
'        Me.priceChangeDate = Date
'    Else
'        Me.priceChangeDate = Format(Me.priceChangeDate, "m/d/yyyy")
'    End If
'End Sub

Private Sub SignName_Click()
    SignName_LostFocus
End Sub

Private Sub SignName_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.SignName, KeyCode, Shift
End Sub

Private Sub SignName_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.SignName, KeyAscii
    If KeyAscii = vbKeyReturn Then
        SignName_LostFocus
    End If
End Sub

Private Sub SignName_LostFocus()
    AutoCompleteLostFocus Me.SignName
    Dim oldsign As String
    If Me.opSign(SINGLE_SIGN) And Me.ItemNumber <> "" And Me.SignName <> "" Then
        oldsign = SignHashIDtoName.item(CInt(DLookup("SignID", "PartNumbers", "ItemNumber='" & Me.ItemNumber & "'")))
        If Me.SignName <> oldsign Then
            If MsgBox("Change sign for " & Me.ItemNumber & "?" & vbCrLf & vbCrLf & "From: " & oldsign & vbCrLf & "To: " & Me.SignName, vbYesNo) = vbYes Then
                DB.Execute "UPDATE PartNumbers SET SignID=(SELECT ID FROM SignNames WHERE SignName='" & Me.SignName & "') WHERE ItemNumber='" & Me.ItemNumber & "'"
            Else
                Me.SignName = oldsign
            End If
        End If
    ElseIf Me.opSign(SINGLE_SIGN) And Me.ItemNumber <> "" Then
        oldsign = SignHashIDtoName.item(CInt(DLookup("SignID", "PartNumbers", "ItemNumber='" & Me.ItemNumber & "'")))
        Me.SignName = oldsign
    End If
End Sub

Private Sub requerySignName(Optional signType As Long = 0)
    Dim rst As ADODB.Recordset
    Select Case signType
        Case Is = GROUP_SIGN
            Set rst = DB.retrieve("SELECT GroupName AS SignName FROM GroupSigns ORDER BY GroupName")
        Case Is = MISC_SIGN
            Set rst = DB.retrieve("SELECT MiscSignName AS SignName FROM MiscSigns ORDER BY MiscSignName")
        Case Else
            Set rst = DB.retrieve("SELECT SignName FROM SignNames ORDER BY SignName")
    End Select
    Me.SignName.Clear
    While Not rst.EOF
        Me.SignName.AddItem rst("SignName")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    currentSignType = signType
    ExpandDropDownToFit Me.SignName
End Sub

Private Sub requeryItemNumber()
    Dim rst As ADODB.Recordset
    If Me.hideDCItems Then
        Set rst = DB.retrieve("SELECT ItemNumber FROM vNonDiscontinuedItems ORDER BY ItemNumber")
    Else
        Set rst = DB.retrieve("SELECT ItemNumber FROM PartNumbers ORDER BY ItemNumber")
    End If
    Me.ItemNumber.Clear
    While Not rst.EOF
        Me.ItemNumber.AddItem (rst("ItemNumber"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryIfNeeded(selection As Integer)
    Select Case selection
        Case Is = GROUP_SIGN
            If currentSignType <> GROUP_SIGN Then
                requerySignName GROUP_SIGN
            End If
        Case Is = MISC_SIGN
            If currentSignType <> MISC_SIGN Then
                requerySignName MISC_SIGN
            End If
        Case Else
            If currentSignType = GROUP_SIGN Or currentSignType = MISC_SIGN Then
                requerySignName
            End If
    End Select
End Sub

Private Sub requeryTriadCode()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT DISTINCT TriadCode FROM PartNumbers")
    While Not rst.EOF
        If Nz(rst("TriadCode")) <> "" Then
            Me.TriadCode.AddItem rst("TriadCode")
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub TriadCode_GotFocus()
    Me.opSign(TRIAD_CODE) = True
End Sub

Private Sub TriadCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
        Exit Sub
    End If
    If Len(Me.TriadCode) >= 3 Then
        KeyAscii = 0
    End If
End Sub
