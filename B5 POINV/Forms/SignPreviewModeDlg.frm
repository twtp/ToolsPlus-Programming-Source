VERSION 5.00
Begin VB.Form SignPreviewModeDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preview / Print Mode"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton opPreview 
      Caption         =   "All In Current Filter"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ComboBox printer 
      Height          =   315
      Left            =   630
      TabIndex        =   8
      Top             =   2040
      Width           =   2115
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2340
      TabIndex        =   7
      Top             =   2490
      Width           =   1425
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "Preview"
      Height          =   435
      Left            =   630
      TabIndex        =   6
      Top             =   2490
      Width           =   1425
   End
   Begin VB.TextBox hiddenSignName 
      Height          =   225
      Left            =   3330
      TabIndex        =   4
      Text            =   "sign"
      Top             =   210
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox hiddenItemNumber 
      Height          =   225
      Left            =   3330
      TabIndex        =   3
      Text            =   "item"
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton opPreview 
      Caption         =   "All SIGNNAME"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3825
   End
   Begin VB.OptionButton opPreview 
      Caption         =   "All LINECODE w/ SIGNNAME"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3825
   End
   Begin VB.OptionButton opPreview 
      Caption         =   "Just ITEMNUMBER"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3825
   End
   Begin VB.Label Label2 
      Caption         =   "Printer:"
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   2070
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Preview / Print Mode"
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
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   3225
   End
End
Attribute VB_Name = "SignPreviewModeDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ONLYITEM As Long = 0
Private Const LCSIGN As Long = 1
Private Const SIGN As Long = 2
Private Const THIS_FILTER As Long = 3

Private Sub btnExit_Click()
    Unload SignPreviewModeDlg
End Sub

Public Sub btnOK_Click()
    Dim iter As Variant
    Dim whereClause As String
    Select Case True
        Case Is = Me.opPreview(ONLYITEM)
            whereClause = "item_number='" & Me.hiddenItemNumber & "'"
        Case Is = Me.opPreview(LCSIGN)
            whereClause = "product_line='" & Left(Me.hiddenItemNumber, 3) & "' AND sign_name='" & Me.hiddenSignName & "'"
        Case Is = Me.opPreview(SIGN)
            whereClause = "sign_name='" & Me.hiddenSignName & "'"
        Case Is = Me.opPreview(THIS_FILTER)
            Dim items() As String
            items = SignMaintenance.GetItemList
            For Each iter In items
                whereClause = whereClause & Chr(0) & "item_number='" & CStr(iter) & "'"
            Next iter
            whereClause = Replace(whereClause, Chr(0), "", 1, 1)
    End Select
    
    Dim clauses As Variant
    clauses = Split(whereClause, Chr(0))
    Open DESTINATION_DIR & "signstoprint.tsv" For Output As #1
    For Each iter In clauses
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT item_number, sign_name FROM vSigns WHERE " & CStr(iter))
        While Not rst.EOF
            Print #1, rst("item_number") & vbTab & rst("sign_name")
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
    Next iter
    Close #1
    
    Shell PERL & " " & "s:\mastest\mas90-signs\signs\print_single_sign.pl " & qq(DESTINATION_DIR & "signstoprint.tsv") & IIf(Me.printer = "Default Printer", "", " " & qq(Me.printer))
End Sub

Private Sub Form_Load()
    Me.opPreview(ONLYITEM) = True
    requeryPrinters
    Me.printer = "Default Printer"
End Sub

Private Sub hiddenItemNumber_Change()
    Dim i As Long
    For i = 0 To Me.opPreview.count - 1
        Me.opPreview(i).caption = Replace(Me.opPreview(i).caption, "ITEMNUMBER", Me.hiddenItemNumber)
        Me.opPreview(i).caption = Replace(Me.opPreview(i).caption, "LINECODE", Left(Me.hiddenItemNumber, 3))
    Next i
End Sub

Private Sub hiddenSignName_Change()
    Dim i As Long
    For i = 0 To Me.opPreview.count - 1
        Me.opPreview(i).caption = Replace(Me.opPreview(i).caption, "SIGNNAME", Me.hiddenSignName)
    Next i
End Sub

Private Sub printer_Click()
    printer_LostFocus
End Sub

Private Sub printer_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.printer, KeyCode, Shift
End Sub

Private Sub printer_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.printer, KeyAscii
    If KeyAscii = vbKeyReturn Then
        printer_LostFocus
    End If
End Sub

Private Sub printer_LostFocus()
    AutoCompleteLostFocus Me.printer
End Sub

Private Sub requeryPrinters()
    Me.printer.Clear
    Me.printer.AddItem "Default Printer"
    Dim printerA() As String
    Dim i As Long
    printerA = getPrinters
    For i = 0 To UBound(printerA)
        Me.printer.AddItem printerA(i)
    Next i
End Sub

