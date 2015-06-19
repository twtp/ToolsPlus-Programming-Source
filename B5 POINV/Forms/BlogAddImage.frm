VERSION 5.00
Object = "{39C208C4-2615-4D49-911A-50F903B45C86}#1.0#0"; "TPControls.ocx"
Begin VB.Form BlogAddImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Picture"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TPControls.Picture Picture1 
      Height          =   3915
      Left            =   4620
      TabIndex        =   10
      Top             =   30
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   6906
   End
   Begin VB.TextBox articleID 
      Height          =   285
      Left            =   2370
      TabIndex        =   9
      Text            =   "hidden articleID"
      Top             =   60
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2580
      TabIndex        =   8
      Top             =   3420
      Width           =   1425
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   930
      TabIndex        =   7
      Top             =   3420
      Width           =   1425
   End
   Begin VB.TextBox description 
      Height          =   765
      Left            =   870
      TabIndex        =   6
      Top             =   2430
      Width           =   3375
   End
   Begin VB.OptionButton opPic 
      Caption         =   "Describe it, and someone will take a picture of it"
      Height          =   405
      Index           =   2
      Left            =   150
      TabIndex        =   5
      Top             =   1920
      Width           =   2625
   End
   Begin VB.CommandButton btnBrowse 
      Caption         =   "Browse ..."
      Height          =   285
      Left            =   2820
      TabIndex        =   4
      Top             =   1320
      Width           =   1245
   End
   Begin VB.OptionButton opPic 
      Caption         =   "From your computer/on network"
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   3
      Top             =   1320
      Width           =   2625
   End
   Begin VB.ComboBox ItemPicker 
      Height          =   315
      Left            =   2340
      TabIndex        =   2
      Top             =   630
      Width           =   2085
   End
   Begin VB.OptionButton opPic 
      Caption         =   "Standard Item Picture"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   660
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Add Picture"
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
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   1935
   End
End
Attribute VB_Name = "BlogAddImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REG_ITEM = 0
Private Const FROM_HDD = 1
Private Const DESCRIBE = 2



Private Sub btnCancel_Click()
    BlogManager.newImageID = -1
    Unload BlogAddImage
End Sub

Private Sub btnOK_Click()
    If Me.opPic(DESCRIBE) And Me.description = "" Then
        MsgBox "You must provide a description!"
    ElseIf Not Me.opPic(DESCRIBE) And Me.Picture1.GetImagePath = "" Then
        MsgBox "You haven't selected an image!"
    Else
        Dim imgid As String
        If Me.opPic(DESCRIBE) Then
            DB.Execute "INSERT INTO BlogImages ( ArticleID, Description ) VALUES ( " & Me.articleID & ", '" & EscapeSQuotes(Me.description) & "' )"
            imgid = DLookup("@@IDENTITY", "BlogImages")
        ElseIf Me.opPic(REG_ITEM) Then
            DB.Execute "INSERT INTO BlogImages ( ArticleID, UNCAddress, Description ) VALUES ( " & Me.articleID & ", '" & EscapeSQuotes(Me.Picture1.GetImagePath) & "', '" & Me.ItemPicker & "' )"
            imgid = DLookup("@@IDENTITY", "BlogImages")
        ElseIf Me.opPic(FROM_HDD) Then
            DB.Execute "INSERT INTO BlogImages ( ArticleID, UNCAddress, Description ) VALUES ( " & Me.articleID & ", '" & EscapeSQuotes(Me.Picture1.GetImagePath) & "', '" & EscapeSQuotes(filenameparse(Me.Picture1.GetImagePath)) & "' )"
            imgid = DLookup("@@IDENTITY", "BlogImages")
            Copy Me.Picture1.GetImagePath, BLOG_PIC_DIR & imgid & ".jpg"
            DB.Execute "UPDATE BlogImages SET UNCAddress='" & EscapeSQuotes(BLOG_PIC_DIR & imgid & ".jpg") & "' WHERE ID=" & imgid
        End If
        BlogManager.newImageID = imgid
        Unload BlogAddImage
    End If
End Sub

Private Sub ItemPicker_GotFocus()
    Me.opPic(REG_ITEM).value = True
End Sub

Private Sub btnBrowse_GotFocus()
    Me.opPic(FROM_HDD).value = True
End Sub

Private Sub description_GotFocus()
    Me.opPic(DESCRIBE).value = True
End Sub

Private Sub Form_Load()
    requeryItems
End Sub

Private Sub requeryItems()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber FROM vNonDiscontinuedItems ORDER BY ItemNumber")
    Me.ItemPicker.Clear
    While Not rst.EOF
        Me.ItemPicker.AddItem rst("ItemNumber")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    ExpandDropDownToFit Me.ItemPicker
End Sub



Private Sub ItemPicker_Click()
    ItemPicker_LostFocus
End Sub

Private Sub ItemPicker_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.ItemPicker, KeyCode, Shift
    If KeyCode = vbKeyReturn Then
        ItemPicker_LostFocus
    End If
End Sub

Private Sub ItemPicker_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.ItemPicker, KeyAscii
End Sub

Private Sub ItemPicker_LostFocus()
    AutoCompleteLostFocus Me.ItemPicker
    display GenerateItemPicPath(Me.ItemPicker, True)
End Sub

Private Sub btnBrowse_Click()
    Dim fp As FilePicker
    Set fp = New FilePicker
    fp.AddFilter "JPEG Images", "*.jpg"
    fp.SetParent Me.hwnd
    Dim fname As String
    fname = fp.ShowDialog
    display fname
End Sub

Private Sub display(path As String)
    If path = "" Then
        Me.Picture1.DisplayImage ""
    ElseIf Me.Picture1.GetImagePath <> path Then
        Me.Picture1.DisplayImage path
    End If
End Sub
