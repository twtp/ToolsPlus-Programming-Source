VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#3.5#0"; "TPControls.ocx"
Begin VB.Form AddCrossSells 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Cross Sells"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   14460
   StartUpPosition =   1  'CenterOwner
   Begin TPControls.Picture ItemPicture 
      Height          =   1815
      Left            =   6480
      TabIndex        =   59
      Top             =   60
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   3201
   End
   Begin VB.CommandButton btnSuperPollinate 
      Caption         =   "Super Pollinate"
      Height          =   495
      Left            =   13350
      TabIndex        =   58
      Top             =   1770
      Width           =   1035
   End
   Begin VB.CommandButton btnCrossPollinate 
      Caption         =   "Cross Pollinate"
      Height          =   495
      Left            =   13350
      TabIndex        =   57
      Top             =   1260
      Width           =   1035
   End
   Begin VB.CommandButton btnMoveLeftAll 
      Caption         =   "<= ALL"
      Height          =   435
      Left            =   6930
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   5550
      Width           =   585
   End
   Begin VB.CommandButton btnMoveRightAll 
      Caption         =   "=> ALL"
      Height          =   435
      Left            =   6930
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   5040
      Width           =   585
   End
   Begin VB.CommandButton btnAccessoriesPaste 
      Caption         =   "Paste"
      Height          =   225
      Left            =   11190
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   2070
      Width           =   855
   End
   Begin VB.CommandButton btnAccessoriesCopy 
      Caption         =   "Copy"
      Height          =   225
      Left            =   10320
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   2070
      Width           =   855
   End
   Begin VB.CommandButton btnJumpAddAccessory 
      Caption         =   "Chosen Cross Sells"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7650
      TabIndex        =   52
      ToolTipText     =   "Click this button to add a new accessory to the current item"
      Top             =   2010
      Width           =   2115
   End
   Begin VB.CommandButton btnJumpToCurrentItem 
      Caption         =   "Current Item:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   30
      TabIndex        =   51
      ToolTipText     =   "Click this button to jump to the item you want"
      Top             =   540
      Width           =   1275
   End
   Begin VB.CommandButton btnJumpToAllItems 
      Caption         =   "All Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   50
      ToolTipText     =   "Click this button to search for an item in the list"
      Top             =   2010
      Width           =   1635
   End
   Begin VB.Frame Frame3 
      Height          =   1515
      Left            =   9570
      TabIndex        =   47
      Top             =   30
      Width           =   3675
      Begin VB.OptionButton opDirection 
         Caption         =   "ACC_CAPTION_2"
         Height          =   585
         Index           =   1
         Left            =   90
         TabIndex        =   49
         Top             =   840
         Width           =   3195
      End
      Begin VB.OptionButton opDirection 
         Caption         =   "ACC_CAPTION_1"
         Height          =   645
         Index           =   0
         Left            =   90
         TabIndex        =   48
         Top             =   180
         Value           =   -1  'True
         Width           =   3195
      End
   End
   Begin VB.Frame FrameFBP 
      Caption         =   "Filter By Path"
      Height          =   2205
      Left            =   30
      TabIndex        =   35
      Top             =   5910
      Visible         =   0   'False
      Width           =   4185
      Begin VB.CommandButton btnPathFilterApply 
         Caption         =   "Apply Filter"
         Height          =   255
         Left            =   720
         TabIndex        =   46
         Top             =   1890
         Width           =   2955
      End
      Begin VB.ComboBox PathFilter 
         Height          =   315
         Index           =   5
         Left            =   900
         Sorted          =   -1  'True
         TabIndex        =   45
         Top             =   1530
         Width           =   3195
      End
      Begin VB.ComboBox PathFilter 
         Height          =   315
         Index           =   4
         Left            =   900
         Sorted          =   -1  'True
         TabIndex        =   44
         Top             =   1200
         Width           =   3195
      End
      Begin VB.ComboBox PathFilter 
         Height          =   315
         Index           =   3
         Left            =   900
         Sorted          =   -1  'True
         TabIndex        =   43
         Top             =   870
         Width           =   3195
      End
      Begin VB.ComboBox PathFilter 
         Height          =   315
         Index           =   2
         Left            =   900
         Sorted          =   -1  'True
         TabIndex        =   42
         Top             =   540
         Width           =   3195
      End
      Begin VB.ComboBox PathFilter 
         Height          =   315
         Index           =   1
         Left            =   900
         Sorted          =   -1  'True
         TabIndex        =   36
         Top             =   210
         Width           =   3195
      End
      Begin VB.Label Label9 
         Caption         =   "Path 5:"
         Height          =   255
         Left            =   210
         TabIndex        =   41
         Top             =   1590
         Width           =   645
      End
      Begin VB.Label Label8 
         Caption         =   "Path 4:"
         Height          =   255
         Left            =   210
         TabIndex        =   40
         Top             =   1260
         Width           =   645
      End
      Begin VB.Label Label7 
         Caption         =   "Path 3:"
         Height          =   255
         Left            =   210
         TabIndex        =   39
         Top             =   930
         Width           =   645
      End
      Begin VB.Label Label6 
         Caption         =   "Path 2:"
         Height          =   255
         Left            =   210
         TabIndex        =   38
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Path 1:"
         Height          =   255
         Left            =   210
         TabIndex        =   37
         Top             =   270
         Width           =   645
      End
   End
   Begin VB.CheckBox chkHideDiscontinued 
      Caption         =   "Hide Discontinued"
      Height          =   345
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8280
      Value           =   1  'Checked
      Width           =   1485
   End
   Begin VB.TextBox ItemSpec 
      Height          =   285
      Index           =   2
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   1560
      Width           =   6075
   End
   Begin VB.TextBox ItemSpec 
      Height          =   285
      Index           =   1
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   1260
      Width           =   6075
   End
   Begin VB.CheckBox chkFillPreview 
      Caption         =   "Fill Preview Info"
      Height          =   225
      Left            =   6960
      TabIndex        =   23
      Top             =   6720
      Value           =   1  'Checked
      Width           =   1425
   End
   Begin VB.Frame Frame2 
      Height          =   2025
      Left            =   6900
      TabIndex        =   22
      Top             =   6690
      Width           =   7515
      Begin TPControls.Picture AccessoryPicture 
         Height          =   1665
         Left            =   6180
         TabIndex        =   60
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   2937
      End
      Begin VB.TextBox PreviewSpec 
         Height          =   285
         Index           =   4
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1590
         Width           =   5955
      End
      Begin VB.TextBox PreviewSpec 
         Height          =   285
         Index           =   3
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1290
         Width           =   5955
      End
      Begin VB.TextBox PreviewSpec 
         Height          =   285
         Index           =   2
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   990
         Width           =   5955
      End
      Begin VB.TextBox PreviewSpec 
         Height          =   285
         Index           =   1
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   690
         Width           =   5955
      End
      Begin VB.TextBox PreviewDescription 
         Height          =   285
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   330
         Width           =   4425
      End
      Begin VB.TextBox PreviewItemNumber 
         Height          =   285
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   330
         Width           =   1455
      End
   End
   Begin VB.ComboBox ItemSelector 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   510
      Width           =   2145
   End
   Begin VB.TextBox GeneralLC 
      Height          =   315
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "LC"
      Top             =   330
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   435
      Left            =   13350
      TabIndex        =   19
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton btnMoveRight 
      Caption         =   "=>"
      Height          =   555
      Left            =   6930
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3450
      Width           =   585
   End
   Begin VB.CommandButton btnMoveLeft 
      Caption         =   "<="
      Height          =   555
      Left            =   6930
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4110
      Width           =   585
   End
   Begin VB.CommandButton btnRefilter 
      Caption         =   "Re-Filter"
      Height          =   285
      Left            =   5700
      TabIndex        =   16
      Top             =   8880
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter:"
      Height          =   615
      Left            =   30
      TabIndex        =   13
      Top             =   8130
      Width           =   3825
      Begin VB.CommandButton btnExpandPathFilter 
         Caption         =   "q"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3300
         TabIndex        =   34
         Top             =   210
         Width           =   375
      End
      Begin VB.OptionButton opAccFilter 
         Caption         =   "By Path"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   33
         Top             =   240
         Width           =   915
      End
      Begin VB.OptionButton opAccFilter 
         Caption         =   "All Lines"
         Height          =   255
         Index           =   1
         Left            =   1350
         TabIndex        =   15
         Top             =   240
         Width           =   1125
      End
      Begin VB.OptionButton opAccFilter 
         Caption         =   "Current Line"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin TPControls.SimpleListView FullList 
      Height          =   5775
      Left            =   30
      TabIndex        =   12
      Top             =   2310
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   10186
      MultiSelect     =   0   'False
      Sorted          =   0   'False
      Enabled         =   -1  'True
   End
   Begin TPControls.SimpleListView CurrentAccessories 
      Height          =   4155
      Left            =   7590
      TabIndex        =   11
      Top             =   2310
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   7329
      MultiSelect     =   0   'False
      Sorted          =   0   'False
      Enabled         =   -1  'True
   End
   Begin VB.TextBox CurrentItemName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   900
      Width           =   6315
   End
   Begin VB.TextBox ItemNumber 
      Height          =   315
      Left            =   3690
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "ItemNumber"
      Top             =   0
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.HScrollBar barPosition 
      Height          =   285
      LargeChange     =   10
      Left            =   900
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8880
      Width           =   2115
   End
   Begin VB.TextBox txtPosition 
      Height          =   285
      Left            =   3930
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   8880
      Width           =   795
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
      Left            =   3540
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   8880
      Width           =   315
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   8430
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
      Left            =   360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   8880
      Width           =   495
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
      Left            =   3030
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Add Cross Sells"
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
      Left            =   180
      TabIndex        =   8
      Top             =   90
      Width           =   3315
   End
   Begin VB.Label lblStatusBar 
      Height          =   255
      Left            =   7020
      TabIndex        =   7
      Top             =   8940
      Width           =   3375
   End
   Begin VB.Label lblMaxAmt 
      Caption         =   "of MAX"
      Height          =   195
      Left            =   4770
      TabIndex        =   6
      Top             =   8940
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   14460
      X2              =   0
      Y1              =   8820
      Y2              =   8820
   End
End
Attribute VB_Name = "AddCrossSells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FLT_LINE As Long = 0
Private Const FLT_NONE As Long = 1
Private Const FLT_PATH As Long = 2
Private Const DIR_REGULAR_ACC As Long = 0
Private Const DIR_REVERSE_ACC As Long = 1

Private Const ACC_CAPTION_1 As String = "Accessories to this item (items in the list below will show up on the %ITEM%'s webpage"
Private Const ACC_CAPTION_2 As String = "Other items that have this as an accessory (%ITEM% will show up on the pages in the list below)"

Private ItemList() As String
Private changed As Boolean
Private whichCtl As String


Private Sub btnAccessoriesPaste_Click()
    Mouse.Hourglass True
    Dim clip As String
    clip = ClipBoard_GetData()
    If Left(clip, 12) = "reg acc for " Or Left(clip, 12) = "rev acc for " Then
        Dim item As String
        item = Mid(clip, 13)
        If ExistsInPartNumbers(item) Then
            Dim direction As String
            direction = Left(clip, 3)
            If direction = "reg" Or direction = "rev" Then
                Dim rst As ADODB.Recordset
                Set rst = DB.retrieve("SELECT " & IIf(direction = "reg", "Accessory", "ItemNumber") & " FROM PartNumberAccessories WHERE " & IIf(direction = "reg", "ItemNumber", "Accessory") & "='" & item & "'")
                While Not rst.EOF
                    If rst(0) = Me.ItemNumber Then
                        rst(0) = item
                    End If
                    Select Case True
                        Case Is = Me.opDirection(DIR_REGULAR_ACC)
                            DB.Execute "INSERT INTO PartNumberAccessories ( ItemNumber, Accessory ) VALUES ( '" & Me.ItemNumber & "', '" & rst(0) & "' )", True
                        Case Is = Me.opDirection(DIR_REVERSE_ACC)
                            DB.Execute "INSERT INTO PartNumberAccessories ( ItemNumber, Accessory ) VALUES ( '" & rst(0) & "', '" & Me.ItemNumber & "' )", True
                    End Select
                    rst.MoveNext
                Wend
                rst.Close
                Set rst = Nothing
                barPosition_Change
            Else
                MsgBox "Invalid clipboard contents, invalid direction (this shouldn't ever happen)."
            End If
        Else
            MsgBox "Invalid clipboard contents, item not found."
        End If
    Else
        MsgBox "Invalid clipboard contents, not a valid copy string."
    End If
    Mouse.Hourglass False
End Sub

Private Sub btnAccessoriesCopy_Click()
    Select Case True
        Case Is = Me.opDirection(DIR_REGULAR_ACC)
            ClipBoard_SetData "reg acc for " & Me.ItemNumber
        Case Is = Me.opDirection(DIR_REVERSE_ACC)
            ClipBoard_SetData "rev acc for " & Me.ItemNumber
    End Select
End Sub

Private Sub loadThisItem(itemOrIndex As String)
    Dim pos As Long
    If IsNumeric(itemOrIndex) Then
        pos = CLng(itemOrIndex)
    Else
        pos = bsearch(itemOrIndex, ItemList)
        If pos = -1 Then
            pos = 0
        End If
    End If
    If Me.barPosition.value = pos Then
        'barPosition_Change
    Else
        'Me.barPosition.value = pos
    End If
End Sub

Private Sub btnCrossPollinate_Click()
    'if msgbox
        Dim rst As ADODB.Recordset
        Select Case True
            Case Is = Me.opDirection(DIR_REGULAR_ACC)
                Set rst = DB.retrieve("SELECT Accessory FROM PartNumberAccessories WHERE ItemNumber='" & Me.ItemNumber & "'")
                While Not rst.EOF
                    DB.Execute "INSERT INTO PartNumberAccessories ( ItemNumber, Accessory ) VALUES ( '" & rst(0) & "', '" & Me.ItemNumber & "' )", True
                    rst.MoveNext
                Wend
                rst.Close
                Set rst = Nothing
            Case Is = Me.opDirection(DIR_REVERSE_ACC)
                Set rst = DB.retrieve("SELECT ItemNumber FROM PartNumberAccessories WHERE Accessory='" & Me.ItemNumber & "'")
                While Not rst.EOF
                    DB.Execute "INSERT INTO PartNumberAccessories ( ItemNumber, Accessory ) VALUES ( '" & Me.ItemNumber & "', '" & rst(0) & "' )", True
                    rst.MoveNext
                Wend
                rst.Close
                Set rst = Nothing
        End Select
        
    'end if
End Sub

Private Sub btnSuperPollinate_Click()
    If vbNo = MsgBox("Do you know what you're doing?" & vbCrLf & vbCrLf & "Seriously, don't use this. At least check to make sure we have good backups first.", vbYesNo + vbDefaultButton2) Then
        Exit Sub
    End If
    
    Dim rst As ADODB.Recordset
    Select Case True
        Case Is = Me.opDirection(DIR_REGULAR_ACC)
            Set rst = DB.retrieve("SELECT Accessory FROM PartNumberAccessories WHERE ItemNumber='" & Me.ItemNumber & "'")
        Case Is = Me.opDirection(DIR_REVERSE_ACC)
            Set rst = DB.retrieve("SELECT ItemNumber FROM PartNumberAccessories WHERE Accessory='" & Me.ItemNumber & "'")
    End Select
    
    'convert to a normal list
    Dim rstArray As PerlArray
    Set rstArray = New PerlArray
    rstArray.Push CStr(Me.ItemNumber)
    While Not rst.EOF
        rstArray.Push CStr(rst(0))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    Dim pairList As PerlArray
    Set pairList = New PerlArray
    Dim i As Long, j As Long
    For i = 0 To rstArray.Scalar - 1
        For j = 0 To rstArray.Scalar - 1
            If i <> j Then
                pairList.Push Array(rstArray.Elem(i), rstArray.Elem(j))
            End If
        Next j
    Next i
    
    Debug.Print pairList.Scalar
    If vbYes = MsgBox("Prepared to insert " & pairList.Scalar & " accessories...continue?", vbYesNo + vbDefaultButton2) Then
        For i = 0 To pairList.Scalar - 1
            DB.Execute "INSERT INTO PartNumberAccessories ( ItemNumber, Accessory ) VALUES ( '" & pairList.Elem(i)(0) & "', '" & pairList.Elem(i)(1) & "' )", True
        Next i
        MsgBox "Done!"
    Else
        MsgBox "Aborted!"
    End If
End Sub

Private Sub btnJumpAddAccessory_Click()
    Dim item As String
    item = InputBox("Enter the item you want to add as an accessory to the " & Me.ItemNumber & " (you can also scan a barcode in here):")
    item = UCase(item)
    If item <> "" Then
        If IsNumeric(item) Then
            MsgBox "This item is not barcoded! Barcode it first, then you can put it in here"
        ElseIf Not ExistsInPartNumbers(item) Then
            MsgBox "This item does not exist in the part # table!", vbCritical
        ElseIf Me.CurrentAccessories.InList(item) Then
            MsgBox "This accessory is already loaded!", vbInformation
        Else
            Select Case True
                Case Is = Me.opDirection(DIR_REGULAR_ACC)
                    DB.Execute "INSERT INTO PartNumberAccessories ( ItemNumber, Accessory ) VALUES ( '" & Me.ItemNumber & "', '" & item & "' )"
                    barPosition_Change
                    MsgBox item & " added successfully to " & Me.ItemNumber
                Case Is = Me.opDirection(DIR_REVERSE_ACC)
                    DB.Execute "INSERT INTO PartNumberAccessories ( ItemNumber, Accessory ) VALUES ( '" & item & "', '" & Me.ItemNumber & "' )"
                    barPosition_Change
                    MsgBox Me.ItemNumber & " added successfully to " & item
            End Select
        End If
    End If
End Sub

Private Sub btnJumpToAllItems_Click()
    Dim i As Long
    Dim found As Boolean
    found = False

    Dim item As String
    item = InputBox("Enter item to find (you can also scan a barcode in here):")
    item = UCase(item)
    If item <> "" Then
        If IsNumeric(item) Then
            MsgBox "This item is not barcoded! Barcode it first, then you can put it in here"
        ElseIf Not ExistsInPartNumbers(item) Then
            MsgBox "Can't find this item in the part # table!"
        ElseIf Me.CurrentAccessories.InList(item) Then
            For i = 0 To Me.CurrentAccessories.ListCount - 1
                If Me.CurrentAccessories.GetText(i, 0) = item Then
                    Me.CurrentAccessories.SelectRow i
                    found = True
                    Exit For
                End If
            Next i
            If Not found Then
                MsgBox "Can't find the item, but .InList() says it's there...Talk to Brian!"
            End If
        Else
retrylist:
            If Me.FullList.InList(item) Then
                For i = 0 To Me.FullList.ListCount - 1
                    If Me.FullList.GetText(i, 0) = item Then
                        Me.FullList.SelectRow i
                        found = True
                        Exit For
                    End If
                Next i
                If Not found Then
                    MsgBox "Can't find the item, but .InList() says it's there...Talk to Brian!"
                End If
            Else
                If Me.opAccFilter(FLT_NONE) = False Then
                    Me.opAccFilter(FLT_NONE) = True
                    GoTo retrylist
                ElseIf Me.chkHideDiscontinued Then
                    Me.chkHideDiscontinued = False
                    GoTo retrylist
                Else
                    MsgBox "Can't find this item in the list!"
                End If
            End If
        End If
    End If
End Sub

Private Sub btnJumpToCurrentItem_Click()
    Dim item As String
    item = InputBox("Enter the item to jump to here (you can also scan a barcode in here):")
    item = UCase(item)
    If item <> "" Then
        If IsNumeric(item) Then
            MsgBox "This item is not barcoded! Barcode it first, then you can put it in here."
        ElseIf Not ExistsInPartNumbers(item) Then
            MsgBox "Can't find this item in the part # table!"
        ElseIf InCombo(item, Me.ItemSelector, True) Then
            Me.ItemSelector = item
        Else
            MsgBox "This item can't be found in the item selector... Talk to Brian!"
        End If
    End If
End Sub

Private Sub barPosition_Change()
    Mouse.Hourglass True
    If changed = True Then
        Select Case whichCtl
            Case Else
                MsgBox "Could not determine which control you were updating, so not saving."
        End Select
    End If
    If ItemList(Me.barPosition.value) = "NORECORDS" Then
        MsgBox "No records found."
    Else
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT PartNumbers.ItemNumber, GeneralProductLine, Desc1, Desc2, Desc3, Spec1HTML, Spec2HTML, vItemStatusBreakdown.DConWeb AS ItemDiscontinued FROM PartNumbers INNER JOIN vGeneralProductLine ON PartNumbers.ItemNumber=vGeneralProductLine.ItemNumber INNER JOIN vItemStatusBreakdown ON PartNumbers.ItemNumber=vItemStatusBreakdown.ItemNumber WHERE PartNumbers.ItemNumber='" & ItemList(Me.barPosition.value) & "'")
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

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnExpandPathFilter_Click()
    If Not Me.FrameFBP.Visible Then
        'expand
        Me.FrameFBP.Visible = True
        Me.btnExpandPathFilter.Caption = "q"
    Else
        'contract
        Me.FrameFBP.Visible = False
        Me.btnExpandPathFilter.Caption = "p"
    End If
End Sub

Private Sub btnFirstItem_Click()
    Me.barPosition.value = Me.barPosition.Min
End Sub

Private Sub btnLastItem_Click()
    Me.barPosition.value = Me.barPosition.max
End Sub

Private Sub btnMoveLeft_Click()
    If Me.CurrentAccessories.SelIndex <> -1 Then
        Dim rst As ADODB.Recordset
        Select Case True
            Case Is = Me.opDirection(DIR_REGULAR_ACC)
                DB.Execute "DELETE FROM PartNumberAccessories WHERE ItemNumber='" & Me.ItemNumber & "' AND Accessory='" & Me.CurrentAccessories.GetTextSelected(0) & "'"
            Case Is = Me.opDirection(DIR_REVERSE_ACC)
                DB.Execute "DELETE FROM PartNumberAccessories WHERE ItemNumber='" & Me.CurrentAccessories.GetTextSelected(0) & "' AND Accessory='" & Me.ItemNumber & "'"
        End Select
        Select Case True
            Case Is = Me.opAccFilter(FLT_NONE)
                Me.FullList.Add Me.CurrentAccessories.GetRowSelected
            Case Is = Me.opAccFilter(FLT_LINE)
                Set rst = DB.retrieve("SELECT COUNT(*) FROM PartNumbers INNER JOIN vGeneralProductLine ON PartNumbers.ItemNumber=vGeneralProductLine.ItemNumber INNER JOIN vItemStatusBreakdown ON PartNumbers.ItemNumber=vItemStatusBreakdown.ItemNumber WHERE PartNumbers.ItemNumber='" & Me.CurrentAccessories.GetTextSelected(0) & "' AND (WebPublished=1 OR WebToBePublished=1) AND " & IIf(Me.chkHideDiscontinued, "vItemStatusBreakdown.DC=0 AND ", "") & "GeneralProductLine='" & Me.GeneralLC & "'")
                If rst(0) <> 0 Then
                    'Me.FullList.Add Me.CurrentAccessories.GetRowSelected
                    'Me.FullList.Add Array(Me.CurrentAccessories.GetTextSelected(0), Me.CurrentAccessories.GetTextSelected(1))
                    Me.FullList.Add Array(Me.CurrentAccessories.GetTextSelected(0), Me.CurrentAccessories.GetTextSelected(3))
                End If
                rst.Close
                Set rst = Nothing
            Case Is = Me.opAccFilter(FLT_PATH)
                'Dim sql As String
                'sql = "SELECT COUNT(*) FROM PartNumbers INNER JOIN vItemStatusBreakdown ON PartNumbers.ItemNumber=vItemStatusBreakdown.ItemNumber WHERE PartNumbers.ItemNumber='" & Me.CurrentAccessories.GetTextSelected(0) & "' AND (WebPublished=1 OR WebToBePublished=1)" & IIf(Me.chkHideDiscontinued, " AND vItemStatusBreakdown.DC=0", "")
                'Dim i As Long
                'For i = 1 To 5
                '    If Me.PathFilter(i) <> "" Then
                '        sql = sql & " AND EXISTS (SELECT * FROM vPNWebPaths WHERE ItemNumber=PartNumbers.ItemNumber AND WebPathName"
                '        If InCombo(Me.PathFilter(i), Me.PathFilter(i), True) Then
                '            sql = sql & "='" & EscapeSQuotes(Replace(Me.PathFilter(i), "'   ", "")) & "'"
                '        Else
                '            sql = sql & " LIKE '%" & EscapeSQuotes(Replace(Me.PathFilter(i), "'   ", "")) & "%'"
                '        End If
                '        sql = sql & " AND PathLevel=" & i & ")"
                '    End If
                'Next i
                'Set rst = DB.retrieve(sql)
                'If rst(0) <> 0 Then
                '    'Me.FullList.Add Me.CurrentAccessories.GetRowSelected
                '    Me.FullList.Add Array(Me.CurrentAccessories.GetTextSelected(0), Me.CurrentAccessories.GetTextSelected(1))
                'End If
                'rst.Close
                'Set rst = Nothing
                MsgBox "filter by path is TODO!"
        End Select
        Me.CurrentAccessories.RemoveSelected
    End If
End Sub

Private Sub btnMoveLeftAll_Click()
    If Me.CurrentAccessories.ListCount > 0 Then
        If vbYes = MsgBox("Are you sure you want to REMOVE all accessories from the " & Me.ItemNumber & "?", vbYesNo) Then
            Dim i As Long
            For i = 0 To Me.CurrentAccessories.ListCount - 1
                Select Case True
                    Case Is = Me.opDirection(DIR_REGULAR_ACC)
                        DB.Execute "DELETE FROM PartNumberAccessories WHERE ItemNumber='" & Me.ItemNumber & "' AND Accessory='" & Me.CurrentAccessories.GetText(i, 0) & "'"
                    Case Is = Me.opDirection(DIR_REVERSE_ACC)
                        DB.Execute "DELETE FROM PartNumberAccessories WHERE ItemNumber='" & Me.CurrentAccessories.GetText(i, 0) & "' AND Accessory='" & Me.ItemNumber & "'"
                End Select
            Next i
            Me.CurrentAccessories.Clear
            fillFullList Me.ItemNumber, Me.GeneralLC
        End If
    End If
End Sub

Private Sub btnMoveRight_Click()
    If Me.FullList.SelIndex <> -1 Then
        Select Case True
            Case Is = Me.opDirection(DIR_REGULAR_ACC)
                DB.Execute "INSERT INTO PartNumberAccessories ( ItemNumber, Accessory ) VALUES ( '" & Me.ItemNumber & "', '" & Me.FullList.GetTextSelected(0) & "' )"
            Case Is = Me.opDirection(DIR_REVERSE_ACC)
                DB.Execute "INSERT INTO PartNumberAccessories ( ItemNumber, Accessory ) VALUES ( '" & Me.FullList.GetTextSelected(0) & "', '" & Me.ItemNumber & "' )"
        End Select
        'Me.CurrentAccessories.Add Me.FullList.GetRowSelected
        'Me.CurrentAccessories.Add Array(Me.FullList.GetTextSelected(0), Me.FullList.GetTextSelected(1), "N")
        'Me.CurrentAccessories.Add Array(Me.FullList.GetTextSelected(0), Me.FullList.GetTextSelected(1))
        Me.CurrentAccessories.Add Array(Me.FullList.GetTextSelected(0), "??", "??", Me.FullList.GetTextSelected(1))
        Me.FullList.RemoveSelected
    End If
End Sub

Private Sub btnMoveRightAll_Click()
    If Me.FullList.ListCount > 0 Then
        If vbYes = MsgBox("Are you sure you want to add all " & Me.FullList.ListCount & " to the " & Me.ItemNumber & "?", vbYesNo) Then
            Dim i As Long
            For i = 0 To Me.FullList.ListCount - 1
                Select Case True
                    Case Is = Me.opDirection(DIR_REGULAR_ACC)
                        DB.Execute "INSERT INTO PartNumberAccessories ( ItemNumber, Accessory ) VALUES ( '" & Me.ItemNumber & "', '" & Me.FullList.GetText(i, 0) & "' )"
                    Case Is = Me.opDirection(DIR_REVERSE_ACC)
                        DB.Execute "INSERT INTO PartNumberAccessories ( ItemNumber, Accessory ) VALUES ( '" & Me.FullList.GetText(i, 0) & "', '" & Me.ItemNumber & "' )"
                End Select
                'Me.CurrentAccessories.Add Array(Me.FullList.GetText(i, 0), Me.FullList.GetText(i, 1), "N")
                Me.CurrentAccessories.Add Array(Me.FullList.GetText(i, 0), "??", "??", Me.FullList.GetText(i, 1))
            Next i
            Me.FullList.Clear
        End If
    End If
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

Private Sub btnPathFilterApply_Click()
    If Not Me.opAccFilter(FLT_PATH) Then
        Me.opAccFilter(FLT_PATH) = True
    End If
    fillFullList Me.ItemNumber, Me.GeneralLC
    btnExpandPathFilter_Click
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

Private Sub btnRefilter_Click()
    requeryForm Me.ItemNumber
End Sub


Private Sub fillForm(item As ADODB.Recordset)
    Me.ItemNumber = item("ItemNumber")
    Me.opDirection(DIR_REGULAR_ACC).Caption = Replace(ACC_CAPTION_1, "%ITEM%", item("ItemNumber"))
    Me.opDirection(DIR_REVERSE_ACC).Caption = Replace(ACC_CAPTION_2, "%ITEM%", item("ItemNumber"))
    Me.ItemSelector = item("ItemNumber")
    Me.GeneralLC = item("GeneralProductLine")
    Me.CurrentItemName = Nz(item("Desc2")) & IIf(Nz(item("Desc1")) <> "", " " & Nz(item("Desc1")), "") & " " & Nz(item("Desc3"))
    If Not CBool(item("ItemDiscontinued")) Then
        Me.ItemSpec(1) = Nz(item("Spec1HTML"))
        Me.ItemSpec(2) = Nz(item("Spec2HTML"))
    Else
        Me.ItemSpec(1) = FillDCSpecs(item("ItemNumber"), 1)
        Me.ItemSpec(2) = FillDCSpecs(item("ItemNumber"), 2)
    End If
    fillCurrentAccessories item("ItemNumber")
    fillFullList item("ItemNumber"), item("GeneralProductLine")
    requeryPaths
    Me.ItemPicture.DisplayImage PictureDBFunctions.GenerateItemPicturePathAny(Me.ItemNumber)
End Sub

Private Sub requeryForm(Optional saveItem As String = "")
    Mouse.Hourglass True
    
    Dim i As Long
    If IsFormLoaded("SignMaintenance") Then
        ItemList = SignMaintenance.GetItemList
    Else
        Dim rst As ADODB.Recordset
        Set rst = DB.retrieve("SELECT ItemNumber FROM vSignsRecSource ORDER BY ItemNumber")
        ReDim ItemList(rst.RecordCount - 1) As String
        For i = 0 To rst.RecordCount - 1
            ItemList(i) = rst("ItemNumber")
            rst.MoveNext
        Next i
        rst.Close
        Set rst = Nothing
    End If
    
    Me.ItemSelector.Clear
    For i = 0 To UBound(ItemList)
        Me.ItemSelector.AddItem ItemList(i)
    Next i
    Me.lblMaxAmt.Caption = "of " & UBound(ItemList) + 1
    'Me.barPosition.Min = 0
    'Me.barPosition.max = UBound(ItemList)
    If IsFormLoaded("SignMaintenance") Then
        loadThisItem SignMaintenance.ItemNumber
    ElseIf saveItem <> "" Then
        loadThisItem saveItem
    Else
        loadThisItem 0
    End If
    
    Mouse.Hourglass False
End Sub

Private Sub chkFillPreview_Click()
    Dim i As Long
    If Me.chkFillPreview Then
        Enable Me.PreviewItemNumber
        Enable Me.PreviewDescription
        For i = Me.PreviewSpec.LBound To Me.PreviewSpec.UBound
            Enable Me.PreviewSpec(i)
        Next i
    Else
        Disable Me.PreviewItemNumber
        Disable Me.PreviewDescription
        For i = Me.PreviewSpec.LBound To Me.PreviewSpec.UBound
            Disable Me.PreviewSpec(i)
        Next i
    End If
End Sub

Private Sub chkHideDiscontinued_Click()
    fillFullList Me.ItemNumber, Me.GeneralLC
End Sub

Private Sub CurrentAccessories_Click(i As Long, j As Long, txt As String)
    fillPreview Me.CurrentAccessories, i, j
End Sub

Private Sub CurrentAccessories_DblClick(i As Long, j As Long, txt As String)
    'If j <> 2 Then
        btnMoveLeft_Click
    'Else
    '    If txt = "Y" Then
    '        DB.Execute "UPDATE PartNumberAccessories SET DirectAccessory=0 WHERE ItemNumber='" & Me.ItemNumber & "' AND Accessory='" & Me.CurrentAccessories.GetTextSelected(0) & "'"
    '        Me.CurrentAccessories.Edit "N", i, j
    '    Else
    '        DB.Execute "UPDATE PartNumberAccessories SET DirectAccessory=1 WHERE ItemNumber='" & Me.ItemNumber & "' AND Accessory='" & Me.CurrentAccessories.GetTextSelected(0) & "'"
    '        Me.CurrentAccessories.Edit "Y", i, j
    '    End If
    'End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyF8 Or KeyCode = vbKeyPageUp) And Me.barPosition <> Me.barPosition.Min Then
        Me.barPosition = Me.barPosition - 1
        KeyCode = 0
    ElseIf (KeyCode = vbKeyF9 Or KeyCode = vbKeyPageDown) And Me.barPosition <> Me.barPosition.max Then
        Me.barPosition = Me.barPosition + 1
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
    Me.ItemPicture.SetCachePointer ImageCache
    Me.AccessoryPicture.SetCachePointer ImageCache
    
    Me.CurrentAccessories.sorted = True
    Me.FullList.sorted = True
    
    'Me.CurrentAccessories.SetColumnNames Array("Item", "Description", "Similar")
    'Me.CurrentAccessories.SetColumnWidths Array("1500", "4500", "500")
    'Me.CurrentAccessories.SetColumnNames Array("Item", "Description")
    'Me.CurrentAccessories.SetColumnWidths Array("1500", "5000")
    Me.CurrentAccessories.SetColumnNames Array("Item", "Sold", "Date Pub", "Description")
    Me.CurrentAccessories.SetColumnWidths Array("1500", "400", "400", "4200")
    Me.FullList.SetColumnNames Array("Item", "Description")
    Me.FullList.SetColumnWidths Array("1500", "5000")
    'this now gets called per-item in requeryForm
    'requeryPaths
    requeryForm
End Sub

Private Sub FullList_Click(i As Long, j As Long, txt As String)
    fillPreview Me.FullList, i, j
End Sub

Private Sub FullList_DblClick(i As Long, j As Long, txt As String)
    btnMoveRight_Click
End Sub

Private Sub ItemSelector_Click()
    If Me.ItemSelector <> Me.ItemNumber Then
        Me.barPosition.value = bsearch(Me.ItemSelector, ItemList)
    End If
End Sub

Private Sub opAccFilter_Click(index As Integer)
    If index = 2 And Not Me.FrameFBP.Visible Then
        btnExpandPathFilter_Click
    Else
        fillFullList Me.ItemNumber, Me.GeneralLC
    End If
End Sub

Private Sub opDirection_Click(index As Integer)
    fillCurrentAccessories Me.ItemNumber
    fillFullList Me.ItemNumber, Me.GeneralLC
End Sub

Private Sub fillCurrentAccessories(item As String)
    Mouse.Hourglass True
    Dim rst As ADODB.Recordset
    Select Case True
        Case Is = Me.opDirection(DIR_REGULAR_ACC)
            'Set rst = DB.retrieve("SELECT Accessory, AccessoryFullDescription, DirectAccessory FROM vAccessorySpecs WHERE ItemNumber='" & item & "'")
            'Set rst = DB.retrieve("SELECT Accessory, AccessoryFullDescription FROM vAccessorySpecs WHERE ItemNumber='" & item & "' AND DirectAccessory='N'")
            Set rst = DB.retrieve("SELECT Accessory, COALESCE(SoldStringTEMP,''), AccessoryPublishedDate, AccessoryFullDescription FROM vAccessorySpecs WHERE ItemNumber='" & item & "' AND DirectAccessory='N'")
        Case Is = Me.opDirection(DIR_REVERSE_ACC)
            'Set rst = DB.retrieve("SELECT ItemNumber, ItemFullDescription, DirectAccessory FROM vAccessorySpecs WHERE Accessory='" & item & "'")
            'Set rst = DB.retrieve("SELECT ItemNumber, ItemFullDescription FROM vAccessorySpecs WHERE Accessory='" & item & "' AND DirectAccessory='N'")
            Set rst = DB.retrieve("SELECT ItemNumber, COALESCE(SoldStringTEMP,''), AccessoryPublishedDate, ItemFullDescription FROM vAccessorySpecs WHERE Accessory='" & item & "' AND DirectAccessory='N'")
    End Select
    Me.CurrentAccessories.Clear
    If Not rst.EOF Then
        Me.CurrentAccessories.Add rst, , True
    End If
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub fillFullList(item As String, PL As String)
    Mouse.Hourglass True
    Dim sql As String
    sql = "SELECT PartNumbers.ItemNumber, " & _
                 "CASE WHEN PartNumbers.Desc1 IS NULL THEN '' WHEN PartNumbers.Desc1='' THEN '' ELSE PartNumbers.Desc1 + ' ' END + COALESCE(PartNumbers.Desc3,'') AS FullDescription " & _
          "FROM PartNumbers INNER JOIN vGeneralProductLine ON PartNumbers.ItemNumber=vGeneralProductLine.ItemNumber INNER JOIN vItemStatusBreakdown ON PartNumbers.ItemNumber=vItemStatusBreakdown.ItemNumber " & _
          "WHERE PartNumbers.BaseItemID=0 AND (PartNumbers.WebPublished=1 OR PartNumbers.WebToBePublished=1)"
          
    If Me.chkHideDiscontinued Then
        sql = sql & " AND vItemStatusBreakdown.DC=0"
    End If
    
    If CBool(Me.CurrentAccessories.ListCount <> 0) Then
        Select Case True
            Case Is = Me.opDirection(DIR_REGULAR_ACC)
                sql = sql & " AND PartNumbers.ItemNumber NOT IN (SELECT Accessory FROM vAccessorySpecs WHERE ItemNumber='" & item & "')"
            Case Is = Me.opDirection(DIR_REVERSE_ACC)
                sql = sql & " AND PartNumbers.ItemNumber NOT IN (SELECT ItemNumber FROM vAccessorySpecs WHERE Accessory='" & item & "')"
        End Select
    End If
    
    Select Case True
        Case Is = Me.opAccFilter(FLT_NONE)
            'noop
        Case Is = Me.opAccFilter(FLT_LINE)
            sql = sql & " AND GeneralProductLine='" & PL & "'"
        Case Is = Me.opAccFilter(FLT_PATH)
            'Dim i As Long
            'For i = 1 To 5
            '    If Me.PathFilter(i) <> "" Then
            '        Dim tmpname As String
            '        tmpname = Replace(Me.PathFilter(i), "'   ", "")
            '        sql = sql & " AND EXISTS (SELECT * FROM vPNWebPaths WHERE ItemNumber=PartNumbers.ItemNumber AND WebPathName"
            '        If InCombo(Me.PathFilter(i), Me.PathFilter(i), True) Then
            '            sql = sql & "='" & EscapeSQuotes(tmpname) & "'"
            '        Else
            '            sql = sql & " LIKE '%" & EscapeSQuotes(tmpname) & "%'"
            '        End If
            '        sql = sql & " AND PathLevel=" & i & ")"
            '    End If
            'Next i
    End Select
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve(sql)
    Me.FullList.Clear
    If Not rst.EOF Then
        Me.FullList.Add rst, , True
    End If
    rst.Close
    Set rst = Nothing
    Mouse.Hourglass False
End Sub

Private Sub fillPreview(ctl As SimpleListView, i As Long, j As Long)
    If Me.chkFillPreview Then
        Mouse.Hourglass True
        Dim thisitem As String
        thisitem = ctl.GetText(i, 0)
        If thisitem <> Me.PreviewItemNumber Then
            Mouse.Hourglass True
            Me.PreviewItemNumber = thisitem
            Me.PreviewDescription = ctl.GetText(i, 1)
            Dim rst As ADODB.Recordset
            Set rst = DB.retrieve("SELECT Spec1HTML, Spec2HTML, Spec3HTML, Spec4HTML FROM PartNumbers WHERE ItemNumber='" & thisitem & "'")
            For i = 1 To 4
                Me.PreviewSpec(i) = Nz(rst("Spec" & i & "HTML"))
            Next i
            rst.Close
            Set rst = Nothing
            Mouse.Hourglass False
        End If
        Me.AccessoryPicture.DisplayImage PictureDBFunctions.GenerateItemPicturePathAny(Me.ItemNumber)
        Mouse.Hourglass False
    End If
End Sub

Private Sub requeryPaths()
    'MsgBox "TODO: fixme!"
    'Dim rst As ADODB.Recordset
    'Dim hash As Dictionary
    'Set hash = New Dictionary
    'If Me.ItemNumber <> "" And Me.ItemNumber <> "ItemNumber" Then
    '    Set rst = DB.retrieve("SELECT WebPathName, PathLevel FROM vPNWebPaths WHERE ItemNumber='" & Me.ItemNumber & "'")
    '    While Not rst.EOF
    '        hash.Add rst("PathLevel") & "|" & rst("WebPathName"), "1"
    '        rst.MoveNext
    '    Wend
    '    rst.Close
    'End If
    'Set rst = DB.retrieve("SELECT WebPathName, PathLevel FROM WebPaths")
    'Dim i As Long
    'For i = 1 To 5
    '    Me.PathFilter(i).Clear
    'Next i
    'While Not rst.EOF
    '    If hash.exists(rst("PathLevel") & "|" & rst("WebPathName")) Then
    '        Me.PathFilter(rst("PathLevel")).AddItem "'   " & rst("WebPathName")
    '    Else
    '        Me.PathFilter(rst("PathLevel")).AddItem rst("WebPathName")
    '    End If
    '    rst.MoveNext
    'Wend
    'rst.Close
    'Set rst = Nothing
End Sub
