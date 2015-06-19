VERSION 5.00
Begin VB.Form AddAccessories 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Accessories - Sign Maintenance"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PoInv.ItemFinder ItemFinder1 
      Height          =   3795
      Left            =   90
      TabIndex        =   8
      Top             =   660
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6641
   End
   Begin VB.CheckBox chkReverseAccessories 
      Caption         =   "hidden reverse accessories"
      Height          =   255
      Left            =   6990
      TabIndex        =   7
      Top             =   90
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   3570
      TabIndex        =   6
      Top             =   4530
      Width           =   1305
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   5100
      TabIndex        =   5
      Top             =   4530
      Width           =   1305
   End
   Begin VB.ListBox chosenAccessories 
      Height          =   3570
      Left            =   7830
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   810
      Width           =   1845
   End
   Begin VB.CommandButton btnMoveLeft 
      Caption         =   "<<"
      Height          =   465
      Left            =   7110
      TabIndex        =   2
      Top             =   3450
      Width           =   585
   End
   Begin VB.CommandButton btnMoveRight 
      Caption         =   ">>"
      Height          =   465
      Left            =   7110
      TabIndex        =   1
      Top             =   2940
      Width           =   585
   End
   Begin VB.Label lblChosenAcc 
      Caption         =   "Chosen Accessories:"
      Height          =   195
      Left            =   7800
      TabIndex        =   4
      Top             =   600
      Width           =   1605
   End
   Begin VB.Label lblTitle 
      Caption         =   "Add Accessories:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5805
   End
End
Attribute VB_Name = "AddAccessories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private changed As Boolean

Private Sub btnCancel_Click()
    Unload AddAccessories
End Sub

Private Sub btnMoveLeft_Click()
    changed = True
    Me.chosenAccessories.RemoveItem Me.chosenAccessories.ListIndex
End Sub

Private Sub btnMoveRight_Click()
    changed = True
    Dim item As String
    item = Me.ItemFinder1.GetSelectedItem
    If Not InList(item, Me.chosenAccessories, True) Then
        Me.chosenAccessories.AddItem item
    End If
End Sub

Private Sub btnOK_Click()
    'ok, good enough for me. doesn't check if it's been changed and then changed back though.
    If changed Then
        'cache the old list of accessories
        Dim rst As ADODB.Recordset
        If Me.chkReverseAccessories Then
            Set rst = DB.retrieve("SELECT ItemNumber AS Accessory FROM PartNumberAccessories WHERE Accessory='" & SignMaintenance.ItemNumber & "'")
        Else
            Set rst = DB.retrieve("SELECT Accessory FROM PartNumberAccessories WHERE ItemNumber='" & SignMaintenance.ItemNumber & "'")
        End If
        Dim lookup As Dictionary
        Set lookup = New Dictionary
        While Not rst.EOF
            lookup.Add rst("Accessory").value, "1"
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
        'add any new accessories
        Dim i As Long
        For i = 0 To Me.chosenAccessories.ListCount - 1
            If Not lookup.exists(Me.chosenAccessories.list(i)) Then
                If Me.chkReverseAccessories Then
                    DB.Execute "INSERT INTO PartNumberAccessories (ItemNumber, Accessory) VALUES ( '" & Me.chosenAccessories.list(i) & "', '" & SignMaintenance.ItemNumber & "' )"
                Else
                    DB.Execute "INSERT INTO PartNumberAccessories (ItemNumber, Accessory) VALUES ( '" & SignMaintenance.ItemNumber & "', '" & Me.chosenAccessories.list(i) & "' )"
                End If
            Else
                lookup.Remove CStr(Me.chosenAccessories.list(i))
            End If
        Next i
        'and delete the old ones
        Dim key As Variant
        For Each key In lookup.Keys
            If Me.chkReverseAccessories Then
                DB.Execute "DELETE FROM PartNumberAccessories WHERE ItemNumber='" & CStr(key) & "' AND Accessory='" & SignMaintenance.ItemNumber & "'"
            Else
                DB.Execute "DELETE FROM PartNumberAccessories WHERE ItemNumber='" & SignMaintenance.ItemNumber & "' AND Accessory='" & CStr(key) & "'"
            End If
        Next key
        Set lookup = Nothing
    End If
    Unload AddAccessories
End Sub

Private Sub chkReverseAccessories_Click()
    If Me.chkReverseAccessories Then
        Me.lblChosenAcc.caption = "Accessory of:"
        Me.lblTitle.caption = "Add " & SignMaintenance.ItemNumber & " as accessory to:"
    Else
        Me.lblChosenAcc.caption = "Chosen Accessories:"
        Me.lblTitle.caption = "Add Accessories to " & SignMaintenance.ItemNumber & ":"
    End If
    requeryChosenAccessories
End Sub

Private Sub chosenAccessories_DblClick()
    btnMoveLeft_Click
End Sub

Private Sub Form_Load()
    Mouse.Hourglass True
    Me.ItemFinder1.FillOnNoSel = True
    Me.ItemFinder1.Initialize
    requeryChosenAccessories
    Mouse.Hourglass False
End Sub





Private Sub requeryChosenAccessories()
    Dim rst As ADODB.Recordset
    If Me.chkReverseAccessories Then
        Set rst = DB.retrieve("SELECT ItemNumber AS Accessory FROM PartNumberAccessories WHERE Accessory='" & SignMaintenance.ItemNumber & "'")
    Else
        Set rst = DB.retrieve("SELECT Accessory FROM PartNumberAccessories WHERE ItemNumber='" & SignMaintenance.ItemNumber & "'")
    End If
    Me.chosenAccessories.Clear
    While Not rst.EOF
        Me.chosenAccessories.AddItem rst("Accessory")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub ItemFinder1_ListAction()
    btnMoveRight_Click
End Sub

'Private Sub saveAccessoriesReg()
'    'cache the old list of accessories
'    Dim rst As ADODB.Recordset
'    Set rst = DB.retrieve("SELECT Accessory FROM PartNumberAccessories WHERE ItemNumber='" & SignMaintenance.ItemNumber & "'")
'    Dim lookup As Dictionary
'    Set lookup = New Dictionary
'    While Not rst.EOF
'        lookup.Add rst("Accessory").value, "1"
'        rst.MoveNext
'    Wend
'    rst.Close
'    Set rst = Nothing
'    'add any new accessories
'    Dim i As Long
'    For i = 0 To Me.chosenAccessories.ListCount - 1
'        If Not lookup.exists(Me.chosenAccessories.list(i)) Then
'            DB.Execute "INSERT INTO PartNumberAccessories (ItemNumber, Accessory) VALUES ( '" & SignMaintenance.ItemNumber & "', '" & Me.chosenAccessories.list(i) & "' )"
'        Else
'            lookup.Remove CStr(Me.chosenAccessories.list(i))
'        End If
'    Next i
'    'and delete the old ones
'    Dim key As Variant
'    For Each key In lookup.Keys
'        DB.Execute "DELETE FROM PartNumberAccessories WHERE ItemNumber='" & SignMaintenance.ItemNumber & "' AND Accessory='" & CStr(key) & "'"
'    Next key
'    Set lookup = Nothing
'End Sub
