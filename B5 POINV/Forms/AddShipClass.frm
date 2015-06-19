VERSION 5.00
Begin VB.Form AddShipClass 
   Caption         =   "Add Shipping Classes"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnDelete 
      Caption         =   "Delete"
      Height          =   315
      Left            =   3750
      TabIndex        =   6
      Top             =   1620
      Width           =   705
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   1568
      TabIndex        =   5
      Top             =   3960
      Width           =   1545
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   2753
      TabIndex        =   4
      Top             =   3120
      Width           =   1005
   End
   Begin VB.TextBox newShipClass 
      Height          =   285
      Left            =   923
      TabIndex        =   3
      Top             =   3120
      Width           =   1785
   End
   Begin VB.ListBox lstShipClasses 
      Height          =   1620
      ItemData        =   "AddShipClass.frx":0000
      Left            =   660
      List            =   "AddShipClass.frx":0002
      TabIndex        =   2
      Top             =   1050
      Width           =   3015
   End
   Begin VB.Label generalLabel 
      Caption         =   "Current Shipping Classes"
      Height          =   285
      Index           =   0
      Left            =   660
      TabIndex        =   1
      Top             =   780
      Width           =   1845
   End
   Begin VB.Label generalLabel 
      Alignment       =   2  'Center
      Caption         =   "Shipping Classes"
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
      Index           =   1
      Left            =   1238
      TabIndex        =   0
      Top             =   60
      Width           =   2205
   End
End
Attribute VB_Name = "AddShipClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAdd_Click()
    If Me.newShipClass <> "" Then
        Dim rst As ADODB.Recordset
        Set rst = db.retrieve("SELECT ShipClassName FROM ShipClass WHERE ShipClassName='" & Me.newShipClass & "'")
        If rst.EOF Then
            db.Execute ("INSERT INTO ShipClass ( ShipClassName ) VALUES ( '" & Me.newShipClass & "' )")
        Else
            MsgBox "A shipping class with that name already exists."
        End If
        rst.Close
        Set rst = Nothing
        requeryShipClass
        Me.newShipClass = ""
    End If
End Sub

Private Sub btnDelete_Click()
    If MsgBox("Delete shipping class """ & Me.lstShipClasses & """?", vbYesNo) = vbYes Then
        db.Execute "DELETE FROM ShipClass WHERE ShipClassName='" & Me.lstShipClasses & "'"
        requeryShipClass
    End If
End Sub

Private Sub btnExit_Click()
    Unload AddShipClass
End Sub

Private Sub Form_Load()
    requeryShipClass
End Sub

Private Sub lstShipClasses_DblClick()
    btnDelete_Click
End Sub

Private Sub requeryShipClass()
    Dim rst As ADODB.Recordset
    Set rst = db.retrieve("SELECT ShipClassName FROM ShipClass ORDER BY ShipClassName")
    Me.lstShipClasses.Clear
    While Not rst.EOF
        Me.lstShipClasses.AddItem (rst("ShipClassName"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub newShipClass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        btnAdd_Click
    End If
End Sub
