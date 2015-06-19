VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#3.5#0"; "TPControls.ocx"
Begin VB.Form InventoryQuantityTriggers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory Quantity Triggers"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4365
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnPrintScreen 
      Caption         =   "Print Screen"
      Height          =   525
      Left            =   3660
      TabIndex        =   18
      Top             =   30
      Width           =   705
   End
   Begin VB.TextBox CurrentTriggerType 
      Height          =   285
      Left            =   2940
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   300
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Frame qtyFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   765
      Left            =   450
      TabIndex        =   12
      Top             =   1500
      Width           =   3255
      Begin VB.TextBox Quantity 
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Top             =   420
         Width           =   855
      End
      Begin VB.ComboBox AboveOrBelow 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   60
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Quantity:"
         Height          =   255
         Left            =   90
         TabIndex        =   16
         Top             =   450
         Width           =   645
      End
      Begin VB.Label lblCurrQty 
         Caption         =   "CURRENT QTY"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   450
         Width           =   1275
      End
   End
   Begin VB.Frame dateFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   465
      Left            =   90
      TabIndex        =   9
      Top             =   1980
      Width           =   3105
      Begin TPControls.DateDropdown DateDropdown 
         Height          =   315
         Left            =   840
         TabIndex        =   10
         Top             =   60
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
      End
      Begin VB.Label Label4 
         Caption         =   "Date:"
         Height          =   255
         Left            =   90
         TabIndex        =   11
         Top             =   90
         Width           =   675
      End
   End
   Begin VB.ComboBox TriggerType 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   630
      Width           =   2655
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   465
      Left            =   2910
      TabIndex        =   6
      Top             =   3300
      Width           =   975
   End
   Begin VB.CommandButton btnRemove 
      Caption         =   "Remove"
      Height          =   465
      Left            =   1770
      TabIndex        =   5
      Top             =   3300
      Width           =   975
   End
   Begin VB.CommandButton btnSet 
      Caption         =   "Set"
      Height          =   465
      Left            =   630
      TabIndex        =   4
      Top             =   3300
      Width           =   975
   End
   Begin VB.TextBox Notes 
      Height          =   885
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2340
      Width           =   3735
   End
   Begin VB.TextBox ItemNumber 
      Height          =   285
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1170
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Type of Trigger:"
      Height          =   255
      Left            =   150
      TabIndex        =   8
      Top             =   660
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "Item:"
      Height          =   285
      Left            =   630
      TabIndex        =   1
      Top             =   1200
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "Inventory Quantity Triggers"
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
      TabIndex        =   0
      Top             =   60
      Width           =   3495
   End
End
Attribute VB_Name = "InventoryQuantityTriggers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private currentTriggers As Dictionary
Private changed As Boolean
Private fillingForm As Boolean

Private Sub AboveOrBelow_Click()
    If fillingForm Then
        Exit Sub
    End If
    
    changed = True
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnPrintScreen_Click()
    ClipBoard_SetImage Utilities.CaptureVB6Form(Me)
    MsgBox "Screenshot captured!"
End Sub

Private Sub btnRemove_Click()
    deleteTrigger
    Unload Me
End Sub

Private Sub btnSet_Click()
    If saveTrigger Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    changed = False
    Me.TriggerType.AddItem "Quantity"
    Me.TriggerType.AddItem "Date"
    
    Me.AboveOrBelow.AddItem "Below"
    Me.AboveOrBelow.AddItem "Above"
    
    Me.ItemNumber = InventoryMaintenance.ItemNumber
    'Me.lblCurrQty.caption = "Currently: " & InventoryMaintenance.QtyOnHand
    'Me.DateDropdown.value = Date
    
    Me.qtyFrame.Top = 1500
    Me.dateFrame.Top = 1500
    Me.qtyFrame.Left = 450
    Me.dateFrame.Left = 450
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT TriggerType, QtyThreshold, Below, DateThreshold, Notes FROM InventoryQuantityTriggers WHERE ItemNumber='" & Me.ItemNumber & "'")
    Set currentTriggers = New Dictionary
    While Not rst.EOF
        Select Case rst("TriggerType")
            Case Is = "qty"
                currentTriggers.Add "qty", CStr(rst("QtyThreshold") & Chr(0) & IIf(rst("Below"), "Below", "Above") & Chr(0) & rst("Notes"))
            Case Is = "date"
                currentTriggers.Add "date", CStr(rst("DateThreshold") & Chr(0) & rst("Notes"))
        End Select
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    If currentTriggers.exists("qty") Then
        fillTrigger "qty"
    ElseIf currentTriggers.exists("date") Then
        fillTrigger "date"
    Else
        'fillingform = True
        'Me.CurrentTriggerType = "qty"
        'Me.TriggerType = "Quantity"
        'Me.AboveOrBelow = "Below"
        'fillingform = False
        'setToDefaults
        fillingForm = True
        Me.TriggerType = "Quantity"
        fillingForm = False
        setToDefaults
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set currentTriggers = Nothing
End Sub

Private Sub Notes_KeyDown(KeyCode As Integer, Shift As Integer)
    changed = True
End Sub

Private Sub Quantity_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
    If KeyAscii <> 0 Then
        changed = True
    End If
End Sub

Private Sub fillTrigger(which As String)
    If currentTriggers.exists(which) Then
        Dim fields As Variant
        fields = Split(currentTriggers.item(which), Chr(0))
        fillingForm = True
        Select Case which
            Case Is = "qty"
                Me.TriggerType = "Quantity"
                Me.Quantity = fields(0)
                Me.AboveOrBelow = fields(1)
            Case Is = "date"
                Me.TriggerType = "Date"
                Me.DateDropdown.value = fields(0)
        End Select
        fillingForm = False
        Me.Notes = fields(UBound(fields))
    Else
        setToDefaults
    End If
    makeVisible which
End Sub

Private Sub makeVisible(which As String)
    Me.qtyFrame.Visible = False
    Me.dateFrame.Visible = False
    Select Case which
        Case Is = "qty"
            Me.qtyFrame.Visible = True
        Case Is = "date"
            Me.dateFrame.Visible = True
    End Select
    Me.CurrentTriggerType = which
End Sub

Private Sub setToDefaults()
    fillingForm = True
    'Me.TriggerType = "Quantity"
    makeVisible "qty"
    Me.ItemNumber = InventoryMaintenance.ItemNumber
    Me.AboveOrBelow = IIf(InventoryMaintenance.QtyOnHand = 0, "Above", "Below")
    Me.Quantity = ""
    Me.lblCurrQty.Caption = "Currently: " & InventoryMaintenance.QtyOnHand
    Me.DateDropdown.value = Date
    Me.Notes = ""
    'changed = False
    fillingForm = False
End Sub

Private Function getTriggerType() As String
    'Select Case Me.TriggerType
    '    Case Is = "Quantity"
    '        getTriggerType = "qty"
    '    Case Is = "Date"
    '        getTriggerType = "date"
    'End Select
    getTriggerType = Me.CurrentTriggerType
End Function

Private Function triggerExists(item As String, whichType As String) As Boolean
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT COUNT(*) FROM InventoryQuantityTriggers WHERE ItemNumber='" & item & "' AND TriggerType='" & whichType & "'")
    triggerExists = CBool(rst(0) = 1)
    rst.Close
    Set rst = Nothing
End Function

Private Sub TriggerType_Click()
    If fillingForm Then
        Exit Sub
    End If
    
    If changed Then
        If vbYes = MsgBox("Save changes to this trigger?", vbYesNo) Then
            If Not saveTrigger() Then
                Exit Sub
            End If
        End If
    End If
    
    Select Case Me.TriggerType
        Case Is = "Quantity"
            fillTrigger "qty"
        Case Is = "Date"
            fillTrigger "date"
    End Select
End Sub

Private Function saveTrigger() As Boolean
    Select Case Me.CurrentTriggerType
        Case Is = "qty"
            If Me.Quantity = "" Then
                MsgBox "Enter a quantity!"
                saveTrigger = False
            Else
                deleteTrigger
                DB.Execute "INSERT INTO InventoryQuantityTriggers ( ItemNumber, TriggerType, QtyThreshold, Below, Notes ) VALUES ( '" & Me.ItemNumber & "', 'qty', " & Me.Quantity & ", " & IIf(Me.AboveOrBelow = "Below", "1", "0") & ", '" & EscapeSQuotes(Me.Notes) & "' )"
                changed = False
                saveTrigger = True
            End If
        Case Is = "date"
            If DateDiff("d", Me.DateDropdown.value, Date) > 0 Then
                MsgBox "Date can't occur in the past!"
                saveTrigger = False
            Else
                deleteTrigger
                DB.Execute "INSERT INTO InventoryQuantityTriggers ( ItemNumber, TriggerType, DateThreshold, Notes ) VALUES ( '" & Me.ItemNumber & "', 'date', '" & Me.DateDropdown.value & "', '" & EscapeSQuotes(Me.Notes) & "' )"
                changed = False
                saveTrigger = True
            End If
        Case Else
            MsgBox "can't save trigger because " & qq(Me.CurrentTriggerType) & " isn't in the lookup!"
    End Select
End Function

Private Sub deleteTrigger()
    DB.Execute "DELETE FROM InventoryQuantityTriggers WHERE ItemNumber='" & Me.ItemNumber & "' AND TriggerType='" & getTriggerType() & "'"
    changed = False
End Sub
