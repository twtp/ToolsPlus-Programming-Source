VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#3.5#0"; "TPControls.ocx"
Begin VB.Form BarcodePrinting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Barcode Printing"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4065
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox ReturnArea 
      Height          =   285
      Left            =   3030
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox PrinterName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   30
      Width           =   3105
   End
   Begin VB.CommandButton PrinterSelectButton 
      Caption         =   "Printer:"
      Height          =   315
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   825
   End
   Begin VB.CommandButton PrintButton 
      Caption         =   "Print Barcodes"
      Height          =   525
      Left            =   660
      TabIndex        =   5
      Top             =   2280
      Width           =   2715
   End
   Begin TPControls.NumericUpDown QtyToPrint 
      Height          =   285
      Left            =   1950
      TabIndex        =   4
      Top             =   1650
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   503
      Value           =   "1"
      Minimum         =   1
      Maximum         =   99999
      Increment       =   1
      DisplayFormat   =   "#0"
      AllowDecimals   =   0   'False
      TimeBeforeHoldDown=   500
      TimeBetweenRepeats=   75
      Enabled         =   -1  'True
      BackColor       =   -2147483643
      DefaultValue    =   0
   End
   Begin VB.TextBox CurrentQOH 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1950
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1260
      Width           =   1005
   End
   Begin VB.ComboBox ItemNumber 
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   690
      Width           =   2715
   End
   Begin VB.Label Label2 
      Caption         =   "Qty to Print:"
      Height          =   255
      Left            =   690
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Current QOH:"
      Height          =   255
      Left            =   690
      TabIndex        =   1
      Top             =   1290
      Width           =   1215
   End
End
Attribute VB_Name = "BarcodePrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private changed As Boolean
Private comboIndex As Dictionary

Private Sub Form_Load()
    requeryItems
End Sub

Private Sub ItemNumber_Click()
    changed = True
    ItemNumber_LostFocus
End Sub

Private Sub ItemNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.ItemNumber, KeyCode, Shift
    changed = True
End Sub

Private Sub ItemNumber_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.ItemNumber, KeyAscii
    changed = True
End Sub

Private Sub ItemNumber_LostFocus()
    AutoCompleteLostFocus Me.ItemNumber
    If changed Then
        If Me.ItemNumber = "" Then
            Me.CurrentQOH = ""
            Me.QtyToPrint.value = 0
        Else
            Dim cid As String
            cid = comboIndex.item(CLng(Me.ItemNumber.ListIndex))
            Dim rst As ADODB.Recordset
            Set rst = DB.retrieve("SELECT SUM(vActualWhseQty.QtyOnHand) FROM InventoryComponentMap INNER JOIN vActualWhseQty ON InventoryComponentMap.ItemID=vActualWhseQty.ItemID WHERE InventoryComponentMap.ComponentID=" & cid)
            If rst.EOF Then
                Me.CurrentQOH = "0"
            Else
                Me.CurrentQOH = Nz(rst(0), "0")
            End If
            'Me.QtyToPrint.value = Me.CurrentQOH
            If CLng(Me.CurrentQOH) <= 0 Then
                Me.QtyToPrint.value = 1
            Else
                Me.QtyToPrint.value = CLng(Me.CurrentQOH)
            End If
        End If
        changed = False
    End If
End Sub

Private Sub PrintButton_Click()
    If Me.PrinterName = "" Then
        MsgBox "No printer selected!"
    ElseIf Me.ItemNumber = "" Then
        MsgBox "No item selected!"
    ElseIf CLng(Me.QtyToPrint.value) < 1 Then
        MsgBox "Select a quantity to print!"
    Else
        Dim p As VB.printer
        Dim found As Boolean
        found = False
        For Each p In VB.printers
            If p.DeviceName = Me.PrinterName Then
                found = True
                Exit For
            End If
        Next p
        
        If found = False Then
            MsgBox "Couldn't find this printer name! How did that happen?"
        Else
            Dim cid As String
            cid = comboIndex.item(CLng(Me.ItemNumber.ListIndex))
            Dim bc As String
            bc = DLookup("Barcode", "InventoryComponentBarcodes", "ComponentID=" & cid)
            Dim epl2str As String
            epl2str = "EPL2" & vbCrLf & _
                      "q456" & vbCrLf & _
                      "Q152,24+0" & vbCrLf & _
                      "S4" & vbCrLf & _
                      "UN" & vbCrLf & _
                      "WN" & vbCrLf & _
                      "ZT" & vbCrLf & _
                      "N" & vbCrLf
            epl2str = epl2str & "B40,10,0,1,2,4,100,N," & qq(bc) & vbCrLf
            epl2str = epl2str & "A80,120,0,3,1,1,N," & qq(Me.ItemNumber) & vbCrLf
            epl2str = epl2str & "P" & CLng(Me.QtyToPrint.value) & vbCrLf
            PrintEPL2Data Me.PrinterName, epl2str, Me.ItemNumber & " barcodes (" & CLng(Me.QtyToPrint.value) & ")"
        End If
    End If
End Sub

Private Sub PrinterSelectButton_Click()
retry:
    Load PrinterSelectionDialog
    PrinterSelectionDialog.SetPrinterSelectionReturnLocation "BarcodePrinting", "ReturnArea"
    PrinterSelectionDialog.SetPrinterSelectionTo Me.PrinterName
    PrinterSelectionDialog.Show MODAL
    If Me.ReturnArea = "" Then
        'do nothing
    ElseIf Me.ReturnArea = Me.PrinterName Then
        'do nothing
    ElseIf CBool(InStr(LCase(Me.ReturnArea), "barcode")) Then
        Me.PrinterName = Me.ReturnArea
    Else
        If vbYes = MsgBox("Are you sure this is a Zebra thermal printer?", vbYesNo) Then
            Me.PrinterName = Me.ReturnArea
        Else
            'redo
            GoTo retry
        End If
    End If
End Sub

Private Sub requeryItems()
    Me.ItemNumber.Clear
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, Component FROM InventoryComponents WHERE ID IN (SELECT ComponentID FROM InventoryComponentBarcodes WHERE Internal=1) ORDER BY Component")
    Dim i As Long
    i = 0
    Set comboIndex = New Dictionary
    While Not rst.EOF
        comboIndex.Add i, CStr(rst("ID"))
        i = i + 1
        Me.ItemNumber.AddItem rst("Component")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub
