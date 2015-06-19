VERSION 5.00
Begin VB.Form AddBaseItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Edit Base Item"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   9015
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton mappingDown 
      Caption         =   "\/"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3750
      TabIndex        =   18
      Top             =   3060
      Width           =   285
   End
   Begin VB.CommandButton mappingUp 
      Caption         =   "/\"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3750
      TabIndex        =   17
      Top             =   2760
      Width           =   285
   End
   Begin VB.CommandButton Exit 
      Caption         =   "OK"
      Height          =   495
      Left            =   1620
      TabIndex        =   16
      Top             =   4230
      Width           =   1365
   End
   Begin VB.CommandButton lengthIncrement 
      Caption         =   "len ++"
      Enabled         =   0   'False
      Height          =   225
      Left            =   3780
      TabIndex        =   15
      Top             =   2010
      Width           =   615
   End
   Begin VB.CommandButton lengthDecrement 
      Caption         =   "len --"
      Enabled         =   0   'False
      Height          =   225
      Left            =   3780
      TabIndex        =   14
      Top             =   1770
      Width           =   615
   End
   Begin VB.CommandButton positionDecrement 
      Caption         =   "pos --"
      Enabled         =   0   'False
      Height          =   225
      Left            =   3780
      TabIndex        =   13
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton positionIncrement 
      Caption         =   "pos ++"
      Enabled         =   0   'False
      Height          =   225
      Left            =   3780
      TabIndex        =   12
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton MoveRightAll 
      Caption         =   "=> ALL"
      Height          =   495
      Left            =   6480
      TabIndex        =   11
      Top             =   1020
      Width           =   585
   End
   Begin VB.TextBox BaseItemID 
      Height          =   285
      Left            =   2550
      TabIndex        =   10
      Text            =   "ID"
      Top             =   180
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.ListBox AttributeMappings 
      Height          =   1425
      Left            =   270
      TabIndex        =   9
      Top             =   2430
      Width           =   3405
   End
   Begin VB.CommandButton RemoveAttribute 
      Caption         =   "Remove This Attribute"
      Enabled         =   0   'False
      Height          =   465
      Left            =   1740
      TabIndex        =   8
      Top             =   690
      Width           =   1185
   End
   Begin VB.CommandButton AddAttribute 
      Caption         =   "Add New Attribute"
      Height          =   465
      Left            =   330
      TabIndex        =   7
      Top             =   690
      Width           =   1275
   End
   Begin VB.ListBox Attributes 
      Height          =   1035
      Left            =   270
      TabIndex        =   6
      Top             =   1170
      Width           =   3435
   End
   Begin VB.TextBox BaseItem 
      Height          =   285
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   180
      Width           =   1515
   End
   Begin VB.CommandButton MoveLeft 
      Caption         =   "<="
      Height          =   435
      Left            =   6480
      TabIndex        =   3
      Top             =   2340
      Width           =   585
   End
   Begin VB.CommandButton MoveRight 
      Caption         =   "=>"
      Height          =   435
      Left            =   6480
      TabIndex        =   2
      Top             =   1890
      Width           =   585
   End
   Begin VB.ListBox SelectedItems 
      Height          =   4155
      Left            =   7080
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   30
      Width           =   1905
   End
   Begin VB.ListBox AllItems 
      Height          =   4155
      Left            =   4530
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "Base Item:"
      Height          =   225
      Left            =   90
      TabIndex        =   5
      Top             =   210
      Width           =   855
   End
End
Attribute VB_Name = "AddBaseItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'me.baseitem or me.baseitemid need to be loaded w/ something
    'check *_change events
    ReDim tabs(2) As Long
    tabs(0) = 72
    tabs(1) = tabs(0) + 36
    tabs(2) = tabs(1) + 300
    SetListTabStops Me.Attributes.hwnd, tabs
    ReDim tabs(0) As Long
    tabs(0) = 72
    SetListTabStops Me.AttributeMappings.hwnd, tabs
End Sub

Private Sub BaseItem_Change()
    If Me.BaseItemID = "ID" Or Me.BaseItemID = "" Then
        Me.BaseItemID = DLookup("ID", "BaseItemList", "BaseItem='" & Me.BaseItem & "'")
        requeryAllItems
        requerySelectedItems
        requeryAttributes
    End If
End Sub

Private Sub BaseItemID_Change()
    If Me.BaseItem = "" Then
        Me.BaseItem = DLookup("BaseItem", "BaseItemList", "ID=" & Me.BaseItemID)
        requeryAllItems
        requerySelectedItems
        requeryAttributes
    End If
End Sub

Private Sub requeryAllItems()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber FROM PartNumbers WHERE ItemNumber LIKE '" & Me.BaseItem & "%' ORDER BY ItemNumber")
    Me.AllItems.Clear
    While Not rst.EOF
        If rst("ItemNumber") <> Me.BaseItem Then
            Me.AllItems.AddItem rst("ItemNumber")
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requerySelectedItems()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber FROM PartNumbers WHERE BaseItemID='" & Me.BaseItemID & "'")
    Me.SelectedItems.Clear
    While Not rst.EOF
        Me.SelectedItems.AddItem rst("ItemNumber")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryAttributes()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, AttributeName, StartPosition, AttributeLength FROM BaseItemAttributes WHERE BaseItemID=" & Me.BaseItemID)
    Me.Attributes.Clear
    Me.Attributes.AddItem "Attribute Name" & vbTab & "Position" & vbTab & "Length"
    While Not rst.EOF
        Me.Attributes.AddItem rst("AttributeName") & vbTab & rst("StartPosition") & vbTab & rst("AttributeLength") & vbTab & rst("ID")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
End Sub

Private Sub requeryAttributeMappings()
    Me.AttributeMappings.Clear
    Me.AttributeMappings.AddItem "Abbreviation" & vbTab & "Readable"
    If Me.Attributes.ListIndex <> 0 Then
        Dim rst As ADODB.Recordset
        Dim pieces As Variant
        pieces = Split(Me.Attributes, vbTab)
        Set rst = DB.retrieve("EXEC spBaseItemAttrMapRequery " & Me.BaseItemID & ", " & pieces(1) & ", " & pieces(2))
        While Not rst.EOF
            Me.AttributeMappings.AddItem rst("Abbreviation") & vbTab & IIf(Nz(rst("Readable")) = "", rst("Abbreviation"), Nz(rst("Readable")))
            rst.MoveNext
        Wend
        rst.Close
        Set rst = Nothing
    End If
End Sub

Private Sub Exit_Click()
    Unload AddBaseItem
End Sub





Private Sub AddAttribute_Click()
    Dim currentAtts As Dictionary
    Set currentAtts = New Dictionary
    Dim i As Long
    For i = 1 To Me.Attributes.ListCount - 1
        currentAtts.Add Left(Me.Attributes.list(i), InStr(Me.Attributes.list(i), vbTab) - 1), "1"
    Next i

    Dim attName As String
    Dim startPos As String
    Dim attLen As String
    Do
        attName = InputBox("Enter attribute name [color, waist size, etc], capitalization doesn't matter:")
        If attName = "" Then
            Exit Sub
        End If
        If currentAtts.exists(attName) Then
            Exit Sub
        End If
    Loop Until Len(attName) < 15
    Set currentAtts = Nothing
    Do
        startPos = InputBox("Enter the starting position of the attribute:" & vbCrLf & vbCrLf & "Ex. FOO123-BLK = 8")
        If startPos = "" Then
            Exit Sub
        End If
    Loop Until IsNumeric(startPos) And startPos < 16
    Do
        attLen = InputBox("Enter the length of the attribute:" & vbCrLf & vbCrLf & "Ex. FOO123-BLK = 3")
        If attLen = "" Then
            Exit Sub
        End If
    Loop Until IsNumeric(attLen) And attLen < 15
    
    Dim newid As String
    DB.Execute "INSERT INTO BaseItemAttributes ( BaseItemID, AttributeName, StartPosition, AttributeLength ) VALUES ( " & Me.BaseItemID & ", '" & EscapeSQuotes(attName) & "', " & startPos & ", " & attLen & " )"
    newid = DLookup("@@IDENTITY", "BaseItemAttributes")
    
    Me.Attributes.AddItem attName & vbTab & startPos & vbTab & attLen & vbTab & newid
End Sub

Private Sub RemoveAttribute_Click()
    If MsgBox("Remove attribute " & Left(Me.Attributes, InStr(Me.Attributes, vbTab) - 1) & "?", vbYesNo) = vbYes Then
        Dim oldid As String
        oldid = Mid(Me.Attributes, InStrRev(Me.Attributes, vbTab) + 1)
        DB.Execute "DELETE FROM BaseItemAttributeMappings WHERE AttributeID=" & oldid
        DB.Execute "DELETE FROM BaseItemAttributes WHERE ID=" & oldid
        Me.Attributes.RemoveItem Me.Attributes.ListIndex
    End If
End Sub

Private Sub Attributes_Click()
    If Me.Attributes.ListIndex <> 0 Then
        Me.RemoveAttribute.Enabled = True
        Me.positionDecrement.Enabled = True
        Me.positionIncrement.Enabled = True
        Me.lengthDecrement.Enabled = True
        Me.lengthIncrement.Enabled = True
    Else
        Me.RemoveAttribute.Enabled = False
        Me.positionDecrement.Enabled = False
        Me.positionIncrement.Enabled = False
        Me.lengthDecrement.Enabled = False
        Me.lengthIncrement.Enabled = False
    End If
    requeryAttributeMappings
End Sub

Private Sub AttributeMappings_Click()
    Dim tabpos As Long
    tabpos = InStr(Me.AttributeMappings, vbTab)
    If Left(Me.AttributeMappings, tabpos - 1) = Mid(Me.AttributeMappings, tabpos + 1) Then
        Me.mappingDown.Enabled = False
        Me.mappingUp.Enabled = False
    Else
        Select Case Me.AttributeMappings.ListIndex
            Case Is = 0
                Me.mappingDown.Enabled = False
                Me.mappingUp.Enabled = False
            Case Is = 1
                Me.mappingDown.Enabled = True
                Me.mappingUp.Enabled = CBool(Me.AttributeMappings.ListCount > 2)
            Case Is = Me.AttributeMappings.ListCount - 1
                Me.mappingDown.Enabled = False
                Me.mappingUp.Enabled = True
            Case Else
                Me.mappingDown.Enabled = True
                Me.mappingUp.Enabled = True
        End Select
    End If
End Sub




Private Sub AttributeMappings_DblClick()
    If Me.AttributeMappings.ListIndex <> 0 Then
        Dim attrid As String
        attrid = Mid(Me.Attributes, InStrRev(Me.Attributes, vbTab) + 1)
        Dim abbrev As String, oldreadable As String, newreadable As String
        abbrev = Left(Me.AttributeMappings, InStr(Me.AttributeMappings, vbTab) - 1)
        oldreadable = Mid(Me.AttributeMappings, InStr(Me.AttributeMappings, vbTab) + 1)
        newreadable = InputBox("What should " & abbrev & " show up as? [currently " & oldreadable & "] Capitalization is important!", , oldreadable)
        If newreadable = "" Then
            'cancel
        ElseIf newreadable = oldreadable Then
            'do nothing
        ElseIf newreadable = abbrev Then
            DB.Execute "DELETE FROM BaseItemAttributeMappings WHERE AttributeID=" & attrid & " AND Abbreviation=" & abbrev
            Me.AttributeMappings.RemoveItem Me.AttributeMappings.ListIndex
        ElseIf abbrev = oldreadable Then
            DB.Execute "INSERT INTO BaseItemAttributeMappings ( AttributeID, Abbreviation, Readable ) VALUES ( " & attrid & ", '" & abbrev & "', '" & EscapeSQuotes(newreadable) & "' )"
            Me.AttributeMappings.list(Me.AttributeMappings.ListIndex) = abbrev & vbTab & newreadable
        Else
            DB.Execute "UPDATE BaseItemAttributeMappings SET Readable='" & EscapeSQuotes(newreadable) & "' WHERE AttributeID=" & attrid & " AND Abbreviation='" & abbrev & "'"
            Me.AttributeMappings.list(Me.AttributeMappings.ListIndex) = abbrev & vbTab & newreadable
        End If
    End If
End Sub






Private Sub MoveLeft_Click()
    DB.Execute "UPDATE PartNumbers SET BaseItemID=0 WHERE ItemNumber='" & Me.SelectedItems & "'"
    Me.SelectedItems.RemoveItem Me.SelectedItems.ListIndex
End Sub

Private Sub MoveRight_Click()
    If Not InList(Me.AllItems, Me.SelectedItems, True) Then
        Me.SelectedItems.AddItem Me.AllItems
        DB.Execute "UPDATE PartNumbers SET BaseItemID='" & Me.BaseItemID & "' WHERE ItemNumber='" & Me.AllItems & "'"
    End If
End Sub

Private Sub MoveRightAll_Click()
    Me.SelectedItems.Clear
    Dim i As Long
    For i = 0 To Me.AllItems.ListCount - 1
        Me.SelectedItems.AddItem Me.AllItems.list(i)
    Next i
    DB.Execute "UPDATE PartNumbers SET BaseItemID='" & Me.BaseItemID & "' WHERE ItemNumber LIKE '" & Me.BaseItem & "%' AND ItemNumber<>'" & Me.BaseItem & "'"
End Sub

Private Sub AllItems_DblClick()
    MoveRight_Click
End Sub

Private Sub SelectedItems_DblClick()
    MoveLeft_Click
End Sub



Private Sub positionDecrement_Click()
    Dim pieces As Variant
    pieces = Split(Me.Attributes, vbTab)
    If pieces(1) >= 2 Then
        pieces(1) = pieces(1) - 1
        DB.Execute "UPDATE BaseItemAttributes SET StartPosition=" & pieces(1) & " WHERE ID=" & pieces(3)
        Me.Attributes.list(Me.Attributes.ListIndex) = Join(pieces, vbTab)
    End If
    requeryAttributeMappings
End Sub

Private Sub positionIncrement_Click()
    Dim pieces As Variant
    pieces = Split(Me.Attributes, vbTab)
    If pieces(1) <= 14 Then
        pieces(1) = pieces(1) + 1
        DB.Execute "UPDATE BaseItemAttributes SET StartPosition=" & pieces(1) & " WHERE ID=" & pieces(3)
        Me.Attributes.list(Me.Attributes.ListIndex) = Join(pieces, vbTab)
    End If
    requeryAttributeMappings
End Sub

Private Sub lengthDecrement_Click()
    Dim pieces As Variant
    pieces = Split(Me.Attributes, vbTab)
    If pieces(2) >= 2 Then
        pieces(2) = pieces(2) - 1
        DB.Execute "UPDATE BaseItemAttributes SET AttributeLength=" & pieces(2) & " WHERE ID=" & pieces(3)
        Me.Attributes.list(Me.Attributes.ListIndex) = Join(pieces, vbTab)
    End If
    requeryAttributeMappings
End Sub

Private Sub lengthIncrement_Click()
    Dim pieces As Variant
    pieces = Split(Me.Attributes, vbTab)
    If pieces(2) <= 14 Then
        pieces(2) = pieces(2) + 1
        DB.Execute "UPDATE BaseItemAttributes SET AttributeLength=" & pieces(2) & " WHERE ID=" & pieces(3)
        Me.Attributes.list(Me.Attributes.ListIndex) = Join(pieces, vbTab)
    End If
    requeryAttributeMappings
End Sub



Private Sub mappingDown_Click()
    Dim attrid As String
    attrid = Mid(Me.Attributes, InStrRev(Me.Attributes, vbTab) + 1)
    'TODO
    MsgBox "TODO"
End Sub

Private Sub mappingUp_Click()
    Dim attrid As String
    attrid = Mid(Me.Attributes, InStrRev(Me.Attributes, vbTab) + 1)
    'TODO
    MsgBox "TODO"
End Sub
