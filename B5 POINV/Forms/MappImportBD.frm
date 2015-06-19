VERSION 5.00
Begin VB.Form MappImportBD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "B&D MAPP Import"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   4725
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox MAPPOldCol 
      Height          =   285
      Left            =   1470
      TabIndex        =   35
      Text            =   "D"
      Top             =   3240
      Width           =   465
   End
   Begin VB.TextBox DateCol 
      Height          =   285
      Left            =   1470
      TabIndex        =   33
      Text            =   "G"
      Top             =   2940
      Width           =   465
   End
   Begin VB.CheckBox chkPrompt 
      Caption         =   "Prompt Before Changes"
      Height          =   255
      Left            =   90
      TabIndex        =   32
      Top             =   3990
      Width           =   2235
   End
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   2670
      TabIndex        =   29
      Top             =   2370
      Width           =   1815
      Begin VB.OptionButton opType 
         Caption         =   "Update PO/Inv"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Value           =   -1  'True
         Width           =   1605
      End
      Begin VB.OptionButton opType 
         Caption         =   "Report Only"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   180
         Width           =   1605
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   405
      Left            =   2550
      TabIndex        =   28
      Top             =   4350
      Width           =   1455
   End
   Begin VB.CommandButton btnImport 
      Caption         =   "Import"
      Height          =   405
      Left            =   750
      TabIndex        =   27
      Top             =   4350
      Width           =   1455
   End
   Begin VB.Frame framePC 
      Height          =   1035
      Left            =   3180
      TabIndex        =   21
      Top             =   1230
      Width           =   1485
      Begin VB.CheckBox EnablePC 
         Caption         =   "P-C"
         Height          =   225
         Left            =   120
         TabIndex        =   24
         Top             =   30
         Width           =   585
      End
      Begin VB.TextBox PCStart 
         Height          =   285
         Left            =   870
         TabIndex        =   23
         Top             =   330
         Width           =   555
      End
      Begin VB.TextBox PCEnd 
         Height          =   285
         Left            =   870
         TabIndex        =   22
         Top             =   630
         Width           =   555
      End
      Begin VB.Label generalLabel 
         Caption         =   "Start Line:"
         Height          =   255
         Index           =   8
         Left            =   90
         TabIndex        =   26
         Top             =   360
         Width           =   825
      End
      Begin VB.Label generalLabel 
         Caption         =   "End Line:"
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   25
         Top             =   660
         Width           =   825
      End
   End
   Begin VB.Frame frameDEL 
      Height          =   1035
      Left            =   1620
      TabIndex        =   15
      Top             =   1230
      Width           =   1485
      Begin VB.CheckBox EnableDEL 
         Caption         =   "DEL"
         Height          =   225
         Left            =   120
         TabIndex        =   18
         Top             =   30
         Width           =   645
      End
      Begin VB.TextBox DELStart 
         Height          =   285
         Left            =   870
         TabIndex        =   17
         Top             =   330
         Width           =   555
      End
      Begin VB.TextBox DELEnd 
         Height          =   285
         Left            =   870
         TabIndex        =   16
         Top             =   630
         Width           =   555
      End
      Begin VB.Label generalLabel 
         Caption         =   "Start Line:"
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   20
         Top             =   360
         Width           =   825
      End
      Begin VB.Label generalLabel 
         Caption         =   "End Line:"
         Height          =   255
         Index           =   6
         Left            =   90
         TabIndex        =   19
         Top             =   660
         Width           =   825
      End
   End
   Begin VB.Frame frameDW 
      Height          =   1035
      Left            =   60
      TabIndex        =   9
      Top             =   1230
      Width           =   1485
      Begin VB.TextBox DWEnd 
         Height          =   285
         Left            =   870
         TabIndex        =   14
         Top             =   630
         Width           =   555
      End
      Begin VB.TextBox DWStart 
         Height          =   285
         Left            =   870
         TabIndex        =   13
         Top             =   330
         Width           =   555
      End
      Begin VB.CheckBox EnableDW 
         Caption         =   "D-W"
         Height          =   225
         Left            =   120
         TabIndex        =   10
         Top             =   30
         Width           =   645
      End
      Begin VB.Label generalLabel 
         Caption         =   "End Line:"
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   12
         Top             =   660
         Width           =   825
      End
      Begin VB.Label generalLabel 
         Caption         =   "Start Line:"
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   11
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.CheckBox chkIncludeAccs 
      Caption         =   "Try Accessory Line Codes"
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   3690
      Value           =   1  'Checked
      Width           =   2235
   End
   Begin VB.TextBox PartNumCol 
      Height          =   285
      Left            =   1470
      TabIndex        =   6
      Text            =   "B"
      Top             =   2340
      Width           =   465
   End
   Begin VB.TextBox MAPPCol 
      Height          =   285
      Left            =   1470
      TabIndex        =   3
      Text            =   "F"
      Top             =   2640
      Width           =   465
   End
   Begin VB.TextBox Filename 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   810
      Width           =   2985
   End
   Begin VB.CommandButton btnBrowseForFile 
      Caption         =   "Browse..."
      Height          =   285
      Left            =   3480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   810
      Width           =   915
   End
   Begin VB.Label generalLabel 
      Caption         =   "Old MAPP Column:"
      Height          =   255
      Index           =   10
      Left            =   60
      TabIndex        =   36
      Top             =   3270
      Width           =   1365
   End
   Begin VB.Label generalLabel 
      Caption         =   "Date Column:"
      Height          =   255
      Index           =   9
      Left            =   60
      TabIndex        =   34
      Top             =   2970
      Width           =   1365
   End
   Begin VB.Label Label3 
      Caption         =   "Black && Decker MAPP Import"
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
      Left            =   120
      TabIndex        =   8
      Top             =   60
      Width           =   3825
   End
   Begin VB.Label generalLabel 
      Caption         =   "Part # Column:"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   2370
      Width           =   1365
   End
   Begin VB.Label generalLabel 
      Caption         =   "MAPP Column:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   2670
      Width           =   1365
   End
   Begin VB.Label generalLabel 
      Caption         =   "File:"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   840
      Width           =   375
   End
End
Attribute VB_Name = "MappImportBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fullreport As String

Private Sub btnBrowseForFile_Click()
    Dim dlg As FilePicker
    Set dlg = New FilePicker
    dlg.SetParent Me.hwnd
    dlg.AddFilter "Excel Documents", "*.xls"
    dlg.AddFilter "Comma-Separated Values Files", "*.csv"
    Me.Filename = dlg.ShowDialog
    Set dlg = Nothing
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnImport_Click()
    Mouse.Hourglass True
    fullreport = "Item" & vbTab & "New MAPP" & vbTab & "Old MAPP" & vbCrLf
    
    If sanityCheck Then
        doImport
    End If
    
    If Me.opType(0) Then
        Open DESTINATION_DIR & "mappreport.txt" For Output As #1
        Print #1, fullreport
        Close #1
        'OpenDefaultApp DESTINATION_DIR & "mappreport.txt"
        Dim excelapp As Excel.Application
        Set excelapp = New Excel.Application
        Dim wks As Excel.Worksheet
        Set wks = excelapp.Workbooks.Open(DESTINATION_DIR & "mappreport.txt").Worksheets(1)
        excelapp.Visible = True
        Set wks = Nothing
        Set excelapp = Nothing
    End If
    Mouse.Hourglass False
End Sub

Private Sub Form_Load()
    Disable Me.DWStart
    Disable Me.DWEnd
    Disable Me.DELStart
    Disable Me.DELEnd
    Disable Me.PCStart
    Disable Me.PCEnd
End Sub

Private Sub MAPPCol_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(LimitToLetters(KeyAscii))))
End Sub

Private Sub opType_Click(Index As Integer)
    If Me.opType(1) Then
        Me.chkPrompt.Enabled = True
    Else
        Me.chkPrompt.Enabled = False
    End If
End Sub

Private Sub PartNumCol_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(LimitToLetters(KeyAscii))))
End Sub

Private Sub DELEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
End Sub

Private Sub DELStart_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
End Sub

Private Sub DWEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
End Sub

Private Sub DWStart_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
End Sub

Private Sub PCEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
End Sub

Private Sub PCStart_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
End Sub

Private Sub EnableDEL_Click()
    If Me.EnableDEL Then
        Enable Me.DELStart
        Enable Me.DELEnd
    Else
        Disable Me.DELStart
        Disable Me.DELEnd
    End If
End Sub

Private Sub EnableDW_Click()
    If Me.EnableDW Then
        Enable Me.DWStart
        Enable Me.DWEnd
    Else
        Disable Me.DWStart
        Disable Me.DWEnd
    End If
End Sub

Private Sub EnablePC_Click()
    If Me.EnablePC Then
        Enable Me.PCStart
        Enable Me.PCEnd
    Else
        Disable Me.PCStart
        Disable Me.PCEnd
    End If
End Sub

Private Function sanityCheck() As Boolean
    sanityCheck = False
    If Me.Filename = "" Then
        MsgBox "No file entered!"
    ElseIf Not FileExists(Me.Filename) Then
        MsgBox "Can't find file " & qq(Me.Filename) & "!"
    ElseIf Me.EnableDW And (Me.DWStart = "" Or Me.DWEnd = "") Then
        MsgBox "D-W not filled out!"
    ElseIf Me.EnableDEL And (Me.DELStart = "" Or Me.DELEnd = "") Then
        MsgBox "DEL not filled out!"
    ElseIf Me.EnablePC And (Me.PCStart = "" Or Me.PCEnd = "") Then
        MsgBox "P-C not filled out!"
    ElseIf Me.PartNumCol = "" Then
        MsgBox "Partnumber column not specified!"
    ElseIf Me.MAPPCol = "" Then
        MsgBox "MAPP column not specified!"
    Else
        sanityCheck = True
    End If
End Function

Private Sub doImport()
    Dim excelapp As Excel.Application
    Set excelapp = New Excel.Application
    excelapp.Workbooks.Open Me.Filename
    
    If Me.EnableDW Then
        Call doLineCode("D-W", IIf(Me.chkIncludeAccs, "D-A", ""), Me.DWStart, Me.DWEnd, excelapp)
    End If
    If Me.EnableDEL Then
        Call doLineCode("DEL", IIf(Me.chkIncludeAccs, "DEA", ""), Me.DELStart, Me.DELEnd, excelapp)
    End If
    If Me.EnablePC Then
        Call doLineCode("P-C", IIf(Me.chkIncludeAccs, "P-A", ""), Me.PCStart, Me.PCEnd, excelapp)
    End If
    excelapp.Quit
    Set excelapp = Nothing
End Sub

Private Sub doLineCode(lc As String, altlc As String, fromLine As Long, toLine As Long, xl As Excel.Application)
    Dim itemH As Dictionary
    Set itemH = New Dictionary
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber, MAPP FROM InventoryMaster WHERE MAPP<>0 AND (ProductLine='" & lc & "'" & IIf(altlc = "", "", " OR ProductLine='" & altlc & "'") & ")")
    While Not rst.EOF
        itemH.Add CStr(rst("ItemNumber")), CCur(rst("MAPP"))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    Dim i As Long
    For i = fromLine To toLine
        Dim item As String
        item = xl.ActiveWorkbook.ActiveSheet.Range(Me.PartNumCol & i).value
        item = Trim(item)
        If item <> "" Then
            Dim mappdate As String
            mappdate = xl.ActiveWorkbook.ActiveSheet.Range(Me.DateCol & i).value
            mappdate = Replace(mappdate, "-", "-01-")
            mappdate = Format(mappdate, "Short Date")
            While Left(item, 1) = "0"
                item = Mid(item, 2)
            Wend
            item = lc & item
            If ExistsInInventoryMaster(item) Then
                If DateDiff("d", mappdate, Date) >= 0 Then
                    doLine item, itemH, xl.ActiveWorkbook.ActiveSheet.Range(Me.MAPPCol & i).value
                Else
                    doLine item, itemH, xl.ActiveWorkbook.ActiveSheet.Range(Me.MAPPOldCol & i).value
                End If
            ElseIf Me.chkIncludeAccs Then
                item = altlc & Mid(item, 4)
                If ExistsInInventoryMaster(item) Then
                    If DateDiff("d", mappdate, Date) >= 0 Then
                        doLine item, itemH, xl.ActiveWorkbook.ActiveSheet.Range(Me.MAPPCol & i).value
                    Else
                        doLine item, itemH, xl.ActiveWorkbook.ActiveSheet.Range(Me.MAPPOldCol & i).value
                    End If
                End If
            End If
        End If
    Next i
    
    Dim thisitem As Variant
    For Each thisitem In itemH.Keys
        If Me.opType(0) Then
            fullreport = fullreport & CStr(thisitem) & vbTab & "N/A" & vbTab & itemH.item(thisitem) & vbCrLf
        Else
            If Me.chkPrompt Then
                If vbYes = MsgBox(thisitem & " not found in xsheet, remove MAPP?", vbYesNo) Then
                    DB.Execute "UPDATE InventoryMaster SET MAPP=0 WHERE ItemNumber='" & thisitem & "'"
                End If
            Else
                DB.Execute "UPDATE InventoryMaster SET MAPP=0 WHERE ItemNumber='" & thisitem & "'"
            End If
        End If
    Next thisitem
    
    Set itemH = Nothing
End Sub

Private Sub doLine(item As String, itemH As Dictionary, newmapp As String)
    If itemH.exists(item) Then
        If itemH.item(item) = CCur(newmapp) Then
            'no change, do nothing
        Else
            'mapp changed
            If Me.opType(0) Then
                fullreport = fullreport & item & vbTab & Format(newmapp, "Currency") & vbTab & Format(itemH.item(item), "Currency") & vbCrLf
            Else
                If Me.chkPrompt Then
                    If vbYes = MsgBox(item & ": Update MAPP to " & Format(newmapp, "Currency") & " (was " & Format(itemH.item(item), "Currency") & ")?", vbYesNo) Then
                        DB.Execute "UPDATE InventoryMaster SET MAPP=" & newmapp & ", DateMAPPEffective=GETDATE() WHERE ItemNumber='" & item & "'"
                    End If
                Else
                    DB.Execute "UPDATE InventoryMaster SET MAPP=" & newmapp & ", DateMAPPEffective=GETDATE() WHERE ItemNumber='" & item & "'"
                End If
            End If
        End If
        itemH.Remove item
    Else
        'new mapp
        If Me.opType(0) Then
            fullreport = fullreport & item & vbTab & Format(newmapp, "Currency") & vbTab & "N/A" & vbCrLf
        Else
            If Me.chkPrompt Then
                If vbYes = MsgBox(item & ": Update MAPP to " & Format(newmapp, "Currency") & " (new)?", vbYesNo) Then
                    DB.Execute "UPDATE InventoryMaster SET MAPP=" & newmapp & ", DateMAPPEffective=GETDATE() WHERE ItemNumber='" & item & "'"
                End If
            Else
                DB.Execute "UPDATE InventoryMaster SET MAPP=" & newmapp & ", DateMAPPEffective=GETDATE() WHERE ItemNumber='" & item & "'"
            End If
        End If
    End If
End Sub
