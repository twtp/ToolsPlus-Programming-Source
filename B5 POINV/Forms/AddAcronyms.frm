VERSION 5.00
Object = "{24892950-B55B-4EAE-811A-DD27F172DE83}#1.0#0"; "TPControls.ocx"
Begin VB.Form AddAcronyms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Acronyms"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6450
   StartUpPosition =   1  'CenterOwner
   Begin TPControls.SimpleListView AcronymList 
      Height          =   3675
      Left            =   150
      TabIndex        =   4
      Top             =   570
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6482
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Close"
      Height          =   345
      Left            =   5040
      TabIndex        =   3
      Top             =   0
      Width           =   1395
   End
   Begin VB.CommandButton DeleteThis 
      Caption         =   "Delete Selected"
      Height          =   345
      Left            =   1950
      TabIndex        =   2
      Top             =   4380
      Width           =   1425
   End
   Begin VB.CommandButton AddNew 
      Caption         =   "Add New"
      Height          =   345
      Left            =   360
      TabIndex        =   0
      Top             =   4380
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "Add/Edit Acronyms"
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
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   2655
   End
End
Attribute VB_Name = "AddAcronyms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AcronymList_DblClick(i As Long, j As Long, txt As String)
    Dim fieldName As String
    Dim newvalue As String
    Select Case j
        Case Is = 0
            fieldName = "Abbreviation"
            newvalue = InputBox("Enter new abbreviation:", , txt)
        Case Is = 1
            fieldName = "Explanation"
            newvalue = InputBox("Enter new explanation:", , txt)
        Case Else
            MsgBox "Whoa, didn't expect that. You clicked on a column I don't know about."
            Exit Sub
    End Select
    If newvalue = "" Then
        Exit Sub
    End If
    If newvalue = txt Then
        Exit Sub
    End If
    DB.Execute "UPDATE Abbreviations SET " & fieldName & "='" & EscapeSQuotes(newvalue) & "' WHERE Abbreviation='" & EscapeSQuotes(Me.AcronymList.GetTextSelected()) & "'"
    Me.AcronymList.Edit newvalue, i, j
End Sub

Private Sub AddNew_Click()
    Dim newabbrev As String
    Dim newexpl As String
    newabbrev = InputBox("Enter new abbreviation:")
    If newabbrev = "" Then
        Exit Sub
    End If
    If DLookup("COUNT(*)", "Abbreviations", "Abbreviation='" & EscapeSQuotes(newabbrev) & "'") <> 0 Then
        MsgBox "This abbreviation already exists in the lookup table."
        Exit Sub
    End If
    newexpl = InputBox("Enter explanation for " & qq(newabbrev) & ":")
    If newexpl = "" Then
        Exit Sub
    End If
    DB.Execute "INSERT INTO Abbreviations ( Abbreviation, Explanation ) VALUES ( '" & EscapeSQuotes(newabbrev) & "', '" & EscapeSQuotes(newexpl) & "' )"
    Me.AcronymList.Add Array(newabbrev, newexpl)
End Sub

Private Sub DeleteThis_Click()
    If Me.AcronymList.SelIndex <> -1 Then
        If vbYes = MsgBox("Remove abbreviation " & qq(Me.AcronymList.GetTextSelected()) & "?", vbYesNo) Then
            DB.Execute "DELETE FROM Abbreviations WHERE Abbreviation='" & EscapeSQuotes(Me.AcronymList.GetTextSelected()) & "'"
            Me.AcronymList.RemoveSelected
        End If
    End If
End Sub

Private Sub Exit_Click()
    Unload AddAcronyms
End Sub

Private Sub Form_Load()
    Me.AcronymList.sorted = True
    Me.AcronymList.SetColumnNames Array("Abbreviation", "Explanation")
    Me.AcronymList.SetColumnWidths Array(1800, 4000)
    requeryAcronyms
End Sub

Private Sub requeryAcronyms()
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT Abbreviation, Explanation FROM Abbreviations")
    Me.AcronymList.Add rst, , True
    rst.Close
    Set rst = Nothing
End Sub
