VERSION 5.00
Begin VB.Form KeywordsEngineCostsMaintenance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keywords - Cost Maintenance"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6345
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbEditMethod 
      Height          =   315
      Left            =   5070
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   990
      Width           =   1215
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   375
      Left            =   5310
      TabIndex        =   11
      Top             =   30
      Width           =   1005
   End
   Begin VB.CommandButton btnImport 
      Caption         =   "Import from Spreadsheet"
      Height          =   465
      Left            =   5070
      TabIndex        =   9
      Top             =   4650
      Width           =   1125
   End
   Begin VB.CommandButton btnSelectNone 
      Caption         =   "Select None"
      Height          =   465
      Left            =   5070
      TabIndex        =   8
      Top             =   3420
      Width           =   1125
   End
   Begin VB.CommandButton btnSelectAll 
      Caption         =   "Select All"
      Height          =   465
      Left            =   5070
      TabIndex        =   7
      Top             =   2940
      Width           =   1125
   End
   Begin VB.CommandButton btnSelectInvert 
      Caption         =   "Invert Selection"
      Height          =   465
      Left            =   5070
      TabIndex        =   6
      Top             =   3900
      Width           =   1125
   End
   Begin VB.CommandButton btnEditSelected 
      Caption         =   "Edit Selected"
      Height          =   465
      Left            =   5070
      TabIndex        =   5
      Top             =   1500
      Width           =   1125
   End
   Begin VB.CommandButton btnEditAll 
      Caption         =   "Edit All"
      Height          =   465
      Left            =   5070
      TabIndex        =   4
      Top             =   1980
      Width           =   1125
   End
   Begin VB.ListBox costList 
      Height          =   4740
      Left            =   60
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   780
      Width           =   4845
   End
   Begin VB.TextBox engineid 
      Height          =   285
      Left            =   3780
      TabIndex        =   0
      Text            =   "engineid"
      Top             =   60
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblEditMethod 
      Caption         =   "Edit method:"
      Height          =   225
      Left            =   5100
      TabIndex        =   12
      Top             =   750
      Width           =   975
   End
   Begin VB.Label lblEngineName 
      Caption         =   "ENGINE_NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   180
      TabIndex        =   10
      Top             =   90
      Width           =   2625
   End
   Begin VB.Label headerLabel 
      Caption         =   "Cost Per Click"
      Height          =   225
      Index           =   1
      Left            =   3690
      TabIndex        =   3
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label headerLabel 
      Caption         =   "ItemNumber / Keyword"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   540
      Width           =   1695
   End
End
Attribute VB_Name = "KeywordsEngineCostsMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private currentMethod As String

Private currentsort As String
Private sortdescending As Boolean



Public Sub LoadEngine(engine As String)
    Me.engineid = DLookup("ID", "KeywordsEngines", "Engine='" & EscapeSQuotes(engine) & "'")
    Me.lblEngineName.Caption = engine
    requeryCostList
End Sub



Private Sub btnEditAll_Click()
    Dim factor As String
    factor = getFactor()
    If factor <> "" Then
        editArray ListBoxAsArray(Me.costList, False), cleanFactor(factor)
        requeryCostList
    End If
End Sub

Private Sub btnEditSelected_Click()
    Dim factor As String
    factor = getFactor()
    If factor <> "" Then
        editArray ListBoxAsArray(Me.costList, True), cleanFactor(factor)
        requeryCostList
    End If
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnImport_Click()
    Load KeywordsEngineCostsImport
    KeywordsEngineCostsImport.LoadEngine Me.lblEngineName
    KeywordsEngineCostsImport.Show MODAL
End Sub

Private Sub btnSelectAll_Click()
    Dim i As Long
    For i = 0 To Me.costList.ListCount - 1
        Me.costList.Selected(i) = True
    Next i
End Sub

Private Sub btnSelectInvert_Click()
    Dim i As Long
    For i = 0 To Me.costList.ListCount - 1
        Me.costList.Selected(i) = Not Me.costList.Selected(i)
    Next i
End Sub

Private Sub btnSelectNone_Click()
    Dim i As Long
    For i = 0 To Me.costList.ListCount - 1
        Me.costList.Selected(i) = False
    Next i
End Sub

Private Sub cmbEditMethod_Click()
    If Me.cmbEditMethod = "" Then
        Me.cmbEditMethod = currentMethod
    Else
        currentMethod = Me.cmbEditMethod
    End If
End Sub

Private Sub Form_Load()
    SetListTabs Me.costList, Array(180)
    requeryEditMethods
    
    currentsort = "SearchPhrase"
    sortdescending = False
End Sub

Private Sub requeryCostList()
    Me.costList.Clear
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT SearchPhrase, CurrentBid FROM KeywordsAttributes WHERE EngineID=" & Me.engineid & " ORDER BY " & currentsort & IIf(sortdescending, " DESC", ""))
    Me.costList.Visible = False
    While Not rst.EOF
        Me.costList.AddItem rst("SearchPhrase") & vbTab & Format(rst("CurrentBid"), "Currency")
        rst.MoveNext
    Wend
    Me.costList.Visible = True
End Sub

Private Sub requeryEditMethods()
    Me.cmbEditMethod.Clear
    Me.cmbEditMethod.AddItem "Absolute"
    Me.cmbEditMethod.AddItem "Relative"
    Me.cmbEditMethod.AddItem "Percent"
    Me.cmbEditMethod = Me.cmbEditMethod.List(0)
    currentMethod = Me.cmbEditMethod.List(0)
    ExpandDropDownToFit Me.cmbEditMethod
End Sub

Private Function getFactor() As String
    Dim factor As String
    Do
        factor = InputBox("Enter the value to change to/by:")
    Loop While Not validFactor(factor)
    getFactor = factor
End Function

Private Function validFactor(factor As String) As Boolean
    If factor = "" Then
        validFactor = True
    Else
        Select Case LCase(currentMethod)
            Case Is = "absolute"
                validFactor = IsNumeric(factor) And Not CBool(InStr(factor, "-"))
            Case Is = "relative"
                validFactor = IsNumeric(factor)
            Case Is = "percent"
                validFactor = IsNumeric(Replace(factor, "%", ""))
            Case Else
                MsgBox "whoever designed this program is an idiot. fix me."
        End Select
    End If
End Function

Private Function cleanFactor(factor As String) As String
    Select Case LCase(currentMethod)
        Case Is = "absolute"
            'no cleaning to do here
        Case Is = "relative"
            'add + sign if increase, otherwise ok
            If Not CBool(InStr(factor, "-")) Then
                factor = "+" & factor
            End If
        Case Is = "percent"
            'remove %, add + sign maybe
            factor = Replace(factor, "%", "")
            If Not CBool(InStr(factor, "-")) Then
                factor = "+" & factor
            End If
        Case Else
                MsgBox "whoever designed this program is an idiot. fix me."
    End Select
    cleanFactor = factor
End Function

Private Sub editArray(editlist() As String, factorValue As String)
    Dim i As Long
    Dim item As String, cpc As String
    Select Case LCase(currentMethod)
        Case Is = "absolute"
            DB.Execute "UPDATE KeywordsAttributes SET CurrentBid=" & factorValue & " WHERE EngineID=" & Me.engineid & " AND SearchPhrase IN (" & ListAsCommaSep(editlist, "'", True, True) & " )"
        Case Is = "relative"
            DB.Execute "UPDATE KeywordsAttributes SET CurrentBid=CurrentBid+" & factorValue & " WHERE EngineID=" & Me.engineid & " AND SearchPhrase IN (" & ListAsCommaSep(editlist, "'", True, True) & " )"
        Case Is = "percent"
            If Left(factorValue, 1) = "+" Then
                factorValue = Mid(factorValue, 2)
                factorValue = IIf(factorValue < 1, factorValue, factorValue / 100) + 1
            ElseIf Left(factorValue, 1) = "-" Then
                factorValue = Mid(factorValue, 2)
                factorValue = 1 - IIf(factorValue < 1, factorValue, factorValue / 100)
            Else
                MsgBox "Won't get here"
            End If
            'not doing it this way because we don't want the extra decimal places to carry over
            'during the calculations...is there a way to do this on the server?
            'dbExecute "UPDATE KeywordsSEParameters SET CostPerClick=CostPerClick*" & factorValue & " WHERE EngineID=" & Me.engineid & " AND ItemNumber IN (" & ListAsCommaSep(editlist, "'", True, True) & " )"
            For i = 0 To UBound(editlist)
                item = Left(editlist(i), InStr(editlist(i), vbTab) - 1)
                cpc = Mid(editlist(i), InStr(editlist(i), vbTab) + 1)
                cpc = Format(cpc * factorValue, "0.00")
                DB.Execute "UPDATE KeywordsAttributes SET CurrentBid=" & cpc & " WHERE EngineID=" & Me.engineid & " AND SearchPhrase='" & EscapeSQuotes(item) & "'"
            Next i
        Case Else
            MsgBox "whoever designed this program is an idiot. fix me."
    End Select
End Sub

Private Sub headerLabel_Click(Index As Integer)
    Dim newsort As String
    Select Case Index
        Case Is = 0
            newsort = "SearchPhrase"
        Case Is = 1
            newsort = "CurrentBid"
    End Select
    If newsort = currentsort Then
        sortdescending = Not sortdescending
    Else
        currentsort = newsort
        sortdescending = False
    End If
    requeryCostList
End Sub
