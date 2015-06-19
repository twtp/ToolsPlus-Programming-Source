VERSION 5.00
Begin VB.Form KeywordsEngineMaintenance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keywords - Edit Services/Engines"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   510
      TabIndex        =   9
      Top             =   2970
      Width           =   3315
      Begin VB.CommandButton btnHolidayTo 
         Caption         =   "To:"
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   690
         Width           =   525
      End
      Begin VB.CommandButton btnHolidayFrom 
         Caption         =   "From:"
         Height          =   255
         Left            =   60
         TabIndex        =   15
         Top             =   360
         Width           =   525
      End
      Begin VB.TextBox holidayPercent 
         Height          =   285
         Left            =   2220
         TabIndex        =   14
         Top             =   570
         Width           =   825
      End
      Begin VB.TextBox holidayTo 
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   690
         Width           =   1095
      End
      Begin VB.TextBox holidayFrom 
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkHolidayIncrease 
         Caption         =   "Holiday Increase"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   30
         Width           =   1485
      End
      Begin VB.Label Label4 
         Caption         =   "% increase:"
         Height          =   225
         Left            =   2220
         TabIndex        =   13
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   375
      Left            =   1650
      TabIndex        =   8
      Top             =   4530
      Width           =   1275
   End
   Begin VB.CommandButton btnNewEngine 
      Caption         =   "Create New"
      Height          =   285
      Left            =   3270
      TabIndex        =   6
      Top             =   360
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   510
      TabIndex        =   5
      Top             =   1860
      Width           =   3315
      Begin VB.CommandButton btnEditCosts 
         Caption         =   "Edit / Import Costs"
         Height          =   705
         Left            =   120
         TabIndex        =   7
         Top             =   210
         Width           =   3105
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   510
      TabIndex        =   2
      Top             =   780
      Width           =   3315
      Begin VB.OptionButton opEngineType 
         Caption         =   "Shopping Engine (products only)"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   4
         Top             =   570
         Width           =   2955
      End
      Begin VB.OptionButton opEngineType 
         Caption         =   "Ads Service (products and keywords)"
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   210
         Width           =   2985
      End
   End
   Begin VB.ComboBox cmbService 
      Height          =   315
      Left            =   1470
      TabIndex        =   0
      Top             =   330
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "Select Service:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "KeywordsEngineMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fillingform As Boolean
Dim currentService As String

Public Sub LoadEngine(engineName As String)
    Me.cmbService = engineName
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ProductsOnly FROM KeywordsEngines WHERE Engine='" & EscapeSQuotes(engineName) & "'")
    fillingform = True
    Me.opEngineType(SQLBool(rst("ProductsOnly"))) = True
    fillingform = False
    'Me.btnEditCosts.Enabled = CBool(rst("CostManual"))
End Sub




Private Sub btnEditCosts_Click()
    Load KeywordsEngineCostsMaintenance
    KeywordsEngineCostsMaintenance.LoadEngine currentService
    KeywordsEngineCostsMaintenance.Show MODAL
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnHolidayFrom_Click()
    Load GenericCalendarPopUp
    GenericCalendarPopUp.SetUpCalendar Me.holidayFrom, "KeywordsEngineMaintenance", "holidayFrom"
    GenericCalendarPopUp.Show MODAL
End Sub

Private Sub btnHolidayTo_Click()
    Load GenericCalendarPopUp
    GenericCalendarPopUp.SetUpCalendar Me.holidayTo, "KeywordsEngineMaintenance", "holidayTo"
    GenericCalendarPopUp.Show MODAL
End Sub

Private Sub btnNewEngine_Click()
    Dim newname As String
    newname = InputBox("Enter name for new engine:")
    If newname <> "" Then
        If 0 = DLookup("COUNT(*)", "KeywordsEngines", "Engine='" & EscapeSQuotes(newname) & "'") Then
            DB.Execute "INSERT INTO KeywordsEngines ( Engine, ProductsOnly ) VALUES ( '" & EscapeSQuotes(newname) & "', 0 )"
        Else
            MsgBox "Already exists!"
        End If
        If currentService <> newname Then
            LoadEngine newname
        End If
    End If
End Sub

Private Sub chkHolidayIncrease_Click()
    Me.holidayFrom.Enabled = CBool(Me.chkHolidayIncrease)
    Me.holidayTo.Enabled = CBool(Me.chkHolidayIncrease)
    Me.holidayPercent.Enabled = CBool(Me.chkHolidayIncrease)
    Me.btnHolidayFrom.Enabled = CBool(Me.chkHolidayIncrease)
    Me.btnHolidayTo.Enabled = CBool(Me.chkHolidayIncrease)
    If Me.chkHolidayIncrease Then
        Me.holidayFrom = Format(Date, "mm/dd/yyyy")
        Me.holidayTo = Format(Date, "mm/dd/yyyy")
        Me.holidayPercent = "0"
    Else
        Me.holidayFrom = ""
        Me.holidayTo = ""
        Me.holidayPercent = ""
    End If
End Sub

Private Sub cmbService_Click()
    cmbService_LostFocus
End Sub

Private Sub cmbService_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoCompleteKeyDown Me.cmbService, KeyCode, Shift
End Sub

Private Sub cmbService_KeyPress(KeyAscii As Integer)
    AutoCompleteKeyPress Me.cmbService, KeyAscii
End Sub

Private Sub cmbService_LostFocus()
    AutoCompleteLostFocus Me.cmbService
    If Me.cmbService = "" Then
        Me.cmbService = currentService
    ElseIf Me.cmbService <> currentService Then
        currentService = Me.cmbService
        LoadEngine Me.cmbService
    End If
End Sub

Private Sub Form_Load()
    requeryServices
End Sub


Private Sub requeryServices()
    Me.cmbService.Clear
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ID, Engine FROM KeywordsEngines ORDER BY ID")
    While Not rst.EOF
        Me.cmbService.AddItem rst("Engine")
        rst.MoveNext
    Wend
    ExpandDropDownToFit Me.cmbService
    Me.cmbService = Me.cmbService.list(0) 'this breaks if we don't have any services...
    currentService = Me.cmbService.list(0)
End Sub

Private Sub holidayFrom_Change()
    DB.Execute "UPDATE KeywordsEngines SET HolidayFrom='" & Me.holidayFrom & "' WHERE Engine='" & currentService & "'"
End Sub

Private Sub holidayPercent_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitToNumbers(KeyAscii)
End Sub

Private Sub holidayPercent_LostFocus()
    DB.Execute "UPDATE KeywordsEngines SET HolidayPercent=" & Me.holidayPercent & " WHERE Engine='" & currentService & "'"
End Sub

Private Sub holidayTo_Change()
    DB.Execute "UPDATE KeywordsEngines SET HolidayTo='" & Me.holidayTo & "' WHERE Engine='" & currentService & "'"
End Sub

Private Sub opEngineType_Click(Index As Integer)
    If Not fillingform Then
        DB.Execute "UPDATE KeywordsEngines SET ProductsOnly=" & IIf(Index = 0, "0", "1")
    End If
End Sub
