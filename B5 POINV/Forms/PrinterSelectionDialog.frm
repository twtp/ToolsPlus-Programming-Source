VERSION 5.00
Begin VB.Form PrinterSelectionDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Printer"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2550
      TabIndex        =   3
      Top             =   2880
      Width           =   1485
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   435
      Left            =   930
      TabIndex        =   2
      Top             =   2880
      Width           =   1485
   End
   Begin VB.ListBox PrinterList 
      Height          =   2010
      Left            =   300
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Select Printer"
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
      TabIndex        =   0
      Top             =   90
      Width           =   1935
   End
End
Attribute VB_Name = "PrinterSelectionDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private formToUpdate As String
Private controlToUpdate As String

Public Sub SetPrinterSelectionReturnLocation(formName As String, controlName As String)
    formToUpdate = formName
    controlToUpdate = controlName
End Sub

Public Sub SetPrinterSelectionTo(PrinterName As String)
    If InList(PrinterName, Me.PrinterList, True) Then
        Me.PrinterList = PrinterName
        Me.OKButton.Enabled = True
    End If
End Sub

Private Sub CancelButton_Click()
    SetControlValue formToUpdate, controlToUpdate, ""
    Unload Me
End Sub

Private Sub Form_Load()
    formToUpdate = ""
    controlToUpdate = ""
    Dim p As VB.printer
    For Each p In VB.printers
        Me.PrinterList.AddItem p.DeviceName
    Next p
End Sub

Private Sub OKButton_Click()
    If formToUpdate <> "" And controlToUpdate <> "" Then
        SetControlValue formToUpdate, controlToUpdate, Me.PrinterList
    Else
        MsgBox "Selected a printer, but not sending the value anywhere! Tell Brian about this."
    End If
    Unload Me
End Sub

Private Sub PrinterList_Click()
    Me.OKButton.Enabled = True
End Sub
