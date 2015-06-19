VERSION 5.00
Begin VB.Form InventoryKitPopup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kit Details"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox ItemListBox 
      Height          =   3180
      Left            =   90
      TabIndex        =   1
      Top             =   900
      Width           =   3525
   End
   Begin VB.Label TitleLabel 
      Caption         =   "%ITEM% Kit Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   3645
   End
End
Attribute VB_Name = "InventoryKitPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SetListTabs Me.ItemListBox, Array(24)
End Sub

Public Sub LoadKit(item As String)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ComponentItemCode, QuantityPerAssembly FROM IM_SalesKitDetail WHERE SalesKitNo='" & item & "' ORDER BY CoreComponent DESC, ComponentItemCode ASC")
    LoadItem item, rst
    rst.Close
    Set rst = Nothing
End Sub

Public Sub LoadComponent(item As String)
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT SalesKitNo, QuantityPerAssembly FROM IM_SalesKitDetail WHERE ComponentItemCode='" & item & "' ORDER BY CoreComponent DESC, SalesKitNo ASC")
    LoadItem item, rst
    rst.Close
    Set rst = Nothing
End Sub

Private Sub LoadItem(item As String, rst As ADODB.Recordset)
    Me.TitleLabel.Caption = Replace(Me.TitleLabel.Caption, "%ITEM%", item)
    Me.ItemListBox.AddItem item
    Me.ItemListBox.AddItem ""
    While Not rst.EOF
        Me.ItemListBox.AddItem rst(1) & vbTab & rst(0)
        rst.MoveNext
    Wend
End Sub

Private Sub ItemListBox_DblClick()
    If Me.ItemListBox = "" Then
        'skip
    Else
        Dim temp As Variant
        temp = Split(Me.ItemListBox, vbTab)
        Dim item As String
        item = temp(IIf(UBound(temp) = 1, 1, 0))
        InventoryMaintenance.GoToItem item, True, False
    End If
End Sub
