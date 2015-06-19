VERSION 5.00
Begin VB.UserControl PricingControl 
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   ScaleHeight     =   2835
   ScaleWidth      =   3015
   Begin VB.Frame PriceFrame 
      Height          =   1995
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.TextBox BreakQuantity 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   570
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox PriceBreak 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1980
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox BreakQuantity 
         Height          =   285
         Index           =   4
         Left            =   570
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1290
         Width           =   675
      End
      Begin VB.TextBox BreakQuantity 
         Height          =   285
         Index           =   5
         Left            =   570
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1620
         Width           =   675
      End
      Begin VB.TextBox BreakQuantity 
         Height          =   285
         Index           =   2
         Left            =   570
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   630
         Width           =   675
      End
      Begin VB.TextBox BreakQuantity 
         Height          =   285
         Index           =   3
         Left            =   570
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   960
         Width           =   675
      End
      Begin VB.TextBox PriceBreak 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1980
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   630
         Width           =   855
      End
      Begin VB.TextBox PriceBreak 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1980
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox PriceBreak 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1980
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1290
         Width           =   855
      End
      Begin VB.TextBox PriceBreak 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1980
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1620
         Width           =   855
      End
      Begin VB.CommandButton CopyOtherPriceStructureButton 
         Caption         =   "Copy OTHER"
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   30
         Width           =   1155
      End
      Begin VB.Label PriceLabel 
         Caption         =   "Price 5"
         Height          =   225
         Index           =   5
         Left            =   1410
         TabIndex        =   22
         Top             =   1650
         Width           =   525
      End
      Begin VB.Label PriceLabel 
         Caption         =   "Price 4"
         Height          =   225
         Index           =   4
         Left            =   1410
         TabIndex        =   21
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label PriceLabel 
         Caption         =   "Price 3"
         Height          =   225
         Index           =   3
         Left            =   1410
         TabIndex        =   20
         Top             =   990
         Width           =   525
      End
      Begin VB.Label PriceLabel 
         Caption         =   "Price 2"
         Height          =   225
         Index           =   2
         Left            =   1410
         TabIndex        =   19
         Top             =   660
         Width           =   525
      End
      Begin VB.Label PriceLabel 
         Caption         =   "Price 1"
         Height          =   225
         Index           =   1
         Left            =   1410
         TabIndex        =   18
         Top             =   330
         Width           =   525
      End
      Begin VB.Label QuantityLabel 
         Caption         =   "Qty 5"
         Height          =   225
         Index           =   5
         Left            =   90
         TabIndex        =   17
         Top             =   1650
         Width           =   465
      End
      Begin VB.Label QuantityLabel 
         Caption         =   "Qty 4"
         Height          =   225
         Index           =   4
         Left            =   90
         TabIndex        =   16
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label QuantityLabel 
         Caption         =   "Qty 3"
         Height          =   225
         Index           =   3
         Left            =   90
         TabIndex        =   15
         Top             =   990
         Width           =   465
      End
      Begin VB.Label QuantityLabel 
         Caption         =   "Qty 2"
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   14
         Top             =   660
         Width           =   465
      End
      Begin VB.Label QuantityLabel 
         Caption         =   "Qty 1"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   13
         Top             =   330
         Width           =   465
      End
      Begin VB.Label PriceTypeLabel 
         Caption         =   "TYPE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   1
         Top             =   0
         Width           =   1035
      End
   End
End
Attribute VB_Name = "PricingControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_priceType As String
Private m_priceTypeOther As String

Public Property Get PriceType() As String
    PriceType = m_priceType
End Property
Public Property Let PriceType(newPriceType As String)
    If CanPropertyChange("PriceType") Then
        m_priceType = newPriceType
        UserControl.PriceTypeLabel.Caption = newPriceType
        PropertyChanged "PriceType"
    End If
End Property

Public Property Get PriceTypeOther() As String
    PriceTypeOther = m_priceTypeOther
End Property
Public Property Let PriceTypeOther(newPriceTypeOther As String)
    If CanPropertyChange("PriceTypeOther") Then
        m_priceTypeOther = newPriceTypeOther
        UserControl.CopyOtherPriceStructureButton.Caption = "Copy " & newPriceTypeOther
        PropertyChanged "PriceTypeOther"
    End If
End Property


Public Sub Fill(rst As ADODB.Recordset)
    Dim i As Long
    For i = 1 To 5
        UserControl.PriceBreak(i) = Format(rst("DiscountMarkupPriceRate" & i), "Currency")
        UserControl.BreakQuantity(i).Enabled = True
    Next i
    UserControl.BreakQuantity(1).Enabled = False
    For i = 2 To 5
        If rst("BreakQty" & i - 1) = 999999 Then
            UserControl.BreakQuantity(i) = rst("BreakQty" & i - 1)
            Disable UserControl.PriceBreak(i)
            i = i + 1
            While i <= 5
                UserControl.BreakQuantity(i) = ""
                UserControl.BreakQuantity(i).Enabled = False
                Disable UserControl.PriceBreak(i)
                i = i + 1
            Wend
        Else
            UserControl.BreakQuantity(i) = rst("BreakQty" & i - 1) + 1
            UserControl.BreakQuantity(i - 1).Enabled = False
            Enable UserControl.PriceBreak(i)
        End If
    Next i
End Sub
