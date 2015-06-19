VERSION 5.00
Begin VB.Form GenericMessageBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pick something!"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton optionButton 
      Caption         =   "placeholder"
      Height          =   405
      Index           =   0
      Left            =   2460
      TabIndex        =   0
      Top             =   1740
      Width           =   1395
   End
   Begin VB.Label messageText 
      Caption         =   "placeholder"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4845
   End
End
Attribute VB_Name = "GenericMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub CustomizeForm(widthTwips As Long, heightTwips As Long, theMessage As String, buttonTexts As Variant, Optional textAlignment As AlignmentConstants = vbCenter)

    Dim numButtons As Long
    If VarType(buttonTexts) = vbString Then
        numButtons = 1
    ElseIf VarType(buttonTexts) And vbArray Then
        numButtons = UBound(buttonTexts) + 1
    Else
        Err.Raise ErrorCodes.GenericMBArgumentTypeMismatch, "CustomizeForm", "buttonTexts must be a string or variant array"
    End If
    
    'sanity check for width, must be wider than buttons area is
    '(button width * #buttons) + (spacing * #spaces) + side space*2
    If widthTwips < (Me.optionButton(0).width * numButtons) + (300 * (numButtons - 1)) + 600 Then
        widthTwips = (Me.optionButton(0).width * numButtons) + (300 * (numButtons - 1)) + 600
    End If
    
    Me.width = widthTwips
    Me.Height = heightTwips
    
    If VarType(buttonTexts) = vbString Then
        Me.optionButton(0).caption = buttonTexts
        Me.optionButton(0).Left = placeButton(1, Me.optionButton(0).width, 1, Me.width)
        Me.optionButton(0).Top = Me.Height - 930
    ElseIf VarType(buttonTexts) And vbArray Then
        Me.optionButton(0).caption = buttonTexts(0)
        Me.optionButton(0).Left = placeButton(1, Me.optionButton(0).width, (UBound(buttonTexts) + 1), Me.width)
        Me.optionButton(0).Top = Me.Height - 930
        
        Dim i As Long
        For i = 1 To UBound(buttonTexts)
            Load Me.optionButton(i)
            Me.optionButton(i).caption = buttonTexts(i)
            Me.optionButton(i).Left = placeButton(i + 1, Me.optionButton(0).width, (UBound(buttonTexts) + 1), Me.width)
            Me.optionButton(i).Top = Me.Height - 930
            Me.optionButton(i).Visible = True
        Next i
    End If

    Me.messageText = theMessage
    Me.messageText.width = Me.width - 240
    Me.messageText.Left = 120
    Me.messageText.Height = Me.Height - 1020
    Me.messageText.Height = 240
    Me.messageText.Alignment = textAlignment

End Sub

Private Sub Form_Load()
    'everything is in CustomizeForm()
End Sub

Private Function placeButton(buttonNumber As Long, buttonWidth As Long, numButtons As Long, formWidth As Long) As Long
    If numButtons = 1 Then
        placeButton = (formWidth - buttonWidth) / 2
    Else
        Dim barwidth As Long, firstButton As Long
        barwidth = ((buttonWidth * numButtons) + (300 * (numButtons - 1)))
        firstButton = (formWidth - barwidth) / 2
        placeButton = firstButton + ((buttonNumber - 1) * (buttonWidth + 300))
    End If
End Function

Private Sub optionButton_Click(Index As Integer)
    GenericMessageBoxFns.BasicMessageBoxReturnValue = Index
    Unload Me
End Sub
