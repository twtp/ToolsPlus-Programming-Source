VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BetterRNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : BetterRNG
' DateTime  : 10/24/2005 16:58
' Author    : briandonorfio
' Purpose   : Marsaglia-esque shufflebox to Rnd()
'---------------------------------------------------------------------------------------

Option Explicit

Private Const BOX_START As Long = 1
Private Const BOX_END   As Long = 10000

Dim box(BOX_START To BOX_END) As Double

Private Sub Class_Initialize()
    Randomize Timer
    Dim i As Integer
    For i = BOX_START To BOX_END
        box(i) = Rnd
    Next i
End Sub

Public Function Between(min As Long, max As Long) As Long
    Dim i As Long, retval As Long
    i = Int((BOX_END - BOX_START + 1) * Rnd + BOX_START)
    retval = Int((max - min + 1) * box(i) + min)
    box(i) = Rnd
    Between = retval
End Function
