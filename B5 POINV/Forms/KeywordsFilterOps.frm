VERSION 5.00
Begin VB.Form KeywordsFilterOps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Keywords - Apply to Filter"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnExit 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1920
      TabIndex        =   3
      Top             =   2190
      Width           =   1965
   End
   Begin VB.CommandButton btnSetActive 
      Caption         =   "Set All to Active Status"
      Height          =   465
      Left            =   210
      TabIndex        =   2
      Top             =   1290
      Width           =   5235
   End
   Begin VB.CommandButton btnSetBidsNoOverwrite 
      Caption         =   "Set Suggested Bid to Current Bid -- Don't overwrite existing new bid"
      Height          =   465
      Left            =   210
      TabIndex        =   1
      Top             =   240
      Width           =   5235
   End
   Begin VB.CommandButton btnSetBidsAll 
      Caption         =   "Set Suggested Bid to Current Bid -- All"
      Height          =   465
      Left            =   210
      TabIndex        =   0
      Top             =   720
      Width           =   5235
   End
   Begin VB.Label lblStatusBar 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1830
      Width           =   3555
   End
End
Attribute VB_Name = "KeywordsFilterOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload KeywordsFilterOps
End Sub

Private Sub btnSetActive_Click()
    Dim i As Long, warn As Boolean
    ReDim PhraseList(KeywordsPhraseList.GetListLength) As String
    PhraseList = KeywordsPhraseList.GetList
    Dim svc As String
    svc = KeywordsPhraseList.GetCurrentService()
    Dim svcid As String
    svcid = KwdGetServiceIDFromName(svc)
    For i = 0 To UBound(PhraseList)
        Me.lblStatusBar.caption = "Working on " & PhraseList(i)
        DoEvents
        'if has everything, and not already active
        If KwdCheckGoodToGo(PhraseList(i), svc) = "OK" Then
            If Not DLookup("Active", "KeywordsAttributes", "SearchPhrase='" & EscapeSQuotes(PhraseList(i)) & "' AND Service='" & svc & "'") Then
                SetKwdStatusActive PhraseList(i), svcid
            End If
        Else
            warn = True
        End If
    Next i
    If warn Then
        MsgBox "Some keywords could not be set active (missing TL, etc.)"
    Else
        MsgBox "Done!"
    End If
    Me.lblStatusBar.caption = ""
    KeywordsPhraseList.RefreshSubform
    Unload KeywordsFilterOps
End Sub

Private Sub btnSetBidsAll_Click()
    bidLogic True
End Sub

Private Sub btnSetBidsNoOverwrite_Click()
    bidLogic False
End Sub

Private Sub bidLogic(Optional overwriteOK As Boolean = False)
    Dim i As Long
    ReDim PhraseList(KeywordsPhraseList.GetListLength) As String
    PhraseList = KeywordsPhraseList.GetList
    Dim rst As ADODB.Recordset
    Dim svc As String
    svc = KeywordsPhraseList.GetCurrentService()
    For i = 0 To UBound(PhraseList)
        Me.lblStatusBar.caption = "Working on " & PhraseList(i)
        DoEvents
        Set rst = DB.retrieve("SELECT SuggestedBid, NewBid, CurrentBid FROM KeywordsAttributes WHERE SearchPhrase='" & EscapeSQuotes(PhraseList(i)) & "' AND Service='" & svc & "'")
        'if valid bid
        If rst("SuggestedBid") <> 0 And rst("CurrentBid") <> rst("SuggestedBid") Then
            'if we're overwriting, or it doesn't have a bid
            If overwriteOK Or rst("NewBid") = 0 Then
                SetBid PhraseList(i), svc
            End If
        End If
        rst.Close
    Next i
    Set rst = Nothing
    MsgBox "Done!"
    Me.lblStatusBar.caption = ""
    KeywordsPhraseList.RefreshSubform
    Unload KeywordsFilterOps
End Sub
