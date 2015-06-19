VERSION 5.00
Begin VB.Form AddItemWebPaths 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Item Web Paths"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox hiddenWebsiteID 
      Height          =   285
      Left            =   2910
      TabIndex        =   19
      Text            =   "WSID"
      Top             =   210
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton btnDeletePath 
      Caption         =   "Del Path"
      Height          =   525
      Left            =   60
      TabIndex        =   17
      Top             =   1500
      Width           =   555
   End
   Begin VB.CommandButton btnAddPath 
      Caption         =   "Add Path"
      Height          =   525
      Left            =   8010
      TabIndex        =   16
      Top             =   1500
      Width           =   555
   End
   Begin VB.OptionButton opPath1Filter 
      Caption         =   "Manufacturers"
      Height          =   255
      Index           =   2
      Left            =   6660
      TabIndex        =   15
      Top             =   660
      Width           =   1395
   End
   Begin VB.OptionButton opPath1Filter 
      Caption         =   "Categories"
      Height          =   255
      Index           =   1
      Left            =   6660
      TabIndex        =   14
      Top             =   360
      Value           =   -1  'True
      Width           =   1395
   End
   Begin VB.OptionButton opPath1Filter 
      Caption         =   "All"
      Height          =   255
      Index           =   0
      Left            =   6660
      TabIndex        =   13
      Top             =   60
      Width           =   1395
   End
   Begin VB.ListBox BrowserDetails 
      Height          =   2010
      Left            =   4650
      TabIndex        =   11
      Top             =   4050
      Width           =   3255
   End
   Begin VB.ListBox Browser 
      Height          =   2790
      Left            =   4680
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   960
      Width           =   3255
   End
   Begin VB.ListBox PathsToRemove 
      Height          =   1620
      Left            =   660
      TabIndex        =   8
      Top             =   4800
      Width           =   3795
   End
   Begin VB.ListBox PathsToAdd 
      Height          =   1620
      Left            =   660
      TabIndex        =   6
      Top             =   2910
      Width           =   3795
   End
   Begin VB.ListBox PathsCurrent 
      Height          =   1620
      Left            =   660
      TabIndex        =   4
      Top             =   960
      Width           =   3795
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   4290
      TabIndex        =   2
      Top             =   6600
      Width           =   1305
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   2580
      TabIndex        =   1
      Top             =   6600
      Width           =   1305
   End
   Begin VB.Label lblPreviousPaths 
      Height          =   615
      Left            =   3600
      TabIndex        =   18
      Top             =   30
      Width           =   2895
   End
   Begin VB.Label generalLabel 
      Caption         =   "Path Details:"
      Height          =   225
      Index           =   4
      Left            =   4770
      TabIndex        =   12
      Top             =   3810
      Width           =   1725
   End
   Begin VB.Label generalLabel 
      Caption         =   "Current Paths:"
      Height          =   225
      Index           =   3
      Left            =   4800
      TabIndex        =   10
      Top             =   720
      Width           =   1725
   End
   Begin VB.Label generalLabel 
      Caption         =   "To Remove:"
      Height          =   225
      Index           =   2
      Left            =   720
      TabIndex        =   7
      Top             =   4560
      Width           =   1725
   End
   Begin VB.Label generalLabel 
      Caption         =   "To Add:"
      Height          =   225
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Top             =   2670
      Width           =   1725
   End
   Begin VB.Label generalLabel 
      Caption         =   "Current Paths:"
      Height          =   225
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Width           =   1725
   End
   Begin VB.Label Label1 
      Caption         =   "Add Item Web Paths:"
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
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   2745
   End
End
Attribute VB_Name = "AddItemWebPaths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private browserHistory As PerlArray
Private browserClicked As String

Private hasSections As Boolean
Private hasItems As Boolean

Private idsToAdd As PerlArray
Private idsToDel As PerlArray

Private Sub btnAddPath_Click()
    If Me.Browser.ListIndex <> -1 Then
        Dim wpid As String
        wpid = Mid(Me.Browser, InStr(Me.Browser, vbTab) + 1)
        If idsToAdd.Find(wpid) <> -1 Then
            Exit Sub
        End If
        If "0" <> DLookup("COUNT(*)", "PartNumberWebPaths", "ItemNumber='" & SignMaintenance.ItemNumber & "' AND WebPathID=" & wpid) Then
            Exit Sub
        End If
        If hasSections Then
            If vbNo = MsgBox("This section has subsections! Are you sure you want to add here?", vbYesNo) Then
                Exit Sub
            End If
        End If
        
        'ok to add
        Dim i As Long
        For i = 1 To browserHistory.Scalar - 1 'skip first -1
            Me.PathsToAdd.AddItem Space(2 * (i - 1)) & WebPathCache_GetName(browserHistory.Elem(i))
        Next i
        Me.PathsToAdd.AddItem "*" & Space(2 * (i - 1)) & Me.Browser
        idsToAdd.Push wpid
    End If
End Sub

Private Sub btnDeletePath_Click()
    If Me.PathsCurrent.ListIndex <> -1 Then
        If Left(Me.PathsCurrent, 1) = "*" Then
            Dim PathName As String, wpid As String
            PathName = Me.PathsCurrent
            PathName = Mid(PathName, 2)
            PathName = Trim(PathName)
            wpid = Mid(PathName, InStr(PathName, vbTab) + 1)
            PathName = Left(PathName, InStr(PathName, vbTab) - 1)
            If idsToDel.Find(wpid) <> "-1" Then
                Exit Sub
            End If
            If vbNo = MsgBox("Remove path " & qq(PathName) & "?", vbYesNo) Then
                Exit Sub
            End If
            
            'ok to remove
            Dim thisid As String
            thisid = wpid
            Dim tempHist As PerlArray
            Set tempHist = New PerlArray
            While thisid <> "-1"
                tempHist.UnShift thisid
                thisid = WebPathCache_GetParent(thisid)
            Wend
            Dim i As Long
            For i = 0 To tempHist.Scalar - 1
                Me.PathsToRemove.AddItem IIf(i = tempHist.Scalar - 1, "*", "") & Space(2 * (i)) & WebPathCache_GetName(tempHist.Elem(i)) & vbTab & tempHist.Elem(i)
            Next i
            idsToDel.Push wpid
        End If
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    saveChanges
    Unload Me
End Sub

Private Sub Form_Load()
    SetListTabs Me.PathsCurrent, Array(300)
    SetListTabs Me.PathsToAdd, Array(300)
    SetListTabs Me.PathsToRemove, Array(300)
    SetListTabs Me.Browser, Array(300)
    SetListTabs Me.BrowserDetails, Array(300)
    
    Set browserHistory = New PerlArray
    browserHistory.Push "-1"
    Set idsToAdd = New PerlArray
    Set idsToDel = New PerlArray
    
    Dim item As String
    item = SignMaintenance.ItemNumber
    Me.hiddenWebsiteID = SignMaintenance.CurrentWebsiteID
    
    Dim i As Long
    For i = 0 To SignMaintenance.WebPathTree.ListCount - 1
        Me.PathsCurrent.AddItem SignMaintenance.WebPathTree.List(i)
    Next i
    
    loadBrowser
End Sub

Private Sub Browser_Click()
    If browserClicked = Me.Browser Then
        Exit Sub
    End If
    
    If Me.Browser = ".." Then
        Me.BrowserDetails.Clear
        browserClicked = Me.Browser
        Exit Sub
    End If
    Dim wpid As String
    wpid = Mid(Me.Browser, InStr(Me.Browser, vbTab) + 1)
    
    Me.BrowserDetails.Clear
    
    Dim kids As Variant
    kids = WebPathCache_GetChildList(wpid, False)
    Dim i As Long
    Me.BrowserDetails.AddItem "Subsections: " & CStr(UBound(kids) + 1)
    For i = 0 To UBound(kids)
        Me.BrowserDetails.AddItem "  " & WebPathCache_GetName(CStr(kids(i)))
    Next i
    hasSections = Not CBool(UBound(kids) = -1)
    
    Dim rst As ADODB.Recordset
    Set rst = DB.retrieve("SELECT ItemNumber FROM PartNumberWebPaths WHERE WebPathID=" & wpid)
    Me.BrowserDetails.AddItem ""
    hasItems = Not CBool(rst.EOF)
    Me.BrowserDetails.AddItem "Items: " & rst.RecordCount
    While Not rst.EOF
        Me.BrowserDetails.AddItem "  " & rst("ItemNumber")
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    
    browserClicked = Me.Browser
End Sub

Private Sub Browser_DblClick()
    browserClicked = ""
    If Me.Browser = ".." Then
        browserHistory.Pop
    Else
        browserHistory.Push Mid(Me.Browser, InStr(Me.Browser, vbTab) + 1)
    End If
    loadBrowser
End Sub

Private Sub loadBrowser()
    Dim wpid As String, i As Long, paths As Variant
    wpid = browserHistory.Elem(browserHistory.Scalar - 1)
    
    If wpid = "-1" Then
        Select Case True
            Case Is = Me.opPath1Filter(0)
                paths = WebPathCache_GetRoots(CLng(Me.hiddenWebsiteID), True, True)
            Case Is = Me.opPath1Filter(1)
                paths = WebPathCache_GetRoots(CLng(Me.hiddenWebsiteID), True, False)
            Case Is = Me.opPath1Filter(2)
                paths = WebPathCache_GetRoots(CLng(Me.hiddenWebsiteID), False, True)
        End Select
    Else
        paths = WebPathCache_GetChildList(wpid, False)
        If UBound(paths) = -1 Then
            'alert? unset the paths, and pop the history, since
            'we already added it
            paths = Empty
            browserHistory.Pop
        End If
    End If
    
    If VarType(paths) <> vbEmpty Then
        Me.Browser.Clear
        Me.BrowserDetails.Clear
        If wpid <> "-1" Then
            Me.Browser.AddItem ".."
        End If
        For i = 0 To UBound(paths)
            Me.Browser.AddItem WebPathCache_GetName(CStr(paths(i))) & vbTab & CStr(paths(i))
        Next i
        
        Me.lblPreviousPaths.Caption = ""
        For i = 1 To browserHistory.Scalar - 1
            Me.lblPreviousPaths.Caption = IIf(Me.lblPreviousPaths.Caption = "", "", Me.lblPreviousPaths.Caption & " > ") & WebPathCache_GetName(browserHistory.Elem(i))
        Next i
    End If
End Sub

Private Sub opPath1Filter_Click(Index As Integer)
    Set browserHistory = New PerlArray
    browserHistory.Push "-1"
    
    loadBrowser
End Sub

Private Sub PathsToAdd_DblClick()
    If Left(Me.PathsToAdd, 1) = "*" Then
        Dim wpid As String, i As Long
        wpid = Mid(Me.PathsToAdd, InStr(Me.PathsToAdd, vbTab) + 1)
        For i = 0 To idsToAdd.Scalar - 1
            If idsToAdd.Elem(i) = wpid Then
                idsToAdd.Elem(i) = ""
                Exit For
            End If
        Next i
        
        Dim startdx As Long
        startdx = Me.PathsToAdd.ListIndex
        Me.PathsToAdd.RemoveItem startdx
        For i = startdx - 1 To 0 Step -1
            If Left(Me.PathsToAdd.List(i), 1) = " " Then
                Me.PathsToAdd.RemoveItem i
            Else
                Me.PathsToAdd.RemoveItem i
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub PathsToRemove_DblClick()
    If Left(Me.PathsToRemove, 1) = "*" Then
        Dim wpid As String, i As Long
        wpid = Mid(Me.PathsToRemove, InStr(Me.PathsToRemove, vbTab) + 1)
        For i = 0 To idsToDel.Scalar - 1
            If idsToDel.Elem(i) = wpid Then
                idsToDel.Elem(i) = ""
                Exit For
            End If
        Next i
        
        Dim startdx As Long
        startdx = Me.PathsToRemove.ListIndex
        For i = Me.PathsToRemove.ListIndex To 0 Step -1
            If Left(Me.PathsToRemove.List(i), 1) = " " Then
                Me.PathsToRemove.RemoveItem i
            Else
                Me.PathsToRemove.RemoveItem i
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub saveChanges()
    Dim i As Long
    For i = 0 To idsToDel.Scalar - 1
        If idsToDel.Elem(i) <> "" Then
            DB.Execute "DELETE FROM PartNumberWebPaths WHERE ItemNumber='" & SignMaintenance.ItemNumber & "' AND WebPathID=" & idsToDel.Elem(i)
        End If
    Next i
    For i = 0 To idsToAdd.Scalar - 1
        If idsToAdd.Elem(i) <> "" Then
            DB.Execute "INSERT INTO PartNumberWebPaths ( ItemNumber, WebPathID ) VALUES ( '" & SignMaintenance.ItemNumber & "', " & idsToAdd.Elem(i) & " )"
        End If
    Next i
End Sub
