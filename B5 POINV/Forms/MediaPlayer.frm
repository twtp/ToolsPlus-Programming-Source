VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LogList 
      Height          =   2205
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   11265
   End
   Begin VB.Timer TimerSeconds 
      Left            =   4200
      Top             =   240
   End
   Begin VB.Label DirectoryLabel 
      Caption         =   "d:\music"
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Top             =   90
      Width           =   4365
   End
   Begin VB.Label MediaNameLabel 
      Caption         =   "media name"
      Height          =   495
      Left            =   90
      TabIndex        =   2
      Top             =   960
      Width           =   4515
   End
   Begin VB.Label StateLabel 
      Caption         =   "state"
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   690
      Width           =   4515
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer 
      Height          =   0
      Left            =   1650
      TabIndex        =   0
      Top             =   480
      Width           =   0
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "invisible"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   0
      _cy             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum wmpps
    Undefined = 0
    Stopped = 1
    Paused = 2
    Playing = 3
    ScanForward = 4
    ScanReverse = 5
    Buffering = 6
    Waiting = 7
    MediaEnded = 8
    Transitioning = 9
    Ready = 10
    Reconnecting = 11
    Last = 12
End Enum

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 255
    cAlternate As String * 14
End Type

Private Const ERROR_NO_MORE_FILES = 18&
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private timerMinutes As Long
Private firstRun As Boolean

Private Sub Form_Load()
    Dim args As String
    args = Command()
    If args <> "" Then
        If Left(args, 1) = """" And Right(args, 1) = """" Then
            args = Mid(args, 2, Len(args) - 2)
        End If
        Me.DirectoryLabel.Caption = args
    End If
    
    Call log("overheadlooper initializing")
    Call Randomize
    
    Call Me.WindowsMediaPlayer.settings.setMode("loop", True)
    Call Me.WindowsMediaPlayer.settings.setMode("shuffle", True)
    
    timerMinutes = 9999
    firstRun = True
    Me.TimerSeconds.Interval = 1000
    Me.TimerSeconds.Enabled = True
End Sub

Private Sub TimerSeconds_Timer()
    Me.TimerSeconds.Enabled = False
    
    timerMinutes = timerMinutes + 1
    'Debug.Print "timer firing, total elapsed time is " & timerMinutes & " minutes"
    If timerMinutes > 720 Then ' every 12 hours, reload pls
        timerMinutes = 0
        'Debug.Print "threshold passed, updating playlist"
        updatePlaylist
    End If
    
    Dim shouldBePlaying As Boolean
    shouldBePlaying = CBool(DatePart("h", DateTime.Now) >= 7 And DatePart("h", DateTime.Now) < 19)
    
    If shouldBePlaying Then
        Select Case Me.WindowsMediaPlayer.playState
            Case Is = wmpps.Undefined, wmpps.Stopped, wmpps.Paused, wmpps.MediaEnded, wmpps.Ready, wmpps.Last
                'Debug.Print "player is off, but should be playing, starting"
                Call log("starting player")
                If Me.WindowsMediaPlayer.currentPlaylist.Count > 0 Then
                    Me.WindowsMediaPlayer.Controls.currentItem = Me.WindowsMediaPlayer.currentPlaylist.Item(Int(Rnd * Me.WindowsMediaPlayer.currentPlaylist.Count))
                    Call Me.WindowsMediaPlayer.Controls.play
                End If
        End Select
    Else
        Select Case Me.WindowsMediaPlayer.playState
            Case Is = wmpps.Undefined, wmpps.Paused, wmpps.Playing, wmpps.ScanForward, wmpps.ScanReverse, wmpps.Buffering, wmpps.Transitioning, wmpps.Ready, wmpps.Reconnecting, wmpps.Last
                Call log("stopping player")
                'Debug.Print "player is on, but should be off, stopping"
                Call Me.WindowsMediaPlayer.Controls.stop
        End Select
    End If
    
    Me.TimerSeconds.Interval = 60000
    Me.TimerSeconds.Enabled = True
End Sub

Private Sub WindowsMediaPlayer_MediaChange(ByVal Item As Object)
    Me.MediaNameLabel.Caption = Me.WindowsMediaPlayer.currentMedia.Name
    'Call log("MediaChange: " & Me.WindowsMediaPlayer.currentMedia.sourceURL)
    'Debug.Print "MediaChange: " & Me.WindowsMediaPlayer.currentMedia.sourceURL
    Me.LogList.AddItem "MediaChange: " & Me.WindowsMediaPlayer.currentMedia.sourceURL
End Sub

Private Sub WindowsMediaPlayer_PlayStateChange(ByVal NewState As Long)
    Select Case NewState
        Case Is = wmpps.Undefined
            Me.StateLabel.Caption = "State: Undefined"
        Case Is = wmpps.Stopped
            Me.StateLabel.Caption = "State: Stopped"
        Case Is = wmpps.Paused
            Me.StateLabel.Caption = "State: Paused"
        Case Is = wmpps.Playing
            Me.StateLabel.Caption = "State: Playing"
        Case Is = ScanForward
            Me.StateLabel.Caption = "State: Scanning Forward"
        Case Is = ScanReverse
            Me.StateLabel.Caption = "State: Scanning Reverse"
        Case Is = Buffering
            Me.StateLabel.Caption = "State: Buffering"
        Case Is = Waiting
            Me.StateLabel.Caption = "State: Waiting"
        Case Is = MediaEnded
            Me.StateLabel.Caption = "State: Media Ended"
        Case Is = Transitioning
            Me.StateLabel.Caption = "State: Transitioning"
        Case Is = Ready
            Me.StateLabel.Caption = "State: Ready"
        Case Is = Reconnecting
            Me.StateLabel.Caption = "State: Reconnecting"
        Case Is = Last
            Me.StateLabel.Caption = "State: Last Playstate"
        Case Else
            Me.StateLabel.Caption = "State: Unknown"
    End Select
    'Debug.Print Me.StateLabel.Caption
    Me.LogList.AddItem Me.StateLabel.Caption
    DoEvents
End Sub

Private Function readPlaylistFolder(dir As Scripting.Folder) As Scripting.Dictionary
    Dim iterFile As Scripting.File, iterFolder As Scripting.Folder, retval As Scripting.Dictionary
    Set retval = New Scripting.Dictionary
    'Debug.Print "----- ITERATING DIRECTORY -----"
    For Each iterFile In dir.Files
        If Right(iterFile.Name, 4) = ".mp3" Or Right(iterFile.Name, 4) = ".wma" Then
            'Debug.Print "found file " & CStr(iterFile)
            retval.Add CStr(iterFile.Path), 1 'full path+fname
        End If
    Next iterFile
    
    For Each iterFolder In dir.SubFolders
        Dim temp As Scripting.Dictionary
        Set temp = readPlaylistFolder(iterFolder)
        Dim iterString As Variant
        For Each iterString In temp.Keys
            retval.Add CStr(iterString), 1
        Next iterString
    Next iterFolder
    
    Set readPlaylistFolder = retval
End Function

Private Sub updatePlaylist()
On Error GoTo errh
    Call log("playlist update started")
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim dir As Scripting.Folder
    Set dir = fso.GetFolder(Me.DirectoryLabel.Caption)
    
    Dim found As Scripting.Dictionary
    Set found = readPlaylistFolder(dir)
    
    'Debug.Print "----- CHECKING PLAYLIST -----"
    ReDim toRemove(0 To 4) As WMPLibCtl.IWMPMedia3
    Dim toRemoveCount As Long
    toRemoveCount = -1
    Dim i As Long
    For i = 0 To Me.WindowsMediaPlayer.currentPlaylist.Count - 1
        If found.Exists(Me.WindowsMediaPlayer.currentPlaylist.Item(i).sourceURL) Then
            'Debug.Print "no change for file " & Me.WindowsMediaPlayer.currentPlaylist.Item(i).sourceURL
            found.Remove Me.WindowsMediaPlayer.currentPlaylist.Item(i).sourceURL
        Else
            'Debug.Print "removing dead file " & Me.WindowsMediaPlayer.currentPlaylist.Item(i).sourceURL
            Call log("removing dead file " & CStr(Me.WindowsMediaPlayer.currentPlaylist.Item(i).sourceURL))
            If toRemoveCount > UBound(toRemove) - 1 Then
                ReDim Preserve toRemove(0 To 2 * UBound(toRemove)) As WMPLibCtl.IWMPMedia3
            End If
            toRemoveCount = toRemoveCount + 1
            Set toRemove(toRemoveCount) = Me.WindowsMediaPlayer.currentPlaylist.Item(i)
        End If
    Next i

    For i = 0 To toRemoveCount
        Me.WindowsMediaPlayer.currentPlaylist.RemoveItem toRemove(i)
    Next i

    Dim addedCount As Long
    addedCount = 0
    Dim iter As Variant
    For Each iter In found.Keys
        'Debug.Print "adding new file " & CStr(iter)
        If Not firstRun Then
            Call log("adding new file " & CStr(iter))
        End If
        Me.WindowsMediaPlayer.currentPlaylist.appendItem Me.WindowsMediaPlayer.newMedia(CStr(iter))
        addedCount = addedCount + 1
    Next iter

    'Debug.Print "----- PLAYLIST UPDATED -----"
    Call log("playlist update completed, " & addedCount & " added, " & toRemoveCount + 1 & " removed")
    Exit Sub
errh:
    Call log("playlist update error " & Err.Number & ": " & Err.Description)
    Err.Clear
    Exit Sub
End Sub

Private Sub log(msg As String)
    Open "overheadlooper.log" For Append As #1
    Print #1, Format(DateTime.Now, "yyyy-MM-dd hh:mm:ss") & ": " & msg
    Close #1
    Me.LogList.AddItem msg
End Sub
