Attribute VB_Name = "Copier"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const DISTRIBUTION_DIR As String = "s:\mastest\mas90-signs\A_Dist\"
Private Const CONF_DIR As String = "s:\mastest\mas90-signs\conf_files\"

Public Sub Main()
    Screen.MousePointer = vbHourglass
    
    Dim exe_name As String
    ReDim Files(0) As String
    
On Error GoTo errh
    Open CONF_DIR & "FECopier.conf" For Input As #1
On Error GoTo 0
    Dim buf As String
    While Not EOF(1)
        Line Input #1, buf
        If buf = "" Then
            'next
        ElseIf Left(buf, 1) = "#" Then
            'next
        Else
            Files(UBound(Files)) = buf
            ReDim Preserve Files(UBound(Files) + 1) As String
        End If
    Wend
    ReDim Preserve Files(UBound(Files) - 1) As String
    Close #1
    
    Dim DESTINATION_DIR As String
    DESTINATION_DIR = Environ("TEMP") & "\poinv\"
    Dim SYSTEM_DIR As String
    SYSTEM_DIR = Environ("SystemRoot") & "\system32\"
    
On Error GoTo errh
    If Not FolderExists(DESTINATION_DIR) Then
        CreateFolder DESTINATION_DIR
    End If
    If Not FolderExists(DISTRIBUTION_DIR) Then
        MsgBox "Unable to find the distribution folder! Try talking to Eric N. or Brian?"
        GoTo done
    End If
    Dim dirContents As Variant
    dirContents = GetFolderContents(DISTRIBUTION_DIR)
    Dim i As Long, j As Long, found As Boolean, needToCopy As Boolean
    ReDim tocopy(UBound(Files)) As Boolean
    Dim path As String
    For i = LBound(Files) To UBound(Files)
        found = False
        For j = LBound(dirContents) To UBound(dirContents)
            If LCase(dirContents(j)) = LCase(Files(i)) Then
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            MsgBox "Unable to find the program's file, " & Files(i) & vbCrLf & vbCrLf & "Brian can probably fix this."
            GoTo done
        Else
            Select Case LCase(Right(Files(i), 4))
                Case Is = ".dll", ".ocx"
                    path = SYSTEM_DIR
                Case Is = ".exp", ".lib"
                    path = SYSTEM_DIR
                Case Else
                    path = DESTINATION_DIR
            End Select
            If GetFileModifiedDate(DISTRIBUTION_DIR & Files(i)) > GetFileModifiedDate(path & Files(i)) Then
                tocopy(i) = True
                needToCopy = True
            End If
            If LCase(Right(Files(i), 4)) = ".exe" Then
                exe_name = Files(i) 'probably the only exe
            End If
        End If
    Next i

    If needToCopy Then
        Load CopierDialog
        CopierDialog.Show
        DoEvents
        For i = LBound(tocopy) To UBound(tocopy)
            If tocopy(i) Then
                Select Case LCase(Right(Files(i), 4))
                    Case Is = ".dll", ".ocx"
                        If FileExists(SYSTEM_DIR & Files(i)) Then
                            ShellWait "regsvr32 /u /s " & """" & SYSTEM_DIR & Files(i) & """", vbHide
                        End If
                        Copy DISTRIBUTION_DIR & Files(i), SYSTEM_DIR & Files(i), True
                        ShellWait "regsvr32 /s " & """" & SYSTEM_DIR & Files(i) & """", vbHide
                    Case Is = ".exp", ".lib"
                        Copy DISTRIBUTION_DIR & Files(i), SYSTEM_DIR & Files(i), True
                    Case Else
                        Copy DISTRIBUTION_DIR & Files(i), DESTINATION_DIR & Files(i), True
                End Select
            End If
        Next i
        Unload CopierDialog
        DoEvents
    End If

exec:
    ShellExecute 0, "open", DESTINATION_DIR & exe_name, "", "c:\", 1
done:
    Screen.MousePointer = vbNormal
    Exit Sub
    
errh:
    If Err.Number = 76 Then
        MsgBox "Can't find the file path! This probably means your network drives are broken!" & vbCrLf & vbCrLf & "An easy way to fix this is to just log out and back in again."
    ElseIf Err.Number = 70 Then
        If MsgBox("An updated version is available, but can't be copied since the file is in use." & vbCrLf & vbCrLf & "Open the old version instead?", vbYesNo) = vbYes Then
            Err.Clear
            GoTo exec
        End If
    Else
        MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & Err.Description
    End If
    GoTo done
End Sub
