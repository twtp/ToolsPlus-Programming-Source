Attribute VB_Name = "ARGeneral"
Option Explicit

Public Function getSubDir(opgroup As Object)
    If opgroup(MUSIC_PLS) = True Then
        getSubDir = MUSIC_SUBDIR
    ElseIf opgroup(ANNOUNCEMENT_PLS) = True Then
        getSubDir = ANNOUNCEMENT_SUBDIR
    Else
        'uh oh
        Err.Raise "1010", "ARGeneral.getSubDir", "No option in group selected"
    End If
End Function
