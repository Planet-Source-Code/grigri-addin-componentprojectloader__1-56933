Attribute VB_Name = "FileSystem"
Option Explicit

Private Const MAX_PATH As Long = 260
Private Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameA" (ByVal lpShortPath As String, ByVal lpLongPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function PathCanonicalize Lib "shlwapi.dll" Alias "PathCanonicalizeA" (ByVal lpDest As String, ByVal lpSrc As String) As Long

' ·······················
' · LongPathFromShort() ·
' ·······················
'
' The 8.3 file format went out with the dinosaurs, but it's still used
' sometimes. This function converts from the 8.3 format to the standard
' "Long Filenames" format.
'
Public Function LongPathFromShort(ByVal sShortPath As String) As String
    LongPathFromShort = Space$(MAX_PATH)
    GetLongPathName sShortPath, LongPathFromShort, MAX_PATH
    LongPathFromShort = Left$(LongPathFromShort, InStr(LongPathFromShort, vbNullChar) - 1)
End Function

' ······················
' · Parent Directory() ·
' ······················
'
' Simple function to get either the directory a folder is in, or the
' parent directory of another directory.
' Keeping the trailing backslash is optional.
'
Public Function ParentDirectory(ByVal sChildPath As String, Optional AddBackSlash As Boolean = True) As String
    Dim i As Integer
    ' Find the last backslash in the path [unless the last character is a backslash]
    i = InStrRev(sChildPath, "\", Len(sChildPath) - 1)
    If AddBackSlash Then
        ParentDirectory = Left$(sChildPath, i)
    Else
        ParentDirectory = Left$(sChildPath, i - 1)
    End If
End Function

' ······················
' · CanonicalizePath() ·
' ······················
'
' This function tidies up paths containing "." or ".." referring to parent
' directories and so on, making them all nice and neat.
' Example: CanonicalizePath("C:\Stuff\Project\Files\..\File.bas")
'          = "C:\Stuff\Project\Files.bas"
'
Public Function CanonicalizePath(ByVal sPath As String) As String
    CanonicalizePath = Space$(MAX_PATH)
    If PathCanonicalize(CanonicalizePath, sPath) Then
        CanonicalizePath = Left$(CanonicalizePath, InStr(CanonicalizePath, vbNullChar) - 1)
    Else
        CanonicalizePath = vbNullString
    End If
End Function
