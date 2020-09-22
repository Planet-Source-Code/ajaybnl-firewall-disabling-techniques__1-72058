Attribute VB_Name = "H"
Option Explicit

Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Const MAX_PATH = 260

Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_NORMAL = &H80




Sub FindFiles(DirPath As String, Optional FileSpec As String = "*.*")
Dim FindData As WIN32_FIND_DATA
Dim FindHandle As Long
Dim FindNextHandle As Long
Dim filestring As String

DirPath = Trim$(DirPath)

If Right(DirPath, 1) <> "\" Then
DirPath = DirPath & "\"
End If

FindHandle = FindFirstFile(DirPath & FileSpec, FindData)
DoEvents
If FindHandle <> 0 Then
  If (FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
    If Left$(FindData.cFileName, 1) <> "." And Left$(FindData.cFileName, 2) <> ".." Then FindFiles DirPath & Left(FindData.cFileName, InStr(1, FindData.cFileName, Chr(0)) - 1), FileSpec
ElseIf Len(Left(FindData.cFileName, InStr(1, FindData.cFileName, Chr(0)) - 1)) > 0 Then
'Process File
Form1.Process DirPath & Left(FindData.cFileName, InStr(1, FindData.cFileName, Chr(0)) - 1)
  End If
End If

' Now loop and find the rest of the files
If FindHandle <> 0 Then
  Do

    FindNextHandle = FindNextFile(FindHandle, FindData)
    If FindNextHandle <> 0 Then
      If (FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
        If Left$(FindData.cFileName, 1) <> "." And Left$(FindData.cFileName, 2) <> ".." Then FindFiles DirPath & Left(FindData.cFileName, InStr(1, FindData.cFileName, Chr(0)) - 1), FileSpec
  ElseIf Len(Left(FindData.cFileName, InStr(1, FindData.cFileName, Chr(0)) - 1)) > 0 Then
'Process File
Form1.Process DirPath & Left(FindData.cFileName, InStr(1, FindData.cFileName, Chr(0)) - 1)
      End If
    Else
    Exit Do
    End If
  Loop
End If

Call FindClose(FindHandle)

End Sub

