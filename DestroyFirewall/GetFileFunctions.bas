Attribute VB_Name = "G"
'Get File From Path + File
Function GetFile(A2 As String) As String
On Error GoTo end1
If Len(A2) <= 3 Then GetFile = A2: Exit Function
Dim A3, a4, a5, a6
For a4 = 0 To Len(A2)
For A3 = 0 To a4
If Left(Right(A2, a4), A3) = "\" Or Left(Right(A2, a4), A3) = "/" Then
GetFile = Right(A2, a4 - 1)
Exit Function
End If
Next A3
Next a4
GetFile = A2
Exit Function
end1:
GetFile = A2
End Function
'Get Extention on File From Path+File or File
Function Remext(A2 As String) As String
Remext = GetFile(A2)
If InStr(1, Remext, ".") > 0 Then
Remext = Left(Remext, Len(Remext) - (Len(GetExt(Remext)) + 1))
End If
End Function
Function GetExt(A2 As String) As String
On Error Resume Next
Dim A3, a4, a5, a6
For a4 = 0 To Len(A2)
For A3 = 0 To a4
If Left(Right(A2, a4), A3) = "." Then
GetExt = Right(A2, a4 - 1)
Exit Function
End If
Next A3
Next a4
End Function
'Get Path With Slash From Path
Function GetPath(A2 As String) As String
On Error Resume Next
Dim A3, a4, a5, a6
For a4 = 0 To Len(A2)
For A3 = 0 To a4
If Left(Right(A2, a4), A3) = "\" Then
GetPath = Replace(A2, Right(A2, a4), "")
GoTo end1
End If
Next A3
Next a4
end1:
End Function
