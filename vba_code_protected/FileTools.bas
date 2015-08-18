Attribute VB_Name = "FileTools"
Option Explicit

Function JoinPath(ByVal path1 As String, ByVal path2 As String) As String
    If path1 <> "" Then
        If Right(path1, 1) = "\" Then
            JoinPath = path1 & path2
        Else
            JoinPath = path1 & "\" & path2
        End If
    Else
        JoinPath = path2
    End If
End Function
Function CvtToAbsFile(ByVal inputFile As String) As String
    If InStr(inputFile, "\") = 1 Or InStr(inputFile, ":\") = 2 Or Len(inputFile) = 0 Then
        ' String is already absolute
        CvtToAbsFile = inputFile
    Else
        CvtToAbsFile = ThisWorkbook.Path & "\" & inputFile
    End If
End Function
