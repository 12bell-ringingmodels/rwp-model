Attribute VB_Name = "FileTools"
Option Explicit

Type FileParts
    path As String
    title As String
    extension As String
End Type
   

Public Function JoinPath(ByVal path1 As String, ByVal path2 As String) As String
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
Public Function CvtToAbsFile(ByVal inputFile As String) As String
    If InStr(inputFile, "\") = 1 Or InStr(inputFile, ":\") = 2 Or Len(inputFile) = 0 Then
        ' String is already absolute
        CvtToAbsFile = inputFile
    Else
        CvtToAbsFile = ThisWorkbook.path & "\" & inputFile
    End If
End Function

Public Function GetFileTitle(ByVal filename As String) As String
    Dim extensionIndex As Integer
    Dim pathIndex As Integer
    Dim idx As Integer
    Dim thisCh As String

    extensionIndex = Len(filename) + 1
    pathIndex = 0
    
    For idx = Len(filename) To 1 Step -1
        thisCh = Mid(filename, idx, 1)
        If thisCh = "." Then
            If extensionIndex >= Len(filename) Then
                extensionIndex = idx
            End If
        ElseIf thisCh = "\" Then
            pathIndex = idx
            Exit For
        End If
    Next
    
    
    GetFileTitle = Mid(filename, pathIndex + 1, extensionIndex - pathIndex - 1)
    
End Function

Public Function SplitFilenameToParts(ByVal inputFile As String) As FileParts
    Dim index As Integer
    Dim testCh As String
    Dim foundItem As Boolean
    
    SplitFilenameToParts.path = ""
    SplitFilenameToParts.title = ""
    SplitFilenameToParts.extension = ""
    
    index = FindLast(inputFile, "\/")
    
    If index = 0 Then
        index = FindLast(inputFile, ":")
        If index > 0 Then
            SplitFilenameToParts.path = Mid(inputFile, 1, index)
        End If
    Else
        If index = 2 And (Mid(inputFile, 1, 1) = "\" Or Mid(inputFile, 1, 1) = "/") Then
            SplitFilenameToParts.path = inputFile
            index = Length(inputFile)
        Else
            SplitFilenameToParts.path = Mid(inputFile, 1, index - 1)
        End If
    End If
    If index = 0 Then
        SplitFilenameToParts.title = inputFile
    Else
        If Len(SplitFilenameToParts.path) > 0 And Right(SplitFilenameToParts.path, 1) = ":" And (Len(SplitFilenameToParts.path) > 2 Or (Len(inputFile) >= 3 And Mid(inputFile, 3, 1) = "\")) Then
            SplitFilenameToParts.path = SplitFilenameToParts.path & "\"
        ElseIf Len(Trim(SplitFilenameToParts.path)) = 0 Then
            SplitFilenameToParts.path = "\"
        End If
        SplitFilenameToParts.title = Mid(inputFile, index + 1, Length(inputFile))
    End If
        
    If Len(SplitFilenameToParts.title) = 0 Then
        Exit Function
    End If
    
    index = FindLast(SplitFilenameToParts.title, ".")
    If index = 0 Then
        Exit Function
    Else
        SplitFilenameToParts.extension = Mid(SplitFilenameToParts.title, index, Length(SplitFilenameToParts.title))
        SplitFilenameToParts.title = Left(SplitFilenameToParts.title, index - 1)
    End If
    
End Function



