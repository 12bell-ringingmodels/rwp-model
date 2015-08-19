Attribute VB_Name = "StringProcessing"
Option Explicit

Public Function SecondsToHoursAndMinutes(ByVal timeInSeconds, Optional ByVal hrMarker = "h", Optional ByVal minMarker = "m") As String
    Dim hours As Integer
    Dim minutes As Integer
    Dim seconds As Integer
    
    If IsNumeric(timeInSeconds) Then
        minutes = Round(timeInSeconds / 60, 0)
        hours = Int(minutes / 60)
        minutes = minutes - 60 * hours
        If hours > 0 Then
            SecondsToHoursAndMinutes = hours & hrMarker & " " & minutes & minMarker
        Else
            SecondsToHoursAndMinutes = minutes & minMarker
        End If
    Else
        SecondsToHoursAndMinutes = "#N/A"
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
