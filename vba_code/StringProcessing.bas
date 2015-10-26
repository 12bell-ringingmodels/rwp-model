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


Public Function FindLast(ByVal InputString As String, ByVal matchCharacters As String) As Integer
    Dim matchLocation As Integer
    
    FindLast = 0
    
    For matchLocation = Len(InputString) To 1 Step -1
        If InStr(matchCharacters, Mid(InputString, matchLocation, 1)) > 0 Then
            FindLast = matchLocation
            Exit Function
        End If
    Next
End Function
