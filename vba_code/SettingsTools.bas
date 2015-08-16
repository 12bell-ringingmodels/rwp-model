Attribute VB_Name = "SettingsTools"
Public Const SETTINGS_WORKSHEET As String = "Settings"

Public Enum PresentationModes
    General = 1
    PracticeFeedback = 2
    JudgesFeedback = 3
    ContestFeedback = 4
End Enum

Public Enum HandstrokeGapModes
    Averages = 1
    MinimumSquaredError = 2
End Enum

    
Public Function GetPresentationMode() As PresentationModes

    If WorksheetExists(SETTINGS_WORKSHEET) Then
        GetPresentationMode = Worksheets(SETTINGS_WORKSHEET).Range("optPresMode").Value
    Else
        GetPresentationMode = PresentationModes.General
    End If

End Function

Public Function GetHandstrokeGapMode() As HandstrokeGapModes
    If WorksheetExists(SETTINGS_WORKSHEET) Then
        GetHandstrokeGapMode = Worksheets(SETTINGS_WORKSHEET).Range("optAnalysisHandstrokeGapMethod").Value
    Else
        GetHandstrokeGapMode = HandstrokeGapModes.Averages
    End If
End Function
