VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CProfiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private profileColumn As Integer
Private profileRow As Integer
Private ticTime As Long
Private lastTockTime As Long

Public Sub StartProfiling(ByVal taskName As String)
    Dim profileWs As Worksheet
    
    Set profileWs = GetProfileWorksheet()
    
    If profileWs Is Nothing Then Exit Sub
    
    profileColumn = profileWs.UsedRange.Columns.Count + 1
    
    profileWs.Cells(1, profileColumn) = taskName
    profileWs.Cells(1, profileColumn + 1) = "Delta"
    profileWs.Cells(1, profileColumn + 2) = "Abs"
    
    profileRow = 1
    
    ticTime = GetTickCount()
    lastTockTime = 0
End Sub

Public Sub EndProfiling(ByVal displayChart As Boolean)
    profileColumn = 0
End Sub

Public Sub Tock(ByVal itemName As String)
    Dim profileWs As Worksheet
    Dim tockTime As Long
    
    If profileColumn <= 0 Then Exit Sub
    
    Set profileWs = GetProfileWorksheet()
    If profileWs Is Nothing Then Exit Sub
    
    tockTime = GetTickCount() - ticTime

    profileWs.Cells(profileRow, profileColumn) = itemName
    profileWs.Cells(profileRow, profileColumn + 1) = tockTime - lastTockTime
    profileWs.Cells(profileRow, profileColumn + 2) = tockTime
    
    profileRow = profileRow + 1
    
    lastTockTime = tockTime
End Sub

Private Function GetProfileWorksheet() As Worksheet
    Const PROFILE_WORKSHEET As String = "ProfileOutput"

    If Not IsProfiling() Then
        Set GetProfileWorksheet = Nothing
    Else
        If WorksheetExists(PROFILE_WORKSHEET) Then
            Set GetProfileWorksheet = Sheets(PROFILE_WORKSHEET)
        Else
            Dim ws As Worksheet
            Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
            ws.Name = PROFILE_WORKSHEET
            Set GetProfileWorksheet = ws
        End If
    End If
End Function

Private Property Get IsProfiling() As Boolean
    If WorksheetExists("Settings") Then
        IsProfiling = (Sheets("Settings").Range("optDevProfile").Value = True)
    Else
        IsProfiling = False
    End If
End Property

Private Sub Class_Initialize()

End Sub
