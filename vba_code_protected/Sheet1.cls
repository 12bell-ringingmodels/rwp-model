VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Browse1_Click()
    Call findfile(4, 3)
End Sub

Private Sub Browse10_Click()
    Call findfile(13, 3)
End Sub

Private Sub Browse2_Click()
    Call findfile(5, 3)
End Sub

Private Sub Browse3_Click()
    Call findfile(6, 3)
End Sub

Private Sub Browse4_Click()
    Call findfile(7, 3)
End Sub

Private Sub Browse5_Click()
    Call findfile(8, 3)
End Sub

Private Sub Browse6_Click()
    Call findfile(9, 3)
End Sub

Private Sub Browse7_Click()
    Call findfile(10, 3)
End Sub

Private Sub Browse8_Click()
    Call findfile(11, 3)
End Sub

Private Sub Browse9_Click()
    Call findfile(12, 3)
End Sub

Private Sub BrowseSound1_Click()
    Call findfile(4, 5)
End Sub

Private Sub BrowseSound10_Click()
    Call findfile(13, 5)
End Sub

Private Sub BrowseSound2_Click()
    Call findfile(5, 5)
End Sub

Private Sub BrowseSound3_Click()
    Call findfile(6, 5)
End Sub

Private Sub BrowseSound4_Click()
    Call findfile(7, 5)
End Sub

Private Sub BrowseSound5_Click()
    Call findfile(8, 5)
End Sub

Private Sub BrowseSound6_Click()
    Call findfile(9, 5)
End Sub

Private Sub BrowseSound7_Click()
    Call findfile(10, 5)
End Sub

Private Sub BrowseSound8_Click()
    Call findfile(11, 5)
End Sub

Private Sub BrowseSound9_Click()
    Call findfile(12, 5)
End Sub

Private Sub BrowseTouch_Click()
    Call findfile(15, 2)
End Sub


Private Sub Load_Click()
    Call loadfiles
End Sub


Private Function GetMediaPlayer() As WindowsMediaPlayer
    
    ' The previous implementation seems to rely on having a media controller
    ' hanging around on a worksheet
    If WorksheetExists("MediaPlayer") Then
        Set mediaWksht = Worksheets("MediaPlayer")
        Set GetMediaPlayer = mediaWksht.WindowsMediaPlayer1
    Else
        Set GetMediaPlayer = Nothing
    End If

End Function

Public Sub PlayFromHere()
Attribute PlayFromHere.VB_ProcData.VB_Invoke_Func = "P\n14"
Dim strTeam As String
Dim mediaPlayer As WindowsMediaPlayer
    'load vars
    Call loadvars
   
    
    Set mediaPlayer = GetMediaPlayer()
    
    If mediaPlayer Is Nothing Then
        ' There's going to be no luck here
        Exit Sub
    End If
        
    'get the team name and asociated soundfile
    strTeam = ActiveSheet.Name
    i = Len(strTeam)
    strTeam = Mid(strTeam, 1, i - 2)
    'MsgBox ("Team : " + strTeam + " cell= " + Str(ActiveCell.Column) + " " + Str(ActiveCell.row))
    
    'find the sound file
    For i = 1 To 10
      If TeamName(i) = strTeam Then Exit For
    Next i
    
    If i < 11 Then
        mediaPlayer.URL = soundfile(i)
        'now look for start point
        mediaPlayer.Controls.currentPosition = (Worksheets(strTeam + " 1").Cells(ActiveCell.row, 2) / 1000) + AudioOffset(i)
    Else
        MsgBox ("No associated Sound file")
    End If
End Sub
Public Sub StopPlaying()
Attribute StopPlaying.VB_ProcData.VB_Invoke_Func = "S\n14"
    Dim mediaPlayer As WindowsMediaPlayer

    Set mediaPlayer = GetMediaPlayer()
    
    If mediaPlayer Is Nothing Then
        ' There's going to be no luck here
        Exit Sub
    End If

    mediaPlayer.Controls.stop
    
End Sub





