Attribute VB_Name = "RodModel2"
Public Const MAXIMUM_TEAMS As Integer = 10
Public Const MAXIMUM_BELLS As Integer = 16

'TODO: Should Offset(-2,2) be Offset(-2,3) - Not sure
'       Range(avgDevTableTL.Offset(-2, 0), accErrTableTL.Offset(-2, 2)).Merge
'        WriteTitle TeamName(i), Range(avgDevTableTL.Offset(-2, 0), accErrTableTL.Offset(-2, 2))

Type Strike
  stroke As String
  bell As String
  time As Long
End Type
Public strikefile(MAXIMUM_TEAMS) As String
Public LoadTime(MAXIMUM_TEAMS, MAXIMUM_BELLS, 6000) As Strike
Public Wholepulls(MAXIMUM_TEAMS, 2 * MAXIMUM_BELLS, 3000) As Strike
'Public CorrectTime(16, 6000) As Strike
'Public OutputArray(96000) As Strike
Public BellPres(10, 16) As Boolean
Public NumRows(10) As Integer
Public NumBells(10) As Integer
Public TeamName(10) As String
Public TimeOrder(16) As Long
Public NumWholepulls(10) As Integer
Public MeanPos(10) As Integer
Public Adjustment(16, 2) As Integer
Public SDPos(10) As Integer
Public AverageSD(10) As Double
Public Best(16) As Double
Public BestTeam(16) As Integer
Public ErrorPos(10) As Integer
Public Mean(10, 2, 16) As Double
Public Mean1(10, 2, 16) As Double
Public Deviation(16, 2) As Double
Public SD(10, 2, 16) As Double
Public SD1(10, 2, 16) As Double
Public TotSD(16) As Double
Public TotSDTeam(16) As Integer
Public Errors(10, 2, 16) As Double
Public Errors1(10, 2, 16) As Double
Public Error50(10, 2, 16) As Integer
Public Error501(10, 2, 16) As Integer
Public Ratio(10, 16) As Double
Public StartAnalysis(10) As Integer
Public EndAnalysis(10) As Integer
Public touchfile As String
Public touchblow(100000) As String
Public touchrung(10, 100000) As String
Public TouchLength As Integer
Public soundfile(10) As String
Public AudioOffset(10) As Double
Public faultThreshold(10) As Double

Public timings(16, 6000) As Long
Public AverageRowGap As Integer
Public AverageBellGap As Integer
Public Strikes As Integer

Public Function TotalTeams() As Integer
    Dim countTeams As Integer
    Dim i As Integer

    countTeams = 0
    For i = 1 To MAXIMUM_TEAMS
        If IsTeamProcessed(i) Then
            countTeams = countTeams + 1
        End If
    Next
    TotalTeams = countTeams
End Function


Public Function getTeamModelSheet1(ByVal teamIndex As Integer, Optional createIfNotExist As Boolean = False, Optional programmaticLabel) As Worksheet

    If Not IsTeamProcessed(teamIndex) Then
        Set getTeamModelSheet1 = Nothing
    Else
    
        Dim sheetName As String
        sheetName = TeamName(teamIndex) + " 1"
        
        If IsMissing(programmaticLabel) Then
            Set getTeamModelSheet1 = FindWorksheet(sheetName:=sheetName, createIfNotExist:=createIfNotExist)
        Else
            Set getTeamModelSheet1 = FindWorksheet(sheetName:=sheetName, createIfNotExist:=createIfNotExist, programmaticLabel:=programmaticLabel)
        End If
    End If
End Function
Public Function getTeamModelSheet2(ByVal teamIndex As Integer, Optional createIfNotExist As Boolean = False, Optional programmaticLabel) As Worksheet

    If Not IsTeamProcessed(teamIndex) Then
        Set getTeamModelSheet2 = Nothing
    Else
    
        Dim sheetName As String
        sheetName = TeamName(teamIndex) + " 2"
        
        If IsMissing(programmaticLabel) Then
            Set getTeamModelSheet2 = FindWorksheet(sheetName:=sheetName, createIfNotExist:=createIfNotExist)
        Else
            Set getTeamModelSheet2 = FindWorksheet(sheetName:=sheetName, createIfNotExist:=createIfNotExist, programmaticLabel:=programmaticLabel)
        End If
    End If
End Function

Public Function IsTeamProcessed(ByVal i As Integer) As Boolean
    Dim FrontSheet As Worksheet
    Set FrontSheet = Worksheets("Front")

    If FrontSheet.Range("TeamNameLabel").Offset(i, 0).Value <> "" And _
        FrontSheet.Range("StrikeFileLabel").Offset(i, 0).Value <> "" And _
        UCase(FrontSheet.Range("AnalyseLabel").Offset(i, 0).Value) = "Y" Then
        IsTeamProcessed = True
    Else
        IsTeamProcessed = False
    End If
End Function

'TODO: Remove alpha dependence - should be possible to manage through range references
Function alpha(ByVal colno As Integer) As String
    If colno < 27 Then
        alpha = Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ", colno, 1)
    Else
        alpha = Mid("AAABACADAEAFAGAHAIAJAKALAMANAOAPAQARASATAUAVAWAXAYAZBABBBCBDBEBFBGBHBIBJBKBLBMBNBOBPBQBRBSBTBUBVBWBXBYBZCACBCCCDCECFCGCHCICJCKCLCMCNCOCPCQCRCSCTCUCVCWCXCYCZ", ((colno - 26) * 2) - 1, 2)
    End If
End Function
'TODO: Remove retstr - shouldn't really be required
Function retstr(ByVal x As Integer) As String
    retstr = Mid(Str(x), 2, 9999)
End Function
Function bell_chartonum(ByVal bellch As String) As Integer
    i = InStr("1234567890ETABCD", bellch)
    If i = 0 Then i = InStr("123456789OETABCD", bellch)
    bell_chartonum = i
End Function
Function bell_numtochar(ByVal bell As Integer) As String
  bell_numtochar = Mid("123456789OETABCD", bell, 1)
End Function
Function hextodec(ByVal hextime As String) As Long
    hextodec = CLng("&H" & hextime)
End Function
Sub findfile(ByVal x As Integer, ByVal y As Integer)
ChDir ("C:\")

filetoopen = Application _
      .GetOpenFilename("Strike Text Files (*.*), *.*")
If filetoopen <> False Then Sheets("Front").Cells(x, y) = filetoopen

End Sub


Function histcount(inrange As Range, valcell As Range, span As Integer) As Integer
  Res = 0
  For Each c In valcell
    compval = c.Value
    Exit For
  Next c
  compval = compval - (span / 2)
  comp2val = compval + span
  For Each c In inrange
    If (c.Value >= compval) And (c.Value < comp2val) Then
      Res = Res + 1
    End If
  Next c
  histcount = Res
End Function
Sub loadvars()
Dim HasRun As Boolean
Dim touchptr As Integer

Erase Error50
Erase Error501

Dim FrontSheet As Worksheet
Dim settingsSheet As Worksheet

Dim teamRWPSheet1 As Worksheet
Dim teamRWPSheet2 As Worksheet

Set FrontSheet = Sheets(FRONT_SHEET_NAME)
Set settingsSheet = Sheets(SETTINGS_SHEET_NAME)

strikeBaseDir = FrontSheet.Range("strikeBaseDir").Value


'get filenames from FRONT sheet
For i = 4 To 13

    strikefile(i - 3) = CvtToAbsFile(JoinPath(strikeBaseDir, FrontSheet.Cells(i, 3)))
    TeamName(i - 3) = FrontSheet.Cells(i, 2)
    soundfile(i - 3) = CvtToAbsFile(FrontSheet.Cells(i, 5))
    AudioOffset(i - 3) = 0
    
    
    'lets see if the model has been run or not
    
    Set teamRWPSheet1 = getTeamModelSheet1(i - 3)
    Set teamRWPSheet2 = getTeamModelSheet2(i - 3)
    HasRun = Not (teamRWPSheet2 Is Nothing)
    
    
    'if it has run then scan the team sheets to populate numwholepulls and numbells
    If HasRun Then
       'find numwholepulls
       For j = 1 To 20000
         If teamRWPSheet2.Cells(j, 1) = "MEAN" Then Exit For
       Next j
       NumWholepulls(i - 3) = j - 3
       MeanPos(i - 3) = j
       SDPos(i - 3) = j + 1
       ErrorPos(i - 3) = j + 7
       
       'find numbells
       For j = 1 To 40
         If teamRWPSheet2.Cells(4, j) = "" Then Exit For
       Next j
       NumBells(i - 3) = (j - 2) / 2
       
       'load Mean, SD and Error arrays - and also average deviation offset (currently hard-coded - DJP)
       
       
        For j = 1 To NumBells(i - 3)
        Mean(i - 3, 1, j) = teamRWPSheet2.Cells(MeanPos(i - 3), 4 * NumBells(i - 3) + 6 + j)
        Mean(i - 3, 2, j) = teamRWPSheet2.Cells(MeanPos(i - 3), 5 * NumBells(i - 3) + 6 + j)
        Mean1(i - 3, 1, j) = teamRWPSheet2.Cells(MeanPos(i - 3), 1 + j)
        Mean1(i - 3, 2, j) = teamRWPSheet2.Cells(MeanPos(i - 3), NumBells(i - 3) + 1 + j)
        Next j
        
        For j = 1 To NumBells(i - 3)
        SD(i - 3, 1, j) = teamRWPSheet2.Cells(SDPos(i - 3), 4 * NumBells(i - 3) + 6 + j)
        SD(i - 3, 2, j) = teamRWPSheet2.Cells(SDPos(i - 3), 5 * NumBells(i - 3) + 6 + j)
        SD1(i - 3, 1, j) = teamRWPSheet2.Cells(SDPos(i - 3), 1 + j)
        SD1(i - 3, 2, j) = teamRWPSheet2.Cells(SDPos(i - 3), NumBells(i - 3) + 1 + j)
        Next j
        
        For j = 1 To NumBells(i - 3)
        Errors(i - 3, 1, j) = teamRWPSheet2.Cells(ErrorPos(i - 3), 4 * NumBells(i - 3) + 6 + j)
        Errors(i - 3, 2, j) = teamRWPSheet2.Cells(ErrorPos(i - 3), 5 * NumBells(i - 3) + 6 + j)
        Errors1(i - 3, 1, j) = teamRWPSheet2.Cells(ErrorPos(i - 3), 1 + j)
        Errors1(i - 3, 2, j) = teamRWPSheet2.Cells(ErrorPos(i - 3), NumBells(i - 3) + 1 + j)
        Next j
        
        avgeInterval = teamRWPSheet1.Cells(NumWholepulls(i - 3) + 5, 27)
        faultThreshold(i - 3) = settingsSheet.Range("optAnalysisFaultPct").Value * avgeInterval
        
        For j = 1 To NumWholepulls(i - 3)
            For k = 1 To NumBells(i - 3)
                
                If teamRWPSheet2.Cells(j + 3, 4 * NumBells(i - 3) + 6 + k) < -faultThreshold(i - 3) Then
                Error50(i - 3, 1, k) = Error50(i - 3, 1, k) + 1
                End If
                If teamRWPSheet2.Cells(j + 3, 1 + k) < -faultThreshold(i - 3) Then
                Error501(i - 3, 1, k) = Error501(i - 3, 1, k) + 1
                End If
            
                If teamRWPSheet2.Cells(j + 3, 5 * NumBells(i - 3) + 6 + k) < -faultThreshold(i - 3) Then
                Error50(i - 3, 1, k) = Error50(i - 3, 1, k) + 1
                End If
                If teamRWPSheet2.Cells(j + 3, NumBells(i - 3) + 1 + k) < -faultThreshold(i - 3) Then
                Error501(i - 3, 1, k) = Error501(i - 3, 1, k) + 1
                End If
                
                If teamRWPSheet2.Cells(j + 3, 4 * NumBells(i - 3) + 6 + k) > faultThreshold(i - 3) Then
                Error50(i - 3, 2, k) = Error50(i - 3, 2, k) + 1
                End If
                If teamRWPSheet2.Cells(j + 3, 1 + k) > faultThreshold(i - 3) Then
                Error501(i - 3, 2, k) = Error501(i - 3, 2, k) + 1
                End If
            
                If teamRWPSheet2.Cells(j + 3, 5 * NumBells(i - 3) + 6 + k) > faultThreshold(i - 3) Then
                Error50(i - 3, 2, k) = Error50(i - 3, 2, k) + 1
                End If
                If teamRWPSheet2.Cells(j + 3, NumBells(i - 3) + 1 + k) > faultThreshold(i - 3) Then
                Error501(i - 3, 2, k) = Error501(i - 3, 2, k) + 1
                End If
                
            Next k
        Next j
    End If
Next i


If FrontSheet.Range("touchFile").Value <> "" Then
    touchfile = CvtToAbsFile(FrontSheet.Range("touchFile").Value)
    If settingsSheet.Range("optAnalysisOpeningRounds") Then
        'TODO: Fix this - probably drift the touch data with the opening rounds
        MsgBox "You have a supplied a touch file and also asked to analyse the opening rounds - that isn't typically a very good combination"
    End If
Else
    touchfile = ""
End If
    

'load touchfile if it exists
TouchLength = 0
If touchfile <> "" Then
  If Dir(touchfile) = "" Then
    MsgBox ("Touchfile : " + touchfile + " does not exist")
    Exit Sub
  End If
  touchptr = 1
  Open touchfile For Input As #1
  Do While Not EOF(1)
    Line Input #1, textline
    TouchLength = TouchLength + 1
    For i = 1 To Len(textline)
      If Mid(textline, i, 1) = "0" Then Mid(textline, i, 1) = "O"
      touchblow(touchptr) = UCase(Mid(textline, i, 1))
      touchptr = touchptr + 1
    Next i
  Loop
  Close #1
Else
    lengthEntry = FrontSheet.Range("touchLength").Value

  If lengthEntry <> 0 Then
    TouchLength = lengthEntry
  Else
    MsgBox ("You must have a touchfile or touch length")
    Exit Sub
  End If
End If
End Sub


Sub loadfiles()
Dim bellptr(16) As Integer
Dim i As Integer
Dim timeloop As Long
Dim lasttimeread As Long
Dim thistimeread As Long
Dim thistiming As Long
Dim firsttiming As Long
Dim lasttiming As Long
Dim Strikes As Long
    Dim touchptr As Long
Dim touchrungptr As Long
Dim CR As String
Dim LF As String

Dim wasSuspended As Boolean

Dim ws As Worksheet

Dim topLevelProfiler As CProfiler


CR = Chr(13)
LF = Chr(10)

' Suspend calculation and redraw
Call SuspendCalculationAndRedraw
' Normally we'd set wasSuspended to the function result. However, because this is a top-level call, we want to make sure that we
' don't leave the output suspended when we finish!!!
wasSuspended = False

Call DeleteWorkingSheets

Set topLevelProfiler = New CProfiler

topLevelProfiler.StartProfiling "RWP"

topLevelProfiler.Tock "PrepSheets"


Call loadvars
topLevelProfiler.Tock "LoadVars"


'reset variables
timeloop = 0
lasttimeread = 0
thistimeread = 0
Strikes = 0

'empty the load array
For i = 1 To 10
  For j = 1 To 16
    For k = 1 To 6000
        LoadTime(i, j, k).stroke = ""
        LoadTime(i, j, k).time = 0
    Next k
  Next j
Next i

'empty the touchrung array
For i = 1 To 10
  For touchrungptr = 1 To 100000
    touchrung(i, touchrungptr) = ""
  Next touchrungptr
Next i


'now load the data up into an array
For i = 1 To 10
    If IsTeamProcessed(i) Then
        thistimeread = 0
        lasttimeread = 0
        thistiming = 0
        firsttiming = 0
        lasttiming = 0
        timeloop = 0
        For j = 1 To 16
          bellptr(j) = 0
          BellPres(i, j) = 0
        Next j
        
        'check to see that the file exists
        If Dir(strikefile(i)) = "" Then
          MsgBox ("File : " + strikefile(i) + " does not exist")
          Exit Sub
        End If
        
        'determine recordlength
        'Open strikefile(i) For Input As #1
        'textline = Input(50, #1)
        'j = InStr(UCase(textline), "X")
        'k = InStr(j + 1, UCase(textline), "X")
        'reclen = k - j
        'Close #1
        

        'now process the file
        Open strikefile(i) For Input As #1
        textfile = Input(LOF(1), #1)
        Close #1
        recstart = 1
        touchrungptr = 0
        
        Do While recstart < Len(textfile)
        'find next end of record
            j = InStr(Mid(textfile, recstart, 999999), CR)
            k = InStr(Mid(textfile, recstart, 999999), LF)
        
            If j > 0 Then
                nextrec = recstart + j
                recend = recstart + j - 1
            End If
            If k > j Then
                nextrec = recstart + k
                If j = 0 Then recend = recstart + k - 1
            End If
          
            textline = Mid(textfile, recstart, recend - recstart + 1)
                
            'process the record

            If Mid(textline, 1, 1) <> "#" Then
                stroke = Mid(textline, 1, 1)
                bell = bell_chartonum(Mid(textline, 3, 1))
                bellptr(bell) = bellptr(bell) + 1
      
                thistimeread = hextodec(Mid(textline, 7, 4))
                If thistimeread < lasttimeread Then timeloop = timeloop + 1
                lasttimeread = thistimeread
    
                thistiming = thistimeread + (65536 * timeloop)
                If firsttiming = 0 Then firsttiming = thistiming
                lasttiming = thistiming
        
                LoadTime(i, bell, bellptr(bell)).time = thistiming
                LoadTime(i, bell, bellptr(bell)).stroke = stroke
                LoadTime(i, bell, bellptr(bell)).bell = bell
                    
                touchrung(i, touchrungptr) = Mid(textline, 3, 1)
                touchrungptr = touchrungptr + 1
                'MsgBox ("Team: " + Str(i) + " bell : " + Str(bell) + " row : " + Str(bellptr(bell)) + " time : " + Str(thistiming))
            Else 'process comment
                If Mid(textline, 4, 11) = "FirstBlowMs" Then
                  AudioOffset(i) = Val(Mid(textline, 17, 9))
                End If
            End If
            recstart = nextrec
        Loop
        Close #1

        NumRows(i) = bellptr(bell)

        'mark the bells that have been detected
        NumBells(i) = 0
        For j = 1 To 16
            BellPres(i, j) = False
            For k = 1 To 10
              If LoadTime(i, j, k).time > 0 Then
                BellPres(i, j) = True
                NumBells(i) = NumBells(i) + 1
                Exit For
              End If
            Next k
        Next j
    End If
Next i

'clear up the loaded arrays - get rid of incomplete recordings etc.
'Call cleanup

'setup the other sheets
Dim pres As Boolean
pres = False

oldStatusBar = Application.DisplayStatusBar
Application.DisplayStatusBar = True

topLevelProfiler.Tock "ModelRun"

For i = 1 To 10
  If IsTeamProcessed(i) Then
    pres = True
    
    Application.StatusBar = "Processing file " & i
    
    'find the touch
    Call findtouch(i)
    Call disp(LoadTime, i)
    Call rwp1(i)
    Call RWP2(i)
    ScreenRefresh
  End If
Next i

topLevelProfiler.Tock "ModelFinish"

RemoveProgrammaticSheetsCreatedBy RWP_WORKING_SHEETS

Worksheets("Front").Activate
    
Application.StatusBar = False
Application.DisplayStatusBar = oldStatusBar

'summary sheet
If pres Then
    Call summary
    Worksheets("Summary").Activate
    Cells(1, 1).Activate
    
    If GetPresentationMode() = ContestFeedback Then
        Call team_charts
        Call Load_Flockton_output
        Call Gen_Contest_PDF
    End If
    
End If


' Resume redraw and calculation
ResumeCalculationAndRedraw wasSuspended


topLevelProfiler.EndProfiling displayChart:=False


End Sub


Sub findtouch(i As Integer)
Dim outofpos As Integer

For j = 1 To NumRows(i)
  outofpos = 0
  For k = 1 To NumBells(i) - 1
    If LoadTime(i, k, j).time > LoadTime(i, k + 1, j).time Then outofpos = outofpos + 1
  Next k
  If outofpos > 3 Then
    If Worksheets("Settings").Range("optAnalysisOpeningRounds").Value = True Then
      StartAnalysis(i) = 1
    Else
      StartAnalysis(i) = j - 2
    End If
    EndAnalysis(i) = j + TouchLength + 1
    Exit Sub
  End If
Next j
  
End Sub
Sub cleanup()
Dim i As Integer
Dim j As Integer
Dim clean As Boolean

For i = 1 To 10
  If IsTeamProcessed(i) Then
    clean = False
     Do While Not clean
       'MsgBox ("team " + Str(i) + "numrows= " + Str(NumRows(i)))
       For j = 1 To NumRows(i) 'loop number of rows
         For k = 1 To NumBells(i) 'check all bells present
           If LoadTime(i, k, j).time = 0 Then
             'remove the whole pull
             'MsgBox ("removing team : " + Str(i) + " row : " + Str(j) + " because of bell : " + Str(k))
             Call removewp(i, j)
             GoTo notclean
           End If
         Next k
       Next j
       clean = True
notclean:
     Loop
  End If
Next i
End Sub
Sub removewp(i As Integer, j As Integer)
  For k = 1 To NumBells(i)
    LoadTime(i, k, j).time = 0
    If j Mod 2 = 0 Then
      LoadTime(i, k, j - 1).time = 0
    Else
      LoadTime(i, k, j + 1).time = 0
    End If
  Next k
  NumRows(i) = NumRows(i) - 2
End Sub
Sub disp(ByRef timearray() As Strike, team As Integer)
Dim oi As Integer
Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim touchptr As Long

oi = 0
touchptr = 0

Sheets.Add After:=Sheets(Sheets.Count)
ActiveSheet.Name = TeamName(team)

RegisterProgrammaticSheet TeamName(team), RWP_WORKING_SHEETS

For i = 1 To 16
  If BellPres(team, i) = True Then
    oi = oi + 1
    For j = 1 To NumRows(team)
      'If timearray(team, i, j).time = 0 Then Exit For
      Sheets(TeamName(team)).Cells(j, oi) = timearray(team, i, j).time
    Next j
  End If
Next i

For i = 1 To NumRows(team)
    For j = 1 To NumBells(team)
        Sheets(TeamName(team)).Cells(i, j + NumBells(team) + 2) = touchrung(team, touchptr)
        touchptr = touchptr + 1
    Next j
Next i

'now create the rows rung next to the timings & store them away
x = (NumBells(team) + 3)
For i = x To x + NumBells(team) - 1
    Columns(alpha(i) + ":" + alpha(i)).ColumnWidth = 2.11
Next i
'For i = 1 To NumRows(team)
'  For j = x To x + NumBells(team) - 1
'    Worksheets(TeamName(team)).Cells(i, j) = "=bell_numtochar(MATCH(+SMALL(A" + retstr(i) + ":" + alpha(NumBells(team)) + retstr(i) + "," + retstr(j - x + 1) + "),A" + retstr(i) + ":" + alpha(NumBells(team)) + retstr(i) + ",0))"
'  Next j
'Next i

'store away the bells rung
'For i = 1 To NumRows(team)
'  For j = x To x + NumBells(team) - 1
'    touchrung(team, touchptr) = Worksheets(TeamName(team)).Cells(i, j)
'    touchptr = touchptr + 1
'  Next j
'Next i

End Sub
Sub order(ByRef inparray() As Strike, team As Integer, row As Integer)
Dim pos As Integer
Dim thistime As Long


  For timeptr = 1 To 16
    TimeOrder(timeptr) = 0
  Next timeptr
  
  timeptr = 1
  
  For pos = 1 To 16
    thistime = inparray(team, pos, row).time
    'MsgBox (thistime)
    If thistime <> 0 Then
      For timeptr = 1 To 16
        If TimeOrder(timeptr) = 0 Then
          TimeOrder(timeptr) = thistime
          Exit For
        End If
        If thistime <= TimeOrder(timeptr) Then
          'shuffle down and insert
          If timeptr < 16 Then
            For j = 16 To timeptr + 1 Step -1
                TimeOrder(j) = TimeOrder(j - 1)
            Next j
          End If
          TimeOrder(timeptr) = thistime
          Exit For
        End If
      Next timeptr
    End If
  Next pos
            
        
End Sub
Sub dispwholepulls(ByRef inparray() As Strike, team As Integer, nam As String, xpos As Integer, ypos As Integer)
Dim optr As Integer
Dim j As Integer
Dim lowesttime As Long
Dim belldone(16) As Boolean
Dim i As Integer

Dim workingSheet As Worksheet

optr = ypos
k = xpos


Set workingSheet = getTeamModelSheet1(team)


For i = StartAnalysis(team) To EndAnalysis(team)
  Call order(inparray, team, i)
  For j = 1 To 16
    'MsgBox ("Check " + Str(j))
    If TimeOrder(j) = 0 Then Exit For
    workingSheet.Cells(optr, k) = TimeOrder(j)
    k = k + 1
    If k > (xpos - 1) + NumBells(team) * 2 Then
      k = xpos
      optr = optr + 1
    End If
  Next j
Next i

End Sub

Sub rwp1(team As Integer)
Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim bell1 As Integer
Dim bell2 As Integer
Dim quickval As Integer
Dim slowval As Integer


Const RWP1_ID = "RWP1"

Dim settingsSheet As Worksheet
Set settingsSheet = Sheets(SETTINGS_SHEET_NAME)


Dim teamRWPSheet1 As Worksheet
Set teamRWPSheet1 = getTeamModelSheet1(team, createIfNotExist:=True, programmaticLabel:=RWP1_ID)


Call dispwholepulls(LoadTime, team, " 1", 2, 3)

'add the headers over the wholepulls
For i = 1 To NumBells(team)
  teamRWPSheet1.Cells(1, i + 1) = i
  teamRWPSheet1.Cells(1, i + 1).Interior.ColorIndex = 4
  teamRWPSheet1.Cells(1, NumBells(team) + i + 1) = i
  teamRWPSheet1.Cells(1, NumBells(team) + i + 1).Interior.ColorIndex = 15
  
  teamRWPSheet1.Cells(2, i + 1) = i
  teamRWPSheet1.Cells(2, NumBells(team) + i + 1) = i + NumBells(team)
Next i

'add the whole pull numbers
For i = 3 To NumRows(team)
  If teamRWPSheet1.Cells(i, 2) = "" Then Exit For
  teamRWPSheet1.Cells(i, 1) = i - 2
Next i

'set the number of whole pulls
NumWholepulls(team) = i - 3
'MsgBox ("setting numwholepulls for team : " + Str(team) + " to : " + Str(NumWholepulls(team)) + " numrows : " + Str(NumRows(team)))

'repeat the handstroke lead in a new column at the end
teamRWPSheet1.Cells(1, (NumBells(team) * 2) + 2) = "1H"
For i = 3 To NumRows(team) + 1
  teamRWPSheet1.Cells(i, (NumBells(team) * 2) + 2) = teamRWPSheet1.Cells((i + 1), 2)
Next i
teamRWPSheet1.Cells(NumWholepulls(team) + 4, (NumBells(team) * 2) + 2) = "=AVERAGE(" + alpha((NumBells(team) * 2) + 2) + "4:" + alpha((NumBells(team) * 2) + 2) + retstr(NumWholepulls(team) + 1) + ")"
teamRWPSheet1.Cells(NumWholepulls(team) + 5, (NumBells(team) * 2) + 2) = "=" + alpha((NumBells(team) * 2) + 2) + retstr(NumWholepulls(team) + 4) + "-" + alpha((NumBells(team) * 2) + 1) + retstr(NumWholepulls(team) + 4)

If GetHandstrokeGapMode() = Averages Then
    'calculate the averages
    For i = 2 To ((NumBells(team) * 2) + 1)
      teamRWPSheet1.Cells(NumWholepulls(team) + 4, i) = "=AVERAGE(" + alpha(i) + "4:" + alpha(i) + retstr(NumWholepulls(team) + 1) + ")"
      If i > 2 Then
        teamRWPSheet1.Cells(NumWholepulls(team) + 5, i) = "=" + alpha(i) + retstr(NumWholepulls(team) + 4) + "-" + alpha(i - 1) + retstr(NumWholepulls(team) + 4)
      End If
    Next i
    
    'calculate the mean
    x = (NumBells(team) * 2) + 3
    teamRWPSheet1.Cells(1, x) = "MEAN"
    Range(alpha(x + 1) + "1:" + alpha(x + 1) + "2").Merge
    Range(alpha(x + 1) + "1:" + alpha(x + 1) + "2").WrapText = True
    teamRWPSheet1.Cells(1, x + 1) = "WP AVERAGE"
    For i = 3 To NumWholepulls(team) + 2
      teamRWPSheet1.Cells(i, x) = "=AVERAGE(" + alpha(2) + retstr(i) + ":" + alpha((NumBells(team) * 2) + 1) + retstr(i) + ")"
      If i > 3 And i < (NumWholepulls(team) + 2) Then
        teamRWPSheet1.Cells(i, x + 1) = "=(" + alpha(x) + retstr(i + 1) + "-" + alpha(x) + retstr(i - 1) + ")/2"
      End If
    Next i
    
    'calculate the average gap overall
    teamRWPSheet1.Cells(NumWholepulls(team) + 5, (NumBells(team) * 2) + 3) = "=AVERAGE(C" + retstr(NumWholepulls(team) + 5) + ":" + alpha((NumBells(team) * 2) + 1) + retstr(NumWholepulls(team) + 5) + ")"
    
    'calculate the ratio of the handstroke gap
    teamRWPSheet1.Cells(NumWholepulls(team) + 5, (NumBells(team) * 2) + 4) = "=" + alpha((NumBells(team) * 2) + 2) + retstr(NumWholepulls(team) + 5) + "/" + alpha((NumBells(team) * 2) + 3) + retstr(NumWholepulls(team) + 5)

ElseIf GetHandstrokeGapMode() = MinimumSquaredError Then

    Dim firstWholepull As Range
    Dim lastWholepull As Range
    
    Dim positionAvgs As Range
    Dim positionDiffs As Range
    
    Dim mseTable As Range
    
    Dim meanWholepullCentre As Range


    Set positionAvgs = teamRWPSheet1.Cells(NumWholepulls(team) + 4, 2)
    Set positionDiffs = positionAvgs.Offset(1, 0)
    
    Set firstWholepull = teamRWPSheet1.Cells(4, 2)
    Set lastWholepull = teamRWPSheet1.Cells(NumWholepulls(team) + 1, 2)
    
    'calculate the averages
    For i = 1 To (NumBells(team) * 2) + 1
        positionAvgs.Offset(0, i - 1) = "=AVERAGE(" & firstWholepull.Offset(0, i - 1).Address(ColumnAbsolute:=False) & ":" & lastWholepull.Offset(0, i - 1).Address(ColumnAbsolute:=False) & ")"
        If i >= 2 Then
            positionDiffs.Offset(0, i - 1) = "=" & positionDiffs.Offset(-1, i - 1).Address & " - " & positionDiffs.Offset(-1, i - 2).Address
        End If
    Next i
    
    ' Minimum squared error handstroke gap calculation
    
    'Average centrepoint of wholepull
    Set meanWholepullCentre = positionAvgs.Offset(0, NumBells(team) * 2 + 1)
    meanWholepullCentre = "=AVERAGE(" & Range(positionAvgs, positionAvgs.Offset(0, NumBells(team) * 2 - 1)).Address & ")"
    
    Set mseTable = positionDiffs.Offset(2, 0)
    
    
    Const diffToCentrepointRow = 1
    Const interbellMultipleRow = 2
    Const vectorMultiplyRow = 3
    
    'Set up diffs to average blows
    For i = 1 To (NumBells(team) * 2)
        mseTable.Offset(diffToCentrepointRow, i - 1) = "=" & positionAvgs.Offset(0, i - 1).Address(RowAbsolute:=False) & "-" & meanWholepullCentre.Address
        mseTable.Offset(interbellMultipleRow, i - 1) = (i - 12.5) / 1150
        mseTable.Offset(vectorMultiplyRow, i - 1) = "=" & mseTable.Offset(diffToCentrepointRow, i - 1).Address & "*" & mseTable.Offset(interbellMultipleRow, i - 1).Address
    Next i
    
    mseTable.Offset(diffToCentrepointRow, NumBells(team) * 2) = "Diff"
    'With mseTable.Offset(diffToCentrepointRow, NumBells(team) * 2).AddComment
    '    .Text = "Difference of average position to average of the centre of the wholepull"
    'End With
    mseTable.Offset(interbellMultipleRow, NumBells(team) * 2) = "BellMultiple"
    'With mseTable.Offset(interbellMultipleRow, NumBells(team) * 2).AddComment
    '    .Text = "This is the number of expected interbell gaps expected"
    'End With
    
    Dim averageInterbellGap As Range
    Dim idealAverageLastBlow As Range
    
    Set averageInterbellGap = mseTable.Offset(vectorMultiplyRow + 1, NumBells(team) * 2 - 1)
    averageInterbellGap = "=SUM(" & Range(mseTable.Offset(vectorMultiplyRow, 0), mseTable.Offset(vectorMultiplyRow, NumBells(team) * 2)).Address & ")"
    averageInterbellGap.Offset(0, 1) = "Average interbell gap"
    Set idealAverageLastBlow = mseTable.Offset(vectorMultiplyRow + 2, NumBells(team) * 2 - 1)
    idealAverageLastBlow = "=11.5*" & mseTable.Offset(vectorMultiplyRow + 1, NumBells(team) * 2 - 1).Address & "+" & meanWholepullCentre.Address
    idealAverageLastBlow.Offset(0, 1) = "Ideal last blow place"
    
    positionDiffs.Offset(0, 2 * NumBells(team) + 1) = averageInterbellGap.Value
    
    mseTable = "Minimum-squared error fit calculation"
    Range(mseTable, mseTable.Offset(0, NumBells(team) * 2)).Merge
    
    Dim wholePullCentrePoint As Range
    
    Set wholePullCentrePoint = teamRWPSheet1.Cells(1, (NumBells(team) * 2) + 3)
    
    'calculate the ratio of the handstroke gap
    Dim handstrokeRatio As Range
    Set handstrokeRatio = teamRWPSheet1.Cells(NumWholepulls(team) + 5, (NumBells(team) * 2) + 4)
    handstrokeRatio = "=(" & positionAvgs.Offset(0, 2 * NumBells(team)).Address & "-" & idealAverageLastBlow.Address & ")/" & averageInterbellGap.Address
    
    'calculate the mean
    'TODO: remove x
    x = (NumBells(team) * 2) + 3
    wholePullCentrePoint = "MEAN"
    Range(alpha(x + 1) + "1:" + alpha(x + 1) + "2").Merge
    Range(alpha(x + 1) + "1:" + alpha(x + 1) + "2").WrapText = True
    teamRWPSheet1.Cells(1, x + 1) = "WP AVERAGE"
    For i = 3 To NumWholepulls(team) + 2
      teamRWPSheet1.Cells(i, x) = "=AVERAGE(" + alpha(2) + retstr(i) + ":" + alpha((NumBells(team) * 2) + 1) + retstr(i) + ")"
      If i > 3 And i < (NumWholepulls(team) + 2) Then
        teamRWPSheet1.Cells(i, x + 1) = "=(" + alpha(x) + retstr(i + 1) + "-" + alpha(x) + retstr(i - 1) + ")/2"
      End If
    Next i

Else
    MsgBox "Invalid method of determining handstroke gap has been selected"
    DeleteWorkingSheets
    settingsSheet.Activate
    settingsSheet.Range("optAnalysisHandstrokeGapMethod").Activate
    End
End If

'calculate the intervals
teamRWPSheet1.Cells(1, (NumBells(team) * 2) + 5) = "INTERVAL"
For i = 4 To NumWholepulls(team) + 1
  teamRWPSheet1.Cells(i, (NumBells(team) * 2) + 5) = "=" + alpha((NumBells(team) * 2) + 4) + retstr(i) + "/(" + retstr((NumBells(team) * 2) - 1) + "+" + alpha((NumBells(team) * 2) + 4) + retstr(NumWholepulls(team) + 5) + ")"
Next i
teamRWPSheet1.Cells(NumWholepulls(team) + 4, ((NumBells(team) * 2) + 5)) = "=MIN(" + alpha((NumBells(team) * 2) + 5) + "4:" + alpha((NumBells(team) * 2) + 5) + retstr(NumWholepulls(team) + 1) + ")"
teamRWPSheet1.Cells(NumWholepulls(team) + 5, ((NumBells(team) * 2) + 5)) = "=MAX(" + alpha((NumBells(team) * 2) + 5) + "4:" + alpha((NumBells(team) * 2) + 5) + retstr(NumWholepulls(team) + 1) + ")"

'calculate alternate row length and difference
x = (NumBells(team) * 2) + 6
Range(alpha(x) + "1:" + alpha(x) + "2").Merge
Range(alpha(x) + "1:" + alpha(x) + "2").WrapText = True
teamRWPSheet1.Cells(1, x) = "WP ACTUAL"
For i = 4 To NumWholepulls(team) + 1
  teamRWPSheet1.Cells(i, x) = "=" + alpha(x - 3) + retstr(i) + "-" + alpha(x - 3) + retstr(i - 1)
  If i > 4 Then
    teamRWPSheet1.Cells(i, x + 1) = "=" + alpha(x) + retstr(i) + "-" + alpha(x) + retstr(i - 1)
  End If
Next i
teamRWPSheet1.Cells(NumWholepulls(team) + 4, x) = "=AVERAGE(" + alpha(x) + "4:" + alpha(x) + retstr(NumWholepulls(team) + 1) + ")"
teamRWPSheet1.Cells(NumWholepulls(team) + 4, x + 1) = "=AVERAGE(" + alpha(x + 1) + "5:" + alpha(x + 1) + retstr(NumWholepulls(team) + 1) + ")"
teamRWPSheet1.Cells(NumWholepulls(team) + 5, x) = "=STDEV(" + alpha(x) + "4:" + alpha(x) + retstr(NumWholepulls(team) + 1) + ")"
teamRWPSheet1.Cells(NumWholepulls(team) + 5, x + 1) = "=STDEV(" + alpha(x + 1) + "5:" + alpha(x + 1) + retstr(NumWholepulls(team) + 1) + ")"

'actual lead interval
x = (NumBells(team) * 2) + 9
teamRWPSheet1.Cells(1, x) = "ACTUAL"
teamRWPSheet1.Cells(2, x) = "LEAD"
teamRWPSheet1.Cells(3, x) = "INTERVAL"
For i = 4 To NumWholepulls(team) + 1
  teamRWPSheet1.Cells(i, x) = "=" + alpha((NumBells(team) * 2) + 2) + retstr(i) + "-" + alpha((NumBells(team) * 2) + 1) + retstr(i)
Next i
teamRWPSheet1.Cells(NumWholepulls(team) + 4, x) = "=AVERAGE(" + alpha(x) + "4:" + alpha(x) + retstr(NumWholepulls(team)) + ")"
teamRWPSheet1.Cells(NumWholepulls(team) + 5, x) = "=STDEV(" + alpha(x) + "4:" + alpha(x) + retstr(NumWholepulls(team)) + ")"

'now setup the gap table
x = (NumBells(team) * 2) + 11
For i = x To x + ((NumBells(team) * 2) - 1)
  bell1 = i - x + 1
  bell2 = i - x + 2
  While bell1 > NumBells(team)
    bell1 = bell1 - 12
  Wend
  While bell2 > NumBells(team)
    bell2 = bell2 - 12
  Wend
  teamRWPSheet1.Cells(2, i) = "'" + bell_numtochar(bell1) + "-" + bell_numtochar(bell2)
  For j = 3 To NumWholepulls(team) + 1
    teamRWPSheet1.Cells(j, i) = "=" + alpha(i - x + 3) + retstr(j) + "-" + alpha(i - x + 2) + retstr(j)
  Next j
  teamRWPSheet1.Cells(NumWholepulls(team) + 4, i) = "=AVERAGE(" + alpha(i) + "4:" + alpha(i) + retstr(NumWholepulls(team) + 1) + ")"
  teamRWPSheet1.Cells(NumWholepulls(team) + 5, i) = "=STDEV(" + alpha(i) + "4:" + alpha(i) + retstr(NumWholepulls(team) + 1) + ")"
  If i = x Then teamRWPSheet1.Cells(NumWholepulls(team) + 7, i) = "=COUNTIF(" + alpha(i) + "4:" + alpha(x + ((NumBells(team) * 2) - 2)) + retstr(NumWholepulls(team) + 1) + ",""<125"")"
Next i

'now set all cells to 0 decimal places
Cells.NumberFormat = "0"

'set decimal point on INTERVAL
Columns(alpha((NumBells(team) * 2) + 5)).NumberFormat = "0.0"

Range(alpha((NumBells(team) * 2) + 4) & (NumWholepulls(team) + 5)).NumberFormat = "0.00"
    
With Rows("1:2")
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
End With

' set cell sizes
Cells.Columns.AutoFit

With Columns(alpha((NumBells(team) * 2) + 3))
    .HorizontalAlignment = xlCenter
    .ColumnWidth = 10
    .VerticalAlignment = xlBottom
    .orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
End With

With Columns(alpha((NumBells(team) * 2) + 4))
    .HorizontalAlignment = xlCenter
    .ColumnWidth = 10
    .VerticalAlignment = xlBottom
    .orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
End With

With Columns(alpha((NumBells(team) * 2) + 5))
    .HorizontalAlignment = xlCenter
    .ColumnWidth = 10
    .VerticalAlignment = xlBottom
    .orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
End With

With Columns(alpha((NumBells(team) * 2) + 6))
    .HorizontalAlignment = xlCenter
    .ColumnWidth = 10
    .VerticalAlignment = xlBottom
    .orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
End With

With Columns(alpha((NumBells(team) * 2) + 8))
    .HorizontalAlignment = xlCenter
    .ColumnWidth = 2.5
    .VerticalAlignment = xlBottom
    .orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
End With

With Columns(alpha((NumBells(team) * 2) + 9))
    .HorizontalAlignment = xlCenter
    .ColumnWidth = 10
    .VerticalAlignment = xlBottom
    .orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
End With

With Columns(alpha((NumBells(team) * 2) + 10))
    .HorizontalAlignment = xlCenter
    .ColumnWidth = 2.5
    .VerticalAlignment = xlBottom
    .orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
End With

'now split the window
'Rows("3:3").Select
With ActiveWindow
   .SplitColumn = 1
   .SplitRow = 2
   .FreezePanes = True
End With

End Sub
Sub RWP2(team As Integer)
Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim bell1 As Integer
Dim bell2 As Integer
Dim numbellssingle As Single
Dim touchptr As Integer
Dim touchblowptr As Integer
Dim bellno As Integer
Dim bell As Integer
Dim bellstr As String
Dim pos(2) As Integer

Dim teamRWPSheet1 As Worksheet
Dim teamRWPSheet2 As Worksheet

Const RWP2_ID = "RWP2"

Set teamRWPSheet1 = getTeamModelSheet1(team)

'create sheet
Set teamRWPSheet2 = getTeamModelSheet2(team, createIfNotExist:=True, programmaticLabel:=RWP2_ID)
'FIXME: Is it possible to get away without activating the sheet?
teamRWPSheet2.Activate


'display the faults by position
numbellssingle = NumBells(team)
numbellssingle = ((numbellssingle * 2) + 1) / 2
For i = 2 To (NumBells(team) * 2) + 1
  For j = 4 To NumWholepulls(team) + 1
    teamRWPSheet2.Cells(j, 1) = j - 3
    teamRWPSheet2.Cells(j, i) = "='" + TeamName(team) + " 1" + "'!" + alpha(i) + retstr(j) + "-'" + TeamName(team) + " 1" + "'!$" + alpha((NumBells(team) * 2) + 3) + retstr(j) + "-('" + TeamName(team) + " 1" + "'!" + alpha(i) + "$2-" + Str(numbellssingle) + ")*'" + TeamName(team) + " 1" + "'!$" + alpha((NumBells(team) * 2) + 5) + retstr(j)
  Next j
  bellno = (i - 1) Mod NumBells(team)
  If bellno = 0 Then bellno = NumBells(team)
  teamRWPSheet2.Cells(1, i) = bell_numtochar(bellno)
  teamRWPSheet2.Cells(NumWholepulls(team) + 3, i) = "=AVERAGE(" + alpha(i) + retstr(4) + ":" + alpha(i) + retstr(NumWholepulls(team) + 1) + ")"
  teamRWPSheet2.Cells(NumWholepulls(team) + 4, i) = "=STDEV(" + alpha(i) + retstr(4) + ":" + alpha(i) + retstr(NumWholepulls(team) + 1) + ")"
  If i > 2 Then teamRWPSheet2.Cells(NumWholepulls(team) + 5, i) = "=CORREL(" + alpha(i - 1) + retstr(4) + ":" + alpha(i - 1) + retstr(NumWholepulls(team) + 1) + "," + alpha(i) + retstr(4) + ":" + alpha(i) + retstr(NumWholepulls(team) + 1) + ")"
  teamRWPSheet2.Cells(NumWholepulls(team) + 8, i) = "=SUMIF(" + alpha(i) + "4:" + alpha(i) + retstr(NumWholepulls(team) + 1) + ",""<0"")"
  teamRWPSheet2.Cells(NumWholepulls(team) + 9, i) = "=SUMIF(" + alpha(i) + "4:" + alpha(i) + retstr(NumWholepulls(team) + 1) + ","">0"")"
  teamRWPSheet2.Cells(NumWholepulls(team) + 10, i) = "=-" + alpha(i) + retstr(NumWholepulls(team) + 8) + "+" + alpha(i) + retstr(NumWholepulls(team) + 9)
Next i

'set up headings
teamRWPSheet2.Cells(NumWholepulls(team) + 3, 1) = "MEAN"
teamRWPSheet2.Cells(NumWholepulls(team) + 4, 1) = "SD"
teamRWPSheet2.Cells(NumWholepulls(team) + 5, 1) = "CORREL"
teamRWPSheet2.Cells(NumWholepulls(team) + 7, 1) = "ERRORS"
teamRWPSheet2.Cells(NumWholepulls(team) + 8, 1) = "FAST"
teamRWPSheet2.Cells(NumWholepulls(team) + 9, 1) = "SLOW"
teamRWPSheet2.Cells(NumWholepulls(team) + 10, 1) = "COMB."

'set color headings and alignment
With Range("B1:" + alpha(NumBells(team) + 1) + "1")
    With .Interior
        .ColorIndex = 4
        .Pattern = xlSolid
    End With
    .HorizontalAlignment = xlCenter
End With

With Range(alpha(NumBells(team) + 2) + "1:" + alpha((NumBells(team) * 2) + 1) + "1")
    With .Interior
        .ColorIndex = 6
        .Pattern = xlSolid
    End With
    .HorizontalAlignment = xlCenter
End With

'find and display biggest quick and slow blows
teamRWPSheet2.Cells(NumWholepulls(team) + 4, (NumBells(team) * 2) + 3) = "=MIN(B4:" + alpha((NumBells(team) * 2) + 1) + retstr(NumWholepulls(team) + 1) + ")"
teamRWPSheet2.Cells(NumWholepulls(team) + 5, (NumBells(team) * 2) + 3) = "=MAX(B4:" + alpha((NumBells(team) * 2) + 1) + retstr(NumWholepulls(team) + 1) + ")"

'now do standard deviations by whole pulls
For i = 4 To NumWholepulls(team) + 1
  teamRWPSheet2.Cells(i, (NumBells(team) * 2) + 4) = "=STDEV(B" + retstr(i) + ":" + alpha((NumBells(team) * 2) + 1) + retstr(i) + ")"
Next i
'use the adjusted timings to calculate the DEVSQ
'teamRWPSheet2.Cells(NumWholepulls(team) + 3, (NumBells(team) * 2) + 4) = "=(DEVSQ(B4:" + alpha((NumBells(team) * 2) + 1) + retstr(NumWholepulls(team) + 1) + ")/(" + retstr(NumWholepulls(team) - 2) + "*" + retstr((NumBells(team) * 2) - 1) + "))^0.5"
teamRWPSheet2.Cells(NumWholepulls(team) + 3, (NumBells(team) * 2) + 4) = "=(DEVSQ(" + alpha((NumBells(team) * 4) + 7) + "4:" + alpha((NumBells(team) * 6) + 6) + retstr(NumWholepulls(team) + 1) + ")/(" + retstr(NumWholepulls(team) - 2) + "*" + retstr((NumBells(team) * 2) - 1) + "))^0.5"
teamRWPSheet2.Cells(NumWholepulls(team) + 4, (NumBells(team) * 2) + 4) = "='" + TeamName(team) + " 1'!" + alpha((NumBells(team) * 2) + 7) + retstr(NumWholepulls(team) + 5)
teamRWPSheet2.Cells(NumWholepulls(team) + 5, (NumBells(team) * 2) + 4) = "=if(" + alpha((NumBells(team) * 2) + 4) + retstr(NumWholepulls(team) + 3) + "<" + alpha((NumBells(team) * 2) + 4) + retstr(NumWholepulls(team) + 4) _
            + ",((" + alpha((NumBells(team) * 2) + 4) + retstr(NumWholepulls(team) + 4) + "^2-" + alpha((NumBells(team) * 2) + 4) + retstr(NumWholepulls(team) + 3) + "^2)/4)^0.5,0)"
teamRWPSheet2.Cells(NumWholepulls(team) + 6, (NumBells(team) * 2 + 4)) = "=(" + alpha((NumBells(team) * 2) + 4) + retstr(NumWholepulls(team) + 3) + "^2+" + alpha((NumBells(team) * 2) + 4) + retstr(NumWholepulls(team) + 5) + "^2)^0.5"

'set up the whole pulls rung
touchptr = (StartAnalysis(team) + 1) * NumBells(team)
touchblowptr = 1
For i = 4 To NumWholepulls(team) + 1
  For j = ((NumBells(team) * 2) + 6) To ((NumBells(team) * 4) + 5)
    teamRWPSheet2.Cells(i, j) = touchrung(team, touchptr)
    If touchrung(team, touchptr) <> touchblow(touchblowptr) Then
      'MsgBox ("Wrong blow at : " + Str(touchptr) + " actual = " + touchrung(team, touchptr) + " should be = " + touchblow(touchptr))
      teamRWPSheet2.Cells(i, j).Interior.ColorIndex = 4
    End If
    touchptr = touchptr + 1
    touchblowptr = touchblowptr + 1
  Next j
Next i

For j = 1 To 16
    Adjustment(j, 1) = Worksheets("Front").Cells(j + 27, 2)
    Adjustment(j, 2) = Worksheets("Front").Cells(j + 27, 3)
Next j

multi_strike_errors = 0
warn_on_multi_strike = True
        
'now setup errors by bell
For i = 4 To NumWholepulls(team) + 1
  For j = ((NumBells(team) * 4) + 7) To ((NumBells(team) * 5) + 6)
     bell = ((j - 7) Mod 12) + 1
     bellstr = bell_numtochar(bell)
     
     L = 1
     For k = ((NumBells(team) * 2) + 6) To ((NumBells(team) * 4) + 5)
       If Cells(i, k) = bellstr Then
        If L <= 2 Then
            pos(L) = k - ((NumBells(team) * 2) + 5)
        Else
            If warn_on_multi_strike Then
                Style = vbYesNo + vbCritical + vbDefaultButton2    ' Define buttons.
                title = "Continue warnings?"    ' Define title.

                Msg = "We have a problem. We have more than two strikes for bell " & bell & " in wholepull " & i & " touch " & TeamName(team)
                Response = MsgBox(Msg, Style, title)
                If Response = vbNo Then
                    warn_on_multi_strike = False
                End If
            End If
            
            multi_strike_errors = multi_strike_errors + 1
        End If
         L = L + 1
       End If
     Next k
     'a$ = InputBox("bellstr=" + bellstr + ", pos(1)=" + Str(pos(1)) + ", pos(2)=" + Str(pos(2)))
     'If a$ = "x" Then Stop
     'This bit is hard-coded for 12 bells as a bit of a test. The idea is to subtract an amount for the smaller bells
     'and add an amount in for bigger bells to compensate for the difference between human ear and Hawkear.
     
     If Adjustment(bell, 1) < 0 Then
          teamRWPSheet2.Cells(i, j) = "=OFFSET(A" + retstr(i) + ",0," + Str(pos(1)) + ") +" + retstr(Adjustment(bell, 1))
     Else
          teamRWPSheet2.Cells(i, j) = "=OFFSET(A" + retstr(i) + ",0," + Str(pos(1)) + ") -" + retstr(Adjustment(bell, 1))
     End If
     
     If Adjustment(bell, 2) < 0 Then
          teamRWPSheet2.Cells(i, j + NumBells(team)) = "=OFFSET(A" + retstr(i) + ",0," + Str(pos(2)) + ") +" + retstr(Adjustment(bell, 2))
     Else
         teamRWPSheet2.Cells(i, j + NumBells(team)) = "=OFFSET(A" + retstr(i) + ",0," + Str(pos(2)) + ") -" + retstr(Adjustment(bell, 2))
     End If
     
  Next j
Next i

'recalculate to establish propoer values
ForceRecalculate

stdFaultPct = Worksheets("Settings").Range("optAnalysisFaultPct").Value
leadFaultPct = Worksheets("Settings").Range("optAnalysisQuickHandstrokeLeadPct").Value

'and setup faults in columns
For i = 4 To NumWholepulls(team) + 1
  'teamRWPSheet2.Cells(i, j + NumBells(team) + 1) = "=COUNTIF(B" + retstr(i) + ",""<-" + retstr((teamRWPSheet1.Cells(i, (NumBells(team) * 2) + 5)) * (Worksheets("Front").Cells(18, 2) * Sqr(2)) / 100) + """)"
  'teamRWPSheet2.Cells(i, j + NumBells(team) + 2) = "=COUNTIF(C" + retstr(i) + ":" + alpha(NumBells(team) + 1) + retstr(i) + ",""<-" + retstr((teamRWPSheet1.Cells(i, (NumBells(team) * 2) + 5)) * (Worksheets("Front").Cells(18, 2)) / 100) + """)"
  'teamRWPSheet2.Cells(i, j + NumBells(team) + 3) = "=COUNTIF(B" + retstr(i) + ":" + alpha(NumBells(team) + 1) + retstr(i) + ","">+" + retstr((teamRWPSheet1.Cells(i, (NumBells(team) * 2) + 5)) * (Worksheets("Front").Cells(18, 2)) / 100) + """)"
  'teamRWPSheet2.Cells(i, j + NumBells(team) + 4) = "=COUNTIF(" + alpha(NumBells(team) + 2) + retstr(i) + ":" + alpha((NumBells(team) * 2) + 1) + retstr(i) + ",""<-" + retstr((teamRWPSheet1.Cells(i, (NumBells(team) * 2) + 5)) * (Worksheets("Front").Cells(18, 2)) / 100) + """)"
  'teamRWPSheet2.Cells(i, j + NumBells(team) + 5) = "=COUNTIF(" + alpha(NumBells(team) + 2) + retstr(i) + ":" + alpha((NumBells(team) * 2) + 1) + retstr(i) + ","">+" + retstr((teamRWPSheet1.Cells(i, (NumBells(team) * 2) + 5)) * (Worksheets("Front").Cells(18, 2)) / 100) + """)"

  teamRWPSheet2.Cells(i, j + NumBells(team) + 1) = "=COUNTIF(" + alpha(2) + retstr(i) + ",""<-" + retstr((teamRWPSheet1.Cells(i, (NumBells(team) * 2) + 5)) * leadFaultPct) + """)"
  teamRWPSheet2.Cells(i, j + NumBells(team) + 2) = "=COUNTIF(" + alpha((NumBells(team) * 4) + 7) + retstr(i) + ":" + alpha((NumBells(team) * 5) + 6) + retstr(i) + ",""<-" + retstr((teamRWPSheet1.Cells(i, (NumBells(team) * 2) + 5)) * stdFaultPct) + """)"
  teamRWPSheet2.Cells(i, j + NumBells(team) + 3) = "=COUNTIF(" + alpha((NumBells(team) * 4) + 7) + retstr(i) + ":" + alpha((NumBells(team) * 5) + 6) + retstr(i) + ","">+" + retstr((teamRWPSheet1.Cells(i, (NumBells(team) * 2) + 5)) * stdFaultPct) + """)"
  teamRWPSheet2.Cells(i, j + NumBells(team) + 4) = "=COUNTIF(" + alpha((NumBells(team) * 5) + 7) + retstr(i) + ":" + alpha((NumBells(team) * 6) + 6) + retstr(i) + ",""<-" + retstr((teamRWPSheet1.Cells(i, (NumBells(team) * 2) + 5)) * stdFaultPct) + """)"
  teamRWPSheet2.Cells(i, j + NumBells(team) + 5) = "=COUNTIF(" + alpha((NumBells(team) * 5) + 7) + retstr(i) + ":" + alpha((NumBells(team) * 6) + 6) + retstr(i) + ","">+" + retstr((teamRWPSheet1.Cells(i, (NumBells(team) * 2) + 5)) * stdFaultPct) + """)"
Next i
       

'do the rows underneath
teamRWPSheet2.Cells(NumWholepulls(team) + 3, ((NumBells(team) * 4) + 6)) = "MEAN"
teamRWPSheet2.Cells(NumWholepulls(team) + 4, ((NumBells(team) * 4) + 6)) = "SD"
teamRWPSheet2.Cells(NumWholepulls(team) + 5, ((NumBells(team) * 4) + 6)) = "COMB."
teamRWPSheet2.Cells(NumWholepulls(team) + 7, ((NumBells(team) * 4) + 6)) = "ERRORS"
teamRWPSheet2.Cells(NumWholepulls(team) + 8, ((NumBells(team) * 4) + 6)) = "FAST"
teamRWPSheet2.Cells(NumWholepulls(team) + 9, ((NumBells(team) * 4) + 6)) = "SLOW"
teamRWPSheet2.Cells(NumWholepulls(team) + 10, ((NumBells(team) * 4) + 6)) = "COMB."
teamRWPSheet2.Cells(NumWholepulls(team) + 11, ((NumBells(team) * 4) + 6)) = "OVERALL"
For i = ((NumBells(team) * 4) + 7) To ((NumBells(team) * 6) + 6)
  teamRWPSheet2.Cells(NumWholepulls(team) + 3, i) = "=AVERAGE(" + alpha(i) + retstr(4) + ":" + alpha(i) + retstr(NumWholepulls(team) + 1) + ")"
  teamRWPSheet2.Cells(NumWholepulls(team) + 4, i) = "=STDEV(" + alpha(i) + retstr(4) + ":" + alpha(i) + retstr(NumWholepulls(team) + 1) + ")"
  teamRWPSheet2.Cells(NumWholepulls(team) + 8, i) = "=SUMIF(" + alpha(i) + retstr(4) + ":" + alpha(i) + retstr(NumWholepulls(team) + 1) + ",""<0"")"
  teamRWPSheet2.Cells(NumWholepulls(team) + 9, i) = "=SUMIF(" + alpha(i) + retstr(4) + ":" + alpha(i) + retstr(NumWholepulls(team) + 1) + ","">0"")"
  teamRWPSheet2.Cells(NumWholepulls(team) + 10, i) = "=-" + alpha(i) + retstr(NumWholepulls(team) + 8) + " + " + alpha(i) + retstr(NumWholepulls(team) + 9)
  If i > ((NumBells(team) * 5) + 6) Then
    teamRWPSheet2.Cells(NumWholepulls(team) + 5, i) = "=((" + alpha(i - 12) + retstr(NumWholepulls(team) + 4) + "^2+" + alpha(i) + retstr(NumWholepulls(team) + 4) + "^2)/2)^0.5"
    teamRWPSheet2.Cells(NumWholepulls(team) + 11, i) = "=" + alpha(i - 12) + retstr(NumWholepulls(team) + 10) + "+" + alpha(i) + retstr(NumWholepulls(team) + 10)
  End If
Next i

' Count the faults
For i = ((NumBells(team) * 6) + 8) To ((NumBells(team) * 6) + 12)
  teamRWPSheet2.Cells(NumWholepulls(team) + 3, i) = "=SUM(" + alpha(i) + retstr(4) + ":" + alpha(i) + retstr(NumWholepulls(team) + 1) + ")"
Next i

'set up column headings
For i = ((NumBells(team) * 4) + 7) To ((NumBells(team) * 6) + 6)
    bellno = (i - (NumBells(team) * 4) + 6) Mod NumBells(team)
    If bellno = 0 Then bellno = NumBells(team)
    teamRWPSheet2.Cells(1, i) = bell_numtochar(bellno)
Next i
teamRWPSheet2.Cells(1, (NumBells(team) * 6) + 8) = "QHL"
teamRWPSheet2.Cells(1, (NumBells(team) * 6) + 9) = "QH"
teamRWPSheet2.Cells(1, (NumBells(team) * 6) + 10) = "SH"
teamRWPSheet2.Cells(1, (NumBells(team) * 6) + 11) = "QB"
teamRWPSheet2.Cells(1, (NumBells(team) * 6) + 12) = "SB"

'set color headings and alignment
i = ((NumBells(team) * 4) + 7)
With Range(alpha(i) + "1:" + alpha(i + (NumBells(team) - 1)) + "1")
    With .Interior
        .ColorIndex = 4
        .Pattern = xlSolid
    End With
    .HorizontalAlignment = xlCenter
End With

With Range(alpha(i + NumBells(team)) + "1:" + alpha(i + ((NumBells(team) * 2) - 1)) + "1")
    With .Interior
        .ColorIndex = 6
        .Pattern = xlSolid
    End With
    .HorizontalAlignment = xlCenter
End With

With Range(alpha(NumBells(team) + 2) + "1:" + alpha((NumBells(team) * 2) + 1) + "1")
    With .Interior
        .ColorIndex = 6
        .Pattern = xlSolid
    End With
    .HorizontalAlignment = xlCenter
End With

'now set all cells to 0 decimal places
Cells.NumberFormat = "0"

'set format
'set decimal point on INTERVAL
Rows(NumWholepulls(team) + 5).NumberFormat = "0.0"
Cells(NumWholepulls(team) + 5, (NumBells(team) * 2) + 3).NumberFormat = "0"


'set cell size
'autofit columns
Cells.Columns.AutoFit

'align wholepulls
    With Range(alpha((NumBells(team) * 2) + 6) + "4:" + alpha((NumBells(team) * 4) + 6) + retstr(NumWholepulls(team) + 3))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

'now split the window
'Rows("3:3").Select
With ActiveWindow
   .SplitColumn = 1
   .SplitRow = 2
   .FreezePanes = True
End With

End Sub


Sub summary()
Dim team As Integer
Dim rowptr As Integer
Dim colptr As Integer
Dim i As Integer
Dim pealspeed As Long
Dim pealspeedh As Integer
Dim pealspeedm As Integer
Dim x As Integer
Dim y As Integer

Dim teamSummaryTC As Range
Dim teamSummaryBC As Range


Dim mainSummaryTableTL As Range
Dim extSummaryTableTL As Range


Dim SummaryWorksheet As Worksheet


'create sheet
Sheets.Add After:=Sheets(Sheets.Count)
ActiveSheet.Name = "SUMMARY"

Set SummaryWorksheet = Worksheets("SUMMARY")

SummaryWorksheet.Activate

'get heading
SummaryWorksheet.Cells(1, 1) = Worksheets("FRONT").Cells(1, 1)


'All cells should default to 0 decimal places
Cells.NumberFormat = "0"

Set mainSummaryTableTL = SummaryWorksheet.Cells(4, 1)

'display averages and time
rowptr = 0
colptr = 0
mainSummaryTableTL.Offset(rowptr, 0) = "AVERAGE INTERVAL (ms)"
rowptr = rowptr + 1
mainSummaryTableTL.Offset(rowptr, 0) = "AVERAGE H/S LEAD INTERVAL (ms)"
rowptr = rowptr + 1
mainSummaryTableTL.Offset(rowptr, 0) = "PEAL SPEED (5040)"
rowptr = rowptr + 1
mainSummaryTableTL.Offset(rowptr, 0) = "OVERALL SD"
rowptr = rowptr + 1
mainSummaryTableTL.Offset(rowptr, 0) = "RATING (SD as % of Interval)"
rowptr = rowptr + 1
mainSummaryTableTL.Offset(rowptr, 0) = "PERCENTAGE for Presentation"

If TotalTeams() > 1 Then
    rowptr = rowptr + 2
    mainSummaryTableTL.Offset(rowptr, 0) = "PLACING"
End If



Dim sdRangeLeft As Range
Dim sdRangeRight As Range

Dim intergapCell As Range
Dim handstrokegapCell As Range
Dim sdCell As Range
Dim adjustedsdCell As Range


Set presentationRangeLeft = Nothing

sdrowptr = 0
colptr = 1
For team = 1 To MAXIMUM_TEAMS

    If IsTeamProcessed(team) Then
        rowptr = 0
    
        mainSummaryTableTL.Offset(rowptr - 1, colptr) = TeamName(team)
        
        mainSummaryTableTL.Offset(rowptr, colptr) = "='" + TeamName(team) + " 1'!" + alpha((NumBells(team) * 2) + 3) + retstr(NumWholepulls(team) + 5)
        Set intergapCell = mainSummaryTableTL.Offset(rowptr, colptr)
        rowptr = rowptr + 1
        mainSummaryTableTL.Offset(rowptr, colptr) = "='" + TeamName(team) + " 1'!" + alpha((NumBells(team) * 2) + 9) + retstr(NumWholepulls(team) + 4)
        Set handstrokegapCell = mainSummaryTableTL.Offset(rowptr, colptr)
        
        rowptr = rowptr + 1
        mainSummaryTableTL.Offset(rowptr, colptr) = "=SecondsToHoursAndMinutes((" + intergapCell.Address + "*23+" + handstrokegapCell.Address + ")*5040/(2*1000))"
        rowptr = rowptr + 1
        mainSummaryTableTL.Offset(rowptr, colptr) = "='" + TeamName(team) + " 2'!" + alpha((NumBells(team) * 2) + 4) + retstr(NumWholepulls(team) + 6)
        Set sdCell = mainSummaryTableTL.Offset(rowptr, colptr)
        sdCell.NumberFormat = "0.0"
        rowptr = rowptr + 1
        mainSummaryTableTL.Offset(rowptr, colptr) = "=" + sdCell.Address + "/" + intergapCell.Address
        sdrowptr = rowptr
        Set adjustedsdCell = mainSummaryTableTL.Offset(rowptr, colptr)
        adjustedsdCell.NumberFormat = "0.0%"
        
        If sdRangeLeft Is Nothing Then
            Set sdRangeLeft = mainSummaryTableTL.Offset(rowptr, colptr)
        End If
        Set sdRangeRight = mainSummaryTableTL.Offset(rowptr, colptr)
        
        rowptr = rowptr + 1
        
        mainSummaryTableTL.Offset(rowptr, colptr) = "=1-4*(1-NORMSDIST(0.25/" + adjustedsdCell.Address + "))"
        mainSummaryTableTL.Offset(rowptr, colptr).NumberFormat = "0.0%"
        rowptr = rowptr + 1
        
        colptr = colptr + 1
    End If
Next team

If TotalTeams() > 1 Then
    For i = 1 To TotalTeams()
      mainSummaryTableTL.Offset(rowptr + 1, i) = "=RANK(" + mainSummaryTableTL.Offset(sdrowptr, i).Address(ColumnAbsolute:=False) + "," + Range(sdRangeLeft, sdRangeRight).Address + ",1)"
    Next i
End If

'reportHandstrokeLeads = Worksheets("Settings").Range("

Set extSummaryTableTL = mainSummaryTableTL.Offset(rowptr + 4, 0)

'setup judging by faults
rowptr = 0

If Not Worksheets("Settings").Range("optDisplayHideQuickLeads") Then
    extSummaryTableTL.Offset(rowptr, 0) = "QUICK HANDSTROKE LEADS"
    rowptr = rowptr + 1
End If
extSummaryTableTL.Offset(rowptr, 0) = "QUICK HANDSTROKES"
rowptr = rowptr + 1
extSummaryTableTL.Offset(rowptr, 0) = "SLOW HANDSTROKES"
rowptr = rowptr + 1
extSummaryTableTL.Offset(rowptr, 0) = "QUICK BACKSTROKES"
rowptr = rowptr + 1
extSummaryTableTL.Offset(rowptr, 0) = "SLOW BACKSTROKES"
rowptr = rowptr + 1
extSummaryTableTL.Offset(rowptr, 0) = "TOTAL"

If TotalTeams() > 1 Then
    rowptr = rowptr + 1
    extSummaryTableTL.Offset(rowptr, 0) = "PLACING"
End If

colptr = 1
For team = 1 To 10
    If IsTeamProcessed(team) Then
        rowptr = 0
        Set teamSummaryTC = extSummaryTableTL.Offset(rowptr, colptr)

        
        If Not Worksheets("Settings").Range("optDisplayHideQuickLeads") Then
            extSummaryTableTL.Offset(rowptr, colptr) = "='" + TeamName(team) + " 2'!" + alpha((NumBells(team) * 6) + 8) + retstr(NumWholepulls(team) + 3)
            rowptr = rowptr + 1
        End If
        extSummaryTableTL.Offset(rowptr, colptr) = "='" + TeamName(team) + " 2'!" + alpha((NumBells(team) * 6) + 9) + retstr(NumWholepulls(team) + 3)
        rowptr = rowptr + 1
        extSummaryTableTL.Offset(rowptr, colptr) = "='" + TeamName(team) + " 2'!" + alpha((NumBells(team) * 6) + 10) + retstr(NumWholepulls(team) + 3)
        rowptr = rowptr + 1
        extSummaryTableTL.Offset(rowptr, colptr) = "='" + TeamName(team) + " 2'!" + alpha((NumBells(team) * 6) + 11) + retstr(NumWholepulls(team) + 3)
        rowptr = rowptr + 1
        extSummaryTableTL.Offset(rowptr, colptr) = "='" + TeamName(team) + " 2'!" + alpha((NumBells(team) * 6) + 12) + retstr(NumWholepulls(team) + 3)
        

        Set teamSummaryBC = extSummaryTableTL.Offset(rowptr, colptr)
        
        rowptr = rowptr + 1
        extSummaryTableTL.Offset(rowptr, colptr) = "=SUM(" & Range(teamSummaryTC, teamSummaryBC).Address(ColumnAbsolute:=False) & ")"
        rowptr = rowptr + 1
        colptr = colptr + 1
    End If
Next team

If TotalTeams() > 1 Then
    Dim totalsRange
    Set totalsRange = Range(extSummaryTableTL.Offset(rowptr - 1, 1), extSummaryTableTL.Offset(rowptr - 1, colptr - 1))
    For i = 1 To colptr - 1
      extSummaryTableTL.Offset(rowptr, i) = "=RANK(" & extSummaryTableTL.Offset(rowptr - 1, i).Address(ColumnAbsolute:=False) & "," & totalsRange.Address & ",1)"
    Next i
End If


'center the columns
With Range("B1:N50")
    .HorizontalAlignment = xlCenter
End With


'set cell size
Cells.Columns.AutoFit
'autofit columns

'set columns the same size
Dim MaxWidth As Double
Dim RR As Range
Dim R As Range
Set RR = Range("B1:N50")
For Each R In RR
    If R.ColumnWidth > MaxWidth Then
        MaxWidth = R.ColumnWidth
    End If
Next R
RR.EntireColumn.ColumnWidth = MaxWidth


    AddMacroButton ButtonLabel:="Error Histograms", LinkedMacro:="graph_histograms", TargetWorksheet:=SummaryWorksheet, FitCell:=Cells(22, 1), RowHeightMultiplier:=1.5
    AddMacroButton ButtonLabel:="WholePull SD graphs", LinkedMacro:="graph_wholepullsd", TargetWorksheet:=SummaryWorksheet, FitCell:=Cells(24, 1), RowHeightMultiplier:=1.5
    AddMacroButton ButtonLabel:="Speed graphs", LinkedMacro:="graph_speed", TargetWorksheet:=SummaryWorksheet, FitCell:=Cells(26, 1), RowHeightMultiplier:=1.5
    AddMacroButton ButtonLabel:="Bell Analysis", LinkedMacro:="bell_table", TargetWorksheet:=SummaryWorksheet, FitCell:=Cells(28, 1), RowHeightMultiplier:=1.5
    AddMacroButton ButtonLabel:="Create Charts", LinkedMacro:="team_charts", TargetWorksheet:=SummaryWorksheet, FitCell:=Cells(30, 1), RowHeightMultiplier:=1.5
    AddMacroButton ButtonLabel:="Touch Summary", LinkedMacro:="team_charts1", TargetWorksheet:=SummaryWorksheet, FitCell:=Cells(32, 1), RowHeightMultiplier:=1.5
    AddMacroButton ButtonLabel:="Average Deviations", LinkedMacro:="average_deviations", TargetWorksheet:=SummaryWorksheet, FitCell:=Cells(34, 1), RowHeightMultiplier:=1.5
    
    AddMacroButton ButtonLabel:="Write Toast", LinkedMacro:="Gen_XML", TargetWorksheet:=SummaryWorksheet, FitCell:=Cells(36, 1), RowHeightMultiplier:=1.5
    
    Cells(1, 1).Select
    
End Sub

Sub AddMacroButton(ByVal ButtonLabel As String, ByVal LinkedMacro As String, TargetWorksheet As Worksheet, FitCell As Range, Optional RowHeightMultiplier)
    Dim workingButton As Button
    
    If Not IsMissing(RowHeightMultiplier) Then
        FitCell.RowHeight = RowHeightMultiplier * FitCell.RowHeight
    End If
    
    Set workingButton = TargetWorksheet.Buttons.Add(FitCell.Left, FitCell.Top, FitCell.Width, FitCell.RowHeight)
    workingButton.Characters.Text = ButtonLabel
    workingButton.OnAction = LinkedMacro
End Sub

Sub graph_histograms()
Dim teamcol As Integer
Dim span As Integer
Dim team As Integer
Dim alertStatus As Boolean
Dim calcWasSuspended As Boolean


'loadvars
Call loadvars


calcWasSuspended = SuspendCalculationAndRedraw()
alertStatus = SuspendApplicationAlerts()

'if it already exists then delete the HISTOGRAMS worksheet
For Each ws In Worksheets
    If ws.Name = "HISTOGRAMS" Then Worksheets(ws.Name).Delete
Next ws

'now create a new sheet
Sheets.Add After:=Sheets(Sheets.Count)
ActiveSheet.Name = "HISTOGRAMS"
Worksheets("HISTOGRAMS").Activate

' set teamnames up
teamcol = 2
For team = 1 To 10
  If IsTeamProcessed(team) Then
    Worksheets("HISTOGRAMS").Cells(1, teamcol) = TeamName(team)
    teamcol = teamcol + 1
  End If
Next team

bin = -200

For i = 2 To 42
  If i > 2 Then
    If i < 42 Then
      Worksheets("HISTOGRAMS").Cells(i, 1) = bin
      span = 10
    Else
      Worksheets("HISTOGRAMS").Cells(i, 1) = 500
      span = 620
    End If
  Else
    Worksheets("HISTOGRAMS").Cells(i, 1) = -500
    span = 620
  End If

  bin = bin + 10
  teamcol = 2
  For team = 1 To 10
    If IsTeamProcessed(team) Then
      Worksheets("HISTOGRAMS").Cells(i, teamcol) = histcount(teamRWPSheet2.Range(alpha((NumBells(team) * 4) + 7) + "4:" + alpha((NumBells(team) * 6) + 6) + retstr(NumWholepulls(team) + 1)), Worksheets("HISTOGRAMS").Cells(i, 1), span)
      teamcol = teamcol + 1
    End If
  Next team
Next i
Worksheets("HISTOGRAMS").Cells(2, 1) = "Other"
Worksheets("HISTOGRAMS").Cells(42, 1) = "Other"

'now do the actual histograms
teamcol = 2
For team = 1 To 10
  If IsTeamProcessed(team) Then
    ' FIXME: Can this select be avoided??
    
    Worksheets("HISTOGRAMS").Cells((teamcol - 1) * 10, 12).Select
    Charts.Add
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.SetSourceData Source:=Sheets("HISTOGRAMS").Range(alpha(teamcol) + "2:" + alpha(teamcol) + "42"), _
        PlotBy:=xlColumns
    ActiveChart.SeriesCollection(1).XValues = "=HISTOGRAMS!R2C1:R42C1"
    ActiveChart.Location Where:=xlLocationAsObject, Name:="HISTOGRAMS"
    With ActiveChart
        .HasTitle = True
        .ChartTitle.Characters.Text = TeamName(team)
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "ms"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "x"
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).MaximumScale = 600
    End With
    'ActiveChart.ChartArea.Select
    'ActiveSheet.Shapes("Chart " + retstr(teamcol - 1)).ScaleHeight 1.37, msoFalse, msoScaleFromTopLeft
    'ActiveSheet.Shapes("Chart " + retstr(teamcol - 1)).ScaleWidth 1.13, msoFalse, _
    '    msoScaleFromBottomRight
    ActiveSheet.Shapes("Chart " + retstr(teamcol - 1)).Top = (teamcol - 2) * 200
    ActiveSheet.Shapes("Chart " + retstr(teamcol - 1)).Left = 250
    
    teamcol = teamcol + 1
  End If
Next team

ResumeCalculationAndRedraw calcWasSuspended
ResumeApplicationAlerts alertStatus

End Sub
Sub WriteTitle(title As String, cell As Range)
    cell = title
    With cell
        .HorizontalAlignment = xlCenter
        With .Font
            .Color = RGB(&H39, &H99, &H99)
            .Italic = True
            .Bold = True
            .Size = 14
        End With
    End With
End Sub
Sub team_charts()
'Dim teamcol As Integer
'Dim span As Integer
Dim workingShape As Shape
Dim workingChart As Chart
Dim workingSheet As Worksheet

Dim avgDevTableTL As Range
Dim stdDevTableTL As Range
Dim accErrTableTL As Range
Dim faultTableTL As Range
Dim sheetExtent As Range


Dim i As Integer

Dim wasSuspended As Boolean
Const TEAM_CHART_ID = "TeamCharts"

'loadvars
Call loadvars

' Suspend calculation and redraw
Call SuspendCalculationAndRedraw
wasSuspended = False

RemoveProgrammaticSheetsCreatedBy TEAM_CHART_ID

embedCharts = (GetPresentationMode() = General)

'now create a new sheet for each team
For i = 1 To 10
    If IsTeamProcessed(i) Then
        Set workingSheet = Sheets.Add(After:=Sheets(Sheets.Count))
        workingSheet.Name = TeamName(i) + " Charts"
        RegisterProgrammaticSheet workingSheet.Name, TEAM_CHART_ID
        workingSheet.Activate
                
        Set avgDevTableTL = workingSheet.Cells(3, 1)
        Set stdDevTableTL = workingSheet.Cells(23, 1)
        Set accErrTableTL = workingSheet.Cells(3, 7)
        Set faultTableTL = workingSheet.Cells(23, 7)
        
        
        If embedCharts Then
            Range(avgDevTableTL.Offset(-2, 0), accErrTableTL.Offset(-2, 4)).Merge
            WriteTitle TeamName(i), Range(avgDevTableTL.Offset(-2, 0), accErrTableTL.Offset(-2, 4))
        Else
            Range(avgDevTableTL.Offset(-2, 0), accErrTableTL.Offset(-2, 2)).Merge
            WriteTitle TeamName(i), Range(avgDevTableTL.Offset(-2, 0), accErrTableTL.Offset(-2, 2))
        End If
        
        'create a chart of the average deviation data
        
        avgDevTableTL.Offset(0, 0) = "Bell"
        avgDevTableTL.Offset(0, 1) = "Handstroke"
        avgDevTableTL.Offset(0, 2) = "Backstroke"
        
        Range(avgDevTableTL.Offset(-1, 0), avgDevTableTL.Offset(-1, 2)).Merge
        WriteTitle "Average Deviation", avgDevTableTL.Offset(-1, 0)
        
        If embedCharts Then
            avgDevTableTL.Offset(0, 3).EntireColumn.ColumnWidth = 2
            avgDevTableTL.Offset(0, 4).EntireColumn.ColumnWidth = 75
            avgDevTableTL.Offset(0, 5).EntireColumn.ColumnWidth = 2
        End If
        
        For j = 1 To NumBells(i)
            avgDevTableTL.Offset(j, 0) = j
            avgDevTableTL.Offset(j, 1) = Mean(i, 1, j)
            avgDevTableTL.Offset(j, 2) = Mean(i, 2, j)
        Next j
        
        With Range(avgDevTableTL, avgDevTableTL.Offset(0, 2))
            .Font.Bold = True
            .EntireColumn.AutoFit
        End With
        
        If embedCharts Then
            Set workingShape = ActiveSheet.Shapes.AddChart
            Set workingChart = workingShape.Chart
        Else
            Set workingChart = Charts.Add
            If Charts.Count > 1 Then
                workingChart.Move After:=Charts(Charts.Count)
            End If
            workingChart.Name = TeamName(i) & " AD Chart"
            RegisterProgrammaticSheet workingChart.Name, TEAM_CHART_ID
        End If
        
        Range(avgDevTableTL.Offset(0, 1), avgDevTableTL.Offset(NumBells(i), 2)).NumberFormat = "0.0"
                
        workingChart.ChartType = xlColumnClustered
        workingChart.SetSourceData Source:=Range(avgDevTableTL.Offset(0, 1), avgDevTableTL.Offset(NumBells(i), 2)), PlotBy:=xlColumns
            
        With workingChart
            .HasTitle = True
            .ChartTitle.Characters.Text = TeamName(i) & " - Average Deviation"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Early --- ms --- Late"
        End With
        
        If embedCharts Then
            With workingShape
                .Top = avgDevTableTL.EntireRow.Top
                .Left = avgDevTableTL.Offset(0, 4).EntireColumn.Left
                .Width = avgDevTableTL.Offset(0, 5).EntireColumn.Left - avgDevTableTL.Offset(0, 4).EntireColumn.Left
            End With
        End If
    
        'create a chart of the standard deviation data
        
        stdDevTableTL.Offset(0, 0) = "Bell"
        stdDevTableTL.Offset(0, 1) = "Handstroke"
        stdDevTableTL.Offset(0, 2) = "Backstroke"
        
        Range(stdDevTableTL.Offset(-1, 0), stdDevTableTL.Offset(-1, 2)).Merge
        WriteTitle "Standard Deviation", stdDevTableTL.Offset(-1, 0)
        
        If embedCharts Then
            stdDevTableTL.Offset(0, 3).EntireColumn.ColumnWidth = 2
            stdDevTableTL.Offset(0, 4).EntireColumn.ColumnWidth = 75
            stdDevTableTL.Offset(0, 5).EntireColumn.ColumnWidth = 2
        End If
        
        k = 999
        L = 999
        m = 999
        
        For j = 1 To NumBells(i)
            stdDevTableTL.Offset(j, 0) = j
            stdDevTableTL.Offset(j, 1) = SD(i, 1, j)
            If SD(i, 1, j) < k Then
                k = SD(i, 1, j)
                x = j
            End If
            stdDevTableTL.Offset(j, 2) = SD(i, 2, j)
            If SD(i, 2, j) < L Then
                L = SD(i, 2, j)
                y = j
            End If
            If SD(i, 2, j) + SD(i, 1, j) < m Then
                m = SD(i, 2, j) + SD(i, 1, j)
                Z = j
            End If
        Next j
        
        If Worksheets("Settings").Range("optDisplayHighlighting").Value Then
            With Cells(x + 19, 2).Interior
                .ColorIndex = 4
                .Pattern = xlSolid
            End With
            With Cells(y + 19, 3).Interior
                .ColorIndex = 4
                .Pattern = xlSolid
            End With
            With Cells(Z + 19, 1).Interior
                .ColorIndex = 4
                .Pattern = xlSolid
            End With
        End If
        
        With Range(stdDevTableTL, stdDevTableTL.Offset(0, 2))
            .Font.Bold = True
            .EntireColumn.AutoFit
        End With
        
        If embedCharts Then
            Set workingShape = ActiveSheet.Shapes.AddChart
            Set workingChart = workingShape.Chart
        Else
            Set workingChart = Charts.Add(After:=Charts(Charts.Count))
            workingChart.Name = TeamName(i) & " SD Chart"
            RegisterProgrammaticSheet workingChart.Name, TEAM_CHART_ID
        End If
        
        Range(stdDevTableTL.Offset(0, 1), stdDevTableTL.Offset(NumBells(i), 2)).NumberFormat = "0.0"
        
        workingChart.ChartType = xlColumnClustered
        workingChart.SetSourceData Source:=Range(stdDevTableTL.Offset(0, 1), stdDevTableTL.Offset(NumBells(i), 2)), PlotBy:=xlColumns
        workingChart.Axes(xlValue).MinimumScale = 15
            
        With workingChart
            .HasTitle = True
            .ChartTitle.Characters.Text = TeamName(i) & " - Standard Deviation"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "ms"
       
        End With
        
        If embedCharts Then
            With workingShape
                .Top = stdDevTableTL.EntireRow.Top
                .Left = stdDevTableTL.Offset(0, 4).EntireColumn.Left
                .Width = stdDevTableTL.Offset(0, 5).EntireColumn.Left - stdDevTableTL.Offset(0, 4).EntireColumn.Left
            End With
        End If
      
          'create a chart of the total error data
        
        Range(accErrTableTL.Offset(-1, 0), accErrTableTL.Offset(-1, 2)).Merge
        WriteTitle "Accumulated error", accErrTableTL.Offset(-1, 0)
        
        accErrTableTL.Offset(0, 0) = "Bell"
        accErrTableTL.Offset(0, 1) = "Handstroke"
        accErrTableTL.Offset(0, 2) = "Backstroke"
        
        If embedCharts Then
            accErrTableTL.Offset(0, 3).EntireColumn.ColumnWidth = 2
            accErrTableTL.Offset(0, 4).EntireColumn.ColumnWidth = 75
            accErrTableTL.Offset(0, 5).EntireColumn.ColumnWidth = 2
        End If
        
        For j = 1 To NumBells(i)
            accErrTableTL.Offset(j, 0) = j
            accErrTableTL.Offset(j, 1) = Errors(i, 1, j)
            accErrTableTL.Offset(j, 2) = Errors(i, 2, j)
        Next j
        
        
        If embedCharts Then
            Set workingShape = ActiveSheet.Shapes.AddChart
            Set workingChart = workingShape.Chart
        Else
            Set workingChart = Charts.Add(After:=Charts(Charts.Count))
            workingChart.Name = TeamName(i) & " AE Chart"
            RegisterProgrammaticSheet workingChart.Name, TEAM_CHART_ID
        End If
        
        Range(accErrTableTL.Offset(0, 1), accErrTableTL.Offset(NumBells(i), 2)).NumberFormat = "0"
        
        With Range(accErrTableTL, accErrTableTL.Offset(0, 2))
            .Font.Bold = True
            .EntireColumn.AutoFit
        End With
        
        
        workingChart.ChartType = xlColumnClustered
        workingChart.SetSourceData Source:=Range(accErrTableTL.Offset(0, 1), accErrTableTL.Offset(NumBells(i), 2)), _
            PlotBy:=xlColumns
            
        With workingChart
            .HasTitle = True
            .ChartTitle.Characters.Text = TeamName(i) & " - Accumulated Error"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "ms"
       
        End With
        
        If embedCharts Then
            With workingShape
                .Top = accErrTableTL.EntireRow.Top
                .Left = accErrTableTL.Offset(0, 4).EntireColumn.Left
                .Width = accErrTableTL.Offset(0, 5).EntireColumn.Left - accErrTableTL.Offset(0, 4).EntireColumn.Left
            End With
        End If
    
    
        Range(faultTableTL.Offset(-1, 0), faultTableTL.Offset(-1, 2)).Merge
        WriteTitle "Faults >" & CStr(Round(faultThreshold(i))) & "ms", faultTableTL.Offset(-1, 0)
                
        'create a chart of errors > 50ms
        faultTableTL.Offset(0, 0) = "Bell"
        faultTableTL.Offset(0, 1) = "Quick"
        faultTableTL.Offset(0, 2) = "Slow"
            
        If embedCharts Then
            faultTableTL.Offset(0, 3).EntireColumn.ColumnWidth = 2
            faultTableTL.Offset(0, 4).EntireColumn.ColumnWidth = 75
            faultTableTL.Offset(0, 5).EntireColumn.ColumnWidth = 2
        End If
        
        k = 999
        x = 0
        y = 0
        
        
        For j = 1 To NumBells(i)
            faultTableTL.Offset(j, 0) = j
            faultTableTL.Offset(j, 1) = Error50(i, 1, j)
            faultTableTL.Offset(j, 2) = Error50(i, 2, j)
            If Error50(i, 2, j) + Error50(i, 1, j) = k And x = 0 Then
                x = j
            End If
            If Error50(i, 2, j) + Error50(i, 1, j) < k Then
                k = Error50(i, 2, j) + Error50(i, 1, j)
                Z = j
                x = 0
            End If
        Next j
        
        If Worksheets("Settings").Range("optDisplayHighlighting").Value Then
            With Cells(Z + 19, 12).Interior
            .ColorIndex = 4
            .Pattern = xlSolid
            End With
            
            If x <> 0 Then
                With Cells(x + 19, 12).Interior
                .ColorIndex = 4
                .Pattern = xlSolid
            End With
            End If
        End If
        
        With Range(faultTableTL, faultTableTL.Offset(0, 2))
            .Font.Bold = True
            .EntireColumn.AutoFit
        End With
        
        If embedCharts Then
            Set workingShape = ActiveSheet.Shapes.AddChart
            Set workingChart = workingShape.Chart
        Else
            Set workingChart = Charts.Add(After:=Charts(Charts.Count))
            workingChart.Name = TeamName(i) & " Fault Chart"
            RegisterProgrammaticSheet workingChart.Name, TEAM_CHART_ID
        End If
        
        Range(faultTableTL.Offset(0, 1), faultTableTL.Offset(NumBells(i), 2)).NumberFormat = "0"
        workingChart.ChartType = xlColumnClustered
        workingChart.SetSourceData Source:=Range(faultTableTL.Offset(0, 1), faultTableTL.Offset(NumBells(i), 2)), _
            PlotBy:=xlColumns
            
        With workingChart
            .HasTitle = True
            .ChartTitle.Characters.Text = TeamName(i) & " - Errors > " + CStr(Round(faultThreshold(i))) + " ms"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Count"
        End With
        
        If embedCharts Then
            With workingShape
                .Top = faultTableTL.EntireRow.Top
                .Left = faultTableTL.Offset(0, 4).EntireColumn.Left
                .Width = faultTableTL.Offset(0, 5).EntireColumn.Left - faultTableTL.Offset(0, 4).EntireColumn.Left
            End With
        End If
        Set sheetExtent = Range(workingSheet.Cells(1, 1), workingSheet.Cells(faultTableTL.row + MAXIMUM_BELLS, faultTableTL.Column + 5))
        
        SetupSinglePagePrint workingSheet, sheetExtent, xlLandscape, cFooter:=TeamName(i)
        
    End If
Next i

Worksheets("Summary").Activate
    Cells(1, 1).Activate
    
    
    ResumeCalculationAndRedraw wasSuspended
End Sub

Sub average_deviations()
    'Dim teamcol As Integer
    'Dim span As Integer
    Dim i As Integer
    Dim teamsSummed As Integer
    Dim bellsInAnalysis As Integer
    Dim outputSheet As Worksheet
    Dim outputChart As Chart
    Dim chartContainer As Shape
    
    Dim outputTableTL As Range
    Dim alertStatus As Boolean
    
    'loadvars
    Call loadvars
    
    alertStatus = SuspendApplicationAlerts()
    
    'if it already exists then delete the existing worksheet
    For Each ws In Worksheets
            If ws.Name = "Average Deviations" Then Worksheets(ws.Name).Delete
    Next ws
    
    For j = 1 To 16
        Deviation(j, 1) = 0
        Deviation(j, 2) = 0
    Next j

    'store the lowest SDs for each bell across all teams
    
    For i = 1 To 10
        If IsTeamProcessed(i) Then
            If bellsInAnalysis = 0 Then
                bellsInAnalysis = NumBells(i)
            Else
                If bellsInAnalysis <> NumBells(i) Then
                    MsgBox "Whoa hold on there - not all of the teams rang the same number of bells. Problems started with band " & i
                    Exit Sub
                End If
            End If
            
            teamsSummed = teamsSummed + 1
            For j = 1 To NumBells(i)
                Deviation(j, 1) = Deviation(j, 1) + Mean(i, 1, j)
                Deviation(j, 2) = Deviation(j, 2) + Mean(i, 2, j)
            Next j
        End If
    Next i
    
    
    If teamsSummed <= 0 Then
        MsgBox "Yeah ... to do that, there will need to be at least one band in the analysis"
        Exit Sub
    End If
    
    'Create a new sheet
    Set outputSheet = Sheets.Add(After:=Sheets(Sheets.Count))
    outputSheet.Name = "Average Deviations"
    
    'output the values to the sheet
    
    'create a chart of the average deviation data
    
    Set outputTableTL = outputSheet.Cells(1, 1)
    
    outputTableTL.Offset(0, 0) = "Bell"
    outputTableTL.Offset(0, 1) = "Handstroke"
    outputTableTL.Offset(0, 2) = "Backstroke"
       
    For j = 1 To bellsInAnalysis
        outputTableTL.Offset(j, 0) = j
        outputTableTL.Offset(j, 1) = Deviation(j, 1) / teamsSummed
        outputTableTL.Offset(j, 2) = Deviation(j, 2) / teamsSummed
    Next j
    
    Set chartContainer = outputSheet.Shapes.AddChart
    Set outputChart = chartContainer.Chart
    outputChart.ChartType = xlColumnClustered
    outputChart.SetSourceData Source:=outputSheet.Range(outputTableTL.Offset(0, 1), outputTableTL.Offset(bellsInAnalysis, 2)), _
        PlotBy:=xlColumns
        
    With outputChart
        .HasTitle = True
        .ChartTitle.Characters.Text = "Average Deviation"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "ms"
    End With
    
    With chartContainer
        .Top = i - 1 * 200
        .Left = 150
    End With
        
    outputSheet.Range(outputTableTL, outputTableTL.Offset(0, 2)).Font.Bold = True
    outputSheet.Range(outputTableTL.Offset(1, 1), outputTableTL.Offset(bellsInAnalysis + 1, 2)).NumberFormat = "0.0"
    
    outputSheet.Range(outputTableTL, outputTableTL.Offset(0, 2)).EntireColumn.AutoFit
    
    ResumeApplicationAlerts alertStatus
End Sub

Sub bell_analysis(team As Integer)
  MsgBox ("Performing bell analysis for team " + TeamName(team) + " - NOT YET IMPLEMENTED")
End Sub

Sub team_charts1()
'Dim teamcol As Integer
'Dim span As Integer


Const TEAM_CHART_BY_POSITION_ID = "TeamChartsByPosition"

Dim workingShape As Shape
Dim workingChart As Chart
Dim workingSheet As Worksheet

Dim avgDevTableTL As Range
Dim stdDevTableTL As Range
Dim accErrTableTL As Range
Dim faultTableTL As Range
Dim sheetExtent As Range



Dim i As Integer
Dim alertStatus  As Boolean
Dim calcWasSuspended As Boolean

'loadvars
Call loadvars

calcWasSuspended = SuspendCalculationAndRedraw()
alertStatus = SuspendApplicationAlerts()


RemoveProgrammaticSheetsCreatedBy TEAM_CHART_BY_POSITION_ID

embedCharts = (GetPresentationMode() = General)

'now create a new sheet for each team
For i = 1 To 10
    If IsTeamProcessed(i) Then
            
        Set workingSheet = Sheets.Add(After:=Sheets(Sheets.Count))
        workingSheet.Name = TeamName(i) + " Touch Summary"
        RegisterProgrammaticSheet workingSheet.Name, TEAM_CHART_BY_POSITION_ID
        workingSheet.Activate
        
        
        Set avgDevTableTL = workingSheet.Cells(3, 1)
        Set stdDevTableTL = workingSheet.Cells(23, 1)
        Set accErrTableTL = workingSheet.Cells(3, 7)
        Set faultTableTL = workingSheet.Cells(23, 7)
        
        'create a chart of the average deviation data
        
        avgDevTableTL.Offset(0, 0) = "Bell"
        avgDevTableTL.Offset(0, 1) = "Handstroke"
        avgDevTableTL.Offset(0, 2) = "Backstroke"
        
        
        If embedCharts Then
            Range(avgDevTableTL.Offset(-2, 0), accErrTableTL.Offset(-2, 4)).Merge
            WriteTitle TeamName(i), Range(avgDevTableTL.Offset(-2, 0), accErrTableTL.Offset(-2, 4))
        Else
            Range(avgDevTableTL.Offset(-2, 0), accErrTableTL.Offset(-2, 2)).Merge
            WriteTitle TeamName(i), Range(avgDevTableTL.Offset(-2, 0), accErrTableTL.Offset(-2, 2))
        End If
        
        Range(avgDevTableTL.Offset(-1, 0), avgDevTableTL.Offset(-1, 2)).Merge
        WriteTitle "Average Deviation", avgDevTableTL.Offset(-1, 0)
        
        
        If embedCharts Then
            avgDevTableTL.Offset(0, 3).EntireColumn.ColumnWidth = 2
            avgDevTableTL.Offset(0, 4).EntireColumn.ColumnWidth = 75
            avgDevTableTL.Offset(0, 5).EntireColumn.ColumnWidth = 2
        End If
        
        For j = 1 To NumBells(i)
            avgDevTableTL.Offset(j, 0) = j
            avgDevTableTL.Offset(j, 1) = Mean1(i, 1, j)
            avgDevTableTL.Offset(j, 2) = Mean1(i, 2, j)
        Next j
        
        With Range(avgDevTableTL, avgDevTableTL.Offset(0, 2))
            .Font.Bold = True
            .EntireColumn.AutoFit
        End With
        
        
        If embedCharts Then
            Set workingShape = ActiveSheet.Shapes.AddChart
            Set workingChart = workingShape.Chart
        Else
            Set workingChart = Charts.Add
            If Charts.Count > 1 Then
                workingChart.Move After:=Charts(Charts.Count)
            End If
            workingChart.Name = TeamName(i) & " AD Chart by pos"
            RegisterProgrammaticSheet workingChart.Name, TEAM_CHART_BY_POSITION_ID
        End If
        
        workingChart.ChartType = xlColumnClustered
        workingChart.SetSourceData Source:=Range(avgDevTableTL.Offset(0, 1), avgDevTableTL.Offset(NumBells(i), 2)), PlotBy:=xlColumns
            
        With workingChart
            .HasTitle = True
            .ChartTitle.Characters.Text = "Average Deviation by position in row"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Early --- ms --- Late"
        End With
        
        
        If embedCharts Then
            With workingShape
                .Top = avgDevTableTL.EntireRow.Top
                .Left = avgDevTableTL.Offset(0, 4).EntireColumn.Left
                .Width = avgDevTableTL.Offset(0, 5).EntireColumn.Left - avgDevTableTL.Offset(0, 4).EntireColumn.Left
            End With
        End If
        
    
        'create a chart of the standard deviation data
        
        stdDevTableTL.Offset(0, 0) = "Bell"
        stdDevTableTL.Offset(0, 1) = "Handstroke"
        stdDevTableTL.Offset(0, 2) = "Backstroke"
        
        
        Range(stdDevTableTL.Offset(-1, 0), stdDevTableTL.Offset(-1, 2)).Merge
        WriteTitle "Standard Deviation", stdDevTableTL.Offset(-1, 0)
    
        k = 999
        L = 999
        m = 999
        
        For j = 1 To NumBells(i)
            stdDevTableTL.Offset(j, 0) = j
            stdDevTableTL.Offset(j, 1) = SD1(i, 1, j)
            If SD1(i, 1, j) < k Then
                k = SD1(i, 1, j)
                x = j
            End If
            stdDevTableTL.Offset(j, 2) = SD1(i, 2, j)
            If SD1(i, 2, j) < L Then
                L = SD1(i, 2, j)
                y = j
            End If
            If SD1(i, 2, j) + SD1(i, 1, j) < m Then
                m = SD1(i, 2, j) + SD1(i, 1, j)
                Z = j
            End If
        Next j
                
        With Range(stdDevTableTL, stdDevTableTL.Offset(0, 2))
            .Font.Bold = True
            .EntireColumn.AutoFit
        End With
        
        If embedCharts Then
            Set workingShape = ActiveSheet.Shapes.AddChart
            Set workingChart = workingShape.Chart
        Else
            Set workingChart = Charts.Add(After:=Charts(Charts.Count))
            workingChart.Name = TeamName(i) & " SD Chart by pos"
            RegisterProgrammaticSheet workingChart.Name, TEAM_CHART_BY_POSITION_ID
        End If
        
        Range(stdDevTableTL.Offset(0, 1), stdDevTableTL.Offset(NumBells(i), 2)).NumberFormat = "0.0"
        
        workingChart.ChartType = xlColumnClustered
        workingChart.SetSourceData Source:=Range(stdDevTableTL.Offset(0, 1), stdDevTableTL.Offset(NumBells(i), 2)), PlotBy:=xlColumns
        workingChart.Axes(xlValue).MinimumScale = 15
            
        With workingChart
            .HasTitle = True
            .ChartTitle.Characters.Text = TeamName(i) & " - Standard Deviation by position in row"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "ms"
       
        End With
        
        If embedCharts Then
            With workingShape
                .Top = stdDevTableTL.EntireRow.Top
                .Left = stdDevTableTL.Offset(0, 4).EntireColumn.Left
                .Width = stdDevTableTL.Offset(0, 5).EntireColumn.Left - stdDevTableTL.Offset(0, 4).EntireColumn.Left
            End With
        End If
      
          'create a chart of the total error data
        
        Range(accErrTableTL.Offset(-1, 0), accErrTableTL.Offset(-1, 2)).Merge
        WriteTitle "Accumulated error", accErrTableTL.Offset(-1, 0)
        
        accErrTableTL.Offset(0, 0) = "Bell"
        accErrTableTL.Offset(0, 1) = "Handstroke"
        accErrTableTL.Offset(0, 2) = "Backstroke"
        
        If embedCharts Then
            accErrTableTL.Offset(0, 3).EntireColumn.ColumnWidth = 2
            accErrTableTL.Offset(0, 4).EntireColumn.ColumnWidth = 75
            accErrTableTL.Offset(0, 5).EntireColumn.ColumnWidth = 2
        End If
        
        For j = 1 To NumBells(i)
            accErrTableTL.Offset(j, 0) = j
            accErrTableTL.Offset(j, 1) = Errors1(i, 1, j)
            accErrTableTL.Offset(j, 2) = Errors1(i, 2, j)
        Next j
        
        
        If embedCharts Then
            Set workingShape = ActiveSheet.Shapes.AddChart
            Set workingChart = workingShape.Chart
        Else
            Set workingChart = Charts.Add(After:=Charts(Charts.Count))
            workingChart.Name = TeamName(i) & " AE Chart by pos"
            RegisterProgrammaticSheet workingChart.Name, TEAM_CHART_BY_POSITION_ID
        End If
        
        Range(accErrTableTL.Offset(0, 1), accErrTableTL.Offset(NumBells(i), 2)).NumberFormat = "0"
        
        With Range(accErrTableTL, accErrTableTL.Offset(0, 2))
            .Font.Bold = True
            .EntireColumn.AutoFit
        End With
        
        
        workingChart.ChartType = xlColumnClustered
        workingChart.SetSourceData Source:=Range(accErrTableTL.Offset(0, 1), accErrTableTL.Offset(NumBells(i), 2)), _
            PlotBy:=xlColumns
            
        With workingChart
            .HasTitle = True
            .ChartTitle.Characters.Text = TeamName(i) & " - Accumulated Error"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "ms"
       
        End With
        
        If embedCharts Then
            With workingShape
                .Top = accErrTableTL.EntireRow.Top
                .Left = accErrTableTL.Offset(0, 4).EntireColumn.Left
                .Width = accErrTableTL.Offset(0, 5).EntireColumn.Left - accErrTableTL.Offset(0, 4).EntireColumn.Left
            End With
        End If
      
        Range(faultTableTL.Offset(-1, 0), faultTableTL.Offset(-1, 2)).Merge
        WriteTitle "Faults >" & CStr(Round(faultThreshold(i))) & "ms", faultTableTL.Offset(-1, 0)
        'create a chart of errors > 50ms
        faultTableTL.Offset(0, 0) = "Bell"
        faultTableTL.Offset(0, 1) = "Quick"
        faultTableTL.Offset(0, 2) = "Slow"
            
        If embedCharts Then
            faultTableTL.Offset(0, 3).EntireColumn.ColumnWidth = 2
            faultTableTL.Offset(0, 4).EntireColumn.ColumnWidth = 75
            faultTableTL.Offset(0, 5).EntireColumn.ColumnWidth = 2
        End If
    
        k = 999
        x = 0
        y = 0
        
        
        For j = 1 To NumBells(i)
            faultTableTL.Offset(j, 0) = j
            faultTableTL.Offset(j, 1) = Error501(i, 1, j)
            faultTableTL.Offset(j, 2) = Error501(i, 2, j)
            If Error501(i, 2, j) + Error501(i, 1, j) = k And x = 0 Then
                x = j
            End If
            If Error501(i, 2, j) + Error501(i, 1, j) < k Then
                k = Error501(i, 2, j) + Error501(i, 1, j)
                Z = j
                x = 0
            End If
        Next j
        
        
        With Range(faultTableTL, faultTableTL.Offset(0, 2))
            .Font.Bold = True
            .EntireColumn.AutoFit
        End With
        
        If embedCharts Then
            Set workingShape = ActiveSheet.Shapes.AddChart
            Set workingChart = workingShape.Chart
        Else
            Set workingChart = Charts.Add(After:=Charts(Charts.Count))
            workingChart.Name = TeamName(i) & " Fault Chart by pos"
            RegisterProgrammaticSheet workingChart.Name, TEAM_CHART_BY_POSITION_ID
        End If
        
        Range(faultTableTL.Offset(0, 1), faultTableTL.Offset(NumBells(i), 2)).NumberFormat = "0"
        workingChart.ChartType = xlColumnClustered
        workingChart.SetSourceData Source:=Range(faultTableTL.Offset(0, 1), faultTableTL.Offset(NumBells(i), 2)), _
            PlotBy:=xlColumns
            
        With workingChart
            .HasTitle = True
            .ChartTitle.Characters.Text = TeamName(i) & " - Errors > " + CStr(Round(faultThreshold(i))) + " ms"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Count"
        End With
        
        If embedCharts Then
            With workingShape
                .Top = faultTableTL.EntireRow.Top
                .Left = faultTableTL.Offset(0, 4).EntireColumn.Left
                .Width = faultTableTL.Offset(0, 5).EntireColumn.Left - faultTableTL.Offset(0, 4).EntireColumn.Left
            End With
        End If
        Set sheetExtent = Range(workingSheet.Cells(1, 1), workingSheet.Cells(faultTableTL.row + MAXIMUM_BELLS, faultTableTL.Column + 5))

        
        Columns("B:C").NumberFormat = "0.0"
        Columns("M:N").NumberFormat = "0"
      
      Cells(1, 1).Select
      
      
        SetupSinglePagePrint workingSheet, sheetExtent, xlLandscape, cFooter:=TeamName(i)
    End If
Next i

Worksheets("Summary").Activate
    Cells(1, 1).Activate

ResumeCalculationAndRedraw calcWasSuspended
ResumeApplicationAlerts alertStatus
End Sub

Sub Load_Flockton_output()
    Dim insertionPoint As Worksheet
    Dim workingSheet As Worksheet
        
    Dim FrontSheet As Worksheet
    Set FrontSheet = Sheets("Front")
    
    Const FLOCKTON_ID = "Flockton"
    
    RemoveProgrammaticSheetsCreatedBy FLOCKTON_ID
    
    flocktonBaseDir = FrontSheet.Range("flocktonBaseDir").Value

    If WorksheetExists("SUMMARY") Then
        Set insertionPoint = Worksheets("SUMMARY")
    Else
        Set insertionPoint = Worksheets(Worksheets.Count)
    End If
    
    For i = 1 To MAXIMUM_TEAMS
        If IsTeamProcessed(i) Then
            file_partial = GetFileTitle(FrontSheet.Range("StrikeFileLabel").Offset(i, 0))
            targetFilename = CvtToAbsFile(JoinPath(flocktonBaseDir, file_partial & "_summary.txt"))
            
            If Dir(targetFilename) = "" Then
                MsgBox "Cannot load Flockton output. Tried to load:" & targetFilename
            Else
        
                Set workingSheet = Worksheets.Add(After:=insertionPoint)
                workingSheet.Name = "Flockton " & TeamName(i)
                
                RegisterProgrammaticSheet workingSheet.Name, FLOCKTON_ID
                workingSheet.Cells.Font.Name = "Consolas"
                
                workingSheet.Cells(1, 1) = "Flockton output for " & TeamName(i)
                workingSheet.Cells(1, 1).Font.Name = "Adobe Gothic Std B"
                workingSheet.Cells(1, 1).Font.Size = 12
                workingSheet.Cells(1, 1).HorizontalAlignment = xlCenter
                
                rowptr = 2
                fhndl = FreeFile
                Open targetFilename For Input As fhndl
                    Do
                        Line Input #fhndl, nextLine
                        workingSheet.Cells(rowptr, 1) = nextLine
                    Loop While Not EOF(fhndl)
                Close fhndl
                
                workingSheet.Columns(1).AutoFit
                
                Set insertionPoint = workingSheet
            End If
        End If
    Next
End Sub
