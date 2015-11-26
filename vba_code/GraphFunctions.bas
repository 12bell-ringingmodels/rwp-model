Attribute VB_Name = "GraphFunctions"
Sub graph_wholepullsd()
Dim team As Integer
Dim alertStatus As Boolean


' load vars
Call loadvars

alertStatus = SuspendApplicationAlerts()


'if it already exists then delete the WHOLEPULLS SD worksheet
For Each ws In Worksheets
    If ws.Name = "WHOLEPULLS SD" Then Worksheets(ws.Name).Delete
Next ws

'now create a new sheet
Sheets.Add After:=Sheets(Sheets.Count)
ActiveSheet.Name = "WHOLEPULLS SD"
Worksheets("WHOLEPULLS SD").Activate

teamcol = 2
For team = 1 To 10
  If IsTeamProcessed(team) Then
    Charts.Add
    ActiveChart.ChartType = xlLineMarkers
    ActiveChart.SetSourceData Source:=Sheets(TeamName(team) + " 2").Range(alpha((NumBells(team) * 2) + 4) + "4:" + alpha((NumBells(team) * 2) + 4) + retstr(NumWholepulls(team) + 1)), PlotBy:=xlColumns
    ActiveChart.Location Where:=xlLocationAsObject, Name:="WHOLEPULLS SD"
    With ActiveChart
        .HasTitle = True
        .ChartTitle.Characters.Text = "SD by WholePull - " + TeamName(team)
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Whole Pull"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "SD"
        .Axes(xlValue).MinimumScale = 10
        .Axes(xlValue).MaximumScale = 70
    End With
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Chart " + retstr(teamcol - 1)).ScaleHeight 1.5, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("Chart " + retstr(teamcol - 1)).ScaleWidth 1.5, msoFalse, _
        msoScaleFromBottomRight
    ActiveSheet.Shapes("Chart " + retstr(teamcol - 1)).Top = (teamcol - 2) * 300
    ActiveSheet.Shapes("Chart " + retstr(teamcol - 1)).Left = 0
    teamcol = teamcol + 1
  End If
Next team

ResumeApplicationAlerts alertStatus

End Sub
Sub graph_speed()

Dim lowy As Integer
Dim highy As Integer
Dim working As Double
Dim team As Integer
Dim alertStatus As Boolean

'load vars
Call loadvars
alertStatus = SuspendApplicationAlerts()

'if it already exists then delete the SPEED worksheet
For Each ws In Worksheets
    If ws.Name = "SPEED" Then Worksheets(ws.Name).Delete
Next ws

'now create a new sheet
Sheets.Add After:=Sheets(Sheets.Count)
ActiveSheet.Name = "SPEED"
Worksheets("SPEED").Activate

highy = 0
lowy = 999

'now determine minimum and maximum speeds
For team = 1 To 10
  If IsTeamProcessed(team) Then
    working = Sheets(TeamName(team) + " 1").Cells((NumWholepulls(team) + 4), ((NumBells(team) * 2) + 5))
    If working < lowy Then
      lowy = Int(working)
    End If
    working = Sheets(TeamName(team) + " 1").Cells((NumWholepulls(team) + 5), ((NumBells(team) * 2) + 5))
    If working > highy Then
      highy = Int(working + 1)
    End If
  End If
Next team

teamcol = 2
For team = 1 To 10
  If IsTeamProcessed(team) Then
    Charts.Add
    ActiveChart.ChartType = xlLineMarkers
    ActiveChart.SetSourceData Source:=Sheets(TeamName(team) + " 1").Range(alpha((NumBells(team) * 2) + 5) + "4:" + alpha((NumBells(team) * 2) + 5) + retstr(NumWholepulls(team) + 1)), PlotBy:=xlColumns
    ActiveChart.Location Where:=xlLocationAsObject, Name:="SPEED"
    With ActiveChart
        .HasTitle = True
        .ChartTitle.Characters.Text = "Speed by WholePull - " + TeamName(team)
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Interbell Gap (ms)"
        .Axes(xlValue).MinimumScale = lowy
        .Axes(xlValue).MaximumScale = highy
        '.Axes(x1Value).MajorUnit = 8
        '.Axes(xlValue).MinorUnit = 10
    End With
    
    ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
    ActiveChart.HasAxis(xlValue) = True
    ActiveSheet.ChartObjects("Chart 1").Activate
    ' ActiveChart.Axes(xlValue).Select
    ActiveChart.HasAxis(xlCategory) = True
    ActiveSheet.ChartObjects("Chart 1").Activate
    ' ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).TickLabelSpacing = 12
    
    
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Chart " + retstr(teamcol - 1)).ScaleWidth 1.8, msoFalse, _
        msoScaleFromBottomRight
    ActiveSheet.Shapes("Chart " + retstr(teamcol - 1)).Top = (teamcol - 2) * 220
    ActiveSheet.Shapes("Chart " + retstr(teamcol - 1)).Left = 0
    teamcol = teamcol + 1
  End If
Next team

ResumeApplicationAlerts alertStatus

End Sub

Sub GatherBellStats(ByVal team As Integer, ByVal Heading As String, ByVal FromCol As Integer, ByVal ToCol As Integer, ByVal FromTime As Integer, ByVal StartRow As Integer)
Dim gap(16, 16) As Integer
Dim gap1(16, 16) As Integer
Dim gap2(16, 16) As Integer
Dim gapcounter(16, 16) As Integer
Dim gap2counter(16, 16) As Integer
Dim Before As Integer
Dim After As Integer
Dim Place As Integer

Dim BeforeTime As Double
Dim AfterTime As Double
Dim AvGap As Double

Dim STime As Double
Dim ATime As Double

Dim i As Integer
Dim j As Integer

'empty the collector tables
For i = 1 To 16
  For j = 1 To 16
    gap(i, j) = 0
    gap1(i, j) = 0
    gap2(i, j) = 0
    gapcounter(i, j) = 0
    gap2counter(i, j) = 0
  Next j
Next i

'now populate the collector table by running through the touch
For j = 4 To NumWholepulls(team) + 1
  For i = FromCol To ToCol - 1
    Before = bell_chartonum(Worksheets(TeamName(team) + " 2").Cells(j, i))
    After = bell_chartonum(Worksheets(TeamName(team) + " 2").Cells(j, i + 1))
    BeforeTime = Worksheets(TeamName(team) + " 1").Cells(j, FromTime + i - FromCol)
    AfterTime = Worksheets(TeamName(team) + " 1").Cells(j, FromTime + i - FromCol + 1)
    STime = Worksheets(TeamName(team) + " 2").Cells(j, FromTime + i - FromCol + 1)
    'MsgBox ("before=" + Str(Before) + " after=" + Str(After) + " BeforeTime=" + Str(BeforeTime) + " AfterTime=" + Str(AfterTime))
    AvGap = Worksheets(TeamName(team) + " 1").Cells(j, (NumBells(team) * 2) + 5)
    gap(After, Before) = gap(After, Before) + AfterTime - BeforeTime - AvGap
    gap1(After, Before) = gap1(After, Before) + STime
    'MsgBox ("after=" + After + " before=" + Before + " gap=" + Str(gap(After, Before)) + " gap1=" + Str(gap1(After, Before)))
    gapcounter(After, Before) = gapcounter(After, Before) + 1
    
     
   ' If Before = After Then MsgBox ("Fault")
    'MsgBox (Gap(After, Before))
  Next i
Next j

If Heading = "HANDSTROKE" Then

For j = 4 To NumWholepulls(team) + 1
  For i = FromCol To ToCol

Before = bell_chartonum(Worksheets(TeamName(team) + " 2").Cells(j, i))

Place = i + 1 - FromCol
    If Place > NumBells(team) Then
          Place = Place - NumBells(team)
    End If
    
    ATime = Worksheets(TeamName(team) + " 2").Cells(j, FromTime + i - FromCol)
    gap2(Before, Place) = gap2(Before, Place) + ATime
    gap2counter(Before, Place) = gap2counter(Before, Place) + 1

  Next i
Next j

End If



If Heading = "BACKSTROKE" Then

For j = 4 To NumWholepulls(team) + 1
  For i = FromCol + 1 To ToCol

Before = bell_chartonum(Worksheets(TeamName(team) + " 2").Cells(j, i))

Place = i - FromCol
    If Place > NumBells(team) Then
          Place = Place - NumBells(team)
    End If
    
    ATime = Worksheets(TeamName(team) + " 2").Cells(j, FromTime + i - FromCol)
    gap2(Before, Place) = gap2(Before, Place) + ATime
    gap2counter(Before, Place) = gap2counter(Before, Place) + 1

  Next i
Next j

End If

If Heading = "ALL STROKES" Then

For j = 4 To NumWholepulls(team) + 1
  For i = FromCol To ToCol

Before = bell_chartonum(Worksheets(TeamName(team) + " 2").Cells(j, i))

Place = i + 1 - FromCol
    If Place > NumBells(team) Then
          Place = Place - NumBells(team)
    End If
    
    ATime = Worksheets(TeamName(team) + " 2").Cells(j, FromTime + i - FromCol)
    gap2(Before, Place) = gap2(Before, Place) + ATime
    gap2counter(Before, Place) = gap2counter(Before, Place) + 1

  Next i
Next j

End If



'establish a table in the worksheet
Cells(1, 1) = TeamName(team)
Cells(StartRow, 1) = Heading
Cells(StartRow, 3) = "Av error (ms) based on last bell plus interbell gap"
Cells(StartRow, 18) = "Av error (ms) based on calculated position"
'Cells(StartRow, NumBells(team) / 2 + 5) = "Proximity to bell"
Cells(StartRow + 1, NumBells(team) / 2) = "After"
Cells(StartRow + 1, NumBells(team) + 3 + (NumBells(team) / 2)) = "After"
Cells(StartRow + NumBells(team) / 2, 1) = "This Bell"
Cells(StartRow + NumBells(team) / 2, 1 + NumBells(team) + 3) = "This Bell"
For i = 1 To NumBells(team)
    Cells(StartRow + 2, i + 2) = i
    Cells(StartRow + 2, i + 2 + NumBells(team) + 3) = i
    Cells(StartRow + 2 + i, 2) = i
    Cells(StartRow + 2 + i, 2 + NumBells(team) + 3) = i
Next i

For i = 1 To 12
  For j = 1 To 12
     If gapcounter(i, j) = 0 Then
       Cells(StartRow + 2 + i, 2 + j) = 0
       Cells(StartRow + 2 + i, 5 + j + NumBells(team)) = 0
     Else
       Cells(StartRow + 2 + i, 2 + j) = gap(i, j) / gapcounter(i, j)
       Cells(StartRow + 2 + i, 5 + j + NumBells(team)) = gap1(i, j) / gapcounter(i, j)
     End If
     Cells(StartRow + 2 + i, 2 + j).NumberFormat = "0"
     Cells(StartRow + 2 + i, 5 + j + NumBells(team)).NumberFormat = "0"
  Next j
Next i


'establish another table in the worksheet
'Cells(StartRow + 55, 1) = Heading
Cells(StartRow, 9 + 2 * NumBells(team)) = "Av error (ms) based on position in the row"
Cells(StartRow + 1, 9 + 2 * NumBells(team) + NumBells(team) / 2) = "Place"
Cells(StartRow + NumBells(team) / 2, 2 * NumBells(team) + 7) = "This Bell"
For i = 1 To NumBells(team)
    Cells(StartRow + 2, i + 8 + 2 * NumBells(team)) = i
    Cells(StartRow + 2 + i, 8 + 2 * NumBells(team)) = i
Next i

For i = 1 To 12
  For j = 1 To 12
        If gap2counter(i, j) = 0 Then
            Cells(StartRow + 2 + i, 8 + 2 * NumBells(team) + j) = "N/A"
        Else
            Cells(StartRow + 2 + i, 8 + 2 * NumBells(team) + j) = gap2(i, j) / gap2counter(i, j)
        End If
     Cells(StartRow + 2 + i, 8 + 2 * NumBells(team) + j).NumberFormat = "0"
   Next j
Next i

'set cell size
'autofit columns
Cells.Columns.AutoFit

Columns("C:C").ColumnWidth = Columns("D:D").ColumnWidth
Columns("R:R").ColumnWidth = Columns("S:S").ColumnWidth
Columns("AG:AG").ColumnWidth = Columns("AH:AH").ColumnWidth
Columns("AM:AM").ColumnWidth = Columns("AN:AN").ColumnWidth
Cells(1, 1).Select

End Sub

Sub Bell_table()
Dim team As Integer

Dim FromCol As Integer
Dim ToCol As Integer
Dim StartRow As Integer
Dim FromTime As Integer
Dim calcWasSuspended As Boolean

'loadvars
Call loadvars

calcWasSuspended = SuspendCalculationAndRedraw()

For team = 1 To 10
  If IsTeamProcessed(team) Then
    'if it already exists then delete the Bell Table worksheet
    Application.DisplayAlerts = False
        For Each ws In Worksheets
      If ws.Name = "BELL ANALYSIS " + TeamName(team) Then Worksheets(ws.Name).Delete
    Next ws

    'now create a new sheet
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "BELL ANALYSIS " + TeamName(team)
    Worksheets("BELL ANALYSIS " + TeamName(team)).Activate
    
    'now populate it with handstrokes
    FromCol = (NumBells(team) * 2) + 6
    ToCol = (NumBells(team) * 3) + 5
    FromTime = 2
    StartRow = 3
    Call GatherBellStats((team), "HANDSTROKE", (FromCol), (ToCol), (FromTime), (StartRow))
    
    'now populate it with backstrokes
    FromCol = (NumBells(team) * 3) + 5
    ToCol = FromCol + NumBells(team)
    FromTime = NumBells(team) + 1
    StartRow = StartRow + NumBells(team) + 5
    Call GatherBellStats((team), "BACKSTROKE", (FromCol), (ToCol), (FromTime), (StartRow))
    
    'now populate it with everything
    FromCol = (NumBells(team) * 2) + 6
    ToCol = (NumBells(team) * 3) + 6 + NumBells(team) - 1
    FromTime = 2
    StartRow = StartRow + NumBells(team) + 5
    Call GatherBellStats((team), "ALL STROKES", (FromCol), (ToCol), (FromTime), (StartRow))
  End If
Next team


ResumeCalculationAndRedraw calcWasSuspended

End Sub

