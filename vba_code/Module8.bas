Attribute VB_Name = "Module8"
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


