Attribute VB_Name = "WorksheetTools"
Option Explicit

Public Const PROGRAMMATIC_SHEET_NAME As String = "ProgrammaticSheets"
Public Const FRONT_SHEET_NAME As String = "Front"
Public Const MEDIA_PLAYER_SHEET_NAME As String = "MediaPlayer"
Public Const ABOUT_SHEET_NAME As String = "About"
Public Const SETTINGS_SHEET_NAME As String = "Settings"

Public Const RWP_WORKING_SHEETS As String = "RWPWorking"

Public Sub ResetForCommit()
    Dim FrontSheet As Worksheet
    Set FrontSheet = Sheets(FRONT_SHEET_NAME)
    Dim exampleFile As String
    Dim row As Integer
    Dim exampleDataIndex As Integer
    Dim newTitle As String
    
    
    If MsgBox("This action will result in all team data and settings being reset to default values. It cannot be undone. Do you wish to continue?", vbYesNo Or vbDefaultButton2, "Reset?") <> vbYes Then
        Exit Sub
    End If
    
    FrontSheet.Range("B4:H13").Clear
    FrontSheet.Range("touchFile") = "ExampleData\Cambridge288.txt"
    FrontSheet.Range("touchFile").Offset(1, 0) = ""

    exampleFile = Dir(CvtToAbsFile("ExampleData\exampleData*.txt"))
    FrontSheet.Range("strikeBaseDir") = ""
    FrontSheet.Range("flocktonBaseDir") = ""
    
    row = 0
    While exampleFile <> ""
        exampleDataIndex = InStr(exampleFile, "exampleData")
        newTitle = "Team " & UCase(Mid(exampleFile, exampleDataIndex + 12, Len(exampleFile) - (exampleDataIndex + 12) - 3))
        FrontSheet.Range("B4").Offset(row, 0) = newTitle
        FrontSheet.Range("B4").Offset(row, 1) = "ExampleData\" & exampleFile
        FrontSheet.Range("B4").Offset(row, 6) = "y"
        
        row = row + 1
        exampleFile = Dir
    Wend
    
    Call ResetSettingsSheetToDefault

    Call DeleteWorkingSheets
End Sub


Public Sub ResetForDistribution()
    Dim FrontSheet As Worksheet
    Set FrontSheet = Sheets(FRONT_SHEET_NAME)
    Dim exampleFile As String
    Dim row As Integer
    Dim exampleDataIndex As Integer
    Dim newTitle As String
    
    
    If MsgBox("This action will result in all team data and settings being reset to default values. It cannot be undone. Do you wish to continue?", vbYesNo Or vbDefaultButton2, "Reset?") <> vbYes Then
        Exit Sub
    End If
    
    FrontSheet.Range("B4:H13").Clear
    FrontSheet.Range("touchFile") = "ExampleData\Cambridge288.txt"
    FrontSheet.Range("touchFile").Offset(1, 0) = ""

    exampleFile = Dir(CvtToAbsFile("ExampleData\exampleData*.txt"))
    FrontSheet.Range("strikeBaseDir") = ""
    FrontSheet.Range("flocktonBaseDir") = ""
    
    row = 0
    While exampleFile <> ""
        exampleDataIndex = InStr(exampleFile, "exampleData")
        newTitle = "Team " & UCase(Mid(exampleFile, exampleDataIndex + 12, Len(exampleFile) - (exampleDataIndex + 12) - 3))
        FrontSheet.Range("B4").Offset(row, 0) = newTitle
        FrontSheet.Range("B4").Offset(row, 1) = "ExampleData\" & exampleFile
        FrontSheet.Range("B4").Offset(row, 6) = "y"
        
        row = row + 1
        exampleFile = Dir
    Wend
    
    Call ResetSettingsSheetToDefault

    Call DeleteWorkingSheets
End Sub
Sub ResetForDeveloperAccess()

End Sub

Private Sub ResetSettingsSheetToDefault()
    Dim settingsSheet As Worksheet
    Set settingsSheet = Sheets(SETTINGS_SHEET_NAME)
    
    settingsSheet.Range("optPresMode") = PresentationModes.General
    settingsSheet.Range("optAnalysisOpeningRounds") = False
    settingsSheet.Range("optAnalysisFaultPctSrc") = 25
    settingsSheet.Range("optAnalysisQuickHandstrokeLeadPctSrc") = "=" & settingsSheet.Range("optAnalysisFaultPctSrc").Address & "*Sqrt(2)"
    settingsSheet.Range("optAnalysisHandstrokeGapMethod") = HandstrokeGapModes.Averages
    settingsSheet.Range("optDisplayHideQuickLeads") = True
    settingsSheet.Range("optDisplayHighlighting") = False
    settingsSheet.Range("optDevProfile") = False
    settingsSheet.Range("optDevSuspendRedraw") = True
End Sub

Public Sub DeleteWorkingSheets()
    Dim wasDisplayingAlerts As Boolean
    Dim ws As Worksheet
    Dim DeleteSheet As Boolean
    Dim alertState As Boolean
    
    Worksheets(FRONT_SHEET_NAME).Activate
    alertState = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For Each ws In Sheets
        DeleteSheet = True
        If ws.Name = FRONT_SHEET_NAME Then DeleteSheet = False
        If ws.Name = MEDIA_PLAYER_SHEET_NAME Then DeleteSheet = False
        If ws.Name = PROGRAMMATIC_SHEET_NAME Then DeleteSheet = False
        If ws.Name = ABOUT_SHEET_NAME Then DeleteSheet = False
        If ws.Name = SETTINGS_SHEET_NAME Then DeleteSheet = False
        If InStr(ws.Name, "NoDelete") = 1 Then DeleteSheet = False
    
        If DeleteSheet Then
            ws.Delete
        End If
    Next ws
    Call ResetProgrammaticSheetList
    
    Application.DisplayAlerts = alertState
End Sub

Public Function WorksheetExists(SheetName As String)
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = SheetName Then
            WorksheetExists = True
            Exit Function
        End If
    Next

    WorksheetExists = False
End Function
Public Function SuspendApplicationAlerts() As Boolean
    SuspendApplicationAlerts = Application.DisplayAlerts
End Function
Public Sub ResumeApplicationAlerts(ByVal wasDisplayingAlerts As Boolean)
    Application.DisplayAlerts = wasDisplayingAlerts
End Sub
Function SuspendCalculationAndRedraw() As Boolean
    Dim doSuspend As Boolean
    
    SuspendCalculationAndRedraw = Application.ScreenUpdating
    
    doSuspend = True
    If WorksheetExists(SETTINGS_SHEET_NAME) Then
        doSuspend = Worksheets(SETTINGS_SHEET_NAME).Range("optSuspendRedraw").Value
    End If
        
    If doSuspend Then
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
    End If
End Function
Sub ScreenRefresh()
    Dim wasUpdating As Boolean
    
    wasUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = wasUpdating
End Sub
Sub ResumeCalculationAndRedraw(ByVal wasSuspended As Boolean)
    If Not wasSuspended Then
        Application.Calculation = xlCalculationAutomatic
        Application.Calculate
        Application.ScreenUpdating = True
    End If
End Sub
Public Sub ForceRecalculate()
    Application.Calculate
End Sub


Public Sub ResetProgrammaticSheetList()
    Dim programmaticSheet As Worksheet
    
    Set programmaticSheet = Worksheets(PROGRAMMATIC_SHEET_NAME)
    programmaticSheet.Range(programmaticSheet.Cells(2, 1), programmaticSheet.Cells(programmaticSheet.UsedRange.Rows.Count, 1)).EntireRow.Delete
    
    programmaticSheet.Cells(1, 1) = "Programmatic sheet"
    programmaticSheet.Cells(1, 2) = "Created by"
    
End Sub

Public Sub RemoveProgrammaticSheetsCreatedBy(ByVal CreatorID As String)
    Dim programmaticSheet As Worksheet
    Dim row_index As Integer
    Dim alertState As Boolean
    
    Set programmaticSheet = Worksheets(PROGRAMMATIC_SHEET_NAME)
    
    alertState = Application.DisplayAlerts
    Application.DisplayAlerts = False
    row_index = 2
    Do
        If programmaticSheet.Cells(row_index, 2) = CreatorID Then
            If WorksheetExists(programmaticSheet.Cells(row_index, 1).Value) Then
                Sheets(programmaticSheet.Cells(row_index, 1).Value).Delete
            End If
            programmaticSheet.Cells(row_index, 2).EntireRow.Delete
        Else
            row_index = row_index + 1
        End If
    Loop While programmaticSheet.Cells(row_index, 1) <> ""
    
    Application.DisplayAlerts = alertState
End Sub
Public Sub RegisterProgrammaticSheet(ByVal CreatedSheet As String, ByVal CreatorID As String)
    Dim programmaticSheet As Worksheet
    Set programmaticSheet = Worksheets(PROGRAMMATIC_SHEET_NAME)
    
    Dim row_index As Integer
    
    row_index = 1
    Do
        row_index = row_index + 1
    Loop While programmaticSheet.Cells(row_index, 1) <> ""
    programmaticSheet.Cells(row_index, 1) = CreatedSheet
    programmaticSheet.Cells(row_index, 2) = CreatorID
End Sub
Public Function GetProgrammaticSheetsCreatedBy(ByVal CreatorID As String) As String()
    Dim resultStrings() As String
    Dim programmaticSheet As Worksheet
    Set programmaticSheet = Worksheets(PROGRAMMATIC_SHEET_NAME)
    
    Dim row_index As Integer
    Dim resultsGenerated As Integer
    
    ReDim resultStrings(0 To 0)
    
    resultsGenerated = 0
    
    row_index = 2
    Do
        If programmaticSheet.Cells(row_index, 2) = CreatorID Then
            resultsGenerated = resultsGenerated + 1
            ReDim Preserve resultStrings(0 To resultsGenerated)
            resultStrings(resultsGenerated) = programmaticSheet.Cells(row_index, 1)
        End If
        row_index = row_index + 1
    Loop While programmaticSheet.Cells(row_index, 1) <> ""
    
    GetProgrammaticSheetsCreatedBy = resultStrings()
End Function
Public Sub SelectSheets(TargetSheets() As String, Optional ByVal replace = False)
    Dim addSheets As Integer
    
    For addSheets = 1 To UBound(TargetSheets)
        Sheets(TargetSheets(addSheets)).Select replace:=replace
    Next

End Sub

Public Sub SetupSinglePagePrint(sourceSheet As Worksheet, inputRange As Range, orientation As XlPageOrientation, Optional cHeader, Optional lHeader, Optional rHeader, Optional cFooter, Optional lFooter, Optional rFooter)
    Application.PrintCommunication = False
    
    If IsMissing(lHeader) Then lHeader = ""
    If IsMissing(rHeader) Then rHeader = ""
    If IsMissing(cHeader) Then cHeader = ""
    If IsMissing(lFooter) Then lFooter = ""
    If IsMissing(rFooter) Then rFooter = ""
    If IsMissing(cFooter) Then cFooter = ""
    
    With sourceSheet.PageSetup
        .PrintArea = inputRange.Address
        .LeftHeader = lHeader
        .CenterHeader = cHeader
        .RightHeader = rHeader
        .LeftFooter = lFooter
        .CenterFooter = cFooter
        .RightFooter = rFooter
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .orientation = orientation
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = lHeader
        .EvenPage.CenterHeader.Text = cHeader
        .EvenPage.RightHeader.Text = rHeader
        .EvenPage.LeftFooter.Text = lFooter
        .EvenPage.CenterFooter.Text = cFooter
        .EvenPage.RightFooter.Text = rFooter
        .FirstPage.LeftHeader.Text = lHeader
        .FirstPage.CenterHeader.Text = cHeader
        .FirstPage.RightHeader.Text = rHeader
        .FirstPage.LeftFooter.Text = lFooter
        .FirstPage.CenterFooter.Text = cFooter
        .FirstPage.RightFooter.Text = rFooter
    End With
    Application.PrintCommunication = True

End Sub

