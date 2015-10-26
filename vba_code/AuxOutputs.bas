Attribute VB_Name = "AuxOutputs"
Option Explicit

Public Sub Gen_Contest_PDF()
    Dim pdf_filename As String
    Sheets("SUMMARY").Select
    
    SelectSheets GetProgrammaticSheetsCreatedBy("TeamCharts")
    SelectSheets GetProgrammaticSheetsCreatedBy("Flockton")
    
    pdf_filename = Application.GetSaveAsFilename(TeamName(1) & ".pdf", "PDF Files,*.pdf", 1, "Select the output PDF name")
    If IsNumeric(pdf_filename) Then
        Exit Sub
    End If
    
    ActiveWorkbook.RunAutoMacros Which:=xlAutoClose
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:=pdf_filename, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End Sub

Public Sub Gen_Toast_XML()
    Dim lclTeams As Integer
    Dim multiOutput As Boolean
    
    lclTeams = TotalTeams()
    multiOutput = False
    

    If lclTeams <= 0 Then
        MsgBox "Please load at least one file"
        Exit Sub
    ElseIf lclTeams > 1 Then
        If MsgBox("Multiple files have been loaded - do you want to generate output for all of them?", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
        multiOutput = True
    End If
    

    xml_filename = Application.GetSaveAsFilename("RodModel.xml", "XML Files,*.pdf", 1, "Select the output XML file name")
End Sub
