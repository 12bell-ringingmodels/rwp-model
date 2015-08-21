Attribute VB_Name = "VersionTools"
Option Explicit

Public Enum ExcelVersions
    ExcelVer_Low = 1
    ExcelVer_97 = 8
    ExcelVer_2000 = 9
    ExcelVer_2002 = 10
    ExcelVer_2003 = 11
    ExcelVer_2007 = 12
    ExcelVer_2010 = 14
    ExcelVer_2013 = 15
End Enum

Function GetExcelVersion() As ExcelVersions
    Dim baseVersionNumber As Integer
    
    baseVersionNumber = Int(Application.Version)
    If baseVersionNumber < ExcelVer_97 Then
        GetExcelVersion = ExcelVer_Low
    ElseIf baseVersionNumber <= ExcelVer_97 + 0.5 Then
        GetExcelVersion = ExcelVer_2000
    ElseIf baseVersionNumber <= ExcelVer_2002 + 0.5 Then
        GetExcelVersion = ExcelVer_2002
    ElseIf baseVersionNumber <= ExcelVer_2003 + 0.5 Then
        GetExcelVersion = ExcelVer_2003
    ElseIf baseVersionNumber <= ExcelVer_2007 + 0.5 Then
        GetExcelVersion = ExcelVer_2007
    ElseIf baseVersionNumber <= ExcelVer_2010 + 0.5 Then
        GetExcelVersion = ExcelVer_2010
    Else
        GetExcelVersion = ExcelVer_2013
    End If
End Function
