VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Sheets("Settings").Range("optDevTeardown") Then
        ExportModules "vba_code", "vba_code_protected"
    End If
End Sub

Private Sub Workbook_Open()
    If Sheets("Settings").Range("optDevTeardown") Then
        ImportModules "vba_code"
    End If
End Sub
