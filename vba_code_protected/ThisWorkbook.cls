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
    ExportModules VBA_CODE_DIRECTORY, VBA_CODE_PROTECTED_DIRECTORY
End Sub

Private Sub Workbook_Open()
    If Sheets("Settings").Range("optDevAutoAssembly") Then
        ImportModules VBA_CODE_DIRECTORY
    End If
End Sub
