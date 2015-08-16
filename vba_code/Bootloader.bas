Attribute VB_Name = "Bootloader"
Option Explicit
Public Function IsProtectedModule(ByVal moduleName As String) As Boolean
    Dim lcaseName As String
    
    lcaseName = LCase(moduleName)
    
    If InStr(lcaseName, "bootloader") > 0 Then
        IsProtectedModule = True
    ElseIf InStr(lcaseName, "filetools") > 0 Then
        IsProtectedModule = True
    Else
        IsProtectedModule = False
    End If
End Function

Public Function IsProtectedClassModule(ByVal moduleName As String) As Boolean
    Dim lcaseName As String
    
    lcaseName = LCase(moduleName)
    
    If InStr(lcaseName, "sheet") > 0 Then
        IsProtectedClassModule = True
    ElseIf InStr(lcaseName, "workbook") > 0 Then
        IsProtectedClassModule = True
    Else
        IsProtectedClassModule = False
    End If
End Function


Public Sub ImportModules(ByVal moduleDirectory As String)
    Dim loadPath As String
    Dim targetModule As String
    
    loadPath = CvtToAbsFile(moduleDirectory)
    
    Dim pVBAProject As VBProject
    Set pVBAProject = ThisWorkbook.VBProject
    
    targetModule = Dir(JoinPath(loadPath, "*.bas"))

    While targetModule <> ""
        If Not IsProtectedModule(targetModule) Then
            pVBAProject.VBComponents.Import JoinPath(loadPath, targetModule)
        End If
        
        targetModule = Dir
    Wend
    
    targetModule = Dir(JoinPath(loadPath, "*.cls"))

    While targetModule <> ""
        If Not IsProtectedModule(targetModule) And Not IsProtectedClassModule(targetModule) Then
            pVBAProject.VBComponents.Import JoinPath(loadPath, targetModule)
        End If
        
        targetModule = Dir
    Wend

End Sub
Public Sub ExportModules(ByVal moduleDirectory As String, Optional ByVal removeModules As Boolean = True)
  Dim pVBAProject As VBProject
  Dim vbComp As VBComponent  'VBA module, form, etc...
  
  Dim savePath As String
  
  savePath = CvtToAbsFile(moduleDirectory)
  
  ' If this folder doesn't exist
  If Dir(savePath, vbDirectory) = "" Then
    ' Create it
    MkDir savePath
  Else
    ' Move old data into a backup directory
    Name savePath As (savePath & "_ASV_" & Format(Now, "yyyymmdd-hhmmss"))
    ' and create a blank directory
    MkDir savePath
  End If
  
  ' Get the VBA project
  Set pVBAProject = ThisWorkbook.VBProject
  
  ' Loop through all the components (modules, forms, etc) in the VBA project
  For Each vbComp In pVBAProject.VBComponents
    Select Case vbComp.Type
    Case vbext_ct_StdModule
      vbComp.Export savePath & "\" & vbComp.Name & ".bas"
    Case vbext_ct_Document, vbext_ct_ClassModule
      ' ThisDocument and class modules
      vbComp.Export savePath & "\" & vbComp.Name & ".cls"
    Case vbext_ct_MSForm
      vbComp.Export savePath & "\" & vbComp.Name & ".frm"
    Case Else
      vbComp.Export savePath & "\" & vbComp.Name
    End Select
    
    If removeModules Then
        ' Currently only rebuild of .BAS files is supported
        If vbComp.Type = vbext_ct_StdModule Or vbComp.Type = vbext_ct_ClassModule Then
            If Not IsProtectedModule(vbComp.Name) Then
                pVBAProject.VBComponents.Remove vbComp
            End If
        End If
    End If
  Next
  
End Sub

