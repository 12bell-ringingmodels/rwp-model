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
Public Sub ExportModules(ByVal moduleDirectory As String, Optional ByVal protectedModuleDirectory, Optional ByVal removeModules As Boolean = True)
    Dim pVBAProject As VBProject
    Dim vbComp As VBComponent  'VBA module, form, etc...
    
    Dim savePathUser As String
    Dim savePathProtect As String
    Dim timestamp As String
    
    timestamp = Format(Now, "yyyymmdd-hhmmss")
    
    savePathUser = CvtToAbsFile(moduleDirectory)
    
    If IsMissing(protectedModuleDirectory) Then
        savePathProtect = savePathUser
    Else
        savePathProtect = CvtToAbsFile(protectedModuleDirectory)
    End If
    
    ' If the user folder doesn't exist
    If Dir(savePathUser, vbDirectory) = "" Then
        ' Create it
        MkDir savePathUser
    Else
        ' Move old data into a backup directory
        Name savePathUser As (savePathUser & "_ASV_" & timestamp)
        ' and create a blank directory
        MkDir savePathUser
    End If
    
    If savePathProtect <> savePathUser Then
        ' If the user folder doesn't exist
        If Dir(savePathProtect, vbDirectory) = "" Then
            ' Create it
            MkDir savePathProtect
        Else
            ' Move old data into a backup directory
            Name savePathProtect As (savePathProtect & "_ASV_" & timestamp)
            ' and create a blank directory
            MkDir savePathProtect
        End If
    End If
    
    ' Get the VBA project
    Set pVBAProject = ThisWorkbook.VBProject
    
    ' Loop through all the components (modules, forms, etc) in the VBA project
    For Each vbComp In pVBAProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule
                If IsProtectedModule(vbComp.Name) Then
                    vbComp.Export savePathProtect & "\" & vbComp.Name & ".bas"
                Else
                    vbComp.Export savePathUser & "\" & vbComp.Name & ".bas"
                End If
            Case vbext_ct_Document, vbext_ct_ClassModule
                ' ThisDocument and class modules
                If IsProtectedClassModule(vbComp.Name) Then
                    vbComp.Export savePathProtect & "\" & vbComp.Name & ".cls"
                Else
                    vbComp.Export savePathUser & "\" & vbComp.Name & ".cls"
                End If
            Case vbext_ct_MSForm
                vbComp.Export savePathProtect & "\" & vbComp.Name & ".frm"
            Case Else
                vbComp.Export savePathProtect & "\" & vbComp.Name
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

