Attribute VB_Name = "ToastOutput"
Option Explicit

Private Function ToastXML_CreateDatasourceElement(toastDOM As MSXML2.DOMDocument60) As IXMLDOMElement
    Dim workingElement As IXMLDOMElement
    Dim returnElement As IXMLDOMElement
    
    Set returnElement = toastDOM.createElement("datasource")
    Set workingElement = toastDOM.createElement("name")
    workingElement.Text = "RodModel2"
    returnElement.appendChild workingElement
    
    Set workingElement = toastDOM.createElement("version")
    workingElement.Text = "Unknown"
    returnElement.appendChild workingElement
    
    'Set workingElement = toastDOM.createElement("comment")
    'returnElement.appendChild workingElement
    
    
   Set ToastXML_CreateDatasourceElement = returnElement
End Function

Public Sub WriteTeamToastXML(ByVal teamIndex As Integer, ByVal outputFile As String)
    Dim toastDOM As MSXML2.DOMDocument60
    Dim toastRoot As IXMLDOMElement
    Dim objXMLelement As IXMLDOMElement
    Dim objXMLattr As IXMLDOMAttribute
    
    Set toastDOM = New MSXML2.DOMDocument60
    
    ' Create a root element
    Set toastRoot = toastDOM.createElement("transcription")
    toastDOM.appendChild toastRoot
    
    ' Define the data source
    Dim toastSources As IXMLDOMElement
    Set toastSources = toastDOM.createElement("dataSources")
    toastSources.appendChild ToastXML_CreateDatasourceElement(toastDOM)
    toastRoot.appendChild toastSources
    
    
    toastDOM.Save outputFile
    
End Sub


Public Sub Gen_XML()
    Dim xml_filename As String
    Dim index_teams As Integer
    
    Sheets("SUMMARY").Select
    
    For index_teams = 1 To MAXIMUM_TEAMS
        If IsTeamProcessed(index_teams) Then
            xml_filename = Application.GetSaveAsFilename(TeamName(index_teams) & ".xml", "XML Files,*.xml", 1, "Select the output XML name for Team " & index_teams)
            If Not IsNumeric(xml_filename) Then
                
            End If
        End If
    Next
        
    xml_filename = Application.GetSaveAsFilename(TeamName(1) & ".xml", "XML Files,*.xml", 1, "Select the output XML name")
    If IsNumeric(xml_filename) Then
        Exit Sub
    End If
    
End Sub
