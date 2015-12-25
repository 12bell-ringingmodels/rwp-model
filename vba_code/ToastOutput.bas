Attribute VB_Name = "ToastOutput"
Option Explicit

Private Const TOAST_RWP_SOURCE As String = "RodModel2"

Private Function ToastXML_CreateDatasourceElement(toastDOM As MSXML2.DOMDocument60) As IXMLDOMElement
    Dim workingElement As IXMLDOMElement
    Dim returnElement As IXMLDOMElement
    
    Set returnElement = toastDOM.createElement("datasource")
    Set workingElement = toastDOM.createElement("name")
    workingElement.Text = TOAST_RWP_SOURCE
    returnElement.appendChild workingElement
    
    Set workingElement = toastDOM.createElement("version")
    workingElement.Text = "Unknown"
    returnElement.appendChild workingElement
    
    'Set workingElement = toastDOM.createElement("comment")
    'returnElement.appendChild workingElement
    
    
   Set ToastXML_CreateDatasourceElement = returnElement
End Function


Private Function ToastXML_CreateStrikeData(ByVal teamIndex As Integer, toastDOM As MSXML2.DOMDocument60) As IXMLDOMElement
    Dim strikeElement As IXMLDOMElement
    Dim workingElement As IXMLDOMElement
    Dim workingAttribute As IXMLDOMAttribute
    Dim returnElement As IXMLDOMElement
    
    Set returnElement = toastDOM.createElement("strikeData")
    
    Dim rowIndex As Integer
    Dim bellIndex As Integer
    
    
    For rowIndex = 1 To NumRows(teamIndex)
        If (rowIndex >= AnalStart(teamIndex)) And (rowIndex <= AnalEnd(teamIndex)) Then
            Set workingElement = toastDOM.createElement("rowDelimiter")
            Set workingAttribute = toastDOM.createAttribute("source")
            workingAttribute.nodeValue = TOAST_RWP_SOURCE
            workingElement.setAttributeNode workingAttribute
            
            returnElement.appendChild workingElement
        End If
        
        For bellIndex = 1 To NumBells(teamIndex)
            Set strikeElement = toastDOM.createElement("strike")
            Set workingElement = toastDOM.createElement("bell")
            workingElement.Text = LoadTime(teamIndex, bellIndex, rowIndex).bell
            strikeElement.appendChild workingElement
            Set workingElement = toastDOM.createElement("original")
            workingElement.Text = 0.001 * (LoadTime(teamIndex, bellIndex, rowIndex).time)
            strikeElement.appendChild workingElement
            
            
            returnElement.appendChild strikeElement
        Next
    Next
    
    
   Set ToastXML_CreateStrikeData = returnElement
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
    
    ' Add in the strike data
    toastRoot.appendChild ToastXML_CreateStrikeData(teamIndex, toastDOM)
    
    ' And we're done
    
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
