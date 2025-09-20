Attribute VB_Name = "modXMLHelper"
Option Explicit




Public Function GetXMLAttributeValue(ByVal SrcXMLPath As String, _
                                     ByVal SrcParam As String, _
                                     ByVal SrcFileName As String) As String
                                     
    On Error Resume Next
    
    
    Dim XDoc As CGoXML
    Dim StrResult As String
    
    Set XDoc = New CGoXML
    
    With XDoc
        .Initialize (pavAuto)
        
        .OpenFromFile (SrcFileName)
        
        StrResult = .ReadAttribute(SrcXMLPath, SrcParam)
    End With
    
    GetXMLAttributeValue = StrResult
End Function


Public Sub UpdateXMLAttributeValue(ByVal SrcXMLPath As String, _
                                     ByVal SrcParam As String, _
                                     ByVal SrcValue As String, _
                                     ByVal SrcFileName As String)
    On Error Resume Next
    
    Dim XDoc As CGoXML
    Set XDoc = New CGoXML
    
    With XDoc
        .Initialize (pavAuto)
        
        .OpenFromFile (SrcFileName)
        
        .WriteAttribute SrcXMLPath, SrcParam, SrcValue
        
        .Save (SrcFileName)
    End With
End Sub


