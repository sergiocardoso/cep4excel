Attribute VB_Name = "cepsergiocardosoorg"
Function CEP(valorcep As String, tipoCampo As String)

Dim oXmlDoc As DOMDocument
Dim oXmlNode As IXMLDOMNode
Dim oXmlNodes As IXMLDOMNodeList

Set oXmlDoc = New DOMDocument
oXmlDoc.async = False
oXmlDoc.Load ("http://sergiocardoso.org/cep/?cep=" + valorcep)

Set oXmlNodes = oXmlDoc.SelectNodes("/endereco/" + tipoCampo)
    
For Each oXmlNode In oXmlNodes
    CEP = oXmlNode.Text
Next

End Function


