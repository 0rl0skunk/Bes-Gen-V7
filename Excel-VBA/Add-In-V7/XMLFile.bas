Attribute VB_Name = "XMLFile"
'@Folder("TinLine")
Option Explicit

Public Function WriteXML(ByRef oXml As MSXML2.DOMDocument60, ByRef ParentNode As IXMLDOMElement)



End Function

Public Function CreateXmlAttribute(Name As String, Bez As String, Wert As String, str As String, nodChild As IXMLDOMElement, oXml As MSXML2.DOMDocument60, nodElement As IXMLDOMElement)
    ' create a TinLine XML Attribute with the given informations
    Dim nodGrandChild        As IXMLDOMElement

    Set nodChild = oXml.createElement(str)
    nodElement.appendChild nodChild

    Set nodGrandChild = oXml.createElement("Name")
    nodGrandChild.text = Name
    nodChild.appendChild nodGrandChild

    Set nodGrandChild = oXml.createElement("Bez")
    nodGrandChild.text = Bez
    nodChild.appendChild nodGrandChild

    Set nodGrandChild = oXml.createElement("Wert")
    nodGrandChild.text = Wert
    nodChild.appendChild nodGrandChild

    CreateXmlAttribute = True

End Function

Public Function CreateXmlIndexAttribute(Index As String, Name As String, Datum As String, Bez As String, NodName As String, nodChild As IXMLDOMElement, oXml As MSXML2.DOMDocument60, nodElement As IXMLDOMElement)
    ' create a TinLine XML Attribute with the given informations
    Dim nodGrandChild        As IXMLDOMElement

    Set nodChild = oXml.createElement(NodName)
    nodElement.appendChild nodChild

    Set nodGrandChild = oXml.createElement("Index")
    nodGrandChild.text = Index
    nodChild.appendChild nodGrandChild

    Set nodGrandChild = oXml.createElement("Name")
    nodGrandChild.text = Name
    nodChild.appendChild nodGrandChild

    Set nodGrandChild = oXml.createElement("Datum")
    nodGrandChild.text = Datum
    nodChild.appendChild nodGrandChild

    Set nodGrandChild = oXml.createElement("Bez")
    nodGrandChild.text = Bez
    nodChild.appendChild nodGrandChild

    CreateXmlIndexAttribute = True

End Function

