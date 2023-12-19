Attribute VB_Name = "XMLFile"
'@Folder("TinLine")
Option Explicit

Public Function CreateXmlAttribute(Name As String, Bez As String, Wert As String, str As String, NodChild As IXMLDOMElement, oXml As MSXML2.DOMDocument60, NodElement As IXMLDOMElement)
    ' create a TinLine XML Attribute with the given informations
    Dim NodGrandChild        As IXMLDOMElement

    Set NodChild = oXml.createElement(str)
    NodElement.appendChild NodChild

    Set NodGrandChild = oXml.createElement("Name")
    NodGrandChild.text = Name
    NodChild.appendChild NodGrandChild

    Set NodGrandChild = oXml.createElement("Bez")
    NodGrandChild.text = Bez
    NodChild.appendChild NodGrandChild

    Set NodGrandChild = oXml.createElement("Wert")
    NodGrandChild.text = Wert
    NodChild.appendChild NodGrandChild

    CreateXmlAttribute = True

End Function

Public Function CreateXmlIndexAttribute(Index As String, Name As String, Datum As String, Bez As String, NodName As String, NodChild As IXMLDOMElement, oXml As MSXML2.DOMDocument60, NodElement As IXMLDOMElement)
    ' create a TinLine XML Attribute with the given informations
    Dim NodGrandChild        As IXMLDOMElement

    Set NodChild = oXml.createElement(NodName)
    NodElement.appendChild NodChild

    Set NodGrandChild = oXml.createElement("Index")
    NodGrandChild.text = Index
    NodChild.appendChild NodGrandChild

    Set NodGrandChild = oXml.createElement("Name")
    NodGrandChild.text = Name
    NodChild.appendChild NodGrandChild

    Set NodGrandChild = oXml.createElement("Datum")
    NodGrandChild.text = Datum
    NodChild.appendChild NodGrandChild

    Set NodGrandChild = oXml.createElement("Bez")
    NodGrandChild.text = Bez
    NodChild.appendChild NodGrandChild

    CreateXmlIndexAttribute = True

End Function

