Attribute VB_Name = "XMLFile"
'@Folder("TinLine")
Option Explicit

Public Function WriteXML(ByRef oxml As MSXML2.DOMDocument60, ByRef ParentNode As IXMLDOMElement)



End Function


Public Function CreateXmlAttribute(Name As String, Bez As String, Wert As String, str As String, nodChild As IXMLDOMElement, oxml As MSXML2.DOMDocument60, nodElement As IXMLDOMElement)
    ' create a TinLine XML Attribute with the given informations
    Dim nodGrandChild        As IXMLDOMElement

    Set nodChild = oxml.createElement(str)
    nodElement.appendChild nodChild

    Set nodGrandChild = oxml.createElement("Name")
    nodGrandChild.Text = Name
    nodChild.appendChild nodGrandChild

    Set nodGrandChild = oxml.createElement("Bez")
    nodGrandChild.Text = Bez
    nodChild.appendChild nodGrandChild

    Set nodGrandChild = oxml.createElement("Wert")
    nodGrandChild.Text = Wert
    nodChild.appendChild nodGrandChild

    CreateXmlAttribute = True

End Function

Public Function CreateXmlIndexAttribute(Index As String, Name As String, Datum As String, Bez As String, NodName As String, nodChild As IXMLDOMElement, oxml As MSXML2.DOMDocument60, nodElement As IXMLDOMElement)
    ' create a TinLine XML Attribute with the given informations
    Dim nodGrandChild        As IXMLDOMElement

    Set nodChild = oxml.createElement(NodName)
    nodElement.appendChild nodChild

    Set nodGrandChild = oxml.createElement("Index")
    nodGrandChild.Text = Index
    nodChild.appendChild nodGrandChild

    Set nodGrandChild = oxml.createElement("Name")
    nodGrandChild.Text = Name
    nodChild.appendChild nodGrandChild

    Set nodGrandChild = oxml.createElement("Datum")
    nodGrandChild.Text = Datum
    nodChild.appendChild nodGrandChild

    Set nodGrandChild = oxml.createElement("Bez")
    nodGrandChild.Text = Bez
    nodChild.appendChild nodGrandChild

    CreateXmlIndexAttribute = True

End Function
