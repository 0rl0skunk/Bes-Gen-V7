Attribute VB_Name = "XMLFile"

'@Folder("TinLine")
'@ModuleDescription "TinLine Spezifische XML-Formatierung"

Option Explicit

Public Function CreateXmlAttribute(ByVal Name As String, ByVal Bez As String, ByVal Wert As String, ByVal str As String, ByRef NodChild As IXMLDOMElement, ByRef oXml As MSXML2.DOMDocument60, ByVal NodElement As IXMLDOMElement)
    ' create a TinLine XML Attribute with the given informations
    Dim NodGrandChild        As IXMLDOMElement

    Set NodChild = oXml.createElement(str)
    NodElement.appendChild NodChild

    Set NodGrandChild = oXml.createElement("Name")
    NodGrandChild.Text = Name
    NodChild.appendChild NodGrandChild

    Set NodGrandChild = oXml.createElement("Bez")
    NodGrandChild.Text = Bez
    NodChild.appendChild NodGrandChild

    Set NodGrandChild = oXml.createElement("Wert")
    NodGrandChild.Text = Wert
    NodChild.appendChild NodGrandChild

    CreateXmlAttribute = True

End Function

Public Function CreateXmlIndexAttribute(ByVal Index As String, ByVal Gezeichnet As String, ByRef Bez As String, ByVal NodName As String, ByRef NodChild As IXMLDOMElement, ByRef oXml As MSXML2.DOMDocument60, ByVal NodElement As IXMLDOMElement)
    ' create a TinLine XML Attribute with the given informations
    Dim NodGrandChild        As IXMLDOMElement
    Dim Person               As String
    Dim Datum                As String
    Dim Text()               As String

    Person = Split(Gezeichnet, ";")(0)
    Datum = Split(Gezeichnet, ";")(1)

    Text = SplitStringByLength(Bez, 100)
    If UBound(Text) > 1 Then
        Dim i                As Long
        Set NodChild = oXml.createElement(NodName)
        NodElement.appendChild NodChild

        Set NodGrandChild = oXml.createElement("Index")
        NodGrandChild.Text = Index
        NodChild.appendChild NodGrandChild

        Set NodGrandChild = oXml.createElement("Name")
        NodGrandChild.Text = Person
        NodChild.appendChild NodGrandChild

        Set NodGrandChild = oXml.createElement("Datum")
        NodGrandChild.Text = Datum
        NodChild.appendChild NodGrandChild

        Set NodGrandChild = oXml.createElement("Bez")
        NodGrandChild.Text = Text(1)
        NodChild.appendChild NodGrandChild
        For i = LBound(Text) + 2 To UBound(Text)
            Set NodChild = oXml.createElement(NodName)
            NodElement.appendChild NodChild

            Set NodGrandChild = oXml.createElement("Index")
            NodChild.appendChild NodGrandChild

            Set NodGrandChild = oXml.createElement("Name")
            NodChild.appendChild NodGrandChild

            Set NodGrandChild = oXml.createElement("Datum")
            NodChild.appendChild NodGrandChild

            Set NodGrandChild = oXml.createElement("Bez")
            NodGrandChild.Text = Text(i)
            NodChild.appendChild NodGrandChild
        Next
    Else

        Set NodChild = oXml.createElement(NodName)
        NodElement.appendChild NodChild

        Set NodGrandChild = oXml.createElement("Index")
        NodGrandChild.Text = Index
        NodChild.appendChild NodGrandChild

        Set NodGrandChild = oXml.createElement("Name")
        NodGrandChild.Text = Person
        NodChild.appendChild NodGrandChild

        Set NodGrandChild = oXml.createElement("Datum")
        NodGrandChild.Text = Datum
        NodChild.appendChild NodGrandChild

        Set NodGrandChild = oXml.createElement("Bez")
        NodGrandChild.Text = Text(1)
        NodChild.appendChild NodGrandChild
    End If
    CreateXmlIndexAttribute = True

End Function

