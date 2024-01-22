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

Public Sub deleteIndexesXml()

    Dim fso                  As Object
    Dim lastrow              As Long
    Dim row                  As Long
    Dim col                  As Long
    Dim lastcol              As Long


    Set fso = CreateObject("scripting.FileSystemObject")

    lastrow = shGebäude.Cells(shGebäude.rows.Count, 2).End(xlUp).row
    lastcol = shGebäude.Cells(1, shGebäude.Columns.Count).End(xlToLeft).Column

    For col = 2 To lastcol Step 2
        For row = 6 To lastrow
            writelog LogInfo, "> empty cell " & row & " " & col & " " & IsEmpty(shGebäude.Cells(row, col)) & " " & shGebäude.Cells(row, col).Address
            If Not IsEmpty(shGebäude.Cells(row, col)) Then
                writelog LogInfo, "> " & shGebäude.Cells(row, col).value
                ' only go if there is something in the cell
                deleteIndexXml row, col
            End If
        Next row
    Next col

    ' ----------------------------- PRINZIPSCHEMAS
    Dim j                    As Long
    Dim i                    As Long
    Dim FileName             As String
    Dim TinLine              As String
    Dim Projektname          As String
    Dim Projektpfad          As String
    Dim EP                   As String
    Dim Gewerk               As String
    Dim GewerkNr             As String
    Dim rng                  As range
    Dim arr()                As Variant
    Dim tmparr()             As Variant


    Set rng = shPData.range("ELE_PRI")
    arr() = rng.Resize(rng.rows.Count, 1)
    tmparr() = RemoveBlanksFromStringArray(arr())

    Projektname = shPData.range("ADM_Projektnummer") & "_" & shPData.range("ADM_Projektbezeichnung")

    For i = LBound(tmparr) To UBound(tmparr)     ' for every Prinzipschema
        Gewerk = rng.Find(tmparr(i)).Offset(0, 1).value
        GewerkNr = rng.Find(tmparr(i)).Offset(0, 2).value
        If Len(GewerkNr) < 2 Then
            GewerkNr = "0" & GewerkNr
        End If
        deleteIndexXml 0, 0, shPData.range("ADM_ProjektpfadCAD").value & "\03_PR\" & GewerkNr & "_" & Gewerk & "\TinPlan_PR_" & Gewerk & ".xml"
    Next i

End Sub

Public Sub deleteIndexXml(ByVal row As Long, ByVal col As Long, Optional ByVal i_xmlfile As String = vbNullString)

    Dim XMLFile              As String
    If i_xmlfile <> vbNullString Then
        XMLFile = i_xmlfile
    Else
        XMLFile = i_xmlfile
    End If

    Dim oXml                 As New MSXML2.DOMDocument60
    oXml.load XMLFile
    Dim nodes                As IXMLDOMNodeList
    Dim node                 As IXMLDOMNode
    Dim root                 As IXMLDOMNode

    Set root = oXml.SelectSingleNode("//tinPlan1")
    Set nodes = oXml.SelectNodes("//tinPlan1/*[contains(local-name(), 'IN')]")
    For Each node In nodes
        root.RemoveChild node
    Next
    On Error Resume Next
    oXml.Save XMLFile
    Set oXml = Nothing
    On Error GoTo 0

End Sub

Public Function getXML(ByVal PCol As Collection) As String
    ' get the xml file path for the genearted PK

    Dim Projektpath          As String
    Dim result               As String

    On Error GoTo ErrHandler
    Projektpath = shPData.range("ADM_ProjektpfadCAD")
    'On Error Resume Next
    Dim buildings            As Boolean
    buildings = Not (shGebäude.range("D1").value = vbNullString)
    If PCol(1) = 0 Then
        ' Plan
        If PCol.Count > 7 Then
            If PCol(7)(1) = "DE" Then
                ' Detail
                result = Projektpath & "\04_DE\TinPlan_DE_" & PCol(15) & ".xml"
            Else
                GoTo Plan
            End If
        Else
Plan:
            If PCol(6)(1) = "TUE" Then
                ' --- TF
                If buildings Then
                    result = Projektpath & "\05_TF\" & PCol(3)(0)(2) & "_" & PCol(3)(0)(1) & "\" & Right$(PCol(3)(1)(2), 2) & "_" & PCol(3)(1)(1) & "\TinPlan_TF_" & PCol(3)(1)(1) & ".xml"
                Else
                    result = Projektpath & "\05_TF\" & Right$(PCol(3)(1)(2), 2) & "_" & PCol(3)(1)(1) & "\TinPlan_TF_" & PCol(3)(1)(1) & ".xml"
                End If
            Else
                ' --- EP
                If buildings Then
                    result = Projektpath & "\01_EP\" & PCol(3)(0)(2) & "_" & PCol(3)(0)(1) & "\" & Right$(PCol(3)(1)(2), 2) & "_" & PCol(3)(1)(1) & "\TinPlan_EP_" & PCol(3)(1)(1) & ".xml"
                Else
                    result = Projektpath & "\01_EP\" & Right$(PCol(3)(1)(2), 2) & "_" & PCol(3)(1)(1) & "\TinPlan_EP_" & PCol(3)(1)(1) & ".xml"
                End If
            End If
        End If
    Else
        ' Prinzip
        result = Projektpath & "\03_PR\" & PCol(6)(2) & "_" & PCol(6)(1) & "\TinPlan_PR_" & PCol(6)(1) & ".xml"
    End If

    getXML = result

    Exit Function

ErrHandler:
    If PCol.Count < 6 Then
        If buildings Then
            result = Projektpath & "\01_EP\" & PCol(3)(0)(2) & "_" & PCol(3)(0)(1) & "\" & Right$(PCol(3)(1)(2), 2) & "_" & PCol(3)(1)(1) & "\TinPlan_EP_" & PCol(3)(1)(1) & ".xml"
        Else
            result = Projektpath & "\01_EP\" & Right$(PCol(3)(1)(2), 2) & "_" & PCol(3)(1)(1) & "\TinPlan_EP_" & PCol(3)(1)(1) & ".xml"
        End If
    End If
    getXML = result

End Function
