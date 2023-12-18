Attribute VB_Name = "CADFolder"
'@Folder("Projekt")
Option Explicit

Public Sub CreateTinLineProjectFolder(ByVal Pläne As Boolean, ByVal Brandschutz As Boolean, ByVal Türfachplanung As Boolean, ByVal Prinzip As Boolean, ByVal Schemata As Boolean)

    If Globals.shGebäude Is Nothing Then Globals.SetWBs
    CreateFoldersTinLine
    If Pläne Then CreateFoldersEP
    If Prinzip Then CreateFoldersPR
    If Schemata Then CreateFoldersES
    If Türfachplanung Then CreateFoldersTF
    If Brandschutz Then CreateFoldersBR

End Sub

Private Sub CreateFoldersTinLine()
    MkDir Globals.Projekt.ProjektOrdnerCAD
    MkDir Globals.Projekt.ProjektOrdnerCAD & "\99 TinConfiguration"
    MkDir Globals.Projekt.ProjektOrdnerCAD & "\99 Planlisten"

    TinLineProjectXML
End Sub

Private Sub CreateFoldersEP()
    Dim Folder               As String
    Folder = Globals.Projekt.ProjektOrdnerCAD & "\01_EP"
    MkDir Globals.Projekt.ProjektOrdnerCAD & "\00_XREF"
    MkDir Folder
    MkDir Globals.Projekt.ProjektOrdnerCAD & "\04_DE"
    GebäudeFolders Folder, "Elektro"
End Sub

Private Sub CreateFoldersPR()
    Dim Folder               As String
    Dim UGewerke()           As Variant
    Folder = Globals.Projekt.ProjektOrdnerCAD & "\03_PR"
    MkDir Folder
    UGewerke = getList("ELE_PRI")
End Sub

Private Sub CreateFoldersES()
    MkDir Globals.Projekt.ProjektOrdnerCAD & "\02_ES"
End Sub

Private Sub CreateFoldersTF()
    Dim Folder               As String
    Folder = Globals.Projekt.ProjektOrdnerCAD & "\05_TF"
    MkDir Folder
    GebäudeFolders Folder, "Türfachplanung"
End Sub

Private Sub CreateFoldersBR()
    Dim Folder               As String
    Folder = Globals.Projekt.ProjektOrdnerCAD & "\06_BR"
    MkDir Folder
    GebäudeFolders Folder, "Brandschutzplanung"
End Sub

Private Sub GebäudeFolders(ByVal Folder As String, ByVal Gewerk As String)
    ' Folder = 01_EP etc.
    Dim Plan                 As IPlankopf
    Dim buildings            As Boolean
    Dim arrGeb() As Variant
    Dim larrGeb(2) As Variant
    Dim arrGes() As Variant
    Dim larrGes(2) As Variant
    Dim col
    Dim rng                  As range
    Dim arr()                As Variant
    Dim tmparr()             As Variant
    Dim ws                   As Worksheet
    Dim i                    As Long
    Dim ii                   As Long
    Dim lastrow              As Long
    Dim Pfad                 As String
    Dim Pfad2                As String

    Set ws = Globals.shGebäude
    arrGeb() = getList("PRO_Gebäude")
    For i = LBound(arrGeb) + 1 To UBound(arrGeb)
        'get all Geschosse from the current Building
        larrGeb(0) = arrGeb(i)
        Set rng = ws.range("PRO_Gebäude")
        Set rng = rng.Resize(1, rng.Columns.Count)
        larrGeb(1) = Application.WorksheetFunction.XLookup(larrGeb(0), rng, rng.Offset(1), "-")
        larrGeb(2) = Application.WorksheetFunction.XLookup(larrGeb(0), rng, rng.Offset(2), "-")

        If ws.range("D1").Value <> "" Then
            ' mehrere Gebäude, für jedes Gebäude ein Unterordner erstellen und die entsprechenden Etagen einfügen.
            buildings = True
            Pfad = Folder & "\" & larrGeb(2) & "_" & larrGeb(1)
            MkDir Pfad
        Else
            ' nur ein Gebäude -> Kein unterordner erstellen
            Pfad = Folder
        End If
        ' Geschoss

        col = Application.WorksheetFunction.Match(arrGeb(i), ws.range("1:1"), 0)
        lastrow = ws.Cells(rows.Count, col).End(xlUp).row
        Set rng = ws.range(ws.Cells(6, col), ws.Cells(lastrow, col + 1))
        arr() = rng.Resize(rng.rows.Count, 1)
        tmparr() = RemoveBlanksFromStringArray(arr())

        For ii = LBound(tmparr) To UBound(tmparr)
            larrGes(0) = tmparr(ii)
            Dim tmpcol
            Dim tmplastrow
            tmpcol = Application.WorksheetFunction.Match(larrGeb(0), ws.range("1:1"), 0)
            tmplastrow = ws.Cells(rows.Count, tmpcol).End(xlUp).row
            Set rng = ws.range(ws.Cells(5, tmpcol), ws.Cells(tmplastrow, tmpcol + 1))
            Set rng = rng.Resize(rng.rows.Count, 1)
            larrGes(1) = Application.WorksheetFunction.XLookup(larrGes(0), rng, rng.Offset(0, 1), , 0)
            larrGes(2) = ws.Cells(rng.Find(larrGes(0)).row, 1).Value

            Pfad2 = Pfad & "\" & larrGes(2) & "_" & larrGes(1)
            MkDir Pfad2
            If buildings Then
                FileCopy VorlageEPDWGGEB, Pfad2 & "\" & Right(Folder, 2) & "_" & larrGes(1) & ".dwg"
            Else
                FileCopy VorlageEPDWG, Pfad2 & "\" & Right(Folder, 2) & "_" & larrGes(1) & ".dwg"
            End If
            
            Set Plan = PlankopfFactory.Create(Projekt:=Projekt, _
           ID:=vbNullString, _
           TinLineID:=vbNullString, _
           Gewerk:=Gewerk, _
           UnterGewerk:=vbNullString, _
           Planart:=vbNullString, _
           Plantyp:="PLA", _
           Gebäude:=CStr(larrGeb(0)), _
           Gebäudeteil:=vbNullString, _
           Geschoss:=CStr(larrGes(0)), _
           Planüberschrift:=vbNullString, _
           Format:=vbNullString, _
           Masstab:=vbNullString, _
           Stand:=vbNullString, _
           GezeichnetPerson:=vbNullString, _
           GezeichnetDatum:=vbNullString, _
           GeprüftPerson:=vbNullString, _
           GeprüftDatum:=vbNullString, _
           SkipValidation:=True, _
           CustomÜberschrift:=False _
           )
            TinLineFloorXML Plan
            TinLinePlan Plan
        Next ii
    Next i

End Sub

Private Sub TinLineFloorXML(ByRef Plan As IPlankopf)
    Dim oXml                 As New MSXML2.DOMDocument60
    Dim oXsl                 As New MSXML2.DOMDocument60
    oXsl.load XMLVorlage

    Dim nodElement           As IXMLDOMElement
    Dim nodChild             As IXMLDOMElement
    Dim nodGrandChild        As IXMLDOMElement

    ' Standard XML Elemente für TinLine erstellen
    oXml.LoadXML ("<tinPlan1></tinPlan1>")
    Set nodElement = oXml.SelectSingleNode("//tinPlan1")

    Set nodChild = oXml.createElement("Attribut")
    nodElement.appendChild nodChild
    Set nodChild = oXml.createElement("PA")
    nodElement.appendChild nodChild
    Set nodGrandChild = oXml.createElement("Name")
    nodGrandChild.text = "PA200"
    nodChild.appendChild nodGrandChild
    Set nodGrandChild = oXml.createElement("Bez")
    nodGrandChild.text = "Gebäude"
    nodChild.appendChild nodGrandChild
    Set nodGrandChild = oXml.createElement("Wert")
    nodGrandChild.text = Plan.Gebäude
    nodChild.appendChild nodGrandChild

    ' XML formatieren
    oXml.Save Plan.XMLFile
    oXml.transformNodeToObject oXsl, oXml
    oXml.Save Plan.XMLFile

End Sub

Private Sub TinLinePlan(ByVal Plan As IPlankopf)
    Dim oXml                 As New MSXML2.DOMDocument60
    Dim oXsl                 As New MSXML2.DOMDocument60
    oXsl.load XMLVorlage

    Dim nodElement           As IXMLDOMElement
    Dim nodChild             As IXMLDOMElement
    Dim nodGrandChild        As IXMLDOMElement

    ' Standard XML Elemente für TinLine erstellen
    oXml.LoadXML ("<tinPlan1></tinPlan1>")
    Set nodElement = oXml.SelectSingleNode("//tinPlan1")

    Set nodChild = oXml.createElement("Attribut")
    nodElement.appendChild nodChild
    ' Index mit 15 Zeilen
    Set nodChild = oXml.createElement("Index")
    nodElement.appendChild nodChild
    Set nodGrandChild = oXml.createElement("Zeile")
    nodGrandChild.text = 15
    nodChild.appendChild nodGrandChild
    ' TODO prüfen ob diese Nodes gebraucht werden damit TinLine funktioniert.
    'Set nodChild = oXml.createElement("PA")
    'nodElement.appendChild nodChild
    'Set nodGrandChild = oXml.createElement("Name")
    'nodGrandChild.text = "PA100"
    'nodChild.appendChild nodGrandChild
    'Set nodGrandChild = oXml.createElement("Bez")
    'nodGrandChild.text = "NICHT VERWENDEN!!!"
    'nodChild.appendChild nodGrandChild
    'Set nodGrandChild = oXml.createElement("Wert")
    'nodGrandChild.text = ""
    'nodChild.appendChild nodGrandChild
    ' XML formatieren
    oXml.Save Plan.XMLFile
    oXml.transformNodeToObject oXsl, oXml
    oXml.Save Plan.XMLFile

End Sub

Private Sub TinLinePrinzip(ByVal Plan As IPlankopf)
    Dim oXml                 As New MSXML2.DOMDocument60
    Dim oXsl                 As New MSXML2.DOMDocument60
    oXsl.load XMLVorlage

    Dim nodElement           As IXMLDOMElement
    Dim nodChild             As IXMLDOMElement
    Dim nodGrandChild        As IXMLDOMElement

    ' Standard XML Elemente für TinLine erstellen
    oXml.LoadXML ("<tinPlan1></tinPlan1>")
    Set nodElement = oXml.SelectSingleNode("//tinPlan1")

    Set nodChild = oXml.createElement("Attribut")
    nodElement.appendChild nodChild
    ' Index mit 15 Zeilen
    Set nodChild = oXml.createElement("Index")
    nodElement.appendChild nodChild
    Set nodGrandChild = oXml.createElement("Zeile")
    nodGrandChild.text = 15
    nodChild.appendChild nodGrandChild
    ' TODO prüfen ob diese Nodes gebraucht werden damit TinLine funktioniert.
    'Set nodChild = oXml.createElement("PA")
    'nodElement.appendChild nodChild
    'Set nodGrandChild = oXml.createElement("Name")
    'nodGrandChild.text = "PA100"
    'nodChild.appendChild nodGrandChild
    'Set nodGrandChild = oXml.createElement("Bez")
    'nodGrandChild.text = "NICHT VERWENDEN!!!"
    'nodChild.appendChild nodGrandChild
    'Set nodGrandChild = oXml.createElement("Wert")
    'nodGrandChild.text = ""
    'nodChild.appendChild nodGrandChild
    ' XML formatieren
    oXml.Save Plan.XMLFile
    oXml.transformNodeToObject oXsl, oXml
    oXml.Save Plan.XMLFile
End Sub

Private Sub TinLineProjectXML()
    Dim oXml                 As New MSXML2.DOMDocument60
    Dim oXsl                 As New MSXML2.DOMDocument60
    oXsl.load XMLVorlage

    Dim nodElement           As IXMLDOMElement
    Dim nodChild             As IXMLDOMElement
    Dim nodGrandChild        As IXMLDOMElement

    If Globals.shPData Is Nothing Then Globals.SetWBs

    ' Standard XML Elemente für TinLine erstellen
    oXml.LoadXML ("<tinPlan1></tinPlan1>")
    Set nodElement = oXml.SelectSingleNode("//tinPlan1")

    Set nodChild = oXml.createElement("Attribut")
    nodElement.appendChild nodChild
    ' Projekt Node
    Set nodChild = oXml.createElement("Project")
    nodElement.appendChild nodChild
    Set nodGrandChild = oXml.createElement("Projektnummer")
    nodGrandChild.text = Globals.Projekt.Projektnummer
    nodChild.appendChild nodGrandChild
    Set nodGrandChild = oXml.createElement("Projektbeschreibung")
    nodGrandChild.text = Globals.Projekt.ProjektBezeichnung
    nodChild.appendChild nodGrandChild
    Set nodGrandChild = oXml.createElement("ProjektMemo")
    nodGrandChild.text = ""
    nodChild.appendChild nodGrandChild
    Set nodGrandChild = oXml.createElement("Language")
    nodGrandChild.text = "DE"
    nodChild.appendChild nodGrandChild
    ' Infos für auf den Plankopf
    CreateXmlAttribute "PA01", "Projekt Name", Globals.Projekt.ProjektBezeichnung, "PA", nodChild, oXml, nodElement
    CreateXmlAttribute "PA02", "Projekt Adresse [Strasse]", Globals.Projekt.Projektadresse.Strasse, "PA", nodChild, oXml, nodElement
    CreateXmlAttribute "PA03", "Projekt Adresse [PLZ]", Globals.Projekt.Projektadresse.PLZ, "PA", nodChild, oXml, nodElement
    CreateXmlAttribute "PA04", "Projekt Adresse [Ort]", Globals.Projekt.Projektadresse.Ort, "PA", nodChild, oXml, nodElement
    CreateXmlAttribute "PA05", "Projektnummer", Globals.Projekt.Projektnummer, "PA", nodChild, oXml, nodElement
    CreateXmlAttribute "PA06", "Projektphase", Globals.Projekt.Projektphase, "PA", nodChild, oXml, nodElement
    ' XML formatieren
    oXml.Save Globals.Projekt.ProjektXML
    oXml.transformNodeToObject oXsl, oXml
    oXml.Save Globals.Projekt.ProjektXML
End Sub

