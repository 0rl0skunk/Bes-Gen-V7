Attribute VB_Name = "CADFolder"
'@Folder("Projekt")
Option Explicit

Global Const OrdnerVorlage As String = "H:\TinLine\01_Standards\00_Vorlageordner"
Global Const VorlageEPDWG As String = "H:\TinLine\01_Standards\EP-Vorlage.dwg"
Global Const VorlageEPDWGGEB As String = "H:\TinLine\01_Standards\EP-Vorlage_GEB.dwg"
Global Const VorlagePRDWG As String = "H:\TinLine\01_Standards\PR-Vorlage.dwg"

Public Sub CreateTinLineProjectFolder(ByVal Pl�ne As Boolean, ByVal Brandschutz As Boolean, ByVal T�rfachplanung As Boolean, ByVal Prinzip As Boolean, ByVal Schemata As Boolean)
    
    Globals.Projekt True
    If Globals.shGeb�ude Is Nothing Then Globals.SetWBs
    If Not CreateFoldersTinLine Then Exit Sub
    If Pl�ne Then CreateFoldersEP
    If Prinzip Then CreateFoldersPR
    If Schemata Then CreateFoldersES
    If T�rfachplanung Then CreateFoldersTF
    If Brandschutz Then CreateFoldersBR
    Select Case MsgBox("Pfad im Explorer �ffnen?", vbYesNo, "Projekt TinLine erstellt")
        Case vbYes
            ' open explorer
            Shell "explorer.exe" & " " & Globals.Projekt.ProjektOrdnerCAD, vbNormalFocus
            Exit Sub
        Case vbNo
            ' exit sub
            Exit Sub
    End Select
End Sub

Private Function CreateFoldersTinLine() As Boolean
    On Error GoTo ErrHandler
    MkDir Globals.Projekt.ProjektOrdnerCAD
    MkDir Globals.Projekt.ProjektOrdnerCAD & "\99 TinConfiguration"
    MkDir Globals.Projekt.ProjektOrdnerCAD & "\99 Planlisten"

    TinLineProjectXML
    Globals.shPData.range("ADM_ProjektPfadCAD").value = Globals.Projekt.ProjektOrdnerCAD
    CreateFoldersTinLine = True
    Exit Function
ErrHandler:
    Select Case err.Number
    Case 75
        MsgBox "Der Ordner besteht bereits!" & vbNewLine & "Stell sicher, dass die Projektnummer korrekt eingetragen wurde." & vbNewLine & vbNewLine & "Wenn die Projektnummer etc. korrekt eingetragen wurde, melde dich beim QS-Verantwortlichen!", vbCritical, "Projekt bereits Vorhanden"
        CreateFoldersTinLine = False
    End Select
End Function

Private Sub CreateFoldersEP()
    Dim folder               As String
    folder = Globals.Projekt.ProjektOrdnerCAD & "\01_EP"
    MkDir Globals.Projekt.ProjektOrdnerCAD & "\00_XREF"
    MkDir folder
    MkDir Globals.Projekt.ProjektOrdnerCAD & "\04_DE"
    Geb�udeFolders folder, "Elektro"
End Sub

Private Sub CreateFoldersPR()
    Dim folder               As String
    Dim UGewerke()           As Variant
    folder = Globals.Projekt.ProjektOrdnerCAD & "\03_PR"
    MkDir folder
    UGewerke = getList("ELE_PRI")
    Dim i As Long
    Dim Plan As IPlankopf
    Dim iStr As String
    For i = LBound(UGewerke) To UBound(UGewerke)
        iStr = i - 1
        If Len(iStr) < 2 Then iStr = "0" & iStr
        
        Set Plan = PlankopfFactory.Create(Projekt:=Projekt, _
           ID:=vbNullString, _
           TinLineID:=vbNullString, _
           Gewerk:="Elektro", _
           UnterGewerk:=UGewerke(i), _
           Planart:=vbNullString, _
           Plantyp:="PRI", _
           Geb�ude:=vbNullString, _
           Geb�udeteil:=vbNullString, _
           Geschoss:=vbNullString, _
           Plan�berschrift:=vbNullString, _
           Format:=vbNullString, _
           Masstab:=vbNullString, _
           Stand:=vbNullString, _
           GezeichnetPerson:=vbNullString, _
           GezeichnetDatum:=vbNullString, _
           Gepr�ftPerson:=vbNullString, _
           Gepr�ftDatum:=vbNullString, _
           SkipValidation:=True, _
           Custom�berschrift:=False _
           )
           MkDir folder & "\" & iStr & "_" & Plan.UnterGewerkKF
            TinLinePrinzip Plan
            FileCopy VorlagePRDWG, Plan.dwgFile
    Next
    
End Sub

Private Sub CreateFoldersES()
    MkDir Globals.Projekt.ProjektOrdnerCAD & "\02_ES"
End Sub

Private Sub CreateFoldersTF()
    Dim folder               As String
    folder = Globals.Projekt.ProjektOrdnerCAD & "\05_TF"
    MkDir folder
    Geb�udeFolders folder, "T�rfachplanung"
End Sub

Private Sub CreateFoldersBR()
    Dim folder               As String
    folder = Globals.Projekt.ProjektOrdnerCAD & "\06_BR"
    MkDir folder
    Geb�udeFolders folder, "Brandschutzplanung"
End Sub

Private Sub Geb�udeFolders(ByVal folder As String, ByVal Gewerk As String)
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

    Set ws = Globals.shGeb�ude
    arrGeb() = getList("PRO_Geb�ude")
    For i = LBound(arrGeb) To UBound(arrGeb)
        'get all Geschosse from the current Building
        larrGeb(0) = arrGeb(i)
        Set rng = ws.range("PRO_Geb�ude")
        Set rng = rng.Resize(1, rng.Columns.Count)
        larrGeb(1) = Application.WorksheetFunction.XLookup(larrGeb(0), rng, rng.Offset(1), "-")
        larrGeb(2) = Application.WorksheetFunction.XLookup(larrGeb(0), rng, rng.Offset(2), "-")

        If ws.range("D1").value <> "" Then
            ' mehrere Geb�ude, f�r jedes Geb�ude ein Unterordner erstellen und die entsprechenden Etagen einf�gen.
            buildings = True
            Pfad = folder & "\" & larrGeb(2) & "_" & larrGeb(1)
            MkDir Pfad
        Else
            ' nur ein Geb�ude -> Kein unterordner erstellen
            Pfad = folder
            buildings = False
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
            larrGes(2) = ws.Cells(rng.Find(larrGes(0)).row, 1).value
            
            Set Plan = PlankopfFactory.Create(Projekt:=Projekt, _
           ID:=vbNullString, _
           TinLineID:=vbNullString, _
           Gewerk:=Gewerk, _
           UnterGewerk:=vbNullString, _
           Planart:=vbNullString, _
           Plantyp:="PLA", _
           Geb�ude:=CStr(larrGeb(0)), _
           Geb�udeteil:=vbNullString, _
           Geschoss:=CStr(larrGes(0)), _
           Plan�berschrift:=vbNullString, _
           Format:=vbNullString, _
           Masstab:=vbNullString, _
           Stand:=vbNullString, _
           GezeichnetPerson:=vbNullString, _
           GezeichnetDatum:=vbNullString, _
           Gepr�ftPerson:=vbNullString, _
           Gepr�ftDatum:=vbNullString, _
           SkipValidation:=True, _
           Custom�berschrift:=False _
           )
            
            If Not CreateObject("Scripting.FileSystemObject").FolderExists(Plan.FolderName) Then
            MkDir Plan.FolderName
            End If
            
            If buildings Then
            FileCopy VorlageEPDWGGEB, Plan.dwgFile
            Else
            FileCopy VorlageEPDWG, Plan.dwgFile
            End If
            
            TinLineFloorXML Plan
            TinLinePlan Plan
        Next ii
    Next i

End Sub

Private Sub TinLineFloorXML(ByRef Plan As IPlankopf)
    Dim oXml                 As New MSXML2.DOMDocument60
    Dim oXsl                 As New MSXML2.DOMDocument60
    oXsl.load XMLVorlage

    Dim NodElement           As IXMLDOMElement
    Dim NodChild             As IXMLDOMElement
    Dim NodGrandChild        As IXMLDOMElement

    ' Standard XML Elemente f�r TinLine erstellen
    oXml.LoadXML ("<tinPlan1></tinPlan1>")
    Set NodElement = oXml.SelectSingleNode("//tinPlan1")

    Set NodChild = oXml.createElement("Attribut")
    NodElement.appendChild NodChild
    Set NodChild = oXml.createElement("PA")
    NodElement.appendChild NodChild
    Set NodGrandChild = oXml.createElement("Name")
    NodGrandChild.text = "PA200"
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Bez")
    NodGrandChild.text = "Geb�ude"
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Wert")
    If Plan.Geb�ude = "Gesamt" And Globals.shGeb�ude.range("D1").value = vbNullString Then: NodGrandChild.text = vbNullString: Else: NodGrandChild.text = Plan.Geb�ude
    NodChild.appendChild NodGrandChild

    ' XML formatieren
    Debug.Print Plan.FolderName & "\TinPlanFloor.xml"
    oXml.save Plan.FolderName & "\TinPlanFloor.xml"
    oXml.transformNodeToObject oXsl, oXml
    oXml.save Plan.FolderName & "\TinPlanFloor.xml"

End Sub

Private Sub TinLinePlan(ByVal Plan As IPlankopf)
    Dim oXml                 As New MSXML2.DOMDocument60
    Dim oXsl                 As New MSXML2.DOMDocument60
    oXsl.load XMLVorlage

    Dim NodElement           As IXMLDOMElement
    Dim NodChild             As IXMLDOMElement
    Dim NodGrandChild        As IXMLDOMElement

    ' Standard XML Elemente f�r TinLine erstellen
    oXml.LoadXML ("<tinPlan1></tinPlan1>")
    Set NodElement = oXml.SelectSingleNode("//tinPlan1")

    Set NodChild = oXml.createElement("Attribut")
    NodElement.appendChild NodChild
    ' Index mit 15 Zeilen
    Set NodChild = oXml.createElement("Index")
    NodElement.appendChild NodChild
    Set NodGrandChild = oXml.createElement("Zeile")
    NodGrandChild.text = 15
    NodChild.appendChild NodGrandChild
    ' 1 PA Node erstellen damit TinLine was zum Anzeigen hat und nicht nichts zeigt.
    Set NodChild = oXml.createElement("PA")
    NodElement.appendChild NodChild
    Set NodGrandChild = oXml.createElement("Name")
    NodGrandChild.text = "PA100"
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Bez")
    NodGrandChild.text = "NICHT VERWENDEN!!!"
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Wert")
    NodGrandChild.text = ""
    NodChild.appendChild NodGrandChild
    ' XML formatieren
    Debug.Print Plan.XMLFile
    oXml.save Plan.XMLFile
    oXml.transformNodeToObject oXsl, oXml
    oXml.save Plan.XMLFile

End Sub

Private Sub TinLinePrinzip(ByVal Plan As IPlankopf)
    Dim oXml                 As New MSXML2.DOMDocument60
    Dim oXsl                 As New MSXML2.DOMDocument60
    oXsl.load XMLVorlage

    Dim NodElement           As IXMLDOMElement
    Dim NodChild             As IXMLDOMElement
    Dim NodGrandChild        As IXMLDOMElement

    ' Standard XML Elemente f�r TinLine erstellen
    oXml.LoadXML ("<tinPlan1></tinPlan1>")
    Set NodElement = oXml.SelectSingleNode("//tinPlan1")

    Set NodChild = oXml.createElement("Attribut")
    NodElement.appendChild NodChild
    ' Index mit 15 Zeilen
    Set NodChild = oXml.createElement("Index")
    NodElement.appendChild NodChild
    Set NodGrandChild = oXml.createElement("Zeile")
    NodGrandChild.text = 15
    NodChild.appendChild NodGrandChild
    ' 1 PA Node erstellen damit TinLine was zum Anzeigen hat und nicht nichts zeigt.
    Set NodChild = oXml.createElement("PA")
    NodElement.appendChild NodChild
    Set NodGrandChild = oXml.createElement("Name")
    NodGrandChild.text = "PA100"
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Bez")
    NodGrandChild.text = "NICHT VERWENDEN!!!"
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Wert")
    NodGrandChild.text = ""
    NodChild.appendChild NodGrandChild
    ' XML formatieren
    Debug.Print Plan.XMLFile
    oXml.save Plan.XMLFile
    oXml.transformNodeToObject oXsl, oXml
    oXml.save Plan.XMLFile
End Sub

Private Sub TinLineProjectXML()
    Dim oXml                 As New MSXML2.DOMDocument60
    Dim oXsl                 As New MSXML2.DOMDocument60
    oXsl.load XMLVorlage

    Dim NodElement           As IXMLDOMElement
    Dim NodChild             As IXMLDOMElement
    Dim NodGrandChild        As IXMLDOMElement

    If Globals.shPData Is Nothing Then Globals.SetWBs

    ' Standard XML Elemente f�r TinLine erstellen
    oXml.LoadXML ("<tinPlan1></tinPlan1>")
    Set NodElement = oXml.SelectSingleNode("//tinPlan1")

    Set NodChild = oXml.createElement("Attribut")
    NodElement.appendChild NodChild
    ' Projekt Node
    Set NodChild = oXml.createElement("Project")
    NodElement.appendChild NodChild
    Set NodGrandChild = oXml.createElement("Projektnummer")
    NodGrandChild.text = Globals.Projekt.Projektnummer
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Projektbeschreibung")
    NodGrandChild.text = Globals.Projekt.ProjektBezeichnung
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("ProjektMemo")
    NodGrandChild.text = ""
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Language")
    NodGrandChild.text = "DE"
    NodChild.appendChild NodGrandChild
    ' Infos f�r auf den Plankopf
    CreateXmlAttribute "PA01", "Projekt Name", Globals.Projekt.ProjektBezeichnung, "PA", NodChild, oXml, NodElement
    CreateXmlAttribute "PA02", "Projekt Adresse [Strasse]", Globals.Projekt.Projektadresse.Strasse, "PA", NodChild, oXml, NodElement
    CreateXmlAttribute "PA03", "Projekt Adresse [PLZ]", Globals.Projekt.Projektadresse.PLZ, "PA", NodChild, oXml, NodElement
    CreateXmlAttribute "PA04", "Projekt Adresse [Ort]", Globals.Projekt.Projektadresse.Ort, "PA", NodChild, oXml, NodElement
    CreateXmlAttribute "PA05", "Projektnummer", Globals.Projekt.Projektnummer, "PA", NodChild, oXml, NodElement
    CreateXmlAttribute "PA06", "Projektphase", Globals.Projekt.Projektphase, "PA", NodChild, oXml, NodElement
    ' XML formatieren
    oXml.save Globals.Projekt.ProjektXML
    oXml.transformNodeToObject oXsl, oXml
    oXml.save Globals.Projekt.ProjektXML
End Sub

