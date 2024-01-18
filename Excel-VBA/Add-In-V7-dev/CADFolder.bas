Attribute VB_Name = "CADFolder"

'@Folder("Projekt")
'@Version "Release V1.0.0"

Option Explicit

Public Const OrdnerVorlage   As String = "H:\TinLine\01_Standards\00_Vorlageordner"
Public Const VorlageEPDWG    As String = "H:\TinLine\01_Standards\EP-Vorlage.dwg"
Public Const VorlageEPDWGGEB As String = "H:\TinLine\01_Standards\EP-Vorlage_GEB.dwg"
Public Const VorlagePRDWG    As String = "H:\TinLine\01_Standards\PR-Vorlage.dwg"

Public Sub CreateTinLineProjectFolder(ByVal Pläne As Boolean, ByVal Brandschutz As Boolean, ByVal Türfachplanung As Boolean, ByVal Prinzip As Boolean, ByVal Schemata As Boolean, ByVal SharePointLink As String)
    Globals.shPData.range("ADM_ProjektPfadSharePoint").value = SharePointLink
    Globals.Projekt True
    If Globals.shGebäude Is Nothing Then Globals.SetWBs
    Globals.shProjekt.range("A1").value = False
    Globals.shProjekt.range("A2").value = False
    Globals.shProjekt.range("A3").value = False
    Globals.shProjekt.range("A4").value = False
    Globals.shProjekt.range("A5").value = False
    If Not CreateFoldersTinLine Then Exit Sub
    If Pläne Then CreateFoldersEP
    If Prinzip Then CreateFoldersPR
    If Schemata Then CreateFoldersES
    If Türfachplanung Then CreateFoldersTF
    If Brandschutz Then CreateFoldersBR
    If Pläne Or Prinzip Or Türfachplanung Or Brandschutz Then CreateFolderXRef
    Select Case MsgBox("Pfad im Explorer öffnen?", vbYesNo, "Projekt TinLine erstellt")
    Case vbYes
        ' open explorer
        Shell "explorer.exe" & " " & Globals.Projekt.ProjektOrdnerCAD, vbNormalFocus
        Exit Sub
    Case vbNo
        ' exit sub
        Exit Sub
    End Select
End Sub

Private Function CreateFolderXRef() As Boolean

    Dim fso As New FileSystemObject
    MkDir Globals.Projekt.ProjektOrdnerCAD & "\00_XREF"
    fso.CopyFolder "H:\TinLine\01_Standards\00_Vorlageordner\00_Xref", Globals.Projekt.ProjektOrdnerCAD & "\00_XREF"

End Function

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
    Globals.shProjekt.range("A1").value = True
    Dim Folder               As String
    Folder = Globals.Projekt.ProjektOrdnerCAD & "\01_EP"
    MkDir Folder
    MkDir Globals.Projekt.ProjektOrdnerCAD & "\04_DE"
    GebäudeFolders Folder, "Elektro"
End Sub

Private Sub CreateFoldersPR()
    Globals.shProjekt.range("A2").value = True
    Dim Folder               As String
    Dim UGewerke()           As Variant
    Folder = Globals.Projekt.ProjektOrdnerCAD & "\03_PR"
    MkDir Folder
    UGewerke = getList("ELE_PRI")
    Dim i                    As Long
    Dim Plan                 As IPlankopf
    Dim iStr                 As String
    For i = LBound(UGewerke) To UBound(UGewerke)
        iStr = i - 1
        If Len(iStr) < 2 Then iStr = "0" & iStr

        Set Plan = PlankopfFactory.Create(Projekt:=Projekt, _
                                          ID:=vbNullString, _
                                          TinLineID:=vbNullString, _
                                          Gewerk:="Elektro", _
                                          UnterGewerk:=UGewerke(i), _
                                          Planart:=vbNullString, _
                                          PLANTYP:="PRI", _
                                          Gebäude:=vbNullString, _
                                          Gebäudeteil:=vbNullString, _
                                          Geschoss:=vbNullString, _
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
        MkDir Folder & "\" & iStr & "_" & Plan.UnterGewerkKF
        TinLinePrinzip Plan
        FileCopy VorlagePRDWG, Plan.dwgFile
    Next

End Sub

Private Sub CreateFoldersES()
    Globals.shProjekt.range("A3").value = True
    MkDir Globals.Projekt.ProjektOrdnerCAD & "\02_ES"
End Sub

Private Sub CreateFoldersTF()
    Globals.shProjekt.range("A4").value = True
    Dim Folder               As String
    Folder = Globals.Projekt.ProjektOrdnerCAD & "\05_TF"
    MkDir Folder
    GebäudeFolders Folder, "Türfachplanung"
End Sub

Private Sub CreateFoldersBR()
    Globals.shProjekt.range("A5").value = True
    Dim Folder               As String
    Folder = Globals.Projekt.ProjektOrdnerCAD & "\06_BS"
    MkDir Folder
    GebäudeFolders Folder, "Brandschutzplanung"
End Sub

Public Sub GebäudeFolders(ByVal Folder As String, ByVal Gewerk As String, Optional ByVal MakeDir As Boolean = True)
    ' Folder = 01_EP etc.
    Dim Plan                 As IPlankopf
    Dim buildings            As Boolean
    Dim arrGeb()             As Variant
    Dim larrGeb(2)           As Variant
    Dim arrGes()             As Variant
    Dim larrGes(2)           As Variant
    Dim col                  As Long
    Dim rng                  As range
    Dim arr()                As Variant
    Dim tmparr()             As Variant
    Dim ws                   As Worksheet
    Dim i                    As Long
    Dim ii                   As Long
    Dim lastrow              As Long
    Dim Pfad                 As String

    Set ws = Globals.shGebäude
    arrGeb() = getList("PRO_Gebäude")
    For i = LBound(arrGeb) To UBound(arrGeb)
        'get all Geschosse from the current Building
        larrGeb(0) = arrGeb(i)
        Set rng = ws.range("PRO_Gebäude")
        Set rng = rng.Resize(1, rng.Columns.Count)
        larrGeb(1) = Application.WorksheetFunction.XLookup(larrGeb(0), rng, rng.Offset(1), "-")
        larrGeb(2) = Application.WorksheetFunction.XLookup(larrGeb(0), rng, rng.Offset(2), "-")

        If ws.range("D1").value <> vbNullString Then
            ' mehrere Gebäude, für jedes Gebäude ein Unterordner erstellen und die entsprechenden Etagen einfügen.
            buildings = True
            Pfad = Folder & "\" & larrGeb(2) & "_" & larrGeb(1)
            If MakeDir Then MkDir Pfad
        Else
            ' nur ein Gebäude -> Kein unterordner erstellen
            Pfad = Folder
            buildings = False
        End If
        ' Geschoss

        col = Application.WorksheetFunction.Match(arrGeb(i), ws.range("1:1"), 0)
        lastrow = ws.Cells(ws.rows.Count, col).End(xlUp).row
        Set rng = ws.range(ws.Cells(6, col), ws.Cells(lastrow, col + 1))
        arr() = rng.Resize(rng.rows.Count, 1)
        tmparr() = RemoveBlanksFromStringArray(arr())

        For ii = LBound(tmparr) To UBound(tmparr)
            larrGes(0) = tmparr(ii)
            Dim tmpcol       As Long
            Dim tmplastrow   As Long
            tmpcol = Application.WorksheetFunction.Match(larrGeb(0), ws.range("1:1"), 0)
            tmplastrow = ws.Cells(ws.rows.Count, tmpcol).End(xlUp).row
            Set rng = ws.range(ws.Cells(5, tmpcol), ws.Cells(tmplastrow, tmpcol + 1))
            Set rng = rng.Resize(rng.rows.Count, 1)
            larrGes(1) = rng.Find(larrGes(0)).Offset(0, 1) 'Application.WorksheetFunction.XLookup(larrGes(0), rng, rng.Offset(0, 1), , 0)
            larrGes(2) = ws.Cells(rng.Find(larrGes(0)).row, 1).value

            Set Plan = PlankopfFactory.Create(Projekt:=Projekt, _
                                              ID:=vbNullString, _
                                              TinLineID:=vbNullString, _
                                              Gewerk:=Gewerk, _
                                              UnterGewerk:=vbNullString, _
                                              Planart:=vbNullString, _
                                              PLANTYP:="PLA", _
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

            If MakeDir Then
                If Not CreateObject("Scripting.FileSystemObject").FolderExists(Plan.FolderName) Then
                    MkDir Plan.FolderName
                End If

                If buildings Then
                    FileCopy VorlageEPDWGGEB, Plan.dwgFile
                Else
                    FileCopy VorlageEPDWG, Plan.dwgFile
                End If
            End If

            TinLineFloorXML Plan
            TinLinePlan Plan
        Next ii
    Next i

End Sub

Private Sub TinLineFloorXML(ByVal Plan As IPlankopf)
    Dim oXml                 As New MSXML2.DOMDocument60
    Dim oXsl                 As New MSXML2.DOMDocument60
    oXsl.load XMLVorlage

    Dim NodElement           As IXMLDOMElement
    Dim NodChild             As IXMLDOMElement
    Dim NodGrandChild        As IXMLDOMElement

    ' Standard XML Elemente für TinLine erstellen
    oXml.LoadXML ("<tinPlan1></tinPlan1>")
    Set NodElement = oXml.SelectSingleNode("//tinPlan1")

    Set NodChild = oXml.createElement("Attribut")
    NodElement.appendChild NodChild
    Set NodChild = oXml.createElement("PA")
    NodElement.appendChild NodChild
    Set NodGrandChild = oXml.createElement("Name")
    NodGrandChild.Text = "PA200"
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Bez")
    NodGrandChild.Text = "Gebäude"
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Wert")
    If Plan.Gebäude = "Gesamt" And Globals.shGebäude.range("D1").value = vbNullString Then: NodGrandChild.Text = vbNullString: Else: NodGrandChild.Text = Plan.Gebäude
        NodChild.appendChild NodGrandChild

        ' XML formatieren
        Debug.Print Plan.FolderName & "\TinPlanFloor.xml"
        oXml.Save Plan.FolderName & "\TinPlanFloor.xml"
        oXml.transformNodeToObject oXsl, oXml
        oXml.Save Plan.FolderName & "\TinPlanFloor.xml"

    End Sub

Private Sub TinLinePlan(ByVal Plan As IPlankopf)
    Dim oXml                 As New MSXML2.DOMDocument60
    Dim oXsl                 As New MSXML2.DOMDocument60
    oXsl.load XMLVorlage

    Dim NodElement           As IXMLDOMElement
    Dim NodChild             As IXMLDOMElement
    Dim NodGrandChild        As IXMLDOMElement

    ' Standard XML Elemente für TinLine erstellen
    oXml.LoadXML ("<tinPlan1></tinPlan1>")
    Set NodElement = oXml.SelectSingleNode("tinPlan1")

    Set NodChild = oXml.createElement("Attribut")
    NodElement.appendChild NodChild
    ' Index mit 15 Zeilen
    Set NodChild = oXml.createElement("Index")
    NodElement.appendChild NodChild
    Set NodGrandChild = oXml.createElement("Zeile")
    NodGrandChild.Text = 15
    NodChild.appendChild NodGrandChild
    ' 1 PA Node erstellen damit TinLine was zum Anzeigen hat und nicht nichts zeigt.
    Set NodChild = oXml.createElement("PA")
    NodElement.appendChild NodChild
    Set NodGrandChild = oXml.createElement("Name")
    NodGrandChild.Text = "PA100"
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Bez")
    NodGrandChild.Text = "NICHT VERWENDEN!!!"
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Wert")
    NodGrandChild.Text = vbNullString
    NodChild.appendChild NodGrandChild
    ' XML formatieren
    Debug.Print Plan.XMLFile
    oXml.Save Plan.XMLFile
    oXml.transformNodeToObject oXsl, oXml
    oXml.Save Plan.XMLFile

End Sub

Private Sub TinLinePrinzip(ByVal Plan As IPlankopf)
    Dim oXml                 As New MSXML2.DOMDocument60
    Dim oXsl                 As New MSXML2.DOMDocument60
    oXsl.load XMLVorlage

    Dim NodElement           As IXMLDOMElement
    Dim NodChild             As IXMLDOMElement
    Dim NodGrandChild        As IXMLDOMElement

    ' Standard XML Elemente für TinLine erstellen
    oXml.LoadXML ("<tinPlan1></tinPlan1>")
    Set NodElement = oXml.SelectSingleNode("//tinPlan1")

    Set NodChild = oXml.createElement("Attribut")
    NodElement.appendChild NodChild
    ' Index mit 15 Zeilen
    Set NodChild = oXml.createElement("Index")
    NodElement.appendChild NodChild
    Set NodGrandChild = oXml.createElement("Zeile")
    NodGrandChild.Text = 15
    NodChild.appendChild NodGrandChild
    ' 1 PA Node erstellen damit TinLine was zum Anzeigen hat und nicht nichts zeigt.
    Set NodChild = oXml.createElement("PA")
    NodElement.appendChild NodChild
    Set NodGrandChild = oXml.createElement("Name")
    NodGrandChild.Text = "PA100"
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Bez")
    NodGrandChild.Text = "NICHT VERWENDEN!!!"
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Wert")
    NodGrandChild.Text = vbNullString
    NodChild.appendChild NodGrandChild
    ' XML formatieren
    Debug.Print Plan.XMLFile
    oXml.Save Plan.XMLFile
    oXml.transformNodeToObject oXsl, oXml
    oXml.Save Plan.XMLFile
End Sub

Private Sub TinLineProjectXML()
    Dim oXml                 As New MSXML2.DOMDocument60
    Dim oXsl                 As New MSXML2.DOMDocument60
    oXsl.load XMLVorlage

    Dim NodElement           As IXMLDOMElement
    Dim NodChild             As IXMLDOMElement
    Dim NodGrandChild        As IXMLDOMElement

    If Globals.shPData Is Nothing Then Globals.SetWBs

    ' Standard XML Elemente für TinLine erstellen
    oXml.LoadXML ("<tinPlan1></tinPlan1>")
    Set NodElement = oXml.SelectSingleNode("//tinPlan1")

    Set NodChild = oXml.createElement("Attribut")
    NodElement.appendChild NodChild
    ' Projekt Node
    Set NodChild = oXml.createElement("Project")
    NodElement.appendChild NodChild
    Set NodGrandChild = oXml.createElement("Projektnummer")
    NodGrandChild.Text = Globals.Projekt.Projektnummer
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Projektbeschreibung")
    NodGrandChild.Text = Globals.Projekt.ProjektBezeichnung
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("ProjektMemo")
    NodGrandChild.Text = vbNullString
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Language")
    NodGrandChild.Text = "DE"
    NodChild.appendChild NodGrandChild
    ' Infos für auf den Plankopf
    CreateXmlAttribute "PA01", "Projekt Name", Globals.Projekt.ProjektBezeichnung, "PA", NodChild, oXml, NodElement
    CreateXmlAttribute "PA02", "Projekt Adresse [Strasse]", Globals.Projekt.Projektadresse.Strasse, "PA", NodChild, oXml, NodElement
    CreateXmlAttribute "PA03", "Projekt Adresse [PLZ]", Globals.Projekt.Projektadresse.PLZ, "PA", NodChild, oXml, NodElement
    CreateXmlAttribute "PA04", "Projekt Adresse [Ort]", Globals.Projekt.Projektadresse.Ort, "PA", NodChild, oXml, NodElement
    CreateXmlAttribute "PA05", "Projektnummer", Globals.Projekt.Projektnummer, "PA", NodChild, oXml, NodElement
    CreateXmlAttribute "PA06", "Projektphase", Globals.Projekt.Projektphase, "PA", NodChild, oXml, NodElement
    ' XML formatieren
    oXml.Save Globals.Projekt.ProjektXML
    oXml.transformNodeToObject oXsl, oXml
    oXml.Save Globals.Projekt.ProjektXML
End Sub

Public Sub RenameFolders()
    Dim fso As New FileSystemObject
    Dim root As String
    Dim PlantypFolder As Object
    Dim GeschossFolder As Object
    Dim GebäudeFolder As Object
    Dim oldfolder As String
    root = Globals.Projekt.ProjektOrdnerCAD
    If Globals.shGebäude.range("D1").value = vbNullString Then
        For Each PlantypFolder In fso.GetFolder(root).SubFolders
            If PlantypFolder.Name = "01_EP" Or PlantypFolder.Name = "05_TF" Or PlantypFolder.Name = "06_BS" Then
                For Each GeschossFolder In PlantypFolder.SubFolders
                    oldfolder = GeschossFolder.Name
                    If Len(Split(GeschossFolder.Name, "_")(0)) < 3 Then
                        GeschossFolder.Name = "0" & GeschossFolder.Name
                        writelog LogInfo, "Renaming Folder " & oldfolder & " to " & GeschossFolder.Name
                    End If
                Next GeschossFolder
            End If
        Next
    Else
        ' Mehrere Gebäude => eine Ebene tiefer für Ordner umbenennen
        For Each PlantypFolder In fso.GetFolder(root).SubFolders
            If PlantypFolder.Name = "01_EP" Or PlantypFolder.Name = "05_TF" Or PlantypFolder.Name = "06_BS" Then
                For Each GebäudeFolder In PlantypFolder.SubFolders
                    For Each GeschossFolder In PlantypFolder.SubFolders
                        oldfolder = GeschossFolder.Name
                        If Len(Split(GeschossFolder.Name, "_")(0)) < 3 Then
                            GeschossFolder.Name = "0" & GeschossFolder.Name
                            writelog LogInfo, "Renaming Folder " & oldfolder & " to " & GeschossFolder.Name
                        End If
                    Next GeschossFolder
                Next GebäudeFolder
            End If
        Next
    End If
End Sub

