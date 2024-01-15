Attribute VB_Name = "PlankopfFactory"
Attribute VB_Description = "Erstellt ein Plankopf-Objekt von welchem die daten einfach ausgelesen werden können."

'@IgnoreModule VariableNotUsed
'@Folder "Plankopf"
'@ModuleDescription "Erstellt ein Plankopf-Objekt von welchem die daten einfach ausgelesen werden können."
'@Version "Release V1.0.0"

Option Explicit

Private oXml                 As New MSXML2.DOMDocument60
Private oXsl                 As New MSXML2.DOMDocument60

Private NodElement           As IXMLDOMElement
Private NodChild             As IXMLDOMElement
Private NodGrandChild        As IXMLDOMElement

Private PKNr                 As Long

Public Function Create( _
       ByVal Projekt As IProjekt, _
       ByVal GezeichnetPerson As String, _
       ByVal GezeichnetDatum As String, _
       ByVal GeprüftPerson As String, _
       ByVal GeprüftDatum As String, _
       ByVal Gebäude As String, _
       ByVal Gebäudeteil As String, _
       ByVal Geschoss As String, _
       ByVal Gewerk As String, _
       ByVal UnterGewerk As String, _
       ByVal Format As String, _
       ByVal Masstab As String, _
       ByVal Stand As String, _
       ByVal Planart As String, _
       Optional ByVal PLANTYP As String, _
       Optional ByVal TinLineID As String, _
       Optional ByVal SkipValidation As Boolean = False, _
       Optional ByVal Planüberschrift As String = "NEW", _
       Optional ByVal ID As String = "NEW", _
       Optional ByVal CustomÜberschrift As Boolean = False, _
       Optional ByVal AnlageTyp As String, _
       Optional ByVal AnlageNummer As String _
       ) As IPlankopf

    Dim NewPlankopf          As New Plankopf
    If NewPlankopf.Filldata( _
       Projekt:=Projekt, _
       GezeichnetPerson:=GezeichnetPerson, _
       GezeichnetDatum:=Replace(GezeichnetDatum, "/", "."), _
       GeprüftPerson:=GeprüftPerson, _
       GeprüftDatum:=Replace(GeprüftDatum, "/", "."), _
       Gebäude:=Gebäude, _
       Gebäudeteil:=Gebäudeteil, _
       Geschoss:=Geschoss, _
       Gewerk:=Gewerk, _
       UnterGewerk:=UnterGewerk, _
       Format:=Format, _
       Masstab:=Masstab, _
       Stand:=Stand, _
       Planart:=Planart, _
       PLANTYP:=PLANTYP, _
       TinLineID:=TinLineID, _
       SkipValidation:=SkipValidation, _
       Planüberschrift:=Planüberschrift, _
       ID:=ID, _
       CustomÜberschrift:=CustomÜberschrift, _
       AnlageTyp:=AnlageTyp, _
       AnlageNummer:=AnlageNummer _
                      ) Then
        Dim row              As Long
        Set Create = NewPlankopf
        IndexFactory.GetIndexes Create
        Exit Function
    Else
        Dim frm              As New UserFormMessage
        frm.Typ typError, "Es wurde kein Plankopf erstellt!"
        frm.Show 1
    End If

End Function

Public Function LoadFromDataBase(ByVal row As Long) As IPlankopf

    Dim NewPlankopf          As Plankopf:    Set NewPlankopf = New Plankopf
    Dim ws                   As Worksheet:    Set ws = Globals.shStoreData
    With ws
        If NewPlankopf.Filldata( _
           Projekt:=Projekt, _
           ID:=.Cells(row, 1).value, _
           TinLineID:=.Cells(row, 2).value, _
           Gewerk:=.Cells(row, 3).value, _
           UnterGewerk:=.Cells(row, 4).value, _
           Planart:=.Cells(row, 5).value, _
           PLANTYP:=.Cells(row, 6).value, _
           Gebäude:=.Cells(row, 7).value, _
           Gebäudeteil:=.Cells(row, 8).value, _
           Geschoss:=.Cells(row, 9).value, _
           Planüberschrift:=.Cells(row, 13).value, _
           Format:=.Cells(row, 15).value, _
           Masstab:=.Cells(row, 16).value, _
           Stand:=.Cells(row, 17).value, _
           GezeichnetPerson:=.Cells(row, 18).value, _
           GezeichnetDatum:=.Cells(row, 19).value, _
           GeprüftPerson:=.Cells(row, 20).value, _
           GeprüftDatum:=.Cells(row, 21).value, _
           SkipValidation:=False, _
           CustomÜberschrift:=.Cells(row, 10).value, _
           AnlageTyp:=.Cells(row, 23).value, _
           AnlageNummer:=.Cells(row, 24).value _
                          ) Then
            Set LoadFromDataBase = NewPlankopf
            LoadFromDataBase.TinLinePKNr = .Cells(row, 22).value
            IndexFactory.GetIndexes LoadFromDataBase
            Exit Function
        Else
            Dim frm          As New UserFormMessage
            frm.Typ TypWarning, "Es wurde kein Plankopf erstellt!"
            frm.Show 1
        End If
    End With

    writelog LogInfo, "Plankopf " & LoadFromDataBase.Plannummer & " geladen"

End Function

Sub PopulatePlankopf(ByRef Plankopf As IPlankopf)
    ' First set NodElement to the <tinPlan1> Node in the xml
    Dim str                  As String
    str = "PK" & Plankopf.TinLinePKNr

    CreateXmlAttribute "PA40", "Plan Überschrift", Plankopf.Planüberschrift, str, NodChild, oXml, NodElement
    CreateXmlAttribute "PA41", "Format", Plankopf.LayoutGrösse(True), str, NodChild, oXml, NodElement
    CreateXmlAttribute "PA42", "Massstab", Plankopf.LayoutMasstab, str, NodChild, oXml, NodElement
    CreateXmlAttribute "PA43", "Plannummer", Plankopf.LayoutName, str, NodChild, oXml, NodElement
    CreateXmlAttribute "PA44", "Planstand", Plankopf.LayoutPlanstand, str, NodChild, oXml, NodElement
    CreateXmlAttribute "PA30", "Gezeichnet", Plankopf.GezeichnetPerson, str, NodChild, oXml, NodElement
    CreateXmlAttribute "PA31", "Datum Gezeichnet", Plankopf.GezeichnetDatum, str, NodChild, oXml, NodElement
    CreateXmlAttribute "PA32", "Geprüft", Plankopf.GeprüftPerson, str, NodChild, oXml, NodElement
    CreateXmlAttribute "PA33", "Datum Geprüft", Plankopf.GeprüftDatum, str, NodChild, oXml, NodElement

    TinLineIndexes Plankopf, NodChild, oXml, NodElement

End Sub

Public Function AddToDatabase(ByVal Plankopf As IPlankopf) As Boolean
    AddToDatabase = False
    Dim ws                   As Worksheet: Set ws = Globals.shStoreData
    Dim row                  As Long: row = ws.range("A1").CurrentRegion.rows.Count + 1

    CopyToClipBoard Plankopf.LayoutName

    If Plankopf.Gewerk = "Elektro" And Globals.shProjekt.range("A1").value And Plankopf.PLANTYP = "PLA" Then NewTinLinePlankopf Plankopf Else writelog Logwarning, "Das Projekt wurde ohne Elektropläne erstellt." & vbNewLine & "Wenn die Pläne im TinLine erstellt werden, bitte den QS-Verantwortlichen kontaktieren" ' Elektroplan
    If Plankopf.Gewerk = "Elektro" And Globals.shProjekt.range("A2").value And Plankopf.PLANTYP = "PRI" Then NewTinLinePlankopf Plankopf Else writelog Logwarning, "Das Projekt wurde ohne Elektro Prinzipschemas erstellt." & vbNewLine & "Wenn die Pläne im TinLine erstellt werden, bitte den QS-Verantwortlichen kontaktieren" ' Elektro Prinzipschema
    If Plankopf.Gewerk = "Türfachplanung" And Globals.shProjekt.range("A4").value Then NewTinLinePlankopf Plankopf Else writelog Logwarning, "Das Projekt wurde ohne Türfachpläne erstellt." & vbNewLine & "Wenn die Pläne im TinLine erstellt werden, bitte den QS-Verantwortlichen kontaktieren" ' Türfachplanung
    If Plankopf.Gewerk = "Brandschutzplanung" And Globals.shProjekt.range("A5").value Then NewTinLinePlankopf Plankopf Else writelog Logwarning, "Das Projekt wurde ohne Brandschutzpläne erstellt." & vbNewLine & "Wenn die Pläne im TinLine erstellt werden, bitte den QS-Verantwortlichen kontaktieren" ' Brandschutzplanung
    With ws
        .Cells(row, 1).value = Plankopf.ID
        .Cells(row, 2).value = Plankopf.IDTinLine
        .Cells(row, 3).value = Plankopf.Gewerk
        .Cells(row, 4).value = Plankopf.UnterGewerk
        .Cells(row, 5).value = Plankopf.Planart
        .Cells(row, 6).value = Plankopf.PLANTYP
        .Cells(row, 7).value = Plankopf.Gebäude
        .Cells(row, 8).value = Plankopf.Gebäudeteil
        .Cells(row, 9).value = Plankopf.Geschoss
        .Cells(row, 10).value = Plankopf.CustomPlanüberschrift
        .Cells(row, 11).value = Plankopf.dwgFile
        .Cells(row, 13).value = Plankopf.Planüberschrift
        .Cells(row, 14).value = Plankopf.Plannummer
        .Cells(row, 15).value = Plankopf.LayoutGrösse
        .Cells(row, 16).value = Plankopf.LayoutMasstab
        .Cells(row, 17).value = Plankopf.LayoutPlanstand
        .Cells(row, 18).value = Plankopf.GezeichnetPerson
        .Cells(row, 19).value = Replace(Plankopf.GezeichnetDatum, ".", "/")
        .Cells(row, 20).value = Plankopf.GeprüftPerson
        .Cells(row, 21).value = Replace(Plankopf.GeprüftDatum, ".", "/")
        .Cells(row, 12).value = Plankopf.CurrentIndex.Index
        .Cells(row, 22).value = Plankopf.TinLinePKNr
        .Cells(row, 23).value = Plankopf.AnlageTyp
        .Cells(row, 24).value = Plankopf.AnlageNummer
    End With
    AddToDatabase = True
    writelog LogInfo, "Plankopf " & Plankopf.Plannummer & " in Datenbank gespeichert"

End Function

Private Function NewTinLinePlankopf(ByRef Plankopf As IPlankopf) As Boolean

    oXsl.load XMLVorlage

    ' Standard XML Elemente für TinLine erstellen / Einlesen
    If Len(dir(Plankopf.XMLFile)) = 0 Then
        oXml.LoadXML "<tinPlan1></tinPlan1>"
    Else
        oXml.load Plankopf.XMLFile
    End If
    writelog LogTrace, "XML geladen: " & Plankopf.XMLFile & vbNewLine & oXml.XML

    If CheckEmptyPlankopf(Plankopf) Then
        ' XML formatieren
        oXml.Save Plankopf.XMLFile
        oXml.transformNodeToObject oXsl, oXml
        oXml.Save Plankopf.XMLFile

        writelog LogInfo, "Plankopf " & Plankopf.Plannummer & " in TinLine geschrieben"
        NewTinLinePlankopf = True
    Else
        NewTinLinePlankopf = False
        writelog Logwarning, "Plankopf " & Plankopf.Plannummer & " nicht erstellt"
    End If

End Function

Private Function CheckEmptyPlankopf(ByRef Plankopf As IPlankopf) As Boolean

load:                                                                     ' load xml file
    On Error GoTo err
    Set NodElement = oXml.SelectSingleNode("tinPlan1")
    Dim oSeqNodes            As IXMLDOMNodeList
    Dim oSeqNode             As IXMLDOMNode
    Dim PKs                  As New Collection
    ' select all PK nodes
    Set oSeqNodes = oXml.SelectNodes("//tinPlan1/PK")

    If oSeqNodes.length = 0 Then
        GoTo err
    End If

    For Each oSeqNode In oSeqNodes
        PKs.Add CInt(oSeqNode.SelectSingleNode("Nr").Text)
    Next

    Dim arrPK()              As Variant
    arrPK = CollectionToArray(PKs)
    PKNr = WorksheetFunction.Max(arrPK)
    Plankopf.TinLinePKNr = PKNr

    'If PKNr = "" Then GoTo err

    ' there is a Plankopf in the xml file
    ' check if the Plankopf is empty
EmptyPK:
    Dim ChildNod             As IXMLDOMNodeList
    Dim Nod                  As IXMLDOMNode
    Set ChildNod = oXml.SelectNodes("tinPlan1/PK" & CStr(PKNr))

FreierPlankopf:
    For Each Nod In ChildNod
        If Nod.FirstChild.Text = "PA40" And Not Nod.LastChild.Text = vbNullString Then
            GoTo err
        End If
        NodElement.RemoveChild Nod
    Next Nod

TinLineID:
    ' get the TinLine ID of the current PK
    Dim TinLineID            As String

    For Each oSeqNode In oSeqNodes
        If Not oSeqNode.SelectSingleNode("Name").Text = Plankopf.LayoutName And oSeqNode.SelectSingleNode("Nr").Text = Plankopf.TinLinePKNr Then
            writelog Logwarning, "Das Layout wurde möglicherweise nicht richtig beschriftet " & Plankopf.LayoutName
            CopyToClipBoard Plankopf.LayoutName
            MsgBox "Das Layout ist möglicherweise falsch bezeichnet." & vbNewLine & "Bitte das Layout :" & vbNewLine & oSeqNode.SelectSingleNode("Name").Text & vbNewLine & " in " & vbNewLine & Plankopf.LayoutName & vbNewLine & " Umbenennen." & vbNewLine & vbNewLine & "Die korrekte Beschriftung ist in der Zwischenablage.", vbExclamation, "Layout Umbenennen"
            oSeqNode.SelectSingleNode("Name").Text = Plankopf.LayoutName
        End If
        If oSeqNode.SelectSingleNode("Nr").Text = PKNr Then
            TinLineID = CStr(oSeqNode.SelectSingleNode("ID").Text)
        End If
    Next
TinLineIDFound:
    Plankopf.IDTinLine = TinLineID
    writelog LogTrace, "TinLine ID in Plankopf eingesetzt " & Plankopf.IDTinLine
    PopulatePlankopf Plankopf
    CheckEmptyPlankopf = True
    Exit Function
err:

    ' there s no Plankopf in the specified file
    writelog LogTrace, "Kein leerer Plankopf in XML " & Plankopf.XMLFile
    Select Case MsgBox("Es besteht kein Leerer Plankopf in der Datei: " & vbNewLine & vbNewLine & Plankopf.XMLFile & vbNewLine & vbNewLine & "Datei im TinLine öffnen?", vbYesNo, "Kein Plankopf!")
    Case vbYes
        CreateObject("Shell.Application").Open (Plankopf.dwgFile)
        writelog LogTrace, "DWG Geöffnet im TinLine " & Plankopf.dwgFile
        Select Case MsgBox("Plankopf im TinLine erstellt?", vbYesNo)
        Case vbYes
            writelog LogTrace, "Plankopf erstellt "
            oXml.load Plankopf.XMLFile
            GoTo load
        Case Else
            writelog LogTrace, "Plankopf NICHT erstellt "
            CheckEmptyPlankopf = False
            Exit Function
        End Select
    Case Else
        writelog LogTrace, "DWG NICHT Geöffnet im TinLine " & Plankopf.dwgFile
        CheckEmptyPlankopf = False
        Exit Function
    End Select

End Function

Public Function ReplaceInDatabase(ByVal Plankopf As IPlankopf) As Boolean
    ReplaceInDatabase = False
    Dim ID                   As String: ID = Plankopf.ID
    Dim ws                   As Worksheet: Set ws = Globals.shStoreData
    Dim row                  As Long: row = ws.range("A:A").Find(ID).row
    With ws
        '.Cells(Row, 1).Value = Plankopf.ID
        '.Cells(Row, 2).Value = Plankopf.IDTinLine
        '.Cells(Row, 3).Value = Plankopf.Gewerk
        '.Cells(Row, 4).Value = Plankopf.UnterGewerk
        '.Cells(Row, 5).Value = Plankopf.Planart
        '.Cells(Row, 6).Value = Plankopf.Plantyp
        '.Cells(Row, 7).Value = Plankopf.Gebäude
        '.Cells(Row, 8).Value = Plankopf.GebäudeTeil
        '.Cells(Row, 9).Value = Plankopf.Geschoss
        .Cells(row, 10).value = Plankopf.CustomPlanüberschrift
        .Cells(row, 11).value = Plankopf.dwgFile
        .Cells(row, 13).value = Plankopf.Planüberschrift
        '.Cells(Row, 14).Value = Plankopf.Plannummer
        .Cells(row, 15).value = Plankopf.LayoutGrösse
        .Cells(row, 16).value = Plankopf.LayoutMasstab
        .Cells(row, 17).value = Plankopf.LayoutPlanstand
        .Cells(row, 18).value = Plankopf.GezeichnetPerson
        .Cells(row, 19).value = Replace(Plankopf.GezeichnetDatum, ".", "/")
        .Cells(row, 20).value = Plankopf.GeprüftPerson
        .Cells(row, 21).value = Replace(Plankopf.GeprüftDatum, ".", "/")
        .Cells(row, 23).value = Plankopf.AnlageTyp
        .Cells(row, 24).value = Plankopf.AnlageNummer
    End With
    ReplaceInDatabase = True
    writelog LogInfo, "Plankopf " & Plankopf.Plannummer & " in Datenbank aktualisiert"

    If Plankopf.Gewerk = "Elektro" And Plankopf.PLANTYP = "PLA" And Globals.shProjekt.range("A1").value Then ChangeTinLinePlankopf Plankopf ' Elektroplan
    If Plankopf.Gewerk = "Elektro" And Plankopf.PLANTYP = "PRI" And Globals.shProjekt.range("A2").value Then ChangeTinLinePlankopf Plankopf ' Elektro Prinzipschema
    If Plankopf.Gewerk = "Türfachplanung" And Globals.shProjekt.range("A4").value Then ChangeTinLinePlankopf Plankopf ' Türfachplanung
    If Plankopf.Gewerk = "Brandschutzplanung" And Globals.shProjekt.range("A5").value Then ChangeTinLinePlankopf Plankopf ' Brandschutzplanung

    writelog LogInfo, "Plankopf " & Plankopf.Plannummer & " im TinLine aktualisiert"

End Function

Private Function ChangeTinLinePlankopf(ByRef Plankopf As IPlankopf) As Boolean

    oXsl.load XMLVorlage

    ' Standard XML Elemente für TinLine erstellen / Einlesen
    If Len(dir(Plankopf.XMLFile)) = 0 Then
        oXml.LoadXML "<tinPlan1></tinPlan1>"
    Else
        oXml.load Plankopf.XMLFile
    End If
    writelog LogTrace, "XML geladen: " & Plankopf.XMLFile & vbNewLine & oXml.XML

    If CheckChangePlankopf(Plankopf) Then
        ' XML formatieren
        oXml.Save Plankopf.XMLFile
        oXml.transformNodeToObject oXsl, oXml
        oXml.Save Plankopf.XMLFile

        writelog LogInfo, "Plankopf " & Plankopf.Plannummer & " in TinLine geschrieben"
        ChangeTinLinePlankopf = True
    Else
        ChangeTinLinePlankopf = False
        writelog Logwarning, "Plankopf " & Plankopf.Plannummer & " nicht geändert"
    End If

End Function

Private Function CheckChangePlankopf(ByRef Plankopf As IPlankopf) As Boolean

load:                                                                     ' load xml file
    On Error GoTo err
    Set NodElement = oXml.SelectSingleNode("tinPlan1")
    Dim oSeqNodes            As IXMLDOMNodeList
    Dim oSeqNode             As IXMLDOMNode
    Dim PKs                  As New Collection
    ' select all PK nodes
    Set oSeqNodes = oXml.SelectNodes("//tinPlan1/PK")

    If oSeqNodes.length = 0 Then
        GoTo err
    End If

    For Each oSeqNode In oSeqNodes
        If oSeqNode.SelectSingleNode("ID").Text = Plankopf.IDTinLine Then
            PKNr = oSeqNode.SelectSingleNode("Nr").Text
            Plankopf.TinLinePKNr = PKNr
        End If
    Next

EmptyPK:
    Dim ChildNod             As IXMLDOMNodeList
    Dim Nod                  As IXMLDOMNode
    Set ChildNod = oXml.SelectNodes("tinPlan1/PK" & CStr(PKNr))

FreierPlankopf:
    For Each Nod In ChildNod
        NodElement.RemoveChild Nod
    Next Nod

TinLineID:
    ' get the TinLine ID of the current PK
    PopulatePlankopf Plankopf
    CheckChangePlankopf = True
    Exit Function
err:

    ' there s no Plankopf in the specified file
    writelog LogTrace, "Kein leerer Plankopf in XML " & Plankopf.XMLFile
    Select Case MsgBox("Es besteht kein Leerer Plankopf in der Datei: " & vbNewLine & vbNewLine & Plankopf.XMLFile & vbNewLine & vbNewLine & "Datei im TinLine öffnen?", vbYesNo, "Kein Plankopf!")
    Case vbYes
        CreateObject("Shell.Application").Open (Plankopf.dwgFile)
        writelog LogTrace, "DWG Geöffnet im TinLine " & Plankopf.dwgFile
        Select Case MsgBox("Plankopf im TinLine erstellt?", vbYesNo)
        Case vbYes
            writelog LogTrace, "Plankopf erstellt "
            oXml.load Plankopf.XMLFile
            GoTo load
        Case Else
            writelog LogTrace, "Plankopf NICHT erstellt "
            CheckEmptyPlankopf = False
            Exit Function
        End Select
    Case Else
        writelog LogTrace, "DWG NICHT Geöffnet im TinLine " & Plankopf.dwgFile
        CheckEmptyPlankopf = False
        Exit Function
    End Select

End Function

Public Function DeleteFromDatabase(ByVal row As Long) As Boolean
    DeleteFromDatabase = False
    Dim ID                   As String
    Dim Plannummer           As String: Plannummer = shStoreData.Cells(row, 14).value
    ID = shStoreData.Cells(row, 1).value
    shStoreData.Cells(row, 1).EntireRow.Delete
    IndexFactory.DeletePlan ID
    DeleteFromDatabase = True
    writelog LogInfo, "Plankopf " & Plannummer & " aus Datenbank gelöscht"

End Function

Public Function RewritePlankopf(ByVal Plankopf As IPlankopf) As Boolean
    oXsl.load XMLVorlage
    oXml.load Plankopf.XMLFile
    Set NodElement = oXml.SelectSingleNode("tinPlan1")

    Set NodChild = oXml.createElement("PK")
    NodElement.appendChild NodChild

    Set NodGrandChild = oXml.createElement("Nr")
    NodGrandChild.Text = Plankopf.TinLinePKNr
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Plankopf")
    NodGrandChild.Text = "LAY_Plankopf.dwg"
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Name")
    NodGrandChild.Text = Plankopf.LayoutName
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("Datum")
    NodGrandChild.Text = Format$(CDate(Plankopf.GezeichnetDatum), "DD/MM/YYYY hh:mm:ss")
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("UDatum")
    NodGrandChild.Text = Format$(Now, "DD/MM/YYYY hh:mm:ss")
    NodChild.appendChild NodGrandChild
    Set NodGrandChild = oXml.createElement("ID")
    NodGrandChild.Text = Plankopf.IDTinLine
    NodChild.appendChild NodGrandChild

    PopulatePlankopf Plankopf

    ' XML formatieren
    oXml.Save Plankopf.XMLFile
    oXml.transformNodeToObject oXsl, oXml
    oXml.Save Plankopf.XMLFile

End Function


