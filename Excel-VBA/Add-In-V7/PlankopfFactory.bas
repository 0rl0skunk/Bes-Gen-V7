Attribute VB_Name = "PlankopfFactory"
Attribute VB_Description = "Erstellt ein Plankopf-Objekt von welchem die daten einfach ausgelesen werden können."
'@IgnoreModule VariableNotUsed
Option Explicit
'@Folder "Plankopf"
'@ModuleDescription "Erstellt ein Plankopf-Objekt von welchem die daten einfach ausgelesen werden können."

Private oXml                 As New MSXML2.DOMDocument60
Private oXsl                 As New MSXML2.DOMDocument60

Private NodElement           As IXMLDOMElement
Private NodChild             As IXMLDOMElement
Private NodGrandChild        As IXMLDOMElement

Private PKNr As Long

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
       Optional ByVal Plantyp As String, _
       Optional ByVal TinLineID As String, _
       Optional ByVal SkipValidation As Boolean = False, _
       Optional ByVal Planüberschrift As String = "NEW", _
       Optional ByVal ID As String = "NEW", _
       Optional ByVal CustomÜberschrift As Boolean = False _
       ) As IPlankopf

    Dim NewPlankopf          As New Plankopf
    If NewPlankopf.Filldata( _
       Projekt:=Projekt, _
       GezeichnetPerson:=GezeichnetPerson, _
       GezeichnetDatum:=GezeichnetDatum, _
       GeprüftPerson:=GeprüftPerson, _
       GeprüftDatum:=GeprüftDatum, _
       Gebäude:=Gebäude, _
       Gebäudeteil:=Gebäudeteil, _
       Geschoss:=Geschoss, _
       Gewerk:=Gewerk, _
       UnterGewerk:=UnterGewerk, _
       Format:=Format, _
       Masstab:=Masstab, _
       Stand:=Stand, _
       Planart:=Planart, _
       Plantyp:=Plantyp, _
       TinLineID:=TinLineID, _
       SkipValidation:=SkipValidation, _
       Planüberschrift:=Planüberschrift, _
       ID:=ID, _
       CustomÜberschrift:=CustomÜberschrift _
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
           Plantyp:=.Cells(row, 6).value, _
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
           CustomÜberschrift:=.Cells(row, 10).value _
                               ) Then
            Set LoadFromDataBase = NewPlankopf
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

Sub PopulatePlankopf(ByVal Plankopf As IPlankopf)
' First set NodElement to the <tinPlan1> Node in the xml
    Dim str                  As String
    str = "PK" & PKNr

   CreateXmlAttribute "PA40", "Plan Überschrift", Plankopf.Planüberschrift, str, NodChild, oXml, NodElement
   CreateXmlAttribute "PA41", "Format", Plankopf.LayoutGrösse(True), str, NodChild, oXml, NodElement
   CreateXmlAttribute "PA42", "Massstab", Plankopf.LayoutMasstab, str, NodChild, oXml, NodElement
   CreateXmlAttribute "PA43", "Plannummer", Plankopf.LayoutName, str, NodChild, oXml, NodElement
   CreateXmlAttribute "PA44", "Planstand", Plankopf.LayoutPlanstand, str, NodChild, oXml, NodElement
   CreateXmlAttribute "PA30", "Gezeichnet", Plankopf.GezeichnetPerson, str, NodChild, oXml, NodElement
   CreateXmlAttribute "PA31", "Datum Gezeichnet", Plankopf.GezeichnetDatum, str, NodChild, oXml, NodElement
   CreateXmlAttribute "PA32", "Geprüft", Plankopf.GeprüftPerson, str, NodChild, oXml, NodElement
   CreateXmlAttribute "PA33", "Datum Geprüft", Plankopf.GeprüftDatum, str, NodChild, oXml, NodElement

End Sub

Public Function AddToDatabase(Plankopf As IPlankopf) As Boolean
    AddToDatabase = False
    Dim ws                   As Worksheet: Set ws = Globals.shStoreData
    Dim row                  As Long: row = ws.range("A1").CurrentRegion.rows.Count + 1
    
    CopyToClipBoard Plankopf.LayoutName
    
    If Plankopf.Gewerk = "Elektro" Then NewTinLinePlankopf (Plankopf)
    With ws
        .Cells(row, 1).value = Plankopf.ID
        .Cells(row, 2).value = Plankopf.IDTinLine
        .Cells(row, 3).value = Plankopf.Gewerk
        .Cells(row, 4).value = Plankopf.UnterGewerk
        .Cells(row, 5).value = Plankopf.Planart
        .Cells(row, 6).value = Plankopf.Plantyp
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
        .Cells(row, 19).value = Plankopf.GezeichnetDatum
        .Cells(row, 20).value = Plankopf.GeprüftPerson
        .Cells(row, 21).value = Plankopf.GeprüftDatum
        .Cells(row, 12).value = Plankopf.CurrentIndex.Index
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
    oXml.save Plankopf.XMLFile
    oXml.transformNodeToObject oXsl, oXml
    oXml.save Plankopf.XMLFile

    writelog LogInfo, "Plankopf " & Plankopf.Plannummer & " in TinLine geschrieben"
    NewTinLinePlankopf = True
    Else
    NewTinLinePlankopf = False
    writelog Logwarning, "Plankopf " & Plankopf.Plannummer & " nicht erstellt"
    End If

End Function

Private Function CheckEmptyPlankopf(ByRef Plankopf As IPlankopf) As Boolean

load:                                            ' load xml file
    On Error GoTo err
    Set NodElement = oXml.SelectSingleNode("tinPlan1")
    Dim oSeqNodes            As IXMLDOMNodeList
    Dim oSeqNode As IXMLDOMNode
    Dim PKs                  As New Collection
    ' select all PK nodes
    Set oSeqNodes = oXml.SelectNodes("//tinPlan1/PK")

    If oSeqNodes.length = 0 Then
        GoTo err
    End If

    For Each oSeqNode In oSeqNodes
        PKs.Add CInt(oSeqNode.SelectSingleNode("Nr").text)
    Next

    Dim arrPK()              As Variant
    arrPK = CollectionToArray(PKs)
    PKNr = WorksheetFunction.Max(arrPK)

    'If PKNr = "" Then GoTo err

    ' there is a Plankopf in the xml file
    ' check if the Plankopf is empty
EmptyPK:
    Dim ChildNod As IXMLDOMNodeList
    Dim Nod As IXMLDOMNode
    Set ChildNod = oXml.SelectNodes("tinPlan1/PK" & CStr(PKNr))

FreierPlankopf:
    For Each Nod In ChildNod
        If Nod.FirstChild.text = "PA40" And Not Nod.LastChild.text = "" Then
            GoTo err
        End If
        NodElement.RemoveChild Nod
    Next Nod

TinLineID:
    ' get the TinLine ID of the current PK
    Dim TinLineID            As String

    For Each oSeqNode In oSeqNodes
        If Not oSeqNode.SelectSingleNode("Name").text = Plankopf.LayoutName Then
            writelog Logwarning, "Das Layout wurde möglicherweise nicht richtig beschriftet " & Plankopf.LayoutName
            CopyToClipBoard Plankopf.LayoutName
            MsgBox "Das Layout ist möglicherweise falsch bezeichnet." & vbNewLine & "Bitte das Layout :" & vbNewLine & oSeqNode.SelectSingleNode("Name").text & vbNewLine & " in " & vbNewLine & Plankopf.LayoutName & vbNewLine & " Umbenennen." & vbNewLine & vbNewLine & "Die korrekte Beschriftung ist in der Zwischenablage.", vbExclamation, "Layout Umbenennen"
            oSeqNode.SelectSingleNode("Name").text = Plankopf.LayoutName
        End If
        If oSeqNode.SelectSingleNode("Nr").text = PKNr Then
            TinLineID = CStr(oSeqNode.SelectSingleNode("ID").text)
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
    Dim answer
    answer = MsgBox("Es besteht kein Leerer Plankopf in der Datei: " & vbNewLine & vbNewLine & Plankopf.XMLFile & vbNewLine & vbNewLine & "Datei im TinLine öffnen?", vbYesNo, "Kein Plankopf!")
    writelog LogTrace, "Kein leerer Plankopf in XML " & Plankopf.XMLFile
    If answer = vbYes Then
        CreateObject("Shell.Application").Open (Plankopf.dwgFile)
        answer = MsgBox("Plankopf im TinLine erstellt?", vbYesNo)
        writelog LogTrace, "DWG Geöffnet im TinLine " & Plankopf.dwgFile
        If answer = vbYes Then
        writelog LogTrace, "Plankopf erstellt "
        oXml.load Plankopf.XMLFile
            GoTo load
        Else
            writelog LogTrace, "Plankopf NICHT erstellt "
            CheckEmptyPlankopf = False
            Exit Function
        End If

    Else
        writelog LogTrace, "DWG NICHT Geöffnet im TinLine " & Plankopf.dwgFile
        CheckEmptyPlankopf = False
        Exit Function
    End If

End Function

Public Function ReplaceInDatabase(Plankopf As IPlankopf) As Boolean
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
        .Cells(row, 19).value = Plankopf.GezeichnetDatum
        .Cells(row, 20).value = Plankopf.GeprüftPerson
        .Cells(row, 21).value = Plankopf.GeprüftDatum
    End With
    ReplaceInDatabase = True
    writelog LogInfo, "Plankopf " & Plankopf.Plannummer & " in Datenbank aktualisiert"
    
    If Plankopf.Gewerk = "Elektro" Then ChangeTinLinePlankopf (Plankopf)
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
    oXml.save Plankopf.XMLFile
    oXml.transformNodeToObject oXsl, oXml
    oXml.save Plankopf.XMLFile

    writelog LogInfo, "Plankopf " & Plankopf.Plannummer & " in TinLine geschrieben"
    ChangeTinLinePlankopf = True
    Else
    ChangeTinLinePlankopf = False
    writelog Logwarning, "Plankopf " & Plankopf.Plannummer & " nicht geändert"
    End If

End Function

Private Function CheckChangePlankopf(ByRef Plankopf As IPlankopf) As Boolean

load:                                            ' load xml file
    On Error GoTo err
    Set NodElement = oXml.SelectSingleNode("tinPlan1")
    Dim oSeqNodes            As IXMLDOMNodeList
    Dim oSeqNode As IXMLDOMNode
    Dim PKs                  As New Collection
    ' select all PK nodes
    Set oSeqNodes = oXml.SelectNodes("//tinPlan1/PK")

    If oSeqNodes.length = 0 Then
        GoTo err
    End If

    For Each oSeqNode In oSeqNodes
        If oSeqNode.SelectSingleNode("ID").text = Plankopf.IDTinLine Then
        PKNr = oSeqNode.SelectSingleNode("Nr").text
        End If
    Next

EmptyPK:
    Dim ChildNod As IXMLDOMNodeList
    Dim Nod As IXMLDOMNode
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
    Dim answer
    answer = MsgBox("Es besteht kein Leerer Plankopf in der Datei: " & vbNewLine & vbNewLine & Plankopf.XMLFile & vbNewLine & vbNewLine & "Datei im TinLine öffnen?", vbYesNo, "Kein Plankopf!")
    writelog LogTrace, "Kein leerer Plankopf in XML " & Plankopf.XMLFile
    If answer = vbYes Then
        CreateObject("Shell.Application").Open (Plankopf.dwgFile)
        answer = MsgBox("Plankopf im TinLine erstellt?", vbYesNo)
        writelog LogTrace, "DWG Geöffnet im TinLine " & Plankopf.dwgFile
        If answer = vbYes Then
        writelog LogTrace, "Plankopf erstellt "
        oXml.load Plankopf.XMLFile
            GoTo load
        Else
            writelog LogTrace, "Plankopf NICHT erstellt "
            CheckChangePlankopf = False
            Exit Function
        End If

    Else
        writelog LogTrace, "DWG NICHT Geöffnet im TinLine " & Plankopf.dwgFile
        CheckChangePlankopf = False
        Exit Function
    End If

End Function

Public Function DeleteFromDatabase(row As Long) As Boolean
    DeleteFromDatabase = False
    Dim ID                   As String
    Dim Plannummer           As String: Plannummer = shStoreData.Cells(row, 14).value
    ID = shStoreData.Cells(row, 1).value
    shStoreData.Cells(row, 1).EntireRow.Delete
    IndexFactory.DeletePlan ID
    DeleteFromDatabase = True
    writelog LogInfo, "Plankopf " & Plannummer & " aus Datenbank gelöscht"

End Function


