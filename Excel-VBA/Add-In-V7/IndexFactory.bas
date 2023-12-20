Attribute VB_Name = "IndexFactory"
Attribute VB_Description = "Erstellt ein Index-Objekt von welchem die daten einfach ausgelesen werden können."
Option Explicit
'@Folder("Index")
'@ModuleDescription "Erstellt ein Index-Objekt von welchem die daten einfach ausgelesen werden können."

Private oXml                 As New MSXML2.DOMDocument60
Private oXsl                 As New MSXML2.DOMDocument60

Private NodElement           As IXMLDOMElement
Private NodChild             As IXMLDOMElement
Private NodGrandChild        As IXMLDOMElement

Private pPlankopf As IPlankopf

Public Function Create( _
       ByVal IDPlan As String, _
       ByVal GezeichnetPerson As String, _
       ByVal GezeichnetDatum As String, _
       ByVal Klartext As String, _
       Optional ByVal ID As String = vbNullString, _
       Optional ByVal Letter As String, _
       Optional ByVal GeprüftPerson As String = vbNullString, _
       Optional ByVal GeprüftDatum As String = vbNullString, _
       Optional ByVal SkipValidation As Boolean _
       ) As IIndex

    Dim newIndex             As New Index
    newIndex.Filldata _
        ID:=ID, _
        IDPlan:=IDPlan, _
        Letter:=Letter, _
        GezeichnetPerson:=GezeichnetPerson, _
        GezeichnetDatum:=GezeichnetDatum, _
        GeprüftPerson:=GeprüftPerson, _
        GeprüftDatum:=GeprüftDatum, _
        Klartext:=Klartext, _
        SkipValidation:=SkipValidation

    Set Create = newIndex

End Function

Public Sub DeleteFromDatabase(ID As String)
    ' löscht den gewählten Index aus der Datenbank
    Globals.shIndex.range("H:H").Find(ID).EntireRow.Delete
    writelog LogInfo, "Index gelöscht"
End Sub

Public Sub AddToDatabase(Index As IIndex)
    ' erstellt einen neuen Index in der Datenbank
    Dim row                  As Long
    Dim Gezeichnet           As String
    Dim Geprüft              As String


    Gezeichnet = Index.Gezeichnet
    Geprüft = Index.Geprüft

    row = Globals.shIndex.range("A1").CurrentRegion.rows.Count + 1

    With Globals.shIndex
        .Cells(row, 1).value = Index.PlanID
        .Cells(row, 2).value = Index.Index
        .Cells(row, 3).value = Split(Gezeichnet, ";")(0)
        .Cells(row, 4).value = Split(Gezeichnet, ";")(1)
        .Cells(row, 5).value = Split(Geprüft, ";")(0)
        .Cells(row, 6).value = Split(Geprüft, ";")(1)
        .Cells(row, 7).value = Index.Klartext
        .Cells(row, 8).value = Index.IndexID
    End With

    writelog LogInfo, "Index für Plankopf erstellt"

End Sub

Public Sub DeletePlan(ByVal ID As String)
    ' Löscht alle Indexe von einem Plan
    Dim row                  As Long
    Dim coll                 As New Collection: Set coll = GetIndexes(ID:=ID)
    With Globals.shIndex
        For row = .range("A1").CurrentRegion.rows.Count To 2 Step -1
            If .Cells(row, 1).value = ID Then: .Cells(row, 1).EntireRow.Delete
        Next
    End With

    writelog LogInfo, coll.Count & " Indexe für Plankopf gelöscht"

End Sub

Public Function GetIndexes(Optional ByRef Plankopf As IPlankopf, Optional ByVal ID As String = vbNullString) As Collection
    ' gibt eine Collection von allen Indexen eines Plankopes zurück

    Dim row                  As Long
    Dim IndexID              As String
    Dim IDPlan               As String
    Dim GezeichnetPerson     As String
    Dim GezeichnetDatum      As String
    Dim Klartext             As String
    Dim Letter               As String
    Dim GeprüftPerson        As String
    Dim GeprüftDatum         As String
    Dim Index                As IIndex
    Dim coll                 As New Collection


    With Globals.shIndex
        For row = 2 To .range("A1").CurrentRegion.rows.Count
            IndexID = .Cells(row, 8).value
            IDPlan = .Cells(row, 1).value
            Letter = .Cells(row, 2).value
            GezeichnetPerson = .Cells(row, 3).value
            GezeichnetDatum = .Cells(row, 4).value
            GeprüftPerson = .Cells(row, 5).value
            GeprüftDatum = .Cells(row, 6).value
            Klartext = .Cells(row, 7).value

            If Not Plankopf Is Nothing Then If IDPlan = Plankopf.ID Then GoTo Matching
            If IDPlan = ID Then
                GoTo Matching
            Else
                GoTo Skip
            End If
Matching:
            Set Index = Create(ID:=IndexID, _
                               IDPlan:=IDPlan, _
                               GezeichnetPerson:=GezeichnetPerson, _
                               GezeichnetDatum:=GezeichnetDatum, _
                               Klartext:=Klartext, _
                               Letter:=Letter, _
                               GeprüftPerson:=GeprüftPerson, _
                               GeprüftDatum:=GeprüftDatum)
            coll.Add Index
            If Not Plankopf Is Nothing Then Plankopf.AddIndex Index
Skip:
        Next
    End With
    Set GetIndexes = coll
    On Error GoTo ErrMsg
    writelog LogInfo, coll.Count & " Indexe für Plankopf" & Plankopf.Plannummer
    Exit Function
ErrMsg:
    writelog LogInfo, "NO Indexe für Plankopf"
End Function

Public Function TinLineIndexes(ByRef Plankopf As IPlankopf, ByRef iNodChild As IXMLDOMElement, ByRef ioXml As MSXML2.DOMDocument60, ByRef iNodElement As IXMLDOMElement)
Set oXml = ioXml
Set NodElement = iNodElement
Set pPlankopf = Plankopf
On Error GoTo 0
DeleteIndexes
WriteIndexes

End Function

Private Sub DeleteIndexes()
Dim seqNode As IXMLDOMNode
For Each seqNode In NodElement.SelectNodes("IN" & pPlankopf.TinLinePKNr)
    NodElement.RemoveChild seqNode
Next
writelog LogInfo, "Alle Indexe für Plankopf " & pPlankopf.XMLFile
End Sub

Private Sub WriteIndexes()
Dim Index As IIndex
For Each Index In pPlankopf.Indexes
    CreateXmlIndexAttribute Index.Index, Index.Gezeichnet, Index.Klartext, "IN" & pPlankopf.TinLinePKNr, NodChild, oXml, NodElement
Next
End Sub
