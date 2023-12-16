Attribute VB_Name = "IndexFactory"
Option Explicit
'@Folder("Index")
'@ModuleDescription "Erstellt ein Index-Objekt von welchem die daten einfach ausgelesen werden können."

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
    writelog "Info", "Index gelöscht"
End Sub

Public Sub AddToDatabase(Index As IIndex)
    ' erstellt einen neuen Index in der Datenbank
    Dim _
    row                      As Long, _
    Gezeichnet               As String, _
    Geprüft                  As String

    Gezeichnet = Index.Gezeichnet
    Geprüft = Index.Geprüft

    row = Globals.shIndex.range("A1").CurrentRegion.rows.Count + 1

    With Globals.shIndex
        .Cells(row, 1).Value = Index.PlanID
        .Cells(row, 2).Value = Index.Index
        .Cells(row, 3).Value = Split(Gezeichnet, ";")(0)
        .Cells(row, 4).Value = Split(Gezeichnet, ";")(1)
        .Cells(row, 5).Value = Split(Geprüft, ";")(0)
        .Cells(row, 6).Value = Split(Geprüft, ";")(1)
        .Cells(row, 7).Value = Index.Klartext
        .Cells(row, 8).Value = Index.IndexID
    End With

    writelog "Info", "Index für Plankopf erstellt"

End Sub

Public Function DeletePlan(ByVal ID As String)
    ' Löscht alle Indexe von einem Plan
    Dim row                  As Long
    Dim coll                 As New Collection: Set coll = GetIndexes(ID:=ID)
    With Globals.shIndex
        For row = .range("A1").CurrentRegion.rows.Count To 2 Step -1
            If .Cells(row, 1).Value = ID Then: .Cells(row, 1).EntireRow.Delete
        Next
    End With

    writelog "Info", coll.Count & " Indexe für Plankopf gelöscht"

End Function

Public Function GetIndexes(Optional ByRef Plankopf As IPlankopf, Optional ByVal ID As String = vbNullString) As Collection
    ' gibt eine Collection von allen Indexen eines Plankopes zurück

    Dim _
    row                      As Long, _
    IndexID                  As String, _
    IDPlan                   As String, _
    GezeichnetPerson         As String, _
    GezeichnetDatum          As String, _
    Klartext                 As String, _
    Letter                   As String, _
    GeprüftPerson            As String, _
    GeprüftDatum             As String, _
    Index                    As IIndex, _
    coll                     As New Collection

    With Globals.shIndex
        For row = 2 To .range("A1").CurrentRegion.rows.Count
            IndexID = .Cells(row, 8).Value
            IDPlan = .Cells(row, 1).Value
            Letter = .Cells(row, 2).Value
            GezeichnetPerson = .Cells(row, 3).Value
            GezeichnetDatum = .Cells(row, 4).Value
            GeprüftPerson = .Cells(row, 5).Value
            GeprüftDatum = .Cells(row, 6).Value
            Klartext = .Cells(row, 7).Value

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
    writelog "Info", coll.Count & " Indexe für Plankopf" & Plankopf.Plannummer
    Exit Function
ErrMsg:
    writelog "Info", "NO Indexe für Plankopf"
End Function


