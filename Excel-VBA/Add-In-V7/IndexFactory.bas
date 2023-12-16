Attribute VB_Name = "IndexFactory"
Option Explicit
'@Folder("Index")
'@ModuleDescription "Erstellt ein Index-Objekt von welchem die daten einfach ausgelesen werden k�nnen."

Public Function Create( _
       ByVal IDPlan As String, _
       ByVal GezeichnetPerson As String, _
       ByVal GezeichnetDatum As String, _
       ByVal Klartext As String, _
       Optional ByVal ID As String = vbNullString, _
       Optional ByVal Letter As String, _
       Optional ByVal Gepr�ftPerson As String = vbNullString, _
       Optional ByVal Gepr�ftDatum As String = vbNullString, _
       Optional ByVal SkipValidation As Boolean _
       ) As IIndex

    Dim newIndex             As New Index
    newIndex.Filldata _
        ID:=ID, _
        IDPlan:=IDPlan, _
        Letter:=Letter, _
        GezeichnetPerson:=GezeichnetPerson, _
        GezeichnetDatum:=GezeichnetDatum, _
        Gepr�ftPerson:=Gepr�ftPerson, _
        Gepr�ftDatum:=Gepr�ftDatum, _
        Klartext:=Klartext, _
        SkipValidation:=SkipValidation

    Set Create = newIndex

End Function

Public Sub DeleteFromDatabase(ID As String)
    ' l�scht den gew�hlten Index aus der Datenbank
    Globals.shIndex.range("H:H").Find(ID).EntireRow.Delete
    writelog "Info", "Index gel�scht"
End Sub

Public Sub AddToDatabase(Index As IIndex)
    ' erstellt einen neuen Index in der Datenbank
    Dim _
    row                      As Long, _
    Gezeichnet               As String, _
    Gepr�ft                  As String

    Gezeichnet = Index.Gezeichnet
    Gepr�ft = Index.Gepr�ft

    row = Globals.shIndex.range("A1").CurrentRegion.rows.Count + 1

    With Globals.shIndex
        .Cells(row, 1).Value = Index.PlanID
        .Cells(row, 2).Value = Index.Index
        .Cells(row, 3).Value = Split(Gezeichnet, ";")(0)
        .Cells(row, 4).Value = Split(Gezeichnet, ";")(1)
        .Cells(row, 5).Value = Split(Gepr�ft, ";")(0)
        .Cells(row, 6).Value = Split(Gepr�ft, ";")(1)
        .Cells(row, 7).Value = Index.Klartext
        .Cells(row, 8).Value = Index.IndexID
    End With

    writelog "Info", "Index f�r Plankopf erstellt"

End Sub

Public Function DeletePlan(ByVal ID As String)
    ' L�scht alle Indexe von einem Plan
    Dim row                  As Long
    Dim coll                 As New Collection: Set coll = GetIndexes(ID:=ID)
    With Globals.shIndex
        For row = .range("A1").CurrentRegion.rows.Count To 2 Step -1
            If .Cells(row, 1).Value = ID Then: .Cells(row, 1).EntireRow.Delete
        Next
    End With

    writelog "Info", coll.Count & " Indexe f�r Plankopf gel�scht"

End Function

Public Function GetIndexes(Optional ByRef Plankopf As IPlankopf, Optional ByVal ID As String = vbNullString) As Collection
    ' gibt eine Collection von allen Indexen eines Plankopes zur�ck

    Dim _
    row                      As Long, _
    IndexID                  As String, _
    IDPlan                   As String, _
    GezeichnetPerson         As String, _
    GezeichnetDatum          As String, _
    Klartext                 As String, _
    Letter                   As String, _
    Gepr�ftPerson            As String, _
    Gepr�ftDatum             As String, _
    Index                    As IIndex, _
    coll                     As New Collection

    With Globals.shIndex
        For row = 2 To .range("A1").CurrentRegion.rows.Count
            IndexID = .Cells(row, 8).Value
            IDPlan = .Cells(row, 1).Value
            Letter = .Cells(row, 2).Value
            GezeichnetPerson = .Cells(row, 3).Value
            GezeichnetDatum = .Cells(row, 4).Value
            Gepr�ftPerson = .Cells(row, 5).Value
            Gepr�ftDatum = .Cells(row, 6).Value
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
                               Gepr�ftPerson:=Gepr�ftPerson, _
                               Gepr�ftDatum:=Gepr�ftDatum)
            coll.Add Index
            If Not Plankopf Is Nothing Then Plankopf.AddIndex Index
Skip:
        Next
    End With
    Set GetIndexes = coll
    On Error GoTo ErrMsg
    writelog "Info", coll.Count & " Indexe f�r Plankopf" & Plankopf.Plannummer
    Exit Function
ErrMsg:
    writelog "Info", "NO Indexe f�r Plankopf"
End Function


