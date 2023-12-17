Attribute VB_Name = "IndexFactory"
Attribute VB_Description = "Erstellt ein Index-Objekt von welchem die daten einfach ausgelesen werden k�nnen."
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
    writelog LogInfo, "Index gel�scht"
End Sub

Public Sub AddToDatabase(Index As IIndex)
    ' erstellt einen neuen Index in der Datenbank
    Dim Row                  As Long
    Dim Gezeichnet           As String
    Dim Gepr�ft              As String


    Gezeichnet = Index.Gezeichnet
    Gepr�ft = Index.Gepr�ft

    Row = Globals.shIndex.range("A1").CurrentRegion.rows.Count + 1

    With Globals.shIndex
        .Cells(Row, 1).Value = Index.PlanID
        .Cells(Row, 2).Value = Index.Index
        .Cells(Row, 3).Value = Split(Gezeichnet, ";")(0)
        .Cells(Row, 4).Value = Split(Gezeichnet, ";")(1)
        .Cells(Row, 5).Value = Split(Gepr�ft, ";")(0)
        .Cells(Row, 6).Value = Split(Gepr�ft, ";")(1)
        .Cells(Row, 7).Value = Index.Klartext
        .Cells(Row, 8).Value = Index.IndexID
    End With

    writelog LogInfo, "Index f�r Plankopf erstellt"

End Sub

Public Sub DeletePlan(ByVal ID As String)
    ' L�scht alle Indexe von einem Plan
    Dim Row                  As Long
    Dim coll                 As New Collection: Set coll = GetIndexes(ID:=ID)
    With Globals.shIndex
        For Row = .range("A1").CurrentRegion.rows.Count To 2 Step -1
            If .Cells(Row, 1).Value = ID Then: .Cells(Row, 1).EntireRow.Delete
        Next
    End With

    writelog LogInfo, coll.Count & " Indexe f�r Plankopf gel�scht"

End Sub

Public Function GetIndexes(Optional ByRef Plankopf As IPlankopf, Optional ByVal ID As String = vbNullString) As Collection
    ' gibt eine Collection von allen Indexen eines Plankopes zur�ck

    Dim Row                  As Long
    Dim IndexID              As String
    Dim IDPlan               As String
    Dim GezeichnetPerson     As String
    Dim GezeichnetDatum      As String
    Dim Klartext             As String
    Dim Letter               As String
    Dim Gepr�ftPerson        As String
    Dim Gepr�ftDatum         As String
    Dim Index                As IIndex
    Dim coll                 As New Collection


    With Globals.shIndex
        For Row = 2 To .range("A1").CurrentRegion.rows.Count
            IndexID = .Cells(Row, 8).Value
            IDPlan = .Cells(Row, 1).Value
            Letter = .Cells(Row, 2).Value
            GezeichnetPerson = .Cells(Row, 3).Value
            GezeichnetDatum = .Cells(Row, 4).Value
            Gepr�ftPerson = .Cells(Row, 5).Value
            Gepr�ftDatum = .Cells(Row, 6).Value
            Klartext = .Cells(Row, 7).Value

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
    writelog LogInfo, coll.Count & " Indexe f�r Plankopf" & Plankopf.Plannummer
    Exit Function
ErrMsg:
    writelog LogInfo, "NO Indexe f�r Plankopf"
End Function


