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
       Optional ByVal Gepr�ftDatum As String = vbNullString _
       ) As IIndex

    Dim newIndex             As Index
    Set newIndex = New Index
    newIndex.FillData _
        ID:=ID, _
        IDPlan:=IDPlan, _
        Letter:=Letter, _
        GezeichnetPerson:=GezeichnetPerson, _
        GezeichnetDatum:=GezeichnetDatum, _
        Gepr�ftPerson:=Gepr�ftPerson, _
        Gepr�ftDatum:=Gepr�ftDatum, _
        Klartext:=Klartext

    Set Create = newIndex

End Function

Public Sub DeleteFromDatabase(ID As String)
    ' l�scht den gew�hlten Index aus der Datenbank
    Globals.shIndex.range("H:H").Find(ID).EntireRow.Delete
    log.write "Info", "Index gel�scht" 
End Sub

Public Sub AddToDatabase(Index As IIndex)
    ' erstellt einen neuen Index in der Datenbank
    Dim _
    row As Long, _
    Gezeichnet As String, _
    Gepr�ft As String
    
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
    dim Plannummer as string: Plannummer = plankopffactory.loadfromdatabase(globals.shstoredata.range("A:A").find(index.planID).row).plannummer
    log.write "Info", "Index f�r Plankopf " & Plannummer & " erstellt" 

End Sub

Public Function DeletePlan(ByVal ID As String)
    ' L�scht alle Indexe von einem Plan
    Dim row As Long
    dim coll as new collection: set coll=    getindexes Id:=id
    dim Plannummer as string: Plannummer = plankopffactory.loadfromdatabase(globals.shstoredata.range("A:A").find(ID).row).plannummer
    With Globals.shIndex
        For row = .range("A1").CurrentRegion.rows.Count To 2 Step -1
        If .Cells(row, 1).Value = ID Then: .Cells(row, 1).EntireRow.Delete
        Next
    End With

    log.write "Info", coll.count & " Indexe f�r Plankopf " & Plannummer & " gel�scht" 

End Function

Public Function GetIndexes(byref optional Plankopf As IPlankopf, byval optional ID as string = vbnullstring) As Collection
    ' bibt eine Collection von allen Indexen eines Plankopes zur�ck
    Globals.SetWBs

    Dim _
    row As Long, _
    IndexID As String, _
    IDPlan As String, _
    GezeichnetPerson As String, _
    GezeichnetDatum As String, _
    Klartext As String, _
    Letter As String, _
    Gepr�ftPerson As String, _
    Gepr�ftDatum As String, _
    Index As IIndex, _
    coll As New Collection

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

            If IDPlan = Plankopf.ID or IDPlan = ID Then
                Set Index = Create(ID:=IndexID, _
                                          IDPlan:=IDPlan, _
                                          GezeichnetPerson:=GezeichnetPerson, _
                                          GezeichnetDatum:=GezeichnetDatum, _
                                          Klartext:=Klartext, _
                                          Letter:=Letter, _
                                          Gepr�ftPerson:=Gepr�ftPerson, _
                                          Gepr�ftDatum:=Gepr�ftDatum)
                coll.Add Index
                if not plankopf is nothing then                Plankopf.AddIndex Index
            End If
        Next
    End With
    Set GetIndexes = coll

    log.write "Info", coll.count & " Indexe f�r Plankopf" & Plankopf.Plannummer

End Function


