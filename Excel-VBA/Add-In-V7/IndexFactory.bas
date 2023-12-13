Attribute VB_Name = "IndexFactory"
Option Explicit
'@Folder("Index")
' a Factory 'creates the new class and binds it to the interface so only the wanted methods are exposed to the user

Public Function Create( _
       ByVal IDPlan As String, _
       ByVal GezeichnetPerson As String, _
       ByVal GezeichnetDatum As String, _
       ByVal Klartext As String, _
       Optional ByVal ID As String = vbNullString, _
       Optional ByVal Letter As String, _
       Optional ByVal GeprüftPerson As String = vbNullString, _
       Optional ByVal GeprüftDatum As String = vbNullString _
       ) As IIndex

    Dim newIndex             As Index
    Set newIndex = New Index
    newIndex.FillData _
        ID:=ID, _
        IDPlan:=IDPlan, _
        Letter:=Letter, _
        GezeichnetPerson:=GezeichnetPerson, _
        GezeichnetDatum:=GezeichnetDatum, _
        GeprüftPerson:=GeprüftPerson, _
        GeprüftDatum:=GeprüftDatum, _
        Klartext:=Klartext

    Set Create = newIndex

End Function

Public Sub DeleteFromDatabase(ID As String)
    
    Globals.shIndex.range("H:H").Find(ID).EntireRow.Delete

End Sub

Public Sub AddToDatabase(Index As IIndex)

    Dim _
    row As Long, _
    Gezeichnet As String, _
    Geprüft As String
    
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

End Sub

Public Function GetIndexes(Plankopf As IPlankopf)

    Globals.SetWBs

    Dim _
    row As Long, _
    IndexID As String, _
    IDPlan As String, _
    GezeichnetPerson As String, _
    GezeichnetDatum As String, _
    Klartext As String, _
    Letter As String, _
    GeprüftPerson As String, _
    GeprüftDatum As String

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

            If IDPlan = Plankopf.ID Then
                Plankopf.AddIndex Create(ID:=IndexID, _
                                          IDPlan:=IDPlan, _
                                          GezeichnetPerson:=GezeichnetPerson, _
                                          GezeichnetDatum:=GezeichnetDatum, _
                                          Klartext:=Klartext, _
                                          Letter:=Letter, _
                                          GeprüftPerson:=GeprüftPerson, _
                                          GeprüftDatum:=GeprüftDatum)
            End If
        Next
    End With

End Function


