Attribute VB_Name = "PlankopfFactory"
Option Explicit
'@Folder "Plankopf"
' a Factory 'creates the new class and binds it to the interface so only the wanted methods are exposed to the user

Public Function Create( _
       ByVal Projekt As IProjekt, _
       ByVal GezeichnetPerson As String, _
       ByVal GezeichnetDatum As String, _
       ByVal GeprüftPerson As String, _
       ByVal GeprüftDatum As String, _
       ByVal Gebäude As String, _
       ByVal GebäudeTeil As String, _
       ByVal Geschoss As String, _
       ByVal Gewerk As String, _
       ByVal UnterGewerk As String, _
       ByVal Format As String, _
       ByVal Masstab As String, _
       ByVal Stand As String, _
       ByVal Klartext As String, _
       ByVal Planart As String, _
       Optional ByVal Plantyp As String, _
       Optional ByVal TinLineID As String, _
       Optional ByVal SkipValidation As Boolean = False, _
       Optional ByVal Planüberschrift As String = "NEW", _
       Optional ByVal ID As String = "NEW" _
       ) As IPlankopf

    Dim NewPlankopf          As Plankopf: Set NewPlankopf = New Plankopf
    If NewPlankopf.FillData( _
       Projekt:=Projekt, _
       GezeichnetPerson:=GezeichnetPerson, _
       GezeichnetDatum:=GezeichnetDatum, _
       GeprüftPerson:=GeprüftPerson, _
       GeprüftDatum:=GeprüftDatum, _
       Gebäude:=Gebäude, _
       GebäudeTeil:=GebäudeTeil, _
       Geschoss:=Geschoss, _
       Gewerk:=Gewerk, _
       UnterGewerk:=UnterGewerk, _
       Format:=Format, _
       Masstab:=Masstab, _
       Stand:=Stand, _
       Klartext:=Klartext, _
       Planart:=Planart, _
       Plantyp:=Plantyp, _
       TinLineID:=TinLineID, _
       SkipValidation:=SkipValidation, _
       Planüberschrift:=Planüberschrift, _
       ID:=ID _
            ) Then
        Dim row As Long
        Set Create = NewPlankopf
        IndexFactory.GetIndexes Create
        Exit Function
    Else
        Dim frm              As New UserFormMessage
        frm.typeWarning "Es wurde kein Plankopf erstellt!"
        frm.Show 1
    End If



End Function

Public Function LoadFromDatabase(ByVal row As Long) As IPlankopf

    Globals.SetWBs

    Dim NewPlankopf          As Plankopf:    Set NewPlankopf = New Plankopf
    Dim ws                   As Worksheet:    Set ws = Globals.shStoreData
    With ws
        If NewPlankopf.FillData( _
           Projekt:=Projekt, _
           ID:=.Cells(row, 1).Value, _
           TinLineID:=.Cells(row, 2).Value, _
           Gewerk:=.Cells(row, 3).Value, _
           UnterGewerk:=.Cells(row, 4).Value, _
           Planart:=.Cells(row, 5).Value, _
           Plantyp:=.Cells(row, 6).Value, _
           Gebäude:=.Cells(row, 7).Value, _
           GebäudeTeil:=.Cells(row, 8).Value, _
           Geschoss:=.Cells(row, 9).Value, _
           Klartext:=.Cells(row, 10).Value, _
           Planüberschrift:=.Cells(row, 13).Value, _
           Format:=.Cells(row, 15).Value, _
           Masstab:=.Cells(row, 16).Value, _
           Stand:=.Cells(row, 17).Value, _
           GezeichnetPerson:=.Cells(row, 18).Value, _
           GezeichnetDatum:=.Cells(row, 19).Value, _
           GeprüftPerson:=.Cells(row, 20).Value, _
           GeprüftDatum:=.Cells(row, 21).Value, _
           SkipValidation:=False _
                            ) Then
            Set LoadFromDatabase = NewPlankopf
            IndexFactory.GetIndexes LoadFromDatabase
            Exit Function
        Else
            Dim frm          As New UserFormMessage
            frm.typeWarning "Es wurde kein Plankopf erstellt!"
            frm.Show 1
        End If
    End With


End Function

Public Function AddToDatabase(Plankopf As IPlankopf) As Boolean

    Dim ws                   As Worksheet: Set ws = Globals.shStoreData
    Dim row                  As Long: row = ws.range("A1").CurrentRegion.rows.Count + 1
    With ws
        .Cells(row, 1).Value = Plankopf.ID
        .Cells(row, 2).Value = Plankopf.IDTinLine
        .Cells(row, 3).Value = Plankopf.Gewerk
        .Cells(row, 4).Value = Plankopf.UnterGewerk
        .Cells(row, 5).Value = Plankopf.Planart
        .Cells(row, 6).Value = Plankopf.Plantyp
        .Cells(row, 7).Value = Plankopf.Gebäude
        .Cells(row, 8).Value = Plankopf.GebäudeTeil
        .Cells(row, 9).Value = Plankopf.Geschoss
        .Cells(row, 10).Value = Plankopf.Klartext
        .Cells(row, 13).Value = Plankopf.Planüberschrift
        .Cells(row, 14).Value = Plankopf.PlanNummer
        .Cells(row, 15).Value = Plankopf.LayoutGrösse
        .Cells(row, 16).Value = Plankopf.LayoutMasstab
        .Cells(row, 17).Value = Plankopf.LayoutPlanstand
        .Cells(row, 18).Value = Plankopf.GezeichnetPerson
        .Cells(row, 19).Value = Plankopf.GezeichnetDatum
        .Cells(row, 20).Value = Plankopf.GeprüftPerson
        .Cells(row, 21).Value = Plankopf.GeprüftDatum
        '.Cells(row, 1).Value = Plankopf.currentIndex.index
    End With

End Function

Public Function ReplaceInDatabase(Plankopf As IPlankopf) As Boolean

    Dim ID As String: ID = Plankopf.ID
    Dim ws                   As Worksheet: Set ws = Globals.shStoreData
    Dim row                  As Long: row = ws.range("A:A").Find(ID).row
    With ws
        .Cells(row, 1).Value = Plankopf.ID
        .Cells(row, 2).Value = Plankopf.IDTinLine
        .Cells(row, 3).Value = Plankopf.Gewerk
        .Cells(row, 4).Value = Plankopf.UnterGewerk
        .Cells(row, 5).Value = Plankopf.Planart
        .Cells(row, 6).Value = Plankopf.Plantyp
        .Cells(row, 7).Value = Plankopf.Gebäude
        .Cells(row, 8).Value = Plankopf.GebäudeTeil
        .Cells(row, 9).Value = Plankopf.Geschoss
        .Cells(row, 10).Value = Plankopf.Klartext
        .Cells(row, 13).Value = Plankopf.Planüberschrift
        .Cells(row, 14).Value = Plankopf.PlanNummer
        .Cells(row, 15).Value = Plankopf.LayoutGrösse
        .Cells(row, 16).Value = Plankopf.LayoutMasstab
        .Cells(row, 17).Value = Plankopf.LayoutPlanstand
        .Cells(row, 18).Value = Plankopf.GezeichnetPerson
        .Cells(row, 19).Value = Plankopf.GezeichnetDatum
        .Cells(row, 20).Value = Plankopf.GeprüftPerson
        .Cells(row, 21).Value = Plankopf.GeprüftDatum
        '.Cells(row, 1).Value = Plankopf.currentIndex.index
    End With

End Function

