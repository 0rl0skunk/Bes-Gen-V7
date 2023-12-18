Attribute VB_Name = "PlankopfFactory"
Attribute VB_Description = "Erstellt ein Plankopf-Objekt von welchem die daten einfach ausgelesen werden können."
'@IgnoreModule VariableNotUsed
Option Explicit
'@Folder "Plankopf"
'@ModuleDescription "Erstellt ein Plankopf-Objekt von welchem die daten einfach ausgelesen werden können."

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
           ID:=.Cells(row, 1).Value, _
           TinLineID:=.Cells(row, 2).Value, _
           Gewerk:=.Cells(row, 3).Value, _
           UnterGewerk:=.Cells(row, 4).Value, _
           Planart:=.Cells(row, 5).Value, _
           Plantyp:=.Cells(row, 6).Value, _
           Gebäude:=.Cells(row, 7).Value, _
           Gebäudeteil:=.Cells(row, 8).Value, _
           Geschoss:=.Cells(row, 9).Value, _
           Planüberschrift:=.Cells(row, 13).Value, _
           Format:=.Cells(row, 15).Value, _
           Masstab:=.Cells(row, 16).Value, _
           Stand:=.Cells(row, 17).Value, _
           GezeichnetPerson:=.Cells(row, 18).Value, _
           GezeichnetDatum:=.Cells(row, 19).Value, _
           GeprüftPerson:=.Cells(row, 20).Value, _
           GeprüftDatum:=.Cells(row, 21).Value, _
           SkipValidation:=False, _
           CustomÜberschrift:=.Cells(row, 10).Value _
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

Public Function AddToDatabase(Plankopf As IPlankopf) As Boolean
    AddToDatabase = False
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
        .Cells(row, 8).Value = Plankopf.Gebäudeteil
        .Cells(row, 9).Value = Plankopf.Geschoss
        .Cells(row, 10).Value = Plankopf.CustomPlanüberschrift
        .Cells(row, 11).Value = Plankopf.DWGFile
        .Cells(row, 13).Value = Plankopf.Planüberschrift
        .Cells(row, 14).Value = Plankopf.Plannummer
        .Cells(row, 15).Value = Plankopf.LayoutGrösse
        .Cells(row, 16).Value = Plankopf.LayoutMasstab
        .Cells(row, 17).Value = Plankopf.LayoutPlanstand
        .Cells(row, 18).Value = Plankopf.GezeichnetPerson
        .Cells(row, 19).Value = Plankopf.GezeichnetDatum
        .Cells(row, 20).Value = Plankopf.GeprüftPerson
        .Cells(row, 21).Value = Plankopf.GeprüftDatum
        .Cells(row, 12).Value = Plankopf.CurrentIndex.Index
    End With
    AddToDatabase = True
    writelog LogInfo, "Plankopf " & Plankopf.Plannummer & " in Datenbank gespeichert"

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
        .Cells(row, 10).Value = Plankopf.CustomPlanüberschrift
        .Cells(row, 11).Value = Plankopf.DWGFile
        .Cells(row, 13).Value = Plankopf.Planüberschrift
        '.Cells(Row, 14).Value = Plankopf.Plannummer
        .Cells(row, 15).Value = Plankopf.LayoutGrösse
        .Cells(row, 16).Value = Plankopf.LayoutMasstab
        .Cells(row, 17).Value = Plankopf.LayoutPlanstand
        .Cells(row, 18).Value = Plankopf.GezeichnetPerson
        .Cells(row, 19).Value = Plankopf.GezeichnetDatum
        .Cells(row, 20).Value = Plankopf.GeprüftPerson
        .Cells(row, 21).Value = Plankopf.GeprüftDatum
    End With
    ReplaceInDatabase = True
    writelog LogInfo, "Plankopf " & Plankopf.Plannummer & " in Datenbank aktualisiert"

End Function

Public Function DeleteFromDatabase(row As Long) As Boolean
    DeleteFromDatabase = False
    Dim ID                   As String
    Dim Plannummer           As String: Plannummer = shStoreData.Cells(row, 14).Value
    ID = shStoreData.Cells(row, 1).Value
    shStoreData.Cells(row, 1).EntireRow.Delete
    IndexFactory.DeletePlan ID
    DeleteFromDatabase = True
    writelog LogInfo, "Plankopf " & Plannummer & " aus Datenbank gelöscht"

End Function


