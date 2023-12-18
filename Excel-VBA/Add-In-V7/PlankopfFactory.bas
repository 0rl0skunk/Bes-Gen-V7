Attribute VB_Name = "PlankopfFactory"
Attribute VB_Description = "Erstellt ein Plankopf-Objekt von welchem die daten einfach ausgelesen werden k�nnen."
'@IgnoreModule VariableNotUsed
Option Explicit
'@Folder "Plankopf"
'@ModuleDescription "Erstellt ein Plankopf-Objekt von welchem die daten einfach ausgelesen werden k�nnen."

Public Function Create( _
       ByVal Projekt As IProjekt, _
       ByVal GezeichnetPerson As String, _
       ByVal GezeichnetDatum As String, _
       ByVal Gepr�ftPerson As String, _
       ByVal Gepr�ftDatum As String, _
       ByVal Geb�ude As String, _
       ByVal Geb�udeteil As String, _
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
       Optional ByVal Plan�berschrift As String = "NEW", _
       Optional ByVal ID As String = "NEW", _
       Optional ByVal Custom�berschrift As Boolean = False _
       ) As IPlankopf

    Dim NewPlankopf          As New Plankopf
    If NewPlankopf.Filldata( _
       Projekt:=Projekt, _
       GezeichnetPerson:=GezeichnetPerson, _
       GezeichnetDatum:=GezeichnetDatum, _
       Gepr�ftPerson:=Gepr�ftPerson, _
       Gepr�ftDatum:=Gepr�ftDatum, _
       Geb�ude:=Geb�ude, _
       Geb�udeteil:=Geb�udeteil, _
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
       Plan�berschrift:=Plan�berschrift, _
       ID:=ID, _
       Custom�berschrift:=Custom�berschrift _
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
           Geb�ude:=.Cells(row, 7).Value, _
           Geb�udeteil:=.Cells(row, 8).Value, _
           Geschoss:=.Cells(row, 9).Value, _
           Plan�berschrift:=.Cells(row, 13).Value, _
           Format:=.Cells(row, 15).Value, _
           Masstab:=.Cells(row, 16).Value, _
           Stand:=.Cells(row, 17).Value, _
           GezeichnetPerson:=.Cells(row, 18).Value, _
           GezeichnetDatum:=.Cells(row, 19).Value, _
           Gepr�ftPerson:=.Cells(row, 20).Value, _
           Gepr�ftDatum:=.Cells(row, 21).Value, _
           SkipValidation:=False, _
           Custom�berschrift:=.Cells(row, 10).Value _
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
        .Cells(row, 7).Value = Plankopf.Geb�ude
        .Cells(row, 8).Value = Plankopf.Geb�udeteil
        .Cells(row, 9).Value = Plankopf.Geschoss
        .Cells(row, 10).Value = Plankopf.CustomPlan�berschrift
        .Cells(row, 11).Value = Plankopf.DWGFile
        .Cells(row, 13).Value = Plankopf.Plan�berschrift
        .Cells(row, 14).Value = Plankopf.Plannummer
        .Cells(row, 15).Value = Plankopf.LayoutGr�sse
        .Cells(row, 16).Value = Plankopf.LayoutMasstab
        .Cells(row, 17).Value = Plankopf.LayoutPlanstand
        .Cells(row, 18).Value = Plankopf.GezeichnetPerson
        .Cells(row, 19).Value = Plankopf.GezeichnetDatum
        .Cells(row, 20).Value = Plankopf.Gepr�ftPerson
        .Cells(row, 21).Value = Plankopf.Gepr�ftDatum
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
        '.Cells(Row, 7).Value = Plankopf.Geb�ude
        '.Cells(Row, 8).Value = Plankopf.Geb�udeTeil
        '.Cells(Row, 9).Value = Plankopf.Geschoss
        .Cells(row, 10).Value = Plankopf.CustomPlan�berschrift
        .Cells(row, 11).Value = Plankopf.DWGFile
        .Cells(row, 13).Value = Plankopf.Plan�berschrift
        '.Cells(Row, 14).Value = Plankopf.Plannummer
        .Cells(row, 15).Value = Plankopf.LayoutGr�sse
        .Cells(row, 16).Value = Plankopf.LayoutMasstab
        .Cells(row, 17).Value = Plankopf.LayoutPlanstand
        .Cells(row, 18).Value = Plankopf.GezeichnetPerson
        .Cells(row, 19).Value = Plankopf.GezeichnetDatum
        .Cells(row, 20).Value = Plankopf.Gepr�ftPerson
        .Cells(row, 21).Value = Plankopf.Gepr�ftDatum
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
    writelog LogInfo, "Plankopf " & Plannummer & " aus Datenbank gel�scht"

End Function


