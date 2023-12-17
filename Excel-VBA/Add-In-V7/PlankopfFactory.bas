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
        Dim Row              As Long
        Set Create = NewPlankopf
        IndexFactory.GetIndexes Create
        Exit Function
    Else
        Dim frm              As New UserFormMessage
        frm.Typ typError, "Es wurde kein Plankopf erstellt!"
        frm.Show 1
    End If



End Function

Public Function LoadFromDataBase(ByVal Row As Long) As IPlankopf

    Dim NewPlankopf          As Plankopf:    Set NewPlankopf = New Plankopf
    Dim WS                   As Worksheet:    Set WS = Globals.shStoreData
    With WS
        If NewPlankopf.Filldata( _
           Projekt:=Projekt, _
           ID:=.Cells(Row, 1).Value, _
           TinLineID:=.Cells(Row, 2).Value, _
           Gewerk:=.Cells(Row, 3).Value, _
           UnterGewerk:=.Cells(Row, 4).Value, _
           Planart:=.Cells(Row, 5).Value, _
           Plantyp:=.Cells(Row, 6).Value, _
           Geb�ude:=.Cells(Row, 7).Value, _
           Geb�udeteil:=.Cells(Row, 8).Value, _
           Geschoss:=.Cells(Row, 9).Value, _
           Plan�berschrift:=.Cells(Row, 13).Value, _
           Format:=.Cells(Row, 15).Value, _
           Masstab:=.Cells(Row, 16).Value, _
           Stand:=.Cells(Row, 17).Value, _
           GezeichnetPerson:=.Cells(Row, 18).Value, _
           GezeichnetDatum:=.Cells(Row, 19).Value, _
           Gepr�ftPerson:=.Cells(Row, 20).Value, _
           Gepr�ftDatum:=.Cells(Row, 21).Value, _
           SkipValidation:=False, _
           Custom�berschrift:=.Cells(Row, 10).Value _
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
    Dim WS                   As Worksheet: Set WS = Globals.shStoreData
    Dim Row                  As Long: Row = WS.range("A1").CurrentRegion.rows.Count + 1
    With WS
        .Cells(Row, 1).Value = Plankopf.ID
        .Cells(Row, 2).Value = Plankopf.IDTinLine
        .Cells(Row, 3).Value = Plankopf.Gewerk
        .Cells(Row, 4).Value = Plankopf.UnterGewerk
        .Cells(Row, 5).Value = Plankopf.Planart
        .Cells(Row, 6).Value = Plankopf.Plantyp
        .Cells(Row, 7).Value = Plankopf.Geb�ude
        .Cells(Row, 8).Value = Plankopf.Geb�udeteil
        .Cells(Row, 9).Value = Plankopf.Geschoss
        .Cells(Row, 10).Value = Plankopf.CustomPlan�berschrift
        .Cells(Row, 11).Value = Plankopf.DWGFile
        .Cells(Row, 13).Value = Plankopf.Plan�berschrift
        .Cells(Row, 14).Value = Plankopf.Plannummer
        .Cells(Row, 15).Value = Plankopf.LayoutGr�sse
        .Cells(Row, 16).Value = Plankopf.LayoutMasstab
        .Cells(Row, 17).Value = Plankopf.LayoutPlanstand
        .Cells(Row, 18).Value = Plankopf.GezeichnetPerson
        .Cells(Row, 19).Value = Plankopf.GezeichnetDatum
        .Cells(Row, 20).Value = Plankopf.Gepr�ftPerson
        .Cells(Row, 21).Value = Plankopf.Gepr�ftDatum
        .Cells(Row, 12).Value = Plankopf.CurrentIndex.Index
    End With
    AddToDatabase = True
    writelog LogInfo, "Plankopf " & Plankopf.Plannummer & " in Datenbank gespeichert"

End Function

Public Function ReplaceInDatabase(Plankopf As IPlankopf) As Boolean
    ReplaceInDatabase = False
    Dim ID                   As String: ID = Plankopf.ID
    Dim WS                   As Worksheet: Set WS = Globals.shStoreData
    Dim Row                  As Long: Row = WS.range("A:A").Find(ID).Row
    With WS
        '.Cells(Row, 1).Value = Plankopf.ID
        '.Cells(Row, 2).Value = Plankopf.IDTinLine
        '.Cells(Row, 3).Value = Plankopf.Gewerk
        '.Cells(Row, 4).Value = Plankopf.UnterGewerk
        '.Cells(Row, 5).Value = Plankopf.Planart
        '.Cells(Row, 6).Value = Plankopf.Plantyp
        '.Cells(Row, 7).Value = Plankopf.Geb�ude
        '.Cells(Row, 8).Value = Plankopf.Geb�udeTeil
        '.Cells(Row, 9).Value = Plankopf.Geschoss
        .Cells(Row, 10).Value = Plankopf.CustomPlan�berschrift
        .Cells(Row, 11).Value = Plankopf.DWGFile
        .Cells(Row, 13).Value = Plankopf.Plan�berschrift
        '.Cells(Row, 14).Value = Plankopf.Plannummer
        .Cells(Row, 15).Value = Plankopf.LayoutGr�sse
        .Cells(Row, 16).Value = Plankopf.LayoutMasstab
        .Cells(Row, 17).Value = Plankopf.LayoutPlanstand
        .Cells(Row, 18).Value = Plankopf.GezeichnetPerson
        .Cells(Row, 19).Value = Plankopf.GezeichnetDatum
        .Cells(Row, 20).Value = Plankopf.Gepr�ftPerson
        .Cells(Row, 21).Value = Plankopf.Gepr�ftDatum
    End With
    ReplaceInDatabase = True
    writelog LogInfo, "Plankopf " & Plankopf.Plannummer & " in Datenbank aktualisiert"

End Function

Public Function DeleteFromDatabase(Row As Long) As Boolean
    DeleteFromDatabase = False
    Dim ID                   As String
    Dim Plannummer           As String: Plannummer = shStoreData.Cells(Row, 14).Value
    ID = shStoreData.Cells(Row, 1).Value
    shStoreData.Cells(Row, 1).EntireRow.Delete
    IndexFactory.DeletePlan ID
    DeleteFromDatabase = True
    writelog LogInfo, "Plankopf " & Plannummer & " aus Datenbank gel�scht"

End Function


