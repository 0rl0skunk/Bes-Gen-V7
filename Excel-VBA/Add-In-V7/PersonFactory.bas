Attribute VB_Name = "PersonFactory"
Attribute VB_Description = "Erstellt ein Index-Objekt von welchem die daten einfach ausgelesen werden können."
Option Explicit
'@Folder("Person")
'@ModuleDescription "Erstellt ein Index-Objekt von welchem die daten einfach ausgelesen werden können."

Public Function Create( _
       ByVal Nachname As String, _
       ByVal Vorname As String, _
       ByVal Firma As String, _
       ByVal Adresse As IAdresse, _
       ByVal EMail As String, _
       Optional ByVal Anrede As String, _
       Optional ByVal ID As String = vbNullString _
       ) As IPerson

    Dim NewPerson            As New Person
    NewPerson.Filldata _
        Vorname:=Vorname, _
        Nachname:=Nachname, _
        Adresse:=Adresse, _
        Anrede:=Anrede, _
        Firma:=Firma, _
        EMail:=EMail, _
        ID:=ID

    Set Create = NewPerson

End Function

Public Sub AddToDatabase(Person As IPerson)
    ' erstellt eine neue Person in der Datenbank
    Dim Row                  As Long
    Dim WS                   As Worksheet


    Set WS = Globals.shAdress

    Row = WS.range("A" & WS.rows.Count).End(xlUp).Row + 1
    With WS
        .Cells(Row, 1).Value = Person.Nachname
        .Cells(Row, 2).Value = Person.Vorname
        .Cells(Row, 3).Value = Person.Firma
        .Cells(Row, 4).Value = Person.Adresse.Strasse
        .Cells(Row, 5).Value = Person.Adresse.PLZ
        .Cells(Row, 6).Value = Person.Adresse.Ort
        .Cells(Row, 7).Value = Person.EMail
        .Cells(Row, 8).Value = Person.Anrede
        .Cells(Row, 9).Value = Person.ID
    End With

    writelog LogInfo, "Person erfasst"

End Sub

Public Function LoadFromDataBase(Row As Long) As IPerson
    ' Lädt die Daten aus der Datenbank
    Dim WS                   As Worksheet
    Dim NewPerson            As New IPerson


    Set WS = Globals.shAdress

    With WS
        Set NewPerson = Create( _
                        Nachname:=.Cells(Row, 1).Value, _
                        Vorname:=.Cells(Row, 2).Value, _
                        Firma:=.Cells(Row, 3).Value, _
                        Adresse:=AdressFactory.Create( _
                                  Strasse:=.Cells(Row, 4).Value, _
                        PLZ:=.Cells(Row, 5).Value, _
                        Ort:=.Cells(Row, 6).Value), _
                        EMail:=.Cells(Row, 7).Value, _
                        Anrede:=.Cells(Row, 8).Value _
                                 )
    End With

    Set LoadFromDataBase = NewPerson

    writelog LogInfo, "Person geladen"

End Function


