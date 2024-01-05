Attribute VB_Name = "PersonFactory"
Attribute VB_Description = "Erstellt ein Index-Objekt von welchem die daten einfach ausgelesen werden können."

'@Folder("Person")
'@ModuleDescription "Erstellt ein Index-Objekt von welchem die daten einfach ausgelesen werden können."
'@Version "Release V1.0.0"

Option Explicit

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

Public Sub AddToDatabase(ByVal Person As IPerson)
    ' erstellt eine neue Person in der Datenbank
    Dim row                  As Long
    Dim ws                   As Worksheet


    Set ws = Globals.shAdress

    row = ws.range("A" & ws.rows.Count).End(xlUp).row + 1
    With ws
        .Cells(row, 1).value = Person.Nachname
        .Cells(row, 2).value = Person.Vorname
        .Cells(row, 3).value = Person.Firma
        .Cells(row, 4).value = Person.Adresse.Strasse
        .Cells(row, 5).value = Person.Adresse.PLZ
        .Cells(row, 6).value = Person.Adresse.Ort
        .Cells(row, 7).value = Person.EMail
        .Cells(row, 8).value = Person.Anrede
        .Cells(row, 9).value = Person.ID
    End With

    writelog LogInfo, "Person erfasst"

End Sub

Public Function LoadFromDataBase(row As Long) As IPerson
    ' Lädt die Daten aus der Datenbank
    Dim ws                   As Worksheet
    Dim NewPerson            As New IPerson


    Set ws = Globals.shAdress

    With ws
        Set NewPerson = Create( _
                        Nachname:=.Cells(row, 1).value, _
                        Vorname:=.Cells(row, 2).value, _
                        Firma:=.Cells(row, 3).value, _
                        Adresse:=AdressFactory.Create( _
                                  Strasse:=.Cells(row, 4).value, _
                        PLZ:=.Cells(row, 5).value, _
                        Ort:=.Cells(row, 6).value), _
                        EMail:=.Cells(row, 7).value, _
                        Anrede:=.Cells(row, 8).value _
                                 )
    End With

    Set LoadFromDataBase = NewPerson

    writelog LogInfo, "Person geladen"

End Function


