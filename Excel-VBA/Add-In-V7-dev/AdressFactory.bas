Attribute VB_Name = "AdressFactory"
Attribute VB_Description = "Erstellt ein Adress-Objekt von welchem die daten einfach ausgelesen werden können."

'@Folder "Adresse"
'@ModuleDescription "Erstellt ein Adress-Objekt von welchem die daten einfach ausgelesen werden können."

Option Explicit

Public Function Create( _
       ByVal Strasse As String, _
       ByVal PLZ As String, _
       ByVal Ort As String _
       ) As IAdresse

    Dim NewAdresse           As New Adresse
    NewAdresse.Filldata _
        Strasse:=Strasse, _
        PLZ:=PLZ, _
        Ort:=Ort

    Set Create = NewAdresse

End Function


