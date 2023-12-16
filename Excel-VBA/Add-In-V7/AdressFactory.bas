Attribute VB_Name = "AdressFactory"
Option Explicit
'@Folder "Adresse"
'@ModuleDescription "Erstellt ein Adress-Objekt von welchem die daten einfach ausgelesen werden können."

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


