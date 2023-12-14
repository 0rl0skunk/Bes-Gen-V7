Attribute VB_Name = "AdressFactory"
Option Explicit
'@Folder "Adresse"
'@ModuleDescription "Erstellt ein Adress-Objekt von welchem die daten einfach ausgelesen werden können."

Public Function Create( _
       ByVal Strasse As String, _
       ByVal PLZ As String, _
       ByVal Ort As String _
       ) As IAdresse

    Dim NewAdresse           As Adresse: Set NewAdresse = New Adresse
    NewAdresse.FillData _
        Strasse, _
        PLZ, _
        Ort

    Set Create = NewAdresse

End Function


