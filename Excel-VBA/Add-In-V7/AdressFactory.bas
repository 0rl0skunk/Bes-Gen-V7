Attribute VB_Name = "AdressFactory"
Option Explicit
'@Folder "Adresse"
' a Factory 'creates the new class and binds it to the interface so only the wanted methods are exposed to the user

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


