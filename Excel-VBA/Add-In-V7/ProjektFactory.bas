Attribute VB_Name = "ProjektFactory"
Option Explicit
'@Folder "Projekt"
'@ModuleDescription "Erstellt ein Projekt-Objekt von welchem die daten einfach ausgelesen werden können."

Public Function Create( _
       ByVal Projektnummer As String, _
       ByVal Projektadresse As IAdresse, _
       ByVal ProjektBezeichnung As String, _
       ByVal Projektphase As String, _
       ByVal ProjektOrdnerSharePoint As String _
       ) As IProjekt

    Dim NewProjekt           As Projekt
    Set NewProjekt = New Projekt
    NewProjekt.FillData _
        Projektnummer, _
        Projektadresse, _
        ProjektBezeichnung, _
        Projektphase, _
        ProjektOrdnerSharePoint

    Set Create = NewProjekt

End Function


