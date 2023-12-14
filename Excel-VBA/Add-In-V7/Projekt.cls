VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Projekt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder "Projekt"
'@ModuleDescription "Projekt Objekt für einfacheres Management vom Projekt."


Implements IProjekt

Private Type TProjekt
    Projektnummer As String
    Projektadresse As IAdresse
    ProjektBezeichnung As String
    Projektphase As String
    ProjektphaseNummer As Long
    ProjektOrdnerSharePoint As String
    ProjektOrdnerCAD As String
End Type

Private This                 As TProjekt

Public Sub FillData( _
       ByVal Projektnummer As String, _
       ByVal Projektadresse As IAdresse, _
       ByVal ProjektBezeichnung As String, _
       ByVal Projektphase As String, _
       ByVal ProjektOrdnerSharePoint As String _
       )

    This.Projektnummer = Projektnummer
    Set This.Projektadresse = Projektadresse
    This.ProjektBezeichnung = ProjektBezeichnung
    This.Projektphase = Projektphase
    This.ProjektphaseNummer = Globals.shPData.range("PRO_Projektphase").Find(This.Projektphase).Offset(0, 1).Value
    This.ProjektOrdnerSharePoint = ProjektOrdnerSharePoint
    GetCADPath

End Sub

Private Property Get IProjekt_Projektnummer() As String
    IProjekt_Projektnummer = This.Projektnummer
End Property

Private Property Get IProjekt_Projektadresse() As IAdresse
    Set IProjekt_Projektadresse = This.Projektadresse
End Property

Private Property Get IProjekt_ProjektBezeichnung() As String
    IProjekt_ProjektBezeichnung = This.ProjektBezeichnung
End Property

Private Property Get IProjekt_ProjektPhase() As String
    IProjekt_ProjektPhase = This.Projektphase
End Property

Private Property Get IProjekt_ProjektPhaseNummer() As String
    IProjekt_ProjektPhaseNummer = This.ProjektphaseNummer
End Property

Private Property Get IProjekt_ProjektOrdnerSharePoint() As String
    IProjekt_ProjektOrdnerSharePoint = This.ProjektOrdnerSharePoint
End Property

Private Property Get IProjekt_ProjektOrdnerCAD() As String
    IProjekt_ProjektOrdnerCAD = This.ProjektOrdnerCAD
End Property

Private Function GetCADPath()

    Dim TinLinePath          As String
    TinLinePath = "H:\TinLine\00_Projekte"

    Dim Jahr                 As String
    Jahr = "20" & Left$(This.Projektnummer, 2)

    This.ProjektOrdnerCAD = TinLinePath & "\" & Jahr & "\" & This.Projektnummer & "_" & This.ProjektBezeichnung

    writelog "Info", "CAD-Pfad für Projekt " & This.Projektnummer & " > " & This.ProjektOrdnerCAD

End Function


