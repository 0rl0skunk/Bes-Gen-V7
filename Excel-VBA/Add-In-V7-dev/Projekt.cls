VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Projekt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Projekt Objekt f�r einfacheres Management vom Projekt."

'@Folder "Projekt"
'@ModuleDescription "Projekt Objekt f�r einfacheres Management vom Projekt."

Option Explicit

Implements IProjekt

Private Type TProjekt
    Projektnummer As String
    ProjektAdresse As IAdresse
    ProjektBezeichnung As String
    Projektphase As String
    ProjektphaseNummer As Long
    ProjektOrdnerSharePoint As String
    ProjektOrdnerCAD As String
    ProjektXML As String
End Type

Private This                 As TProjekt

Public Sub Filldata( _
       ByVal Projektnummer As String, _
       ByVal ProjektAdresse As IAdresse, _
       ByVal ProjektBezeichnung As String, _
       ByVal Projektphase As String, _
       ByVal ProjektOrdnerSharePoint As String _
       )

    This.Projektnummer = Projektnummer
    Set This.ProjektAdresse = ProjektAdresse
    This.ProjektBezeichnung = ProjektBezeichnung
    This.Projektphase = Projektphase
    This.ProjektphaseNummer = Globals.shPZM.range("PRO_Projektphase").Find(This.Projektphase).Offset(0, 1).value
    This.ProjektOrdnerSharePoint = ProjektOrdnerSharePoint
    GetCADPath

End Sub

Private Property Get IProjekt_Projektnummer() As String
    IProjekt_Projektnummer = This.Projektnummer
End Property

Private Property Get IProjekt_ProjektXML() As String
    IProjekt_ProjektXML = This.ProjektXML
End Property

Private Property Get IProjekt_Projektadresse() As IAdresse
    Set IProjekt_Projektadresse = This.ProjektAdresse
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

Private Sub GetCADPath()

    Dim TinLinePath          As String
    TinLinePath = "H:\TinLine\00_Projekte"

    Dim Jahr                 As String
    Jahr = "20" & Left$(This.Projektnummer, 2)

    If Globals.shPData.range("ADM_ProjektPfadCAD").value = vbNullString Then
        This.ProjektOrdnerCAD = TinLinePath & "\" & Jahr & "\" & This.Projektnummer & "_" & This.ProjektBezeichnung
    Else
        This.ProjektOrdnerCAD = Globals.shPData.range("ADM_ProjektPfadCAD").value
    End If

    This.ProjektXML = This.ProjektOrdnerCAD & "\99 TinConfiguration\TinPlanProject.xml"

    writelog LogInfo, "CAD-Pfad f�r Projekt " & This.Projektnummer & " > " & This.ProjektOrdnerCAD

End Sub


