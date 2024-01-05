VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IProjekt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface für 'Projekt.cls'."
'@Folder "Projekt"
'@Interface
'@ModuleDescription "Interface für 'Projekt.cls'."


Option Explicit

                                
Public Property Get Projektnummer() As String
End Property

                                
Public Property Get ProjektXML() As String
End Property

                                
Public Property Get Projektadresse() As IAdresse
End Property

                                
Public Property Get ProjektBezeichnung() As String
End Property

                                
Public Property Get Projektphase() As String
End Property

                                
Public Property Get ProjektphaseNummer() As String
End Property

                                
Public Property Get ProjektOrdnerSharePoint() As String
End Property

                                
Public Property Get ProjektOrdnerCAD() As String
End Property

                                
