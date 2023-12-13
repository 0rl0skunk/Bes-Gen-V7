VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Adresse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Erstellt ein Adress-Objekt von welchem die daten einfach ausgelesen werden können."
'@IgnoreModule PublicImplementationShouldBePrivate
Option Explicit
'@ModuleDescription "Erstellt ein Adress-Objekt von welchem die daten einfach ausgelesen werden können."
Implements IAdresse

'@Folder "Adresse"
Private Type TAdresse
    Strasse As String
    PLZ As String
    Ort As String
End Type

Private This                 As TAdresse

Public Sub FillData( _
       ByVal Strasse As String, _
       ByVal PLZ As String, _
       ByVal Ort As String _
       )

    This.Strasse = Strasse
    This.PLZ = PLZ
    This.Ort = Ort

End Sub

Private Property Get IAdresse_Strasse() As String
    IAdresse_Strasse = This.Strasse
End Property

Private Property Get IAdresse_PLZ() As String
    IAdresse_PLZ = This.PLZ
End Property

Private Property Get IAdresse_Ort() As String
    IAdresse_Ort = This.Ort
End Property

Public Property Get IAdresse_Komplett() As String
    IAdresse_Komplett = This.Strasse & vbNewLine & This.PLZ & " " & This.Ort
End Property


