VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAdresse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface für 'Adresse.cls'."

'@Folder "Adresse"
'@Interface
'@ModuleDescription "Interface für 'Adresse.cls'."

Option Explicit

Public Property Get Strasse() As String
End Property

Public Property Get PLZ() As String
End Property

Public Property Get Ort() As String
End Property

Public Property Get Komplett() As String
End Property

