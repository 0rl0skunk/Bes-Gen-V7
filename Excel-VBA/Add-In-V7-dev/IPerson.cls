VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface für 'Person.cls'."

'@Folder("Person")
'@Interface
'@ModuleDescription "Interface für 'Person.cls'."

Option Explicit

Public Property Get Nachname() As String
End Property

Public Property Get Vorname() As String
End Property

Public Property Get Anrede() As String
End Property

Public Property Get Adresse() As IAdresse
End Property

Public Property Get Firma() As String
End Property

Public Property Get EMail() As String
End Property

Public Property Get ID() As String
End Property

