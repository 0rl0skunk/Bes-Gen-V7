VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IPerson

'@Folder("Person")

Private Type TPerson
    Nachname As String
    Vorname As String
    Firma As String
    Anrede As String
    Adresse As IAdresse
    EMail As String
    ID As String
End Type

                                
Private This                 As TPerson

                                
Public Sub Filldata( _
       ByVal Nachname As String, _
       ByVal Vorname As String, _
       ByVal Firma As String, _
       ByVal Adresse As IAdresse, _
       ByVal EMail As String, _
       Optional ByVal Anrede As String, _
       Optional ByVal ID As String = vbNullString _
       )
    If ID = vbNullString Then ID = getNewID(IDPerson)
    This.ID = ID
    This.Nachname = Nachname
    This.Vorname = Vorname
    This.Firma = Firma
    This.EMail = EMail
    Set This.Adresse = Adresse
    This.Anrede = Anrede
End Sub

                                
Private Property Get IPerson_Nachname() As String
    IPerson_Nachname = This.Nachname
End Property

                                
Private Property Get IPerson_Vorname() As String
    IPerson_Vorname = This.Vorname
End Property

                                
Private Property Get IPerson_Anrede() As String
    IPerson_Anrede = This.Anrede
End Property

                                
Private Property Get IPerson_Adresse() As IAdresse
    Set IPerson_Adresse = This.Adresse
End Property

                                
Private Property Get IPerson_Firma() As String
    IPerson_Firma = This.Firma
End Property

                                
Private Property Get IPerson_EMail() As String
    IPerson_EMail = This.EMail
End Property

                                
Private Property Get IPerson_ID() As String
    IPerson_ID = This.ID
End Property

                                

