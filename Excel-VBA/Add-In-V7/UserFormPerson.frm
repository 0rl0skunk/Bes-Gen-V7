VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPerson 
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3720
   OleObjectBlob   =   "UserFormPerson.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'@Folder("Person")
Option Explicit
Private icons                As UserFormIconLibrary

Private Sub ComboBoxPersonFirma_Change()
    LoadAdress Me.ComboBoxPersonFirma.Value
End Sub

Private Sub CommandButtonClose_Click()
    Unload Me
End Sub

Private Sub CommandButtonCreate_Click()
    PersonFactory.AddToDatabase FormToPerson
End Sub

Private Function FormToPerson() As IPerson
    Dim NewPerson            As New IPerson
    Set NewPerson = PersonFactory.Create( _
                    Vorname:=Me.TextBoxPersonVorname.Value, _
                    Nachname:=Me.TextBoxPersonNachname.Value, _
                    Firma:=Me.ComboBoxPersonFirma.Value, _
                    Anrede:=Me.ComboBoxPersonAnrede.Value, _
                    Adresse:=AdressFactory.Create( _
                              Strasse:=Me.TextBoxADRStrasse.Value, _
                    PLZ:=Me.TextBoxADRPLZ.Value, _
                    Ort:=Me.TextBoxADROrt.Value), _
                    EMail:=Me.TextBoxPersonEMail.Value _
                            )
    Set FormToPerson = NewPerson
End Function

Private Sub UserForm_Initialize()
    Me.ComboBoxPersonAnrede.List = Array("Herr", "Frau", "Du")

    Dim e                    As range
    With CreateObject("Scripting.Dictionary")
        For Each e In Globals.shAdress.range("ADM_Firmen")
            If Not .Exists(e.Value) Then
                .Add e.Value, Nothing
            End If
        Next e

        Me.ComboBoxPersonFirma.List = .Keys
    End With

End Sub

Private Sub LoadAdress(ByVal Firma As String)
    Dim e                    As range
    For Each e In Globals.shAdress.range("ADM_Firmen")
        If e.Value = Firma Then
            Me.TextBoxADRStrasse.Value = e.Offset(0, 1).Value
            Me.TextBoxADRPLZ.Value = e.Offset(0, 2).Value
            Me.TextBoxADROrt.Value = e.Offset(0, 3).Value
        End If
    Next
End Sub


