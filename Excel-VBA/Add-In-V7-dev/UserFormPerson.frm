VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPerson 
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3720
   OleObjectBlob   =   "UserFormPerson.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Einfaches Auswählen und definieren eines Layouts für Plan und Prinzip Planköpfe."


















'@Folder("Person")
'@ModuleDescription "Einfaches Auswählen und definieren eines Layouts für Plan und Prinzip Planköpfe."

Option Explicit

Private icons                As UserFormIconLibrary

Private Sub ComboBoxPersonFirma_Change()

    LoadAdress Me.ComboBoxPersonFirma.value

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
                    Vorname:=Me.TextBoxPersonVorname.value, _
                    Nachname:=Me.TextBoxPersonNachname.value, _
                    Firma:=Me.ComboBoxPersonFirma.value, _
                    Anrede:=Me.ComboBoxPersonAnrede.value, _
                    Adresse:=AdressFactory.Create( _
                              Strasse:=Me.TextBoxADRStrasse.value, _
                    PLZ:=Me.TextBoxADRPLZ.value, _
                    Ort:=Me.TextBoxADROrt.value), _
                    EMail:=Me.TextBoxPersonEMail.value _
                            )
    Set FormToPerson = NewPerson

End Function

Private Sub UserForm_Initialize()

    Set icons = New UserFormIconLibrary
    Me.TitleIcon.Picture = icons.IconAddressBook.Picture
    Me.TitleLabel.Caption = "Adresse / Person erfassen"
    Me.LabelInstructions.Caption = "Neue Adresse oder Person für E-Mail Versand erfassen"

    Me.ComboBoxPersonAnrede.List = Array("Herr", "Frau", "Du")

    Dim e                    As range
    With CreateObject("Scripting.Dictionary")
        For Each e In Globals.shAdress.range("ADM_Firmen")
            If Not .Exists(e.value) Then
                .Add e.value, Nothing
            End If
        Next e

        Me.ComboBoxPersonFirma.List = .Keys
    End With

End Sub

Private Sub LoadAdress(ByVal Firma As String)
    ' Adresse in Userform anzeigen
    Dim e                    As range
    For Each e In Globals.shAdress.range("ADM_Firmen")
        If e.value = Firma Then
            Me.TextBoxADRStrasse.value = e.Offset(0, 1).value
            Me.TextBoxADRPLZ.value = e.Offset(0, 2).value
            Me.TextBoxADROrt.value = e.Offset(0, 3).value
        End If
    Next

End Sub


