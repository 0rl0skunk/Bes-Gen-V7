VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormProjektErstellen 
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3240
   OleObjectBlob   =   "UserFormProjektErstellen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormProjektErstellen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Projekt")
Private Sub CommandButtonClose_Click()
    Unload Me
End Sub

Private Sub CommandButtonErstellen_Click()

CreateTinLineProjectFolder Me.CheckBoxEP.Value, Me.CheckBoxBR.Value, Me.CheckBoxTF.Value, Me.CheckBoxPR.Value, Me.CheckBoxES.Value

End Sub

Private Sub UserForm_Initialize()
    Me.TitleLabel.Caption = "Projekt erstellen"
    Me.LabelInstructions.Caption = Globals.Projekt.ProjektOrdnerCAD
End Sub

