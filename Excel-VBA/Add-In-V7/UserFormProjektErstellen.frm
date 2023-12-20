VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormProjektErstellen 
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6120
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
CreateTinLineProjectFolder Me.CheckBoxEP.value, Me.CheckBoxBR.value, Me.CheckBoxTF.value, Me.CheckBoxPR.value, Me.CheckBoxES.value, Me.TextBoxSPLink.value
If err.Number = 75 Then Application.StatusBar = "Das Projekt wurde nicht erstellt!"
Unload Me
End Sub

Private Sub CommandButtonSharePoint_Click()
ActiveWorkbook.FollowHyperlink Address:="https://rebsamennet.sharepoint.com/:f:/r/sites/PZM-ZH/03_Pub/00_Projekte?csf=1&web=1&e=EGLXoZ"
End Sub

Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ActiveWorkbook.FollowHyperlink Address:="https://rebsamennet.sharepoint.com/:f:/r/sites/PZM-ZH/03_Pub/00_Projekte?csf=1&web=1&e=EGLXoZ"
End Sub

Private Sub UserForm_Initialize()
    Me.TitleLabel.Caption = "Projekt erstellen"
    Me.LabelInstructions.Caption = Globals.Projekt.ProjektOrdnerCAD
End Sub

