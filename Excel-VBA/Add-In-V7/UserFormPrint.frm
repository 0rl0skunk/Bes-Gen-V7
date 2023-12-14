VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPrint 
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600.001
   OleObjectBlob   =   "UserFormPrint.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@Folder "Print"
Private icons                As UserFormIconLibrary

Private Sub UserForm_Initialize()

    Set icons = New UserFormIconLibrary
    Me.TitleIcon.Picture = icons.IconPrint.Picture

End Sub

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

