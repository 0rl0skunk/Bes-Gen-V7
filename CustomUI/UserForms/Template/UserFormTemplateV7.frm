VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTemplateV7 
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9585.001
   OleObjectBlob   =   "UserFormTemplateV7.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormTemplateV7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@Folder "Objektdaten"
Private icons      As UserFormIconLibrary

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

Public Property Let Title(Value As String)

    Me.TitleLabel.Caption = Value

End Property

Public Property Let icon(Value As String)

    Set icons = New UserFormIconLibrary
    Dim icon As MSForms.control
    
    For Each icon In icons.Controls
    If icon.Tag = "icon" And icon.Name = Value Then
        Me.TitleIcon.Picture = icon.Picture
    End If
    Next

End Property

Public Property Let Instruction(Value As String)

    Me.LabelInstructions.Caption = Value

End Property
