VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormMessage 
   ClientHeight    =   3480
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5160
   OleObjectBlob   =   "UserFormMessage.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "UserFormMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder "Templates"
Option Explicit
Private icons                As UserFormIconLibrary

Public Sub typeError(ByVal MessageText As String, Optional ByVal Title As String = "Ein Fehler ist aufgetreten!", Optional ByVal OpenLog As Boolean = False)

    Me.TitleIcon.Picture = icons.IconError.Picture
    Me.TitleLabel.Caption = Title
    Me.LabelMessage.Value = MessageText
    If OpenLog Then Me.CommandButtonLog.Visible = True

End Sub

Public Sub typeWarning(ByVal MessageText As String, Optional ByVal Title As String = "Achtung", Optional ByVal OpenLog As Boolean = False)

    Me.TitleIcon.Picture = icons.IconWarning.Picture
    Me.TitleLabel.Caption = Title
    Me.LabelMessage.Value = MessageText
    If OpenLog Then Me.CommandButtonLog.Visible = True

End Sub

Public Sub typeInfo(ByVal MessageText As String, Optional ByVal Title As String = "Info", Optional ByVal OpenLog As Boolean = False)

    Me.TitleIcon.Picture = icons.IconInfo.Picture
    Me.TitleLabel.Caption = Title
    Me.LabelMessage.Value = MessageText
    If OpenLog Then Me.CommandButtonLog.Visible = True

End Sub

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

Private Sub CommandButtonLog_Click()

    CreateObject("Shell.Application").Open (LOG.LOGFile)

End Sub

Private Sub UserForm_Initialize()

    Set icons = New UserFormIconLibrary

End Sub

