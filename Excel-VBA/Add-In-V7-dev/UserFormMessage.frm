VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormMessage 
   ClientHeight    =   3480
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5160
   OleObjectBlob   =   "UserFormMessage.frx":0000
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "UserFormMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Erster Versuch f�r eine Custom Fehlermeldung. Implementierung folgt zu einem sp�teren Zeitpunkt."
















'@Folder "Templates"
'@ModuleDescription "Erster Versuch f�r eine Custom Fehlermeldung. Implementierung folgt zu einem sp�teren Zeitpunkt."

Option Explicit

Private icons                As UserFormIconLibrary

Public Enum MSGTyp
    typError = 0
    TypWarning = 1
    TypInfo = 2
End Enum

Public Sub Typ(ByVal MessageType As MSGTyp, ByVal MessageText As String, Optional ByVal Title As String = "Ein Fehler ist aufgetreten!", Optional ByVal OpenLog As Boolean = False)

    Select Case MessageType
        Case 0                                   ' Error
            Me.TitleIcon.Picture = icons.IconError.Picture ' Icon setzen
            Me.TitleLabel.Caption = Title        ' Titel gem�ss Eingabe
            Me.LabelMessage.value = MessageText  ' Message gem�ss Eingabe
            If OpenLog Then Me.CommandButtonLog.Visible = True ' der Button f�r die Anzeige vom Log kann �ber den TYP definiert werden.
        Case 1                                   ' Warning
            Me.TitleIcon.Picture = icons.IconWarning.Picture
            Me.TitleLabel.Caption = Title
            Me.LabelMessage.value = MessageText
            If OpenLog Then Me.CommandButtonLog.Visible = True
        Case 2                                   ' Info
            Me.TitleIcon.Picture = icons.IconInfo.Picture
            Me.TitleLabel.Caption = Title
            Me.LabelMessage.value = MessageText
            If OpenLog Then Me.CommandButtonLog.Visible = True
        Case Else
            Me.TitleIcon.Picture = icons.IconInfo.Picture
            Me.TitleLabel.Caption = Title
            Me.LabelMessage.value = MessageText
            If OpenLog Then Me.CommandButtonLog.Visible = True
    End Select
End Sub

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

Private Sub CommandButtonLog_Click()
    ' �ffnet die Log-Datei im standard-Programm f�r *.log dateien.
    CreateObject("Shell.Application").Open (logger.LogFile)

End Sub

Private Sub UserForm_Initialize()

    Set icons = New UserFormIconLibrary

End Sub

