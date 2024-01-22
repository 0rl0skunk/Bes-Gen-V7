VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormInfo 
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600.001
   OleObjectBlob   =   "UserFormInfo.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







'@Folder("Info Version")
'@ModuleDescription "Log-Anzeige damit nicht immer die Datei geffnet werden muss."

Option Explicit

Private icons                As UserFormIconLibrary

Private Sub CommandButtonClear_Click()

    LogClear

    Dim fso                  As New FileSystemObject
    Dim strText              As TextStream
    Set strText = fso.OpenTextFile(logger.LogFile, ForReading)
    Me.TextBoxLOG.value = strText.ReadAll

End Sub

Private Sub UserForm_Initialize()

    Set icons = New UserFormIconLibrary
    Me.TitleIcon.Picture = icons.IconXML.Picture
    Me.TitleLabel.Caption = "LOG"
    Me.LabelInstructions.Caption = logger.LogFile

    Dim fso                  As New FileSystemObject
    Dim strText              As TextStream
    Set strText = fso.OpenTextFile(logger.LogFile, ForReading)
    Me.TextBoxLOG.value = strText.ReadAll
    strText.Close

End Sub

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

