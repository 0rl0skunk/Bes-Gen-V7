VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTaskDetail 
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600.001
   OleObjectBlob   =   "UserFormTaskDetail.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormTaskDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False













'@Folder("Tasks")
'@Version "Release V1.0.0"

Option Explicit

Private pTask                As Task
Private icons                As UserFormIconLibrary

Private Sub CommandButtonGebäude_Click()

    'TODO Geschoss UserForm

End Sub

Private Sub CommandButtonGewerk_Click()

    'TODO Gewerk UserForm

End Sub

Private Sub UserForm_Initialize()

    Set icons = New UserFormIconLibrary
    Me.TitleIcon.Picture = icons.IconTodoList.Picture

End Sub

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

Public Sub LoadClass(Task As Task)

    Set pTask = Task
    Set Task = Nothing

    Me.ComboBoxPriorität.value = pTask.Priorität

End Sub

