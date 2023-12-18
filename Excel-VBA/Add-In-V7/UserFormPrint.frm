VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPrint 
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13440
   OleObjectBlob   =   "UserFormPrint.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








'@Folder "Print"
Option Explicit
Private icons                As UserFormIconLibrary
Private pPlanköpfe           As Collection

Private Sub CommandButtonPrint_Click()
    Dim li                   As ListItem
    Set pPlanköpfe = New Collection
    For Each li In Me.ListViewPlankopf.ListItems
        If li.Checked Then
            pPlanköpfe.Add PlankopfFactory.LoadFromDataBase(Globals.shStoreData.range("A:A").Find(li.ListSubItems.Item(1).text).row)
        End If
    Next

    CreatePlotList pPlanköpfe

End Sub

Private Sub UserForm_Initialize()

    Set icons = New UserFormIconLibrary
    Me.TitleIcon.Picture = icons.IconPrint.Picture
    LoadListViewPlan Me.ListViewPlankopf

End Sub

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

