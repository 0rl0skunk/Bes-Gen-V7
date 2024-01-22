VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPrint 
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13440
   OleObjectBlob   =   "UserFormPrint.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Planköpfe als PDF publizieren. Momentan nur für TinLine Pläne / Elektro"






'@Folder "Print"
'@ModuleDescription "Planköpfe als PDF publizieren. Momentan nur für TinLine Pläne / Elektro"

Option Explicit

Private icons                As UserFormIconLibrary
Private pPlanköpfe           As Collection

Private Sub CheckBoxSelectAll_Click()
    Dim li As ListItem
    For Each li In Me.ListViewPlankopf.ListItems
        li.Checked = Me.CheckBoxSelectAll.value
    Next li
End Sub

Private Sub CommandButtonPrint_Click()
    ' alle ausgewählten Planköpfe publizieren
    Dim li                   As ListItem
    Set pPlanköpfe = New Collection
    For Each li In Me.ListViewPlankopf.ListItems
        If li.Checked Then
            ' für alle publizierbaren Planköpfe schauen ob diese ausgewählt sind, wenn ja zu der collection hinzufügen und sonst überspringen
            pPlanköpfe.Add PlankopfFactory.LoadFromDataBase(Globals.shStoreData.range("A:A").Find(li.ListSubItems.Item(1).Text).row)
        End If
    Next

    ' *.dsd Datei erstellen und publizieren
    CreatePlotList pPlanköpfe

End Sub

Private Sub ListViewPlankopf_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Me.CheckBoxSelectAll.value = False
End Sub

Private Sub UserForm_Initialize()

    Set icons = New UserFormIconLibrary
    Me.TitleIcon.Picture = icons.IconPrint.Picture
    Me.TitleLabel.Caption = "Pläne Publizieren"
    Me.LabelInstructions.Caption = "Planköpfe vom TinLine in PDFs publizieren"

    LoadListViewPlan Me.ListViewPlankopf

End Sub

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

