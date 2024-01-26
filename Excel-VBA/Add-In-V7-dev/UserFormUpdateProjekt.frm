VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormUpdateProjekt 
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5040
   OleObjectBlob   =   "UserFormUpdateProjekt.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormUpdateProjekt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'@Folder "Projekt"

Option Explicit

Private Sub CommandButtonClose_Click()
    Unload Me
End Sub

Private Sub CommandButtonCreate_Click()
    ' Projektphase updaten
    If Me.CheckBoxPRO Then
        Globals.SetWBs
        Globals.shPData.range("ADM_Projektphase").value = Me.ComboBoxProjektPhase.value
        CADFolder.TinLineProjectXML
    End If

    ' Planstand updaten
    Dim rng                  As range
    Dim row                  As range
    Dim ResizeRows           As Long
    If Me.CheckBoxPLA Then
        Set rng = shStoreData.range("A1").CurrentRegion.Offset(2, 0)
        If rng.rows.Count - 3 = 0 Then ResizeRows = 1 Else ResizeRows = rng.rows.Count - 2

        For Each row In rng.Resize(ResizeRows, 1)
            Globals.shStoreData.Cells(row.row, 17).value = Me.ComboBoxStand.value
        Next
    End If

    ' Indexe löschen
    Dim pPlankopf            As IPlankopf
    If Me.CheckBoxIND Then
        For Each pPlankopf In Globals.planköpfe
            pPlankopf.ClearIndex
            PlankopfFactory.ReplaceInDatabase pPlankopf
        Next pPlankopf
    End If

    Unload Me

End Sub

Private Sub UserForm_Initialize()
    Dim arr()                As Variant
    ' Planstand
    Me.ComboBoxStand.Clear
    arr() = getList("PLA_Planstand")
    Me.ComboBoxStand.List = arr()

    ' Projektphase
    Me.ComboBoxProjektPhase.Clear
    arr() = getList("PRO_Projektphase")
    Me.ComboBoxProjektPhase.List = arr()

    Me.TitleLabel.Caption = "Projekt aktualisieren"
    Me.LabelInstructions.Caption = "Alle Planköpfe mit folgenden Informationen aktualisieren."

    Call FormToTaskBar( _
         Form:=Me, _
         IconFromPic:=Me.TitleIcon.Picture, _
         ThumbnailTooltip:=Me.TitleLabel.Caption, _
         HideExcel:=False _
                     )
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ResetTaskbar
End Sub


