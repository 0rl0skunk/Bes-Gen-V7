VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormRepair 
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   OleObjectBlob   =   "UserFormRepair.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormRepair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Repariert das TinLine Projekt, wenn Fehler mit den Plank�pfen entstehen."









'@Folder("Repair")
'@ModuleDescription "Repariert das TinLine Projekt, wenn Fehler mit den Plank�pfen entstehen."

Option Explicit

Private icons                As UserFormIconLibrary

Private Sub CommandButtonRepair_Click()

    Application.Cursor = xlWait
    If Me.CheckBoxPLAELE.value Then PlanBereinigen "01_EP", "Elektro"
    If Me.CheckBoxPLATF.value Then PlanBereinigen "05_TF", "T�rfachplanung"
    If Me.CheckBoxPLABF.value Then PlanBereinigen "06_BS", "Brandschutzplanung"
    MsgBox "Das Projekt wurde bereinigt.", vbInformation, "Bereinigen abgaschlossen"
    Application.StatusBar = False
    Unload Me
    Application.Cursor = xlDefault

End Sub

Private Sub UserForm_Initialize()

    Set icons = New UserFormIconLibrary
    Me.TitleIcon.Picture = icons.IconRepair.Picture
    Me.TitleLabel.Caption = "Projekt Bereinigen"
    Me.LabelInstructions.Caption = "W�hle aus was alles bereinigt werden soll."


    ' setz die Sichtbarkeit f�r die Checkboxen, damit keine Dateien bereinigt werden welche nicht bestehen.
    ' EP
    Me.CheckBoxPLAELE.Visible = Globals.shProjekt.range("A1").value
    ' PR
    Me.CheckBox1.Visible = Globals.shProjekt.range("A2").value = True
    ' TF
    Me.CheckBoxPLATF.Visible = Globals.shProjekt.range("A4").value = True
    ' BS
    Me.CheckBoxPLABF.Visible = Globals.shProjekt.range("A5").value = True

End Sub

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

Private Sub PlanBereinigen(ByVal Folder As String, ByVal Gewerk As String)
    Dim Plankopf             As IPlankopf

    ' schreibt alle TinPlan und TinPrinzip *.xml files neu
    Geb�udeFolders Globals.Projekt.ProjektOrdnerCAD & "\" & Folder, Gewerk, False

    Dim i                    As Long
    Dim pPlank�pfe           As New Collection
    Set pPlank�pfe = Globals.GetPlank�pfe(Gewerk)
    i = 1

    For Each Plankopf In pPlank�pfe
        ' f�r jeden Plankopf in den zu reparierenden Plank�pfe ...
        Application.StatusBar = "Updating Plankopf " & Plankopf.ID & " | " & i & " von " & pPlank�pfe.Count ' ... schreibt eine Statusmeldung
        PlankopfFactory.RewritePlankopf Plankopf ' ... schreibt den Plankopf neu in die *.xml Datei
        i = i + 1
    Next
End Sub

