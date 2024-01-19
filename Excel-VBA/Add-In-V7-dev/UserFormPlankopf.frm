VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPlankopf 
   ClientHeight    =   11760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960.001
   OleObjectBlob   =   "UserFormPlankopf.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormPlankopf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Erstellen von Plank�pfen f�r alle Gewerke. Automatisches Einf�gen der Plank�pfe f�r Elektropl�ne �ber das Modul PlankopfFactory"


'@Folder "Plankopf"
'@ModuleDescription "Erstellen von Plank�pfen f�r alle Gewerke. Automatisches Einf�gen der Plank�pfe f�r Elektropl�ne �ber das Modul PlankopfFactory"

Option Explicit

Public Enum EnumIcon
    Add = 0
    Edit = 1
End Enum

Private icons                As UserFormIconLibrary
Private pPlankopf            As IPlankopf
Public PlankopfCopyFrom      As IPlankopf
Private pProjekt             As IProjekt
Private shPData              As Worksheet
Private shGeb�ude            As Worksheet

Public Sub setIcons(ByVal icon As EnumIcon)
    ' Icon anpassen f�r erstellen oder Bearbeiten
    Select Case icon
    Case 0
        Me.TitleIcon.Picture = icons.IconAddProperties.Picture
        Me.TitleLabel.Caption = "Plankopf erstellen"
    Case 1
        Me.TitleIcon.Picture = icons.IconEditProperties.Picture
        Me.TitleLabel.Caption = "Plankopf bearbeiten"
    End Select

End Sub

Private Function validateUserForm(Optional skipIndex As Boolean = False) As Boolean
    ' sind alle wichtigen Infos mitgegeben, nicht korrekt ausgef�llte Infos werden markiert
    Dim oControl As MSForms.control
    Dim oComboBox As MSForms.ComboBox
    Dim oTextBox As MSForms.TextBox
    
    Dim warningColor: warningColor = RGB(255, 255, 0)
    Dim errorColor: errorColor = RGB(255, 0, 0)
    
    validateUserForm = True
    
    Select Case Me.MultiPageTyp.value
    Case 0                                       ' Plan
        For Each oControl In Me.Frame11.Controls
            If oControl.Name Like "*ComboBox*" Then
                Set oComboBox = oControl
                If oComboBox.value = "-- Bitte w�hlen --" Then oComboBox.BackColor = errorColor: validateUserForm = False
            End If
        Next oControl
    Case 1                                       ' Schema
        For Each oControl In Me.Frame10.Controls
            If oControl.Name Like "*ComboBox*" Then
                Set oComboBox = oControl
                If oComboBox.value = "-- Bitte w�hlen --" Then oComboBox.BackColor = errorColor: validateUserForm = False
            ElseIf oControl.Name Like "*TextBox*" Then
                Set oTextBox = oControl
                If oTextBox.value = vbNullString Then oTextBox.BackColor = errorColor: validateUserForm = False
            End If
        Next oControl
    Case 2                                       ' Prinzip
        For Each oControl In Me.Frame12.Controls
            If oControl.Name Like "*ComboBox*" Then
                Set oComboBox = oControl
                If oComboBox.value = "-- Bitte w�hlen --" Then oComboBox.BackColor = errorColor: validateUserForm = False
            End If
        Next oControl
    End Select
    
    ' Projektinfos
    For Each oControl In Me.Frame3.Controls
        If oControl.Name Like "*ComboBox*" Then
            Set oComboBox = oControl
            If oComboBox.value = "-- Bitte w�hlen --" Then oComboBox.BackColor = errorColor: validateUserForm = False
        End If
    Next oControl
    
    ' Layout
    For Each oControl In Me.FrameLayout.Controls
        If oControl.Name Like "*ComboBox*" Then
            Set oComboBox = oControl
            If oComboBox.value = "-- Bitte w�hlen --" Then oComboBox.BackColor = errorColor: validateUserForm = False
        End If
    Next oControl
        
    ' Allgemeine Infos
    For Each oControl In Me.Frame6.Controls
        If oControl.Name Like "*ComboBox*" Then
            Set oComboBox = oControl
            If oComboBox.value = "-- Bitte w�hlen --" Then oComboBox.BackColor = errorColor: validateUserForm = False
        End If
    Next oControl
        
    ' Gepr�ft
    For Each oControl In Me.FramePlaninfo.Controls
        If oControl.Name Like "*TextBox*" Then
            Set oTextBox = oControl
            If oTextBox.value = vbNullString Then oTextBox.BackColor = warningColor
        End If
    Next oControl
    
    If Not Me.TextBoxIndexKlartext.value = vbNullString Then
        ' Wenn ein Index erstellt wurde aber nicht hinzugef�gt ist.
        Select Case MsgBox("Es wurde ein Index erstellt jedoch nicht korrekt erfasst." & vbNewLine & "Soll der Index hinzugef�gt werden?", vbYesNo, "Index erfassen")
        Case vbYes
            CommandButtonIndexErstellen_Click    ' Index erstellen
        Case vbNo
        End Select
    End If
End Function

Private Sub CommandButtonCreate_Click()
    ' Plankopf in Datenbank schreiben
    If Not validateUserForm Then: MsgBox "Einige Angaben sind nicht korrekt ausgef�llt!" & vbNewLine & "Bitte Pr�fe deine Eingaben.", vbCritical, "Eingaben pr�fen": Exit Sub
    
    If Me.CommandButtonCreate.Caption = "Update" Then
        ' Ersetzen / Updaten
        If PlankopfFactory.ReplaceInDatabase(FormToPlankopf) Then Unload Me
    Else
        ' Neu erstellen
        If PlankopfFactory.AddToDatabase(FormToPlankopf) Then Unload Me
    End If

End Sub

Private Sub CommandButtonBeschriftungAktualisieren_Click()
    ' Beschriftungen und Plannummer neu erstellen
    If Not validateUserForm(True) Then: Me.TextBoxBeschriftungDateiname.value = vbNullString: Me.TextBoxBeschriftungPlannummer.value = vbNullString: Me.TextBoxPlan�berschrift.value = vbNullString: Exit Sub
    
    Set pPlankopf = FormToPlankopf
    Me.TextBoxBeschriftungPlannummer.value = pPlankopf.Plannummer
    Me.TextBoxBeschriftungDateiname.value = pPlankopf.PDFFileName
    Me.TextBoxPlan�berschrift.value = pPlankopf.Plan�berschrift
    Me.BesID.Caption = pPlankopf.ID
    Me.LabelDWGFileName.Caption = pPlankopf.DWGFileName
    Me.LabelXMLFileName.Caption = pPlankopf.XMLFileName
    Me.LabelFolderName.Caption = pPlankopf.FolderName

End Sub

Private Sub CommandButtonIndexErstellen_Click()
    ' Neuer Index f�r den ge�ffneten Plankopf erstellen
    Dim Index                As IIndex: Set Index = IndexFactory.Create( _
        IDPlan:=pPlankopf.ID, _
        GezeichnetPerson:=Me.TextBoxIndexGez.value, _
        GezeichnetDatum:=Me.TextBoxIndexGezDatum.value, _
        Klartext:=Me.TextBoxIndexKlartext.value, _
        Letter:=Me.TextBoxIndexLetter.value _
                 )
    IndexFactory.AddToDatabase Index
    pPlankopf.AddIndex Index

    LoadIndexes

    Me.TextBoxIndexGez.value = vbNullString
    Me.TextBoxIndexGezDatum.value = vbNullString
    Me.TextBoxIndexKlartext.value = vbNullString
    Me.TextBoxIndexLetter.value = vbNullString

End Sub

Private Sub CommandButtonIndexL�schen_Click()
    ' Ausgew�hlten Index l�schen
    Dim ID                   As String
    Dim li As ListItem
    
    For Each li In Me.ListViewIndex.ListItems
        If li.Checked Then
            ID = li.ListSubItems(1)
            IndexFactory.DeleteFromDatabase ID
        End If
    Next

    pPlankopf.ClearIndex
    IndexFactory.GetIndexes pPlankopf

    LoadIndexes

End Sub

Private Sub CommandLayoutW�hlen_Click()
    ' UserFormLayout �ffnen und diese �bernehmen
    Dim frm                  As New UserFormLayout
    frm.load Me.ComboBoxLayoutFormat.value, Me.TextBoxLayoutMasstab.value, Me.MultiPageTyp.value
    frm.Show 1
    If frm.CheckBoxLoad Then
        Me.ComboBoxLayoutFormat.value = frm.TextBoxFormatH.value & "H" & frm.TextBoxFormatB.value & "B"
    End If
    Set frm = Nothing

End Sub

'@Ignore ProcedureNotUsed
Private Sub EditDWG_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' DWG-Datei im TinLine �ffnen
    TinLine.setTinProject pProjekt.ProjektOrdnerCAD
    Select Case Me.MultiPageTyp.value
    Case 0                                       'Plan
        TinLine.setTinPlanBibliothek
    Case 1                                       'Prinzip
        TinLine.setTinPrinzipBibiothek
    End Select

    CreateObject("Shell.Application").Open (FormToPlankopf.dwgFile)

End Sub

Private Sub MultiPageTyp_Change()
    ' Anpassungen wenn der Plantyp ge�ndert wird
    ' TODO Remove Geschoss "Gesamt" from Plan and Schema Beschriftungen
    Select Case Me.MultiPageTyp.value
    Case 0                                       'PLA
        Me.ComboBoxGeb�ude.Enabled = True
        Me.ComboBoxGeb�udeTeil.Enabled = True
        Me.ComboBoxGeschoss.Enabled = True
    Case 1                                       'SCH
        Me.ComboBoxGeb�ude.Enabled = True
        Me.ComboBoxGeb�udeTeil.Enabled = True
        Me.ComboBoxGeschoss.Enabled = True
    Case 2                                       'PRI
        Me.ComboBoxGeb�ude.value = "Gesamt"
        Me.ComboBoxGeb�udeTeil.value = "Gesamt"
        Me.ComboBoxGeschoss.value = "Gesamt"
        Me.ComboBoxGeb�ude.Enabled = False
        Me.ComboBoxGeb�udeTeil.Enabled = False
        Me.ComboBoxGeschoss.Enabled = False
    End Select

    If Me.ComboBoxGeb�ude.ListCount = 1 Then
        Me.ComboBoxGeb�ude.value = Me.ComboBoxGeb�ude.List(0)
        Me.ComboBoxGeb�ude.Enabled = False
    Else
        Me.ComboBoxGeb�ude.Enabled = True
    End If

    If Me.ComboBoxGeb�udeTeil.ListCount = 1 Then
        Me.ComboBoxGeb�udeTeil.value = Me.ComboBoxGeb�udeTeil.List(0)
        Me.ComboBoxGeb�udeTeil.Enabled = False
    Else
        Me.ComboBoxGeb�udeTeil.Enabled = True
    End If

End Sub

'@Ignore ProcedureNotUsed
Private Sub Preview_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Plankopfpreview �ffnen
    Dim frm                  As New UserFormPlankopfPreview
    frm.LoadClass FormToPlankopf, pProjekt
    frm.Show 1

End Sub

Private Sub TextBoxPlanInfoDatumGepr�ft_Change()
    Me.TextBoxPlanInfoDatumGepr�ft.BackColor = SystemColorConstants.vbWindowBackground
End Sub

Private Sub TextBoxPlanInfoK�rzelGepr�ft_Change()
    Me.TextBoxPlanInfoK�rzelGepr�ft.BackColor = SystemColorConstants.vbWindowBackground
End Sub

Private Sub UserForm_Initialize()

    Set icons = New UserFormIconLibrary

    ' ComboBox Listen auf�llen

    ' Unterprojekt
    ' Array mit Unterprojekt Name und Nummer nebeneinander

    Dim arr()                As Variant

    ' populate unterprojekt if there is only one
    'arr() = getList("Unterprojekte")
    Me.ComboBoxUnterprojekt.List = getList("PRO_Unterprojekte")
    If Me.ComboBoxUnterprojekt.ListCount = 1 Then
        Me.ComboBoxUnterprojekt.value = Me.ComboBoxUnterprojekt.List(0)
        Me.ComboBoxUnterprojekt.Enabled = False
    End If
    Me.LabelProjektphase.Caption = Globals.shPData.range("ADM_Projektphase").value

    ' Planstand
    Me.ComboBoxStand.Clear
    arr() = getList("PLA_Planstand")
    Me.ComboBoxStand.List = arr()

    ' Planart
    Me.ComboBoxEPArt.Clear

    ' Haupt Gewerk
    Me.ComboBoxEPHGewerk.Clear
    Me.ComboBoxESHGewerk.Clear
    Me.ComboBoxPRHGewerk.Clear
    arr() = getList("PRO_Hauptgewerk")
    Me.ComboBoxEPHGewerk.List = arr()
    Me.ComboBoxESHGewerk.List = arr()
    Me.ComboBoxPRHGewerk.List = arr()

    ' Geb�udeTeil
    Me.ComboBoxGeb�udeTeil.Clear
    Me.ComboBoxGeb�udeTeil.List = getList("PRO_Geb�udeteil")
    If Me.ComboBoxGeb�udeTeil.ListCount = 1 Then
        Me.ComboBoxGeb�udeTeil.value = Me.ComboBoxGeb�udeTeil.List(0)
        Me.ComboBoxGeb�udeTeil.Enabled = False
    Else
        Me.ComboBoxGeb�ude.Enabled = True
    End If
    ' Geb�ude
    Me.ComboBoxGeb�ude.Clear
    Me.ComboBoxGeb�ude.List = getList("PRO_Geb�ude")

    Me.MultiPageTyp.value = 0
    ' Formate
    Me.ComboBoxLayoutFormat.Clear
    arr() = getList("PLA_Format")
    Me.ComboBoxLayoutFormat.List = arr()

    ' Massstab
    Me.TextBoxLayoutMasstab.value = "1:50"
    Me.LabelProjektnummer.Caption = Globals.shPData.range("ADM_Projektnummer").value

    Me.TextBoxPlanInfoDatumGezeichnet.value = Format$(Now, "DD.MM.YYYY")
    Me.TextBoxPlanInfoK�rzelGezeichnet.value = getUserName
    
    Me.TextBoxIndexGez = getUserName
    Me.TextBoxIndexGezDatum = Format$(Now, "DD.MM.YYYY")

    writelog LogInfo, "UserFormPlankopf > Inizialise complete"
    
    Application.Cursor = xlDefault

End Sub

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

Private Sub LoadIndexes()
    ' Indexe vom Plankopf laden und ListView abf�llen
    Dim ind                  As IIndex
    Dim li                   As ListItem

    With Me.ListViewIndex
        .ListItems.Clear
        .View = lvwReport
        .CheckBoxES = True
        .Gridlines = True
        .FullRowSelect = True
        With .ColumnHeaders
            .Clear
            .Add , , vbNullString, 20
            .Add , , vbNullString, 0
            .Add , , "Index", 20
            .Add , , "Gezeichnet", 40
            .Add , , "Datum", 60
            .Add , , "Beschreibung", 250
        End With

        For Each ind In pPlankopf.Indexes
            Set li = .ListItems.Add()
            li.ListSubItems.Add , , ind.IndexID
            li.ListSubItems.Add , , ind.Index
            li.ListSubItems.Add , , Split(ind.Gezeichnet, " ; ")(0)
            li.ListSubItems.Add , , Split(ind.Gezeichnet, " ; ")(1)
            li.ListSubItems.Add , , ind.Klartext
        Next
    End With

End Sub

Public Sub LoadClass(Plankopf As IPlankopf, ByVal Projekt As IProjekt, Optional ByVal copy As Boolean = False)
    ' Usercontrols von Klasse laden
    Set pProjekt = Projekt

    Set pPlankopf = Plankopf
    Set Plankopf = Nothing
    Dim Planstand            As String
    Dim PLANTYP              As Long
    Dim Gewerk               As String
    Dim UnterGewerk          As String


    Select Case pPlankopf.PLANTYP
    Case "PLA"
        Me.MultiPageTyp.value = 0
        Me.ComboBoxEPHGewerk.value = pPlankopf.Gewerk
        Me.ComboBoxEPUGewerk.value = pPlankopf.UnterGewerk
        Me.ComboBoxEPArt.value = pPlankopf.Planart
    Case "SCH"
        Me.MultiPageTyp.value = 1
        Me.ComboBoxESHGewerk.value = pPlankopf.Gewerk
        Me.ComboBoxESUGewerk.value = pPlankopf.UnterGewerk
    Case "PRI"
        Me.MultiPageTyp.value = 2
        Me.ComboBoxPRHGewerk.value = pPlankopf.Gewerk
        Me.ComboBoxPRUGewerk.value = pPlankopf.UnterGewerk
    End Select

    ' f�llt die Eingabefelder gem�ss geladenem Objekt aus
    Me.ComboBoxGeb�ude.value = pPlankopf.Geb�ude
    Me.ComboBoxGeb�udeTeil.value = pPlankopf.Geb�udeteil
    Me.ComboBoxGeschoss.value = pPlankopf.Geschoss
    Me.ComboBoxLayoutFormat.value = pPlankopf.LayoutGr�sse
    Me.TextBoxLayoutMasstab.value = pPlankopf.LayoutMasstab
    Me.TextBoxPlanInfoDatumGezeichnet.value = pPlankopf.GezeichnetDatum
    Me.TextBoxPlanInfoK�rzelGezeichnet.value = pPlankopf.GezeichnetPerson
    Me.TextBoxPlanInfoDatumGepr�ft.value = pPlankopf.Gepr�ftDatum
    Me.TextBoxPlanInfoK�rzelGepr�ft.value = pPlankopf.Gepr�ftPerson
    Me.TextBoxPlan�berschrift.value = pPlankopf.Plan�berschrift
    Me.LabelDWGFileName.Caption = pPlankopf.DWGFileName
    Me.LabelXMLFileName.Caption = pPlankopf.XMLFileName
    Me.LabelFolderName.Caption = pPlankopf.FolderName
    Me.TBAnlageteil.value = pPlankopf.AnlageNummer
    Me.ComboBoxESAnlageTyp.value = pPlankopf.AnlageTyp
    LoadIndexes

    Me.ComboBoxStand.value = pPlankopf.LayoutPlanstand

    If Not copy Then
        ' disable all inputs which should only be set once
        Me.MultiPageTyp.Enabled = False
        Me.ComboBoxEPArt.Enabled = False
        Me.ComboBoxEPHGewerk.Enabled = False
        Me.ComboBoxEPUGewerk.Enabled = False
        Me.ComboBoxESAnlageTyp.Enabled = False
        Me.ComboBoxESHGewerk.Enabled = False
        Me.ComboBoxESUGewerk.Enabled = False
        Me.ComboBoxGeb�ude.Enabled = False
        Me.ComboBoxGeb�udeTeil.Enabled = False
        Me.ComboBoxGeschoss.Enabled = False
        Me.ComboBoxPRHGewerk.Enabled = False
        Me.ComboBoxPRUGewerk.Enabled = False

        Me.CommandButtonCreate.Caption = "Update"
        Me.BesID.Caption = pPlankopf.ID
        Me.TinLineID.Caption = pPlankopf.IDTinLine
    Else
        Me.BesID.Caption = getNewID(IDPlankopf)
        pPlankopf.ID = Me.BesID.Caption
        Dim Index            As IIndex
        For Each Index In pPlankopf.Indexes
            Index.PlanID = pPlankopf.ID
            IndexFactory.AddToDatabase Index
        Next
    End If

    CommandButtonBeschriftungAktualisieren_Click

End Sub

Public Sub CopyPlankopf(Plankopf As IPlankopf, ByVal Projekt As IProjekt, ByVal CopyIndex As Boolean)
    ' Plankopf Kopieren mit oder ohne Indexe
    If CopyIndex Then
        Set Plankopf.Indexes = PlankopfCopyFrom.Indexes
        Set PlankopfCopyFrom = Nothing
    End If

    LoadClass Plankopf, Projekt, True

End Sub

Private Function FormToPlankopf() As IPlankopf
    ' UserForm in ein Plankopf-Objekt umwandeln
    Dim PLANTYP              As String
    Dim Gewerk               As String
    Dim UnterGewerk          As String
    Dim ID                   As String

    If Me.BesID.Caption = "ID" Then ID = getNewID(IDPlankopf)

    Select Case Me.MultiPageTyp.value
    Case 0
        PLANTYP = "PLA"
        Gewerk = Me.ComboBoxEPHGewerk.value
        UnterGewerk = Me.ComboBoxEPUGewerk.value
    Case 1
        PLANTYP = "SCH"
        Gewerk = Me.ComboBoxESHGewerk.value
        UnterGewerk = Me.ComboBoxESUGewerk.value
    Case 2
        PLANTYP = "PRI"
        Gewerk = Me.ComboBoxPRHGewerk.value
        UnterGewerk = Me.ComboBoxPRUGewerk.value
    Case Else
        PLANTYP = "PLA"
        Gewerk = Me.ComboBoxEPHGewerk.value
        UnterGewerk = Me.ComboBoxEPUGewerk.value
    End Select

    If pProjekt Is Nothing Then Set pProjekt = Globals.Projekt
    Set FormToPlankopf = PlankopfFactory.Create( _
                         Projekt:=pProjekt, _
                         GezeichnetPerson:=Me.TextBoxPlanInfoK�rzelGezeichnet.value, _
                         GezeichnetDatum:=Me.TextBoxPlanInfoDatumGezeichnet.value, _
                         Gepr�ftPerson:=Me.TextBoxPlanInfoK�rzelGepr�ft.value, _
                         Gepr�ftDatum:=Me.TextBoxPlanInfoDatumGepr�ft.value, _
                         Geb�ude:=Me.ComboBoxGeb�ude.value, _
                         Geb�udeteil:=Me.ComboBoxGeb�udeTeil.value, _
                         Gewerk:=Gewerk, _
                         UnterGewerk:=UnterGewerk, _
                         Geschoss:=Me.ComboBoxGeschoss.value, _
                         Format:=Me.ComboBoxLayoutFormat.value, _
                         Masstab:=Me.TextBoxLayoutMasstab.value, _
                         Stand:=Me.ComboBoxStand.value, _
                         PLANTYP:=PLANTYP, _
                         Planart:=Me.ComboBoxEPArt.value, _
                         TinLineID:=Me.TinLineID.Caption, _
                         SkipValidation:=False, _
                         Plan�berschrift:=Me.TextBoxPlan�berschrift.value, _
                         ID:=Me.BesID.Caption, _
                         AnlageTyp:=Me.ComboBoxESAnlageTyp.value, _
                         AnlageNummer:=Me.TBAnlageteil.value _
                                        )

End Function

'-------------------------------------------------------- ComboBox_Change Events ---------------------------------------------------------

Private Sub ComboBoxEPArt_Change()

    Me.ComboBoxEPArt.BackColor = SystemColorConstants.vbWindowBackground

End Sub

Private Sub ComboBoxEPUGewerk_Change()

    Me.ComboBoxEPUGewerk.BackColor = SystemColorConstants.vbWindowBackground

    If Me.ComboBoxEPUGewerk.value = vbNullString Then
        Me.ComboBoxEPUGewerk.value = "-- Bitte w�hlen --"
    End If

End Sub

Private Sub ComboBoxEPHGewerk_Change()

    Dim row                  As Variant          ' Reihe in welcher der Kontext gefunden wurde
    Dim col                  As Long             ' Spalte in welcher der Kontext gefunden wurde
    Dim lastrow              As Long             ' Die Letzte verwendete Zeile in der Spalte
    Dim ws                   As Worksheet: Set ws = Globals.shPData

    If Me.ComboBoxEPHGewerk.value = "-- Bitte w�hlen --" Then
        ' wenn keine Auswahl getroffen wurde
        Me.ComboBoxEPUGewerk.Enabled = False
        Me.ComboBoxEPUGewerk.Clear
        Me.ComboBoxEPUGewerk.value = "-- Bitte w�hlen --"
        Me.ComboBoxEPArt.Enabled = False
        Me.ComboBoxEPArt.Clear
        Me.ComboBoxEPArt.value = "-- Bitte w�hlen --"
        Exit Sub
    End If

    If Me.ComboBoxEPHGewerk.value = vbNullString Then Exit Sub

    Me.ComboBoxEPArt.Enabled = True
    Me.ComboBoxEPUGewerk.Enabled = True

    Me.ComboBoxEPHGewerk.BackColor = SystemColorConstants.vbWindowBackground
    Dim HGewerk              As String
    HGewerk = WLookup(Me.ComboBoxEPHGewerk.value, ws.range("PRO_Hauptgewerk"), 2)

    If Not IsError(Application.Match(HGewerk & " PLA", ws.range("10:10"), 0)) Then
        ' checkt ob das Gewerk vorhanden ist und verwendet werden kann
1       col = Application.Match(HGewerk & " PLA", ws.range("10:10"), 0) ' findet die aktuelle Spalte mit dem ausgew�hlten Wert f�r das Hauptgewerk
        lastrow = Application.CountA(ws.Cells(13, col).EntireColumn) + 10 ' findet die Letzte Reihe in welcher der Wert ausgew�hlt wurde

        Me.ComboBoxEPUGewerk.Clear               ' l�scht die aktuelle Liste der ComboBox

        For row = 13 To lastrow                  ' loopt durch alle Reihen und f�gt diese der Liste hinzu wenn diese nicht leer sind
            If ws.Cells(row, col).value <> vbNullString Then
                Me.ComboBoxEPUGewerk.AddItem ws.Cells(row, col).value
            End If
        Next row

        Me.ComboBoxEPUGewerk.value = "-- Bitte w�hlen --" ' Setzt den default wert der ComboBox

        ' --- Planart ---
2       col = Application.Match(HGewerk, ws.range("9:9"), 0) ' findet die aktuelle Spalte mit dem ausgew�hlten Wert f�r das Hauptgewerk
        lastrow = Application.CountA(ws.Cells(13, col).EntireColumn) + 10 ' findet die Letzte Reihe in welcher der Wert ausgew�hlt wurde

        Me.ComboBoxEPArt.Clear                   ' l�scht die aktuelle Liste der ComboBox

        For row = 13 To lastrow                  ' loopt durch alle Reihen und f�gt diese der Liste hinzu wenn diese nicht leer sind
            If ws.Cells(row, col).value <> vbNullString Then
                Me.ComboBoxEPArt.AddItem ws.Cells(row, col).value
            End If
        Next row

        Me.ComboBoxEPArt.value = "-- Bitte w�hlen --" ' Setzt den default wert der ComboBox

    End If

End Sub

Private Sub ComboBoxESAnlageTyp_Change()

    Me.ComboBoxESAnlageTyp.BackColor = SystemColorConstants.vbWindowBackground
    If Me.ComboBoxESAnlageTyp.value = "Steuerung" Then
        Me.ComboBoxESAnlageTyp.ControlTipText = "Genaue Steuerung im Klartext definieren!"
    Else
        Me.ComboBoxESAnlageTyp.ControlTipText = "W�hle den Anlagentyp des zu beschriftenden Schemas aus."
    End If

End Sub

Private Sub ComboBoxESHGewerk_Change()
    ' Funktionsweise gem. Kommentaren ComboboxEPHGewerk
    Dim row                  As Variant
    Dim col                  As Long
    Dim lastrow              As Long
    Dim ws                   As Worksheet: Set ws = Globals.shPData

    Dim HGewerk              As String
    HGewerk = WLookup(Me.ComboBoxESHGewerk.value, ws.range("PRO_Hauptgewerk"), 2)

    Me.ComboBoxESHGewerk.BackColor = SystemColorConstants.vbWindowBackground
    If Me.ComboBoxESHGewerk.value = "-- Bitte W�hlen --" Then
        Me.ComboBoxESAnlageTyp.Enabled = False
        Me.ComboBoxESUGewerk.Enabled = False
        Me.ComboBoxESAnlageTyp.Clear
        Me.ComboBoxESUGewerk.Clear
        Me.ComboBoxESAnlageTyp.value = "-- Bitte w�hlen --"
        Me.ComboBoxESUGewerk.value = "-- Bitte w�hlen --"
        Exit Sub
    End If

1   col = Application.WorksheetFunction.Match(HGewerk & " SCH", ws.range("10:10"), 0)
2   lastrow = Application.WorksheetFunction.CountA(ws.Cells(13, col).EntireColumn) + 11

    Me.ComboBoxESUGewerk.Clear

    Me.ComboBoxESAnlageTyp.Enabled = True
    Me.ComboBoxESUGewerk.Enabled = True

    For row = 13 To lastrow
        If ws.Cells(row, col).value <> vbNullString Then
            Me.ComboBoxESUGewerk.AddItem ws.Cells(row, col).value
        End If
    Next row

    Me.ComboBoxESUGewerk.value = "-- Bitte w�hlen --"

End Sub

Private Sub ComboBoxESUGewerk_Change()
    ' Funktionsweise gem. Kommentaren ComboboxEPHGewerk
    Dim col                  As Variant
    Dim row                  As Variant
    Dim lastrow              As Variant
    Dim ws                   As Worksheet: Set ws = Globals.shPData

    Me.ComboBoxESUGewerk.BackColor = SystemColorConstants.vbWindowBackground

    If Me.ComboBoxESUGewerk.value = "-- Bitte w�hlen --" Then Exit Sub
    If Me.ComboBoxESUGewerk.value = vbNullString Then Exit Sub
    Select Case Me.ComboBoxESHGewerk.value
    Case "Elektro"
        If Not IsError(Application.Match("Anlagetyp " & Me.ComboBoxESUGewerk.value, ws.range("12:12"), 0)) Then
1           col = Application.Match("Anlagetyp " & Me.ComboBoxESUGewerk.value, ws.range("12:12"), 0)
            lastrow = Application.WorksheetFunction.CountA(ws.Cells(13, col).EntireColumn) + 12
            Me.ComboBoxESAnlageTyp.Clear
            For row = 13 To lastrow
                If ws.Cells(row, col).value <> vbNullString Then
                    Me.ComboBoxESAnlageTyp.AddItem ws.Cells(row, col).value
                End If
            Next row
            Me.ComboBoxESAnlageTyp.value = "-- Bitte w�hlen --"
        Else
            Me.ComboBoxESAnlageTyp.Clear
            Me.ComboBoxESAnlageTyp.value = "-- Bitte w�hlen --"
        End If
    Case vbNullString
    Case Else
        Me.ComboBoxESAnlageTyp.Clear
        Me.ComboBoxESAnlageTyp.value = "-- Bitte w�hlen --"
    End Select

End Sub

Private Sub ComboBoxPRHGewerk_Change()
    ' Funktionsweise gem. Kommentaren ComboboxEPHGewerk
    Dim row                  As Variant
    Dim col                  As Long
    Dim lastrow              As Long
    Dim ws                   As Worksheet: Set ws = Globals.shPData

    Me.ComboBoxPRHGewerk.BackColor = SystemColorConstants.vbWindowBackground

    If Me.ComboBoxPRHGewerk.value = "-- Bitte w�hlen --" Then
        Me.ComboBoxPRUGewerk.Enabled = False
        Me.ComboBoxPRUGewerk.Clear
        Me.ComboBoxPRUGewerk.value = "-- Bitte w�hlen --"
        Exit Sub
    End If

    Dim HGewerk              As String
    HGewerk = WLookup(Me.ComboBoxPRHGewerk.value, ws.range("PRO_Hauptgewerk"), 2)

    If Not IsError(Application.WorksheetFunction.Match(HGewerk & " PRI", ws.range("10:10"), 0)) Then

1       col = Application.WorksheetFunction.Match(HGewerk & " PRI", ws.range("10:10"), 0)
        lastrow = Application.WorksheetFunction.CountA(ws.Cells(13, col).EntireColumn) + 10

        Me.ComboBoxPRUGewerk.Clear

        Me.ComboBoxPRUGewerk.Enabled = True

        For row = 13 To lastrow
            If ws.Cells(row, col).value <> vbNullString Then
                Me.ComboBoxPRUGewerk.AddItem ws.Cells(row, col).value
            End If
        Next row

    Else
        Me.ComboBoxPRUGewerk.value = "-- Bitte w�hlen --"
    End If

    Me.ComboBoxPRHGewerk.BackColor = SystemColorConstants.vbWindowBackground

End Sub

Private Sub ComboBoxPRUGewerk_Change()

    Me.ComboBoxPRUGewerk.BackColor = SystemColorConstants.vbWindowBackground

End Sub

Private Sub ComboBoxGeb�ude_Change()
    ' Funktionsweise gem. Kommentaren ComboboxEPHGewerk
    Me.ComboBoxGeb�ude.BackColor = SystemColorConstants.vbWindowBackground
    If Me.MultiPageTyp.value = 2 Then Exit Sub
    Dim col                  As Long
    Dim lastrow              As Long
    Dim arr()                As Variant
    Dim tmparr()             As Variant
    Dim rng                  As range
    Dim ws                   As Worksheet
    Set ws = Globals.shGeb�ude

    If Me.ComboBoxGeb�ude.value = "-- Bitte w�hlen --" Then
        Me.ComboBoxGeschoss.Enabled = False
        Me.ComboBoxGeschoss.Clear
        Me.ComboBoxGeschoss.value = "-- Bitte w�hlen --"
        Exit Sub
    End If

    If Not IsError(ws.range("1:1").Find(Me.ComboBoxGeb�ude.value).Column) Then

1       col = ws.range("1:1").Find(Me.ComboBoxGeb�ude.value).Column
        lastrow = ws.Cells(ws.rows.Count, col).End(xlUp).row

        Me.ComboBoxGeschoss.Clear

        Me.ComboBoxGeschoss.Enabled = True
        Set rng = ws.range(Globals.shGeb�ude.Cells(5, col), ws.Cells(lastrow, col + 1))
        arr() = rng.Resize(rng.rows.Count, 1).Offset(1, 0)
        tmparr() = RemoveBlanksFromStringArray(arr())
        Me.ComboBoxGeschoss.List = tmparr()
        Me.ComboBoxGeschoss.value = "-- Bitte w�hlen --"
    Else
        Me.ComboBoxGeschoss.value = "-- Bitte w�hlen --"
    End If

End Sub

Private Sub ComboBoxGeschoss_Change()

    Me.ComboBoxGeschoss.BackColor = SystemColorConstants.vbWindowBackground

End Sub

Private Sub ComboBoxStand_Change()
    Me.ComboBoxStand.BackColor = SystemColorConstants.vbWindowBackground
End Sub

Private Sub ComboBoxGeb�udeTeil_Change()
    Me.ComboBoxGeb�udeTeil.BackColor = SystemColorConstants.vbWindowBackground
End Sub


