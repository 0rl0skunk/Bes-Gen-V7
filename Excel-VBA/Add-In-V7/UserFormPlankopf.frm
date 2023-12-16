VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPlankopf 
   ClientHeight    =   12240
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

Option Explicit
'@IgnoreModule IntegerDataType, EmptyStringLiteral
'@Folder "Plankopf"

Private icons                As UserFormIconLibrary
Private pPlankopf            As IPlankopf
Public PlankopfCopyFrom      As IPlankopf
Private pProjekt             As IProjekt
Private shPData              As Worksheet
Private shGebäude            As Worksheet

Public Sub setIcons(ByVal icon As String)

    Select Case icon
        Case "add"
            Me.TitleIcon.Picture = icons.IconAddProperties.Picture
            Me.TitleLabel.Caption = "Plankopf erstellen"
        Case "delete"
            Me.TitleIcon.Picture = icons.IconImportantProperty.Picture
            Me.TitleLabel.Caption = "Plankopf löschen"
        Case "edit"
            Me.TitleIcon.Picture = icons.IconEditProperties.Picture
            Me.TitleLabel.Caption = "Plankopf bearbeiten"
    End Select

End Sub

Private Sub CommandButtonCreate_Click()

    If Me.CommandButtonCreate.Caption = "Update" Then
        PlankopfFactory.ReplaceInDatabase FormToPlankopf
    Else
        PlankopfFactory.AddToDatabase FormToPlankopf
    End If
    Unload Me

End Sub

Private Sub CommandButtonBeschriftungAktualisieren_Click()

    Set pPlankopf = FormToPlankopf
    Me.TextBoxBeschriftungPlannummer.Value = pPlankopf.Plannummer
    Me.TextBoxBeschriftungDateiname.Value = pPlankopf.PDFFileName
    Me.TextBoxPlanüberschrift.Value = pPlankopf.Planüberschrift

End Sub

Private Sub CommandButtonIndexErstellen_Click()

    Dim Index                As IIndex: Set Index = IndexFactory.Create( _
        IDPlan:=pPlankopf.ID, _
        GezeichnetPerson:=Me.TextBoxIndexGez.Value, _
        GezeichnetDatum:=Me.TextBoxIndexGezDatum.Value, _
        Klartext:=Me.TextBoxIndexKlartext.Value, _
        Letter:=Me.TextBoxIndexLetter.Value _
                 )
    IndexFactory.AddToDatabase Index
    pPlankopf.AddIndex Index

    LoadIndexes

    Me.TextBoxIndexGez.Value = vbNullString
    Me.TextBoxIndexGezDatum.Value = vbNullString
    Me.TextBoxIndexKlartext.Value = vbNullString
    Me.TextBoxIndexLetter.Value = vbNullString

End Sub

Private Sub CommandButtonIndexLöschen_Click()

    Dim ID                   As String
    ID = Me.ListViewIndex.SelectedItem.ListSubItems(1)
    IndexFactory.DeleteFromDatabase ID

    pPlankopf.ClearIndex
    IndexFactory.GetIndexes pPlankopf

    LoadIndexes

End Sub

Private Sub CommandLayoutWählen_Click()

    Dim frm                  As New UserFormLayout
    frm.load Me.ComboBoxLayoutFormat.Value, Me.TextBoxLayoutMasstab.Value, Me.MultiPageTyp.Value
    frm.Show 1
    If frm.CheckBoxLoad Then
        Me.ComboBoxLayoutFormat.Value = frm.TextBoxFormatH.Value & "H" & frm.TextBoxFormatB.Value & "B"
    End If
    Set frm = Nothing

End Sub

'@Ignore ProcedureNotUsed
Private Sub EditDWG_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    TinLine.setTinProject pProjekt.ProjektOrdnerCAD
    Select Case Me.MultiPageTyp.Value
        Case 0                                   'Plan
            TinLine.setTinPlanBibliothek
        Case 1                                   'Prinzip
            TinLine.setTinPrinzipBibiothek
    End Select

    CreateObject("Shell.Application").Open (FormToPlankopf.DWGFile)

End Sub

'@Ignore ProcedureNotUsed
Private Sub Preview_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim frm                  As New UserFormPlankopfPreview
    frm.LoadClass FormToPlankopf, pProjekt
    frm.Show 1

End Sub

Private Sub UserForm_Initialize()

    Set icons = New UserFormIconLibrary

    ' ComboBox Listen aufüllen

    ' Unterprojekt
    ' Array mit Unterprojekt Name und Nummer nebeneinander

    Dim arr()

    ' populate unterprojekt if there is only one
    'arr() = getList("Unterprojekte")
    Me.ComboBoxUnterprojekt.List = getList("PRO_Unterprojekte")
    If Me.ComboBoxUnterprojekt.ListCount = 1 Then
        Me.ComboBoxUnterprojekt.Value = Me.ComboBoxUnterprojekt.List(0)
        Me.ComboBoxUnterprojekt.Enabled = False
    End If
    Me.LabelProjektphase.Caption = Globals.shPData.range("ADM_Projektphase").Value

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

    ' GebäudeTeil
    Me.ComboBoxGebäudeTeil.Clear
    Me.ComboBoxGebäudeTeil.List = getList("PRO_Gebäudeteil")
    If Me.ComboBoxGebäudeTeil.ListCount = 1 Then
        Me.ComboBoxGebäudeTeil.Value = Me.ComboBoxGebäudeTeil.List(0)
        Me.ComboBoxGebäudeTeil.Enabled = False
    Else
        Me.ComboBoxGebäude.Enabled = True
    End If
    ' Gebäude
    Me.ComboBoxGebäude.Clear
    Me.ComboBoxGebäude.List = getList("PRO_Gebäude")
    If Me.ComboBoxGebäude.ListCount = 1 Then
        Me.ComboBoxGebäude.Value = Me.ComboBoxGebäude.List(0)
        Me.ComboBoxGebäude.Enabled = False
    Else
        Me.ComboBoxGebäude.Enabled = True
    End If

    ' Geschoss
    'Me.ComboBoxGeschoss.Clear
    If Me.ComboBoxGebäude.ListCount = 1 Then
        Me.ComboBoxGeschoss.Enabled = True
    Else
        Me.ComboBoxGeschoss.Enabled = False
    End If

    Me.MultiPageTyp.Value = 0
    ' Formate
    Me.ComboBoxLayoutFormat.Clear
    arr() = getList("PLA_Format")
    Me.ComboBoxLayoutFormat.List = arr()

    ' Massstab
    Me.TextBoxLayoutMasstab.Value = "1:50"
    Me.LabelProjektnummer.Caption = Globals.shPData.range("ADM_Projektnummer").Value

    Me.TextBoxPlanInfoDatumGezeichnet.Value = Format(Now, "DD.MM.YYYY")
    Me.TextBoxPlanInfoKürzelGezeichnet.Value = getUserName

    writelog "Info", "UserFormPlankopf > Inizialise complete"

End Sub

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

Private Sub LoadIndexes()

    Dim ind                  As IIndex
    Dim li                   As ListItem

    With Me.ListViewIndex
        .ListItems.Clear
        .View = lvwReport
        .CheckBoxes = True
        .Gridlines = True
        .FullRowSelect = True
        With .ColumnHeaders
            .Clear
            .Add , , "", 20
            .Add , , "", 0
            .Add , , "Index", 20
            .Add , , "Gezeichnet", 40
            .Add , , "Datum", 60
            .Add , , "Beschreibung", 250
        End With

        For Each ind In pPlankopf.indexes
            Set li = .ListItems.Add()
            li.ListSubItems.Add , , ind.IndexID
            li.ListSubItems.Add , , ind.Index
            li.ListSubItems.Add , , Split(ind.Gezeichnet, " ; ")(0)
            li.ListSubItems.Add , , Split(ind.Gezeichnet, " ; ")(1)
            li.ListSubItems.Add , , ind.Klartext
        Next
    End With

End Sub

Public Sub LoadClass(Plankopf As IPlankopf, ByVal Projekt As IProjekt, Optional copy As Boolean = False)

    Set pProjekt = Projekt

    Set pPlankopf = Plankopf
    Set Plankopf = Nothing
    Dim Planstand            As String, _
    Plantyp                  As Integer, _
    Gewerk                   As String, _
    UnterGewerk              As String

    Select Case pPlankopf.Plantyp
        Case "PLA"
            Me.MultiPageTyp.Value = 0
            Me.ComboBoxEPHGewerk.Value = pPlankopf.Gewerk
            Me.ComboBoxEPUGewerk.Value = pPlankopf.UnterGewerk
            Me.ComboBoxEPArt.Value = pPlankopf.Planart
        Case "SCH"
            Me.MultiPageTyp.Value = 1
            Me.ComboBoxESHGewerk.Value = pPlankopf.Gewerk
            Me.ComboBoxESUGewerk.Value = pPlankopf.UnterGewerk
        Case "PRI"
            Me.MultiPageTyp.Value = 2
            Me.ComboBoxPRHGewerk.Value = pPlankopf.Gewerk
            Me.ComboBoxPRUGewerk.Value = pPlankopf.UnterGewerk
    End Select
    Me.ComboBoxGebäude.Value = pPlankopf.Gebäude
    Me.ComboBoxGebäudeTeil.Value = pPlankopf.GebäudeTeil
    Me.ComboBoxGeschoss.Value = pPlankopf.Geschoss
    Me.ComboBoxLayoutFormat.Value = pPlankopf.LayoutGrösse
    Me.TextBoxKlartext.Value = pPlankopf.Klartext
    Me.TextBoxLayoutMasstab.Value = pPlankopf.LayoutMasstab
    Me.TextBoxPlanInfoDatumGezeichnet.Value = pPlankopf.GezeichnetDatum
    Me.TextBoxPlanInfoKürzelGezeichnet.Value = pPlankopf.GezeichnetPerson
    Me.TextBoxPlanInfoDatumGeprüft.Value = pPlankopf.GeprüftDatum
    Me.TextBoxPlanInfoKürzelGeprüft.Value = pPlankopf.GeprüftPerson
    Me.TextBoxPlanüberschrift.Value = pPlankopf.Planüberschrift
    LoadIndexes

    Me.ComboBoxStand.Value = pPlankopf.LayoutPlanstand

    If Not copy Then
        ' disable all inputs which should only be set once
        Me.MultiPageTyp.Enabled = False
        Me.ComboBoxEPArt.Enabled = False
        Me.ComboBoxEPHGewerk.Enabled = False
        Me.ComboBoxEPUGewerk.Enabled = False
        Me.ComboBoxESAnlageTyp.Enabled = False
        Me.ComboBoxESHGewerk.Enabled = False
        Me.ComboBoxESUGewerk.Enabled = False
        Me.ComboBoxGebäude.Enabled = False
        Me.ComboBoxGebäudeTeil.Enabled = False
        Me.ComboBoxGeschoss.Enabled = False
        Me.ComboBoxPRHGewerk.Enabled = False
        Me.ComboBoxPRUGewerk.Enabled = False

        Me.CommandButtonCreate.Caption = "Update"
        Me.BesID.Caption = pPlankopf.ID
        Me.TinLineID.Caption = pPlankopf.IDTinLine
    Else
        Me.BesID.Caption = getNewID(6, Globals.shStoreData, shStoreData.range("A1").CurrentRegion, 1)
        pPlankopf.ID = Me.BesID.Caption
        Dim Index            As IIndex
        For Each Index In pPlankopf.indexes
            Index.PlanID = pPlankopf.ID
            IndexFactory.AddToDatabase Index
        Next
    End If

    CommandButtonBeschriftungAktualisieren_Click

End Sub

Public Sub CopyPlankopf(Plankopf As IPlankopf, ByVal Projekt As IProjekt, ByVal CopyIndex As Boolean)

    If CopyIndex Then
        Set Plankopf.indexes = PlankopfCopyFrom.indexes
        Set PlankopfCopyFrom = Nothing
    End If
    LoadClass Plankopf, Projekt, True

End Sub

Private Function FormToPlankopf() As IPlankopf

    Dim Plantyp              As String, _
    Gewerk                   As String, _
    UnterGewerk              As String
    Select Case Me.MultiPageTyp.Value
        Case 0
            Plantyp = "PLA"
            Gewerk = Me.ComboBoxEPHGewerk.Value
            UnterGewerk = Me.ComboBoxEPUGewerk.Value
        Case 1
            Plantyp = "SCH"
            Gewerk = Me.ComboBoxESHGewerk.Value
            UnterGewerk = Me.ComboBoxESUGewerk.Value
        Case 2
            Plantyp = "PRI"
            Gewerk = Me.ComboBoxPRHGewerk.Value
            UnterGewerk = Me.ComboBoxPRUGewerk.Value
        Case Else
            Plantyp = "PLA"
            Gewerk = Me.ComboBoxEPHGewerk.Value
            UnterGewerk = Me.ComboBoxEPUGewerk.Value
    End Select
    If pProjekt Is Nothing Then Set pProjekt = Globals.Projekt
    Set FormToPlankopf = PlankopfFactory.Create( _
                         Projekt:=pProjekt, _
                         GezeichnetPerson:=Me.TextBoxPlanInfoKürzelGezeichnet.Value, _
                         GezeichnetDatum:=Me.TextBoxPlanInfoDatumGezeichnet.Value, _
                         GeprüftPerson:=Me.TextBoxPlanInfoKürzelGeprüft.Value, _
                         GeprüftDatum:=Me.TextBoxPlanInfoDatumGeprüft.Value, _
                         Gebäude:=Me.ComboBoxGebäude.Value, _
                         GebäudeTeil:=Me.ComboBoxGebäudeTeil.Value, _
                         Gewerk:=Gewerk, _
                         UnterGewerk:=UnterGewerk, _
                         Geschoss:=Me.ComboBoxGeschoss.Value, _
                         Format:=Me.ComboBoxLayoutFormat.Value, _
                         Masstab:=Me.TextBoxLayoutMasstab.Value, _
                         Stand:=Me.ComboBoxStand.Value, _
                         Klartext:=Me.TextBoxKlartext, _
                         Plantyp:=Plantyp, _
                         Planart:=Me.ComboBoxEPArt.Value, _
                         TinLineID:=Me.TinLineID.Caption, _
                         SkipValidation:=False, _
                         Planüberschrift:=Me.TextBoxPlanüberschrift.Value, _
                         ID:=Me.BesID.Caption _
                              )

End Function

'-------------------------------------------------------- ComboBox_Change Events ---------------------------------------------------------

Private Sub ComboBoxEPArt_Change()

    Me.ComboBoxEPArt.BackColor = SystemColorConstants.vbWindowBackground

End Sub

Private Sub ComboBoxEPUGewerk_Change()

    Me.ComboBoxEPUGewerk.BackColor = SystemColorConstants.vbWindowBackground

    If Me.ComboBoxEPUGewerk.Value = "" Then
        Me.ComboBoxEPUGewerk.Value = "-- Bitte wählen --"
    End If

End Sub

Private Sub ComboBoxEPHGewerk_Change()

    Dim row, col             As Integer, lastrow As Long

    'If Not Dev Then On Error GoTo ErrMsg

    If Me.ComboBoxEPHGewerk.Value = "-- Bitte wählen --" Then
        Me.ComboBoxEPUGewerk.Enabled = False
        Me.ComboBoxEPUGewerk.Clear
        Me.ComboBoxEPUGewerk.Value = "-- Bitte wählen --"
        Me.ComboBoxEPArt.Enabled = False
        Me.ComboBoxEPArt.Clear
        Me.ComboBoxEPArt.Value = "-- Bitte wählen --"
        Exit Sub
    End If

    If Me.ComboBoxEPHGewerk.Value = "" Then Exit Sub

    Me.ComboBoxEPArt.Enabled = True
    Me.ComboBoxEPUGewerk.Enabled = True

    Me.ComboBoxEPHGewerk.BackColor = SystemColorConstants.vbWindowBackground
    Dim HGewerk              As String
    HGewerk = WLookup(Me.ComboBoxEPHGewerk.Value, Globals.shPData.range("PRO_Hauptgewerk"), 2)

    If Not IsError(Application.Match(HGewerk & " PLA", range("10:10"), 0)) Then
1       col = Application.Match(HGewerk & " PLA", Globals.shPData.range("10:10"), 0)
        lastrow = Application.CountA(Globals.shPData.Cells(12, col).EntireColumn) + 9
        Me.ComboBoxEPUGewerk.Clear
        For row = 12 To lastrow
            If Globals.shPData.Cells(row, col).Value <> "" Then
                Me.ComboBoxEPUGewerk.AddItem Globals.shPData.Cells(row, col).Value
            End If
        Next row
        Me.ComboBoxEPUGewerk.Value = "-- Bitte wählen --"
2       col = Application.Match(HGewerk, Globals.shPData.range("9:9"), 0)
        lastrow = Application.CountA(Globals.shPData.Cells(12, col).EntireColumn) + 9
        Me.ComboBoxEPArt.Clear
        For row = 12 To lastrow
            If Globals.shPData.Cells(row, col).Value <> "" Then
                Me.ComboBoxEPArt.AddItem Globals.shPData.Cells(row, col).Value
            End If
        Next row
        Me.ComboBoxEPArt.Value = "-- Bitte wählen --"
    End If

    Exit Sub

ErrMsg:

End Sub

Private Sub ComboBoxESAnlageTyp_Change()

    Me.ComboBoxESAnlageTyp.BackColor = SystemColorConstants.vbWindowBackground
    If Me.ComboBoxESAnlageTyp.Value = "Steuerung" Then
        Me.ComboBoxESAnlageTyp.ControlTipText = "Genaue Steuerung im Klartext definieren!"
    Else
        Me.ComboBoxESAnlageTyp.ControlTipText = "Wähle den Anlagentyp des zu beschriftenden Schemas aus."
    End If

End Sub

Private Sub ComboBoxESHGewerk_Change()

    Dim row, col             As Integer, lastrow As Long

    'If Not Dev Then On Error GoTo ErrMsg

    Dim HGewerk              As String
    HGewerk = WLookup(Me.ComboBoxESHGewerk.Value, Globals.shPData.range("PRO_Hauptgewerk"), 2)

    Me.ComboBoxESHGewerk.BackColor = SystemColorConstants.vbWindowBackground
    If Me.ComboBoxESHGewerk.Value = "-- Bitte Wählen --" Then
        Me.ComboBoxESAnlageTyp.Enabled = False
        Me.ComboBoxESUGewerk.Enabled = False
        Me.ComboBoxESAnlageTyp.Clear
        Me.ComboBoxESUGewerk.Clear
        Me.ComboBoxESAnlageTyp.Value = "-- Bitte wählen --"
        Me.ComboBoxESUGewerk.Value = "-- Bitte wählen --"
        Exit Sub
    End If
1   col = Application.WorksheetFunction.Match(HGewerk & " SCH", Globals.shPData.range("10:10"), 0) 'get collumn of currently selected Gewerk
2   lastrow = Application.WorksheetFunction.CountA(Globals.shPData.Cells(12, col).EntireColumn) + 9 'get last row of said collumn
    Me.ComboBoxESUGewerk.Clear
    Me.ComboBoxESAnlageTyp.Enabled = True
    Me.ComboBoxESUGewerk.Enabled = True
    For row = 12 To lastrow
        If Globals.shPData.Cells(row, col).Value <> "" Then
            Me.ComboBoxESUGewerk.AddItem Globals.shPData.Cells(row, col).Value
        End If
    Next row
    Me.ComboBoxESUGewerk.Value = "-- Bitte wählen --"

    Me.ComboBoxESHGewerk.BackColor = SystemColorConstants.vbWindowBackground

    Exit Sub

ErrMsg:

End Sub

Private Sub ComboBoxESUGewerk_Change()

    Dim col, row, lastrow

    Me.ComboBoxESUGewerk.BackColor = SystemColorConstants.vbWindowBackground

    If Me.ComboBoxESUGewerk.Value = "-- Bitte wählen --" Then Exit Sub
    If Me.ComboBoxESUGewerk.Value = "" Then Exit Sub
    Select Case Me.ComboBoxESHGewerk.Value
        Case "Elektro"
            If Not IsError(Application.Match("Anlagetyp " & Me.ComboBoxESUGewerk.Value, Globals.shPData.range("11:11"), 0)) Then
1               col = Application.Match("Anlagetyp " & Me.ComboBoxESUGewerk.Value, Globals.shPData.range("11:11"), 0)
                lastrow = Application.WorksheetFunction.CountA(Globals.shPData.Cells(12, col).EntireColumn) + 10
                Me.ComboBoxESAnlageTyp.Clear
                For row = 12 To lastrow
                    If Globals.shPData.Cells(row, col).Value <> "" Then
                        Me.ComboBoxESAnlageTyp.AddItem Globals.shPData.Cells(row, col).Value
                    End If
                Next row
                Me.ComboBoxESAnlageTyp.Value = "-- Bitte wählen --"
            Else
                Me.ComboBoxESAnlageTyp.Clear
                Me.ComboBoxESAnlageTyp.Value = "-- Bitte wählen --"
            End If
        Case ""

        Case Else
            'HLKKS
            Me.ComboBoxESAnlageTyp.Clear
            Me.ComboBoxESAnlageTyp.Value = "-- Bitte wählen --"
    End Select

End Sub

Private Sub ComboBoxPRHGewerk_Change()

    Dim row, col             As Integer, lastrow As Long

    If Not Dev Then On Error GoTo ErrMsg

    Me.ComboBoxPRHGewerk.BackColor = SystemColorConstants.vbWindowBackground

    If Me.ComboBoxPRHGewerk.Value = "-- Bitte wählen --" Then
        Me.ComboBoxPRUGewerk.Enabled = False
        Me.ComboBoxPRUGewerk.Clear
        Me.ComboBoxPRUGewerk.Value = "-- Bitte wählen --"
        Exit Sub
    End If

    Dim HGewerk              As String
    HGewerk = WLookup(Me.ComboBoxPRHGewerk.Value, shPData.range("PRO_Hauptgewerk"), 2)

    If Not IsError(Application.WorksheetFunction.Match(HGewerk & " PRI", shPData.range("10:10"), 0)) Then
1       col = Application.WorksheetFunction.Match(HGewerk & " PRI", shPData.range("10:10"), 0)
        lastrow = Application.WorksheetFunction.CountA(shPData.Cells(12, col).EntireColumn) + 9
        Me.ComboBoxPRUGewerk.Clear
        Me.ComboBoxPRUGewerk.Enabled = True
        For row = 12 To lastrow
            If shPData.Cells(row, col).Value <> "" Then
                Me.ComboBoxPRUGewerk.AddItem shPData.Cells(row, col).Value
            End If
        Next row
    Else
        Me.ComboBoxPRUGewerk.Value = "-- Bitte wählen --"
    End If

    Me.ComboBoxPRHGewerk.BackColor = SystemColorConstants.vbWindowBackground

    Exit Sub

ErrMsg:

End Sub

Private Sub ComboBoxPRUGewerk_Change()

    Me.ComboBoxPRUGewerk.BackColor = SystemColorConstants.vbWindowBackground

End Sub

Private Sub ComboBoxGebäude_Change()

    Me.ComboBoxGebäude.BackColor = SystemColorConstants.vbWindowBackground
    'get current building column
    Dim col                  As Long, lastrow As Long
    Dim arr(), tmparr()
    Dim rng                  As range

    'If Not Dev Then On Error GoTo ErrMsg

    If Me.ComboBoxGebäude.Value = "-- Bitte wählen --" Then
        Me.ComboBoxGeschoss.Enabled = False
        Me.ComboBoxGeschoss.Clear
        Me.ComboBoxGeschoss.Value = "-- Bitte wählen --"
        Exit Sub
    End If
    'On Error Resume Next
    If Not IsError(Globals.shGebäude.range("1:1").Find(Me.ComboBoxGebäude.Value).Column) Then
1       col = Globals.shGebäude.range("1:1").Find(Me.ComboBoxGebäude.Value).Column
        lastrow = Globals.shGebäude.Cells(Globals.shGebäude.rows.Count, col).End(xlUp).row
        Me.ComboBoxGeschoss.Clear
        Me.ComboBoxGeschoss.Enabled = True
        Set rng = Globals.shGebäude.range(Globals.shGebäude.Cells(5, col), Globals.shGebäude.Cells(lastrow, col + 1))
        arr() = rng.Resize(rng.rows.Count, 1)
        tmparr() = RemoveBlanksFromStringArray(arr())
        Me.ComboBoxGeschoss.List = tmparr()
        Me.ComboBoxGeschoss.Value = "-- Bitte wählen --"
    Else
        Me.ComboBoxGeschoss.Value = "-- Bitte wählen --"
    End If

    If Me.ComboBoxGebäude.ListCount = 1 Then
        ' if there is only one listitem in Gebäude
        If Not IsError(Globals.shGebäude.range("1:1").Find(Me.ComboBoxGebäude.Value).Column) Then
2           col = Globals.shGebäude.range("1:1").Find(Me.ComboBoxGebäude.Value).Column
            lastrow = Globals.shGebäude.Cells(Globals.shGebäude.rows.Count, col).End(xlUp).row
            Me.ComboBoxGeschoss.Clear
            Me.ComboBoxGeschoss.Enabled = True
            Set rng = Globals.shGebäude.range(Globals.shGebäude.Cells(5, col), Globals.shGebäude.Cells(lastrow, col + 1))
Debug.Print rng.Address
            arr() = rng.Resize(rng.rows.Count, 1)
            tmparr() = RemoveBlanksFromStringArray(arr())
            Me.ComboBoxGeschoss.List = tmparr()
            Me.ComboBoxGeschoss.Value = "-- Bitte wählen --"
        Else
            Me.ComboBoxGeschoss.Value = "-- Bitte wählen --"
        End If
    End If
    On Error GoTo 0

    Exit Sub

ErrMsg:


End Sub

Private Sub ComboBoxGeschoss_Change()

    Me.ComboBoxGeschoss.BackColor = SystemColorConstants.vbWindowBackground

End Sub


