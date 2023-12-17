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
'@Folder "Plankopf"
Option Explicit

Public Enum EnumIcon
    Add = 0
    Edit = 1
End Enum

Private icons                As UserFormIconLibrary
Private pplankopf            As IPlankopf
Public PlankopfCopyFrom      As IPlankopf
Private pProjekt             As IProjekt
Private shPData              As Worksheet
Private shGeb�ude            As Worksheet

Public Sub setIcons(ByVal icon As EnumIcon)

    Select Case icon
        Case 0
            Me.TitleIcon.Picture = icons.IconAddProperties.Picture
            Me.TitleLabel.Caption = "Plankopf erstellen"
        Case 1
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

    Set pplankopf = FormToPlankopf
    Me.TextBoxBeschriftungPlannummer.Value = pplankopf.Plannummer
    Me.TextBoxBeschriftungDateiname.Value = pplankopf.PDFFileName
    Me.TextBoxPlan�berschrift.Value = pplankopf.Plan�berschrift
    
    Me.LabelDWGFileName.Caption = pplankopf.DWGFileName
    Me.LabelXMLFileName.Caption = pplankopf.XMLFileName
    Me.LabelFolderName.Caption = pplankopf.FolderName

End Sub

Private Sub CommandButtonIndexErstellen_Click()

    Dim Index                As IIndex: Set Index = IndexFactory.Create( _
        IDPlan:=pplankopf.ID, _
        GezeichnetPerson:=Me.TextBoxIndexGez.Value, _
        GezeichnetDatum:=Me.TextBoxIndexGezDatum.Value, _
        Klartext:=Me.TextBoxIndexKlartext.Value, _
        Letter:=Me.TextBoxIndexLetter.Value _
                 )
    IndexFactory.AddToDatabase Index
    pplankopf.AddIndex Index

    LoadIndexes

    Me.TextBoxIndexGez.Value = vbNullString
    Me.TextBoxIndexGezDatum.Value = vbNullString
    Me.TextBoxIndexKlartext.Value = vbNullString
    Me.TextBoxIndexLetter.Value = vbNullString

End Sub

Private Sub CommandButtonIndexL�schen_Click()

    Dim ID                   As String
    ID = Me.ListViewIndex.SelectedItem.ListSubItems(1)
    IndexFactory.DeleteFromDatabase ID

    pplankopf.ClearIndex
    IndexFactory.GetIndexes pplankopf

    LoadIndexes

End Sub

Private Sub CommandLayoutW�hlen_Click()

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

Private Sub MultiPageTyp_Change()
' TODO Remove Geschoss "Gesamt" from Plan and Schema Beschriftungen
    Select Case Me.MultiPageTyp.Value
    Case 0 'PLA
        Me.ComboBoxGeb�ude.Enabled = True
        Me.ComboBoxGeb�udeTeil.Enabled = True
        Me.ComboBoxGeschoss.Enabled = True
    Case 1 'SCH
        Me.ComboBoxGeb�ude.Enabled = True
        Me.ComboBoxGeb�udeTeil.Enabled = True
        Me.ComboBoxGeschoss.Enabled = True
    Case 2 'PRI
        Me.ComboBoxGeb�ude.Value = Me.ComboBoxGeb�ude.List(0)
        Me.ComboBoxGeb�udeTeil.Value = Me.ComboBoxGeb�udeTeil.List(0)
        Me.ComboBoxGeschoss.Value = Me.ComboBoxGeschoss.List(0)
        Me.ComboBoxGeb�ude.Enabled = False
        Me.ComboBoxGeb�udeTeil.Enabled = False
        Me.ComboBoxGeschoss.Enabled = False
    End Select
    
End Sub

'@Ignore ProcedureNotUsed
Private Sub Preview_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim frm                  As New UserFormPlankopfPreview
    frm.LoadClass FormToPlankopf, pProjekt
    frm.Show 1

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

    ' Geb�udeTeil
    Me.ComboBoxGeb�udeTeil.Clear
    Me.ComboBoxGeb�udeTeil.List = getList("PRO_Geb�udeteil")
    If Me.ComboBoxGeb�udeTeil.ListCount = 1 Then
        Me.ComboBoxGeb�udeTeil.Value = Me.ComboBoxGeb�udeTeil.List(0)
        Me.ComboBoxGeb�udeTeil.Enabled = False
    Else
        Me.ComboBoxGeb�ude.Enabled = True
    End If
    ' Geb�ude
    Me.ComboBoxGeb�ude.Clear
    Me.ComboBoxGeb�ude.List = getList("PRO_Geb�ude")
    If Me.ComboBoxGeb�ude.ListCount = 1 Then
        Me.ComboBoxGeb�ude.Value = Me.ComboBoxGeb�ude.List(0)
        Me.ComboBoxGeb�ude.Enabled = False
    Else
        Me.ComboBoxGeb�ude.Enabled = True
    End If

    ' Geschoss
    'Me.ComboBoxGeschoss.Clear
    Me.ComboBoxGeschoss.Enabled = Me.ComboBoxGeb�ude.ListCount = 1

    Me.MultiPageTyp.Value = 0
    ' Formate
    Me.ComboBoxLayoutFormat.Clear
    arr() = getList("PLA_Format")
    Me.ComboBoxLayoutFormat.List = arr()

    ' Massstab
    Me.TextBoxLayoutMasstab.Value = "1:50"
    Me.LabelProjektnummer.Caption = Globals.shPData.range("ADM_Projektnummer").Value

    Me.TextBoxPlanInfoDatumGezeichnet.Value = Format(Now, "DD.MM.YYYY")
    Me.TextBoxPlanInfoK�rzelGezeichnet.Value = getUserName

    writelog LogInfo, "UserFormPlankopf > Inizialise complete"

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

        For Each ind In pplankopf.Indexes
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

    Set pplankopf = Plankopf
    Set Plankopf = Nothing
    Dim Planstand            As String
    Dim Plantyp              As Integer
    Dim Gewerk               As String
    Dim UnterGewerk          As String


    Select Case pplankopf.Plantyp
        Case "PLA"
            Me.MultiPageTyp.Value = 0
            Me.ComboBoxEPHGewerk.Value = pplankopf.Gewerk
            Me.ComboBoxEPUGewerk.Value = pplankopf.UnterGewerk
            Me.ComboBoxEPArt.Value = pplankopf.Planart
        Case "SCH"
            Me.MultiPageTyp.Value = 1
            Me.ComboBoxESHGewerk.Value = pplankopf.Gewerk
            Me.ComboBoxESUGewerk.Value = pplankopf.UnterGewerk
        Case "PRI"
            Me.MultiPageTyp.Value = 2
            Me.ComboBoxPRHGewerk.Value = pplankopf.Gewerk
            Me.ComboBoxPRUGewerk.Value = pplankopf.UnterGewerk
    End Select
    Me.ComboBoxGeb�ude.Value = pplankopf.Geb�ude
    Me.ComboBoxGeb�udeTeil.Value = pplankopf.Geb�udeteil
    Me.ComboBoxGeschoss.Value = pplankopf.Geschoss
    Me.ComboBoxLayoutFormat.Value = pplankopf.LayoutGr�sse
    Me.TextBoxLayoutMasstab.Value = pplankopf.LayoutMasstab
    Me.TextBoxPlanInfoDatumGezeichnet.Value = pplankopf.GezeichnetDatum
    Me.TextBoxPlanInfoK�rzelGezeichnet.Value = pplankopf.GezeichnetPerson
    Me.TextBoxPlanInfoDatumGepr�ft.Value = pplankopf.Gepr�ftDatum
    Me.TextBoxPlanInfoK�rzelGepr�ft.Value = pplankopf.Gepr�ftPerson
    Me.TextBoxPlan�berschrift.Value = pplankopf.Plan�berschrift
    Me.LabelDWGFileName.Caption = pplankopf.DWGFileName
    Me.LabelXMLFileName.Caption = pplankopf.XMLFileName
    Me.LabelFolderName.Caption = pplankopf.FolderName
    LoadIndexes

    Me.ComboBoxStand.Value = pplankopf.LayoutPlanstand

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
        Me.BesID.Caption = pplankopf.ID
        Me.TinLineID.Caption = pplankopf.IDTinLine
    Else
        Me.BesID.Caption = getNewID(6, Globals.shStoreData, shStoreData.range("A1").CurrentRegion, 1)
        pplankopf.ID = Me.BesID.Caption
        Dim Index            As IIndex
        For Each Index In pplankopf.Indexes
            Index.PlanID = pplankopf.ID
            IndexFactory.AddToDatabase Index
        Next
    End If

    CommandButtonBeschriftungAktualisieren_Click

End Sub

Public Sub CopyPlankopf(Plankopf As IPlankopf, ByVal Projekt As IProjekt, ByVal CopyIndex As Boolean)

    If CopyIndex Then
        Set Plankopf.Indexes = PlankopfCopyFrom.Indexes
        Set PlankopfCopyFrom = Nothing
    End If
    LoadClass Plankopf, Projekt, True

End Sub

Private Function FormToPlankopf() As IPlankopf

    Dim Plantyp              As String
    Dim Gewerk               As String
    Dim UnterGewerk          As String

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
                         GezeichnetPerson:=Me.TextBoxPlanInfoK�rzelGezeichnet.Value, _
                         GezeichnetDatum:=Me.TextBoxPlanInfoDatumGezeichnet.Value, _
                         Gepr�ftPerson:=Me.TextBoxPlanInfoK�rzelGepr�ft.Value, _
                         Gepr�ftDatum:=Me.TextBoxPlanInfoDatumGepr�ft.Value, _
                         Geb�ude:=Me.ComboBoxGeb�ude.Value, _
                         Geb�udeteil:=Me.ComboBoxGeb�udeTeil.Value, _
                         Gewerk:=Gewerk, _
                         UnterGewerk:=UnterGewerk, _
                         Geschoss:=Me.ComboBoxGeschoss.Value, _
                         Format:=Me.ComboBoxLayoutFormat.Value, _
                         Masstab:=Me.TextBoxLayoutMasstab.Value, _
                         Stand:=Me.ComboBoxStand.Value, _
                         Plantyp:=Plantyp, _
                         Planart:=Me.ComboBoxEPArt.Value, _
                         TinLineID:=Me.TinLineID.Caption, _
                         SkipValidation:=False, _
                         Plan�berschrift:=Me.TextBoxPlan�berschrift.Value, _
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
        Me.ComboBoxEPUGewerk.Value = "-- Bitte w�hlen --"
    End If

End Sub

Private Sub ComboBoxEPHGewerk_Change()

    Dim Row                  As Variant
    Dim col                  As Integer
    Dim lastrow              As Long
Dim WS As Worksheet: Set WS = Globals.shPData


    'If Not Dev Then On Error GoTo ErrMsg

    If Me.ComboBoxEPHGewerk.Value = "-- Bitte w�hlen --" Then
        Me.ComboBoxEPUGewerk.Enabled = False
        Me.ComboBoxEPUGewerk.Clear
        Me.ComboBoxEPUGewerk.Value = "-- Bitte w�hlen --"
        Me.ComboBoxEPArt.Enabled = False
        Me.ComboBoxEPArt.Clear
        Me.ComboBoxEPArt.Value = "-- Bitte w�hlen --"
        Exit Sub
    End If

    If Me.ComboBoxEPHGewerk.Value = "" Then Exit Sub

    Me.ComboBoxEPArt.Enabled = True
    Me.ComboBoxEPUGewerk.Enabled = True

    Me.ComboBoxEPHGewerk.BackColor = SystemColorConstants.vbWindowBackground
    Dim HGewerk              As String
    HGewerk = WLookup(Me.ComboBoxEPHGewerk.Value, WS.range("PRO_Hauptgewerk"), 2)

    If Not IsError(Application.Match(HGewerk & " PLA", WS.range("10:10"), 0)) Then
1       col = Application.Match(HGewerk & " PLA", WS.range("10:10"), 0)
        lastrow = Application.CountA(WS.Cells(13, col).EntireColumn) + 10
        Me.ComboBoxEPUGewerk.Clear
        For Row = 13 To lastrow
            If WS.Cells(Row, col).Value <> "" Then
                Me.ComboBoxEPUGewerk.AddItem WS.Cells(Row, col).Value
            End If
        Next Row
        Me.ComboBoxEPUGewerk.Value = "-- Bitte w�hlen --"
2       col = Application.Match(HGewerk, WS.range("9:9"), 0)
        lastrow = Application.CountA(WS.Cells(13, col).EntireColumn) + 10
        Me.ComboBoxEPArt.Clear
        For Row = 13 To lastrow
            If WS.Cells(Row, col).Value <> "" Then
                Me.ComboBoxEPArt.AddItem WS.Cells(Row, col).Value
            End If
        Next Row
        Me.ComboBoxEPArt.Value = "-- Bitte w�hlen --"
    End If

    Exit Sub

End Sub

Private Sub ComboBoxESAnlageTyp_Change()

    Me.ComboBoxESAnlageTyp.BackColor = SystemColorConstants.vbWindowBackground
    If Me.ComboBoxESAnlageTyp.Value = "Steuerung" Then
        Me.ComboBoxESAnlageTyp.ControlTipText = "Genaue Steuerung im Klartext definieren!"
    Else
        Me.ComboBoxESAnlageTyp.ControlTipText = "W�hle den Anlagentyp des zu beschriftenden Schemas aus."
    End If

End Sub

Private Sub ComboBoxESHGewerk_Change()

    Dim Row                  As Variant
    Dim col                  As Integer
    Dim lastrow              As Long
    Dim WS As Worksheet: Set WS = Globals.shPData

    'If Not Dev Then On Error GoTo ErrMsg

    Dim HGewerk              As String
    HGewerk = WLookup(Me.ComboBoxESHGewerk.Value, WS.range("PRO_Hauptgewerk"), 2)

    Me.ComboBoxESHGewerk.BackColor = SystemColorConstants.vbWindowBackground
    If Me.ComboBoxESHGewerk.Value = "-- Bitte W�hlen --" Then
        Me.ComboBoxESAnlageTyp.Enabled = False
        Me.ComboBoxESUGewerk.Enabled = False
        Me.ComboBoxESAnlageTyp.Clear
        Me.ComboBoxESUGewerk.Clear
        Me.ComboBoxESAnlageTyp.Value = "-- Bitte w�hlen --"
        Me.ComboBoxESUGewerk.Value = "-- Bitte w�hlen --"
        Exit Sub
    End If
1   col = Application.WorksheetFunction.Match(HGewerk & " SCH", WS.range("10:10"), 0) 'get collumn of currently selected Gewerk
2   lastrow = Application.WorksheetFunction.CountA(WS.Cells(13, col).EntireColumn) + 11 'get last row of said collumn
    Me.ComboBoxESUGewerk.Clear
    Me.ComboBoxESAnlageTyp.Enabled = True
    Me.ComboBoxESUGewerk.Enabled = True
    For Row = 13 To lastrow
        If WS.Cells(Row, col).Value <> "" Then
            Me.ComboBoxESUGewerk.AddItem WS.Cells(Row, col).Value
        End If
    Next Row
    Me.ComboBoxESUGewerk.Value = "-- Bitte w�hlen --"

    Me.ComboBoxESHGewerk.BackColor = SystemColorConstants.vbWindowBackground

    Exit Sub

End Sub

Private Sub ComboBoxESUGewerk_Change()

    Dim col                  As Variant
    Dim Row                  As Variant
    Dim lastrow              As Variant
    Dim WS As Worksheet: Set WS = Globals.shPData

    Me.ComboBoxESUGewerk.BackColor = SystemColorConstants.vbWindowBackground

    If Me.ComboBoxESUGewerk.Value = "-- Bitte w�hlen --" Then Exit Sub
    If Me.ComboBoxESUGewerk.Value = "" Then Exit Sub
    Select Case Me.ComboBoxESHGewerk.Value
        Case "Elektro"
            If Not IsError(Application.Match("Anlagetyp " & Me.ComboBoxESUGewerk.Value, WS.range("11:11"), 0)) Then
1               col = Application.Match("Anlagetyp " & Me.ComboBoxESUGewerk.Value, WS.range("11:11"), 0)
                lastrow = Application.WorksheetFunction.CountA(WS.Cells(13, col).EntireColumn) + 11
                Me.ComboBoxESAnlageTyp.Clear
                For Row = 12 To lastrow
                    If WS.Cells(Row, col).Value <> "" Then
                        Me.ComboBoxESAnlageTyp.AddItem WS.Cells(Row, col).Value
                    End If
                Next Row
                Me.ComboBoxESAnlageTyp.Value = "-- Bitte w�hlen --"
            Else
                Me.ComboBoxESAnlageTyp.Clear
                Me.ComboBoxESAnlageTyp.Value = "-- Bitte w�hlen --"
            End If
        Case ""

        Case Else
            'HLKKS
            Me.ComboBoxESAnlageTyp.Clear
            Me.ComboBoxESAnlageTyp.Value = "-- Bitte w�hlen --"
    End Select

End Sub

Private Sub ComboBoxPRHGewerk_Change()

    Dim Row                  As Variant
    Dim col                  As Integer
    Dim lastrow              As Long
    Dim WS As Worksheet: Set WS = Globals.shPData


    If Not Dev Then On Error GoTo ErrMsg

    Me.ComboBoxPRHGewerk.BackColor = SystemColorConstants.vbWindowBackground

    If Me.ComboBoxPRHGewerk.Value = "-- Bitte w�hlen --" Then
        Me.ComboBoxPRUGewerk.Enabled = False
        Me.ComboBoxPRUGewerk.Clear
        Me.ComboBoxPRUGewerk.Value = "-- Bitte w�hlen --"
        Exit Sub
    End If

    Dim HGewerk              As String
    HGewerk = WLookup(Me.ComboBoxPRHGewerk.Value, WS.range("PRO_Hauptgewerk"), 2)

    If Not IsError(Application.WorksheetFunction.Match(HGewerk & " PRI", WS.range("10:10"), 0)) Then
1       col = Application.WorksheetFunction.Match(HGewerk & " PRI", WS.range("10:10"), 0)
        lastrow = Application.WorksheetFunction.CountA(WS.Cells(13, col).EntireColumn) + 10
        Me.ComboBoxPRUGewerk.Clear
        Me.ComboBoxPRUGewerk.Enabled = True
        For Row = 13 To lastrow
            If WS.Cells(Row, col).Value <> "" Then
                Me.ComboBoxPRUGewerk.AddItem WS.Cells(Row, col).Value
            End If
        Next Row
    Else
        Me.ComboBoxPRUGewerk.Value = "-- Bitte w�hlen --"
    End If

    Me.ComboBoxPRHGewerk.BackColor = SystemColorConstants.vbWindowBackground

    Exit Sub

ErrMsg:

End Sub

Private Sub ComboBoxPRUGewerk_Change()

    Me.ComboBoxPRUGewerk.BackColor = SystemColorConstants.vbWindowBackground

End Sub

Private Sub ComboBoxGeb�ude_Change()

    Me.ComboBoxGeb�ude.BackColor = SystemColorConstants.vbWindowBackground
    'get current building column
    Dim col                  As Long
    Dim lastrow              As Long

    Dim arr()                As Variant
    Dim tmparr()             As Variant

    Dim rng                  As range

    'If Not Dev Then On Error GoTo ErrMsg

    If Me.ComboBoxGeb�ude.Value = "-- Bitte w�hlen --" Then
        Me.ComboBoxGeschoss.Enabled = False
        Me.ComboBoxGeschoss.Clear
        Me.ComboBoxGeschoss.Value = "-- Bitte w�hlen --"
        Exit Sub
    End If
    'On Error Resume Next
    If Not IsError(Globals.shGeb�ude.range("1:1").Find(Me.ComboBoxGeb�ude.Value).Column) Then
1       col = Globals.shGeb�ude.range("1:1").Find(Me.ComboBoxGeb�ude.Value).Column
        lastrow = Globals.shGeb�ude.Cells(Globals.shGeb�ude.rows.Count, col).End(xlUp).Row
        Me.ComboBoxGeschoss.Clear
        Me.ComboBoxGeschoss.Enabled = True
        Set rng = Globals.shGeb�ude.range(Globals.shGeb�ude.Cells(5, col), Globals.shGeb�ude.Cells(lastrow, col + 1))
        arr() = rng.Resize(rng.rows.Count, 1)
        tmparr() = RemoveBlanksFromStringArray(arr())
        Me.ComboBoxGeschoss.List = tmparr()
        Me.ComboBoxGeschoss.Value = "-- Bitte w�hlen --"
    Else
        Me.ComboBoxGeschoss.Value = "-- Bitte w�hlen --"
    End If

    If Me.ComboBoxGeb�ude.ListCount = 1 Then
        ' if there is only one listitem in Geb�ude
        If Not IsError(Globals.shGeb�ude.range("1:1").Find(Me.ComboBoxGeb�ude.Value).Column) Then
2           col = Globals.shGeb�ude.range("1:1").Find(Me.ComboBoxGeb�ude.Value).Column
            lastrow = Globals.shGeb�ude.Cells(Globals.shGeb�ude.rows.Count, col).End(xlUp).Row
            Me.ComboBoxGeschoss.Clear
            Me.ComboBoxGeschoss.Enabled = True
            Set rng = Globals.shGeb�ude.range(Globals.shGeb�ude.Cells(5, col), Globals.shGeb�ude.Cells(lastrow, col + 1))
Debug.Print rng.Address
            arr() = rng.Resize(rng.rows.Count, 1)
            tmparr() = RemoveBlanksFromStringArray(arr())
            Me.ComboBoxGeschoss.List = tmparr()
            Me.ComboBoxGeschoss.Value = "-- Bitte w�hlen --"
        Else
            Me.ComboBoxGeschoss.Value = "-- Bitte w�hlen --"
        End If
    End If
    On Error GoTo 0

    Exit Sub

End Sub

Private Sub ComboBoxGeschoss_Change()

    Me.ComboBoxGeschoss.BackColor = SystemColorConstants.vbWindowBackground

End Sub


