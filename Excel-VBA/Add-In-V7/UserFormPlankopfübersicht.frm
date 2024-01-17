VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPlankopfübersicht 
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17280
   OleObjectBlob   =   "UserFormPlankopfübersicht.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormPlankopfübersicht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Übersicht aller erstellten Planköpfe im Projekt. Die Planköpfe können hier drüber erstellt, angepasst und kopiert werden."

'@Folder "Plankopf"
'@ModuleDescription "Übersicht aller erstellten Planköpfe im Projekt. Die Planköpfe können hier drüber erstellt, angepasst und kopiert werden."
'@Version "Release V1.0.0"

Option Explicit

Private icons                As UserFormIconLibrary
Private Planköpfe            As New Collection
Private Filters              As Boolean

Private Sub CommandButtonAdd_Click()
    ' neuer Plankopf erstellen
    Dim frm                  As New UserFormPlankopf
    frm.setIcons Add
    frm.Show 1

    ' nach dem schliessen die Planliste aktualisieren
    LoadListViewPlan Me.ListViewPlankopf

    SetFilters

End Sub

Private Sub CommandButtonEdit_Click()
    ' ausgewählten Plankopf bearbeiten
    Dim row                  As Long
    row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.Item(1).Text).row ' die zu ladende Reihe aus der Datenbank finden
    Dim frm                  As New UserFormPlankopf
    frm.LoadClass PlankopfFactory.LoadFromDataBase(row), Projekt ' ein Plankopf-Objekt aus der Reihe erstellen und im UserForm laden
    frm.setIcons Edit
    frm.Show 1

    ' nach dem schliessen die Planliste aktualisieren
    LoadListViewPlan Me.ListViewPlankopf

    SetFilters

End Sub

Private Sub CommandButtonFilters_Click()

    ShowFilter

End Sub

Private Sub ShowFilter()
    ' Filter anzeigen oder verstecken
    If Filters Then
        Me.CommandButtonFilters.Caption = "< Filter"
        Me.CommandButtonFilters.Left = 708
        Me.CommandButtonClose.Left = 786
        Me.width = 876
    Else
        Me.CommandButtonFilters.Caption = "Filter >"
        Me.CommandButtonFilters.Left = 396
        Me.CommandButtonClose.Left = 588
        Me.width = 678
    End If

    Filters = Not Filters

End Sub

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

Public Property Let Title(ByVal value As String)
    ' titel setzen
    Me.TitleLabel.Caption = value

End Property

Public Property Let icon(ByVal value As String)
    ' Icons setzen
    Set icons = New UserFormIconLibrary
    Dim icon                 As MSForms.control

    For Each icon In icons.Controls
        If icon.Tag = "icon" And icon.Name = value Then
            Me.TitleIcon.Picture = icon.Picture
        End If
    Next

End Property

Public Property Let Instruction(ByVal value As String)
    ' Kurzbeschrieb setzen
    Me.LabelInstructions.Caption = value

End Property

Private Sub CommandButtonCopy_Click()

    Dim row                  As Long
    Dim Plankopf             As IPlankopf

    If Globals.shStoreData.Cells(4, 1).value = vbNullString Then
        ' wenn nur ein Plankopf erstellt wurde funktioniert die match funktion nicht richtig deshalb hier eine spezielle condtion.
        row = 3
    Else
        ' ansonsten die Reihe mit der Match Funktion finden
        row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.Item(1).Text).row
    End If

    Set Plankopf = PlankopfFactory.LoadFromDataBase(row)
    Dim frm                  As New UserFormPlankopf
    Dim answer               As Boolean

    If IndexFactory.GetIndexes(PlankopfFactory.LoadFromDataBase(row)).Count > 0 Then
        ' Wenn der zu kopierende Plan Indexe hat soll man diese nach bedarf kopieren können oder nicht.
        Select Case MsgBox("Vorhandene Indexe kopieren?", vbYesNo, "Indexe kopieren?")
        Case vbYes
            answer = True
        Case vbNo
            answer = False
            Plankopf.ClearIndex
        End Select
    Else
        answer = False
        Plankopf.ClearIndex
    End If

    Set frm.PlankopfCopyFrom = Plankopf
    frm.CopyPlankopf Plankopf, Projekt, answer
    frm.Show 1

    ' nach dem schliessen die Planliste aktualisieren
    LoadListViewPlan Me.ListViewPlankopf

    SetFilters

End Sub

Private Sub CommandButtonDelete_Click()
    ' Plankopf löschen
    Dim row                  As Long
    If Globals.shStoreData.Cells(4, 1).value = vbNullString Then
        ' wenn nur ein Plankopf erstellt wurde funktioniert die match funktion nicht richtig deshalb hier eine spezielle condtion.
        row = 3
    Else
        ' ansonsten die Reihe mit der Match Funktion finden
        row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.Item(1).Text).row
    End If

    With Globals.shStoreData
        Dim info             As String
        info = vbNewLine & .Cells(row, 14).value & vbNewLine & IndexFactory.GetIndexes(PlankopfFactory.LoadFromDataBase(row)).Count & " Indexe"
    End With

    Select Case MsgBox("Bist du sicher dass du den Plankopf löschen willst?" & info, vbYesNo, "Plankopf löschen")
        ' zusätzliche bestätigung um versehentliche löschungen zu vermeiden.
    Case vbYes
        PlankopfFactory.DeleteFromDatabase row

        ' nach dem schliessen die Planliste aktualisieren
        LoadListViewPlan Me.ListViewPlankopf

        SetFilters

    Case vbNo
        Exit Sub
    End Select

End Sub

Private Sub CommandButtonFilterReset_Click()

    LoadListViewPlan Me.ListViewPlankopf

End Sub

Private Sub FilterListView(ByVal Index As String, ByVal FilterValue As String)
    ' Filter sind immernoch Work In Progress (WIP)
    Dim e                    As ListItem
StartOver:
    For Each e In Me.ListViewPlankopf.ListItems
        Debug.Print e.ListSubItems.Item(Index).Text
        If FilterValue <> "Alles" Then
            If e.ListSubItems.Item(Index).Text <> FilterValue Then
                Me.ListViewPlankopf.ListItems.Remove e.Index
                GoTo StartOver
            End If
        End If
    Next e

End Sub

Private Sub CommandButtonSetFilter_Click()

    FilterListView 3, Me.ComboBoxFilterGeschoss.value
    FilterListView 4, Me.ComboBoxFilterGebäude.value
    FilterListView 5, Me.ComboBoxFilterGebäudeteil.value
    FilterListView 6, Me.ComboBoxFilterGewerk.value
    FilterListView 7, Me.ComboBoxFilterUnterGewerk.value
    FilterListView 8, Me.ComboBoxFilterPlanart.value

End Sub

Private Sub ListViewPlankopf_DblClick()

    Dim row                  As Long
    If Globals.shStoreData.Cells(4, 1).value = vbNullString Then
        row = 3
    Else
        row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.Item(1).Text).row
    End If
    Dim frm                  As New UserFormPlankopf
    frm.LoadClass PlankopfFactory.LoadFromDataBase(row), Projekt
    frm.setIcons Edit
    frm.Show 1

End Sub

Private Sub UserForm_Initialize()

    Application.Cursor = xlWait
    LoadListViewPlan Me.ListViewPlankopf
    Filters = False
    ShowFilter

    Me.TitleLabel.Caption = "Plankopfübersicht"
    Me.LabelInstructions.Caption = "Planköpfe erstellen, bearbeiten und löschen"

    If Me.ListViewPlankopf.ListItems.Count < 1 Then CommandButtonAdd_Click
    Application.Cursor = xlDefault

End Sub

Private Sub SetFilters()

    LoadFilters Me.ComboBoxFilterGebäude, "Gebäude"
    LoadFilters Me.ComboBoxFilterGebäudeteil, "Gebäudeteil"
    LoadFilters Me.ComboBoxFilterGeschoss, "Geschoss"
    LoadFilters Me.ComboBoxFilterGewerk, "Gewerk"
    LoadFilters Me.ComboBoxFilterUnterGewerk, "Untergewerk"
    LoadFilters Me.ComboBoxFilterPlanart, "Planart"

End Sub

Private Sub LoadFilters(ByVal Filter As MSForms.ComboBox, ByVal FilterText As String)

    Dim e                    As range
    Dim col                  As Long
    Dim lastrow              As Long: lastrow = Globals.shStoreData.range("A1").CurrentRegion.rows.Count
    Dim ws                   As Worksheet: Set ws = Globals.shStoreData

    Select Case FilterText
    Case "Gebäude"
        col = 7
    Case "Geschoss"
        col = 9
    Case "Gebäudeteil"
        col = 8
    Case "Gewerk"
        col = 3
    Case "Untergewerk"
        col = 4
    Case "Planart"
        col = 5
    End Select

    Filter.Clear
    With CreateObject("Scripting.Dictionary")
        .Add "Alles", Nothing
        For Each e In ws.range(ws.Cells(3, col), ws.Cells(lastrow, col))
            If Not .Exists(e.value) Then
                .Add e.value, Nothing
            End If
        Next e

        Filter.List = .Keys
    End With
    Filter.value = Filter.List(0)

End Sub

