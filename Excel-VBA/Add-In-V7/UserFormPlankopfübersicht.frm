VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPlankopfübersicht 
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17280
   OleObjectBlob   =   "UserFormPlankopfübersicht.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormPlankopfübersicht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








'@Folder "Plankopf"
Option Explicit
Private icons                As UserFormIconLibrary
Private planköpfe            As New Collection
Private Filters              As Boolean

Private Sub CommandButtonAdd_Click()

    Dim frm                  As New UserFormPlankopf
    frm.setIcons Add
    frm.Show 1
    LoadListViewPlan Me.ListViewPlankopf

    SetFilters

End Sub

Private Sub CommandButtonEdit_Click()

    Dim row                  As Long
    row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.Item(1).text).row
    Dim frm                  As New UserFormPlankopf
    frm.LoadClass PlankopfFactory.LoadFromDataBase(row), Projekt
    frm.setIcons Edit
    frm.Show 1

    LoadListViewPlan Me.ListViewPlankopf

    SetFilters

End Sub

Private Sub CommandButtonFilters_Click()

    ShowFilter

End Sub

Private Sub ShowFilter()

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

Public Property Let Title(Value As String)

    Me.TitleLabel.Caption = Value

End Property

Public Property Let icon(Value As String)

    Set icons = New UserFormIconLibrary
    Dim icon                 As MSForms.control

    For Each icon In icons.Controls
        If icon.Tag = "icon" And icon.Name = Value Then
            Me.TitleIcon.Picture = icon.Picture
        End If
    Next

End Property

Public Property Let Instruction(Value As String)

    Me.LabelInstructions.Caption = Value

End Property

Private Sub CommandButtonCopy_Click()

    Dim row                  As Long
    Dim Plankopf             As IPlankopf
    If Globals.shStoreData.Cells(4, 1).Value = vbNullString Then
        row = 3
    Else
        row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.Item(1).text).row
    End If
    Set Plankopf = PlankopfFactory.LoadFromDataBase(row)
    Dim frm                  As New UserFormPlankopf
    Dim answer               As Boolean
    If IndexFactory.GetIndexes(PlankopfFactory.LoadFromDataBase(row)).Count > 0 Then
        Select Case MsgBox("Vorhandene Indexe kopieren?", vbYesNo, "Indexe kopieren?")
            Case vbYes
                answer = True
            Case vbNo
                answer = False
        End Select
    Else
        answer = False
    End If
    Set frm.PlankopfCopyFrom = Plankopf
    frm.CopyPlankopf Plankopf, Projekt, answer
    frm.Show 1

    LoadListViewPlan Me.ListViewPlankopf

    SetFilters

End Sub

Private Sub CommandButtonDelete_Click()

    Dim row                  As Long
    If Globals.shStoreData.Cells(4, 1).Value = vbNullString Then
        row = 3
    Else
        row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.Item(1).text).row
    End If
    With Globals.shStoreData
        Dim info             As String: info = vbNewLine & .Cells(row, 14).Value & vbNewLine & IndexFactory.GetIndexes(PlankopfFactory.LoadFromDataBase(row)).Count & " Indexe"
    End With
    Select Case MsgBox("Bist du sicher dass du den Plankopf löschen willst?" & info, vbYesNo, "Plankopf löschen")
        Case vbYes
            PlankopfFactory.DeleteFromDatabase row
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

    Dim e                    As ListItem
StartOver:
    For Each e In Me.ListViewPlankopf.ListItems
Debug.Print e.ListSubItems.Item(Index).text
        If FilterValue <> "Alles" Then
            If e.ListSubItems.Item(Index).text <> FilterValue Then
                Me.ListViewPlankopf.ListItems.Remove e.Index
                GoTo StartOver
            End If
        End If
    Next e

End Sub

Private Sub CommandButtonSetFilter_Click()

    FilterListView 3, Me.ComboBoxFilterGeschoss.Value
    FilterListView 4, Me.ComboBoxFilterGebäude.Value
    FilterListView 5, Me.ComboBoxFilterGebäudeteil.Value
    FilterListView 6, Me.ComboBoxFilterGewerk.Value
    FilterListView 7, Me.ComboBoxFilterUnterGewerk.Value
    FilterListView 8, Me.ComboBoxFilterPlanart.Value

End Sub

Private Sub ListViewPlankopf_DblClick()

    Dim row                  As Long
    If Globals.shStoreData.Cells(4, 1).Value = vbNullString Then
        row = 3
    Else
        row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.Item(1).text).row
    End If
    Dim frm                  As New UserFormPlankopf
    frm.LoadClass PlankopfFactory.LoadFromDataBase(row), Projekt
    frm.setIcons Edit
    frm.Show 1

End Sub

Private Sub UserForm_Initialize()

    LoadListViewPlan Me.ListViewPlankopf
    Filters = False
    ShowFilter

    If Me.ListViewPlankopf.ListItems.Count < 1 Then CommandButtonAdd_Click

End Sub

Private Sub SetFilters()

    LoadFilters Me.ComboBoxFilterGebäude, "Gebäude"
    LoadFilters Me.ComboBoxFilterGebäudeteil, "Gebäudeteil"
    LoadFilters Me.ComboBoxFilterGeschoss, "Geschoss"
    LoadFilters Me.ComboBoxFilterGewerk, "Gewerk"
    LoadFilters Me.ComboBoxFilterUnterGewerk, "Untergewerk"
    LoadFilters Me.ComboBoxFilterPlanart, "Planart"

End Sub

Private Sub LoadFilters(ByRef Filter As MSForms.ComboBox, ByVal FilterText As String)

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
            If Not .Exists(e.Value) Then
                .Add e.Value, Nothing
            End If
        Next e

        Filter.List = .Keys
    End With
    Filter.Value = Filter.List(0)

End Sub

