VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPlankopf�bersicht 
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17280
   OleObjectBlob   =   "UserFormPlankopf�bersicht.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormPlankopf�bersicht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@IgnoreModule VariableNotUsed


Option Explicit
'@Folder "Plankopf"
Private icons                As UserFormIconLibrary
Private Plank�pfe            As New Collection
Private Filters              As Boolean

Private Sub CommandButton1_Click()

    Dim frm                  As New UserFormPlankopf
    frm.setIcons Add
    frm.Show 1
    LoadListView

End Sub

Private Sub CommandButton2_Click()

    Dim row                  As Long
    row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.Item(1).Text).row
    Dim frm                  As New UserFormPlankopf
    frm.LoadClass PlankopfFactory.LoadFromDataBase(row), Projekt
    frm.setIcons Edit
    frm.Show 1

    LoadListView

End Sub

Private Sub CommandButton3_Click()

    ShowFilter

End Sub

Private Sub ShowFilter()

    If Filters Then
        Me.CommandButton3.Caption = "< Filter"
        Me.CommandButton3.Left = 708
        Me.CommandButtonClose.Left = 786
        Me.width = 876
    Else
        Me.CommandButton3.Caption = "Filter >"
        Me.CommandButton3.Left = 396
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
        row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.Item(1).Text).row
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

    LoadListView

End Sub

Private Sub CommandButtonDelete_Click()

    Dim row                  As Long
    If Globals.shStoreData.Cells(4, 1).Value = vbNullString Then
        row = 3
    Else
        row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.Item(1).Text).row
    End If
    With Globals.shStoreData
        Dim info             As String: info = vbNewLine & .Cells(row, 14).Value & vbNewLine & IndexFactory.GetIndexes(PlankopfFactory.LoadFromDataBase(row)).Count & " Indexe"
    End With
    Select Case MsgBox("Bist du sicher dass du den Plankopf l�schen willst?" & info, vbYesNo, "Plankopf l�schen")
        Case vbYes
            PlankopfFactory.DeleteFromDatabase row
            LoadListView
        Case vbNo
            Exit Sub
    End Select

End Sub

Private Sub CommandButtonFilterReset_Click()

    LoadListView

End Sub

Private Sub FilterListView(ByVal Index As String, ByVal FilterValue As String)

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

    FilterListView 3, Me.ComboBoxFilterGeschoss.Value
    FilterListView 4, Me.ComboBoxFilterGeb�ude.Value
    FilterListView 5, Me.ComboBoxFilterGeb�udeteil.Value
    FilterListView 6, Me.ComboBoxFilterGewerk.Value
    FilterListView 7, Me.ComboBoxFilterUnterGewerk.Value
    FilterListView 8, Me.ComboBoxFilterPlanart.Value

End Sub

Private Sub ListViewPlankopf_DblClick()

    Dim row                  As Long
    If Globals.shStoreData.Cells(4, 1).Value = vbNullString Then
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

    LoadListView
    Filters = False
    ShowFilter

End Sub

Private Sub LoadListView()

    Dim Pla                  As IPlankopf
    Dim li                   As ListItem

    Dim row                  As Long
    Dim lastrow              As Long


    With Me.ListViewPlankopf
        .ListItems.Clear
        .View = lvwReport
        .CheckBoxes = True
        .Gridlines = True
        .FullRowSelect = True
        With .ColumnHeaders
            .Clear
            .Add , , vbNullString, 20            ' 0
            .Add , , "ID", 0                     ' 1
            .Add , , "Plannummer"                ' 2
            .Add , , "Geschoss"                  ' 3
            .Add , , "Geb�ude"                   ' 4
            .Add , , "Geb�udeteil"               ' 5
            .Add , , "Gewerk", 0                 ' 6
            .Add , , "Untergewerk", 0            ' 7
            .Add , , "Planart", 0                ' 8
            .Add , , "Gezeichnet"                ' 9
            .Add , , "Gepr�ft"                   ' 10
            .Add , , "Index"                     ' 11
        End With
        If Globals.shStoreData Is Nothing Then Globals.SetWBs
        lastrow = Globals.shStoreData.range("A1").CurrentRegion.rows.Count
        For row = 3 To lastrow
            Set Pla = PlankopfFactory.LoadFromDataBase(row)
            'Plank�pfe.Add Pla                    ', Pla.ID
            Set li = .ListItems.Add()
            li.ListSubItems.Add , , Pla.ID
            li.ListSubItems.Add , , Pla.Plannummer
            li.ListSubItems.Add , , Pla.Geschoss
            li.ListSubItems.Add , , Pla.Geb�ude
            li.ListSubItems.Add , , Pla.Geb�udeTeil
            li.ListSubItems.Add , , Pla.Gewerk
            li.ListSubItems.Add , , Pla.UnterGewerk
            li.ListSubItems.Add , , Pla.Planart
            li.ListSubItems.Add , , Pla.Gezeichnet
            li.ListSubItems.Add , , Pla.Gepr�ft
            li.ListSubItems.Add , , Pla.currentIndex.Index
        Next row
    End With

    LoadFilters Me.ComboBoxFilterGeb�ude, "Geb�ude"
    LoadFilters Me.ComboBoxFilterGeb�udeteil, "Geb�udeteil"
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
        Case "Geb�ude"
            col = 7
        Case "Geschoss"
            col = 9
        Case "Geb�udeteil"
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

