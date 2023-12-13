VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPlankopfÜbersicht 
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9585.001
   OleObjectBlob   =   "UserFormPlankopfÜbersicht.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormPlankopfÜbersicht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
'@Folder "Plankopf"
Private icons                As UserFormIconLibrary
Private Planköpfe            As New Collection

Private Sub CommandButton1_Click()

    Dim frm                  As New UserFormPlankopf
    frm.Show 1
    LoadListView

End Sub

Private Sub CommandButton2_Click()

    Dim row As Long
    row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.item(1).Text).row
    Dim frm As New UserFormPlankopf
    frm.LoadClass PlankopfFactory.LoadFromDatabase(row), Projekt
    frm.Show 1

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

Private Sub ListViewPlankopf_DblClick()

    Dim row As Long
    row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.item(1).Text).row
    Dim frm As New UserFormPlankopf
    frm.LoadClass PlankopfFactory.LoadFromDatabase(row), Projekt
    frm.Show 1

End Sub

Private Sub UserForm_Activate()

    LoadListView

End Sub

Private Sub UserForm_Initialize()

    Globals.SetWBs
    LoadListView

End Sub

Private Sub LoadListView()

    Dim Pla                  As IPlankopf
    Dim Li                   As ListItem

    Dim row                  As Long, _
    lastrow                  As Long
    With Me.ListViewPlankopf
        .ListItems.Clear
        .View = lvwReport
        .CheckBoxes = True
        .Gridlines = True
        .FullRowSelect = True
        With .ColumnHeaders
            .Clear
            .Add , , "", 20
            .Add , , "ID", 0
            .Add , , "Plannummer"
            .Add , , "Geschoss"
            .Add , , "Gebäude"
            .Add , , "Gebäudeteil"
            .Add , , "Gezeichnet"
            .Add , , "Geprüft"
            .Add , , "Index"
        End With
        If Globals.shStoreData Is Nothing Then Globals.SetWBs
        lastrow = Globals.shStoreData.range("A1").CurrentRegion.rows.Count
        For row = 3 To lastrow
            Set Pla = PlankopfFactory.LoadFromDatabase(row)
            'Planköpfe.Add Pla                    ', Pla.ID
            Set Li = .ListItems.Add()
            Li.ListSubItems.Add , , Pla.ID
            Li.ListSubItems.Add , , Pla.PlanNummer
            Li.ListSubItems.Add , , Pla.Geschoss
            Li.ListSubItems.Add , , Pla.Gebäude
            Li.ListSubItems.Add , , Pla.GebäudeTeil
            Li.ListSubItems.Add , , Pla.Gezeichnet
            Li.ListSubItems.Add , , Pla.Geprüft
            'Li.ListSubItems.Add , , Pla.currentIndex.index
        Next row
    End With
End Sub


