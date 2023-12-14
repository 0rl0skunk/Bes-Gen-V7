VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPlankopfübersicht 
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13320
   OleObjectBlob   =   "UserFormPlankopfübersicht.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormPlankopfübersicht"
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

    Dim row                  As Long
    row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.item(1).Text).row
    Dim frm                  As New UserFormPlankopf
    frm.LoadClass PlankopfFactory.LoadFromDatabase(row), Projekt
    frm.Show 1

    LoadListView

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
        row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.item(1).Text).row
    End If
    Set Plankopf = PlankopfFactory.LoadFromDatabase(row)
    Dim frm                  As New UserFormPlankopf
    Dim answer               As Boolean
    If IndexFactory.GetIndexes(PlankopfFactory.LoadFromDatabase(row)).Count > 0 Then
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
        row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.item(1).Text).row
    End If
    With Globals.shStoreData
        Dim info             As String: info = vbNewLine & .Cells(row, 14).Value & vbNewLine & IndexFactory.GetIndexes(PlankopfFactory.LoadFromDatabase(row)).Count & " Indexe"
    End With
    Select Case MsgBox("Bist du sicher dass du den Plankopf löschen willst?" & info, vbYesNo, "Plankopf löschen")
        Case vbYes
            PlankopfFactory.DeleteFromDatabase row
            LoadListView
        Case vbNo
            Exit Sub
    End Select

End Sub

Private Sub ListViewPlankopf_DblClick()

    Dim row                  As Long
    If Globals.shStoreData.Cells(4, 1).Value = vbNullString Then
        row = 3
    Else
        row = Globals.shStoreData.range("A:A").Find(Me.ListViewPlankopf.SelectedItem.ListSubItems.item(1).Text).row
    End If
    Dim frm                  As New UserFormPlankopf
    frm.LoadClass PlankopfFactory.LoadFromDatabase(row), Projekt
    frm.Show 1

End Sub

Private Sub UserForm_Initialize()

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
            .Add , , "Gebüude"
            .Add , , "Gebüudeteil"
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
            Li.ListSubItems.Add , , Pla.Plannummer
            Li.ListSubItems.Add , , Pla.Geschoss
            Li.ListSubItems.Add , , Pla.Gebäude
            Li.ListSubItems.Add , , Pla.GebäudeTeil
            Li.ListSubItems.Add , , Pla.Gezeichnet
            Li.ListSubItems.Add , , Pla.Geprüft
            'Li.ListSubItems.Add , , Pla.currentIndex.index
        Next row
    End With
End Sub


