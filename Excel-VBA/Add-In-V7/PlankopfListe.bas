Attribute VB_Name = "PlankopfListe"

'@Folder("Plankopf")
'@Version "Release V1.0.0"

Option Explicit

Public Sub LoadListViewPlan(ByRef control As ListView)

    Dim Pla                  As IPlankopf
    Dim li                   As ListItem
    Dim row                  As Long
    Dim lastrow              As Long

    With control
        .ListItems.Clear
        .View = lvwReport
        .Gridlines = False
        .FullRowSelect = True
        .FlatScrollBar = False
        With .ColumnHeaders
            .Clear
            If control.CheckBoxES Then
            .Add , , vbNullString, 20            ' 0
            Else
            .Add , , vbNullString, 0    '0
            End If
            .Add , , "ID", 0                     ' 1
            .Add , , "Plannummer", 70               ' 2
            .Add , , "Geschoss", 70                 ' 3
            .Add , , "Geb�ude", 70                  ' 4
            .Add , , "Geb�udeteil", 70              ' 5
            .Add , , "Gewerk", 70                 ' 6
            .Add , , "Untergewerk", 70            ' 7
            .Add , , "Planart", 70                ' 8
            .Add , , "Gezeichnet", 70               ' 9
            .Add , , "Gepr�ft", 70                  ' 10
            .Add , , "Index", 20                    ' 11
        End With
        If Globals.shStoreData Is Nothing Then Globals.SetWBs
        lastrow = Globals.shStoreData.range("A1").CurrentRegion.rows.Count
        For row = 3 To lastrow
            Application.StatusBar = "L�dt Plankopf " & row - 2 & " von " & lastrow - 2
            Set Pla = PlankopfFactory.LoadFromDataBase(row)
            'Plank�pfe.Add Pla                    ', Pla.ID
            Set li = .ListItems.Add()
            li.ListSubItems.Add , , Pla.ID
            li.ListSubItems.Add , , Pla.Plannummer
            li.ListSubItems.Add , , Pla.Geschoss
            li.ListSubItems.Add , , Pla.Geb�ude
            li.ListSubItems.Add , , Pla.Geb�udeteil
            li.ListSubItems.Add , , Pla.Gewerk
            li.ListSubItems.Add , , Pla.UnterGewerk
            li.ListSubItems.Add , , Pla.Planart
            li.ListSubItems.Add , , Pla.Gezeichnet
            li.ListSubItems.Add , , Pla.Gepr�ft
            li.ListSubItems.Add , , Pla.CurrentIndex.Index
        Next row
        .Refresh
    End With
    
    Application.StatusBar = False

End Sub

