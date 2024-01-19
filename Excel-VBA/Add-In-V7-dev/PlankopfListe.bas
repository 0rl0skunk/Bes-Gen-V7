Attribute VB_Name = "PlankopfListe"

'@Folder("Plankopf")
'@ModuleDescription "Alle Plankopflisten werden über dieses Modul geladen, damit alle gleich aussehen."

Option Explicit

Public Sub LoadListViewPlan(ByRef control As ListView)

    Dim pla                  As IPlankopf
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
                .Add , , vbNullString, 20        ' 0
            Else
                .Add , , vbNullString, 0         '0
            End If
            .Add , , "ID", 0                     ' 1
            .Add , , "Plannummer", 70            ' 2
            .Add , , "Geschoss", 70              ' 3
            .Add , , "Gebäude", 70               ' 4
            .Add , , "Gebäudeteil", 70           ' 5
            .Add , , "Gewerk", 70                ' 6
            .Add , , "Untergewerk", 70           ' 7
            .Add , , "Planart", 70               ' 8
            .Add , , "Gezeichnet", 70            ' 9
            .Add , , "Geprüft", 70               ' 10
            .Add , , "Index", 20                 ' 11
        End With
        If Globals.shStoreData Is Nothing Then Globals.SetWBs
        lastrow = Globals.shStoreData.range("A1").CurrentRegion.rows.Count
        For row = 3 To lastrow
            Application.StatusBar = "Lädt Plankopf " & row - 2 & " von " & lastrow - 2
            Set pla = PlankopfFactory.LoadFromDataBase(row)
            'Planköpfe.Add Pla                    ', Pla.ID
            Set li = .ListItems.Add()
            li.ListSubItems.Add , , pla.ID
            li.ListSubItems.Add , , pla.Plannummer
            li.ListSubItems.Add , , pla.Geschoss
            li.ListSubItems.Add , , pla.Gebäude
            li.ListSubItems.Add , , pla.Gebäudeteil
            li.ListSubItems.Add , , pla.Gewerk
            li.ListSubItems.Add , , pla.UnterGewerk
            li.ListSubItems.Add , , pla.Planart
            li.ListSubItems.Add , , pla.Gezeichnet
            li.ListSubItems.Add , , pla.Geprüft
            li.ListSubItems.Add , , pla.CurrentIndex.Index
        Next row
        .Refresh
    End With
    
    Application.StatusBar = False

End Sub

