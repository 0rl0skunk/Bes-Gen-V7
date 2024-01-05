Attribute VB_Name = "PlankopfListe"
'@Folder("Plankopf")
Option Explicit

                                
Public Sub LoadListViewPlan(ByRef control As ListView)
    
    Dim Pla                  As IPlankopf
    Dim li                   As ListItem
    
    Dim row                  As Long
    Dim lastrow              As Long
    
    
    With control
        .ListItems.Clear
        .View = lvwReport
        .CheckBoxES = True
        .Gridlines = True
        .FullRowSelect = True
        With .ColumnHeaders
            .Clear
            .Add , , vbNullString, 20                                     ' 0
            .Add , , "ID", 0                                              ' 1
            .Add , , "Plannummer"                                         ' 2
            .Add , , "Geschoss"                                           ' 3
            .Add , , "Gebäude"                                            ' 4
            .Add , , "Gebäudeteil"                                        ' 5
            .Add , , "Gewerk", 0                                          ' 6
            .Add , , "Untergewerk", 0                                     ' 7
            .Add , , "Planart", 0                                         ' 8
            .Add , , "Gezeichnet"                                         ' 9
            .Add , , "Geprüft"                                            ' 10
            .Add , , "Index"                                              ' 11
        End With
        If Globals.shStoreData Is Nothing Then Globals.SetWBs
        lastrow = Globals.shStoreData.range("A1").CurrentRegion.rows.Count
        For row = 3 To lastrow
            Set Pla = PlankopfFactory.LoadFromDataBase(row)
            'Planköpfe.Add Pla                    ', Pla.ID
            Set li = .ListItems.Add()
            li.ListSubItems.Add , , Pla.ID
            li.ListSubItems.Add , , Pla.Plannummer
            li.ListSubItems.Add , , Pla.Geschoss
            li.ListSubItems.Add , , Pla.Gebäude
            li.ListSubItems.Add , , Pla.Gebäudeteil
            li.ListSubItems.Add , , Pla.Gewerk
            li.ListSubItems.Add , , Pla.UnterGewerk
            li.ListSubItems.Add , , Pla.Planart
            li.ListSubItems.Add , , Pla.Gezeichnet
            li.ListSubItems.Add , , Pla.Geprüft
            li.ListSubItems.Add , , Pla.CurrentIndex.Index
        Next row
    End With
    
End Sub

                                
