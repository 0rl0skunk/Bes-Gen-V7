## Program Flows
### Plankopfübersicht
onAction Button
    Globals.SetWBs
        globals.Projekt
            ProjektFactory.Create
                Adressfactory.create
                    Adress.FillData
                Projekt.Filldata
            write Log
        write log
    UserFormPlankopfÜbersicht.Initialize
        LoadListView
            For each Plankopf
                Plankopffactory.loadFromDatabase
                    Plankopf.Filldata
                        ProjektFactory.Create
                            Adressfactory.create
                                Adress.FillData
                            Projekt.Filldata
                        write Log
                        getNewID
                    write Log
                Indexfactory.getindexes
                    for each Row in shIndex
                        check if the PlanID match
                write log
            add Indexes to the Plankopf
### IndexErstellen Plankopf
Indexfactory.create
    Index.FillData
        NewID
        Write Log
Indexfactory.AddToDatabase
plankopf.addIndex
LoadIndexes
    