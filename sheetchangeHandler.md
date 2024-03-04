Handle change events in the addin
check if a new folder must be created or not.

```vb
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Source As range)
    Debug.Print "Changed " & Sh.Name
    On Error GoTo Finish
    Dim fso As New FileSystemObject
    Dim Folder As Object
    App.EnableEvents = False ' disable infinity Loop
    
    writelog LogWarning, Join(Array("Changed", Source.Address, "in", Sh.Name), " ")
        
    If Not Globals.shPData.range("ADM_ProjektPfadCAD").value = vbNullString Then
    Dim LikePath As String
    Dim CADPath As String: CADPath = Globals.Projekt.ProjektOrdnerCAD
    Dim GebäudeCode As String
    Dim GeschossCode As String
    Dim isGebäude As Boolean
    Dim isAdresse As Boolean
    Dim isGeschoss As Boolean
    Select Case Sh.Name
    Case Globals.shGebäude.Name
    ' something changed in shGebäude after creating the Project
    GebäudeCode = Sh.Cells(3, Source.Column).value
    GeschossCode = Sh.Cells(Source.row, 1).value
    
    isGebäude = Application.Intersect(Globals.shGebäude.range("B1:AQ2"), Source)
    isAdresse = Application.Intersect(Globals.shGebäude.range("B4:AQ6"), Source)
    isGeschoss = Application.Intersect(Globals.shGebäude.range("B9:AQ98"), Source)
    
    If Globals.shGebäude.Cells(1, 4).value = vbNullString Then
    ' only one Building
        If isAdresse Then
        ' change the adresse in tinProject
        End If
        If isGeschoss Then
        ' get like folder
        LikePath = Globals.shGebäude.Cells(Source.row, 1).value & "_*"
        End If
    Else
    ' more buildings
        
    End If
    
    If Globals.shProjekt.range("A1").value = False Then ' EP
        For Each Folder In fso.GetFolder(CADFolder & "\01_EP").SubFolders
            If Folder.Name Like likefolder Then
            ' if the foldernames are simmilar
            magbox "Soll der Ordner " & Folder.Name & " in " & likefolder & " umbenannt werden?"
            End If
        Next Folder
    End If
    If Globals.shProjekt.range("A4").value = False Then ' TF
    
    End If
    If Globals.shProjekt.range("A5").value = False Then ' BS
    
    End If
    
    ' check if a Folder with the same code exists. if yes rename it?
    ' if no create a new one?
    ' check if it is a Gebäude or a Geschoss
    Case Globals.shAnsichten.Name
    ' something changed in shAnsichten after creating the Project
    
    If Globals.shProjekt.range("A6").value = False Then ' DE
    
    End If
    ' check if it is a new one or an old one
    ' rename or create it?
    End Select
    End If
Finish:
    App.EnableEvents = True
End Sub
```