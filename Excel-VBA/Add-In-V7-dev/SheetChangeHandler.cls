VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetChangeHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1


Private Sub Class_Initialize()
Set App = Application
End Sub

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
    Dim Geb�udeCode As String
    Dim GeschossCode As String
    Dim isGeb�ude As Boolean
    Dim isAdresse As Boolean
    Dim isGeschoss As Boolean
    Select Case Sh.Name
    Case Globals.shGeb�ude.Name
    ' something changed in shGeb�ude after creating the Project
    Geb�udeCode = Sh.Cells(3, Source.Column).value
    GeschossCode = Sh.Cells(Source.row, 1).value
    
    isGeb�ude = Application.Intersect(Globals.shGeb�ude.range("B1:AQ2"), Source)
    isAdresse = Application.Intersect(Globals.shGeb�ude.range("B4:AQ6"), Source)
    isGeschoss = Application.Intersect(Globals.shGeb�ude.range("B9:AQ98"), Source)
    
    If Globals.shGeb�ude.Cells(1, 4).value = vbNullString Then
    ' only one Building
        If isAdresse Then
        ' change the adresse in tinProject
        End If
        If isGeschoss Then
        ' get like folder
        LikePath = Globals.shGeb�ude.Cells(Source.row, 1).value & "_*"
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
    ' check if it is a Geb�ude or a Geschoss
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
