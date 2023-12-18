Attribute VB_Name = "ModuleTest"
Attribute VB_Description = "Modul mit Test funktionen."
'@IgnoreModule EmptyStringLiteral
'@Folder "Objektdaten"
'@ModuleDescription "Modul mit Test funktionen."

Option Explicit

Public Sub testProjekt()

    Dim proj                 As IProjekt

    Set proj = ProjektFactory.Create( _
               "220033", _
               AdressFactory.Create( _
               "Ebenaustrasse 10", _
               "6048", _
               "Horw" _
               ), _
               "Umbau Kaserne Auenfeld, Frauenfeld", _
               "Ausführung", _
               "https://rebsamennet.sharepoint.com/:f:/r/sites/PZM-ZH/03_Pub/00_Projekte/Auft.2022/220033_LU?csf=1&web=1&e=HHSqga" _
               )

Debug.Print proj.ProjektBezeichnung

End Sub

Public Sub test2()

    Dim ttask                As New Task
    ttask.Filldata _
        ErfasstAm:="24.11.2023", _
        ErfasstVon:="BaOr", _
        FälligAm:="25.011.2023", _
        Gewerk:="Elektro", _
        Gebäude:="Haus1", _
        Gebäudeteil:="West", _
        Geschoss:="Erdgeschoss", _
        Erledigt:=False, _
        Priorität:=2, _
        text:="Lorem ipsum dolor sit amet"

    Dim frm                  As New UserFormTaskDetail
    frm.LoadClass ttask
    frm.Show 1


End Sub

Sub test3()

    Application.ActiveWorkbook.Names.Add "BesGenVersion"
    Application.ActiveWorkbook.Names("BesGenVersion").value = "V7"

End Sub


