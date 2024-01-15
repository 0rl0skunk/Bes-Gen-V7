Attribute VB_Name = "Globals"
Attribute VB_Description = "Beinhaltet Globale Variabeln und Funktionen auf welche von mehreren orten zugriff gew�rt sein muss."
'@Folder "Excel-Items"
'@IgnoreModule VariableNotUsed
'@ModuleDescription "Beinhaltet Globale Variabeln und Funktionen auf welche von mehreren orten zugriff gew�rt sein muss."
'@Version "Release V1.0.0"

Option Explicit

Public Const Version         As Double = 7#
Public Const maxlen          As Long = 35        'Maximale Anzahl Zeichen der Plan�berschrift im Modul 'Plankopf.cls'
Public Const TinLineProjekte As String = "H:\TinLine\00_Projekte\"
Public Const XMLVorlage      As String = "H:\TinLine\01_Standards\transform.xsl"
Public Const TemplatePagesXslm As String = "H:\TinLine\01_Standards\Beschriftungsgenerator\Bes-Gen-PZM_Templates.xlsm"

Public WB                    As Workbook
Public shPData               As Worksheet
Public shStoreData           As Worksheet
Public shAdress              As Worksheet
Public shVersand             As Worksheet
Public shIndex               As Worksheet
Public shPlanListe           As Worksheet
Public shGeb�ude             As Worksheet
Public shSPSync              As Worksheet
Public shProjekt             As Worksheet
Public xlsmPages             As Workbook
Public CopyrightSTR          As String
Private pProjekt             As IProjekt
Private pPlank�pfe           As Collection

Public Function Projekt(Optional ByVal ForceNew As Boolean = False) As IProjekt
    With Application.ActiveWorkbook.Sheets("Projektdaten")
        If pProjekt Is Nothing Or ForceNew Then
            Set pProjekt = _
                         ProjektFactory.Create( _
                         .range("ADM_Projektnummer").value, _
                         AdressFactory.Create _
                         (.range("ADM_ADR_Strasse").value, _
                          .range("ADM_ADR_PLZ").value, _
                          .range("ADM_ADR_Ort").value), _
                         .range("ADM_Projektbezeichnung").value, _
                         .range("ADM_Projektphase").value, _
                         .range("ADM_ProjektpfadSharePoint").value)
            writelog LogInfo, "Created Projekt " & pProjekt.Projektnummer
        Else
            writelog LogInfo, "Projekt already exists " & pProjekt.Projektnummer
        End If
    End With
    Set Projekt = pProjekt
End Function

Public Function Plank�pfe() As Collection
    If pPlank�pfe Is Nothing Then GetPlank�pfe
    Set Plank�pfe = pPlank�pfe
End Function

Public Function GetPlank�pfe(Optional ByVal Gewerk As String = vbNullString, Optional ByVal Planart As String = vbNullString) As Collection
    'TODO Create Plank�pfe from Workbook / Database
    Set pPlank�pfe = New Collection
    Dim row                  As range
    Dim ResizeRows           As Long
    Dim rng                  As range
    Set rng = shStoreData.range("A1").CurrentRegion.Offset(2, 0)
    If rng.rows.Count - 3 = 0 Then ResizeRows = 1 Else ResizeRows = rng.rows.Count - 3
    'select what filters matter
    Dim bGewerk              As Boolean: If Gewerk = vbNullString Then bGewerk = False Else bGewerk = True
    Dim bPlanart             As Boolean: If Planart = vbNullString Then bPlanart = False Else Planart = True
    ' if it is a Prinzipschema
    If Planart = "Prinzipschema" Then
        For Each row In rng.Resize(ResizeRows, 1)
            If bPlanart And Globals.shStoreData.Cells(row.row, 5).value = Gewerk Then pPlank�pfe.Add PlankopfFactory.LoadFromDataBase(row.row)
        Next
        GoTo Loaded
    End If
    ' check if the Gewerk is applicable
    If bGewerk Then
        For Each row In rng.Resize(ResizeRows, 1)
            If Globals.shStoreData.Cells(row.row, 3).value = Gewerk Then pPlank�pfe.Add PlankopfFactory.LoadFromDataBase(row.row)
        Next
        GoTo Loaded
    End If
Loaded:
    Set GetPlank�pfe = pPlank�pfe
    writelog LogInfo, "Loaded " & pPlank�pfe.Count & " Plank�pfe from the Database"
End Function

Public Function SetWBs() As Boolean
    ' Setzt alle Workbooks und Worksheets welche vom Add-In verwendet werden.
    SetWBs = False
    If WB Is Nothing Then Set WB = Application.ActiveWorkbook
    Dim i                    As Long
    Set shAdress = WB.Sheets("Adressverzeichnis")
    Set shStoreData = WB.Sheets("Datenbank")
    Set shIndex = WB.Sheets("Index")
    Set shPlanListe = WB.Sheets("Planlisten")
    Set shVersand = WB.Sheets("Versand")
    Set shGeb�ude = WB.Sheets("Geb�ude")
    Set shPData = WB.Sheets("Projektdaten")
    Set shSPSync = WB.Sheets("SharePointSync")
    Set shProjekt = WB.Sheets("Projekterstellen")

    Globals.Projekt
    SetWBs = True
    writelog LogInfo, "Loaded all Workbooks in Globals Module"
End Function


