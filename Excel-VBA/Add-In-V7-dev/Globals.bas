Attribute VB_Name = "Globals"
Attribute VB_Description = "Beinhaltet Globale Variabeln und Funktionen auf welche von mehreren orten zugriff gew�rt sein muss."

'@Folder "Excel-Items"
'@IgnoreModule VariableNotUsed
'@ModuleDescription "Beinhaltet Globale Variabeln und Funktionen auf welche von mehreren orten zugriff gew�rt sein muss."

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
Public shPZM                 As Worksheet
Public shAnsichten           As Worksheet

Public MySheetHandler        As SheetChangeHandler

Public CopyrightSTR          As String
Private pProjekt             As IProjekt
Private pPlank�pfe           As Collection

Public Sub Unprotect()
    SetWBs
    shPData.Unprotect "Reb$1991"
    shGeb�ude.Unprotect "Reb$1991"
    shProjekt.Unprotect "Reb$1991"
    shIndex.Unprotect "Reb$1991"
    shStoreData.Unprotect "Reb$1991"
    shVersand.Unprotect "Reb$1991"
End Sub

Public Sub Protect()
    SetWBs
    shPData.Protect "Reb$1991"
    shPZM.Protect "Reb$1991"
    shGeb�ude.Protect "Reb$1991"
End Sub

Public Function Projekt(Optional ByVal ForceNew As Boolean = False) As IProjekt
    If shPData Is Nothing Then Globals.SetWBs
    With shPData
        If pProjekt Is Nothing Or ForceNew Then
            Set pProjekt = _
                         ProjektFactory.Create( _
                         .range("ADM_Projektnummer").value, _
                         AdressFactory.Create _
                         (shGeb�ude.range("ADM_ADR_Strasse").value, _
                          shGeb�ude.range("ADM_ADR_PLZ").value, _
                          shGeb�ude.range("ADM_ADR_Ort").value), _
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

Public Function plank�pfe() As Collection
    If pPlank�pfe Is Nothing Then GetPlank�pfe
    Set plank�pfe = pPlank�pfe
End Function

Public Function GetPlank�pfe(Optional ByVal Gewerk As String = vbNullString, Optional ByVal Planart As String = vbNullString) As Collection
    'TODO Create Plank�pfe from Workbook / Database
    Set pPlank�pfe = New Collection
    Dim row                  As range
    Dim ResizeRows           As Long
    Dim rng                  As range
    Set rng = shStoreData.range("A1").CurrentRegion.Offset(2, 0)
    If rng.rows.Count - 3 = 0 Then ResizeRows = 1 Else ResizeRows = rng.rows.Count - 2
    'select what filters matter
    Dim bGewerk              As Boolean: If Gewerk = vbNullString Then bGewerk = False Else bGewerk = True
    Dim bPlanart             As Boolean: If Planart = vbNullString Then bPlanart = False Else Planart = True
    ' if it is a Prinzipschema
    If bPlanart Then
        For Each row In rng.Resize(ResizeRows, 1)
            If Globals.shStoreData.Cells(row.row, 5).value = Planart Then pPlank�pfe.Add PlankopfFactory.LoadFromDataBase(row.row)
        Next
        GoTo Loaded

        ' check if the Gewerk is applicable
    ElseIf bGewerk Then
        For Each row In rng.Resize(ResizeRows, 1)
            If Globals.shStoreData.Cells(row.row, 3).value = Gewerk Then pPlank�pfe.Add PlankopfFactory.LoadFromDataBase(row.row)
        Next
        GoTo Loaded
    Else
        For Each row In rng.Resize(ResizeRows, 1)
            pPlank�pfe.Add PlankopfFactory.LoadFromDataBase(row.row)
        Next
    End If
Loaded:
    Set GetPlank�pfe = pPlank�pfe
    writelog LogInfo, "Loaded " & pPlank�pfe.Count & " Plank�pfe from the Database"
End Function

Public Function SetWBs() As Boolean
    ' Setzt alle Workbooks und Worksheets welche vom Add-In verwendet werden.
    SetWBs = False
    On Error Resume Next
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
    Set shPZM = WB.Sheets("PZM")
    Set shAnsichten = WB.Sheets("Ansichten-Schnitte")
    On Error GoTo 0
    Globals.Projekt
    SetWBs = True
    writelog LogInfo, "Loaded all Workbooks in Globals Module"
End Function


