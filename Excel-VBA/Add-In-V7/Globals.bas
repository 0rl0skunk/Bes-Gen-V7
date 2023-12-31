Attribute VB_Name = "Globals"
Attribute VB_Description = "Beinhaltet Globale Variabeln und Funktionen auf welche von mehreren orten zugriff gew�rt sein muss."
'@IgnoreModule VariableNotUsed
Option Explicit
'@Folder "Excel-Items"
'@ModuleDescription "Beinhaltet Globale Variabeln und Funktionen auf welche von mehreren orten zugriff gew�rt sein muss."

Global Const Version As Double = 5#

Global Const maxlen As Long = 35                 'Maximale Anzahl Zeichen der Plan�berschrift im Modul 'Plankopf.cls'
Global Const TinLineProjekte As String = "H:\TinLine\00_Projekte\"
Global Const XMLVorlage As String = "H:\TinLine\01_Standards\transform.xsl"
Global Const TemplatePagesXslm As String = "H:\TinLine\01_Standards\Beschriftungsgenerator\Bes-Gen-PZM_Templates.xlsm"
Global Const LogDepth As Double = 0
' 3 = Trace
' 2 = Info
' 1 = Warnings
' 0 = Errors


Public WB                    As Workbook
Public shPData               As Worksheet
Public shStoreData           As Worksheet
Public shAdress              As Worksheet
Public shVersand             As Worksheet
Public shIndex               As Worksheet
Public shPlanListe           As Worksheet
Public shGeb�ude             As Worksheet
Public shSPSync              As Worksheet
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

Private Sub GetPlank�pfe()

    'TODO Create Plank�pfe from Workbook / Database
    Dim row                  As range
    For Each row In shStoreData.range("A1").CurrentRegion
        pPlank�pfe.Add PlankopfFactory.LoadFromDataBase(row.row)
    Next
    writelog LogInfo, "Loaded " & pPlank�pfe.Count & " Plank�pfe from the Database"
End Sub

Function Initialize() As Boolean
    Initialize = False
    Set WB = ActiveWorkbook

    CopyrightSTR = _
                 "Release: " & Version & vbLf _
               & ChrW(&HA9) & Format(Now(), "yyyy") & " Orlando Bassi"

    Set shPData = WB.Sheets("Projektdaten")

    On Error GoTo 0
    ' Version checking
    If shPData.range("B4").value < Version Then
        Dim curVersion       As String
        Dim shouldVersion    As String

        curVersion = "Bes-Gen-PZM-Add-In-V" & Version
        Select Case shPData.range("B4").value
            Case vbNullString
                shouldVersion = "Bes-Gen-PZM-Add-In"
            Case Else
                If shPData.range("B4").value > Version Then
                    shouldVersion = "Bes-Gen-PZM-Add-In-V" & shPData.range("B4").value & " oder neuer"
                End If
        End Select
        MsgBox "Die Arbeitsmappe wird von dieser Version vom Beschriftungsgenerator nicht unterst�zt!" & vbLf _
             & "Bitte Update die Arbeitsmappe oder lade eine �ltere Version vom Beschriftungsgenerator." & vbLf _
             & vbLf & "aktuelle Version" & vbLf & curVersion & vbLf _
             & "zu verwendende Version: " & vbLf & shouldVersion _
               , vbCritical, "Nicht unterst�tzte Arbeitsmappe"
        Exit Function
    End If

    On Error Resume Next
    WB.Activate

    SetWBs

    xlsmPages.Close False
    Set xlsmPages = Nothing
    On Error GoTo 0
    Initialize = True
    shPData.Activate

End Function

Public Function SetWBs() As Boolean
    ' Setzt alle Workbooks und Worksheets welche vom Add-In verwendet werden.
    SetWBs = False
    If WB Is Nothing Then Set WB = Application.ActiveWorkbook
    Dim i                    As Integer
    Set shAdress = WB.Sheets("Adressverzeichnis")
    If err Then
        Set xlsmPages = Workbooks.Open(TemplatePagesXslm)
        xlsmPages.Sheets("Adressverzeichnis").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shAdress = WB.Sheets("Adressverzeichnis")
    End If
    Set shStoreData = WB.Sheets("Datenbank")
    If err Then
        Set xlsmPages = Workbooks.Open(TemplatePagesXslm)
        xlsmPages.Sheets("Datenbank").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shStoreData = WB.Sheets("Datenbank")
    End If
    Set shIndex = WB.Sheets("Index")
    If err Then
        Set xlsmPages = Workbooks.Open(TemplatePagesXslm)
        xlsmPages.Sheets("Index").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shIndex = WB.Sheets("Index")
    End If
    Set shPlanListe = WB.Sheets("Planlisten")
    If err Then
        Set xlsmPages = Workbooks.Open(TemplatePagesXslm)
        xlsmPages.Sheets("Planlisten").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shPlanListe = WB.Sheets("Planlisten")
    End If
    Set shVersand = WB.Sheets("Versand")
    If err Then
        Set xlsmPages = Workbooks.Open(TemplatePagesXslm)
        xlsmPages.Sheets("Versand").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shVersand = WB.Sheets("Versand")
    End If
    Set shGeb�ude = WB.Sheets("Geb�ude")
    If err Then
        Set xlsmPages = Workbooks.Open(TemplatePagesXslm)
        xlsmPages.Sheets("Geb�ude").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shGeb�ude = WB.Sheets("Geb�ude")
    End If
    Set shPData = WB.Sheets("Projektdaten")
    If err Then
        Set xlsmPages = Workbooks.Open(TemplatePagesXslm)
        xlsmPages.Sheets("Projektdaten").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shPData = WB.Sheets("Projektdaten")
    End If
    Set shSPSync = WB.Sheets("SharePointSync")
    If err Then
        Set xlsmPages = Workbooks.Open(TemplatePagesXslm)
        xlsmPages.Sheets("SharePointSync").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shPData = WB.Sheets("SharePointSync")
    End If


    Globals.Projekt
    SetWBs = True
    writelog LogInfo, "Loaded all Workbooks in Globals Module"

End Function


