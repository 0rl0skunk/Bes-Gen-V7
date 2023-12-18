Attribute VB_Name = "Globals"
Attribute VB_Description = "Beinhaltet Globale Variabeln und Funktionen auf welche von mehreren orten zugriff gewärt sein muss."
'@IgnoreModule VariableNotUsed
Option Explicit
'@Folder "Excel-Items"
'@ModuleDescription "Beinhaltet Globale Variabeln und Funktionen auf welche von mehreren orten zugriff gewärt sein muss."

Global Const Version As Double = 5#

Global Const maxlen As Long = 35                 'Maximale Anzahl Zeichen der Planüberschrift im Modul 'Plankopf.cls'
Global Const OrdnerVorlage As String = "H:\TinLine\01_Standards\00_Vorlageordner" 'TODO Create Folder from Excel
Global Const VorlageEPDWG As String = "H:\TinLine\01_Standards\EP-Vorlage.dwg" 'TODO Create Folder from Excel
Global Const VorlageEPDWGGEB As String = "H:\TinLine\01_Standards\EP-Vorlage_GEB.dwg" 'TODO Create Folder from Excel
Global Const VorlagePRDWG As String = "H:\TinLine\01_Standards\PR-Vorlage.dwg" 'TODO Create Folder from Excel
Global Const TinLineProjekte As String = "H:\TinLine\00_Projekte\"
Global Const XMLVorlage As String = "H:\TinLine\01_Standards\transform.xsl"
Global Const TemplatePagesXslm As String = "H:\TinLine\01_Standards\Beschriftungsgenerator\Bes-Gen-PZM_Templates.xlsm"
Global Const LogDepth As Double = 1#
' 3= everything > Slowest
' 2= warnings and errors
' 1= Errors only


Public WB                    As Workbook
Public shPData               As Worksheet
Public shStoreData           As Worksheet
Public shAdress              As Worksheet
Public shVersand             As Worksheet
Public shIndex               As Worksheet
Public shPlanListe           As Worksheet
Public shGebäude             As Worksheet
Public shSPSync              As Worksheet
Public xlsmPages             As Workbook
Public CopyrightSTR          As String

Private pProjekt             As IProjekt
Private pPlanköpfe           As Collection

Public Function Projekt() As IProjekt
    With Application.ActiveWorkbook.Sheets("Projektdaten")
        If pProjekt Is Nothing Then
            Set pProjekt = _
                         ProjektFactory.Create( _
                         .range("ADM_Projektnummer").Value, _
                         AdressFactory.Create _
                         (.range("ADM_ADR_Strasse").Value, _
                          .range("ADM_ADR_PLZ").Value, _
                          .range("ADM_ADR_Ort").Value), _
                         .range("ADM_Projektbezeichnung").Value, _
                         .range("ADM_Projektphase").Value, _
                         .range("ADM_ProjektpfadSharePoint").Value)
            writelog LogInfo, "Created Projekt " & pProjekt.Projektnummer
        Else
            writelog LogInfo, "Projekt already exists " & pProjekt.Projektnummer
        End If

    End With
    Set Projekt = pProjekt
End Function

Public Function planköpfe() As Collection

    If pPlanköpfe Is Nothing Then GetPlanköpfe
    Set planköpfe = pPlanköpfe

End Function

Private Sub GetPlanköpfe()

    'TODO Create Planköpfe from Workbook / Database
    Dim row                  As range
    For Each row In shStoreData.range("A1").CurrentRegion
        pPlanköpfe.Add PlankopfFactory.LoadFromDataBase(row.row)
    Next
    writelog LogInfo, "Loaded " & pPlanköpfe.Count & " Planköpfe from the Database"
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
    If shPData.range("B4").Value < Version Then
        Dim curVersion       As String
        Dim shouldVersion    As String

        curVersion = "Bes-Gen-PZM-Add-In-V" & Version
        Select Case shPData.range("B4").Value
            Case vbNullString
                shouldVersion = "Bes-Gen-PZM-Add-In"
            Case Else
                If shPData.range("B4").Value > Version Then
                    shouldVersion = "Bes-Gen-PZM-Add-In-V" & shPData.range("B4").Value & " oder neuer"
                End If
        End Select
        MsgBox "Die Arbeitsmappe wird von dieser Version vom Beschriftungsgenerator nicht unterstüzt!" & vbLf _
             & "Bitte Update die Arbeitsmappe oder lade eine ältere Version vom Beschriftungsgenerator." & vbLf _
             & vbLf & "aktuelle Version" & vbLf & curVersion & vbLf _
             & "zu verwendende Version: " & vbLf & shouldVersion _
               , vbCritical, "Nicht unterstützte Arbeitsmappe"
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
    If ERR Then
        Set xlsmPages = Workbooks.Open(TemplatePagesXslm)
        xlsmPages.Sheets("Adressverzeichnis").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shAdress = WB.Sheets("Adressverzeichnis")
    End If
    Set shStoreData = WB.Sheets("Datenbank")
    If ERR Then
        Set xlsmPages = Workbooks.Open(TemplatePagesXslm)
        xlsmPages.Sheets("Datenbank").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shStoreData = WB.Sheets("Datenbank")
    End If
    Set shIndex = WB.Sheets("Index")
    If ERR Then
        Set xlsmPages = Workbooks.Open(TemplatePagesXslm)
        xlsmPages.Sheets("Index").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shIndex = WB.Sheets("Index")
    End If
    Set shPlanListe = WB.Sheets("Planlisten")
    If ERR Then
        Set xlsmPages = Workbooks.Open(TemplatePagesXslm)
        xlsmPages.Sheets("Planlisten").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shPlanListe = WB.Sheets("Planlisten")
    End If
    Set shVersand = WB.Sheets("Versand")
    If ERR Then
        Set xlsmPages = Workbooks.Open(TemplatePagesXslm)
        xlsmPages.Sheets("Versand").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shVersand = WB.Sheets("Versand")
    End If
    Set shGebäude = WB.Sheets("Gebäude")
    If ERR Then
        Set xlsmPages = Workbooks.Open(TemplatePagesXslm)
        xlsmPages.Sheets("Gebäude").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shGebäude = WB.Sheets("Gebäude")
    End If
    Set shPData = WB.Sheets("Projektdaten")
    If ERR Then
        Set xlsmPages = Workbooks.Open(TemplatePagesXslm)
        xlsmPages.Sheets("Projektdaten").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shPData = WB.Sheets("Projektdaten")
    End If
    Set shSPSync = WB.Sheets("SharePointSync")
    If ERR Then
        Set xlsmPages = Workbooks.Open(TemplatePagesXslm)
        xlsmPages.Sheets("SharePointSync").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shPData = WB.Sheets("SharePointSync")
    End If


    Globals.Projekt
    SetWBs = True
    writelog LogInfo, "Loaded all Workbooks in Globals Module"

End Function


