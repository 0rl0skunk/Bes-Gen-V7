Attribute VB_Name = "Globals"
Option Explicit
'@Folder "Excel-Items"

' defines all nescessairy links and other global variables

Global Const Version = 5#

Global Const maxlen = 35                         'Maximale Anzahl Zeichen in der Planbschriftung
Global Const OrdnerVorlage = "H:\TinLine\01_Standards\00_Vorlageordner"
Global Const VorlageEPDWG = "H:\TinLine\01_Standards\EP-Vorlage.dwg"
Global Const VorlageEPDWGGEB = "H:\TinLine\01_Standards\EP-Vorlage_GEB.dwg"
Global Const VorlagePRDWG = "H:\TinLine\01_Standards\PR-Vorlage.dwg"
Global Const TinLineProjekte = "H:\TinLine\00_Projekte\"
Global Const XMLVorlage = "H:\TinLine\01_Standards\transform.xsl"
Global Const LogPath = "C:\Users\Public\Documents\Bes-Gen-V2.1.log"
Global Const listCols = 30
Global Const minListHeight = 20
Global Const HighlightColor = vbCyan

Public WB                    As Workbook
Public shPData               As Worksheet
Public shStoreData           As Worksheet
Public shAdress              As Worksheet
Public shVersand             As Worksheet
Public shIndex               As Worksheet
Public shPlanListe           As Worksheet
Public shGeb�ude             As Worksheet
Public xlsmPages             As Workbook
Public CopyrightSTR          As String
Public UserName              As String

Private pProjekt             As IProjekt
Private pPlank�pfe           As Collection
Public pWorkbook             As Workbook
Public isBesGen              As Boolean

Public Function Projekt() As IProjekt
    With Application.ActiveWorkbook.Sheets("Projektdaten")
        If pProjekt Is Nothing Then Set pProjekt = _
           ProjektFactory.Create( _
           .range("ADM_Projektnummer").Value, _
           AdressFactory.Create _
           (.range("ADM_ADR_Strasse").Value, _
            .range("ADM_ADR_PLZ").Value, _
            .range("ADM_ADR_Ort").Value), _
           .range("ADM_Projektbezeichnung").Value, _
           .range("ADM_Projektphase").Value, _
           .range("ADM_ProjektpfadSharePoint").Value)
    End With
    Set Projekt = pProjekt

End Function

Public Function Plank�pfe() As Collection

    If pPlank�pfe Is Nothing Then GetPlank�pfe
    Set Plank�pfe = pPlank�pfe

End Function

Private Sub GetPlank�pfe()

    'TODO Create Plank�pfe from Workbook / Database

End Sub

Function Initialize() As Boolean

    Set WB = ActiveWorkbook

    CopyrightSTR = _
                 "Release: " & Version & vbLf _
               & ChrW(&HA9) & Format(Now(), "yyyy") & " Orlando Bassi"

    Set shPData = WB.Sheets("Projektdaten")

    On Error GoTo 0
    ' Version checking
    If shPData.range("B4").Value < Version Then
        Dim curVersion       As String, shouldVersion As String
        curVersion = "Bes-Gen-PZM-Add-In-V" & Version
        Select Case shPData.range("B4").Value
            Case ""
                shouldVersion = "Bes-Gen-PZM-Add-In"
            Case Else
                If shPData.range("B4").Value > Version Then
                    shouldVersion = "Bes-Gen-PZM-Add-In-V" & shPData.range("B4").Value & " oder neuer"
                End If
        End Select
        MsgBox "Die Arbeitsmappe wird von dieser Version vom Beschriftungsgenerator nicht unterst�zt!" & vbLf _
             & "Bitte Update die Arbeitsmappe oder lade eine �ltere Version vom Beschriftungsgenerator." & vbLf _
             & vbLf & "aktuelle Version" & vbLf & curVersion & vbLf _
             & "zu verwendende Version: " & vbLf & shouldVersion _
               , vbCritical, "Nicht unterst�tzte Arbeitsmappe"
        Exit Function
    Else
    End If

    On Error Resume Next
    WB.Activate

    SetWBs

    xlsmPages.Close False
    Set xlsmPages = Nothing
    On Error GoTo 0

    shPData.Activate

End Function

Public Function SetWBs()
    ' set workbooks to easily reference them from within the Add-In
    ' reference the worksheeet by 'Globals.shAdress' and the workbook by 'Globals.WB'
    If WB Is Nothing Then Set WB = Application.ActiveWorkbook
    Dim i                    As Integer
    Set shAdress = WB.Sheets("Adressverzeichnis")
    If ERR Then
        Set xlsmPages = Workbooks.Open("H:\TinLine\01_Standards\Beschriftungsgenerator\Bes-Gen-PZM_Templates.xlsm")
        xlsmPages.Sheets("Adressverzeichnis").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shAdress = WB.Sheets("Adressverzeichnis")
    End If
    Set shStoreData = WB.Sheets("Datenbank")
    If ERR Then
        Set xlsmPages = Workbooks.Open("H:\TinLine\01_Standards\Beschriftungsgenerator\Bes-Gen-PZM_Templates.xlsm")
        xlsmPages.Sheets("Datenbank").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shStoreData = WB.Sheets("Datenbank")
    End If
    Set shIndex = WB.Sheets("Index")
    If ERR Then
        Set xlsmPages = Workbooks.Open("H:\TinLine\01_Standards\Beschriftungsgenerator\Bes-Gen-PZM_Templates.xlsm")
        xlsmPages.Sheets("Index").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shIndex = WB.Sheets("Index")
    End If
    Set shPlanListe = WB.Sheets("Planlisten")
    If ERR Then
        Set xlsmPages = Workbooks.Open("H:\TinLine\01_Standards\Beschriftungsgenerator\Bes-Gen-PZM_Templates.xlsm")
        xlsmPages.Sheets("Planlisten").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shPlanListe = WB.Sheets("Planlisten")
    End If
    Set shVersand = WB.Sheets("Versand")
    If ERR Then
        Set xlsmPages = Workbooks.Open("H:\TinLine\01_Standards\Beschriftungsgenerator\Bes-Gen-PZM_Templates.xlsm")
        xlsmPages.Sheets("Versand").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shVersand = WB.Sheets("Versand")
    End If
    Set shGeb�ude = WB.Sheets("Geb�ude")
    If ERR Then
        Set xlsmPages = Workbooks.Open("H:\TinLine\01_Standards\Beschriftungsgenerator\Bes-Gen-PZM_Templates.xlsm")
        xlsmPages.Sheets("Geb�ude").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shGeb�ude = WB.Sheets("Geb�ude")
    End If
    Set shPData = WB.Sheets("Projektdaten")
    If ERR Then
        Set xlsmPages = Workbooks.Open("H:\TinLine\01_Standards\Beschriftungsgenerator\Bes-Gen-PZM_Templates.xlsm")
        xlsmPages.Sheets("Projektdaten").copy after:=WB.Sheets(WB.Sheets.Count)
        Set shPData = WB.Sheets("Projektdaten")
    End If

    Globals.Projekt

End Function

