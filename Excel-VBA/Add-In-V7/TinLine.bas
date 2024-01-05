Attribute VB_Name = "TinLine"
Attribute VB_Description = "Funktionen welche das TinLine Betreffen."

'@Folder("TinLine")
'@ModuleDescription "Funktionen welche das TinLine Betreffen."
'@Version "Release V1.0.0"

Option Explicit

Public Sub setTinProject(ByVal path As String)
    ' Aktuelles Projekt als aktiv setzen
    Dim Projekt()            As String

    Projekt = Split(path, "\")
    ReDim Preserve Projekt(0 To 3)
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "ProgrammPath", "Projekte", Join(Projekt(), "\")
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "Projekt", "AktivProjekt", path

End Sub

Public Sub setTinPrinzipBibiothek()
    ' SymbolBibliothek auf Prinzip Symbole setzen
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "ProgrammPath", "SymbolleistePlan", "181-PR-PZM"
End Sub

Public Sub setTinPlanBibliothek()
    ' SymbolBibliothek auf Plan Symbole setzen
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "ProgrammPath", "SymbolleistePlan", "181-EP-PZM"
End Sub

