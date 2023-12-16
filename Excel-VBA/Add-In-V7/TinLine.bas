Attribute VB_Name = "TinLine"
Attribute VB_Description = "Funktionen welche das TinLine Betreffen."
Option Explicit
'@Folder("TinLine")
'@ModuleDescription "Funktionen welche das TinLine Betreffen."

Public Sub setTinProject(ByVal Path As String)
    ' Aktuelles Projekt als aktiv setzen
    Dim Projekt()            As String

    Projekt = Split(Path, "\")
    ReDim Preserve Projekt(0 To 3)
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "ProgrammPath", "Projekte", Join(Projekt(), "\")
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "Projekt", "AktivProjekt", Path

End Sub

Public Sub setTinPrinzipBibiothek()
    ' SymbolBibliothek auf Prinzip Symbole setzen
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "ProgrammPath", "SymbolleistePlan", "181-PR-PZM"
End Sub

Public Sub setTinPlanBibliothek()
    ' SymbolBibliothek auf Plan Symbole setzen
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "ProgrammPath", "SymbolleistePlan", "181-EP-PZM"
End Sub

