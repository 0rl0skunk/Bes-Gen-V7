Attribute VB_Name = "TinLine"
Attribute VB_Description = "Funktionen welche das TinLine Betreffen."

'@Folder("TinLine")
'@ModuleDescription "Funktionen welche das TinLine Betreffen."

Public Enum TinBibliothek
    EP = 1
    PR = 2
    ES = 3
    TF = 5
    BS = 6
End Enum

Option Explicit

Public Sub setTinProject(ByVal path As String)
    ' Aktuelles Projekt als aktiv setzen
    Dim Projekt()            As String

    Projekt = Split(path, "\")
    ReDim Preserve Projekt(0 To 3)
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "ProgrammPath", "Projekte", Join(Projekt(), "\")
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "Projekt", "AktivProjekt", path

End Sub

Public Sub setBibliothek(ByVal Bibliothek As TinBibliothek)
    ' SymbolBibliothek auf Plan Symbole setzen
    Select Case Bibliothek
    Case 1
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "ProgrammPath", "SymbolleistePlan", "181-EP-PZM"
    Case 2
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "ProgrammPath", "SymbolleistePlan", "181-PR-PZM"
    Case 3
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "ProgrammPath", "SymbolleistePlan", "182-Elektroschema"
    Case 5
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "ProgrammPath", "SymbolleistePlan", "181-TF-PZM"
    Case 6
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "ProgrammPath", "SymbolleistePlan", "181-Brandschutz"
    End Select
End Sub
