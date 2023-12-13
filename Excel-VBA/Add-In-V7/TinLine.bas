Attribute VB_Name = "TinLine"
Option Explicit
'@Folder("TinLine")

Public Sub setTinProject(ByVal Path As String)
    ' --- set the curent Projectfile as active Directory for TinLine
    Dim Projekt()            As String

    Projekt = Split(Path, "\")
    ReDim Preserve Projekt(0 To 3)
    '@Ignore DefaultMemberRequired
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "ProgrammPath", "Projekte", Join(Projekt(), "\")
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "Projekt", "AktivProjekt", Path

End Sub

Public Sub setTinPrinzipBibiothek()
    ' --- change the SymbolBibliothek to Prinzip Symbole
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "ProgrammPath", "SymbolleistePlan", "181-PR-PZM"
End Sub

Public Sub setTinPlanBibliothek()
    ' --- change the SymbolBibliothek to Plan Symbole
    WriteIni Environ$("APPDATA") & "\TinLine\TinLine 23-Deu\R23\deu\TinLine\tinlokal.ini", "ProgrammPath", "SymbolleistePlan", "181-EP-PZM"
End Sub

