Attribute VB_Name = "Stapelplot"
'@Folder("Print")
Option Explicit

Private a
Private dsd                  As String           ' Dateiname von Stapelplott Datei
Private OutputFolder         As String
Private NewFiles As Long
Private OldFiles As Long
Private pPlanköpfe As Collection

Sub plotPlanliste()

    Dim fs, a, i             As Integer, search As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim scr                  As String
    scr = "C:\Users\Public\Documents\plotter.scr"
    Set a = fs.CreateTextFile(scr, True)
    a.WriteLine ("-PUBLISH")
    a.WriteLine (dsd)
    a.Close

    Dim wsh                  As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn         As Boolean: waitOnReturn = True
    Dim windowStyle          As Integer: windowStyle = 1
    Dim errorCode            As Integer
    
    OldFiles = CountFiles(OutputFolder)
    
    wsh.Run """C:\Program Files\TinLine\TinLine 23-Deu\accoreconsole.exe"" /i ""H:\TinLine\01_Standards\TinBlank.dwg"" /s ""C:\Users\Public\Documents\plotter.scr"" /l EN-US", windowStyle, waitOnReturn
    
    ' File Counting
    If Not CountFiles(OutputFolder) = OldFiles + NewFiles Then
    Select Case MsgBox("Es wurden nicht alle Pläne geplottet." & vbNewLine & "Soll geprüft werden welche Pläne fehlen?", vbYesNo, "Fehler beim plotten")
    Case vbYes
        Checkplot
    Case vbNo
    End Select
    End If
    Dim CreatedFiles
    CreatedFiles = CountFiles(OutputFolder) - OldFiles
    
    Select Case MsgBox("Es wurden " & CreatedFiles & " von " & NewFiles & " Plänen erstellt." & vbNewLine & "Pfad im Explorer öffnen?", vbYesNo, "Pläne erstellt")
        Case vbYes
            ' open explorer
            Shell "explorer.exe" & " " & OutputFolder, vbNormalFocus
            Exit Sub
        Case vbNo
            ' exit sub
            Exit Sub
    End Select

End Sub

Public Sub Checkplot()
    Dim fso As New FileSystemObject
    Dim File As scripting.File
    Dim PDFFile As IPlankopf
    Dim i As Long
    
    For Each File In fso.GetFolder(OutputFolder).files
    i = 1
        For Each PDFFile In pPlanköpfe
            If File.Name = PDFFile.PDFFileName & ".pdf" Then
                pPlanköpfe.Remove i
                Exit For
            End If
            i = i + 1
        Next
    Next
    
    Dim msg As String
    msg = "Folgende Pläne müssen überprüft werden:" & vbNewLine
    For Each PDFFile In pPlanköpfe
    msg = msg & vbNewLine & PDFFile.dwgFile & " | " & PDFFile.LayoutName
    Next
    
    MsgBox msg, vbInformation, "Fehlerhafte Pläne"
    
End Sub

Public Function CreatePlotList(ByVal Planköpfe As Collection) As String

    Dim folder               As String, strFolderExists As String
    Dim outputCol            As New Collection
    Dim Plan                 As IPlankopf
    
    Set pPlanköpfe = Planköpfe
    Set Planköpfe = Nothing
    
    NewFiles = pPlanköpfe.Count
    
    If Globals.shPData Is Nothing Then Globals.SetWBs

    folder = Globals.Projekt.ProjektOrdnerCAD & "\99_Planlisten"
    strFolderExists = dir(folder)

    'If strFolderExists = "" Then MkDir folder
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then                       ' if OK is pressed
            OutputFolder = .SelectedItems(1)
        End If
    End With


    Dim filename             As String: filename = Format(Now, "YYMMDDhhmmss")
    Dim i                    As Integer, search As String
    Set a = CreateObject("Scripting.FileSystemObject").CreateTextFile(Globals.Projekt.ProjektOrdnerCAD & "\99 Planlisten\" & filename & ".dsd", True)
    dsd = Globals.Projekt.ProjektOrdnerCAD & "\99 Planlisten\" & filename & ".dsd"
    a.WriteLine ("[DWF6Version]")
    a.WriteLine ("Ver=1")
    a.WriteLine ("[DWF6MinorVersion]")
    a.WriteLine ("MinorVer=1")

    For Each Plan In pPlanköpfe
        Eintrag Plan
    Next




    a.WriteLine ( _
                "[Target]" & vbLf & _
                "Type=2" & vbLf & _
                "DWF=" & vbLf & _
                "OUT=" & OutputFolder & vbLf & _
                "PWD=")


    a.WriteLine ( _
                "[PdfOptions]" & vbCrLf & _
                "IncludeHyperlinks=TRUE" & vbCrLf & _
                "CreateBookmarks=TRUE" & vbCrLf & _
                "CaptureFontsInDrawing=TRUE" & vbCrLf & _
                "ConvertTextToGeometry=FALSE" & vbCrLf & _
                "VectorResolution=1200" & vbCrLf & _
                "RasterResolution=400")

    a.WriteLine ("[AutoCAD Block Data]" & vbCrLf & _
                 "IncludeBlockInfo=0" & vbCrLf & _
                 "BlockTmplFilePath=" & vbCrLf & _
                 "[SheetSet Properties]" & vbCrLf & _
                 "IsSheetSet=FALSE" & vbCrLf & _
                 "IsHomogeneous=FALSE" & vbCrLf & _
                 "SheetSet Name=" & vbCrLf & _
                 "NoOfCopies=1" & vbCrLf & _
                 "PlotStampOn=FALSE" & vbCrLf & _
                 "ViewFile=TRUE" & vbCrLf & _
                 "JobID=0" & vbCrLf & _
                 "SelectionSetName=" & vbCrLf & _
                 "AcadProfile=<<Unbenanntes Profil>>" & vbCrLf & _
                 "CategoryName=" & vbCrLf & _
                 "LogFilePath=" & vbCrLf & _
                 "IncludeLayer=TRUE" & vbCrLf & _
                 "LineMerge=FALSE" & vbCrLf & _
                 "CurrentPrecision=" & vbCrLf & _
                 "PromptForDwfName=TRUE" & vbCrLf & _
                 "PwdProtectPublishedDWF=FALSE" & vbCrLf _
               & "PromptForPwd=FALSE" & vbCrLf & _
                 "RepublishingMarkups=FALSE" & vbCrLf & _
                 "PublishSheetSetMetadata=FALSE" & vbCrLf & _
                 "PublishSheetMetadata=FALSE" & vbCrLf & _
                 "3DDWFOptions=0 1")
    a.Close

    plotPlanliste

    CreatePlotList = OutputFolder

End Function

Private Sub Eintrag(Plan As IPlankopf)
    a.WriteLine ("[DWF6Sheet:" & Plan.PDFFileName & "]") ' PDF Ablage
    a.WriteLine ("DWG=" & Plan.dwgFile)          ' DWG Ablage
    a.WriteLine ("Layout=" & Plan.LayoutName)    ' Plannummer / Layoutname
    a.WriteLine ("Setup=")
    a.WriteLine ("OriginalSheetPath=" & Plan.dwgFile) ' DWG Ablage
    a.WriteLine ("Has Plot Port=0")
    a.WriteLine ("Has3DDWF=0")
    a.WriteLine (" ")
End Sub


