Attribute VB_Name = "CustomUI"
Attribute VB_Description = "Handelt die Interaktion mit dem 'Custom Ribbon' welches beim �ffnen von Excel erstellt wird."

'@Folder "Custom UI"
'@IgnoreModule ProcedureNotUsed, VariableNotUsed
'@ModuleDescription "Handelt die Interaktion mit dem 'Custom Ribbon' welches beim �ffnen von Excel erstellt wird."

Option Explicit

Private isUILocked           As Boolean
Public myRibbon              As IRibbonUI
Const OneNoteLink = "https://rebsamennet.sharepoint.com/sites/00004ProjekteInternQS/_layouts/OneNote.aspx?id=%2Fsites%2F00004ProjekteInternQS%2FSiteAssets%2F00004%20QS%20f%C3%BCr%20Mitarbeiter"
Const OneNoteAppLink = "onenote:https://rebsamennet.sharepoint.com/sites/00004ProjekteInternQS/SiteAssets/00004%20QS%20f�r%20Mitarbeiter/"

#If VBA7 Then
    Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
#Else
    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
#End If

#If VBA7 Then

Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
#Else
Function GetRibbon(ByVal lRibbonPointer As Long) As Object
#End If

Dim objRibbon                As Object

CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
Set GetRibbon = objRibbon
Set objRibbon = Nothing

End Function

Sub isVisibleGroup(control As IRibbonControl, ByRef returnedVal As Variant)
    ' ist das Projekt schon erstellt. Pl�ne etc. sollen erst erstellt werden k�nnen wenn das Projekt auch vorhanden ist.
    If Application.ActiveWorkbook.FileFormat <> 50 Then
        returnedVal = False
        If control.ID = "customGroupNoBesGen" Then returnedVal = True Else returnedVal = False
    Else
        If Globals.shPData Is Nothing Then Globals.SetWBs
        If Globals.shPData.range("ADM_ProjektPfadCAD").value = vbNullString Then
            ' Projekt nicht erstellt ---
            Select Case control.ID
                Case "customGroupPanels"
                    returnedVal = False
                Case "customGroupBuildings"
                    returnedVal = True
                Case "customGroupExplorer"
                    returnedVal = False
                Case "customGroupHelp"
                    returnedVal = True
                Case "customGroupCreateProject"
                    returnedVal = True
                Case "customGroupNoBesGen"
                    returnedVal = False
                Case "customGroupQuickAdd"
                    returnedVal = False
                Case Else
                    returnedVal = False
            End Select
        ElseIf Globals.shPData.range("ADM_ProjektPfadCAD").value = "HLKS" Then
            ' Projekt erstellt HLKS ---
            Select Case control.ID
                Case "customGroupPanels"
                    returnedVal = True
                Case "customGroupBuildings"
                    returnedVal = False
                Case "customGroupExplorer"
                    returnedVal = False
                Case "customGroupHelp"
                    returnedVal = True
                Case "customGroupCreateProject"
                    returnedVal = False
                Case "customGroupNoBesGen"
                    returnedVal = False
                Case "customGroupQuickAdd"
                    returnedVal = True
            End Select
        Else
         ' Projekt erstellt Elektro ---
            Select Case control.ID
                Case "customGroupPanels"
                    returnedVal = True
                Case "customGroupBuildings"
                    returnedVal = False
                Case "customGroupExplorer"
                    returnedVal = True
                Case "customGroupHelp"
                    returnedVal = True
                Case "customGroupCreateProject"
                    returnedVal = False
                Case "customGroupNoBesGen"
                    returnedVal = False
                Case "customGroupQuickAdd"
                    returnedVal = True
            End Select
        End If
    End If
    writelog LogInfo, " CUSTOM UI | " & control.ID & " is visible = " & returnedVal
End Sub

Sub onLoad(ribbon As IRibbonUI)
    'PURPOSE: Run code when Ribbon loads the UI to store Ribbon Object's Pointer ID code
#If VBA7 Then
        Dim StoreRibbonPointer As LongPtr
#Else
        Dim StoreRibbonPointer As Long
#End If

    'Store Ribbon Object to Public variable
    Set myRibbon = ribbon
    isUILocked = False
    'Store pointer to IRibbonUI in a Named Range within add-in file
    StoreRibbonPointer = ObjPtr(ribbon)
    ThisWorkbook.Names.Add Name:="RibbonID", RefersTo:=StoreRibbonPointer

    writelog LogInfo, " CUSTOM UI | " & "CustomRibbon successfully Loaded"

End Sub

Public Sub RefreshRibbon()
    'PURPOSE: Refresh Ribbon UI

    Dim myRibbon             As Object

    On Error GoTo RestartExcel
    If myRibbon Is Nothing Then
        Set myRibbon = GetRibbon(Replace(ThisWorkbook.Names("RibbonID").RefersTo, "=", vbNullString))
    End If

    'Redo Ribbon Load
    myRibbon.Invalidate
    On Error GoTo 0

    Exit Sub

    'ERROR MESSAGES:
RestartExcel:
    MsgBox "Please restart Excel for Ribbon UI changes to take affect", , "Ribbon UI Refresh Failed"
    writelog LogError, " CUSTOM UI | " & "trying to refresh CustomRibbon" & vbNewLine & _
                      err.Number & vbNewLine & err.Description & vbNewLine & err.Source

End Sub

Sub onActionButton(control As IRibbonControl)
    writelog LogInfo, " CUSTOM UI | " & "Button " & control.ID & " pressed" & vbNewLine & "---------------------------"
    If Globals.shPData Is Nothing Then Globals.SetWBs
    Dim FolderPath           As String
    Dim frmPKadd             As New UserFormPlankopf
    Select Case control.ID
        Case "Objektdaten"
            ActiveWorkbook.Sheets("Geb�ude").Activate
        Case "Person"
            Dim frmAdresse   As New UserFormPerson
            frmAdresse.Show 0
        Case "CADFolder"
            FolderPath = Globals.Projekt.ProjektOrdnerCAD
            writelog LogInfo, "Opening CAD-Folder" & vbNewLine & FolderPath
            Shell "explorer.exe " & FolderPath, vbNormalFocus
        Case "XREFFolder"
            FolderPath = Globals.Projekt.ProjektOrdnerCAD & "\00_XREF"
            writelog LogInfo, "Opening CAD-Folder" & vbNewLine & FolderPath
            Shell "explorer.exe " & FolderPath, vbNormalFocus
        Case "SharePoint"
            Dim folderSP     As String: folderSP = Globals.Projekt.ProjektOrdnerSharePoint
            writelog LogInfo, " CUSTOM UI | " & "Opening SharePoint-Folder" & vbNewLine & folderSP
            If folderSP <> vbNullString Then
                ActiveWorkbook.FollowHyperlink Address:=folderSP
            Else
                MsgBox "Es ist kein SharePoint Pfad beim erstellen des Projektes eingef�gt worden." & vbNewLine & _
                       "Dieser Kann nachtr�glich in der Zelle 'D8' im Blatt 'Projektdaten' eingef�gt werden", vbInformation, "Kein SharePoint Ordner"
            End If
        Case "Drucken"
            Dim frmPrint     As New UserFormPrint
            frmPrint.Show 0
        Case "Repair"
            Dim frmRepair    As New UserFormRepair
            frmRepair.Show 0
        Case "�bersicht"
            Globals.shPData.Activate
            Dim frm�bersicht As New UserFormPlankopf�bersicht
            frm�bersicht.Show 0
        Case "Version"
            Dim frmVersion   As New UserFormInfo
            frmVersion.Show
        Case "Chat"
            'TODO E-Mail oder Teams �ffnen
        Case "Adresse"
            Globals.shAdress.Activate
            Dim frmPerson    As New UserFormPerson
            frmPerson.Show 0
        Case "Bot"
            'TODO ChatbotIntegration / URL �ffnen
        Case "Mail"
            Dim frmOutlook   As New UserFormOutlook
            frmOutlook.Show 0
        Case "CADElektro"
            Dim frmCreateElektro As New UserFormProjektErstellen
            frmCreateElektro.Show 0
        Case "HLKSElektro"
            Globals.shPData.range("ADM_ProjektPfadCAD").value = "HLKS"
        Case "Upgrade"
            Dim frmUpgrade   As New UserFormUpgrade
            frmUpgrade.Show 0
        Case "OneNote"
            ActiveWorkbook.FollowHyperlink Address:=OneNoteAppLink
        Case "PlotFolder"
            FolderPath = Environ("localappdata") & "\Bes-Gen-V7\Plot"
            writelog LogInfo, "Opening Plot-Folder" & vbNewLine & FolderPath
            Shell "explorer.exe " & FolderPath, vbNormalFocus
            ' Quick-Add
        Case "Plan"
            frmPKadd.setIcons Add
            frmPKadd.MultiPageTyp.value = 0
            frmPKadd.Show 0
        Case "Schema"
            frmPKadd.setIcons Add
            frmPKadd.MultiPageTyp.value = 1
            frmPKadd.Show 0
        Case "Prinzip"
            frmPKadd.setIcons Add
            frmPKadd.MultiPageTyp.value = 2
            frmPKadd.Show 0
        Case "Detail"
            frmPKadd.setIcons Add
            frmPKadd.MultiPageTyp.value = 3
            frmPKadd.Show 0
        Case "UpdateProject"
            Dim frmUpdateProjekt As New UserFormUpdateProjekt
            frmUpdateProjekt.Show 1
    End Select
    CustomUI.RefreshRibbon
End Sub

Sub isButtonEnabled(control As IRibbonControl, ByRef returnedVal As Variant)
    Select Case control.ID
        Case "Objektdaten"
            returnedVal = Not isUILocked
        Case Else
            returnedVal = True
    End Select
    writelog LogInfo, " CUSTOM UI | " & control.ID & " is enabled = " & returnedVal
End Sub


