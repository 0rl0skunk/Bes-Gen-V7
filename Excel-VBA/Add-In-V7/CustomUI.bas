Attribute VB_Name = "CustomUI"
Attribute VB_Description = "Handelt die Interaktion mit dem 'Custom Ribbon' welches beim öffnen von Excel erstellt wird."
'@IgnoreModule ProcedureNotUsed, VariableNotUsed
Option Explicit
'@Folder "Custom UI"
'@ModuleDescription "Handelt die Interaktion mit dem 'Custom Ribbon' welches beim öffnen von Excel erstellt wird."

Private isUILocked           As Boolean
Public myRibbon              As IRibbonUI

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
                    returnedVal = False
                Case "customGroupCreateProject"
                    returnedVal = True
                Case "customGroupNoBesGen"
                    returnedVal = False
                Case Else
                    returnedVal = False
            End Select
        Else
        ' Projekt erstellt ---
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
            End Select
        End If
    End If
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

    writelog LogInfo, "CustomRibbon successfully Loaded"

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
    writelog LogError, "trying to refresh CustomRibbon" & vbNewLine & _
                      err.Number & vbNewLine & err.description & vbNewLine & err.source

End Sub

Sub onActionButton(control As IRibbonControl)
    writelog LogInfo, "Button " & control.ID & " pressed" & vbNewLine & "---------------------------"
    If Globals.shPData Is Nothing Then Globals.SetWBs
    Select Case control.ID
        Case "Objektdaten"
            ActiveWorkbook.Sheets("Gebäude").Activate
        Case "Person"
            Dim frmAdresse   As New UserFormPerson
            frmAdresse.Show 1
        Case "CADFolder"
            Dim folderpath   As String: folderpath = Globals.Projekt.ProjektOrdnerCAD
            writelog LogInfo, "Opening CAD-Folder" & vbNewLine & folderpath
            Shell "explorer.exe " & folderpath, vbNormalFocus
        Case "SharePoint"
            Dim folderSP     As String: folderSP = Globals.Projekt.ProjektOrdnerSharePoint
            writelog LogInfo, "Opening SharePoint-Folder" & vbNewLine & folderSP
            ActiveWorkbook.FollowHyperlink Address:=folderSP
        Case "Drucken"
            Dim frmPrint     As New UserFormPrint
            frmPrint.Show 1
        Case "Repair"
            Dim frmRepair    As New UserFormRepair
            frmRepair.Show 1
        Case "Übersicht"
            Globals.shPData.Activate
            Dim frmÜbersicht As New UserFormPlankopfübersicht
            frmÜbersicht.Show
        Case "Version"
            Dim frmVersion   As New UserFormInfo
            frmVersion.Show 1
        Case "Chat"
            'TODO E-Mail oder Teams öffnen
        Case "Adresse"
            Globals.shAdress.Activate
            Dim frmPerson    As New UserFormPerson
            frmPerson.Show 0
        Case "Bot"
            'TODO ChatbotIntegration / URL öffnen
        Case "Mail"
            Dim frmOutlook   As New UserFormOutlook
            frmOutlook.Show 1
        Case "CADElektro"
            Dim frmCreateElektro As New UserFormProjektErstellen
            frmCreateElektro.Show 1
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
    writelog LogInfo, control.ID & " is enabled = " & returnedVal
End Sub


