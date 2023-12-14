Attribute VB_Name = "CustomUI"
Option Explicit
'@Folder "Custom UI"
'@Ignore ProcedureNotUsed
'@ModuleDescription "Handelt die Interaktion mit dem 'Custom Ribbon' welches beim öffnen von Excel erstellt wird."

Private isUILocked           As Boolean

Private Type TpData
    Number As String
    Name As String
    Phase As String
End Type

Private TextNew              As TpData
Private TextOld              As TpData

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

Sub isVisibleGroup(control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        Case "customGroupPanels"
        Case "customGroupSIA"
        Case "customGroupBuildings"
        Case "customGroupExplorer"
        Case "customGroupHelp"
        Case Else
    End Select
    returnedVal = True
End Sub

Sub IsButtonVisible(control As IRibbonControl, ByRef returnedVal)

    Select Case control.ID
        Case "unLockProjekt"
            returnedVal = isUILocked
        Case "LockProjekt"
            returnedVal = Not isUILocked
    End Select

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

    writelog "Info", "CustomRibbon successfully Loaded"

End Sub

Public Sub RefreshRibbon()
    'PURPOSE: Refresh Ribbon UI

    Dim myRibbon             As Object

    On Error GoTo RestartExcel
    If myRibbon Is Nothing Then
        Set myRibbon = GetRibbon(Replace(ThisWorkbook.Names("RibbonID").RefersTo, "=", ""))
    End If

    'Redo Ribbon Load
    myRibbon.Invalidate
    On Error GoTo 0

    Exit Sub

    'ERROR MESSAGES:
RestartExcel:
    MsgBox "Please restart Excel for Ribbon UI changes to take affect", , "Ribbon UI Refresh Failed"
    writelog "Error", "trying to refresh CustomRibbon" & vbNewLine & _
                     ERR.Number & vbNewLine & ERR.description & vbNewLine & ERR.source

End Sub

Sub onActionButton(control As IRibbonControl)
    Globals.SetWBs
    If Not isUILocked Then
        Select Case control.ID
            Case "Objektdaten"
                'TODO Objektdaten Input UserForm
                'Dim frmObj As New UserFormTemplateV7
                'frmObj.Show 0
                ActiveWorkbook.Sheets("Gebäude").Activate
        End Select
    End If
    Select Case control.ID
        Case "CADFolder"
            Shell "explorer.exe " & Globals.Projekt.ProjektOrdnerCAD, vbNormalFocus
        Case "SharePoint"
            ActiveWorkbook.FollowHyperlink Address:=Globals.Projekt.ProjektOrdnerSharePoint
        Case "Drucken"
            'TODO Drucken UserForm
            Dim frmPrint     As New UserFormPrint
            frmPrint.Show 1
        Case "Repair"
            'TODO Reparieren UserForm
            Dim frmRepair    As New UserFormRepair
            frmRepair.Show 1
        Case "Übersicht"
            'TODO Planübersicht UserForm
            Globals.shPData.Activate
            Dim frmÜbersicht As New UserFormPlankopfübersicht
            frmÜbersicht.Show
        Case "Version"
            Dim frmVersion   As New UserFormInfo
            frmVersion.Show 1
            'TODO Übersicht Planköpfe UserForm
        Case "Chat"
            'TODO E-Mail oder Teams öffnen
        Case "Bot"
            'TODO ChatbotIntegration / URL öffnen
        Case "LockProjekt"
            isUILocked = Not isUILocked
            CustomUI.RefreshRibbon
        Case "unLockProjekt"
            isUILocked = Not isUILocked
            CustomUI.RefreshRibbon
    End Select

    writelog "Info", "Button " & control.ID & " pressed"
End Sub

Sub onChange(control As IRibbonControl, Text As String)
    TextOld = TextNew
    Select Case control.ID
        Case "Projektnummer"
            Application.ActiveWorkbook.Sheets("Projektdaten").range("ADM_Projektnummer").Value = Text
        Case "Projektname"
            Application.ActiveWorkbook.Sheets("Projektdaten").range("ADM_ProjektBezeichnung").Value = Text
        Case "comboBoxProjektphase"
            Application.ActiveWorkbook.Sheets("Projektdaten").range("ADM_Projektphase").Value = Text
    End Select

End Sub

Sub CallBackGetText(control As IRibbonControl, ByRef returnedVal)

    On Error Resume Next
    Select Case control.ID
        Case "Projektnummer"
            returnedVal = Application.ActiveWorkbook.Sheets("Projektdaten").range("ADM_Projektnummer").Value
        Case "Projektname"
            returnedVal = Application.ActiveWorkbook.Sheets("Projektdaten").range("ADM_ProjektBezeichnung").Value
        Case "comboBoxProjektphase"
            returnedVal = Application.ActiveWorkbook.Sheets("Projektdaten").range("ADM_Projektphase").Value
    End Select
    On Error GoTo 0
    writelog "Info", control.ID & " Text set to: " & returnedVal
End Sub

Sub isButtonEnabled(control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        Case "Objektdaten"
            returnedVal = Not isUILocked
        Case Else
            returnedVal = True
    End Select
    writelog "Info", control.ID & " is enabled = " & returnedVal
End Sub

Sub isTextEnabled(control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        Case "Projektnummer"
            returnedVal = Not isUILocked
        Case "Projektname"
            returnedVal = Not isUILocked
        Case "comboBoxProjektphase"
            returnedVal = Not isUILocked
    End Select
    writelog "Info", control.ID & " is enabled = " & returnedVal
End Sub


