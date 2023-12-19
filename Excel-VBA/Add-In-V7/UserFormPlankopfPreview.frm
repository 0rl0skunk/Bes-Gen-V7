VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPlankopfPreview 
   ClientHeight    =   4800
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11880
   OleObjectBlob   =   "UserFormPlankopfPreview.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormPlankopfPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Plankopf"
Option Explicit

Private pPlankopfnummer      As Long
Private icons                As UserFormIconLibrary

Private Sub CommandButtonClose_Click()
    Unload Me
End Sub

Public Sub LoadClass(ByVal Plankopf As IPlankopf, ByVal Projekt As IProjekt)

    Me.PA40.Caption = Plankopf.Planüberschrift
    Me.PA41.Caption = Plankopf.LayoutGrösse
    Me.PA42.Caption = Plankopf.LayoutMasstab
    Me.PA43.Caption = Plankopf.Plannummer
    Me.PA44.Caption = Plankopf.LayoutPlanstand
    Me.PA30.Caption = Split(Plankopf.Gezeichnet, " ; ")(0)
    Me.PA31.Caption = Split(Plankopf.Gezeichnet, " ; ")(1)
    Me.PA32.Caption = Split(Plankopf.Geprüft, " ; ")(0)
    Me.PA33.Caption = Split(Plankopf.Geprüft, " ; ")(1)
    '--- Projektaddresse
    Me.LabelProjektAdresse.Caption = Projekt.ProjektBezeichnung & vbNewLine & Projekt.Projektadresse.Komplett
    Me.Projektnummer.Caption = Projekt.Projektnummer
    Me.Projektphase.Caption = Projekt.ProjektphaseNummer & " - " & Projekt.Projektphase
    Me.Plot.Caption = Format$(Now(), "DD.MM.YYYY HH:mm")

End Sub

Public Sub LoadXML(ByVal filepath As String, ByVal Plankopfnummer As Long)

    Dim xmlDOMDoc            As New MSXML2.DOMDocument60
    xmlDOMDoc.load filepath

    pPlankopfnummer = Plankopfnummer
    Dim PKNr                 As String: PKNr = "PK" & pPlankopfnummer

    Dim ParentNode           As MSXML2.IXMLDOMElement
    Set ParentNode = xmlDOMDoc.DocumentElement

    Dim ChildNode            As MSXML2.IXMLDOMElement
    Dim GrandChildNode       As MSXML2.IXMLDOMElement

    For Each ChildNode In ParentNode.ChildNodes
        If ChildNode.HasChildNodes And ChildNode.BaseName = PKNr Then
            For Each GrandChildNode In ChildNode.ChildNodes
                Select Case GrandChildNode.text
                    Case "PA40": Me.PA40.Caption = GrandChildNode.NextSibling.NextSibling.text
                    Case "PA41": Me.PA41.Caption = GrandChildNode.NextSibling.NextSibling.text
                    Case "PA42": Me.PA42.Caption = GrandChildNode.NextSibling.NextSibling.text
                    Case "PA43": Me.PA43.Caption = GrandChildNode.NextSibling.NextSibling.text
                    Case "PA44": Me.PA44.Caption = GrandChildNode.NextSibling.NextSibling.text
                    Case "PA30": Me.PA30.Caption = GrandChildNode.NextSibling.NextSibling.text
                    Case "PA31": Me.PA31.Caption = GrandChildNode.NextSibling.NextSibling.text
                    Case "PA32": Me.PA32.Caption = GrandChildNode.NextSibling.NextSibling.text
                    Case "PA33": Me.PA33.Caption = GrandChildNode.NextSibling.NextSibling.text
                End Select
            Next
        End If
    Next

End Sub

