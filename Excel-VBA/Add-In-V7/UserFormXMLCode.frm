VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormXMLCode 
   ClientHeight    =   11130
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9600.001
   OleObjectBlob   =   "UserFormXMLCode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormXMLCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








'@Folder "Plankopf"
Option Explicit
Private icons                As UserFormIconLibrary

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

Public Sub load(ByVal filepath As String, ByVal Plankopfnummer As Long)

    Dim xmlDOMDoc            As New MSXML2.DOMDocument60
    xmlDOMDoc.load filepath
    Me.LabelInstructions.Caption = filepath

    Dim PKnr                 As String: PKnr = "PK" & Plankopfnummer
    Dim XMLstr               As String * 1024

    Dim RootNode             As MSXML2.IXMLDOMElement
    Set RootNode = xmlDOMDoc.DocumentElement

    Dim ChildNode            As MSXML2.IXMLDOMElement

    For Each ChildNode In RootNode.ChildNodes
        If ChildNode.BaseName = "PK" And ChildNode.FirstChild.text = CStr(Plankopfnummer) Then XMLstr = XMLstr & vbNewLine & CStr(ChildNode.xml)
        If ChildNode.HasChildNodes And ChildNode.BaseName = PKnr Then
            XMLstr = XMLstr & vbNewLine & CStr(ChildNode.xml)
        End If
    Next

    Me.TextBox1.Value = xmlDOMDoc.xml
    Me.TextBox1.TextAlign = fmTextAlignLeft

End Sub

Private Sub UserForm_Initialize()

    Me.TextBox1.Value = "--- No XML-File was loaded ---"
    Me.TextBox1.TextAlign = fmTextAlignCenter

End Sub

