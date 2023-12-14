VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Task"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder "Tasks"
'@ModuleDescription "Task Objekt für einfacheres Management von Tasks im Projekt."


Private Type ttask
    ErfasstAm As String
    ErfasstVon As String
    FälligAm As String
    Gewerk As String
    Gebäude As String
    GebäudeTeil As String
    Geschoss As String
    Erledigt As Boolean
    Priorität As Long
    Text As String * 1024
    Pläne As New Collection
End Type

Private This                 As ttask
Private Inputs               As ttask

Public Sub FillData( _
       ErfasstAm As String, _
       ErfasstVon As String, _
       FälligAm As String, _
       Gewerk As String, _
       Gebäude As String, _
       GebäudeTeil As String, _
       Geschoss As String, _
       Erledigt As Boolean, _
       Priorität As Long, _
       Text As String _
       )

    Inputs.ErfasstAm = ErfasstAm
    Inputs.ErfasstVon = ErfasstVon
    Inputs.FälligAm = FälligAm
    Inputs.Gewerk = Gewerk
    Inputs.Gebäude = Gebäude
    Inputs.Geschoss = Geschoss
    Inputs.GebäudeTeil = GebäudeTeil
    Inputs.Erledigt = Erledigt
    Inputs.Priorität = Priorität
    Inputs.Text = Text

    If ValidInputs Then
        This = Inputs
        'Set Inputs.Pläne = GetPläne
    End If

End Sub

Public Property Get ErfasstAm() As String
    ErfasstAm = This.ErfasstAm
End Property

Public Property Get ErfasstVon() As String
    ErfasstVon = This.ErfasstVon
End Property

Public Property Get FälligAm() As String
    FälligAm = This.FälligAm
End Property

Public Property Get Gewerk() As String
    Gewerk = This.Gewerk
End Property

Public Property Get Gebäude() As String
    Gebäude = This.Gebäude
End Property

Public Property Get GebäudeTeil() As String
    GebäudeTeil = This.GebäudeTeil
End Property

Public Property Get Geschoss() As String
    Geschoss = This.Geschoss
End Property

Public Property Get Erledigt() As Boolean
    Erledigt = This.Erledigt
End Property

Public Property Get Priorität() As Long
    Priorität = This.Priorität
End Property

Public Property Get Text() As String
    Text = This.Text
End Property

Private Function ValidInputs() As Boolean

    ValidInputs = False

    ValidInputs = True

End Function

Private Function GetPläne() As Collection

    Dim lCol                 As New Collection
    Dim lPlankopf            As IPlankopf

    'TODO implement Gewerk to Plankopf
    'For Each lPlankopf In Planköpfe
    '    If _
    '    lPlankopf.Gewerk = This.Gewerk And _
    '                       lPlankopf.Gebäude = This.Gebäude And _
    '                       lPlankopf.Gebäudeteil = This.Gebäudeteil And _
    '                       lPlankopf.Geschoss = This.Geschoss _
    '                       Then
    '        lCol.Add lPlankopf
    '    End If
    'Next

End Function


