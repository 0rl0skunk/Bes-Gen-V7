VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Task"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Task Objekt f�r einfacheres Management von Tasks im Projekt."
'@IgnoreModule VariableNotUsed
Option Explicit
'@Folder "Tasks"
'@ModuleDescription "Task Objekt f�r einfacheres Management von Tasks im Projekt."


Private Type ttask
    ErfasstAm As String
    ErfasstVon As String
    F�lligAm As String
    Gewerk As String
    Geb�ude As String
    Geb�udeteil As String
    Geschoss As String
    Erledigt As Boolean
    Priorit�t As Long
    text As String * 1024
    Pl�ne As New Collection
End Type

Private This                 As ttask
Private Inputs               As ttask

Public Sub Filldata( _
       ErfasstAm As String, _
       ErfasstVon As String, _
       F�lligAm As String, _
       Gewerk As String, _
       Geb�ude As String, _
       Geb�udeteil As String, _
       Geschoss As String, _
       Erledigt As Boolean, _
       Priorit�t As Long, _
       text As String _
       )

    Inputs.ErfasstAm = ErfasstAm
    Inputs.ErfasstVon = ErfasstVon
    Inputs.F�lligAm = F�lligAm
    Inputs.Gewerk = Gewerk
    Inputs.Geb�ude = Geb�ude
    Inputs.Geschoss = Geschoss
    Inputs.Geb�udeteil = Geb�udeteil
    Inputs.Erledigt = Erledigt
    Inputs.Priorit�t = Priorit�t
    Inputs.text = text

    If ValidInputs Then
        This = Inputs
        'Set Inputs.Pl�ne = GetPl�ne
    End If

End Sub

Public Property Get ErfasstAm() As String
    ErfasstAm = This.ErfasstAm
End Property

Public Property Get ErfasstVon() As String
    ErfasstVon = This.ErfasstVon
End Property

Public Property Get F�lligAm() As String
    F�lligAm = This.F�lligAm
End Property

Public Property Get Gewerk() As String
    Gewerk = This.Gewerk
End Property

Public Property Get Geb�ude() As String
    Geb�ude = This.Geb�ude
End Property

Public Property Get Geb�udeteil() As String
    Geb�udeteil = This.Geb�udeteil
End Property

Public Property Get Geschoss() As String
    Geschoss = This.Geschoss
End Property

Public Property Get Erledigt() As Boolean
    Erledigt = This.Erledigt
End Property

Public Property Get Priorit�t() As Long
    Priorit�t = This.Priorit�t
End Property

Public Property Get text() As String
    text = This.text
End Property

Private Function ValidInputs() As Boolean

    ValidInputs = False

    ValidInputs = True

End Function

Private Function GetPl�ne() As Collection

    Dim lCol                 As New Collection
    Dim lPlankopf            As IPlankopf

    'TODO implement Gewerk to Plankopf
    'For Each lPlankopf In Plank�pfe
    '    If _
    '    lPlankopf.Gewerk = This.Gewerk And _
    '                       lPlankopf.Geb�ude = This.Geb�ude And _
    '                       lPlankopf.Geb�udeteil = This.Geb�udeteil And _
    '                       lPlankopf.Geschoss = This.Geschoss _
    '                       Then
    '        lCol.Add lPlankopf
    '    End If
    'Next

End Function


