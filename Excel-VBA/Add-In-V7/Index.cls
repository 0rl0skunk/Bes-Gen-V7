VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule EmptyStringLiteral
'@Folder("Index")
Option Explicit

Implements IIndex

Private Type TIndex
    '--- Identifiers ---
    index As String
    ID As String
    IDPlan As String
    '--- Zeichner ---
    GezeichnetPerson As String
    GezeichnetDatum As String
    GeprüftPerson As String
    GeprüftDatum As String
    '--- Planbezeichnung ---
    Klartext As String
    '--- File Path ---
    DWGFileName As String
    DWGFilePath As String
    XMLFileName As String
    XMLFilePath As String
End Type

Private This                 As TIndex
Private Inputs               As TIndex

Private Property Get xmlfile() As String
    xmlfile = This.XMLFilePath & "\" & This.XMLFileName
End Property

Private Property Get DWGFile() As String
    DWGFile = This.DWGFilePath & "\" & This.DWGFileName
End Property

Public Sub FillData( _
       ByVal IDPlan As String, _
       ByVal GezeichnetPerson As String, _
       ByVal GezeichnetDatum As String, _
       ByVal Klartext As String, _
       Optional ByVal Letter As String = vbNullString, _
       Optional ByVal GeprüftPerson As String = vbNullString, _
       Optional ByVal GeprüftDatum As String = vbNullString _
       )

    Inputs.IDPlan = IDPlan
    Inputs.index = Letter
    Inputs.GezeichnetPerson = GezeichnetPerson
    Inputs.GezeichnetDatum = GezeichnetDatum
    Inputs.GeprüftPerson = GeprüftPerson
    Inputs.GeprüftDatum = GeprüftDatum
    Inputs.Klartext = Klartext

    If validateInputs Then
        This = Inputs
    End If

End Sub

Private Function validateInputs() As Boolean

    validateInputs = False
    If Len(Inputs.IDPlan) = 0 Then ERR.Raise 1, "Index > ValidateInputs", "No 'Plan ID' given"
    '--- Zeichner ---
    If Len(Inputs.GezeichnetPerson) = 0 Then ERR.Raise 1, "Index > ValidateInputs", "No 'GezeichnetPerson' given"
    If Len(Inputs.GezeichnetDatum) = 0 Then ERR.Raise 1, "Index > ValidateInputs", "No 'GezeichnetDatum' given"
    If Inputs.GeprüftPerson = "NONE" Then Inputs.GeprüftPerson = vbNullString
    If Inputs.GeprüftDatum = "NONE" Then Inputs.GeprüftDatum = ""
    '--- Planbezeichnung ---
    If Len(Inputs.Klartext) = 0 Then ERR.Raise 1, "Index > ValidateInputs", "No 'Klartext' given"
    validateInputs = True
    
    If This.index = vbNullString Then
    Globals.SetWBs
        Dim indexes(1 To 26) As String, LastIndex As String
        Dim i As Integer
        For i = 65 To 90
            indexes(i - 64) = Chr(i)
        Next
        If IsError(LastIndex = Globals.shIndex.range("A:A").Find(Inputs.IDPlan, searchdirection:=xlPrevious).Offset(0, 1).Value) Then
            This.index = "A"
        Else
            This.index = indexes(Application.WorksheetFunction.Match(LastIndex, indexes) + 1)
        End If
    End If

End Function

Private Property Get IIndex_Index() As String
    IIndex_Index = This.index
End Property

Private Property Get IIndex_Gezeichnet() As String
    IIndex_Gezeichnet = This.GezeichnetPerson & " ; " & This.GezeichnetDatum
End Property

Private Property Get IIndex_Geprüft() As String
    IIndex_Geprüft = This.GeprüftPerson & " ; " & This.GeprüftDatum
End Property

Private Property Get IIndex_Klartext() As String
    IIndex_Klartext = This.Klartext
End Property

Private Property Let IIndex_Index(ByVal NewValue As String)
    This.index = NewValue
End Property


