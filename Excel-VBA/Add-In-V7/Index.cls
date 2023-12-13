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
    Index As String
    ID As String
    IDPlan As String
    '--- Zeichner ---
    GezeichnetPerson As String
    GezeichnetDatum As String
    Gepr�ftPerson As String
    Gepr�ftDatum As String
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
       Optional ByVal ID As String = vbNullString, _
       Optional ByVal Letter As String = vbNullString, _
       Optional ByVal Gepr�ftPerson As String = vbNullString, _
       Optional ByVal Gepr�ftDatum As String = vbNullString _
       )

    Inputs.ID = ID
    Inputs.IDPlan = IDPlan
    Inputs.Index = Letter
    Inputs.GezeichnetPerson = GezeichnetPerson
    Inputs.GezeichnetDatum = GezeichnetDatum
    Inputs.Gepr�ftPerson = Gepr�ftPerson
    Inputs.Gepr�ftDatum = Gepr�ftDatum
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
    If Inputs.Gepr�ftPerson = "NONE" Then Inputs.Gepr�ftPerson = vbNullString
    If Inputs.Gepr�ftDatum = "NONE" Then Inputs.Gepr�ftDatum = ""
    '--- Planbezeichnung ---
    If Len(Inputs.Klartext) = 0 Then ERR.Raise 1, "Index > ValidateInputs", "No 'Klartext' given"
    validateInputs = True
    
    If This.ID = vbNullString Then This.ID = getNewID(4, Globals.shIndex, Globals.shIndex.range("A1").CurrentRegion, 1)
    
    If This.Index = vbNullString Then
    Globals.SetWBs
        Dim indexes(1 To 26) As String, LastIndex As String
        Dim i As Integer
        For i = 65 To 90
            indexes(i - 64) = Chr(i)
        Next
    End If

End Function

Private Property Get IIndex_IndexID() As String
    IIndex_IndexID = This.ID
End Property

Private Property Get IIndex_Index() As String
    IIndex_Index = This.Index
End Property

Private Property Let IIndex_Index(NewValue As String)
    This.Index = NewValue
End Property

Private Property Get IIndex_Gezeichnet() As String
    IIndex_Gezeichnet = This.GezeichnetPerson & " ; " & This.GezeichnetDatum
End Property

Private Property Get IIndex_Gepr�ft() As String
    IIndex_Gepr�ft = This.Gepr�ftPerson & " ; " & This.Gepr�ftDatum
End Property

Private Property Get IIndex_Klartext() As String
    IIndex_Klartext = This.Klartext
End Property

Private Property Get IIndex_PlanID() As String
    IIndex_PlanID = This.IDPlan
End Property

Private Property Let IIndex_PlanID(ByVal NewValue As String)
    This.IDPlan = NewValue
End Property

