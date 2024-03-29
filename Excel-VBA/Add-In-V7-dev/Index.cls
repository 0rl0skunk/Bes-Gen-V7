VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Index Objekt f�r einfacheres Management von Indexen."

'@Folder("Index")
'@IgnoreModule VariableNotUsed, EmptyStringLiteral
'@ModuleDescription "Index Objekt f�r einfacheres Management von Indexen."

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
    SkipValidation As Boolean
End Type

Private This                 As TIndex
Private Inputs               As TIndex
Private created              As Boolean

Public Function Filldata( _
       ByVal IDPlan As String, _
       ByVal GezeichnetPerson As String, _
       ByVal GezeichnetDatum As String, _
       ByVal Klartext As String, _
       Optional ByVal ID As String = vbNullString, _
       Optional ByVal Letter As String = vbNullString, _
       Optional ByVal Gepr�ftPerson As String = vbNullString, _
       Optional ByVal Gepr�ftDatum As String = vbNullString, _
       Optional ByVal SkipValidation As Boolean _
       ) As Boolean
    created = False
    Inputs.ID = ID
    Inputs.IDPlan = IDPlan
    Inputs.Index = Letter
    Inputs.GezeichnetPerson = GezeichnetPerson
    Inputs.GezeichnetDatum = GezeichnetDatum
    Inputs.Gepr�ftPerson = Gepr�ftPerson
    Inputs.Gepr�ftDatum = Gepr�ftDatum
    Inputs.Klartext = Klartext
    Inputs.SkipValidation = SkipValidation

    If validateInputs Then
        This = Inputs
        created = True
    End If

    writelog LogInfo, "Index erstellt"

End Function

Private Function validateInputs() As Boolean

    If Inputs.SkipValidation Then: validateInputs = True: Exit Function

    ' sind alle ben�tigten Daten gegeben und korrekt
    validateInputs = False
    If Len(Inputs.IDPlan) = 0 Then err.Raise 1, "Index > ValidateInputs", "No 'Plan ID' given"
    '--- Zeichner ---
    If Len(Inputs.GezeichnetPerson) = 0 Then writelog LogError, "Index > ValidateInputs > No 'GezeichnetPerson' given": created = False
    If Len(Inputs.GezeichnetDatum) = 0 Then writelog LogError, "Index > ValidateInputs > No 'GezeichnetDatum' given": created = False
    If Inputs.Gepr�ftPerson = "NONE" Then Inputs.Gepr�ftPerson = vbNullString
    If Inputs.Gepr�ftDatum = "NONE" Then Inputs.Gepr�ftDatum = vbNullString

    '--- Planbezeichnung ---
    If Len(Inputs.Klartext) = 0 Then writelog LogError, "Index > ValidateInputs > No 'Klartext' given"
    validateInputs = True

    If This.ID = vbNullString Then This.ID = getNewID(IDIndex)

    If This.Index = vbNullString Then
        Dim Indexes(1 To 26) As String
        Dim LastIndex        As String

        Dim i                As Long
        'For i = 65 To 90
        '    indexes(i - 64) = Chr(i)
        'Next
    Else
    End If

End Function

Private Property Get XMLFile() As String
    XMLFile = This.XMLFilePath & "\" & This.XMLFileName
End Property

Private Property Get dwgFile() As String
    dwgFile = This.DWGFilePath & "\" & This.DWGFileName
End Property

Private Property Get IIndex_IndexID() As String
    IIndex_IndexID = This.ID
End Property

Private Property Get IIndex_Index() As String
    IIndex_Index = This.Index
End Property
Private Property Let IIndex_Index(ByVal NewValue As String)
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


