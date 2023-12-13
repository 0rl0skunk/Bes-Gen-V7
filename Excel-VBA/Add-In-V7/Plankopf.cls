VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plankopf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Plankopf"
' To make a method available to the user it must have the prefix of the Interface
' in this case 'IPlankopf_'

Option Explicit

Implements IPlankopf

Private Type TPlankopf
    '--- Identifiers ---
    ID As String
    IDTinLine As String
    Planart As String: PlanartKF As String
    Plantyp As String
    '--- Zeichner ---
    GezeichnetPerson As String
    GezeichnetDatum As String
    Gepr�ftPerson As String
    Gepr�ftDatum As String
    '--- Geb�ude ---
    Geb�ude As String: Geb�udeKF As String
    Geb�udeTeil As String: Geb�udeteilKF As String
    Geschoss As String: GeschossKF As String: GeschossCode As String
    '--- Gewerk ---
    Gewerk As String: GewerkKF As String
    UnterGewerk As String: UnterGewerkKF As String
    '--- Planbezeichnung ---
    Klartext As String
    Plan�berschrift As String
    '--- Layout ---
    LayoutName As String
    LayoutGr�sse As String
    LayoutMasstab As String
    LayoutPlanstand As String
    '--- File Path ---
    PDFFileName As String
    DWGFileName As String
    DWGFilePath As String
    XMLFileName As String
    XMLFilePath As String
    '--- Index ---
    currentIndex As IIndex
    indexes As New Collection
End Type

Private This                 As TPlankopf
Private Inputs               As TPlankopf
Private pProjekt             As IProjekt

' -------------------------------
' --- Properties
' -------------------------------
Private Property Get IPlankopf_Gewerk() As String
    IPlankopf_Gewerk = This.Gewerk
End Property
Private Property Get IPlankopf_UnterGewerk() As String
    IPlankopf_UnterGewerk = This.UnterGewerk
End Property
Private Property Get IPlankopf_ID() As String
    IPlankopf_ID = This.ID
End Property

Private Property Let IPlankopf_ID(ByVal NewValue As String)
    This.ID = NewValue
End Property

Private Property Get IPlankopf_IDTinLine() As String
    IPlankopf_IDTinLine = This.IDTinLine
End Property

Private Property Get IPlankopf_Plantyp() As String
    IPlankopf_Plantyp = This.Plantyp
End Property

Private Property Get IPlankopf_Planart() As String
    IPlankopf_Planart = This.Planart
End Property

Private Property Get IPlankopf_GezeichnetPerson() As String
    IPlankopf_GezeichnetPerson = This.GezeichnetPerson
End Property

Private Property Get IPlankopf_GezeichnetDatum() As String
    IPlankopf_GezeichnetDatum = This.GezeichnetDatum
End Property

Private Property Get IPlankopf_Gezeichnet() As String
    IPlankopf_Gezeichnet = This.GezeichnetPerson & " ; " & This.GezeichnetDatum
End Property

Private Property Get IPlankopf_Gepr�ftPerson() As String
    IPlankopf_Gepr�ftPerson = This.Gepr�ftPerson
End Property

Private Property Get IPlankopf_Gepr�ftDatum() As String
    IPlankopf_Gepr�ftDatum = This.Gepr�ftDatum
End Property

Private Property Get IPlankopf_Gepr�ft() As String
    IPlankopf_Gepr�ft = This.Gepr�ftPerson & " ; " & This.Gepr�ftDatum
End Property

Private Property Get IPlankopf_Geb�ude() As String
    IPlankopf_Geb�ude = This.Geb�ude
End Property

Private Property Get IPlankopf_Geb�udeteil() As String
    IPlankopf_Geb�udeteil = This.Geb�udeTeil
End Property

Private Property Get IPlankopf_Geschoss() As String
    IPlankopf_Geschoss = This.Geschoss
End Property

Private Property Get IPlankopf_Klartext() As String
    IPlankopf_Klartext = This.Klartext
End Property

Private Property Get IPlankopf_Plan�berschrift() As String
    IPlankopf_Plan�berschrift = This.Plan�berschrift
End Property

Private Property Get IPlankopf_LayoutName() As String
    IPlankopf_LayoutName = This.LayoutName
End Property

Private Property Get IPlankopf_PlanNummer() As String
    IPlankopf_PlanNummer = This.LayoutName
End Property

Private Property Get IPlankopf_LayoutGr�sse() As String
    IPlankopf_LayoutGr�sse = This.LayoutGr�sse
End Property

Private Property Get IPlankopf_LayoutMasstab() As String
    IPlankopf_LayoutMasstab = This.LayoutMasstab
End Property

Private Property Get IPlankopf_LayoutPlanstand() As String
    IPlankopf_LayoutPlanstand = This.LayoutPlanstand
End Property

Private Property Get IPlankopf_PDFFileName() As String
    IPlankopf_PDFFileName = This.PDFFileName
End Property

Private Property Get IPlankopf_DWGFileName() As String
    IPlankopf_DWGFileName = This.DWGFileName
End Property

Private Property Get IPlankopf_DWGFile() As String
    IPlankopf_DWGFile = This.DWGFilePath & "\" & This.DWGFileName & ".dwg"
End Property

Private Property Get IPlankopf_DWGFilePath() As String
    IPlankopf_DWGFilePath = This.DWGFilePath
End Property

Private Property Get IPlankopf_XMLFileName() As String
    IPlankopf_XMLFileName = This.XMLFileName
End Property

Private Property Get IPlankopf_XMLFilePath() As String
    IPlankopf_XMLFilePath = This.XMLFilePath
End Property

Private Property Get IPlankopf_xmlfile() As String
    IPlankopf_xmlfile = This.XMLFilePath & "\" & This.XMLFileName
End Property

Private Property Get IPlankopf_currentIndex() As IIndex
    Set IPlankopf_currentIndex = This.currentIndex
End Property

Private Property Get IPlankopf_Indexes() As Collection
    Set IPlankopf_Indexes = This.indexes
End Property

Private Property Set IPlankopf_Indexes(ByVal NewIndexes As Collection)
    Set This.indexes = NewIndexes
End Property

' -------------------------------
' --- Functions
' -------------------------------

Public Function FillData( _
       ByVal Projekt As IProjekt, _
       ByVal GezeichnetPerson As String, _
       ByVal GezeichnetDatum As String, _
       ByVal Gepr�ftPerson As String, _
       ByVal Gepr�ftDatum As String, _
       ByVal Geb�ude As String, _
       ByVal Geb�udeTeil As String, _
       ByVal Geschoss As String, _
       ByVal Gewerk As String, _
       ByVal UnterGewerk As String, _
       ByVal Format As String, _
       ByVal Masstab As String, _
       ByVal Stand As String, _
       ByVal Klartext As String, _
       ByVal Planart As String, _
       Optional ByVal Plantyp As String, _
       Optional ByVal TinLineID As String = "NEW", _
       Optional ByVal SkipValidation As Boolean = False, _
       Optional ByVal Plan�berschrift As String = "NEW", _
       Optional ByVal ID As String = "NEW" _
       ) As Boolean
    ' skip validation should be used as a means of debuging or when creating a new Plankopf to give some template information but not validate it while loading
    With Inputs
        .IDTinLine = TinLineID
        .GezeichnetPerson = GezeichnetPerson
        .GezeichnetDatum = GezeichnetDatum
        .Gepr�ftPerson = Gepr�ftPerson
        .Gepr�ftDatum = Gepr�ftDatum
        .Geb�ude = Geb�ude
        .Geb�udeTeil = Geb�udeTeil
        .Gewerk = Gewerk
        .UnterGewerk = UnterGewerk
        .Geschoss = Geschoss
        .Klartext = Klartext
        .Plan�berschrift = Plan�berschrift
        .LayoutGr�sse = Format
        .LayoutMasstab = Masstab
        .LayoutPlanstand = Stand
        .Planart = Planart
        .Plantyp = Plantyp
    End With
    Set pProjekt = Projekt
    'On Error GoTo NotValid
    If SkipValidation Or ValidInputs Then
        This = Inputs
        If ID = "NEW" Then
            This.ID = getNewID(6, Globals.shStoreData, shStoreData.range("A1").CurrentRegion, 1)
        Else
            This.ID = ID
        End If
        
        GetKF
        
        This.DWGFilePath = pProjekt.ProjektOrdnerCAD & "\01_EP\02_EG"
        This.XMLFilePath = This.DWGFilePath
        This.DWGFileName = GetDWGFileName
        This.Plan�berschrift = GetPlan�berschrift
        Refresh
    End If
    On Error GoTo 0
    FillData = True
    Exit Function

NotValid:
    On Error GoTo -1
    On Error GoTo 0
    Dim frm                  As New UserFormMessage
    frm.typeWarning "Es wurde kein Plankopf erstellt." & vbNewLine & _
                    "�ffne den Log um die Fehlenden Inhalte zu sehen.", _
                    "Kein Plankopf erstellt", True
    frm.Show
    FillData = False
    Exit Function

End Function

Private Function GetKF()

On Error GoTo ERR
        This.Geb�udeKF = Globals.shGeb�ude.range("1:1").Find(This.Geb�ude).Offset(1).Value
        This.GeschossCode = Globals.shGeb�ude.range("B:B").Find(This.Geschoss).Offset(, -1).Value
        This.GeschossKF = Globals.shGeb�ude.range("B:B").Find(This.Geschoss).Offset(, 1).Value
        This.Geb�udeteilKF = Globals.shPData.range("PRO_Geb�udeteil").Find(This.Geb�udeTeil).Offset(0, 1).Value
        This.GewerkKF = Globals.shPData.range("PRO_Hauptgewerk").Find(This.Gewerk).Offset(0, 1).Value
        
        Dim col As Long
        Select Case This.Plantyp
        Case "PLA"
1           col = Application.Match(This.GewerkKF & " PLA", Globals.shPData.range("10:10"), 0)
            This.UnterGewerkKF = Globals.shPData.Columns(col).Find(This.UnterGewerk).Offset(0, 1).Value
2           col = Application.Match(This.GewerkKF, Globals.shPData.range("9:9"), 0)
            This.PlanartKF = Globals.shPData.Columns(col).Find(This.Planart).Offset(0, 1).Value
        Case "SCH"
3           col = Application.Match(This.GewerkKF & " SCH", Globals.shPData.range("10:10"), 0)
            This.UnterGewerkKF = Globals.shPData.Columns(col).Find(This.UnterGewerk).Offset(0, 1).Value
4           col = Application.Match(This.GewerkKF, Globals.shPData.range("9:9"), 0)
            This.PlanartKF = Globals.shPData.Columns(col).Find(This.Planart).Offset(0, 1).Value
        Case "PRI"
5           col = Application.Match(This.GewerkKF & " PRI", Globals.shPData.range("10:10"), 0)
            This.UnterGewerkKF = Globals.shPData.Columns(col).Find(This.UnterGewerk).Offset(0, 1).Value
6           col = Application.Match(This.GewerkKF, Globals.shPData.range("9:9"), 0)
            This.PlanartKF = Globals.shPData.Columns(col).Find(This.Planart).Offset(0, 1).Value
        End Select
Exit Function
ERR:
    Debug.Print ">>>   "; ERR.Description

End Function

Private Function GetDWGFileName() As String
    GetDWGFileName = "EP_EG.dwg"
End Function

Private Function GetPlan�berschrift() As String

    '--- log
    Log.Log "GetPlan�berschrift start"
    '--- log end

    Dim Plan�berschrift      As String


    If This.Plantyp = "PLA" Then
        ' �berschrift f�r Elektropl�ne
        Plan�berschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerk, This.Geschoss, This.Geb�udeTeil)
    Else
        ' �berschrift f�r Schemas
        Plan�berschrift = "Prinzipschema " & This.UnterGewerk

        If Len(Plan�berschrift) > Globals.maxlen Then
            ' the Description is too long to fit in the Plankopf
            ' use shorter forms
            Plan�berschrift = "Prinzipschema " & This.UnterGewerk
            If Len(Plan�berschrift) > maxlen Then
                Plan�berschrift = "PR " & This.UnterGewerk
            End If
        End If
    End If

    GetPlan�berschrift = Plan�berschrift
    '--- log
    Log.Log "GetPlan�berschrift END >> " & GetPlan�berschrift
    '--- log end

End Function

Private Function ValidInputs() As Boolean

    ValidInputs = False
    Dim ErrorSource          As String
    ErrorSource = "Plankopf > ValidInputs > "
    If Len(Inputs.GezeichnetPerson) = 0 Then Log.Log ErrorSource & "no 'INPUT GEZEICHNET PERSON' given"
    If Len(Inputs.GezeichnetDatum) = 0 Then Log.Log ErrorSource & "no 'INPUT GEZEICHNET DATUM' given"
    If Len(Inputs.Gepr�ftPerson) = 0 Then Log.Log ErrorSource & "no 'INPUT GEPR�FT PERSON' given"
    If Len(Inputs.Gepr�ftDatum) = 0 Then Log.Log ErrorSource & "no 'INPUT GEPR�FT DATUM' given"
    If Len(Inputs.Geb�ude) = 0 Then Log.Log ErrorSource & "no 'INPUT GEB�UDE' given"
    If Len(Inputs.Geb�udeTeil) = 0 Then Log.Log ErrorSource & "no 'INPUT GEB�UDETEIL' given"
    If Len(Inputs.Geschoss) = 0 Then Log.Log ErrorSource & "no 'INPUT GESCHOSS' given"
    If Len(Inputs.Plan�berschrift) = 0 Then Log.Log ErrorSource & "no 'INPUT PLAN�BERSCHRIFT' given"
    If Len(Inputs.LayoutGr�sse) = 0 Then Log.Log ErrorSource & "no 'INPUT FORMAT' given"
    If Len(Inputs.LayoutMasstab) = 0 Then Log.Log ErrorSource & "no 'INPUT MASSTAB' given"
    If Len(Inputs.LayoutPlanstand) = 0 Then Log.Log ErrorSource & "no 'INPUT PLANSTAND' given"
    ValidInputs = True

End Function

Public Function Refresh()

    '--- log
    Log.Log "Output start"
    '--- log end

    Dim dateToday            As String
    dateToday = Format(Now(), "yymmdd")
    'Select Case This.PlanTyp
DWGEP:
    This.LayoutName = WorksheetFunction.TextJoin _
                      ("-", True, This.Geb�udeKF, This.Geb�udeteilKF, This.GeschossKF, This.PlanartKF, This.UnterGewerkKF)
PDFEP:
    This.PDFFileName = WorksheetFunction.TextJoin _
                       ("_", True, shPData.range("ADM_Projektnummer").Value, "04" & This.GewerkKF, This.GeschossCode, This.LayoutName, dateToday) & ".PDF"
End Function

' -------------------------------
' --- Subs
' -------------------------------
Private Sub IPlankopf_AddIndex(NewValue As IIndex)

    This.indexes.Add NewValue
    Set This.currentIndex = NewValue

End Sub
Private Sub IPlankopf_ClearIndex()

    Set This.indexes = New Collection

End Sub

