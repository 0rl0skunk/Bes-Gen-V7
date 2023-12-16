VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plankopf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Plankopf Objekt für einfacheres Management von Planköpfen."
Option Explicit
'@Folder "Plankopf"
'@ModuleDescription "Plankopf Objekt für einfacheres Management von Planköpfen."


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
    GeprüftPerson As String
    GeprüftDatum As String
    '--- Gebäude ---
    Gebäude As String: GebäudeKF As String
    GebäudeTeil As String: GebäudeteilKF As String
    Geschoss As String: GeschossKF As String: GeschossCode As String
    '--- Gewerk ---
    Gewerk As String: GewerkKF As String
    UnterGewerk As String: UnterGewerkKF As String
    '--- Planbezeichnung ---
    Klartext As String
    Planüberschrift As String
    '--- Layout ---
    LayoutName As String
    LayoutGrösse As String
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

Private Property Get IPlankopf_GeprüftPerson() As String
    IPlankopf_GeprüftPerson = This.GeprüftPerson
End Property

Private Property Get IPlankopf_GeprüftDatum() As String
    IPlankopf_GeprüftDatum = This.GeprüftDatum
End Property

Private Property Get IPlankopf_Geprüft() As String
    IPlankopf_Geprüft = This.GeprüftPerson & " ; " & This.GeprüftDatum
End Property

Private Property Get IPlankopf_Gebäude() As String
    IPlankopf_Gebäude = This.Gebäude
End Property

Private Property Get IPlankopf_Gebäudeteil() As String
    IPlankopf_Gebäudeteil = This.GebäudeTeil
End Property

Private Property Get IPlankopf_Geschoss() As String
    IPlankopf_Geschoss = This.Geschoss
End Property

Private Property Get IPlankopf_Klartext() As String
    IPlankopf_Klartext = This.Klartext
End Property

Private Property Get IPlankopf_Planüberschrift() As String
    IPlankopf_Planüberschrift = This.Planüberschrift
End Property

Private Property Get IPlankopf_LayoutName() As String
    IPlankopf_LayoutName = This.LayoutName
End Property

Private Property Get IPlankopf_PlanNummer() As String
    IPlankopf_PlanNummer = This.LayoutName
End Property

Private Property Get IPlankopf_LayoutGrösse() As String
    IPlankopf_LayoutGrösse = This.LayoutGrösse
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
    If This.currentIndex Is Nothing Then Set This.currentIndex = IndexFactory.Create(This.ID, vbNullString, vbNullString, vbNullString, vbNullString, " ", vbNullString, vbNullString, True)
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

Public Function Filldata( _
       ByVal Projekt As IProjekt, _
       ByVal GezeichnetPerson As String, _
       ByVal GezeichnetDatum As String, _
       ByVal GeprüftPerson As String, _
       ByVal GeprüftDatum As String, _
       ByVal Gebäude As String, _
       ByVal GebäudeTeil As String, _
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
       Optional ByVal Planüberschrift As String = "NEW", _
       Optional ByVal ID As String = "NEW" _
       ) As Boolean
    ' SkipValidation ist zu debug zwecken zu gebrauchen
    With Inputs
        .IDTinLine = TinLineID
        .GezeichnetPerson = GezeichnetPerson
        .GezeichnetDatum = GezeichnetDatum
        .GeprüftPerson = GeprüftPerson
        .GeprüftDatum = GeprüftDatum
        .Gebäude = Gebäude
        .GebäudeTeil = GebäudeTeil
        .Gewerk = Gewerk
        .UnterGewerk = UnterGewerk
        .Geschoss = Geschoss
        .Klartext = Klartext
        .Planüberschrift = Planüberschrift
        .LayoutGrösse = Format
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
        GetPlanüberschrift
        Refresh
    End If
    On Error GoTo 0
    Filldata = True
    Exit Function

    On Error GoTo -1
    On Error GoTo 0
    Dim frm                  As New UserFormMessage
    frm.typeWarning "Es wurde kein Plankopf erstellt." & vbNewLine & _
                    "öffne den Log um die Fehlenden Inhalte zu sehen.", _
                    "Kein Plankopf erstellt", True
    frm.Show
    Filldata = False
    Exit Function

End Function

Private Sub GetKF()

    On Error GoTo ERR
    This.GebäudeKF = Globals.shGebäude.range("1:1").Find(This.Gebäude).Offset(1).Value
    This.GeschossCode = Globals.shGebäude.range("B:B").Find(This.Geschoss).Offset(, -1).Value
    This.GeschossKF = Globals.shGebäude.range("B:B").Find(This.Geschoss).Offset(, 1).Value
    This.GebäudeteilKF = Globals.shPData.range("PRO_Gebäudeteil").Find(This.GebäudeTeil).Offset(0, 1).Value
    This.GewerkKF = Globals.shPData.range("PRO_Hauptgewerk").Find(This.Gewerk).Offset(0, 1).Value

    Dim col                  As Long
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

    writelog "Info", "Kurzformen erstellt"
    Exit Sub
ERR:
Debug.Print ">>>   "; ERR.description
    writelog "Error", "Beim erstellen der Kurzformen" & vbNewLine & _
                     ERR.Number & vbNewLine & ERR.description & vbNewLine & ERR.source

End Sub

Private Function GetDWGFileName() As String
    GetDWGFileName = "EP_EG.dwg"
End Function

Private Sub GetPlanüberschrift()

    Dim Planüberschrift      As String

    If This.Plantyp = "PLA" Then
        ' Überschrift für Elektropläne
        Planüberschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerk, This.Geschoss, This.GebäudeTeil)
        If Len(Planüberschrift) > Globals.maxlen Then
            Planüberschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerk, This.GeschossKF, This.GebäudeTeil)
            If Len(Planüberschrift) > maxlen Then
                Planüberschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerk, This.GeschossKF, This.GebäudeteilKF)
                If Len(Planüberschrift) > maxlen Then
                    Planüberschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerkKF, This.GeschossKF, This.GebäudeteilKF)
                    If Len(Planüberschrift) > maxlen Then
                        Planüberschrift = Application.WorksheetFunction.TextJoin(" ", True, This.PlanartKF, This.UnterGewerkKF, This.GeschossKF, This.GebäudeteilKF)
                    End If
                End If
            End If
        End If
    Else
        ' Überschrift für Schemas
        Planüberschrift = "Prinzipschema " & This.UnterGewerk
        If Len(Planüberschrift) > Globals.maxlen Then
            Planüberschrift = "Prinzipschema " & This.UnterGewerk
            If Len(Planüberschrift) > maxlen Then
                Planüberschrift = "PR " & This.UnterGewerk
            End If
        End If
    End If

    This.Planüberschrift = Planüberschrift

    writelog "Info", "Planüberschrift erstellt: " & This.Planüberschrift

End Sub

Private Function ValidInputs() As Boolean

    ValidInputs = False
    Dim ErrorSource          As String
    ErrorSource = "Plankopf > ValidInputs > "
    If Len(Inputs.GezeichnetPerson) = 0 Then writelog "error", ErrorSource & "no 'INPUT GEZEICHNET PERSON' given"
    If Len(Inputs.GezeichnetDatum) = 0 Then writelog "error", ErrorSource & "no 'INPUT GEZEICHNET DATUM' given"
    If Len(Inputs.GeprüftPerson) = 0 Then writelog "error", ErrorSource & "no 'INPUT GEPRÜFT PERSON' given"
    If Len(Inputs.GeprüftDatum) = 0 Then writelog "error", ErrorSource & "no 'INPUT GEPRÜFT DATUM' given"
    If Len(Inputs.Gebäude) = 0 Then writelog "error", ErrorSource & "no 'INPUT GEBÄUDE' given"
    If Len(Inputs.GebäudeTeil) = 0 Then writelog "error", ErrorSource & "no 'INPUT GEBÄUDETEIL' given"
    If Len(Inputs.Geschoss) = 0 Then writelog "error", ErrorSource & "no 'INPUT GESCHOSS' given"
    If Len(Inputs.Planüberschrift) = 0 Then writelog "error", ErrorSource & "no 'INPUT PLANÜBERSCHRIFT' given"
    If Len(Inputs.LayoutGrösse) = 0 Then writelog "error", ErrorSource & "no 'INPUT FORMAT' given"
    If Len(Inputs.LayoutMasstab) = 0 Then writelog "error", ErrorSource & "no 'INPUT MASSTAB' given"
    If Len(Inputs.LayoutPlanstand) = 0 Then writelog "error", ErrorSource & "no 'INPUT PLANSTAND' given"
    ValidInputs = True

    writelog "Info", "Plankopf erstellt"

End Function

Public Sub Refresh()

    Dim dateToday            As String
    dateToday = Format(Now(), "yymmdd")
    'Select Case This.PlanTyp
    This.LayoutName = WorksheetFunction.TextJoin _
                      ("-", True, This.GebäudeKF, This.GebäudeteilKF, This.GeschossKF, This.PlanartKF, This.UnterGewerkKF)
    This.PDFFileName = WorksheetFunction.TextJoin _
                       ("_", True, shPData.range("ADM_Projektnummer").Value, "04" & This.GewerkKF, This.GeschossCode, This.LayoutName, dateToday) & ".PDF"

    writelog "Info", "Planablage-Namen aktualisiert" & This.PDFFileName
End Sub

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


